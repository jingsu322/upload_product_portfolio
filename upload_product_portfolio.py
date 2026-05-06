#!/usr/bin/env python3
"""
upload_product_portfolio.py

Reads a successful_domains.xlsx (from crawl4ai analysis) and a HubSpot Companies
export, matches domains -> HubSpot record IDs, and updates the `top_ingredients`
property in HubSpot via the Companies batch update API.

Behavior summary
----------------
Skips a source row when:
  - The source has empty `product_portfolio`
  - The domain is not present in the HubSpot export (assumed deleted / missing)
  - The merged value equals the value already in HubSpot (no-op)

For records that already have a value, the merge strategy is configurable via --mode:
  - append   (default): merge old + new, case-insensitive dedup, preserve old order,
                        new items appended at the end. Joined with ", ".
  - replace            : overwrite entirely with the new value.
  - skip               : leave existing value unchanged (only fill empty fields).

Outputs (in --output-dir, default ./outputs):
  - rollback_log_<ts>.csv : record_id, domain, old_value, new_value, status, applied_at
                            (use this to revert: write old_value back to record_id)
  - skipped_<ts>.csv      : domain, reason (everything that did NOT result in an API call)
  - upload_summary_<ts>.txt : human-readable stats

Auth / config from .env (loaded automatically):
  HUBSPOT_PRIVATE_APP_TOKEN=pat-na1-...
  HUBSPOT_DELAY=0.2
  # Optional override (defaults to "top_ingredients"):
  # HUBSPOT_TOP_INGREDIENTS_FIELD=top_ingredients

Usage:
  python upload_product_portfolio.py \
    --successful-domains successful_domains.xlsx \
    --hubspot-export hubspot-crm-exports.xlsx \
    [--mode append|replace|skip] \
    [--dry-run] \
    [--yes] \
    [--output-dir ./outputs]
"""

from __future__ import annotations

import argparse
import csv
import os
import sys
import time
from datetime import datetime
from pathlib import Path
from typing import Optional

import pandas as pd
import requests
from dotenv import load_dotenv

HUBSPOT_BATCH_UPDATE_URL = "https://api.hubapi.com/crm/v3/objects/companies/batch/update"
BATCH_SIZE = 100  # HubSpot batch endpoint limit
SEPARATOR = ", "

SRC_DOMAIN_COL = "company_domain"
SRC_PORTFOLIO_COL = "product_portfolio"

HUB_ID_COL = "Record ID"
HUB_DOMAIN_COL = "Company Domain Name"
HUB_PORTFOLIO_COL = "Product Portfolio"


# ---------- helpers ----------

def normalize_domain(d) -> str:
    if d is None:
        return ""
    s = str(d).strip().lower()
    # strip common protocol/path noise just in case
    for prefix in ("https://", "http://", "www."):
        if s.startswith(prefix):
            s = s[len(prefix):]
    s = s.split("/", 1)[0]
    return s.strip()


_OPEN_BRACKETS = "([{"
_CLOSE_BRACKETS = ")]}"


def parse_portfolio(s) -> list[str]:
    """
    Split a portfolio string into items on top-level commas only, ignoring
    commas that are inside (), [], or {}. Trim whitespace; drop empties.

    The data contains items like:
        "Flavoring materials (bonito powder, bonito extract, kelp powder)"
        "Chicken seasoning powder (contains wheat, soybean, pork)"
    A naive .split(",") would shred those, which combined with dedup would
    silently corrupt the value. Tracking bracket depth avoids that.
    """
    if s is None:
        return []
    if isinstance(s, float) and pd.isna(s):
        return []
    text = str(s).strip()
    if not text:
        return []

    items: list[str] = []
    buf: list[str] = []
    depth = 0
    for ch in text:
        if ch in _OPEN_BRACKETS:
            depth += 1
            buf.append(ch)
        elif ch in _CLOSE_BRACKETS:
            if depth > 0:
                depth -= 1
            buf.append(ch)
        elif ch == "," and depth == 0:
            piece = "".join(buf).strip()
            if piece:
                items.append(piece)
            buf = []
        else:
            buf.append(ch)
    tail = "".join(buf).strip()
    if tail:
        items.append(tail)
    return items


def merge_portfolios(old_value: str, new_value: str, mode: str) -> Optional[str]:
    """
    Returns the value to write to HubSpot.
    Returns None to signal "skip this record entirely" (only used when
    mode=='skip' and old already populated).
    """
    old_items = parse_portfolio(old_value)
    new_items = parse_portfolio(new_value)

    if mode == "replace":
        return SEPARATOR.join(new_items)

    if mode == "skip":
        if old_items:
            return None
        return SEPARATOR.join(new_items)

    # mode == "append": case-insensitive dedup, keep old order, append truly new
    seen_lower = {x.lower() for x in old_items}
    merged = list(old_items)
    for item in new_items:
        key = item.lower()
        if key not in seen_lower:
            merged.append(item)
            seen_lower.add(key)
    return SEPARATOR.join(merged)


def fmt_record_id(v) -> Optional[str]:
    """HubSpot Record ID may come in as int / float / str. Normalize to a clean string."""
    if v is None:
        return None
    if isinstance(v, float):
        if pd.isna(v):
            return None
        return str(int(v))
    if isinstance(v, int):
        return str(v)
    s = str(v).strip()
    return s or None


# ---------- main ----------

def main() -> int:
    ap = argparse.ArgumentParser(
        description="Upload product portfolio data to HubSpot Companies (top_ingredients property)",
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    ap.add_argument("--successful-domains", required=True,
                    help="Path to successful_domains.xlsx")
    ap.add_argument("--hubspot-export", required=True,
                    help="Path to HubSpot Companies export .xlsx")
    ap.add_argument("--mode", choices=["append", "replace", "skip"], default="append",
                    help="Merge strategy when target already has a value (default: append)")
    ap.add_argument("--dry-run", action="store_true",
                    help="Compute everything and write rollback_log preview, but do NOT call HubSpot")
    ap.add_argument("--yes", action="store_true",
                    help="Skip the interactive confirmation before applying updates")
    ap.add_argument("--output-dir", default="./outputs",
                    help="Where to write logs (default: ./outputs)")
    args = ap.parse_args()

    load_dotenv()
    token = os.getenv("HUBSPOT_PRIVATE_APP_TOKEN")
    delay = float(os.getenv("HUBSPOT_DELAY", "0.2"))
    field_name = os.getenv("HUBSPOT_TOP_INGREDIENTS_FIELD", "top_ingredients")

    if not args.dry_run and not token:
        print("ERROR: HUBSPOT_PRIVATE_APP_TOKEN not set in .env (or environment).",
              file=sys.stderr)
        return 1

    out_dir = Path(args.output_dir)
    out_dir.mkdir(parents=True, exist_ok=True)
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")

    # ---- load source ----
    print(f"Loading source: {args.successful_domains}")
    src_df = pd.read_excel(args.successful_domains)
    for col in (SRC_DOMAIN_COL, SRC_PORTFOLIO_COL):
        if col not in src_df.columns:
            print(f"ERROR: column '{col}' missing in source file. Found: {list(src_df.columns)}",
                  file=sys.stderr)
            return 1
    print(f"  {len(src_df):,} source rows")

    # ---- load HubSpot export ----
    print(f"Loading HubSpot export: {args.hubspot_export}")
    hub_df = pd.read_excel(args.hubspot_export)
    for col in (HUB_ID_COL, HUB_DOMAIN_COL):
        if col not in hub_df.columns:
            print(f"ERROR: column '{col}' missing in HubSpot export. Found: {list(hub_df.columns)}",
                  file=sys.stderr)
            return 1
    print(f"  {len(hub_df):,} HubSpot rows")

    # ---- build domain -> [(record_id, current_value), ...] ----
    hub_map: dict[str, list[tuple[str, str]]] = {}
    null_domain_rows = 0
    null_id_rows = 0
    for _, row in hub_df.iterrows():
        d = normalize_domain(row.get(HUB_DOMAIN_COL))
        if not d:
            null_domain_rows += 1
            continue
        rid = fmt_record_id(row.get(HUB_ID_COL))
        if rid is None:
            null_id_rows += 1
            continue
        cur = row.get(HUB_PORTFOLIO_COL) if HUB_PORTFOLIO_COL in hub_df.columns else ""
        cur_str = "" if (cur is None or (isinstance(cur, float) and pd.isna(cur))) else str(cur)
        hub_map.setdefault(d, []).append((rid, cur_str))

    duplicate_domains = {d: recs for d, recs in hub_map.items() if len(recs) > 1}
    if duplicate_domains:
        print(f"  Note: {len(duplicate_domains):,} domains map to >1 HubSpot record "
              f"(all matching records will be updated)")

    # ---- compute updates ----
    updates: list[dict] = []
    skipped: list[tuple[str, str]] = []  # (domain, reason)

    for _, row in src_df.iterrows():
        raw_domain = row.get(SRC_DOMAIN_COL)
        domain = normalize_domain(raw_domain)
        new_raw = row.get(SRC_PORTFOLIO_COL)

        if not domain:
            skipped.append((str(raw_domain) if raw_domain is not None else "", "empty source domain"))
            continue
        if new_raw is None or (isinstance(new_raw, float) and pd.isna(new_raw)) or not str(new_raw).strip():
            skipped.append((domain, "empty product_portfolio in source"))
            continue
        if domain not in hub_map:
            skipped.append((domain, "not found in HubSpot export"))
            continue

        new_value_raw = str(new_raw)
        for rid, old_value in hub_map[domain]:
            merged = merge_portfolios(old_value, new_value_raw, args.mode)
            if merged is None:
                skipped.append((domain, f"record {rid}: mode=skip and target already populated"))
                continue
            if merged.strip() == (old_value or "").strip():
                skipped.append((domain, f"record {rid}: no change (merged value identical)"))
                continue
            updates.append({
                "record_id": rid,
                "domain": domain,
                "old_value": old_value,
                "new_value": merged,
            })

    # ---- preflight summary ----
    not_found = sum(1 for _, r in skipped if "not found" in r)
    no_change = sum(1 for _, r in skipped if "no change" in r)
    src_empty = sum(1 for _, r in skipped if "empty product_portfolio" in r)
    skip_existing = sum(1 for _, r in skipped if "mode=skip" in r)

    print()
    print("=" * 60)
    print("Pre-flight summary")
    print("=" * 60)
    print(f"  Mode                          : {args.mode}")
    print(f"  HubSpot field (internal name) : {field_name}")
    print(f"  Dry run                       : {args.dry_run}")
    print(f"  Source rows                   : {len(src_df):,}")
    print(f"  Updates to apply              : {len(updates):,}")
    print(f"  Skipped                       : {len(skipped):,}")
    print(f"    - not in HubSpot export     : {not_found:,}")
    print(f"    - empty source value        : {src_empty:,}")
    print(f"    - no change needed          : {no_change:,}")
    if skip_existing:
        print(f"    - target already populated  : {skip_existing:,}")

    # ---- write skipped log ----
    skipped_path = out_dir / f"skipped_{ts}.csv"
    with open(skipped_path, "w", newline="", encoding="utf-8") as f:
        w = csv.writer(f)
        w.writerow(["domain", "reason"])
        w.writerows(skipped)
    print(f"  Wrote: {skipped_path}")

    rollback_path = out_dir / f"rollback_log_{ts}.csv"

    if not updates:
        print("\nNothing to update. Done.")
        return 0

    # ---- dry run early exit ----
    if args.dry_run:
        with open(rollback_path, "w", newline="", encoding="utf-8") as f:
            w = csv.writer(f)
            w.writerow(["record_id", "domain", "old_value", "new_value", "status", "applied_at"])
            for u in updates:
                w.writerow([u["record_id"], u["domain"], u["old_value"], u["new_value"], "DRY_RUN", ""])
        print(f"  Wrote: {rollback_path} (dry run preview)")
        print("\nDry run complete. Re-run without --dry-run to apply.")
        return 0

    # ---- confirm ----
    if not args.yes:
        print()
        try:
            ans = input(f"Apply {len(updates):,} updates to HubSpot? [y/N]: ").strip().lower()
        except EOFError:
            ans = ""
        if ans not in ("y", "yes"):
            print("Aborted by user.")
            return 1

    # ---- apply in batches ----
    headers = {
        "Authorization": f"Bearer {token}",
        "Content-Type": "application/json",
    }

    succeeded = 0
    failed = 0
    results: list[tuple[dict, str, str]] = []  # (update, status, applied_at)

    n_batches = (len(updates) + BATCH_SIZE - 1) // BATCH_SIZE
    print(f"\nApplying {len(updates):,} updates in {n_batches:,} batch(es) of up to {BATCH_SIZE}...")

    for bi in range(n_batches):
        batch = updates[bi * BATCH_SIZE:(bi + 1) * BATCH_SIZE]
        payload = {
            "inputs": [
                {"id": u["record_id"], "properties": {field_name: u["new_value"]}}
                for u in batch
            ]
        }
        now_iso = datetime.now().isoformat(timespec="seconds")
        try:
            resp = requests.post(HUBSPOT_BATCH_UPDATE_URL, headers=headers,
                                 json=payload, timeout=60)
        except requests.exceptions.RequestException as e:
            err = f"ERROR: {type(e).__name__}: {e}"
            for u in batch:
                results.append((u, err, now_iso))
                failed += 1
            print(f"  Batch {bi + 1}/{n_batches}: network error - {e}", file=sys.stderr)
        else:
            if resp.status_code in (200, 201, 207):
                try:
                    body = resp.json()
                except ValueError:
                    body = {}
                ok_ids = {str(item.get("id")) for item in body.get("results", []) or []}

                # Map errors back to ids when possible.
                err_by_id: dict[str, str] = {}
                for err_obj in (body.get("errors") or []):
                    msg = err_obj.get("message", "unknown error")
                    ctx = err_obj.get("context") or {}
                    ids_field = ctx.get("ids") or ctx.get("objectId") or []
                    if isinstance(ids_field, (list, tuple)):
                        for v in ids_field:
                            err_by_id[str(v)] = msg
                    else:
                        err_by_id[str(ids_field)] = msg

                for u in batch:
                    if u["record_id"] in ok_ids:
                        results.append((u, "OK", now_iso))
                        succeeded += 1
                    else:
                        msg = err_by_id.get(u["record_id"], "no result returned")
                        results.append((u, f"ERROR: {msg}", now_iso))
                        failed += 1
            else:
                err = f"HTTP {resp.status_code}: {resp.text[:300]}"
                for u in batch:
                    results.append((u, f"ERROR: {err}", now_iso))
                    failed += 1
                print(f"  Batch {bi + 1}/{n_batches}: {err}", file=sys.stderr)

        done = (bi + 1) * BATCH_SIZE
        done = min(done, len(updates))
        print(f"  Batch {bi + 1}/{n_batches} done ({done}/{len(updates)}, "
              f"ok={succeeded}, fail={failed})")

        if bi < n_batches - 1:
            time.sleep(delay)

    # ---- write rollback log ----
    with open(rollback_path, "w", newline="", encoding="utf-8") as f:
        w = csv.writer(f)
        w.writerow(["record_id", "domain", "old_value", "new_value", "status", "applied_at"])
        for u, status, when in results:
            w.writerow([u["record_id"], u["domain"], u["old_value"], u["new_value"], status, when])
    print(f"\nWrote rollback log: {rollback_path}")

    # ---- final summary ----
    summary_lines = [
        "=" * 60,
        "Upload summary",
        "=" * 60,
        f"  Mode                       : {args.mode}",
        f"  HubSpot field              : {field_name}",
        f"  Source rows                : {len(src_df):,}",
        f"  Updates attempted          : {len(updates):,}",
        f"  Succeeded                  : {succeeded:,}",
        f"  Failed                     : {failed:,}",
        f"  Skipped (pre-API)          : {len(skipped):,}",
        f"    - not in HubSpot export  : {not_found:,}",
        f"    - empty source value     : {src_empty:,}",
        f"    - no change needed       : {no_change:,}",
        f"    - target already populated: {skip_existing:,}" if skip_existing else "",
        f"  Domains w/ multiple records: {len(duplicate_domains):,}",
        f"  Rollback log               : {rollback_path}",
        f"  Skipped log                : {skipped_path}",
    ]
    summary_lines = [ln for ln in summary_lines if ln]
    summary_text = "\n".join(summary_lines)
    print()
    print(summary_text)

    summary_path = out_dir / f"upload_summary_{ts}.txt"
    with open(summary_path, "w", encoding="utf-8") as f:
        f.write(summary_text + "\n")
    print(f"\nWrote summary: {summary_path}")

    return 0 if failed == 0 else 2


if __name__ == "__main__":
    sys.exit(main())
