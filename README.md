# product_portfolio uploader for HubSpot

Updates the `top_ingredients` property on HubSpot Company records (display
label "Product Portfolio"), using two Excel files as input:

1. **`successful_domains.xlsx`** — produced by your crawl4ai analysis. Must
   contain at least these columns:
   - `company_domain`
   - `product_portfolio`
2. **HubSpot Companies export** (the one exported from
   `https://app.hubspot.com/contacts/<portal>/objects/0-2/views/<view>/list`).
   Must contain at least:
   - `Record ID`
   - `Company Domain Name`
   - `Product Portfolio` (used to compute the merge with existing data)

The script joins the two on domain (case-insensitive, with `https://`,
`http://`, `www.` and trailing paths stripped), then writes the merged value
back to HubSpot via the Companies batch update API.

---

## Install

Python 3.9+.

```bash
pip install -r requirements.txt
```

## Configure

Copy `.env.example` to `.env` and fill it in:

```env
HUBSPOT_PRIVATE_APP_TOKEN=pat-na1-XXXXXXXX
HUBSPOT_DELAY=0.2
# Optional. Defaults to "top_ingredients" — only set this if your custom
# property uses a different internal name.
# HUBSPOT_TOP_INGREDIENTS_FIELD=top_ingredients
```

The token must come from a HubSpot **Private App** with the
`crm.objects.companies.write` scope (and `read` for completeness).

---

## Run

Always start with `--dry-run` to see the plan and inspect the rollback log
preview before touching HubSpot:

```bash
python upload_product_portfolio.py \
    --successful-domains successful_domains.xlsx \
    --hubspot-export hubspot-crm-exports-product-portfolio-upload-check-2026-05-04.xlsx \
    --dry-run
```

Then commit:

```bash
python upload_product_portfolio.py \
    --successful-domains successful_domains.xlsx \
    --hubspot-export hubspot-crm-exports-product-portfolio-upload-check-2026-06-04.xlsx \
    --output-dir ./outputs
```

You'll be prompted to confirm before any HubSpot calls. Pass `--yes` to skip
the prompt (e.g. for automation).

### CLI flags

| Flag | Default | Meaning |
| --- | --- | --- |
| `--successful-domains` | (required) | Path to the crawl4ai output xlsx. |
| `--hubspot-export` | (required) | Path to the HubSpot CSV/xlsx export. |
| `--mode` | `append` | How to combine with existing HubSpot value (see below). |
| `--dry-run` | off | Compute and write rollback preview. Do **not** call HubSpot. |
| `--yes` | off | Skip the interactive "Apply N updates?" prompt. |
| `--output-dir` | `./outputs` | Where logs are written. |

### Merge modes (`--mode`)

When a HubSpot record already has something in `top_ingredients`:

- **`append`** *(default, recommended)* — merge old + new, **case-insensitive
  dedup**, preserve the existing order, and append truly-new items at the end.
  So `"Bonito Extract"` and `"bonito extract"` are treated as the same item;
  this avoids slow drift from re-runs.
- **`replace`** — overwrite with the new value entirely. Use this only if you
  trust the source completely.
- **`skip`** — only fill records where the field is empty; leave populated
  records untouched.

In every mode, items containing commas inside parentheses (e.g.
`"Flavoring materials (bonito powder, bonito extract, kelp powder)"`) are
treated as **single items** — the splitter is bracket-aware so it doesn't
shred them.

---

## What gets skipped

A source row is skipped (not sent to HubSpot) and recorded in
`skipped_<ts>.csv` when:

- The source row's `product_portfolio` is empty.
- The domain isn't present in the HubSpot export — assumed deleted in HubSpot
  since the export was taken. Re-export from HubSpot if you want to retry.
- The merged result is identical to what's already in HubSpot (no-op).
- (`--mode skip` only) The HubSpot record already has a value.

## Duplicate domains in HubSpot

If multiple HubSpot company records share a domain, **all of them get
updated**. The pre-flight summary tells you how many domains this affects.
Each affected `record_id` gets its own row in the rollback log.

## Outputs

All in `--output-dir` (default `./outputs`), suffixed with a timestamp:

- **`rollback_log_<ts>.csv`** — one row per record_id we attempted to update.
  Columns: `record_id, domain, old_value, new_value, status, applied_at`.
  - `status` is `OK`, `DRY_RUN`, or `ERROR: <reason>`.
  - To roll a record back, set `top_ingredients` for that `record_id` back
    to the `old_value` from this file. (See "Rolling back" below.)
- **`skipped_<ts>.csv`** — domains/records that were skipped and why.
- **`upload_summary_<ts>.txt`** — the final stats block, also printed to stdout.

## Rolling back

If you need to revert, the rollback log has every `(record_id, old_value)`
pair you need. A minimal one-off rollback script would look like:

```python
import csv, os, time, requests
from dotenv import load_dotenv
load_dotenv()
TOKEN = os.environ["HUBSPOT_PRIVATE_APP_TOKEN"]
URL = "https://api.hubapi.com/crm/v3/objects/companies/batch/update"
H = {"Authorization": f"Bearer {TOKEN}", "Content-Type": "application/json"}

with open("outputs/rollback_log_XXXXXX.csv") as f:
    rows = [r for r in csv.DictReader(f) if r["status"] == "OK"]

for i in range(0, len(rows), 100):
    batch = rows[i:i+100]
    requests.post(URL, headers=H, json={
        "inputs": [
            {"id": r["record_id"], "properties": {"top_ingredients": r["old_value"]}}
            for r in batch
        ]
    }).raise_for_status()
    time.sleep(0.2)
```

Only `OK` rows need rolling back — `ERROR` rows were never written, and
`DRY_RUN` rows came from a dry-run preview.

---

## Notes on rate / batch behavior

- The Companies batch update endpoint accepts up to **100 records per call**.
- Sleep between batches is configurable via `HUBSPOT_DELAY` (default `0.2`s).
  HubSpot Private App default rate limit is 100 requests per 10 seconds, so
  0.2s between batches is conservative.
- HubSpot may return HTTP `207` (multi-status) — partial success. The script
  handles this and reports per-record status in the rollback log.

## Common errors

- **`OBJECT_NOT_FOUND`** — the record was deleted between the time you
  exported and the time you ran the script. The script skips deleted
  domains *up front* (they aren't in the export), but if a record is
  deleted in the gap, you'll see this in the rollback log status.
- **`PROPERTY_DOESNT_EXIST`** — your custom property isn't named
  `top_ingredients` in HubSpot. Set `HUBSPOT_TOP_INGREDIENTS_FIELD` in
  `.env` to its actual internal name.
- **Property value too long** — HubSpot text fields have length limits. If
  this fires, consider switching the property type to "Multi-line text" in
  HubSpot, or pre-truncate.
