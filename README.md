# Batch Start to Picked date — automated refresh

Rebuilds `Batch_Start_to_Picked_date.xlsx` twice a day (9am / 9pm IST) with no manual
CSV export/upload step: enrollment numbers come straight from Metabase (saved question
**12957 "Batch Status"**), and the "Picked" rows come straight from the FlyWheel Google
Sheet via the Sheets API — pulling only the 4 columns actually used, not the whole
82-column sheet (the "import just the important parts" version of IMPORTRANGE, since
IMPORTRANGE itself only works Sheet-to-Sheet, not into an .xlsx file).

Everything about *where* the data lives — Metabase URL, question id, FlyWheel sheet
id/tab name — is hardcoded at the top of `refresh_batch_pipeline.py`. Only two secrets
are needed to run it.

## Files
- `refresh_batch_pipeline.py` — the whole pipeline (Metabase query -> Sheets API pull -> xlsx build).
- `.github/workflows/refresh_batch_pipeline.yml` — GitHub Actions cron, 03:30 & 15:30 UTC (= 9am/9pm IST).
- `requirements.txt`

## One-time setup — 2 secrets only

Repo -> Settings -> Secrets and variables -> Actions -> New repository secret:

### 1. `METABASE_API_KEY`
Metabase -> Admin -> Settings -> Authentication -> API Keys -> create one with read
access to database id 4 (Newton School). Paste the key value as this secret.

### 2. `GOOGLE_SERVICE_ACCOUNT_JSON`
1. Google Cloud Console -> a project -> IAM & Admin -> Service Accounts -> Create.
2. Create a JSON key for it, download it.
3. Open the FlyWheel sheet -> Share -> paste the service account's `client_email`
   (looks like `xxx@yyy.iam.gserviceaccount.com`) -> Viewer access.
4. Paste the **entire contents of the downloaded JSON file** as this secret's value
   (GitHub secrets support multi-line text — no base64 encoding needed, the script
   detects and handles raw JSON automatically).

That's it — no other secrets to set.

## Hardcoded config (edit `refresh_batch_pipeline.py` directly if any of these change)
```python
METABASE_URL          = "https://metabase-lierhfgoeiwhr.newtonschool.co"
METABASE_DATABASE_ID  = 4
METABASE_QUESTION_ID  = 12957   # saved question "Batch Status"
FLYWHEEL_SHEET_ID      = "1Ue49enEEpgNaOEdQVgwgsehWvHb3HEI0Q-qekvAzYyU"
FLYWHEEL_TAB_NAME      = "Prog<>Placement"
```
`FLYWHEEL_TAB_NAME` was confirmed from the tab bar in a screenshot of the live sheet
(the tab selected/highlighted for gid=1147350782). If the first run still 400s with
something like "Unable to parse range", double-check that tab's exact label — Sheets
tab names are picky about exact characters/spacing.

## Column positions
The pipeline reads columns **A** (UserID), **D** (Batch), **U** (Status), **V** (Picked
Date) from the FlyWheel tab — matches the export this was built from (0-indexed cols
0, 3, 20, 21). If anyone reorders columns in that sheet, update the `ranges` list in
`fetch_picked_rows()`.

## What's NOT automated
- The **Batch (A/B)** column in the "Batch Start to Picked date" tab is manual by
  design — nothing in Metabase or the FlyWheel sheet backs an A/B split for these
  batches (checked `courses_subbatch`, the FlyWheel `Batch` column, and the original
  pivot's blank "A?B" column — none of them have it). This script always rebuilds the
  sheet from scratch, so it won't currently preserve hand-typed A/B values across a
  run — say if you want that preserved and I'll add a step that reads the previous
  output first and carries that column forward.

## Run locally to test
```bash
pip install -r requirements.txt
export METABASE_API_KEY=...
export GOOGLE_SERVICE_ACCOUNT_JSON="$(cat service-account.json)"
python refresh_batch_pipeline.py
```
