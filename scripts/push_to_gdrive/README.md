# Push To Google Drive

Utilities for pushing local workbook artifacts into Google Drive/Sheets.

## Philosophy

- Treat local `.xlsx` files as the source of truth.
- Preserve the target Google Sheet URL when updating a prepared template.
- Prefer whole-workbook replacement for large instrument workbooks; this avoids slow, brittle cell-by-cell writes and correctly expands large sheets.
- Verify the target URL with `--dry-run` before uploading.

## Push Instrument Workbook

`push_instrument_workbook.py` replaces an existing native Google Sheet with a local instrument workbook while preserving the Google Sheet file ID and URL.

Run with:

```bash
python3 scripts/push_to_gdrive/push_instrument_workbook.py \
  --target "https://docs.google.com/spreadsheets/d/SPREADSHEET_ID/edit" \
  --workbook studies/STUDY/data/cleaned/assessments/IRB-instrument.xlsx
```

Dry run first:

```bash
python3 scripts/push_to_gdrive/push_instrument_workbook.py \
  --target "https://docs.google.com/spreadsheets/d/SPREADSHEET_ID/edit" \
  --workbook studies/STUDY/data/cleaned/assessments/IRB-instrument.xlsx \
  --dry-run
```

The `--target` value may be:

- a Google Sheet URL
- a raw Google spreadsheet ID
- a local text file containing either one

## Authentication

The scripts need a Google OAuth access token. `push_instrument_workbook.py` needs Drive write access; `fill_in_overview.py` needs Sheets write access.

Use one of:

```bash
GOOGLE_OAUTH_ACCESS_TOKEN="ya29..." python3 scripts/push_to_gdrive/push_instrument_workbook.py \
  --target "https://docs.google.com/spreadsheets/d/SPREADSHEET_ID/edit" \
  --workbook studies/STUDY/data/cleaned/assessments/IRB-instrument.xlsx
```

or:

```bash
python3 scripts/push_to_gdrive/push_instrument_workbook.py \
  --target "https://docs.google.com/spreadsheets/d/SPREADSHEET_ID/edit" \
  --workbook studies/STUDY/data/cleaned/assessments/IRB-instrument.xlsx \
  --access-token "ya29..."
```

If `--access-token` and `GOOGLE_OAUTH_ACCESS_TOKEN` are omitted, the script tries:

```bash
gcloud auth print-access-token
gcloud auth application-default print-access-token
```

The token must have permission to update the target Drive file.

## Fill In Overview

`fill_in_overview.py` writes a local study overview workbook into an existing Google Sheet by matching:

- workbook sheet name to Google Sheet tab name
- workbook column header to Google Sheet column header

Only matched columns are written. Extra Google Sheet columns are preserved.

Run with:

```bash
python3 scripts/push_to_gdrive/fill_in_overview.py \
  --target "https://docs.google.com/spreadsheets/d/SPREADSHEET_ID/edit" \
  --overview-file studies/STUDY/overview/STUDY.xlsx
```

Dry run first:

```bash
python3 scripts/push_to_gdrive/fill_in_overview.py \
  --target "https://docs.google.com/spreadsheets/d/SPREADSHEET_ID/edit" \
  --overview-file studies/STUDY/overview/STUDY.xlsx \
  --dry-run
```

By default, matched target columns are cleared from row 2 down before new values are written, which prevents stale rows from lingering. To only write over the necessary cells:

```bash
python3 scripts/push_to_gdrive/fill_in_overview.py \
  --target "https://docs.google.com/spreadsheets/d/SPREADSHEET_ID/edit" \
  --overview-file studies/STUDY/overview/STUDY.xlsx \
  --no-clear-existing
```

## Push Data Map

`push_data_map.py` writes local `*-data-map.xlsx` workbooks into an existing Google Sheet by matching:

- data-map workbook filename to Google Sheet tab name
- workbook column header to Google Sheet column header

For example, `assessments-data-map.xlsx` writes into an `assessments` tab, and `platforms-data-map.xlsx` writes into a `platforms` tab. Matching ignores case, spaces, underscores, and punctuation.

Run for a study folder:

```bash
python3 scripts/push_to_gdrive/push_data_map.py \
  --target "https://docs.google.com/spreadsheets/d/SPREADSHEET_ID/edit" \
  --study-folder studies/STUDY
```

Run for a specific data-map directory:

```bash
python3 scripts/push_to_gdrive/push_data_map.py \
  --target "https://docs.google.com/spreadsheets/d/SPREADSHEET_ID/edit" \
  --data-map-dir studies/STUDY/data-map
```

Run selected files only:

```bash
python3 scripts/push_to_gdrive/push_data_map.py \
  --target "https://docs.google.com/spreadsheets/d/SPREADSHEET_ID/edit" \
  --data-map-file studies/STUDY/data-map/platforms-data-map.xlsx \
  --data-map-file studies/STUDY/data-map/assessments-data-map.xlsx
```

Dry run first:

```bash
python3 scripts/push_to_gdrive/push_data_map.py \
  --target "https://docs.google.com/spreadsheets/d/SPREADSHEET_ID/edit" \
  --study-folder studies/STUDY \
  --dry-run
```

## Caveat

`push_instrument_workbook.py` is a workbook replacement. It preserves the target file URL/ID, but the imported workbook contents can replace prior Google-specific table formatting objects.

`fill_in_overview.py` is a range writer. It preserves tabs and extra columns, but only writes into tabs and columns whose names match the overview workbook.

`push_data_map.py` is also a range writer. It preserves tabs and extra columns, but only writes into tabs that match local `*-data-map.xlsx` file names and columns that match workbook headers.

## Test

```bash
python3 -m unittest \
  scripts.push_to_gdrive.test_push_instrument_workbook \
  scripts.push_to_gdrive.test_fill_in_overview \
  scripts.push_to_gdrive.test_push_data_map
```
