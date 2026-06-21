# Create Study Folder In Google Drive

This workflow initializes a Google Drive study folder from a template and pushes
local cleaned study outputs into that folder.

## Philosophy

- Keep local cleaned outputs as the source of truth.
- Use native Google Drive copy operations for template files so Google-native
  layouts, sheets, and formatting are preserved as much as the Drive API allows.
- Fill Google Sheets by matching tab names and column names instead of relying
  on fixed coordinates.
- Upload cleaned files into the same class structure used locally under
  `data/cleaned`.
- Log the Google Drive folder URL, uploads, failures, and data-map status in the
  local study history log.
- Treat this as a publishing step, not a cleaning step. Run local verification
  first.

## Expected Local Inputs

```text
studies/STUDY/
├── data/cleaned/
├── data-map/
├── overview/
└── histories/YYYY-MM-DD/log.md
```

The workflow reads:

- local overview workbook from `overview/*.xlsx`
- local data-map workbooks from `data-map/*-data-map.xlsx`
- local cleaned outputs from `data/cleaned/**`

## Google Template Expectations

The Google Drive template folder should include:

- `IRB-meta` at the template root
- `Overview/STUDY_IRB`
- `Data (internal/approved-access)/No-PHI Data (internal/approved-access)/blank_templates/`
- `Data Map (internal/approved-access)/IRB-data-map`
- template sheets named `REDCap_INSTRUMENTS` and `BLANK` inside `blank_templates`

During initialization, filenames and folder names replace:

- `STUDY` -> supplied `--study-name`
- `IRB` -> supplied `--irb`

`IRB-meta` is treated as a reserved metadata filename and is not renamed.

## Authentication

Use one of:

```bash
export GOOGLE_OAUTH_ACCESS_TOKEN="ya29..."
```

or pass:

```bash
--access-token "ya29..."
```

If neither is provided, the script tries `gcloud auth print-access-token` and
`gcloud auth application-default print-access-token`.

## Stage 1: Initialize Google Drive Study Folder

This copies the template folder into the destination parent folder, renames
template placeholders, fills `IRB-meta`, and logs the initialized folder link.

For your template and personal Google Drive root:

```bash
python3 scripts/workflows/create_study_folder_gdrive/run.py \
  --stage initialize \
  --study-folder studies/53879-62882-OCD-TMS \
  --study-name "OCD-TMS" \
  --irb "53879-62882" \
  --template "https://drive.google.com/drive/folders/1OLGHgxPg9UsBDbOH6vgSjiS6dUW0lBCF?usp=sharing" \
  --destination root
```

If the destination should be a specific Drive folder:

```bash
python3 scripts/workflows/create_study_folder_gdrive/run.py \
  --stage initialize \
  --study-folder studies/53879-62882-OCD-TMS \
  --study-name "OCD-TMS" \
  --irb "53879-62882" \
  --template "https://drive.google.com/drive/folders/1OLGHgxPg9UsBDbOH6vgSjiS6dUW0lBCF?usp=sharing" \
  --destination "https://drive.google.com/drive/folders/DESTINATION_FOLDER_ID"
```

## Stage 2: Upload Overview And Cleaned Data

Use this after initialization. If you pass `--initialized-folder-id`, the script
tries to find the standard Overview, No-PHI Data, and template folders under it.

```bash
python3 scripts/workflows/create_study_folder_gdrive/run.py \
  --stage upload \
  --study-folder studies/53879-62882-OCD-TMS \
  --study-name "OCD-TMS" \
  --irb "53879-62882" \
  --initialized-folder-id "https://drive.google.com/drive/folders/INITIALIZED_STUDY_FOLDER_ID"
```

Override destinations when needed:

```bash
python3 scripts/workflows/create_study_folder_gdrive/run.py \
  --stage upload \
  --study-folder studies/53879-62882-OCD-TMS \
  --study-name "OCD-TMS" \
  --irb "53879-62882" \
  --overview-destination "https://docs.google.com/spreadsheets/d/OVERVIEW_SHEET_ID/edit" \
  --cleaned-data-folder-id "https://drive.google.com/drive/folders/NO_PHI_DATA_FOLDER_ID" \
  --templates-folder-id "https://drive.google.com/drive/folders/BLANK_TEMPLATES_FOLDER_ID"
```

Template choice:

- workbooks with `raw`, `raw_labels`, `cleaned`, `timepoint_dictionary`,
  `column_variable_dictionary`, and `excluded_rows` use `REDCap_INSTRUMENTS`
- other `.xlsx`, `.xlsm`, and `.csv` files use `BLANK`
- non-spreadsheet files are uploaded directly

## Stage 3: Fill Data Map

After cleaned files are uploaded, the workflow rewrites local data-map
`location` cells to the uploaded Google Drive links where available, then fills
the Google data-map sheet by matching tab and column names.

```bash
python3 scripts/workflows/create_study_folder_gdrive/run.py \
  --stage data-map \
  --study-folder studies/53879-62882-OCD-TMS \
  --study-name "OCD-TMS" \
  --irb "53879-62882" \
  --initialized-folder-id "https://drive.google.com/drive/folders/INITIALIZED_STUDY_FOLDER_ID" \
  --data-map-destination "https://docs.google.com/spreadsheets/d/DATA_MAP_SHEET_ID/edit"
```

## All-In-One

This runs initialize, overview/data upload, data-map fill, and extra local folder
upload in one process. The all-in-one mode can infer default Google locations
from the copied template folder.

```bash
python3 scripts/workflows/create_study_folder_gdrive/run.py \
  --stage all \
  --study-folder studies/53879-62882-OCD-TMS \
  --study-name "OCD-TMS" \
  --irb "53879-62882" \
  --template "https://drive.google.com/drive/folders/1OLGHgxPg9UsBDbOH6vgSjiS6dUW0lBCF?usp=sharing" \
  --destination root
```

## Human Review Checkpoints

1. Confirm the template folder is the intended Google Drive template.
2. Confirm `IRB-meta` was filled correctly.
3. Confirm overview tabs and columns were filled in the expected Google Sheet.
4. Confirm cleaned data files uploaded into the expected class folders.
5. Review upload failures in `histories/YYYY-MM-DD/log.md`.
6. Confirm data-map `location` cells point to Google Drive links.

## Tests

```bash
python3 -m unittest scripts.workflows.create_study_folder_gdrive.test_create_study_folder_gdrive
```
