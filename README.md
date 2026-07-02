# BSL Data Workflows

Reusable Python workflows for turning raw study exports into organized,
auditable study folders.

The repository is built around the study-folder pattern:

```text
studies/STUDY_NAME/
  data/
    raw_exports/
    cleaned/
  overview/
  data-map/
  histories/YYYY-MM-DD/
```

Most scripts are command-line tools under `scripts/`. The most common path is:

```text
raw REDCap export
  -> discover IDs, events, and instruments
  -> review dictionary
  -> clean one workbook per instrument
  -> classify and organize instruments
  -> final verification
  -> subject timepoints
  -> overview and data maps
  -> optional merge and Google Drive publishing
```

## Guiding Principles

- **Raw exports remain the source of truth.** Cleaned data should preserve a
  trail back to DATA CSV, DATA_LABELS CSV, and codebook inputs.
- **Stable leftmost indexes make workbooks comparable.** Cleaned REDCap
  instrument sheets start with:

```text
IRB, subid, arm, visit, date, record_id, redcap_event_name
```

- **Each instrument workbook is self-contained.** REDCap workbooks carry `raw`,
  `raw_labels`, `cleaned`, `timepoint_dictionary`,
  `column_variable_dictionary`, and `excluded_rows`.
- **Discovery is deterministic, review is human.** The pipeline proposes IDs,
  events, instruments, labels, classes, and exclusions; humans review the
  dictionary, abbreviations, unknown classifications, and warnings.
- **Sensitive material is excluded early.** Contact, phone, email, address, MRN,
  consent/admin, deprecated, and not-used instruments/columns are excluded from
  cleaned outputs.
- **Logs matter.** Major stages write to `histories/YYYY-MM-DD/log.md` and
  intermediate CSV/XLSX files so decisions can be audited.

## Install / Environment

The workflows are plain Python scripts. A typical environment needs:

```bash
python3 -m pip install pandas openpyxl matplotlib
```

Optional features may need extra packages:

```bash
python3 -m pip install sentence-transformers google-api-python-client google-auth google-auth-httplib2
```

Run commands from the repository root.

## End-To-End REDCap Cleaning

Set study-specific paths once:

```bash
STUDY_FOLDER="studies/EXAMPLE_STUDY"
IRB="12345"
DATA_CSV="$STUDY_FOLDER/data/raw_exports/redcap/all/example_DATA.csv"
LABELS_CSV="$STUDY_FOLDER/data/raw_exports/redcap/all/example_DATA_LABELS.csv"
CODEBOOK="$STUDY_FOLDER/data/raw_exports/redcap/all/example-Codebook.xlsx"
```

### 1. Discovery

Discovers subject IDs, events/arms, instruments, and creates the first
dictionary.

```bash
python3 scripts/workflows/clean_redcap_instruments/run.py \
  --study-folder "$STUDY_FOLDER" \
  --stage discovery \
  --irb "$IRB" \
  --data-csv "$DATA_CSV" \
  --labels-csv "$LABELS_CSV" \
  --codebook "$CODEBOOK"
```

Review after this step:

```text
<study-folder>/data/cleaned/redcap/dictionary.xlsx
<study-folder>/histories/YYYY-MM-DD/log.md
```

Human checkpoint:

- confirm subject ID patterns;
- review codebook mismatches;
- fill/correct event abbreviations such as `V1`, `V2`, `V8`;
- review instrument labels before cleaning.

### 2. Clean Instruments

Builds one workbook per retained instrument.

```bash
python3 scripts/workflows/clean_redcap_instruments/run.py \
  --study-folder "$STUDY_FOLDER" \
  --stage clean \
  --irb "$IRB"
```

Human checkpoint:

- inspect a few representative instrument workbooks;
- review excluded rows and excluded columns;
- confirm sensitive columns are absent from `raw`, `raw_labels`, and `cleaned`.

### 3. Postprocess

Drops stale workbooks, classifies instruments, moves workbooks into class
folders, and promotes the dictionary.

```bash
python3 scripts/workflows/clean_redcap_instruments/run.py \
  --study-folder "$STUDY_FOLDER" \
  --stage postprocess \
  --irb "$IRB"
```

Typical output:

```text
<study-folder>/data/cleaned/dictionary.xlsx
<study-folder>/data/cleaned/subjects/
<study-folder>/data/cleaned/assessments/
<study-folder>/data/cleaned/treatments/
<study-folder>/data/cleaned/neuroimaging/
<study-folder>/data/cleaned/biologics_biometrics/
<study-folder>/data/cleaned/safety_regulatory/
```

### 4. Final Verify

Checks that dictionaries, class folders, workbooks, required sheets, index
columns, classification locations, and excluded columns are consistent.

```bash
python3 scripts/workflows/clean_redcap_instruments/steps/final_verify.py \
  --study-folder "$STUDY_FOLDER"
```

Output:

```text
<study-folder>/histories/YYYY-MM-DD/final_verification.xlsx
```

Failures should be resolved. Warnings, such as missing abbreviations, should be
reviewed and documented.

### 5. Subject Timepoints

Summarizes observed dates across all cleaned instrument workbooks.

```bash
python3 scripts/subject_timepoints/run.py \
  --study-folder "$STUDY_FOLDER"
```

Output:

```text
<study-folder>/data/cleaned/subjects/subject_timepoints.xlsx
```

Human checkpoint:

- review large date spans;
- review blank dates;
- use source columns when a timepoint has conflicting dates.

## Optional Imaging Metadata

Clean a Flywheel MRI session CSV against the REDCap-derived subject timepoints.

```bash
MRI_SESSION_CSV="$STUDY_FOLDER/data/raw_exports/imaging/mri_session.csv"

python3 scripts/workflows/clean_mri_session/run.py \
  --csv "$MRI_SESSION_CSV" \
  --study-folder "$STUDY_FOLDER"
```

Output:

```text
<study-folder>/data/cleaned/neuroimaging/mri_session.xlsx
```

Rows with uncertain visit matching intentionally keep the `visit` cell blank.

## Overview And Data Maps

### Study Overview

Creates a concise study-facing overview workbook from cleaned class folders.

```bash
python3 scripts/create_study_overview/run.py \
  --study-folder "$STUDY_FOLDER"
```

Output:

```text
<study-folder>/overview/<study-folder-name>.xlsx
```

### Data Maps

Creates lightweight data-map workbooks for each cleaned class folder.

```bash
python3 scripts/create_data_maps/run.py \
  --study-folder "$STUDY_FOLDER"
```

Output:

```text
<study-folder>/data-map/platforms-data-map.xlsx
<study-folder>/data-map/<class>-data-map.xlsx
```

Human checkpoint:

- fill raw/platform placeholder rows;
- check descriptions and locations;
- review before pushing into Google Sheets templates.

## Merge Cleaned Studies

Merge two already-cleaned and verified study folders.

```bash
STUDY1="studies/12345-STUDY-A"
STUDY2="studies/56789-STUDY-B"
MERGED_STUDY="studies/12345-56789-STUDY"

python3 scripts/merge_studies/run.py \
  --study1 "$STUDY1" \
  --study2 "$STUDY2" \
  --out "$MERGED_STUDY"
```

The merged folder includes:

```text
<merged-study>/data/cleaned/
<merged-study>/data/cleaned/subjects/subject_timepoints.xlsx
<merged-study>/studies/<source-study-1>/
<merged-study>/studies/<source-study-2>/
```

Human checkpoint:

- inspect semicolon-separated conflicts;
- confirm matched instruments are truly comparable;
- rerun final verification on the merged folder.

## Visualize Data Volume

Create basic REDCap cleaned-data volume plots by participant and visit.

```bash
python3 scripts/visualize/subject_data_volume.py \
  --study-folder "$STUDY_FOLDER"
```

Plots are written to:

```text
<study-folder>/histories/YYYY-MM-DD/plots/
```

## Publish To Google Drive

After local cleaning, verification, overview, and data maps are ready, publish
the study folder into the Google Drive study-folder template.

```bash
python3 scripts/workflows/create_study_folder_gdrive/run.py \
  --stage all \
  --study-folder "$STUDY_FOLDER" \
  --study-name "Example Study" \
  --irb "$IRB" \
  --template "https://drive.google.com/drive/folders/TEMPLATE_FOLDER_ID" \
  --destination root
```

Useful options:

```bash
--existing-file-policy update-or-create
--share-sheet-editor person1@example.edu
--share-sheet-editor person2@example.edu
```

The publishing workflow:

- copies the Google Drive template;
- fills `IRB-meta`;
- writes the local overview into the Google overview sheet;
- uploads cleaned workbooks into class folders;
- writes data maps and replaces relative locations with Drive links;
- logs upload successes, failures, and permission warnings.

There is also a Colab notebook in the tools repo:

```text
https://github.com/tracyhann/bsl-data-tools/tree/main/create-study-folder-gdrive
```

## Tool Reference

| Tool | Purpose |
| --- | --- |
| [`scripts/workflows/clean_redcap_instruments`](scripts/workflows/clean_redcap_instruments/README.md) | Main staged REDCap discovery, cleaning, postprocess, and verification workflow. |
| [`scripts/standardize_record_id`](scripts/standardize_record_id/README.md) | REDCap subject ID discovery and standardization. |
| [`scripts/standardize_date`](scripts/standardize_date/README.md) | Date detection and YYYY-MM-DD standardization. |
| [`scripts/discover_events`](scripts/discover_events/README.md) | Event and arm discovery from REDCap event columns. |
| [`scripts/discover_instruments`](scripts/discover_instruments/README.md) | Instrument block discovery from REDCap columns. |
| [`scripts/classify_instruments`](scripts/classify_instruments/README.md) | Keyword/semantic instrument categorization. |
| [`scripts/subject_timepoints`](scripts/subject_timepoints/README.md) | Cross-instrument subject timepoint date summary. |
| [`scripts/workflows/clean_mri_session`](scripts/workflows/clean_mri_session/README.md) | MRI session CSV cleaning and REDCap timepoint matching. |
| [`scripts/merge_studies`](scripts/merge_studies/README.md) | Merge two cleaned study folders. |
| [`scripts/create_study_overview`](scripts/create_study_overview/README.md) | Study-facing overview workbook generation. |
| [`scripts/create_data_maps`](scripts/create_data_maps/README.md) | Data-map workbook generation. |
| [`scripts/push_to_gdrive`](scripts/push_to_gdrive/README.md) | Lower-level helpers for writing Excel/overview/data-map values into Google Sheets. |
| [`scripts/workflows/create_study_folder_gdrive`](scripts/workflows/create_study_folder_gdrive/README.md) | End-to-end Google Drive publishing workflow. |

## Tests

Run focused tests for the tool you changed. Common examples:

```bash
python3 -m unittest scripts.workflows.clean_redcap_instruments.test.test_workflows
python3 -m unittest scripts.classify_instruments.test_classify_instruments
python3 -m unittest scripts.subject_timepoints.test_subject_timepoints
python3 -m unittest scripts.merge_studies.test_merge_studies
python3 -m unittest scripts.workflows.create_study_folder_gdrive.test_create_study_folder_gdrive
```

## Human Review Checklist

Use this list before sharing or publishing a study folder:

- `dictionary.xlsx` has reviewed event abbreviations and instrument labels.
- `instrument_classification.xlsx` has no unexplained `unknown` classes.
- Excluded instruments and columns are expected.
- `final_verification.xlsx` has no unresolved failures.
- `subject_timepoints.xlsx` has reviewed wide spans and missing dates.
- Overview descriptions and collection methods are human-readable.
- Data-map locations are correct.
- Google Drive upload logs have no unresolved upload or permission warnings.
