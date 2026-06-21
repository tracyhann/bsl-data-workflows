# Clean REDCap Instruments Workflow

This workflow converts REDCap exports into a study-folder package of cleaned,
auditable instrument workbooks.

## Philosophy

This workflow is intentionally conservative. It should make REDCap data easier
to use without hiding, deleting, or overwriting the evidence trail.

Rules we implement:

- **Raw exports remain the source of truth.** Discovery and cleaning should be
  traceable back to the DATA CSV, DATA_LABELS CSV, and codebook.
- **Index every cleaned sheet the same way.** Every cleaned instrument workbook
  starts with the same leftmost index columns:

```text
IRB, subid, arm, visit, date, record_id, redcap_event_name
```

- **Do not overwrite study data.** Standardized fields are added as index
  columns; raw values are preserved in `raw`, `raw_labels`, and
  `excluded_rows`.
- **Make each instrument workbook self-contained.** Every workbook carries the
  instrument's raw data, label-version raw data, cleaned sheet, timepoint
  dictionary, column dictionary, and excluded rows.
- **Use the labels export for human-readable values.** When a matching
  DATA_LABELS CSV is supplied, cleaned instrument values use labels where
  possible.
- **Keep discovery deterministic, but allow human review.** The pipeline creates
  `dictionary.xlsx`, then humans may edit abbreviations or labels before the
  cleaning stage.
- **Exclude sensitive and administrative material before organization.** Contact
  fields, consent/admin forms, deprecated forms, and empty instruments are moved
  out before final cleaned folders are produced.
- **Never silently trust the codebook over the export.** Codebook checks are
  advisory; the raw export is prioritized when they disagree.
- **Log every major decision.** Discovery, cleaning, exclusions, stale workbook
  deletion, sorting, and final verification all write to the dated history log.

## Layout

- `run.py`: wrapper for the staged workflow.
- `steps/`: individual workflow steps used by the wrapper.
- `test/`: unit and integration tests for this workflow.

## Primary Inputs

The wrapper expects a study folder plus REDCap export files:

- `--study-folder`: root folder for one study.
- `--irb`: study/IRB identifier used for standardized IDs and filenames.
- `--data-csv`: REDCap DATA export.
- `--labels-csv`: matching REDCap DATA_LABELS export.
- `--codebook`: REDCap codebook workbook.

Discovery copies these files into:

```text
<study-folder>/data/raw_exports/redcap/all/
```

Later stages infer the inputs from that location when explicit file paths are not
provided.

## Copy-Paste Command Template

Set these variables once per study, then run the stage commands below:

```bash
STUDY_FOLDER="studies/EXAMPLE_STUDY"
IRB="12345"
DATA_CSV="studies/EXAMPLE_STUDY/data/raw_exports/redcap/all/example_DATA.csv"
LABELS_CSV="studies/EXAMPLE_STUDY/data/raw_exports/redcap/all/example_DATA_LABELS.csv"
CODEBOOK="studies/EXAMPLE_STUDY/data/raw_exports/redcap/all/example-Codebook.xlsx"
```

## Streamlined Process With Human Checkpoints

Use this as the human-in-the-loop runbook. Automation proposes structure; humans
approve the places where study meaning, de-identification risk, or category
judgment matter.

| Step | Run or inspect | Human checkpoint |
| --- | --- | --- |
| 0. Raw export intake | Confirm DATA CSV, DATA_LABELS CSV, codebook, IRB, and study folder | Make sure exports are current, paired correctly, and not partial exports unless intentionally scoped. |
| 1. Discovery | `--stage discovery` | Review subject ID patterns, event/arm discovery, instrument discovery, and codebook mismatch logs. |
| 2. Dictionary review | `data/cleaned/redcap/dictionary.xlsx` | Fill or correct event abbreviations such as `V1`, `V2`, `V8`; review event labels, instrument labels, and suspicious subject variants. |
| 3. Classification review | `histories/YYYY-MM-DD/instrument_classification.xlsx` | Review low-confidence or `unknown` instruments; update `keywords.json` or `semantics.json` when a recurring instrument lands in the wrong class. |
| 4. Exclusion review | Clean-stage log sections and excluded instruments/columns | Confirm admin/contact/consent/deprecated/not-used exclusions are correct and no scientific variable was dropped by keyword collision. |
| 5. Clean instrument workbooks | `--stage clean` | Spot-check `raw`, `raw_labels`, `cleaned`, `column_variable_dictionary`, and `excluded_rows` for representative instruments. |
| 6. Postprocess organization | `--stage postprocess` | Confirm workbooks are sorted into the right `data/cleaned/<class>/` folders and `unknown/` is empty or understood. |
| 7. Final verification | `steps/final_verify.py` | Resolve failures. Read warnings, especially missing abbreviations or category mismatches. |
| 8. Subject timepoints | `scripts/subject_timepoints/run.py` | Review unusually wide date spans, missing visits, merged IRBs, and visit ordering. |
| 9. Imaging/session cleaning | `scripts/workflows/clean_mri_session/run.py`, when applicable | Review rows with blank visits, scan-date/timepoint mismatches, and subject ID matches. |
| 10. Study overview | `scripts/create_study_overview/run.py` | Review description quality, collection methods, and timepoint counts. |
| 11. Data maps | `scripts/create_data_maps/run.py` | Fill raw/platform placeholders and confirm relative file locations. |
| 12. Google template publishing | Existing Google Sheets tables, if used | Confirm target folder/template, replace-vs-append behavior, header names, and table ranges before writing. |

## Stages

### 1. Discovery

Run with:

```bash
python3 scripts/workflows/clean_redcap_instruments/run.py \
  --study-folder "$STUDY_FOLDER" \
  --stage discovery \
  --irb "$IRB" \
  --data-csv "$DATA_CSV" \
  --labels-csv "$LABELS_CSV" \
  --codebook "$CODEBOOK"
```

Responsibilities:

- Standardize the primary REDCap record ID column.
- Discover REDCap events from the event column.
- Discover instruments from column blocks ending in `*_complete`.
- Check discovered events and instruments against the supplied codebook.
- Create the initial dictionary workbook.
- Write discovery logs and intermediate CSVs.

Primary outputs:

```text
<study-folder>/data/cleaned/redcap/dictionary.xlsx
<study-folder>/histories/YYYY-MM-DD/log.md
<study-folder>/histories/YYYY-MM-DD/*_record_id_*.csv
<study-folder>/histories/YYYY-MM-DD/*_events_*.csv
<study-folder>/histories/YYYY-MM-DD/*_instrument_*.csv
```

Human responsibility after discovery:

- Review `dictionary.xlsx`.
- Fill or correct event abbreviations such as `V1`, `V2`, `V8`.
- Review instrument labels and category-sensitive assumptions.
- Resolve codebook mismatches when they reflect real export problems.

### 2. Clean

Run with:

```bash
python3 scripts/workflows/clean_redcap_instruments/run.py \
  --study-folder "$STUDY_FOLDER" \
  --stage clean \
  --irb "$IRB"
```

Responsibilities:

- Exclude instruments that are administrative, consent/contact-related,
  deprecated, or have no event coverage.
- Build one workbook per retained instrument.
- Prefer label-export values in human-facing sheets when a labels CSV is
  available.
- Standardize the leftmost cleaned columns:

```text
IRB, subid, arm, visit, date, record_id, redcap_event_name
```

- Drop or flag sensitive columns such as email, phone, address, contact, and MRN.
- Drop fully empty data columns and sparse incomplete rows when configured.

Each instrument workbook contains:

- `raw`
- `raw_labels`
- `cleaned`
- `timepoint_dictionary`
- `column_variable_dictionary`
- `excluded_rows`

Primary output:

```text
<study-folder>/data/cleaned/redcap/<IRB>-<instrument>.xlsx
```

Human responsibility after cleaning:

- Review excluded instruments and excluded rows.
- Inspect `column_variable_dictionary` when a needed variable is unexpectedly
  dropped.
- Confirm that visit abbreviations in the dictionary produce the intended cleaned
  `visit` values.

### 3. Postprocess

Run with:

```bash
python3 scripts/workflows/clean_redcap_instruments/run.py \
  --study-folder "$STUDY_FOLDER" \
  --stage postprocess \
  --irb "$IRB"
```

Responsibilities:

- Delete stale instrument workbooks whose `cleaned` sheet has no data rows.
- Classify retained instruments into categories.
- Move workbooks from the temporary REDCap folder into category folders.
- Promote the dictionary to the cleaned root.
- Remove the temporary `data/cleaned/redcap/` folder after it is emptied.

Primary outputs:

```text
<study-folder>/data/cleaned/dictionary.xlsx
<study-folder>/data/cleaned/subjects/
<study-folder>/data/cleaned/assessments/
<study-folder>/data/cleaned/treatments/
<study-folder>/data/cleaned/neuroimaging/
<study-folder>/data/cleaned/biologics_biometrics/
<study-folder>/data/cleaned/safety_regulatory/
```

Human responsibility after postprocess:

- Review `unknown/` if present.
- Update classification keywords or semantics when a recurring instrument lands
  in the wrong folder.
- Rerun postprocess after classification rule updates.

## Step Modules

- `steps/create_redcap_index_dictionary.py`: builds the subject, event, and
  instrument dictionary sheets from discovery outputs.
- `steps/discover_and_standardize.py`: coordinates record ID, event, and
  instrument discovery and writes the discovery log section.
- `steps/match_text_labels.py`: fills event and instrument labels from the
  labels export and verifies against the codebook.
- `steps/create_instrument_excels.py`: builds per-instrument workbooks and
  cleaned sheets.
- `steps/exclude_column.py`: flags sensitive columns to remove from exported
  workbooks.
- `steps/exclude_instruments.py`: excludes admin, consent/contact, deprecated,
  and zero-event instruments.
- `steps/drop_stale_instruments.py`: removes empty cleaned instrument workbooks.
- `steps/organize_instruments.py`: moves retained instrument workbooks into
  category folders and promotes `dictionary.xlsx`.
- `steps/final_verify.py`: verifies the organized cleaned output structure.

## Final Verification

Run final verification after postprocess and before regenerating subject
timepoints:

```bash
python3 scripts/workflows/clean_redcap_instruments/steps/final_verify.py \
  --study-folder "$STUDY_FOLDER"
```

It writes:

```text
<study-folder>/histories/YYYY-MM-DD/final_verification.xlsx
```

Verification fails loudly for structural/data-integrity problems and warns for
advisory issues such as missing abbreviations.

After final verification passes, run `scripts/subject_timepoints/run.py` to
refresh the cross-instrument subject timepoint summary.

## Typical Full Sequence

Set paths:

```bash
STUDY_FOLDER="studies/EXAMPLE_STUDY"
IRB="12345"
DATA_CSV="studies/EXAMPLE_STUDY/data/raw_exports/redcap/all/example_DATA.csv"
LABELS_CSV="studies/EXAMPLE_STUDY/data/raw_exports/redcap/all/example_DATA_LABELS.csv"
CODEBOOK="studies/EXAMPLE_STUDY/data/raw_exports/redcap/all/example-Codebook.xlsx"
```

Run all stages:

```bash
python3 scripts/workflows/clean_redcap_instruments/run.py \
  --study-folder "$STUDY_FOLDER" \
  --stage discovery \
  --irb "$IRB" \
  --data-csv "$DATA_CSV" \
  --labels-csv "$LABELS_CSV" \
  --codebook "$CODEBOOK"

python3 scripts/workflows/clean_redcap_instruments/run.py \
  --study-folder "$STUDY_FOLDER" \
  --stage clean \
  --irb "$IRB"

python3 scripts/workflows/clean_redcap_instruments/run.py \
  --study-folder "$STUDY_FOLDER" \
  --stage postprocess \
  --irb "$IRB"

python3 scripts/workflows/clean_redcap_instruments/steps/final_verify.py \
  --study-folder "$STUDY_FOLDER"

python3 scripts/subject_timepoints/run.py \
  --study-folder "$STUDY_FOLDER"
```
