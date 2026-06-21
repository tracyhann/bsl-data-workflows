# Clean MRI Session Workflow

This workflow converts Flywheel MRI session CSV exports into a cleaned,
study-folder workbook that can line up with the REDCap visit dictionary.

## Philosophy

MRI session exports are useful indexing metadata, but they should not invent
study timepoints by themselves. This workflow uses the REDCap-derived cleaned
study folder as the authority for subject IDs and visit labels.

Rules we implement:

- **Preserve the raw imaging export.** The `raw` sheet is copied directly from
  the input CSV.
- **Standardize only the indexing fields.** The cleaned sheet creates a stable
  leftmost index using `IRB`, `subid`, `arm`, `visit`, and `date`.
- **Use REDCap timepoints to verify visits.** MRI `session.label` can hint at a
  visit, but the script only fills `visit` when the standardized MRI date falls
  inside a known subject timepoint range and the session label agrees with that
  visit.
- **Leave uncertain visits blank.** A blank visit is preferable to a confident
  wrong visit.
- **Keep the workbook self-contained.** The workbook includes raw MRI rows, the
  REDCap timepoint dictionary, and the cleaned MRI session view.

## Layout

- `run.py`: CLI and implementation for MRI session cleaning.
- `test_clean_mri_session.py`: unit tests for subject/date/visit matching.

## Primary Inputs

The input CSV must contain these columns:

```text
subject.label, session.label, session.timestamp, session.timezone, session.url, errors
```

The study folder should already contain:

```text
<study-folder>/data/cleaned/dictionary.xlsx
<study-folder>/data/cleaned/subjects/subject_timepoints.xlsx
```

If `subject_timepoints.xlsx` is missing, the script will build it from the
cleaned instrument workbooks.

## Copy-Paste Command Template

Set paths:

```bash
STUDY_FOLDER="studies/EXAMPLE_STUDY"
MRI_SESSION_CSV="$STUDY_FOLDER/data/raw_exports/imaging/mri_session.csv"
MRI_SESSION_OUT="$STUDY_FOLDER/data/cleaned/neuroimaging/mri_session.xlsx"
```

Run:

```bash
python3 scripts/workflows/clean_mri_session/run.py \
  --csv "$MRI_SESSION_CSV" \
  --study-folder "$STUDY_FOLDER" \
  --out "$MRI_SESSION_OUT"
```

## Outputs

Default output:

```text
<study-folder>/data/cleaned/neuroimaging/mri_session.xlsx
```

Sheets:

- `raw`: exact MRI session CSV values.
- `timepoint_dictionary`: copy of the REDCap event dictionary.
- `cleaned`: standardized MRI session index.

The `cleaned` sheet columns are:

```text
IRB, subid, arm, visit, date, subject.label, session.label, session.url, errors, session.timestamp, session.timezone
```

## Human Responsibilities

- Review rows with blank `visit`; those rows could not be confidently aligned
  to a REDCap timepoint.
- Check that `session.label` naming conventions are consistent enough for visit
  matching.
- Confirm that the REDCap subject timepoints are up to date before using MRI
  visit labels for downstream analysis.

## Tests

Run:

```bash
python3 -m unittest scripts.workflows.clean_mri_session.test_clean_mri_session
```
