# Merge Studies

This tool merges two already-cleaned study folders into one combined cleaned
study folder.

## Philosophy

Merging happens after each study has already been cleaned and verified. The goal
is to combine comparable cleaned workbooks while preserving study provenance.

Rules we implement:

- **Merge cleaned outputs, not raw REDCap exports.** Each input study should
  already have `data/cleaned/dictionary.xlsx` and category folders.
- **Archive source studies.** The merged folder contains a `studies/` directory
  with copies of both original study folders.
- **Match instruments conservatively.** Instruments match by instrument name,
  then by instrument label when names differ.
- **Preserve columns from both studies.** Extra columns from either study are
  kept, with blanks filled where the other study lacks that column.
- **Merge same subject, arm, and visit rows carefully.** Equal values collapse;
  conflicting values are retained with semicolon-separated entries in study
  input order.
- **Keep the earliest indexing date.** When merged rows have multiple dates, the
  `date` column keeps the earliest date.
- **Regenerate subject timepoints.** After merging, subject timepoints are built
  for the merged folder.

## Layout

- `run.py`: CLI and merge implementation.

## Required Input Structure

Each input study folder should contain:

```text
<study-folder>/data/cleaned/dictionary.xlsx
<study-folder>/data/cleaned/<category>/<IRB>-<instrument>.xlsx
```

## Copy-Paste Command Template

Set paths:

```bash
STUDY1="studies/12345-STUDY-A"
STUDY2="studies/56789-STUDY-B"
MERGED_STUDY="studies/12345-56789-STUDY"
```

Run:

```bash
python3 scripts/merge_studies/run.py \
  --study1 "$STUDY1" \
  --study2 "$STUDY2" \
  --out "$MERGED_STUDY"
```

Overwrite an existing merged folder:

```bash
python3 scripts/merge_studies/run.py \
  --study1 "$STUDY1" \
  --study2 "$STUDY2" \
  --out "$MERGED_STUDY" \
  --overwrite
```

## Outputs

The merged folder contains:

```text
<merged-study>/data/cleaned/dictionary.xlsx
<merged-study>/data/cleaned/<category>/<IRB1>-<IRB2>-<instrument>.xlsx
<merged-study>/data/cleaned/subjects/subject_timepoints.xlsx
<merged-study>/studies/<source-study-1>/
<merged-study>/studies/<source-study-2>/
```

The command prints:

```text
<merged-study>
dictionary=<merged-study>/data/cleaned/dictionary.xlsx
merged_workbooks=<count>
copied_workbooks=<count>
archived_study_folders=2
```

## Human Responsibilities

- Inspect merged rows that contain semicolon-separated conflicts.
- Confirm that same-named instruments are truly comparable across studies.
- Re-run final verification on the merged folder after merging.
- Check `subject_timepoints.xlsx` for unexpected date spans.

## Tests

Run:

```bash
python3 -m unittest scripts.merge_studies.test_merge_studies
```
