# Subject Timepoints

This tool summarizes observed dates across all cleaned instrument workbooks in a
study folder.

## Philosophy

Subject timepoints are built from cleaned study data, not from a single form.
The output is an index that helps identify visit date ranges and supports later
modalities such as MRI session matching.

Rules we implement:

- **Read every cleaned instrument workbook.** Category folders under
  `data/cleaned/` are scanned.
- **Use the standardized leftmost index.** Rows are grouped by `subid`, `arm`,
  and `visit`, while IRB values are combined when the same subject appears under
  multiple IRBs in a merged study.
- **Summarize date ranges.** Each timepoint records earliest date, latest date,
  span in days, and all observed dates.
- **Trace sources only when useful.** Source paths are written for earliest and
  latest dates when there is a nonzero date span.
- **Keep this as an audit table.** The output should make date disagreement
  visible rather than resolving it silently.

## Layout

- `run.py`: CLI and workbook scanner.

## Required Input Structure

The study folder should contain cleaned instrument workbooks under category
folders:

```text
<study-folder>/data/cleaned/<category>/<IRB>-<instrument>.xlsx
```

Each workbook must have a `cleaned` sheet with at least:

```text
IRB, subid, arm, visit, date
```

## Copy-Paste Command Template

Set paths:

```bash
STUDY_FOLDER="studies/EXAMPLE_STUDY"
TIMEPOINTS_OUT="$STUDY_FOLDER/data/cleaned/subjects/subject_timepoints.xlsx"
```

Run:

```bash
python3 scripts/subject_timepoints/run.py \
  --study-folder "$STUDY_FOLDER" \
  --out "$TIMEPOINTS_OUT"
```

## Outputs

Default output:

```text
<study-folder>/data/cleaned/subjects/subject_timepoints.xlsx
```

Sheet:

```text
subject_timepoints
```

Columns:

```text
IRB, subid, arm, visit, earliest_entry_date, earliest_date_source, latest_entry_date, latest_date_source, span, values
```

Rows are sorted by:

```text
subid, arm, earliest_entry_date, visit
```

## Human Responsibilities

- Review large `span` values. They often indicate that one visit has dates
  spread across multiple forms or modalities.
- Use the source columns to investigate disagreements when `span` is greater
  than zero.
- Re-run this tool after merging studies or regenerating cleaned instruments.

## Tests

Run:

```bash
python3 -m unittest scripts.subject_timepoints.test_subject_timepoints
```
