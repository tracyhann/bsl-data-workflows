# Standardize REDCap Record ID

This tool detects and standardizes REDCap subject IDs from one selected table
column.

## Philosophy

Record ID standardization should target the primary subject ID field first. A
wide REDCap export can contain many other IDs in notes, Flywheel links, or
cross-study references, and those should not be treated as the row's true
subject ID.

Rules we implement:

- **Process one selected column.** The default is column index `0`, usually
  `record_id` or `Record ID`.
- **Standardize IDs to `IRB_s0*`.** Examples: `12345_s025`,
  `12345s025`, and `12345_b025` can resolve to the same canonical style when
  they are valid subject IDs.
- **Reject unlikely record ID cells.** Very long text values are treated as
  unlikely primary record IDs.
- **Preserve raw values.** Output appends standardized ID columns instead of
  replacing source values.
- **Summarize naming formats.** A separate summary table reports coded format
  families and example values.

## Layout

- `run.py`: CLI for applying REDCap subject ID recognition to one table column.
- `presidio/redcap_subid_presidio.py`: reusable recognizer and canonicalizer.

## Copy-Paste Command Template

Set paths:

```bash
INPUT_TABLE="studies/EXAMPLE_STUDY/data/raw_exports/redcap/all/example_DATA.csv"
STANDARDIZED_OUT="/tmp/example_redcap_subid_standardized.csv"
SUMMARY_OUT="/tmp/example_record_id_format_summary.csv"
```

Run by column name:

```bash
python3 scripts/standardize_record_id/run.py \
  "$INPUT_TABLE" \
  --column-name "record_id" \
  --out "$STANDARDIZED_OUT" \
  --summary-out "$SUMMARY_OUT" \
  --summary-examples 5 \
  --random-seed 42
```

Run by zero-based column index:

```bash
python3 scripts/standardize_record_id/run.py \
  "$INPUT_TABLE" \
  --column-index 0 \
  --out "$STANDARDIZED_OUT" \
  --summary-out "$SUMMARY_OUT" \
  --summary-examples 5 \
  --random-seed 42
```

Run against an Excel sheet:

```bash
INPUT_XLSX="studies/EXAMPLE_STUDY/data/cleaned/dictionary.xlsx"
INPUT_SHEET="subject_id"

python3 scripts/standardize_record_id/run.py \
  "$INPUT_XLSX" \
  --sheet "$INPUT_SHEET" \
  --column-name "source_ID" \
  --out "$STANDARDIZED_OUT" \
  --summary-out "$SUMMARY_OUT" \
  --summary-examples 5 \
  --random-seed 42
```

## Outputs

The standardized output CSV contains the original columns plus:

```text
redcap_subid_detected, redcap_subid_canonicalized, redcap_subid_match, redcap_subid_entity_type, redcap_subid_score, redcap_subid_start, redcap_subid_end
```

The format summary CSV groups values into coded naming families and includes
random examples per family.

## Human Responsibilities

- Use the primary `record_id` column for subject indexing summaries.
- Review rows where the selected record ID does not standardize.
- Investigate any row that appears to contain more than one plausible subject
  ID.
- Keep the summary output with the run history so naming drift is auditable.

## Tests

Run:

```bash
python3 -m unittest scripts.standardize_record_id.test_standardize_record_id
```
