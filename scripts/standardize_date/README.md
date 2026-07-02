# Standardize Date

This tool detects date-like values in one selected CSV, TSV, or Excel column and
writes a CSV with standardized date annotations.

## Philosophy

Date standardization should be strict enough to avoid fake dates, but flexible
enough to handle common REDCap and spreadsheet formats.

Rules we implement:

- **Process one chosen column at a time.** This avoids guessing across unrelated
  columns.
- **Standardize valid dates to `YYYY-MM-DD`.**
- **Reject impossible dates.** Values like `2026-99-99` are not silently turned
  into real dates.
- **Preserve the original table.** Output appends annotation columns instead of
  replacing the source value.
- **Record parse status and match span.** Downstream scripts can audit what text
  was recognized and where it came from.

## Layout

- `run.py`: CLI for applying date recognition to one table column.
- `date_standardizer.py`: reusable date parsing and standardization functions.

## Copy-Paste Command Template

Set paths:

```bash
INPUT_TABLE="studies/EXAMPLE_STUDY/data/raw_exports/redcap/all/example_DATA.csv"
DATE_OUT="/tmp/example_date_standardized.csv"
```

Run by column name:

```bash
python3 scripts/standardize_date/run.py \
  "$INPUT_TABLE" \
  --column-name "visit_date" \
  --out "$DATE_OUT"
```

Run by zero-based column index:

```bash
python3 scripts/standardize_date/run.py \
  "$INPUT_TABLE" \
  --column-index 4 \
  --out "$DATE_OUT"
```

Run against an Excel sheet:

```bash
INPUT_XLSX="studies/EXAMPLE_STUDY/data/cleaned/assessments/example.xlsx"
INPUT_SHEET="cleaned"

python3 scripts/standardize_date/run.py \
  "$INPUT_XLSX" \
  --sheet "$INPUT_SHEET" \
  --column-name "date" \
  --out "$DATE_OUT"
```

## Outputs

The output CSV contains the original columns plus:

```text
date_detected, date_standardized, date_match, date_parse_status, date_score, date_start, date_end
```

If no valid date is detected, `date_detected` is `False` and
`date_standardized` is `NA`.

## Human Responsibilities

- Choose the correct date column.
- Review low-confidence or unexpected parse statuses before using dates for
  subject timepoint matching.
- Keep invalid placeholders as missing values rather than forcing them into a
  calendar date.

## Tests

Run:

```bash
python3 -m unittest scripts.standardize_date.test_standardize_date
```
