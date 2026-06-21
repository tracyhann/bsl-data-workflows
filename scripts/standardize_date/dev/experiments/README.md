# Date Standardization Experiment

This experiment builds a two-column dataset for testing date detection and standardization.

## Generate Data

```bash
python3 scripts/standardize_date/dev/experiments/generate_data.py
```

Default output:

```text
scripts/standardize_date/dev/experiments/data.csv
```

Columns:

```text
raw,canonicalized
```

Rules:

- Full dates are standardized to `YYYY-MM-DD`, for example `2026-06-15`.
- Non-dates, invalid dates, and risky partial dates are labeled `NA`.
- U.S. numeric slash dates default to month/day/year.
- If a numeric slash date starts with a value greater than 12, it is treated as day/month/year.

## Apply Detector

```bash
python3 scripts/standardize_date/run.py input.csv --column-name visit_date --out output.csv
```

If no column name is supplied, the runner uses zero-based column index `0`.

For Excel:

```bash
python3 scripts/standardize_date/run.py input.xlsx --sheet Sheet1 --column-name "Visit Date" --out output.csv
```
