# REDCap Subject ID Standardization Experiment

This experiment builds a two-column dataset for testing REDCap subject ID detection and standardization.

## Generate Data

```bash
python3 scripts/standardize_record_id/dev/experiments/generate_data.py
```

Default output:

```text
scripts/standardize_record_id/dev/experiments/data.csv
```

Columns:

```text
raw,canonicalized
```

Rules:

- REDCap subject IDs are standardized to `IRB_s0*`, for example `58807_s025`.
- Non-subject IDs and distractor values are labeled `NA`.
- Suffix variants such as `58807_s025a` and `58807s025XOVER` map to the base subject ID.

## Apply Detector

```bash
python3 scripts/standardize_record_id/run.py input.csv --column-name record_id --out output.csv
```

If no column name is supplied, the runner uses zero-based column index `0`.

For Excel:

```bash
python3 scripts/standardize_record_id/run.py input.xlsx --sheet Sheet1 --column-name "Record ID" --out output.csv
```



subid identification and standardization (standardize to IRB_s0*)

date identification and standardization (YYYY-MM-DD) 

columns to instrument discovery and grouping

any columns that contain contact information or person's name should be dropped