# Discover REDCap Instruments

This tool discovers REDCap instrument blocks from column headers in a raw export.

## Philosophy

REDCap exports are wide tables where instruments appear as contiguous blocks of
columns. This tool uses those local column patterns instead of a hardcoded list
of form names.

Rules we implement:

- **Use column order as evidence.** Instruments are detected as contiguous
  blocks in the order they appear in the export.
- **Treat `*_complete` as a stop signal.** REDCap completion-status columns mark
  the end of an instrument block.
- **Use prefixes and suffixes as supporting evidence.** Shared prefixes such as
  `madrs_` help name and group nearby columns.
- **Do not require a codebook.** Discovery can run on the raw export alone.
- **Verify against the codebook when supplied.** The codebook tells us what was
  expected, while the export tells us what is actually present.

## Layout

- `run.py`: CLI for reading CSV, TSV, XLSX, or XLS inputs.
- `instrument_discovery.py`: instrument block discovery logic.
- `test_discover_instruments.py`: tests for instrument boundary detection.

## Copy-Paste Command Template

Set paths:

```bash
INPUT_CSV="studies/EXAMPLE_STUDY/data/raw_exports/redcap/all/example_DATA.csv"
OUT_DIR="studies/EXAMPLE_STUDY/histories/YYYY-MM-DD"
```

Run without codebook verification:

```bash
python3 scripts/discover_instruments/run.py \
  "$INPUT_CSV" \
  --out-dir "$OUT_DIR"
```

Run with codebook verification:

```bash
CODEBOOK="studies/EXAMPLE_STUDY/data/raw_exports/redcap/all/example-Codebook.xlsx"
INSTRUMENT_CODEBOOK_SHEET="instruments"

python3 scripts/discover_instruments/run.py \
  "$INPUT_CSV" \
  --out-dir "$OUT_DIR" \
  --codebook "$CODEBOOK" \
  --codebook-sheet "$INSTRUMENT_CODEBOOK_SHEET"
```

Run against an Excel sheet:

```bash
INPUT_XLSX="studies/EXAMPLE_STUDY/data/raw_exports/redcap/all/example_DATA.xlsx"
INPUT_SHEET="Sheet1"

python3 scripts/discover_instruments/run.py \
  "$INPUT_XLSX" \
  --sheet "$INPUT_SHEET" \
  --out-dir "$OUT_DIR"
```

## Outputs

For an input named `example_DATA.csv`, the output directory contains:

```text
example_DATA_instrument_summary.csv
example_DATA_instrument_columns.csv
example_DATA_instrument_codebook_verification.csv
```

`instrument_codebook_verification.csv` is only written when `--codebook` and
`--codebook-sheet` are supplied.

## Human Responsibilities

- Review instruments that are missing from discovery but present in the
  codebook. They may be absent from the export, or the export may not include
  those form columns.
- Review instruments that are extra in discovery. They may be real export
  content that the codebook sheet did not list.
- Confirm that discovered instrument names are suitable for filenames and
  dictionary labels.

## Tests

Run:

```bash
python3 -m unittest scripts.discover_instruments.test_discover_instruments
```
