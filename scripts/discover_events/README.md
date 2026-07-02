# Discover REDCap Events

This tool discovers REDCap event names and arm groupings from a raw DATA or
DATA_LABELS export.

## Philosophy

Event discovery should be driven by the export itself. The codebook is useful
for verification, but the raw export is the true source for what was actually
exported.

Rules we implement:

- **Use the event column only.** By default this is column index `1`, usually
  `redcap_event_name`, `event_name`, or `Event Name`.
- **Normalize arm spelling.** Variants such as `arm2`, `Arm 2`, `arm_2`, and
  longer arm suffixes are grouped into the same arm when the numeric arm is the
  same.
- **Preserve row order.** Event order within each arm follows first appearance
  in the raw export.
- **Keep raw event text.** Output rows retain the first raw event value so a
  human can check parsing decisions.
- **Use codebook verification as an audit.** Missing or extra events are logged
  when a codebook sheet is supplied.

## Layout

- `run.py`: CLI for reading CSV, TSV, XLSX, or XLS inputs.
- `event_discovery.py`: parsing and grouping logic.

## Copy-Paste Command Template

Set paths:

```bash
INPUT_CSV="studies/EXAMPLE_STUDY/data/raw_exports/redcap/all/example_DATA.csv"
OUT_DIR="studies/EXAMPLE_STUDY/histories/YYYY-MM-DD"
```

Run without codebook verification:

```bash
python3 scripts/discover_events/run.py \
  "$INPUT_CSV" \
  --out-dir "$OUT_DIR" \
  --column-name "redcap_event_name"
```

Run with codebook verification:

```bash
CODEBOOK="studies/EXAMPLE_STUDY/data/raw_exports/redcap/all/example-Codebook.xlsx"
EVENT_CODEBOOK_SHEET="events"

python3 scripts/discover_events/run.py \
  "$INPUT_CSV" \
  --out-dir "$OUT_DIR" \
  --column-name "redcap_event_name" \
  --codebook "$CODEBOOK" \
  --codebook-sheet "$EVENT_CODEBOOK_SHEET"
```

If the event column has a different name or only an index is known:

```bash
python3 scripts/discover_events/run.py \
  "$INPUT_CSV" \
  --out-dir "$OUT_DIR" \
  --column-index 1
```

## Outputs

For an input named `example_DATA.csv`, the output directory contains:

```text
example_DATA_event_values.csv
example_DATA_events_by_arm.csv
example_DATA_event_codebook_verification.csv
```

`event_codebook_verification.csv` is only written when `--codebook` and
`--codebook-sheet` are supplied.

## Human Responsibilities

- Confirm that the event column selected is the true REDCap event column.
- Review codebook mismatches. Some mismatches are expected when a form exists in
  the codebook but is not present in the export.
- Fill study-specific event abbreviations later in `dictionary.xlsx`, such as
  `V1`, `V2`, and `V8`.

## Tests

Run:

```bash
python3 -m unittest scripts.discover_events.test_discover_events
```
