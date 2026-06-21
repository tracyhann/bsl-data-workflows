# REDCap Event Discovery Experiment

This experiment builds a dataset for testing REDCap event arm discovery and event grouping.

## Generate Data

```bash
python3 scripts/discover_events/dev/experiments/generate_data.py
```

Default output:

```text
scripts/discover_events/dev/experiments/data.csv
```

Columns:

```text
raw,arm,event_name,source
```

Rules:

- Arm variants such as `arm2`, `arm 2`, `arm 2 a`, `arm 2a`, `arm_2`, and `arm_2fxyz` are normalized to arm `2`.
- REDCap label forms such as `Screening (Visit 1) (Arm 1: Screening)` are parsed to arm `1` and event `Screening (Visit 1)`.
- REDCap unique-name forms such as `screening_visit_1_arm_1` are parsed to arm `1` and event `screening_visit_1`.

## Apply Discovery

```bash
python3 scripts/discover_events/run.py input.csv --out-dir event_discovery
```

By default, the runner uses column index `1`, which matches raw REDCap exports with `Event Name` or `redcap_event_name` next to the record ID column. If a known event column name appears elsewhere, it is used automatically.

Outputs:

```text
*_event_values.csv
*_events_by_arm.csv
```

`*_events_by_arm.csv` is sorted by first appearance in the original input, with event order counted separately within each arm.
