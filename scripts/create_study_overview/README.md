# Create Study Overview

This workflow creates a concise study overview workbook from the cleaned study
folder. It summarizes which REDCap instruments are available, when they were
collected, and how they were collected.

## Philosophy

The overview is a human-facing summary, not a replacement for cleaned workbooks.

Rules we implement:

- **Only summarize REDCap instruments with matched labels.** Files without a
  dictionary instrument label are skipped so the overview stays study-facing.
- **Exclude intentionally unused or training material.** Items containing
  terms such as `trainee`, `do not use`, `not used`, or `deprecated` are not
  included.
- **Use cleaned sheets as the source for collected visits.** The overview reads
  the standardized indexing `visit` column from each instrument workbook.
- **Prefer clinical semantic definitions.** Descriptions come from
  `scripts/classify_instruments/semantics.json` when a match is available.
- **Do not infer beyond the cleaned data.** Collection timepoints are counted
  only from visits that actually appear in each cleaned workbook.
- **Make collection tools explicit.** REDCap instruments list the observed IRB
  values from the cleaned sheet as `REDCap (IRB: ...)`.

## Input

Required:

```text
<study-folder>/data/cleaned/
<study-folder>/data/cleaned/dictionary.xlsx
```

The dictionary must include the `instrument` sheet with matched
`instrument_label` values.

Optional:

```text
scripts/classify_instruments/semantics.json
```

The semantics file is used to fill the overview `description` column.

## Output

Default output:

```text
<study-folder>/overview/<study-folder-name>.xlsx
```

The workbook contains one sheet per cleaned data class:

```text
assessments
neuroimaging
treatments
biologics_biometrics
subjects
safety_regulatory
```

Each sheet has:

```text
data
description
# of timepoints (collected)
timepoints (collected)
collection method
collection tool
```

## Copy-Paste Commands

Run for one study:

```bash
python3 scripts/create_study_overview/run.py \
  --study-folder studies/54909-58807-BRAINS
```

Run with an explicit output workbook:

```bash
python3 scripts/create_study_overview/run.py \
  --study-folder studies/54909-58807-BRAINS \
  --out studies/54909-58807-BRAINS/overview/54909-58807-BRAINS.xlsx
```

Run with a custom semantics file:

```bash
python3 scripts/create_study_overview/run.py \
  --study-folder studies/54909-58807-BRAINS \
  --semantics scripts/classify_instruments/semantics.json
```

Current common study commands:

```bash
python3 scripts/create_study_overview/run.py --study-folder studies/53879-62882-OCD-TMS
```

```bash
python3 scripts/create_study_overview/run.py --study-folder studies/63771-LEAP
```

```bash
python3 scripts/create_study_overview/run.py --study-folder studies/71771-LEAP-OL
```

## Workflow Checkpoint

Run this after:

1. REDCap discovery, cleaning, postprocess, and final verification are complete.
2. `scripts/subject_timepoints/run.py` has refreshed cross-instrument
   timepoints.
3. Humans have reviewed the dictionary abbreviations, instrument labels, and
   classification folders.

Review the generated overview before sharing it outside the cleaning workflow.
This is where humans should catch overly generic descriptions, missing
collection methods, or implausible collected timepoint counts.

## Collection Method Rules

The `collection method` column is inferred from the instrument label:

- `survey` -> `Survey`
- `mentor rater`, `certified assessor`, or `clinician` ->
  `Clinician administered`
- `self report` or `self-report` -> `Self-report`

If no rule matches, the cell is left blank for human review.

## Human Responsibilities

After generation:

- Review blank `collection method` values.
- Review descriptions that feel too generic.
- Update `scripts/classify_instruments/semantics.json` for recurring clinical
  terms or abbreviations that should be expanded.
- Confirm excluded labels are genuinely not meant for the overview.
- Regenerate the overview after dictionary, classification, or semantics edits.
