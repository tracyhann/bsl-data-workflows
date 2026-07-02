# Create Data Maps

This workflow creates lightweight data-map workbooks from a cleaned study
folder. The data maps are meant to be easy to copy into a data dictionary,
Google Sheet template, or study-facing documentation.

## Philosophy

Data maps are an index, not another cleaned dataset.

Rules we implement:

- **Point to data, do not duplicate it.** Data-map rows describe where cleaned
  files live under `data/cleaned/`.
- **Keep locations relative to the study folder.** Paths are written as
  `./data/cleaned/...` so the study folder can move without breaking the map.
- **Never leave blank descriptions when a fallback exists.** Descriptions are
  filled from semantic definitions, then dictionary instrument labels, then the
  file stem.
- **Keep raw placeholders visible.** Each class map starts with a `raw`
  placeholder row so humans can document raw data sources later.
- **Make manual platform entry easy.** `platforms-data-map.xlsx` starts with ten
  blank rows for human-populated stage, privacy, and platform metadata.
- **Use stable filenames.** All generated files use the pattern
  `[platforms/class]-data-map.xlsx`.

## Input

Required:

```text
<study-folder>/data/cleaned/
```

Expected cleaned class folders include:

```text
subjects
assessments
neuroimaging
treatments
biologics_biometrics
safety_regulatory
```

Optional but recommended:

```text
<study-folder>/data/cleaned/dictionary.xlsx
scripts/classify_instruments/semantics.json
```

The dictionary and semantics files are used to make descriptions more readable.

## Output

Default output directory:

```text
<study-folder>/data-map/
```

Generated files:

```text
<study-folder>/data-map/platforms-data-map.xlsx
<study-folder>/data-map/subjects-data-map.xlsx
<study-folder>/data-map/assessments-data-map.xlsx
<study-folder>/data-map/neuroimaging-data-map.xlsx
<study-folder>/data-map/treatments-data-map.xlsx
<study-folder>/data-map/biologics_biometrics-data-map.xlsx
<study-folder>/data-map/safety_regulatory-data-map.xlsx
```

Only class folders that exist under `data/cleaned/` get a map workbook.

## Copy-Paste Commands

Run for one study:

```bash
python3 scripts/create_data_maps/run.py \
  --study-folder studies/12345-56789-STUDY
```

Run with an explicit output directory:

```bash
python3 scripts/create_data_maps/run.py \
  --study-folder studies/12345-56789-STUDY \
  --out-dir studies/12345-56789-STUDY/data-map
```

Run with a custom semantics file:

```bash
python3 scripts/create_data_maps/run.py \
  --study-folder studies/12345-56789-STUDY \
  --semantics scripts/classify_instruments/semantics.json
```

Current common study commands:

```bash
python3 scripts/create_data_maps/run.py --study-folder studies/12345-56789-STUDY
```

```bash
python3 scripts/create_data_maps/run.py --study-folder studies/12345-STUDY
```

```bash
python3 scripts/create_data_maps/run.py --study-folder studies/12345-STUDY-OL
```

## Workflow Checkpoint

Run this after:

1. The cleaned study folder is organized under `data/cleaned/<class>/`.
2. Final verification has passed or all warnings are understood.
3. Study overview generation has been reviewed, if the overview is part of the
   delivery package.

Review the generated data maps before publishing them to Google Sheets or other
study documentation. This is where humans should fill raw-source placeholders,
stage/privacy/platform metadata, and any descriptions that need study-specific wording.

## Workbook Structure

`platforms-data-map.xlsx` has:

```text
stage | privacy | description | location
```

Each class data-map workbook has:

```text
stage | description | location
```

The first class-map row is:

```text
raw |  |
```

Cleaned files are listed as:

```text
cleaned/processed | <description> | ./data/cleaned/<class>/<file>.xlsx
```

## Human Responsibilities

After generation:

- Fill in the `raw` placeholder row for each class map.
- Fill or refine `platforms-data-map.xlsx`.
- Review descriptions that fell back to file names.
- Confirm generated paths point to the intended cleaned files.
- If the same vague description recurs, update
  `scripts/classify_instruments/semantics.json` and regenerate.
