# Classify Instruments

This tool classifies discovered REDCap instruments into broad study-data
categories so cleaned workbooks can be organized into meaningful folders.

## Philosophy

Instrument classification should help humans navigate cleaned data, not become
an irreversible truth source.

Rules we implement:

- **Use both instrument names and labels.** Coded REDCap form names and natural
  text labels are combined before classification.
- **Prefer transparent evidence.** Keyword rules are explicit in
  `keywords.json`; semantic descriptions are explicit in `semantics.json`.
- **Use semantics as supporting evidence.** The default mode uses lexical
  semantic overlap. A local sentence-transformer model can be enabled, but the
  workflow still writes a simple category and confidence score.
- **Keep unknowns honest.** If evidence is weak, the instrument should remain
  `unknown` so a human can update keywords or semantics.
- **Treat confidence as review guidance.** Low confidence means review the
  instrument, not that the data are unusable.

## Categories

The predefined classes are:

```text
subjects, assessments, treatments, neuroimaging, biologics_biometrics, safety_regulatory, admin
```

## Layout

- `run.py`: CLI and classification logic.
- `keywords.json`: observed and related keyword rules by category.
- `semantics.json`: category captions and expanded clinical abbreviations.
- `test_classify_instruments.py`: unit tests for the classifier.

## Primary Inputs

The usual input is a cleaned REDCap dictionary:

```text
<study-folder>/data/cleaned/dictionary.xlsx
```

The dictionary must contain an `instrument` sheet with `instrument` and,
ideally, `instrument_label` columns.

## Copy-Paste Command Template

Set paths:

```bash
STUDY_FOLDER="studies/EXAMPLE_STUDY"
RUN_DATE="YYYY-MM-DD"
DICTIONARY="$STUDY_FOLDER/data/cleaned/dictionary.xlsx"
CLASSIFICATION_OUT="$STUDY_FOLDER/histories/$RUN_DATE/instrument_classification.xlsx"
```

Run with the default keyword and semantic rules:

```bash
python3 scripts/classify_instruments/run.py \
  --dictionary "$DICTIONARY" \
  --out "$CLASSIFICATION_OUT"
```

Run with the optional sentence-transformer scorer:

```bash
python3 scripts/classify_instruments/run.py \
  --dictionary "$DICTIONARY" \
  --out "$CLASSIFICATION_OUT" \
  --use-sentence-transformer \
  --model-name "sentence-transformers/all-MiniLM-L6-v2"
```

Classify a short manual list:

```bash
python3 scripts/classify_instruments/run.py \
  --instrument "madrs::MADRS Montgomery Asberg Depression Rating Scale" \
  --instrument "mri_checklist::MRI checklist" \
  --out "/tmp/instrument_classification.xlsx"
```

## Outputs

The workbook contains one sheet named `instrument_classification` with:

```text
instrument_name, instrument_label, class, confidence
```

This output is used by the REDCap cleaning workflow to move workbooks into
category folders.

## Human Responsibilities

- Review low-confidence and `unknown` classifications.
- Add recurring clinical instruments to `keywords.json` and `semantics.json`.
- Keep definitions in `semantics.json` in this style:

```text
ABBREVIATION Expanded Form brief caption of what it measures
```

- Rerun classification after changing keywords or semantics.

## Tests

Run:

```bash
python3 -m unittest scripts.classify_instruments.test_classify_instruments
```
