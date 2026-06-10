---
name: create-data-dictionary
description: >-
  Generate a plain-English data dictionary from a Tableau workbook by parsing its XML
  (a .twbx is a zip containing a .twb XML file). Documents every datasource, calculated
  field (with internal IDs like [Calculation_8273...] resolved back to readable names),
  parameter, dependency, and per-sheet usage as a filterable Excel file. Use whenever the
  user has a Tableau workbook, .twbx, or .twb and wants to document it, explain its
  calculated fields, build a data dictionary or data catalog, or understand the business
  logic inside it — even if they don't say "data dictionary" explicitly (e.g. "document my
  Tableau dashboard", "what do all these calcs mean", "catalog the fields in this workbook").
---

# Tableau Data Dictionary

A `.twbx` is a zip; inside is a `.twb` XML file holding every field and formula. This skill
parses it and produces a clean, filterable Excel data dictionary.

## Setup
```bash
pip install -r requirements.txt   # openpyxl
```

## Division of labour
The scripts do the deterministic work (parse, resolve internal calc IDs to readable names,
trace dependencies, compute usage). Writing what each calculation *means to the business* is
the valuable judgement — that's the step you add. A dictionary that explains "Year-over-year
change in sales" beats one that only echoes a formula.

## Workflow

1. Parse and inspect every calculated field's resolved formula:
   ```bash
   python scripts/parse_workbook.py WORKBOOK.twbx -o fields.json
   ```
2. Read the printed formulas and write a `descriptions.json` mapping each calculated field's
   caption to a one-line plain-English description:
   ```json
   {"Profit Ratio": "Profit as a share of sales (margin).",
    "Days to Ship Actual": "Days between order date and ship date."}
   ```
3. Build the Excel dictionary:
   ```bash
   python scripts/build_dictionary.py WORKBOOK.twbx -d descriptions.json -o Data_Dictionary.xlsx
   ```
   Descriptions are optional, but strongly recommended — without them the Description column
   falls back to blank/formula text.

## Output
A two-sheet `.xlsx`: **Summary** (datasource/sheet/field counts + a list of unused calcs) and
**Data Dictionary** (one filterable row per field: name, datasource, kind, role, datatype,
your description, the resolved formula, dependencies, what it feeds, sheets it's used in, and a
status flag). Rows are colour-banded by kind; unused calcs are flagged.

## Notes
- Always work on a copy; the scripts never modify the input, they write new files.
- A bundled `.hyper` extract is binary data, not XML — this documents logic/structure, not raw rows.
- Identically-named fields in different datasources are kept distinct (keyed on datasource+caption).
