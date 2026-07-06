---
name: create-data-dictionary
description: >-
  Inspect and document a Tableau workbook (read-only) by parsing its XML (a .twbx is a zip
  containing a .twb XML file). Two outputs: (1) a plain-English data dictionary as a filterable
  Excel file — every datasource, calculated field (with internal IDs like [Calculation_8273...]
  resolved back to readable names), parameter, dependency, and per-sheet usage; and (2) a lint /
  health-check report flagging dead calculated fields, calcs missing a description, high-complexity
  calcs (LOD / table-calc / deeply-nested), auto-generated names, and captions reused across
  datasources. Use whenever the user has a Tableau workbook, .twbx, or .twb and wants to document
  it, catalog or explain its calculated fields, understand the business logic inside it, OR audit /
  lint / health-check it, find unused or risky calcs, or clean it up — even if they don't say "data
  dictionary" explicitly (e.g. "document my Tableau dashboard", "what do all these calcs mean",
  "catalog the fields", "audit this workbook", "find dead calculations", "any problems in this .twb").
---

# Tableau Workbook — Inspect (dictionary + lint)

A `.twbx` is a zip; inside is a `.twb` XML file holding every field and formula. This skill parses
it (read-only) and produces two things: a **filterable Excel data dictionary**, and a **lint /
health-check report**. Both share one parser (`tableau_xml.py`); pick either or run both.

## Setup
```bash
pip install -r requirements.txt   # openpyxl (+ matplotlib, only for --graph)
```

## Division of labour
The scripts do the deterministic work (parse, resolve internal calc IDs to readable names,
trace dependencies, compute usage). Writing what each calculation *means to the business* is
the valuable judgement — that's the step you add. A dictionary that explains "Year-over-year
change in sales" beats one that only echoes a formula.

## Workflow — data dictionary

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

   The **Dependency Graph** sheet is added automatically (needs matplotlib). Pass `--no-graph`
   to skip it, e.g. if matplotlib isn't available:
   ```bash
   python scripts/build_dictionary.py WORKBOOK.twbx -d descriptions.json --no-graph -o Data_Dictionary.xlsx
   ```

## Workflow — lint / health-check

One command, no descriptions needed — surfaces maintainability and governance issues:
```bash
python scripts/lint_workbook.py WORKBOOK.twbx -o Lint_Report.xlsx   # .xlsx → Excel, .md → Markdown
```
The output format follows the `-o` extension: `.xlsx` writes a two-sheet Excel report (**Lint
Summary** with per-check counts + a filterable **Findings** sheet — `Check · Field · Datasource ·
Detail · Formula · Suggested fix`, colour-banded per check), `.md` (the default) writes a Markdown
report browsable in Obsidian. Each finding carries a **suggested fix** (for high-complexity calcs
it is tailored to the flag — LOD, table-calc, branching, or length). Both print a per-check count
summary to the console and cover the same checks:
- **Dead calculated fields** — on no sheet and feeding no other field (safe-to-remove candidates).
- **Calcs missing a description** — no in-workbook description; meaning lives only in the formula.
- **High-complexity calcs** — LOD, table-calc, deeply-nested (≥3-way branch), or long (>200 chars).
- **Auto-generated names** — left as Tableau's default `Calculation_…` (rename for clarity).
- **Captions reused across datasources** — same name in more than one datasource (ambiguous).

Report the headline counts to the user and offer to act on the findings (e.g. write the missing
descriptions — which also feed the dictionary's Description column). The lint never edits the
workbook.

## Output
A four-sheet `.xlsx` (three with `--no-graph`):
- **Summary** — datasource/sheet/field counts, a list of unused calcs, and a *Health check (lint)*
  section with the per-check counts.
- **Data Dictionary** — one filterable row per field: name, datasource, kind, role, datatype, your
  description, the resolved formula, dependencies, what it feeds, sheets it's used in, status flag.
  Rows are colour-banded by kind; unused calcs are flagged.
- **Lint** — the health-check findings, embedded automatically: one filterable row per issue
  (`Check · Field · Datasource · Detail · Formula · Suggested fix`), colour-banded per check.
- **Dependency Graph** — the left→right picture (see below; skipped with `--no-graph`).

So `build_dictionary.py` produces one consolidated workbook (dictionary **and** lint). The standalone
`lint_workbook.py` is still there when you want a lint-only report, or a Markdown one for Obsidian.

The third **Dependency Graph** sheet shows a left→right picture of how source columns and
parameters feed each calculated field (arrows = "used by"; red boxes = calculations nothing
depends on). Only *connected* fields appear — standalone calcs are omitted from the picture, so
the Summary sheet stays the authoritative list. Large workbooks produce a dense graph (a note
prints past ~60 connected fields). If matplotlib isn't installed the graph is skipped with a
message and the other two sheets still build.

## Notes
- Always work on a copy; the scripts never modify the input, they write new files.
- A bundled `.hyper` extract is binary data, not XML — this documents logic/structure, not raw rows.
- Identically-named fields in different datasources are kept distinct (keyed on datasource+caption).
