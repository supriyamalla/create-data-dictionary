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
  ALSO design-QCs the workbook: inconsistent fonts and font sizes, colours outside the scheme, and
  number formats that disagree (percent at 0 vs 1 vs 2 decimals, currency scaled differently) — use
  it for "check my dashboard formatting", "are the fonts consistent", "do the numbers format the
  same way", "QC this dashboard before it ships", "does this follow our style guide".
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

## Workflow — design QC (style)

The formatting review a human does by eye before a dashboard ships. One command, no setup:
```bash
python scripts/style_check.py WORKBOOK.twbx
```
It prints the findings; `build_dictionary.py` also embeds them as the **Style** sheet.

Two ways to decide what counts as a violation, and they combine:
- **Majority-wins (default).** The convention is inferred from the workbook itself — "3 of 5 places
  use Tableau Medium, 2 use Tableau Regular". No configuration, so it runs on any workbook,
  including a client's. Where a group is too small or genuinely split, it says *no consistent
  standard* and lists the variants rather than crowning an arbitrary winner.
- **Style guide (optional).** Pass approved values and it also flags a workbook that is
  consistently off-brand:
  ```bash
  python scripts/style_check.py WORKBOOK.twbx -g style_guide.json
  python scripts/build_dictionary.py WORKBOOK.twbx -g style_guide.json -o Dictionary.xlsx
  ```
  ```json
  {"fonts":   {"family": ["Tableau Book"], "size": ["9", "10", "12"]},
   "colors":  ["#1f77b4", "#2ca02c", "#d62728"],
   "number_formats": {"currency": {"decimals": 0}, "percent": {"decimals": 1}}}
  ```
  Every key is optional — supply only the ones your team has actually standardised.

The checks:
- **Font family / font size** — per element role (title, header, axis, legend, quick-filter, labels,
  worksheet body), the value most places use is the standard; deviations are listed with their sheet.
- **Number format** — every field's format code is classified into *family + decimals + scale*, so
  `p0%`, `p0.0%` and `n#,##0.0"%"` are compared as the same kind of thing (all percentages) even
  though they are written differently. Flags mismatched precision and mismatched scaling ("$12K" vs
  "$12,000").
- **Off-palette colour** — with a style guide, any colour outside the approved scheme; without one,
  colours used in only one or two places while the rest of the workbook shares a scheme. Greys,
  black, white and transparent are treated as Tableau defaults, not scheme choices.

Deliberately out of scope in this version: tooltip typography (high volume, low review value) and
font weight. Number formats left at the locale default are not counted as a precision decision.

## Output
A six-sheet `.xlsx` (five with `--no-graph`):
- **Summary** — datasource/sheet/field counts, dashboard-action count, a list of unused calcs, and a
  *Health check (lint)* section and a *Design QC (style)* section, each with per-check counts.
- **Data Dictionary** — one filterable row per field: name, datasource, kind, role, datatype, your
  description, the resolved formula, dependencies, what it feeds, sheets it's used in, status flag.
  Rows are colour-banded by kind; unused calcs are flagged.
- **Actions** — the dashboard actions: `Action · Type · Trigger · Source · Target`, colour-banded by
  type (filter / highlight / URL / parameter). Shows "— no dashboard actions —" if the workbook has none.
- **Lint** — the health-check findings, embedded automatically: one filterable row per issue
  (`Check · Field · Datasource · Detail · Formula · Suggested fix`), colour-banded per check.
- **Style** — the design-QC findings: one filterable row per issue (`Check · Where · Element ·
  Found · Workbook standard · Suggested fix`), colour-banded per check. Shows "— no style findings —"
  if the workbook is clean.
- **Dependency Graph** — the left→right picture (see below; skipped with `--no-graph`).

Each colour-banded sheet (Data Dictionary, Actions, Lint, Style) carries a small **colour key** at the top
explaining its fills; the Dependency Graph has its own inline legend.

So `build_dictionary.py` produces one consolidated workbook (dictionary, lint **and** design QC). The standalone
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
