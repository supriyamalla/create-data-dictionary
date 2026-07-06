# create-data-dictionary — a Claude Skill for Tableau

Turn any Tableau workbook into a plain-English **data dictionary** — and a **lint / health-check** report — without opening Tableau.

A `.twbx` is just a zip containing a `.twb` XML file. This Skill parses that XML and documents
every datasource, calculated field (with internal IDs like `[Calculation_8273…]` resolved back
to readable names), parameter, dependency, and per-sheet usage — then writes it all to a clean,
filterable Excel file with unused calculated fields flagged.

## What you get
A filterable `.xlsx` that both **documents** and **audits** the workbook (up to four sheets):
- **Summary** — datasource / worksheet / dashboard / field counts, unused calculated fields, and a
  *Health check (lint)* count section.
- **Data Dictionary** — one filterable row per field: name, datasource, kind, role, data type, a
  plain-English description, the resolved formula, what it depends on, what it feeds, the sheets it's
  used in, and a status flag.
- **Lint** — a health-check: dead calculated fields, calcs missing a description, high-complexity calcs
  (LOD / table-calc / deep branching) each with a suggested fix, auto-generated names, and captions
  reused across datasources.
- **Dependency Graph** — a left→right picture of how columns and parameters feed each calculated field.

A standalone `lint_workbook.py` can also emit a lint-only report as Excel or Markdown.

## Install

### One command (recommended)
Run this in any coding agent with a terminal (Claude Code, Cursor, etc.):

```bash
npx skills@latest add supriyamalla/create-data-dictionary
```

This launches an interactive picker — select the skill and which agent(s) to install it to, and it copies
everything into the right place for you. No download, no manual steps. (Powered by the open-source
[`skills` CLI](https://github.com/vercel-labs/skills).)

### Manual copy (Claude Code)
Prefer not to use the CLI? Copy the `skills/create-data-dictionary/` folder into `~/.claude/skills/`
(personal) or your project's `.claude/skills/` directory. Claude Code discovers it automatically.

### On Claude.ai (web)
Claude.ai can't install from GitHub — it only accepts zip uploads:
1. Download **`create-data-dictionary.zip`** from this repo.
2. Open **Settings → Capabilities (Features) → Skills** and upload the zip as a custom Skill.

## Use
Start a chat, attach your `.twbx`, and ask something like *"Build a data dictionary for this Tableau
workbook."* The skill is picked up automatically.

> Note: Skills don't sync across surfaces. If you use both Claude.ai and Claude Code, install it to each separately.

## Requirements
Python with `openpyxl` (listed in `requirements.txt`). `matplotlib` is optional — only needed for the
Dependency Graph sheet; without it the other sheets still build.

## How it works (the short version)
The scripts do the deterministic work — unzip, parse, resolve calc IDs, trace dependencies, compute
usage. The *plain-English descriptions* of what each calculation means to the business are authored
during the run (that's the part domain knowledge adds). See `skills/create-data-dictionary/SKILL.md` for the
full workflow.

## Files
```
skills/create-data-dictionary/
  SKILL.md              # the Skill definition Claude reads
  requirements.txt      # openpyxl (+ matplotlib for the dependency graph)
  scripts/
    tableau_xml.py      # shared parser: .twbx -> structured field catalog
    parse_workbook.py   # dump fields + resolved formulas (for authoring descriptions)
    build_dictionary.py # build the Excel dictionary (+ embedded Lint & Dependency Graph sheets)
    lint_workbook.py    # health-check: dead / undocumented / complex / duplicate calcs
    dependency_graph.py # render the dependency-graph sheet
create-data-dictionary.zip  # zip for Claude.ai upload
```

## License
MIT — see [LICENSE](LICENSE). Feel free to use, modify, and share.
