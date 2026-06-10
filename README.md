# create-data-dictionary — a Claude Skill for Tableau

Turn any Tableau workbook into a plain-English **data dictionary** — without opening Tableau.

A `.twbx` is just a zip containing a `.twb` XML file. This Skill parses that XML and documents
every datasource, calculated field (with internal IDs like `[Calculation_8273…]` resolved back
to readable names), parameter, dependency, and per-sheet usage — then writes it all to a clean,
filterable Excel file with unused calculated fields flagged.

## What you get
A two-sheet `.xlsx`:
- **Summary** — datasource / worksheet / dashboard / field counts, plus a list of unused calculated fields.
- **Data Dictionary** — one filterable row per field: name, datasource, kind, role, data type, a
  plain-English description, the resolved formula, what it depends on, what it feeds, the sheets it's
  used in, and a status flag.

## Install & use

### On Claude.ai (easiest)
1. Download **`create-data-dictionary.zip`** from this repo.
2. In Claude.ai, open **Settings → Capabilities (Features) → Skills** and upload the zip as a custom Skill.
3. Start a chat, attach your `.twbx`, and ask something like *"Build a data dictionary for this Tableau workbook."* Claude picks up the Skill automatically.

### In Claude Code
Copy the `create-data-dictionary/` folder into `~/.claude/skills/` (personal) or your project's
`.claude/skills/` directory. Claude Code discovers it automatically — just ask it to document a workbook.

> Note: Skills don't sync across surfaces. If you use both Claude.ai and the Claude API, upload it to each separately.

## Requirements
Python with `openpyxl` (listed in `requirements.txt`). The dependency map / diff / recolor companion
Skills are separate; this one only needs `openpyxl`.

## How it works (the short version)
The scripts do the deterministic work — unzip, parse, resolve calc IDs, trace dependencies, compute
usage. The *plain-English descriptions* of what each calculation means to the business are authored
during the run (that's the part domain knowledge adds). See `create-data-dictionary/SKILL.md` for the
full workflow.

## Files
```
create-data-dictionary/
  SKILL.md              # the Skill definition Claude reads
  requirements.txt      # openpyxl
  scripts/
    tableau_xml.py      # shared parser: .twbx -> structured field catalog
    parse_workbook.py   # dump fields + resolved formulas (for authoring descriptions)
    build_dictionary.py # build the Excel data dictionary
```

## License
MIT — see [LICENSE](LICENSE). Feel free to use, modify, and share.
