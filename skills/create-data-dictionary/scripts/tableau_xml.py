"""
Shared library for reading a Tableau workbook (.twbx or .twb) as XML.

A .twbx is a zip; the .twb inside is XML. This module parses that XML into a
structured catalog of every field (physical columns, calculated fields,
parameters), resolves the internal calculation IDs (e.g. [Calculation_8273...])
back to human-readable field names, and computes dependencies and per-sheet usage.

Used by build_dictionary.py, dependency_map.py and diff_workbooks.py.
"""
import xml.etree.ElementTree as ET
import zipfile, re, io, os


def load_twb_text(path):
    """Return the .twb XML text from a .twbx (zip) or a .twb file."""
    if path.lower().endswith(".twbx"):
        with zipfile.ZipFile(path) as z:
            twb = [n for n in z.namelist() if n.lower().endswith(".twb")]
            if not twb:
                raise ValueError("No .twb found inside the .twbx archive.")
            return z.read(twb[0]).decode("utf-8", errors="replace"), twb[0]
    with open(path, encoding="utf-8", errors="replace") as f:
        return f.read(), os.path.basename(path)


def parse_workbook(path):
    """Parse a workbook into a structured dict. Works on any .twbx/.twb."""
    text, twb_name = load_twb_text(path)
    root = ET.fromstring(text)
    top = root.find("datasources")
    datasources = top.findall("datasource") if top is not None else []

    # internal column name -> caption (for rewriting formulas to readable form)
    name2cap = {}
    for ds in datasources:
        for col in ds.findall("column"):
            name2cap[col.get("name")] = col.get("caption") or col.get("name").strip("[]")

    def readable(formula):
        if not formula:
            return ""
        # datasource-qualified refs: [Datasource].[col]
        f = re.sub(r"\[([^\]]+)\]\.\[([^\]]+)\]",
                   lambda m: "[" + name2cap.get("[" + m.group(2) + "]", m.group(2)) + "]", formula)
        # plain refs: [Calculation_x] / [name]
        f = re.sub(r"\[([^\]]+)\]",
                   lambda m: "[" + name2cap.get("[" + m.group(1) + "]", m.group(1)) + "]", f)
        return re.sub(r"\s+", " ", f).strip()

    fields = []
    for ds in datasources:
        dsname = ds.get("caption") or ds.get("name")
        for col in ds.findall("column"):
            calc = col.find("calculation")
            is_calc = (calc is not None and calc.get("class") == "tableau"
                       and calc.get("formula") is not None)
            is_param = col.get("param-domain-type") is not None
            # the field's own in-workbook description (<desc><formatted-text><run>...)
            desc_text = " ".join(r.text or "" for r in col.findall(".//desc//run")).strip()
            fields.append({
                "datasource": dsname,
                "name": col.get("name"),
                "caption": col.get("caption") or col.get("name").strip("[]"),
                "datatype": col.get("datatype"),
                "role": col.get("role"),
                "kind": "parameter" if is_param else ("calculated" if is_calc else "field"),
                "formula": readable(calc.get("formula")) if is_calc else "",
                "param_default": col.get("value") if is_param else None,
                "description": desc_text,
            })

    # dependencies (from readable formulas, matched against known captions)
    caps = {f["caption"] for f in fields}
    for f in fields:
        deps = set(re.findall(r"\[([^\]]+)\]", f["formula"]))
        f["dependencies"] = sorted(d for d in deps if d in caps and d != f["caption"])

    # usage across worksheets
    worksheets = [w.get("name") for w in root.findall(".//worksheets/worksheet")]
    dashboards = [d.get("name") for d in root.findall(".//dashboards/dashboard")]
    usage = {f["caption"]: set() for f in fields}
    for ws in root.findall(".//worksheets/worksheet"):
        refs = set()
        for c in ws.findall(".//datasource-dependencies/column"):
            refs.add(c.get("name"))
        for ci in ws.findall(".//column-instance"):
            refs.add(ci.get("column"))
        for r in refs:
            if r in name2cap:
                usage.setdefault(name2cap[r], set()).add(ws.get("name"))
    for f in fields:
        f["used_in"] = sorted(usage.get(f["caption"], set()))

    # reverse deps + unused flag
    used_by = {f["caption"]: [] for f in fields}
    for f in fields:
        for d in f["dependencies"]:
            used_by.setdefault(d, []).append(f["caption"])
    for f in fields:
        f["feeds_into"] = sorted(set(used_by.get(f["caption"], [])))
        f["unused"] = (f["kind"] == "calculated" and not f["used_in"] and not f["feeds_into"])

    return {
        "source": os.path.basename(path),
        "twb_name": twb_name,
        "datasources": [ds.get("caption") or ds.get("name") for ds in datasources],
        "worksheets": worksheets,
        "dashboards": dashboards,
        "fields": fields,
    }
