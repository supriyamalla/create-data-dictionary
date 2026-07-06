"""
Parse a Tableau workbook and dump its field catalog (for inspection / authoring
plain-English descriptions before building the dictionary).

Usage:
  python parse_workbook.py WORKBOOK.twbx [-o fields.json]

Prints a readable summary (counts + every calculated field with its resolved
formula + the unused-calc list) and saves the full catalog as JSON.
"""
import argparse, json, os, sys
from tableau_xml import parse_workbook

def main():
    # Windows consoles default to cp1252 and crash printing emoji in field names/formulas.
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass
    ap=argparse.ArgumentParser()
    ap.add_argument("workbook"); ap.add_argument("-o","--output")
    a=ap.parse_args()
    wb=parse_workbook(a.workbook)
    out=a.output or os.path.splitext(os.path.basename(a.workbook))[0]+"_fields.json"
    json.dump(wb,open(out,"w"),indent=1)
    from collections import Counter
    k=Counter(f["kind"] for f in wb["fields"])
    print(f"{wb['source']}: {len(wb['datasources'])} datasources, {len(wb['worksheets'])} worksheets, "
          f"{len(wb['dashboards'])} dashboards, {len(wb['fields'])} fields "
          f"({k['calculated']} calculated, {k['parameter']} parameters, {k['field']} physical)")
    print("\nCALCULATED FIELDS (resolved formulas):")
    for f in wb["fields"]:
        if f["kind"]=="calculated":
            print(f"  • {f['caption']}  [{f['datasource']}]\n      = {f['formula']}")
    unused=[f['caption'] for f in wb["fields"] if f["unused"]]
    print(f"\nUNUSED calculated fields (no sheet, nothing depends on them): {unused or 'none'}")
    print(f"\nCatalog saved: {out}")

if __name__=="__main__": main()
