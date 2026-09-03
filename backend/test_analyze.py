# -*- coding: utf-8 -*-
"""
Quick test: analyze the generated Excel template and print the parsed schema.
Run from the excelmapping root folder:
    .venv\Scripts\python.exe backend\test_analyze.py
"""
import sys
from pathlib import Path

backend_path = Path(__file__).resolve().parent
if str(backend_path) not in sys.path:
    sys.path.insert(0, str(backend_path))

from openpyxl import load_workbook
from excel_mapping.excel_manage_helper import ExcelManageHelper

TEMPLATE = backend_path / "templates" / "01_auto_policy_application.xlsx"

wb  = load_workbook(TEMPLATE, data_only=False)
h   = ExcelManageHelper()
schema = h.analyze_template_workbook(wb, file_name=TEMPLATE.name)

sheet = schema["sheets"][0]
print(f"\n=== Sheet: {sheet['name']} ===")
print(f"  value_column detected : {sheet['value_column']}")
print(f"  sections found        : {len(sheet['sections'])}")
print(f"  orphan_fields         : {len(sheet['orphan_fields'])}")
print()

def dump_section(s, indent=0):
    prefix = "  " * indent
    marker = "[ROOT]" if s["level"] == 0 else f"[SEC L{s['level']}]"
    print(f"{prefix}{marker} '{s['name']}'  (row {s['row']}, conf={s['confidence']:.2f})")
    for f in s.get("fields", []):
        print(f"{prefix}  [FIELD] '{f['name']}'  label={f['label_cell']}  value={f['value_cell']}  conf={f['confidence']:.2f}")
    for child in s.get("children", []):
        dump_section(child, indent + 1)

for sec in sheet["sections"]:
    dump_section(sec)

if sheet["orphan_fields"]:
    print("\n--- Orphan Fields ---")
    for f in sheet["orphan_fields"]:
        print(f"  [ORPHAN] '{f['name']}'  label={f['label_cell']}  value={f['value_cell']}")

print(f"\n  needs_review rows: {sheet['needs_review']}")
wb.close()
