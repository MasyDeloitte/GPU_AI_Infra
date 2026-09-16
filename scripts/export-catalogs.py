from __future__ import annotations

import json
from datetime import date
from pathlib import Path

import openpyxl

ROOT = Path(__file__).resolve().parents[1]
WORKBOOK = ROOT / "Forge_Catalog_Reference_v1.xlsx"
OUTPUT = ROOT / "catalogs"
SHEETS = ["Model_Mapping_Catalog", "Model_Selection", "Hardware_Mapping_Catalog", "GPU_Server_Selection"]


def cell_value(cell):
    value = cell.value
    if hasattr(value, "isoformat"):
        return value.isoformat()
    return value


def export_sheet(workbook, sheet_name):
    sheet = workbook[sheet_name]
    rows = []
    for row in sheet.iter_rows():
        values = [cell_value(cell) for cell in row]
        if any(value is not None for value in values):
            rows.append(values)
    return {
        "sheet": sheet_name,
        "range": sheet.calculate_dimension(),
        "rows": rows,
    }


def main():
    workbook = openpyxl.load_workbook(WORKBOOK, data_only=False)
    OUTPUT.mkdir(exist_ok=True)
    catalog_files = {}
    for sheet_name in SHEETS:
        filename = f"{sheet_name}.json"
        (OUTPUT / filename).write_text(
            json.dumps(export_sheet(workbook, sheet_name), indent=2, ensure_ascii=True),
            encoding="utf-8",
        )
        catalog_files[sheet_name] = filename
    manifest = {
        "catalogVersion": date.today().isoformat(),
        "sourceWorkbook": WORKBOOK.name,
        "sourceSheets": SHEETS,
        "files": catalog_files,
        "product": "InfraPilot AI",
        "formulaPolicy": "Keep formulas in source workbook; JSON export preserves formula text for audit and server-side calculation migration.",
    }
    (OUTPUT / "manifest.json").write_text(json.dumps(manifest, indent=2), encoding="utf-8")
    print(json.dumps(manifest, indent=2))


if __name__ == "__main__":
    main()
