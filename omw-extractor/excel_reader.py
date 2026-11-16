"""Excel Reader Module: Safely loads Excel workbooks and extracts raw cells, rows, and merged data."""

from pathlib import Path
from typing import Any, Dict, List, Tuple
import openpyxl
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.worksheet import Worksheet


def load_workbook_safely(file_path: Path) -> openpyxl.Workbook:
    """Load Excel workbook without executing macros."""
    if not file_path.exists():
        raise FileNotFoundError(f"File not found: {file_path}")
    if file_path.suffix.lower() not in [".xlsx", ".xlsm"]:
        raise ValueError(f"Invalid format '{file_path.suffix}'. Only .xlsx and .xlsm are supported.")
    return openpyxl.load_workbook(file_path, data_only=True, keep_vba=False)


def get_merged_cells_map(sheet: Worksheet) -> Dict[Tuple[int, int], Tuple[Any, str]]:
    """Map every merged cell (r, c) to (top_left_value, top_left_coord)."""
    merged_map = {}
    for rng in sheet.merged_cells.ranges:
        min_col, min_row, max_col, max_row = rng.bounds
        top_left_coord = f"{get_column_letter(min_col)}{min_row}"
        top_val = sheet.cell(row=min_row, column=min_col).value
        for r in range(min_row, max_row + 1):
            for c in range(min_col, max_col + 1):
                merged_map[(r, c)] = (top_val, top_left_coord)
    return merged_map


def read_worksheet(sheet: Worksheet) -> Dict[str, Any]:
    """Extract raw cells, coordinates, and merged ranges from a worksheet."""
    merged_map = get_merged_cells_map(sheet)
    max_r, max_c = sheet.max_row or 0, sheet.max_column or 0
    rows_data: List[Dict[str, Any]] = []

    for r in range(1, max_r + 1):
        cells, has_val = {}, False
        for c in range(1, max_c + 1):
            letter = get_column_letter(c)
            coord = f"{letter}{r}"
            val, parent, is_merged = sheet.cell(row=r, column=c).value, None, False

            if (r, c) in merged_map:
                val, parent = merged_map[(r, c)]
                is_merged = True

            if val is not None and str(val).strip() != "":
                has_val = True

            cells[letter] = {
                "coord": coord,
                "col": c,
                "value": val,
                "is_merged": is_merged,
                "merged_parent": parent,
            }

        if has_val:
            rows_data.append({"row_number": r, "cells": cells})

    return {
        "sheet_name": sheet.title,
        "dimensions": {"max_row": max_r, "max_column": max_c},
        "is_empty": len(rows_data) == 0,
        "non_empty_rows_count": len(rows_data),
        "rows": rows_data,
    }


def extract_workbook(file_path: Path) -> Dict[str, Any]:
    """Extract all worksheets and raw data from an Excel workbook."""
    wb = load_workbook_safely(file_path)
    try:
        sheets = [read_worksheet(wb[name]) for name in wb.sheetnames]
        return {
            "filename": file_path.name,
            "filepath": str(file_path.resolve()),
            "total_sheets": len(sheets),
            "sheets": sheets,
        }
    finally:
        wb.close()
