"""OMW Risk Register - Local Excel-to-JSON Pipeline (main.py).

Run:
    python main.py
or:
    python main.py "path\\to\\OMW_Risk_Register.xlsx"
"""

import json
from pathlib import Path
import sys
from typing import Any, Dict, List

from excel_reader import extract_workbook
from header_mapper import map_sheet_headers
from risk_extractor import extract_sheet_records, is_risk_worksheet


def run_pipeline(excel_path: Path) -> None:
    """Execute complete extraction pipeline and write 3 JSON artifacts."""
    print(f"\n[1/4] Loading and parsing workbook: {excel_path.name}...")
    raw_data = extract_workbook(excel_path)

    output_dir = Path(__file__).resolve().parent / "output"
    output_dir.mkdir(parents=True, exist_ok=True)

    print("[2/4] Detecting headers and mapping master fields...")
    all_risk_records: List[Dict[str, Any]] = []
    sheet_summaries: List[Dict[str, Any]] = []

    for sheet in raw_data["sheets"]:
        sheet_name = sheet["sheet_name"]
        h_row, mapped_cols, unmapped = map_sheet_headers(sheet)
        records = extract_sheet_records(sheet, h_row, mapped_cols)
        is_risk = is_risk_worksheet(sheet_name, mapped_cols, len(records))

        if is_risk:
            all_risk_records.extend(records)

        sheet_summaries.append({
            "sourceSheet": sheet_name,
            "isRiskSheet": is_risk,
            "dimensions": sheet["dimensions"],
            "nonEmptyRows": sheet["non_empty_rows_count"],
            "headerRowNumber": h_row,
            "totalMappedColumns": sum(1 for c in mapped_cols if c["field"]),
            "unmappedColumns": unmapped,
            "totalRiskRecords": len(records),
        })

    print(f"[3/4] Extracted {len(all_risk_records)} risk records across {len(sheet_summaries)} worksheets.")
    print("[4/4] Writing output JSON files...")

    # 1. Complete raw extraction
    raw_file = output_dir / "workbook_raw.json"
    raw_file.write_text(json.dumps(raw_data, indent=2, default=str), encoding="utf-8")

    # 2. Master risk records
    records_file = output_dir / "risk_records.json"
    records_file.write_text(json.dumps(all_risk_records, indent=2, default=str), encoding="utf-8")

    # 3. Extraction summary report
    report_file = output_dir / "extraction_report.json"
    report_data = {
        "workbook": excel_path.name,
        "totalWorksheets": len(sheet_summaries),
        "riskWorksheetsCount": sum(1 for s in sheet_summaries if s["isRiskSheet"]),
        "auxiliaryWorksheetsCount": sum(1 for s in sheet_summaries if not s["isRiskSheet"]),
        "totalRiskRecords": len(all_risk_records),
        "sheets": sheet_summaries,
    }
    report_file.write_text(json.dumps(report_data, indent=2, default=str), encoding="utf-8")

    print("\n" + "=" * 60)
    print("EXTRACTION COMPLETE! Output files generated in:")
    print(f"  Folder: {output_dir.resolve()}")
    print(f"  1. {raw_file.name}       ({raw_file.stat().st_size:,} bytes)")
    print(f"  2. {records_file.name}     ({records_file.stat().st_size:,} bytes, {len(all_risk_records)} records)")
    print(f"  3. {report_file.name}  ({report_file.stat().st_size:,} bytes)")
    print("=" * 60)


def main() -> None:
    """CLI Entry point."""
    print("=" * 60)
    print("  OMW Risk Register Excel-to-JSON Local Extractor")
    print("=" * 60)

    if len(sys.argv) > 1:
        raw_input_path = sys.argv[1]
    else:
        raw_input_path = input("\nEnter Excel workbook file path: ")

    # Strip whitespace and surrounding quotes from copy-pasting paths
    clean_path = raw_input_path.strip().strip('"').strip("'")
    if not clean_path:
        print("Error: No file path provided.")
        sys.exit(1)

    file_path = Path(clean_path)
    if not file_path.exists():
        print(f"Error: File does not exist at: {file_path}")
        sys.exit(1)

    if file_path.suffix.lower() not in [".xlsx", ".xlsm"]:
        print(f"Error: Invalid file format '{file_path.suffix}'. Please provide a .xlsx or .xlsm file.")
        sys.exit(1)

    try:
        run_pipeline(file_path)
    except Exception as e:
        print(f"\nPipeline Error: {e}", file=sys.stderr)
        sys.exit(2)


if __name__ == "__main__":
    main()
