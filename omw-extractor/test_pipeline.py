"""Automated unit and end-to-end tests for omw-extractor local pipeline."""

import json
from pathlib import Path
import jsonschema
import openpyxl
import pytest

from excel_reader import extract_workbook, load_workbook_safely
from header_mapper import map_sheet_headers, match_header
from main import run_pipeline
from risk_extractor import extract_sheet_records, is_risk_worksheet

SCHEMA_PATH = Path(__file__).resolve().parent / "master_schema.json"


@pytest.fixture
def sample_workbook_path(tmp_path) -> Path:
    wb = openpyxl.Workbook()

    # Sheet 1: Top Hole Batch Campaign
    ws1 = wb.active
    ws1.title = "Top Hole Batch Campaign"
    ws1.merge_cells("A1:C1")
    ws1["A1"] = "Hazard Identification"
    ws1.merge_cells("D1:F1")
    ws1["D1"] = "Initial Risk Evaluation"
    ws1.merge_cells("G1:H1")
    ws1["G1"] = "Risk Treatment"

    ws1.append([
        "Risk No", "Risk Title", "Risk Description",
        "Likelihood", "Consequence", "Rating",
        "Mitigation Action", "Action Owner"
    ])

    ws1.append(["L1", "Loss of Returns", "Severe losses at 850m", 3, 4, 12, "Pump LCM pills", "Mud Engineer"])
    ws1.append(["N1", "Shallow Gas", "Gas pocket detected", 2, 5, 10, "Divert through separator", "Drilling Supt"])

    # Sheet 2: C-1 (With duplicate risk identifier L1)
    ws2 = wb.create_sheet(title="C-1")
    ws2.append(["Risk Number", "Risk Title", "Consequence", "Likelihood", "Mitigation Action"])
    ws2.append(["L1", "Casing Wear", 3, 2, "Daily caliper log"])

    # Sheet 3: Non-risk sheet
    ws3 = wb.create_sheet(title="Instructions")
    ws3.append(["Guidance notes for risk assessment"])

    file_path = tmp_path / "test_omw_register.xlsx"
    wb.save(file_path)
    wb.close()
    return file_path


def test_excel_reader(sample_workbook_path):
    raw_data = extract_workbook(sample_workbook_path)
    assert raw_data["total_sheets"] == 3
    assert raw_data["filename"] == "test_omw_register.xlsx"

    sheet1 = next(s for s in raw_data["sheets"] if s["sheet_name"] == "Top Hole Batch Campaign")
    assert sheet1["non_empty_rows_count"] >= 3
    row1 = sheet1["rows"][0]
    assert row1["cells"]["A"]["value"] == "Hazard Identification"
    assert row1["cells"]["B"]["value"] == "Hazard Identification"
    assert row1["cells"]["B"]["is_merged"] is True
    assert row1["cells"]["B"]["merged_parent"] == "A1"


def test_header_mapper_disambiguation():
    sec1, fld1 = match_header("Consequence", parent="Initial Risk")
    assert sec1 == "initialRiskEvaluation"
    assert fld1 == "consequence"

    sec2, fld2 = match_header("Consequence", parent="Residual Risk (Post-Mitigation)")
    assert sec2 == "postResponseRiskEvaluation"
    assert fld2 == "consequence"


def test_risk_extractor_alphanumeric_and_duplicates(sample_workbook_path):
    raw_data = extract_workbook(sample_workbook_path)
    sheet1 = next(s for s in raw_data["sheets"] if s["sheet_name"] == "Top Hole Batch Campaign")
    h_row, mapped_cols, unmapped = map_sheet_headers(sheet1)
    records = extract_sheet_records(sheet1, h_row, mapped_cols)

    assert len(records) == 2
    assert records[0]["hazardIdentification"]["riskNo"] == "L1"
    assert records[1]["hazardIdentification"]["riskNo"] == "N1"


def test_end_to_end_and_strict_json_schema_validation(sample_workbook_path):
    run_pipeline(sample_workbook_path)

    output_dir = Path(__file__).resolve().parent / "output"
    assert output_dir.exists()

    records_file = output_dir / "risk_records.json"
    assert records_file.exists()

    records_data = json.loads(records_file.read_text(encoding="utf-8"))
    assert len(records_data) == 3

    # Load Master Schema provided by user
    schema = json.loads(SCHEMA_PATH.read_text(encoding="utf-8"))

    # Validate EVERY extracted record against the Master JSON Schema
    for rec in records_data:
        jsonschema.validate(instance=rec, schema=schema)

    # Verify duplicate risk numbers are preserved
    risk_nos = [r["hazardIdentification"]["riskNo"] for r in records_data]
    assert risk_nos.count("L1") == 2


def test_generic_headers_and_value_profiling(tmp_path):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Rig Test"
    ws.append(["Column1", "Column 2", "Column3", "Column4", "Risk Title", "Consequence3"])
    ws.append([1, "PLAN", "All Wells", '37" Hole', "Deviated wellbore", "C3"])
    ws.append([2, "PLAN", "All Wells", '37" Hole', "Soft seabed conditions", "C2"])
    
    file_path = tmp_path / "generic_headers_test.xlsx"
    wb.save(file_path)
    wb.close()

    raw_data = extract_workbook(file_path)
    sheet = raw_data["sheets"][0]
    h_row, mapped_cols, unmapped = map_sheet_headers(sheet)
    records = extract_sheet_records(sheet, h_row, mapped_cols)

    assert len(records) == 2
    r1 = records[0]
    assert r1["hazardIdentification"]["riskNo"] == "1"
    assert r1["hazardIdentification"]["source"] == "PLAN"
    assert r1["hazardIdentification"]["well"] == "All Wells"
    assert r1["hazardIdentification"]["phase"] == '37" Hole'
    assert r1["hazardIdentification"]["riskTitle"] == "Deviated wellbore"
    assert r1["initialRiskEvaluation"]["consequence"] == "C3"


def test_persistent_layer_and_fuzzy_matching():
    from header_mapper import get_persistent_mapping, match_header, save_persistent_mapping
    # 1. Save a unique non-generic header
    save_persistent_mapping("barrier description", "hazardIdentification", "existingMitigatingMeasures")
    
    # 2. Check exact lookup
    sec, fld = get_persistent_mapping("barrier description")
    assert sec == "hazardIdentification" and fld == "existingMitigatingMeasures"
    
    # 3. Check >= 80% fuzzy match with typo "barrier descripton" (~94% similarity)
    sec_f, fld_f = match_header("barrier descripton")
    assert sec_f == "hazardIdentification" and fld_f == "existingMitigatingMeasures"

    # 4. Verify generic headers like 'Column 1' are NOT persisted
    save_persistent_mapping("column 1", "hazardIdentification", "riskNo")
    sec_g, fld_g = get_persistent_mapping("column 1")
    assert sec_g is None and fld_g is None
