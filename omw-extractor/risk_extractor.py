"""Risk Extractor Module: Extracts risk records under mapped headers into Master JSON structure."""

from datetime import date, datetime
from typing import Any, Dict, List, Optional

VALID_SHEET_NAMES = [
    "Top Hole Batch Campaign",
    "C-1",
    "C-2",
    "C-3",
    "Pilot Hole",
    '7" Liner Contingency',
    '6" Cont. Res. Sec.',
    "Logging",
    "P&A",
    "Sidetrack",
]


def normalize_value(field_name: str, val: Any) -> Optional[str]:
    """Normalize raw cell value to string or None, respecting schema format."""
    if val is None or str(val).strip() == "":
        return None

    # Handle datetime/date objects from openpyxl
    if isinstance(val, (datetime, date)):
        return val.strftime("%Y-%m-%d")

    # If numeric score or rating e.g. 12 or 3.0
    if isinstance(val, float) and val.is_integer():
        return str(int(val))

    return str(val).strip()


def is_risk_worksheet(sheet_name: str, mapped_columns: List[Dict[str, Any]], records_count: int) -> bool:
    """Identify whether a sheet contains operational risk records."""
    name_clean = sheet_name.strip()
    if name_clean in VALID_SHEET_NAMES:
        return True

    name_lower = name_clean.lower()
    if any(k in name_lower for k in ["instruction", "guidance", "legend", "matrix", "config", "lookup"]):
        return False

    mapped_fields = {c["field"] for c in mapped_columns if c["field"]}
    has_core = bool({"riskNo", "riskTitle", "riskDescription", "cause"} & mapped_fields)
    return has_core and records_count > 0


def extract_sheet_records(
    sheet_data: Dict[str, Any], header_row: int, mapped_cols: List[Dict[str, Any]]
) -> List[Dict[str, Any]]:
    """Extract every risk record from a worksheet strictly formatted to Master Schema."""
    source_sheet = sheet_data["sheet_name"].strip()
    records = []
    col_lookup = {c["letter"]: (c["section"], c["field"]) for c in mapped_cols if c["field"]}

    for row in sheet_data.get("rows", []):
        r_num = row["row_number"]
        if r_num <= header_row:
            continue

        cells = row["cells"]
        row_text = " ".join(str(c.get("value", "")).lower() for c in cells.values() if c.get("value") is not None)
        if "risk no" in row_text or "hazard title" in row_text:
            continue  # Repeated header banner

        hazard_id = {k: None for k in ["riskNo", "source", "category", "well", "phase", "riskTitle", "cause", "riskDescription", "existingMitigatingMeasures"]}
        initial_eval = {k: None for k in ["eventType", "consequence", "likelihood", "rating"]}
        treatment = {k: None for k in ["mitigationAction", "actionDueDate", "actionOwner", "actionStatus"]}
        post_eval = {k: None for k in ["consequence", "likelihood", "rating"]}
        audit = {k: None for k in ["comments", "lastUpdatedBy", "lastUpdatedDate", "pimsReference"]}

        sections = {
            "hazardIdentification": hazard_id,
            "initialRiskEvaluation": initial_eval,
            "riskTreatment": treatment,
            "postResponseRiskEvaluation": post_eval,
            "commentsAndAudit": audit,
        }

        has_substance = False
        for letter, (sec, fld) in col_lookup.items():
            if letter in cells:
                raw_val = cells[letter].get("value")
                cleaned = normalize_value(fld, raw_val)
                if cleaned is not None:
                    has_substance = True
                    sections[sec][fld] = cleaned

        # Schema requires riskNo to be non-null string
        if not hazard_id["riskNo"]:
            if any([hazard_id["riskTitle"], hazard_id["riskDescription"], hazard_id["cause"]]):
                hazard_id["riskNo"] = f"R{r_num}"
            else:
                continue

        if has_substance:
            rec = {
                "sourceSheet": source_sheet,
                "hazardIdentification": hazard_id,
                "initialRiskEvaluation": initial_eval,
                "riskTreatment": treatment,
                "postResponseRiskEvaluation": post_eval,
                "commentsAndAudit": audit,
            }
            records.append(rec)

    return records
