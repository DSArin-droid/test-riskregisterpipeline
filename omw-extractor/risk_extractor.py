"""Risk Extractor Module: Extracts risk records under mapped headers into Master JSON structure."""

from datetime import date, datetime
from typing import Any, Dict, List, Optional, Tuple

NON_RISK_TITLE_KEYWORDS = [
    "instruction",
    "guidance",
    "read me",
    "readme",
    "how to",
    "lookup",
    "dropdown",
    "legend",
    "matrix lookup",
    "version history",
    "metadata",
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


def classify_worksheet(
    sheet_name: str, mapped_columns: List[Dict[str, Any]], records_count: int
) -> Tuple[bool, str, Dict[str, Any]]:
    """Dynamically classify a sheet based on Content Fingerprinting (zero hardcoding).

    Returns:
        (is_risk_sheet: bool, reason: str, metadata: dict)
    """
    name_lower = sheet_name.strip().lower()

    # Rule 1: Exclude documentation, legend, and lookup sheets
    for kw in NON_RISK_TITLE_KEYWORDS:
        if kw in name_lower:
            return (
                False,
                f"Excluded: Matched auxiliary/instruction keyword '{kw}' in sheet title",
                {"keyword_matched": kw},
            )

    # Rule 2: Must contain substantive data rows
    if records_count <= 0:
        return (
            False,
            "Excluded: Contains 0 extracted risk records",
            {"records_count": 0},
        )

    # Rule 3: Content Fingerprinting (Hazard Anchor + Action/Evaluation Anchor)
    mapped_fields = {c["field"] for c in mapped_columns if c["field"]}
    hazard_anchors = {"riskNo", "riskTitle", "riskDescription", "cause"} & mapped_fields
    eval_action_anchors = {
        "consequence", "likelihood", "rating", "mitigationAction", "existingMitigatingMeasures"
    } & mapped_fields

    if hazard_anchors and eval_action_anchors:
        return (
            True,
            f"Classified as Risk Sheet: Found Hazard anchors {sorted(hazard_anchors)} and Evaluation/Action anchors {sorted(eval_action_anchors)} with {records_count} data records",
            {
                "hazard_anchors": sorted(hazard_anchors),
                "eval_action_anchors": sorted(eval_action_anchors),
                "records_count": records_count,
            },
        )

    return (
        False,
        "Excluded: Does not meet dual-anchor risk table threshold (missing core hazard or evaluation columns)",
        {
            "hazard_anchors": sorted(hazard_anchors),
            "eval_action_anchors": sorted(eval_action_anchors),
        },
    )


def is_risk_worksheet(sheet_name: str, mapped_columns: List[Dict[str, Any]], records_count: int) -> bool:
    """Helper returning boolean whether sheet is a risk sheet."""
    return classify_worksheet(sheet_name, mapped_columns, records_count)[0]


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
