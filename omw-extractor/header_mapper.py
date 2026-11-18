"""Header Mapper Module: Maps raw worksheet headers to the 24 Master Header fields across 5 sections."""

from difflib import SequenceMatcher
import json
from pathlib import Path
import re
from typing import Any, Dict, List, Optional, Tuple

PERSISTENT_FILE = Path(__file__).resolve().parent / "persistent_headers.json"

MASTER_SCHEMA = {
    "hazardIdentification": {
        "riskNo": ["risk no", "risk no.", "risk number", "risk id", "id", "risk #", "hazard no"],
        "source": ["source", "risk source", "hazard source", "source / risk assessment", "origin", "identified by"],
        "category": ["category", "risk category", "hazard category", "discipline", "type of risk"],
        "well": ["well", "well name", "well id", "well / rig", "well/rig", "wellbore"],
        "phase": ["phase", "operational phase", "operation phase", "well phase", "drilling phase", "activity"],
        "riskTitle": ["risk title", "hazard title", "title", "short description", "risk headline"],
        "cause": ["cause", "root cause", "causes", "hazard cause", "potential causes", "trigger"],
        "riskDescription": ["risk description", "description", "hazard description", "event description", "risk description (consequence)", "consequence / description"],
        "existingMitigatingMeasures": ["existing mitigating measures", "existing controls", "current controls", "preventive measures", "existing mitigations", "controls in place", "measures in place"],
    },
    "initialRiskEvaluation": {
        "eventType": ["event type", "type of event", "hazard event", "consequence category", "loss type"],
        "consequence": ["initial consequence", "initial severity", "initial impact", "consequence", "severity", "impact", "cons", "c"],
        "likelihood": ["initial likelihood", "initial probability", "initial frequency", "likelihood", "probability", "frequency", "freq", "f", "l"],
        "rating": ["initial rating", "initial score", "initial risk rating", "rating", "risk rating", "risk score", "score", "rr"],
    },
    "riskTreatment": {
        "mitigationAction": ["mitigation action", "mitigation action / contingency", "contingency", "proposed mitigation", "corrective action", "recommended action", "action"],
        "actionDueDate": ["action due date", "due date", "target date", "deadline", "completion date"],
        "actionOwner": ["action owner", "owner", "responsible party", "assigned to", "focal point", "action party"],
        "actionStatus": ["action status", "status", "implementation status", "current status", "state"],
    },
    "postResponseRiskEvaluation": {
        "consequence": ["post-response consequence", "residual consequence", "post consequence", "post-mitigation consequence", "consequence", "severity", "impact"],
        "likelihood": ["post-response likelihood", "residual likelihood", "post likelihood", "post-mitigation likelihood", "residual frequency", "post frequency", "likelihood", "probability", "frequency", "freq", "f"],
        "rating": ["post-response rating", "residual rating", "post rating", "post-mitigation rating", "rating", "risk rating", "score"],
    },
    "commentsAndAudit": {
        "comments": ["comments", "comment", "remarks", "notes", "audit notes"],
        "lastUpdatedBy": ["last updated by", "updated by", "modified by", "editor"],
        "lastUpdatedDate": ["last updated date", "last updated", "update date", "date updated", "modified date"],
        "pimsReference": ["pims reference", "pims ref", "pims #", "pims id", "pims number", "reference no"],
    }
}


def get_persistent_mapping(norm: str) -> Tuple[Optional[str], Optional[str]]:
    """Check persistent layer: exact or >=80% fuzzy match."""
    if not PERSISTENT_FILE.exists() or re.match(r'^(column\s*\d*|\d+)$', norm.lower()):
        return None, None
    cache = json.loads(PERSISTENT_FILE.read_text(encoding="utf-8"))
    if norm in cache:
        return tuple(cache[norm])
    for key, target in cache.items():
        if SequenceMatcher(None, norm, key).ratio() >= 0.80:
            return tuple(target)
    return None, None


def save_persistent_mapping(norm: str, sec: str, fld: str) -> None:
    """Save valid, non-generic, non-ambiguous headers to persistent layer."""
    if not norm or len(norm) > 60 or len(norm.split()) > 8 or fld in {"consequence", "likelihood", "rating"} or re.match(r'^(column\s*\d*|\d+)$', norm.lower()):
        return
    cache = json.loads(PERSISTENT_FILE.read_text(encoding="utf-8")) if PERSISTENT_FILE.exists() else {}
    if norm not in cache:
        cache[norm] = [sec, fld]
        PERSISTENT_FILE.write_text(json.dumps(cache, indent=2), encoding="utf-8")


def clean_text(text: Optional[Any]) -> str:
    """Normalize header text, fixing typos and stripping trailing Excel digits."""
    if text is None:
        return ""
    s = re.sub(r'[\r\n\t_]+', ' ', str(text).strip().lower()).replace("occuring", "occurring")
    s = " ".join(re.sub(r'[^\w\s\./#"-]', '', s).split())
    return re.sub(r'([a-z])\d+$', r'\1', s)


def profile_column_values(samples: List[str]) -> Tuple[Optional[str], Optional[str]]:
    """Infer (section, field) from sampled column values when header is generic."""
    s = " ".join(samples).lower()
    if any(k in s for k in ["plan", "hazid", "dwop", "ops"]):
        return "hazardIdentification", "source"
    if "well" in s:
        return "hazardIdentification", "well"
    if any(k in s for k in ['"', "hole", "spud", "batch", "p&a", "drilling"]):
        return "hazardIdentification", "phase"
    if all(re.match(r'^\d+$|^[a-z]{1,3}-\d+$', v.strip(), re.I) for v in samples):
        return "hazardIdentification", "riskNo"
    if all(re.match(r'^[c]\d$', v.strip(), re.I) for v in samples):
        return "initialRiskEvaluation", "consequence"
    if all(re.match(r'^[lf]\d$', v.strip(), re.I) for v in samples):
        return "initialRiskEvaluation", "likelihood"
    return None, None


def detect_parent_section(text: str) -> Optional[str]:
    """Detect if parent header banner specifies a section."""
    norm = clean_text(text)
    if any(k in norm for k in ["post", "residual", "after mitigation"]):
        return "postResponseRiskEvaluation"
    if any(k in norm for k in ["initial", "inherent", "unmitigated", "pre-mitigation", "gross"]):
        return "initialRiskEvaluation"
    if any(k in norm for k in ["hazard", "identification"]):
        return "hazardIdentification"
    if any(k in norm for k in ["treatment", "action", "mitigation"]):
        return "riskTreatment"
    if any(k in norm for k in ["audit", "comment", "admin"]):
        return "commentsAndAudit"
    return None


def match_header(raw_header: str, parent: Optional[str] = None, seen_initial: bool = False) -> Tuple[Optional[str], Optional[str]]:
    """Match raw header to (section, master_field)."""
    norm = clean_text(raw_header)
    if not norm:
        return None, None

    cached = get_persistent_mapping(norm)
    if cached[0] and cached[1]:
        return cached

    parent_sec = detect_parent_section(parent or "")
    sections_order = (
        ["hazardIdentification", "riskTreatment", "commentsAndAudit", "postResponseRiskEvaluation", "initialRiskEvaluation"]
        if (parent_sec == "postResponseRiskEvaluation" or (seen_initial and not parent_sec))
        else ["hazardIdentification", "riskTreatment", "commentsAndAudit", "initialRiskEvaluation", "postResponseRiskEvaluation"]
    )
    if parent_sec and parent_sec in sections_order:
        sections_order.remove(parent_sec)
        sections_order.insert(0, parent_sec)

    # 1. Exact match
    for sec in sections_order:
        for fld, synonyms in MASTER_SCHEMA[sec].items():
            if norm in synonyms:
                save_persistent_mapping(norm, sec, fld)
                return sec, fld

    # Partial prefix or substring match for longer phrases (5+ chars)
    for sec in sections_order:
        for fld, synonyms in MASTER_SCHEMA[sec].items():
            if any(len(syn) >= 5 and (norm.startswith(syn) or syn in norm) for syn in synonyms):
                save_persistent_mapping(norm, sec, fld)
                return sec, fld

    return None, None


def map_sheet_headers(sheet_data: Dict[str, Any]) -> Tuple[int, List[Dict[str, Any]], List[str]]:
    """Find header row and map columns to master fields."""
    rows = sheet_data.get("rows", [])
    if not rows:
        return 1, [], []

    # Score candidate header rows
    best_row_idx, max_matches = rows[0]["row_number"], 0
    row_map = {r["row_number"]: r for r in rows}

    for r_info in rows[:6]:
        if any(isinstance(c.get("value"), (int, float)) for c in r_info["cells"].values()):
            continue
        r_num = r_info["row_number"]
        matches = sum(1 for c in r_info["cells"].values() if c.get("value") and match_header(str(c["value"]))[0])
        if matches > max_matches:
            max_matches, best_row_idx = matches, r_num

    parent_row = row_map.get(best_row_idx - 1) if best_row_idx > 1 else None
    header_row = row_map.get(best_row_idx, rows[0])

    mapped_cols, unmapped, seen_initial = [], [], set()

    for letter, cell in sorted(header_row["cells"].items(), key=lambda x: x[1]["col"]):
        raw = cell.get("value")
        if raw is None or str(raw).strip() == "":
            continue
        raw_str = str(raw).strip()
        parent_str = str(parent_row["cells"][letter]["value"]).strip() if (parent_row and letter in parent_row["cells"] and parent_row["cells"][letter].get("value")) else None

        sec, fld = match_header(raw_str, parent_str, "consequence" in seen_initial)
        if not (sec and fld):
            samples = [str(r["cells"][letter]["value"]).strip() for r in rows if r["row_number"] > best_row_idx and letter in r["cells"] and r["cells"][letter].get("value") is not None][:5]
            sec, fld = profile_column_values(samples) if samples else (None, None)
            if sec and fld:
                save_persistent_mapping(clean_text(raw_str), sec, fld)

        if sec and fld:
            if sec == "initialRiskEvaluation":
                seen_initial.add(fld)
            mapped_cols.append({"letter": letter, "raw_header": raw_str, "parent_header": parent_str, "section": sec, "field": fld})
        else:
            unmapped.append(f"{cell['coord']} ({letter}): '{raw_str}'")
            mapped_cols.append({"letter": letter, "raw_header": raw_str, "parent_header": parent_str, "section": None, "field": None})

    return best_row_idx, mapped_cols, unmapped
