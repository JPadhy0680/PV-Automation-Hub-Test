# app.py
import streamlit as st
from shared_context import init_shared_context
from listedness_service import dataframe_to_index, assess, render_listedness_updater
init_shared_context()
import pandas as pd
import xml.etree.ElementTree as ET
from datetime import datetime, date
import io
import zipfile
import re
import calendar
import html
from pathlib import Path
from typing import Optional, Set, Tuple, List, Dict
# Ensure multi-line cells render properly
st.markdown(""" """, unsafe_allow_html=True)
st.title("\U0001F4CA\U0001F9E0 E2B_R3 XML Triage Application \U0001F6E0\ufe0f \U0001F680")

# ---------------------------------------------------------------------------------------------------------
# v1.10.9 - editable validity + revised XML naming convention + molecule-only filenames
# - Event Details column shows ONLY clinical details (no "Listedness:" fragments).
# - Listedness column:
#    * If exactly one Celix suspect product: show per-event lines (e.g., "Event 1: Listed").
#    * If 2+ Celix suspects: show one line per product as:
#        "<Drug> - Event 1: Listed; Event 2: Unlisted; ..."
# - Listedness is blank for Non-Valid cases.
# - Case-level Listedness column removed. App Assessment removed. Read-only table.
# ---------------------------------------------------------------------------------------------------------

with st.expander("\U0001F4D6 Instructions"):
    st.markdown("""
- Upload **multiple E2B XML files**.
- The app automatically loads the permanent **MedDRA** Excel from the GitHub `data/` folder to enrich event names.
- The app automatically loads the permanent **Listedness_CX** Excel from the GitHub `data/` folder with four columns: **Active Ingredients**, **PT**, **Expectedness**, **Comment**.
  We compute **Listedness per event** and show it in a separate **Listedness** column.
- If the case has **2 or more Celix suspect products**, the **Listedness** column shows one line per product:
  `Drug X - Event 1: Listed; Event 2: Unlisted; ...`
- Parsed data appears in the **Export & Edit** tab. **Validity** is editable; other columns are read-only.
""")

def _digits_only(s: str) -> str:
    return re.sub(r"\D", "", (s or "").strip())

def format_date(date_str: str) -> str:
    if not date_str:
        return ""
    digits = _digits_only(date_str)
    try:
        if len(digits) >= 8:
            dt = datetime.strptime(digits[:8], "%Y%m%d").date()
            return dt.strftime("%d-%b-%Y")
        elif len(digits) >= 6:
            year = int(digits[:4])
            month = int(digits[4:6])
            return datetime(year, month, 1).strftime("%b-%Y")
        elif len(digits) >= 4:
            year = int(digits[:4])
            return f"{year}"
        else:
            return ""
    except Exception:
        return ""

def parse_date_obj(date_str: str) -> Optional[date]:
    if not date_str:
        return None
    digits = _digits_only(date_str)
    try:
        if len(digits) >= 8:
            return datetime.strptime(digits[:8], "%Y%m%d").date()
        elif len(digits) >= 6:
            year = int(digits[:4])
            month = int(digits[4:6])
            last_day = calendar.monthrange(year, month)[1]
            return datetime(year, month, last_day).date()
        elif len(digits) >= 4:
            year = int(digits[:4])
            return datetime(year, 12, 31).date()
        else:
            return None
    except Exception:
        return None

def map_reporter(code):
    return {
        "1": "Physician",
        "2": "Pharmacist",
        "3": "Other health professional",
        "4": "Lawyer",
        "5": "Consumer or other non-health professional"
    }.get(code, "Unknown")

def map_gender(code):
    return {"1": "Male", "2": "Female"}.get(code, "Unknown")

def map_outcome(code):
    return {
        "1": "Recovered/Resolved",
        "2": "Recovering/Resolving",
        "3": "Not recovered/Ongoing",
        "4": "Recovered with sequelae",
        "5": "Fatal",
        "0": "Unknown"
    }.get(code, "Unknown")

AGE_UNIT_MAP = {"a": "year", "b": "month"}

def map_age_unit(raw_unit: str) -> str:
    if raw_unit is None:
        return ""
    ru = str(raw_unit).strip().lower()
    return AGE_UNIT_MAP.get(ru, ru)

UNKNOWN_TOKENS = {"unk", "asku", "unknown"}

def is_unknown(value: str) -> bool:
    if value is None:
        return True
    v = str(value).strip()
    if not v:
        return True
    return v.lower() in UNKNOWN_TOKENS

def clean_value(value: str) -> str:
    return "" if is_unknown(value) else str(value)

def extract_td_value(report_date_display: str) -> str:
    if not report_date_display:
        return ""
    for line in str(report_date_display).splitlines():
        if line.strip().upper().startswith("TD:"):
            return line.split(":", 1)[1].strip()
    return ""

def build_tracker_comment(validity_value: str, comment_value: str) -> str:
    parts: List[str] = []
    validity_text = clean_value(validity_value)
    comment_text = clean_value(comment_value)

    m = re.match(r"Non-Valid\s*\((.*?)\)", validity_text, flags=re.IGNORECASE | re.DOTALL)
    if m:
        reason = " ".join(m.group(1).split())
        if reason:
            parts.append(reason)
        remainder = validity_text[m.end():].strip()
        if remainder:
            parts.append(" ".join(remainder.split()))
    elif validity_text and validity_text != "Valid":
        parts.append(" ".join(validity_text.split()))

    if comment_text:
        parts.append(" ".join(comment_text.split()))

    deduped: List[str] = []
    seen: Set[str] = set()
    for item in parts:
        key = item.lower()
        if item and key not in seen:
            deduped.append(item)
            seen.add(key)
    return "; ".join(deduped)

def simplify_tracker_validity(validity_value: str) -> str:
    validity_text = clean_value(validity_value)
    if validity_text.startswith("Non-Valid"):
        return "Non-Valid"
    return validity_text

def simplify_tracker_product(product_value: str) -> str:
    if not product_value:
        return ""

    products = [p.strip() for p in str(product_value).split(";") if p.strip()]
    cleaned_products: List[str] = []
    seen: Set[str] = set()

    for product in products:
        norm = normalize_text(product)
        matched = ""

        for prod in company_products:
            pnorm = normalize_text(prod)
            if not pnorm:
                continue
            pattern = r'\b' + re.escape(pnorm) + r'\b'
            if re.search(pattern, norm):
                matched = prod.title()
                break

        if not matched:
            temp = product
            temp = re.sub(r'(?i)\b\d+(?:\.\d+)?\s*(mg|mcg|g|ml)\b', '', temp)
            temp = re.sub(
                r'(?i)\b(tab|tabs|tablet|tablets|cap|caps|capsule|capsules|syrup|inj|injection|vial|amp|ampoule|suspension|oral|solution|cream|gel|ointment)\b',
                '',
                temp,
            )
            temp = re.sub(r'\s+', ' ', temp).strip(" -_,;/")
            matched = temp.title()

        key = matched.lower()
        if matched and key not in seen:
            cleaned_products.append(matched)
            seen.add(key)

    return "; ".join(cleaned_products)

def normalize_text(s: str) -> str:
    s = (s or "").lower()
    s = re.sub(r'[^a-z0-9\s\+\-]', ' ', s)
    s = re.sub(r'\s+', ' ', s).strip()
    return s

# --- Listedness helpers ---
def to_pair_set(df: pd.DataFrame):
    """Build the Active Ingredient + PT expectedness index."""
    return dataframe_to_index(df)

# PL pattern e.g., "PL 12345/6789", "PLGB 12345/6789"
PL_PATTERN = re.compile(r'\b(PL|PLGB|PLNI)\s*([0-9]{5})\s*/\s*([0-9]{4,5})\b', re.IGNORECASE)

def extract_pl_numbers(text: str):
    out = []
    if not text:
        return out
    for m in PL_PATTERN.finditer(text):
        prefix = m.group(1).upper()
        company_code = m.group(2)
        product_code = m.group(3)
        out.append(f"{prefix} {company_code}/{product_code}")
    return out

MY_COMPANY_NAME = "celix"
DEFAULT_COMPETITOR_NAMES = {"glenmark", "cipla", "sun pharma", "dr reddy", "dr. reddy", "torrent", "lupin", "intas", "mankind", "micro labs", "zydus"}

def contains_competitor_name(lot_text: str, competitor_names: Set[str]) -> bool:
    if not lot_text:
        return False
    norm = lot_text.lower()
    if MY_COMPANY_NAME.lower() in norm:
        return False
    for name in competitor_names:
        nm = (name or "").lower().strip()
        if nm and nm in norm:
            return True
    return False

company_products = [
    "abiraterone", "apixaban", "apremilast", "bexarotene", "brivaracetam", "clobazam", "clonazepam",
    "cyanocobalamin", "dabigatran", "dapagliflozin","dimethyl fumarate", "edoxaban", "empagliflozin", "famotidine",
    "fesoterodine", "icatibant", "itraconazole", "linagliptin", "linagliptin + metformin", "metformin", "letrozole", "nintedanib", "pirfenidone", "raltegravir", "ranolazine", "rivaroxaban", "safinamide", "saxagliptin",
    "sitagliptin", "sacubritril + valsartan", "sacubritril", "valsartan", "tamsulosin + solifenacin", "tapentadol", "ticagrelor", "tamsulosin",
    "solifenacin"
]

category2_products = {
    "clobazam", "clonazepam", "cyanocobalamin", "famotidine", "itraconazole","letrozole", "tamsulosin", "solifenacin", "tapentadol"
}

def parse_dd_mmm_yy(s):
    return datetime.strptime(s, "%d-%b-%y").date()

LAUNCH_INFO = {
    "abiraterone": ("launched", parse_dd_mmm_yy("08-Sep-22")),
    "apixaban": ("launched", parse_dd_mmm_yy("26-Feb-25")),
    "apremilast": ("launched", parse_dd_mmm_yy("19-Jan-26")),
    "bexarotene": ("launched", parse_dd_mmm_yy("19-Jan-23")),
    "brivaracetam": ("launched_by_strength", {
        10.0: parse_dd_mmm_yy("21-Aug-26"),
        25.0: parse_dd_mmm_yy("21-Aug-26"),
        50.0: parse_dd_mmm_yy("21-Aug-26"),
        75.0: parse_dd_mmm_yy("21-Aug-26"), 
        100.0: parse_dd_mmm_yy("21-Aug-26")
    }),
    "clobazam": ("launched", parse_dd_mmm_yy("26-Sep-24")),
    "clonazepam": ("launched", parse_dd_mmm_yy("20-Jan-25")),
    "clobazam": ("launched_by_strength", {
        10.0: parse_dd_mmm_yy("26-Sep-24"),
        20.0: None 
    }),
    "cyanocobalamin": ("awaited", None),
    "dabigatran": ("launched_by_strength", {
        75.0: parse_dd_mmm_yy("02-Jul-26"), 
        110.0: parse_dd_mmm_yy("02-Jul-26"),
        150.0: parse_dd_mmm_yy("02-Jul-26")
    }),
    "dapagliflozin": ("launched_by_strength", {
        10.0: parse_dd_mmm_yy("26-Aug-25"), 
        5.0: parse_dd_mmm_yy("10-Sep-25")
    }),
    "dimethyl fumarate": ("launched", parse_dd_mmm_yy("05-Feb-24")),
    "edoxaban": ("awaited", None),
    "empagliflozin": ("awaited", None),
    "famotidine": ("launched", parse_dd_mmm_yy("21-Feb-25")),
    "fesoterodine": ("yet", None),
    "icatibant": ("launched", parse_dd_mmm_yy("28-Jul-22")),
    "itraconazole": ("awaited", None),
    "linagliptin": ("yet", None),
    "linagliptin + metformin": ("yet", None),
    "metformin": ("yet", None),
    "letrozole": ("awaited", None),
    "nintedanib": ("launched", parse_dd_mmm_yy("28-Apr-26")),
    "pirfenidone": ("launched_by_strength", {
        267.0: parse_dd_mmm_yy("29-Jun-22"),
        801.0: parse_dd_mmm_yy("29-Jun-22"),
        534.0: None 
    }),
    "raltegravir": ("yet", None),
    "ranolazine": ("launched", parse_dd_mmm_yy("20-Jul-23")),
    "rivaroxaban": ("launched_by_strength", {
        2.5: parse_dd_mmm_yy("02-Apr-24"),
        10.0: parse_dd_mmm_yy("23-May-24"),
        15.0: parse_dd_mmm_yy("23-May-24"),
        20.0: parse_dd_mmm_yy("23-May-24")
    }),
    "safinamide": ("launched", parse_dd_mmm_yy("28-Apr-26")), 
    "saxagliptin": ("launched_by_strength", {
        5.0: parse_dd_mmm_yy("17-Apr-26"),
        2.5: parse_dd_mmm_yy("20-May-26") 
    }),
    "sitagliptin": ("launched", parse_dd_mmm_yy("12-Mar-26")),
    "sacubritril + valsartan": ("awaited", None),
    "tamsulosin + solifenacin": ("launched", parse_dd_mmm_yy("08-May-23")),
    "tamsulosin": ("launched", parse_dd_mmm_yy("08-May-23")),
    "solifenacin": ("launched", parse_dd_mmm_yy("08-May-23")),
    "tapentadol": ("launched", parse_dd_mmm_yy("01-Feb-24")),
    "ticagrelor": ("yet", None),
    }

# Strength master list supplied by Jagamohan (values are in mg).
PRODUCT_STRENGTHS_MG: Dict[str, Set[Tuple[float, ...]]] = {
    "abiraterone": {(500.0,)}, "apixaban": {(2.5,), (5.0,)}, "apremilast": {(10.0,), (20.0,), (30.0,)},
    "bexarotene": {(75.0,)}, "brivaracetam": {(10.0,), (25.0,), (50.0,), (75.0,), (100.0,)},
    "clobazam": {(10.0,)}, "clonazepam": {(0.5,), (2.0,)}, "cyanocobalamin": {(1.0,)},
    "dabigatran": {(75.0,), (110.0,), (150.0,)}, "dapagliflozin": {(5.0,), (10.0,)}, "dapagliflozine": {(5.0,), (10.0,)},
    "dimethyl fumarate": {(120.0,), (240.0,)}, "edoxaban": {(15.0,), (30.0,), (60.0,)}, "empagliflozin": {(10.0,), (25.0,)},
    "famotidine": {(20.0,), (40.0,)}, "fesoterodine": {(4.0,), (8.0,)}, "icatibant": {(30.0,)}, "itraconazole": {(100.0,)},
    "linagliptin": {(5.0,)}, "linagliptin + metformin": {(2.5, 850.0), (2.5, 1000.0)}, "metformin": {(850.0,), (1000.0,)},
    "nintedanib": {(100.0,), (150.0,)}, "pirfenidone": {(267.0,), (801.0,)}, "raltegravir": {(600.0,)},
    "ranolazine": {(375.0,), (500.0,), (750.0,)}, "rivaroxaban": {(2.5,), (10.0,), (15.0,), (20.0,)},
    "safinamide": {(50.0,), (100.0,)}, "saxagliptin": {(2.5,), (5.0,)}, "sitagliptin": {(25.0,), (50.0,), (100.0,)},
    "sacubritril + valsartan": {(24.0, 26.0), (49.0, 51.0), (97.0, 103.0)}, "sacubitril + valsartan": {(24.0, 26.0), (49.0, 51.0), (97.0, 103.0)},
    "sacubritril": {(24.0,), (49.0,), (97.0,)}, "valsartan": {(26.0,), (51.0,), (103.0,)},
    "tamsulosin + solifenacin": {(6.0, 0.4)}, "tamsulosin": {(0.4,)}, "solifenacin": {(6.0,)},
    "tapentadol": {(50.0,), (100.0,), (150.0,), (200.0,), (250.0,)}, "ticagrelor": {(60.0,), (90.0,)},
}

def _norm_strength_number(num: float) -> float:
    try: return round(float(num), 6)
    except Exception: return num

def _norm_strength_tuple(values) -> Tuple[float, ...]:
    return tuple(_norm_strength_number(v) for v in values)

def extract_strengths_mg(*texts: str) -> Tuple[float, ...]:
    combined = " ".join(str(t or "") for t in texts)
    if not combined.strip(): return tuple()
    combined = combined.replace("&amp;", "&").replace("&", " ").replace("/", " / ").replace("+", " + ").replace(",", " , ")
    return tuple(_norm_strength_number(float(m.group(1))) for m in re.finditer(r'(?<![A-Za-z0-9])([0-9]+(?:\.[0-9]+)?)\s*mg\b', combined, flags=re.IGNORECASE))

def extract_ingredient_strengths_mg(drug, ns) -> Tuple[float, ...]:
    values, seen = [], set()
    numerator_paths = [
        './/hl7:kindOfProduct/hl7:manufacturedProduct/hl7:ingredient/hl7:quantity/hl7:numerator',
        './/hl7:manufacturedProduct/hl7:ingredient/hl7:quantity/hl7:numerator',
        './/hl7:ingredient/hl7:quantity/hl7:numerator',
    ]
    for path in numerator_paths:
        for numerator in drug.findall(path, ns):
            value = numerator.attrib.get('value', '')
            unit = numerator.attrib.get('unit', '')
            key = (value, unit)
            if key in seen: continue
            seen.add(key)
            if value and str(unit).strip().lower() == 'mg':
                try: values.append(_norm_strength_number(float(value)))
                except Exception: pass
    return tuple(values)

def allowed_strengths_for_product(product_name: str) -> Set[Tuple[float, ...]]:
    return PRODUCT_STRENGTHS_MG.get(normalize_text(product_name), set())

def is_strength_allowed_for_product(product_name: str, observed_strengths) -> bool:
    allowed = allowed_strengths_for_product(product_name)
    if not allowed or observed_strengths is None: return True
    observed = (_norm_strength_number(observed_strengths),) if isinstance(observed_strengths, (int, float)) else _norm_strength_tuple(observed_strengths)
    if not observed: return True
    if observed in allowed: return True
    if all(len(a) == 1 for a in allowed):
        allowed_single_values = {a[0] for a in allowed}
        return all(v in allowed_single_values for v in observed)
    return False

def format_strength_tuple(strengths) -> str:
    if strengths is None: return ""
    if isinstance(strengths, (int, float)): strengths = (strengths,)
    out = []
    for v in strengths:
        fv = float(v)
        out.append(str(int(fv)) if fv.is_integer() else str(fv))
    return " / ".join(out) + " mg" if out else ""

def strength_comment_for_product(product_name: str, observed_strengths) -> str:
    if observed_strengths and not is_strength_allowed_for_product(product_name, observed_strengths):
        allowed = allowed_strengths_for_product(product_name)
        allowed_text = ", ".join(sorted(format_strength_tuple(a) for a in allowed)) if allowed else "not configured"
        return f"Product strength mismatch for {product_name}: extracted {format_strength_tuple(observed_strengths)}; allowed strength(s): {allowed_text}."
    return ""

def sanitize_filename_component(value: str, fallback: str = "Unknown") -> str:
    value = clean_value(value)
    value = str(value or "").strip() or fallback
    value = re.sub(r'[<>:"/\\|?*]+', '_', value)
    value = re.sub(r'\s+', '_', value)
    value = re.sub(r'_+', '_', value).strip('._ ')
    return value or fallback

def build_suggested_xml_filename(safety_id: str, validity_value: str, is_serious: bool, molecule_names: List[str], fallback_stem: str = "XML") -> str:
    """Valid cases: SafetyID_S/NS_Molecule.xml; Non-Valid cases: SafetyID_NV_Molecule.xml."""
    safety_part = sanitize_filename_component(safety_id, fallback_stem)
    clean_molecules = [sanitize_filename_component(m, "") for m in molecule_names if sanitize_filename_component(m, "")]
    molecule_part = "+".join(clean_molecules) if clean_molecules else "UnknownMolecule"
    if isinstance(validity_value, str) and validity_value.startswith("Non-Valid"):
        return f"{safety_part}_NV_{molecule_part}.xml"
    seriousness_part = "S" if is_serious else "NS"
    return f"{safety_part}_{seriousness_part}_{molecule_part}.xml"

def make_unique_filename(filename: str, used_names: Set[str]) -> str:
    base = filename[:-4] if filename.lower().endswith('.xml') else filename
    candidate, counter = f"{base}.xml", 2
    while candidate.lower() in used_names:
        candidate = f"{base}_{counter}.xml"
        counter += 1
    used_names.add(candidate.lower())
    return candidate

def get_launch_date(product_name: str, strength_mg) -> Optional[date]:
    key = normalize_text(product_name)
    info = LAUNCH_INFO.get(key)
    if not info:
        return None
    if not is_strength_allowed_for_product(product_name, strength_mg):
        return None
    status, payload = info
    if status == "launched":
        return payload
    if status == "launched_by_strength":
        if isinstance(payload, dict) and payload:
            strength_key = None
            if isinstance(strength_mg, (tuple, list)):
                if len(strength_mg) > 0:
                    strength_key = strength_mg[0]
            elif strength_mg is not None:
                strength_key = strength_mg

            if strength_key is not None:
                try:
                    if strength_key in payload:
                        return payload.get(strength_key)
                    return payload.get(float(strength_key))
                except Exception:
                    pass

            dates = [d for d in payload.values() if d is not None]
            return min(dates) if dates else None
    return None

def get_launch_status(product_name: str) -> Optional[str]:
    key = normalize_text(product_name)
    info = LAUNCH_INFO.get(key)
    if not info:
        return None
    return info[0]



def local_tag_name(tag: str) -> str:
    return str(tag).split('}', 1)[-1].lower() if tag else ''


def xml_text_or_attr(elem, *attr_names: str) -> str:
    if elem is None:
        return ''
    text_value = clean_value(elem.text or '')
    if text_value:
        return text_value
    for attr_name in attr_names:
        attr_value = clean_value(elem.attrib.get(attr_name, ''))
        if attr_value:
            return attr_value
    return ''


def _first_xml_value_by_local_names(root: ET.Element, local_names, attr_names=('extension', 'value', 'root')) -> str:
    wanted = {str(name).strip().lower() for name in local_names if str(name).strip()}
    if root is None or not wanted:
        return ''
    for elem in root.iter():
        if local_tag_name(elem.tag) in wanted:
            value = xml_text_or_attr(elem, *attr_names)
            if value:
                return value
    return ''


def ack_zip_group_from_first_tab_validity(validity_value: str) -> str:
    validity_text = clean_value(validity_value)
    if validity_text == 'Valid':
        return 'Valid'
    if validity_text.startswith('Non-Valid'):
        return 'Non-Valid'
    return 'Other'


def build_first_tab_validity_lookup(rows: List[Dict]) -> Dict[str, str]:
    lookup: Dict[str, str] = {}
    for row in rows or []:
        key = clean_value(row.get('Sender ID', ''))
        if key and key not in lookup:
            lookup[key] = clean_value(row.get('Validity', ''))
    return lookup


def classify_ack_validity(ack_code: str, error_comment: str) -> str:
    ack_code_norm = clean_value(ack_code).zfill(2)
    error_present = bool(clean_value(error_comment))
    if ack_code_norm in {'01', '1'} and not error_present:
        return 'Valid'
    return 'Non-Valid'


def resolve_ack_case_id(root: ET.Element, original_name: str) -> str:
    fallback_stem = Path(original_name).stem
    case_id = _first_xml_value_by_local_names(root, ['safetyreportid'])
    if case_id:
        return case_id
    case_id = _first_xml_value_by_local_names(
        root,
        ['senderidentifier', 'senderid', 'sender', 'sendername', 'sendingorganization', 'sendingorganizationidentifier']
    )
    return case_id or fallback_stem


def parse_ack_xml_metadata(xml_bytes: bytes, original_name: str) -> Dict[str, str]:
    fallback_stem = Path(original_name).stem
    try:
        root = ET.fromstring(xml_bytes)
    except Exception as exc:
        fallback_case_id = sanitize_filename_component(fallback_stem)
        return {
            'Original File Name': original_name,
            'Case ID': fallback_case_id,
            'ACK Code': '',
            'ACK Derived Status': 'Non-Valid',
            'Mapped First-Tab Validity': '',
            'ZIP Group': 'Other',
            'Error Comment': f'XML parse error: {exc}',
            'Renamed File Name': f'{fallback_case_id}_Ack.xml',
        }

    case_id = resolve_ack_case_id(root, original_name)
    ack_code = _first_xml_value_by_local_names(root, ['reportacknowledgmentcode', 'acknowledgmentcode', 'ackcode'])
    error_comment = _first_xml_value_by_local_names(root, ['errormessagecomment', 'errorcomment', 'errortext'])
    ack_status = classify_ack_validity(ack_code, error_comment)
    renamed_file_name = f"{sanitize_filename_component(case_id)}_Ack.xml"

    return {
        'Original File Name': original_name,
        'Case ID': case_id,
        'ACK Code': ack_code,
        'ACK Derived Status': ack_status,
        'Mapped First-Tab Validity': '',
        'ZIP Group': 'Other',
        'Error Comment': error_comment,
        'Renamed File Name': renamed_file_name,
    }

# -------------------------------- UI: Upload & Parse --------------------------

tab1, tab2, tab3, tab4 = st.tabs(["Upload & Parse", "Export & Edit", "Tracker Copy", "ACK Renamer"])
edited_df = None
if "uploader_version" not in st.session_state:
    st.session_state["uploader_version"] = 0

all_rows_display: List[Dict] = []
renamed_xml_meta: Dict[int, Dict] = {}
current_date = datetime.now().strftime("%d-%b-%Y")

# ---------------- Permanent Excel master files ----------------
# Keep these files in the GitHub repository under the data folder:
# data/MedDRA.xlsx
# data/Listedness_CX.xlsx
from shared_context import BASE_DIR, DATA_DIR
MEDDRA_MASTER_BASENAME = "MedDRA"
LISTEDNESS_MASTER_BASENAME = "Listedness_CX"

def find_master_excel_file(base_name: str) -> Optional[Path]:
    """Find a permanent Excel master file in the app data folder.

    Preferred names are base_name.xlsx, base_name.xlsm, or base_name.xls.
    Matching is case-insensitive as a fallback.
    """
    allowed_suffixes = (".xlsx", ".xlsm", ".xls")
    if not DATA_DIR.exists():
        return None

    for suffix in allowed_suffixes:
        candidate = DATA_DIR / f"{base_name}{suffix}"
        if candidate.exists():
            return candidate

    wanted = {f"{base_name}{suffix}".lower() for suffix in allowed_suffixes}
    for candidate in DATA_DIR.iterdir():
        if candidate.is_file() and candidate.name.lower() in wanted:
            return candidate
    return None

def excel_cache_key(file_path: Path) -> Tuple[str, int, int]:
    """Create a cache key that changes whenever the Excel file changes.

    Streamlit cache normally reuses data if function arguments are unchanged.
    Including modified time and size means a GitHub/deployment update to the same
    filename is loaded automatically after redeploy/reboot.
    """
    stat = file_path.stat()
    return (str(file_path.resolve()), int(stat.st_mtime_ns), int(stat.st_size))

@st.cache_data(show_spinner=False)
def load_excel_master(file_path_text: str, modified_time_ns: int, file_size: int) -> pd.DataFrame:
    """Load an Excel master file from repository storage.

    modified_time_ns and file_size are intentionally included in the function
    signature so Streamlit refreshes the cached data whenever the file changes.
    """
    file_path = Path(file_path_text)
    suffix = file_path.suffix.lower()
    engine = "openpyxl" if suffix in {".xlsx", ".xlsm"} else None
    return pd.read_excel(file_path, engine=engine)

def read_master_excel(file_path: Path) -> pd.DataFrame:
    file_path_text, modified_time_ns, file_size = excel_cache_key(file_path)
    return load_excel_master(file_path_text, modified_time_ns, file_size)

def load_permanent_meddra_mapping() -> Optional[pd.DataFrame]:
    file_path = find_master_excel_file(MEDDRA_MASTER_BASENAME)
    if not file_path:
        st.warning("Permanent MedDRA Excel not found. Add data/MedDRA.xlsx to GitHub.")
        return None
    try:
        df = read_master_excel(file_path)
        if "LLT Code" in df.columns:
            df["LLT Code"] = df["LLT Code"].astype(str).str.strip()
        st.caption(
            f"Using permanent MedDRA master: {file_path.name} "
            f"(updated {datetime.fromtimestamp(file_path.stat().st_mtime).strftime('%d-%b-%Y %H:%M')})"
        )
        return df
    except Exception as e:
        st.error(f"Failed to read permanent MedDRA Excel ({file_path.name}): {e}")
        return None

def load_permanent_listedness_pairs() -> Set[Tuple[str, str]]:
    file_path = find_master_excel_file(LISTEDNESS_MASTER_BASENAME)
    if not file_path:
        st.warning("Permanent Listedness_CX Excel not found. Add data/Listedness_CX.xlsx to GitHub.")
        return set()
    try:
        df = read_master_excel(file_path)
        pairs = to_pair_set(df)
        st.caption(
            f"Using permanent listedness master: {file_path.name} "
            f"({len(pairs)} pairs loaded; updated {datetime.fromtimestamp(file_path.stat().st_mtime).strftime('%d-%b-%Y %H:%M')})"
        )
        if not pairs:
            st.info("Listedness_CX file loaded but produced no valid pairs. Check columns: Active Ingredients, PT, Expectedness and Comment.")
        return pairs
    except Exception as e:
        st.error(f"Failed to read permanent Listedness_CX Excel ({file_path.name}): {e}")
        return set()

with tab1:
    st.markdown("### \U0001F50E Upload Files \U0001F5C2\ufe0f")
    if st.button("Clear Inputs", help="Clear uploaded XMLs and parsed data (keep access)."):
        for k in ["df_display", "edited_df"]:
            st.session_state.pop(k, None)
        st.session_state["uploader_version"] = st.session_state.get("uploader_version", 0) + 1
        st.rerun()

    ver = st.session_state.get("uploader_version", 0)
    uploaded_files = st.file_uploader(
        "Upload E2B XML files",
        type=["xml"],
        accept_multiple_files=True,
        help="Upload one or more E2B XML files for parsing.",
        key=f"xml_uploader_{ver}"
    )
    st.info("MedDRA and Listedness_CX are loaded automatically from the GitHub repository data folder. Upload only XML files here.")

    competitor_names: Set[str] = set(DEFAULT_COMPETITOR_NAMES)

    mapping_df = load_permanent_meddra_mapping()
    listedness_pairs: Set[Tuple[str, str]] = load_permanent_listedness_pairs()

    seriousness_map = {
        "resultsInDeath": "Death",
        "isLifeThreatening": "LT",
        "requiresInpatientHospitalization": "Hospital",
        "resultsInPersistentOrSignificantDisability": "Disability",
        "congenitalAnomalyBirthDefect": "Congenital",
        "otherMedicallyImportantCondition": "IME"
    }

    if uploaded_files:
        st.markdown("### \u23f3 Parsing Files...")
        progress = st.progress(0)
        total_files = len(uploaded_files)
        parsed_rows = 0

        for idx, uploaded_file in enumerate(uploaded_files, start=1):
            warnings: List[str] = []
            comments: List[str] = []
            try:
                xml_bytes = uploaded_file.getvalue()
                root = ET.fromstring(xml_bytes)
                original_file_stem = Path(getattr(uploaded_file, 'name', f'XML_{idx}')).stem
            except Exception as e:
                st.error(f"Failed to parse XML file {getattr(uploaded_file, 'name', '(unnamed)')}: {e}")
                progress.progress(idx / total_files)
                continue

            ns = {'hl7': 'urn:hl7-org:v3', 'xsi': 'http://www.w3.org/2001/XMLSchema-instance'}

            # Sender
            sender_elem = root.find('.//hl7:id[@root="2.16.840.1.113883.3.989.2.1.3.1"]', ns)
            sender_id = clean_value(sender_elem.attrib.get('extension', '') if sender_elem is not None else '')

            # TD fallback (for case age)
            creation_elem = root.find('.//hl7:creationTime', ns)
            creation_raw = creation_elem.attrib.get('value', '') if creation_elem is not None else ''
            td_fallback = clean_value(format_date(creation_raw))

            # Reporter Qualification
            reporter_elem = root.find('.//hl7:asQualifiedEntity/hl7:code', ns)
            reporter_qualification = clean_value(map_reporter(reporter_elem.attrib.get('code', '') if reporter_elem is not None else ''))

            # Patient details
            gender_elem = root.find('.//hl7:administrativeGenderCode', ns)
            gender_mapped = map_gender(gender_elem.attrib.get('code', '') if gender_elem is not None else '')
            gender = clean_value(gender_mapped)

            age_elem = root.find('.//hl7:code[@displayName="age"]/../hl7:value', ns)
            age = ""
            if age_elem is not None:
                age_val = age_elem.attrib.get('value', '')
                raw_unit = age_elem.attrib.get('unit', '')
                unit_text = map_age_unit(raw_unit)
                age_val = clean_value(age_val)
                unit_text_disp = clean_value(unit_text)
                if age_val:
                    try:
                        n = float(age_val)
                        if unit_text_disp in ("year", "month"):
                            unit_text_disp = unit_text_disp + ("s" if n != 1 else "")
                    except Exception:
                        pass
                age = f"{age_val}" + (f" {unit_text_disp}" if age_val and unit_text_disp else "") if age_val else ""

            weight_elem = root.find('.//hl7:code[@displayName="bodyWeight"]/../hl7:value', ns)
            weight_val = clean_value(weight_elem.attrib.get('value', '') if weight_elem is not None else '')
            weight_unit = clean_value(weight_elem.attrib.get('unit', '') if weight_elem is not None else '')
            weight = f"{weight_val}" + (f" {weight_unit}" if weight_val and weight_unit else "") if weight_val else ""

            height_elem = root.find('.//hl7:code[@displayName="height"]/../hl7:value', ns)
            height_val = clean_value(height_elem.attrib.get('value', '') if height_elem is not None else '')
            height_unit = clean_value(height_elem.attrib.get('unit', '') if height_elem is not None else '')
            height = f"{height_val}" + (f" {height_unit}" if height_val and height_unit else "") if height_val else ""

            patient_initials = ""
            name_elem = root.find('.//hl7:player1/hl7:name', ns)
            if name_elem is not None:
                if 'nullFlavor' in name_elem.attrib and name_elem.attrib.get('nullFlavor') == 'MSK':
                    patient_initials = "Masked"
                else:
                    init_parts = []
                    for g in name_elem.findall('hl7:given', ns):
                        if g.text and g.text.strip():
                            init_parts.append(g.text.strip()[0].upper())
                    fam = name_elem.find('hl7:family', ns)
                    if fam is not None and fam.text and fam.text.strip():
                        init_parts.append(fam.text.strip()[0].upper())
                    if init_parts:
                        patient_initials = "".join(init_parts)
                    else:
                        if name_elem.text and name_elem.text.strip():
                            patient_initials = name_elem.text.strip()
            patient_initials = clean_value(patient_initials)

            age_group_map = {"0": "Foetus", "1": "Neonate", "2": "Infant", "3": "Child", "4": "Adolescent", "5": "Adult", "6": "Elderly"}
            age_group_elem = root.find('.//hl7:code[@displayName="ageGroup"]/../hl7:value', ns)
            age_group = ""
            if age_group_elem is not None:
                code_val = age_group_elem.attrib.get('code', '')
                null_flavor = age_group_elem.attrib.get('nullFlavor', '')
                if code_val in age_group_map:
                    age_group = age_group_map[code_val]
                elif null_flavor in ["MSK", "UNK", "ASKU", "NI"] or code_val in ["MSK", "UNK", "ASKU", "NI"]:
                    age_group = "[Masked/Unknown]"
            age_group = clean_value(age_group)

            # Patient Record Number (OID)
            patient_record_no = ''
            oid = "2.16.840.1.113883.3.989.2.1.3.7"
            for id_elem in root.findall('.//hl7:id', ns):
                if id_elem.attrib.get('root') == oid:
                    nf = id_elem.attrib.get('nullFlavor', '')
                    ext = id_elem.attrib.get('extension', '')
                    if nf == 'MSK':
                        patient_record_no = 'Masked'
                    elif ext:
                        patient_record_no = ext.strip()
                    break

            patient_parts = []
            if patient_initials:
                patient_parts.append(f"Initials: {patient_initials}")
            if gender:
                patient_parts.append(f"Gender: {gender}")
            if age_group:
                patient_parts.append(f"Age Group: {age_group}")
            if age:
                patient_parts.append(f"Age: {age}")
            if height:
                patient_parts.append(f"Height: {height}")
            if weight:
                patient_parts.append(f"Weight: {weight}")
            if patient_record_no:
                patient_parts.append(f"Record No: {patient_record_no}")
            patient_detail = ", ".join(patient_parts)

            has_any_patient_detail = any([patient_initials, gender, age_group, age, height, weight,patient_record_no])

            # Identify suspect products (value==1)
            suspect_ids: List[str] = []
            for causality in root.findall('.//hl7:causalityAssessment', ns):
                val_elem = causality.find('.//hl7:value', ns)
                if val_elem is not None and val_elem.attrib.get('code') == '1':
                    subj_id_elem = causality.find('.//hl7:subject2/hl7:productUseReference/hl7:id', ns)
                    if subj_id_elem is not None:
                        suspect_ids.append(subj_id_elem.attrib.get('root', ''))

            product_details_list: List[str] = []
            case_has_category2 = False
            case_drug_dates_display: List[Tuple[str, Optional[float], Optional[date], Optional[date]]] = []
            case_event_dates: List[Tuple[str, Optional[date], Optional[date]]] = []
            case_displayed_mahs: List[str] = []
            case_products_norm: Set[str] = set()
            product_norm_to_pretty: Dict[str, str] = {}
            suspect_product_names_all: List[str] = []

            displayed_drugs_assessment: List[Tuple[str, str]] = []

            for drug in root.findall('.//hl7:substanceAdministration', ns):
                id_elem = drug.find('.//hl7:id', ns)
                drug_id = id_elem.attrib.get('root', '') if id_elem is not None else ''
                if drug_id in suspect_ids:
                    name_elem_drug = drug.find('.//hl7:kindOfProduct/hl7:name', ns)
                    raw_drug_text = ""
                    if name_elem_drug is not None:
                        if name_elem_drug.text and name_elem_drug.text.strip():
                            raw_drug_text = name_elem_drug.text.strip()
                        else:
                            orig = name_elem_drug.find('hl7:originalText', ns)
                            if orig is not None and orig.text and orig.text.strip():
                                raw_drug_text = orig.text.strip()
                        if not raw_drug_text and 'displayName' in name_elem_drug.attrib:
                            raw_drug_text = name_elem_drug.attrib.get('displayName', '').strip()
                    if not raw_drug_text:
                        alt_name = drug.find('.//hl7:manufacturedProduct/hl7:name', ns)
                        if alt_name is not None and alt_name.text and alt_name.text.strip():
                            raw_drug_text = alt_name.text.strip()

                    if raw_drug_text:
                        cleaned_suspect_name = clean_value(raw_drug_text)
                        if cleaned_suspect_name:
                            suspect_product_names_all.append(cleaned_suspect_name)

                    def contains_company_product(text: str, company_products: list) -> str:
                        norm = normalize_text(text)
                        for prod in company_products:
                            pnorm = normalize_text(prod)
                            if not pnorm:
                                continue
                            pattern = r'\b' + re.escape(pnorm) + r'\b'
                            if re.search(pattern, norm):
                                return prod
                        return ""

                    matched_company_prod = contains_company_product(raw_drug_text, company_products)
                    if matched_company_prod:
                        norm_key = normalize_text(matched_company_prod)
                        case_products_norm.add(norm_key)
                        pretty_name = raw_drug_text if raw_drug_text else matched_company_prod.title()
                        product_norm_to_pretty.setdefault(norm_key, clean_value(pretty_name))
                        if norm_key in category2_products:
                            case_has_category2 = True

                    text_elem = drug.find('.//hl7:text', ns)
                    dose_elem = drug.find('.//hl7:doseQuantity', ns)
                    dose_val_raw = dose_elem.attrib.get('value', '') if dose_elem is not None else ''
                    dose_unit_raw = dose_elem.attrib.get('unit', '') if dose_elem is not None else ''
                    dose_val = clean_value(dose_val_raw)
                    dose_unit = clean_value(dose_unit_raw)

                    start_elem = drug.find('.//hl7:low', ns)
                    stop_elem = drug.find('.//hl7:high', ns)
                    start_date_str = start_elem.attrib.get('value', '') if start_elem is not None else ''
                    stop_date_str = stop_elem.attrib.get('value', '') if stop_elem is not None else ''
                    start_date_disp = clean_value(format_date(start_date_str))
                    stop_date_disp = clean_value(format_date(stop_date_str))
                    start_date_obj = parse_date_obj(start_date_str)
                    stop_date_obj = parse_date_obj(stop_date_str)

                    mah_name_raw = ''
                    for path in [
                        './/hl7:playingOrganization/hl7:name',
                        './/hl7:manufacturerOrganization/hl7:name',
                        './/hl7:asManufacturedProduct/hl7:manufacturerOrganization/hl7:name',
                    ]:
                        node = drug.find(path, ns)
                        if node is not None and node.text and node.text.strip():
                            mah_name_raw = node.text.strip()
                            break
                    mah_name_clean = clean_value(mah_name_raw)

                    if matched_company_prod:
                        parts = []
                        display_name_for_detail = raw_drug_text if raw_drug_text else matched_company_prod.title()
                        display_name_for_detail = clean_value(display_name_for_detail)
                        if display_name_for_detail:
                            parts.append(f"Drug: {display_name_for_detail}")

                        text_clean = ""
                        if text_elem is not None and text_elem.text:
                            text_clean = clean_value(text_elem.text)
                        if text_clean:
                            parts.append(f"Dosage: {text_clean}")

                        if dose_val or dose_unit:
                            if dose_val and dose_unit:
                                parts.append(f"Dose: {dose_val} {dose_unit}")
                            elif dose_val:
                                parts.append(f"Dose: {dose_val}")
                            elif dose_unit:
                                parts.append(f"Dose Unit: {dose_unit}")

                        if start_date_disp:
                            parts.append(f"Start Date: {start_date_disp}")
                        if stop_date_disp:
                            parts.append(f"Stop Date: {stop_date_disp}")

                        form_elem = drug.find('.//hl7:formCode/hl7:originalText', ns)
                        form_clean = ""
                        if form_elem is not None and form_elem.text:
                            form_clean = clean_value(form_elem.text)
                        if form_clean:
                            parts.append(f"Formulation: {form_clean}")

                        ingredient_strengths = extract_ingredient_strengths_mg(drug, ns)
                        product_name_strengths = extract_strengths_mg(display_name_for_detail)
                        observed_strengths = ingredient_strengths or product_name_strengths

                        if ingredient_strengths:
                            parts.append(f"XML Ingredient Strength: {format_strength_tuple(ingredient_strengths)}")
                        if product_name_strengths:
                            parts.append(f"Product Name Strength: {format_strength_tuple(product_name_strengths)}")
                        if ingredient_strengths and product_name_strengths and ingredient_strengths != product_name_strengths:
                            comments.append(f"Product strength mismatch between XML ingredient field and product name for {display_name_for_detail}: XML ingredient strength {format_strength_tuple(ingredient_strengths)} vs product name strength {format_strength_tuple(product_name_strengths)}.")
                        strength_comment = strength_comment_for_product(matched_company_prod, observed_strengths)
                        if strength_comment:
                            comments.append(strength_comment)

                        lot_elem = drug.find('.//hl7:lotNumberText', ns)
                        lot_clean = ""
                        if lot_elem is not None and lot_elem.text:
                            lot_clean = clean_value(lot_elem.text)
                        if lot_clean:
                            parts.append(f"Lot No: {lot_clean}")

                        if re.search(r'[A-Za-z0-9]', lot_clean):
                            comments.append('Verify Lot No with Celix-Lot No List')

                        if mah_name_clean:
                            parts.append(f"MAH: {mah_name_clean}")
                        case_displayed_mahs.append(mah_name_clean)

                        for t in [display_name_for_detail, text_clean, form_clean, lot_clean]:
                            for pl in extract_pl_numbers(t):
                                comments.append(
                                    f"plz check product name as {display_name_for_detail} {pl} given"
                                    if display_name_for_detail else f"plz check product name: {pl} given"
                                )
                        if lot_clean and contains_competitor_name(lot_clean, competitor_names):
                            comments.append(f"Lot number '{lot_clean}' may belong to another company — please verify.")
                        if mah_name_clean and MY_COMPANY_NAME.lower() not in mah_name_clean.lower():
                            comments.append(f"MAH '{mah_name_clean}' differs from Celix — please verify.")

                        if parts:
                            product_details_list.append("\n ".join(parts))

                        non_valid_reason = ""
                        if not has_any_patient_detail:
                            non_valid_reason = "No patient details"
                        else:
                            status = get_launch_status(matched_company_prod)
                            if status in ("yet", "awaited"):
                                non_valid_reason = "Product not Launched"
                            else:
                                launch_dt = get_launch_date(matched_company_prod, observed_strengths)
                                exposure_reasons = []
                                # We'll use FRD/LRD computed later
                                drug_prior = (start_date_obj and start_date_obj < (launch_dt or start_date_obj)) if launch_dt else False
                                if launch_dt and drug_prior:
                                    exposure_reasons.append("Drug")
                                if exposure_reasons:
                                    non_valid_reason = f"Drug exposure prior to Launch; {', '.join(sorted(set(exposure_reasons)))}"
                        displayed_drugs_assessment.append((display_name_for_detail or "Unknown product", non_valid_reason))

                        case_drug_dates_display.append((matched_company_prod, observed_strengths, start_date_obj, stop_date_obj))

            seriousness_criteria = list(seriousness_map.keys())
            event_details_list: List[str] = []
            event_llts_norm: List[str] = []
            event_pts: List[str] = []
            event_count = 1
            case_has_serious_event = False

            for reaction in root.findall('.//hl7:observation', ns):
                code_elem = reaction.find('hl7:code', ns)
                if code_elem is not None and code_elem.attrib.get('displayName') == 'reaction':
                    value_elem = reaction.find('hl7:value', ns)
                    llt_code = value_elem.attrib.get('code', '') if value_elem is not None else ''
                    llt_term, pt_term = "", ""

                    if mapping_df is not None and llt_code:
                        try:
                            llt_code_str = str(llt_code).strip()
                            row = mapping_df[mapping_df['LLT Code'] == llt_code_str]
                            if not row.empty:
                                llt_term = str(row['LLT Term'].values[0])
                                pt_term = str(row['PT Term'].values[0])
                            else:
                                warnings.append(f"LLT code {llt_code_str} not found in mapping file — LLT/PT terms unavailable for this event.")
                        except Exception as e:
                            warnings.append(f"LLT mapping failed for code {llt_code}: {e}")
                    elif llt_code:
                        warnings.append(f"LLT mapping file not provided — LLT/PT terms unavailable for code {llt_code}.")

                    if not llt_term and value_elem is not None:
                        llt_term = value_elem.attrib.get('displayName', '') or llt_term

                    llt_norm = normalize_text(llt_term)
                    event_llts_norm.append(normalize_text(pt_term or llt_term))
                    event_pts.append(pt_term or llt_term)

                    seriousness_flags = []
                    for criterion in seriousness_criteria:
                        criterion_elem = reaction.find(f'.//hl7:code[@displayName="{criterion}"]/../hl7:value', ns)
                        if criterion_elem is not None and criterion_elem.attrib.get('value') == 'true':
                            seriousness_flags.append(seriousness_map.get(criterion, criterion))
                    seriousness_display = "Non-serious" if not seriousness_flags else ", ".join(seriousness_flags)
                    if seriousness_flags:
                        case_has_serious_event = True

                    outcome_elem = reaction.find('.//hl7:code[@displayName="outcome"]/../hl7:value', ns)
                    outcome = map_outcome(outcome_elem.attrib.get('code', '') if outcome_elem is not None else '')
                    outcome = clean_value(outcome)

                    evt_low = reaction.find('.//hl7:effectiveTime/hl7:low', ns)
                    evt_high = reaction.find('.//hl7:effectiveTime/hl7:high', ns)
                    evt_low_str = evt_low.attrib.get('value', '') if evt_low is not None else ''
                    evt_high_str = evt_high.attrib.get('value', '') if evt_high is not None else ''
                    evt_low_disp = clean_value(format_date(evt_low_str))
                    evt_high_disp = clean_value(format_date(evt_high_str))
                    evt_low_obj = parse_date_obj(evt_low_str)
                    evt_high_obj = parse_date_obj(evt_high_str)
                    case_event_dates.append(("event", evt_low_obj, evt_high_obj))

                    base = f"Event {event_count}: {llt_term} ({pt_term})" if pt_term else f"Event {event_count}: {llt_term}"
                    details_parts = [base, f"Seriousness: {seriousness_display}"]
                    if outcome:
                        details_parts.append(f"Outcome: {outcome}")
                    if evt_low_disp:
                        details_parts.append(f"Event Start: {evt_low_disp}")
                    if evt_high_disp:
                        details_parts.append(f"Event End: {evt_high_disp}")
                    event_details_list.append("; ".join(details_parts))

                    event_count += 1

            event_details_combined_display = "\n".join(event_details_list)

            reportability = "Category 2, serious, reportable case" if (case_has_serious_event and case_has_category2) else "Non-Reportable"

            global_dates = {
                "FRD_raw": "",
                "LRD_raw": "",
                "TD_raw": "",
                "FRD": "",
                "LRD": "",
                "TD": "",
            }
            try:
                # TD
                for el in root.iter():
                    if el.tag.endswith('creationTime'):
                        val = el.attrib.get('value')
                        if val:
                            global_dates["TD_raw"] = val
                            global_dates["TD"] = format_date(val)
                            break
                # FRD (last low), LRD (first availabilityTime)
                last_low_value = None
                for el in root.iter():
                    ln = el.tag.split('}')[-1] if '}' in el.tag else el.tag
                    if ln == 'low':
                        v = el.attrib.get('value')
                        if v:
                            last_low_value = v
                    elif ln == 'availabilityTime':
                        v = el.attrib.get('value')
                        if v and not global_dates.get("LRD_raw"):
                            global_dates["LRD_raw"] = v
                            global_dates["LRD"] = format_date(v)
                            break
                if last_low_value:
                    global_dates["FRD_raw"] = last_low_value
                    global_dates["FRD"] = format_date(last_low_value)
            except Exception:
                pass

            frd_disp = global_dates["FRD"]
            lrd_disp = global_dates["LRD"]
            td_disp = global_dates["TD"] or td_fallback

            case_age_days = ""
            if global_dates["TD_raw"]:
                td_obj = parse_date_obj(global_dates["TD_raw"])
                if td_obj:
                    case_age_days = (datetime.now().date() - td_obj).days
                    if case_age_days < 0:
                        case_age_days = 0

            validity_reason: Optional[str] = None
            has_any_suspect = bool(suspect_ids)
            has_celix_suspect = bool(case_products_norm)

            if not has_any_patient_detail:
                validity_reason = "No patient details"

            if validity_reason is None and has_any_suspect and not has_celix_suspect:
                validity_reason = "Non-company product"

            if validity_reason is None and case_displayed_mahs:
                if any(name and MY_COMPANY_NAME.lower() not in name.lower() for name in case_displayed_mahs):
                    validity_reason = "Non-company product"

            # --------------------- VALIDITY LOGIC (UPDATED) ---------------------

            validity_reason: Optional[str] = None
            validity_comment: Optional[str] = None
            
            has_any_suspect = bool(suspect_ids)
            has_celix_suspect = bool(case_products_norm)
            
            # Rule 1: Patient details
            if not has_any_patient_detail:
                validity_reason = "No patient details"
            
            # Rule 2: Suspect but non-company
            if validity_reason is None and has_any_suspect and not has_celix_suspect:
                validity_reason = "Non-company product"
            
            # Rule 3: MAH check
            if validity_reason is None and case_displayed_mahs:
                if any(
                    name and MY_COMPANY_NAME.lower() not in name.lower()
                    for name in case_displayed_mahs
                ):
                    validity_reason = "Non-company product"
            
            # Rule 4: Product launch status (FIXED for multi-drug)
            if validity_reason is None and case_drug_dates_display:
            
                any_valid_drug = False
                non_valid_products = []
            
                for prod, strength_mg, sdt, edt in case_drug_dates_display:
                    status = get_launch_status(prod)
            
                    if status in ("yet", "awaited"):
                        non_valid_products.append(prod)
                    else:
                        any_valid_drug = True
            
                # Case-level decision
                if not any_valid_drug:
                    # ALL suspect drugs are non-valid
                    validity_reason = "Product not Launched"
                else:
                    # Mixed validity → case is VALID
                    if non_valid_products:
                        non_valid_list = ", ".join(sorted(set(non_valid_products)))
                        validity_comment = (
                            "Note: The following suspect drug(s) were assessed as non-valid "
                            f"due to launch status: {non_valid_list}. "
                            "Case validity is based on the presence of at least one valid suspect drug."
                        )

            earliest_launch_dt = None
            for prod, strength_mg, sdt, edt in case_drug_dates_display:
                if prod:
                    ld = get_launch_date(prod, strength_mg)
                    if ld:
                        earliest_launch_dt = ld if (earliest_launch_dt is None or ld < earliest_launch_dt) else earliest_launch_dt

            frd_raw_obj = parse_date_obj(global_dates["FRD_raw"]) if global_dates["FRD_raw"] else None
            lrd_raw_obj = parse_date_obj(global_dates["LRD_raw"]) if global_dates["LRD_raw"] else None
            exposure_reasons = []
            if validity_reason is None and earliest_launch_dt is not None:
                if frd_raw_obj and frd_raw_obj < earliest_launch_dt:
                    exposure_reasons.append("FRD")
                if lrd_raw_obj and lrd_raw_obj < earliest_launch_dt:
                    exposure_reasons.append("LRD")
                event_prior = any(
                    (evt_start and evt_start < earliest_launch_dt) or
                    (evt_stop and evt_stop < earliest_launch_dt)
                    for _, evt_start, evt_stop in case_event_dates
                )
                if event_prior:
                    exposure_reasons.append("Event")
                drug_prior = any(
                    (drug_start and drug_start < earliest_launch_dt)
                    for prod, _, drug_start, _ in case_drug_dates_display
                    if prod
                )
                if drug_prior:
                    exposure_reasons.append("Drug")
                if exposure_reasons:
                    validity_reason = f"Drug exposure prior to Launch; {', '.join(sorted(set(exposure_reasons)))}"

            narrative_elem = root.find('.//hl7:code[@code="PAT_ADV_EVNT"]/../hl7:text', ns)
            narrative_full_raw = narrative_elem.text if narrative_elem is not None else ''
            narrative_full = clean_value(narrative_full_raw)

            # OFEV check applies only for Nintedanib suspect product(s).
            has_ofev_in_narrative = bool(re.search(r'\bofev\b', normalize_text(narrative_full)))
            has_nintedanib_suspect = ('nintedanib' in case_products_norm)
            if has_ofev_in_narrative and has_nintedanib_suspect:
                adjusted_assessments: List[Tuple[str, str]] = []
                nintedanib_rows_found = False
                for nm, rsn in displayed_drugs_assessment:
                    if re.search(r'\bnintedanib\b', normalize_text(nm)):
                        nintedanib_rows_found = True
                        rsn = "Non-company product OFEV"
                    adjusted_assessments.append((nm, rsn))
                displayed_drugs_assessment = adjusted_assessments

                # Preserve existing case logic: make case non-valid only when
                # no suspect drug remains valid.
                if validity_reason != "No patient details" and nintedanib_rows_found:
                    if len(displayed_drugs_assessment) == 1 or all(rsn for _, rsn in displayed_drugs_assessment):
                        validity_reason = "Non-company product OFEV"

            # Nullification/amendment check applies to every case.
            has_nullification_or_amendment = False
            nullification_reason_elem = root.find(
                './/hl7:code[@displayName="nullificationAmendmentReason"]/../hl7:value/hl7:originalText',
                ns
            )
            if (
                nullification_reason_elem is not None
                and nullification_reason_elem.text
                and nullification_reason_elem.text.strip()
            ):
                has_nullification_or_amendment = True
            else:
                nullification_code_elem = root.find(
                    './/hl7:code[@displayName="nullificationAmendmentCode"]/../hl7:value',
                    ns
                )
                if nullification_code_elem is not None:
                    code_val = clean_value(nullification_code_elem.attrib.get('code', ''))
                    if code_val:
                        has_nullification_or_amendment = True

            if has_nullification_or_amendment:
                comments.append("check nullification or amendment comment")

            validity_value = f"Non-Valid ({validity_reason})" if validity_reason else "Valid"

            if comments and validity_reason is None:
                validity_value = "Kindly check comment and assess validity manually"

            if isinstance(validity_value, str) and validity_value.startswith("Non-Valid"):
                reportability = "NA"

            is_non_valid_case = isinstance(validity_value, str) and validity_value.startswith("Non-Valid")

            report_date_parts = []
            if frd_disp:
                report_date_parts.append(f"FRD: {frd_disp}")
            if lrd_disp:
                report_date_parts.append(f"LRD: {lrd_disp}")
            if td_disp:
                report_date_parts.append(f"TD: {td_disp}")
            report_date_display = "\n".join(report_date_parts)

            per_drug_nonvalid_lines = [f"{nm}: {rsn}" for nm, rsn in displayed_drugs_assessment if rsn]
            show_per_drug_comment = (len(displayed_drugs_assessment) > 1) and (len(per_drug_nonvalid_lines) == len(displayed_drugs_assessment))
            if show_per_drug_comment and isinstance(validity_value, str) and validity_value.startswith("Non-Valid"):
                validity_value = f"{validity_value} \n Drug-wise: " + "; ".join(per_drug_nonvalid_lines)

            # ---- LISTEDNESS (EVENT ONLY; PER-PRODUCT SUMMARY WHEN MULTI-PRODUCT) ----
            event_wise_listedness_display = ""
            if not is_non_valid_case and event_llts_norm:
                if len(case_products_norm) <= 1:
                    lines = []
                    products_to_check = list(case_products_norm) if case_products_norm else []
                    for i, llt_norm in enumerate(event_llts_norm, start=1):
                        results = [assess(listedness_pairs, pnorm, llt_norm)[0] for pnorm in products_to_check]
                        status = 'Expected' if 'Expected' in results else 'Unexpected'
                        lines.append(f"Event {i}: {status}")
                    event_wise_listedness_display = "\n".join(lines)
                else:
                    prod_lines: List[str] = []
                    for pnorm in sorted(list(case_products_norm), key=lambda k: product_norm_to_pretty.get(k, k)):
                        pretty = product_norm_to_pretty.get(pnorm, pnorm)
                        statuses = []
                        for i, llt_norm in enumerate(event_llts_norm, start=1):
                            status, _ = assess(listedness_pairs, pnorm, llt_norm)
                            statuses.append(f"Event {i}: {status}")
                        prod_lines.append(f"{pretty} - " + "; ".join(statuses))
                    event_wise_listedness_display = "\n".join(prod_lines)

            molecule_names_for_filename = [pnorm.title() for pnorm in sorted(list(case_products_norm))]
            suspect_products_preferred = [
                product_norm_to_pretty.get(pnorm, pnorm.title())
                for pnorm in sorted(list(case_products_norm), key=lambda k: product_norm_to_pretty.get(k, k))
            ]
            suspect_product_display = "; ".join(suspect_products_preferred) if suspect_products_preferred else "; ".join(dict.fromkeys(suspect_product_names_all))
            seriousness_value = "Serious" if case_has_serious_event else "Non-serious"
            comment_value = "; ".join(sorted(set(comments))) if comments else ""
            tracker_comment_value = build_tracker_comment(validity_value, comment_value)
            suggested_xml_filename = build_suggested_xml_filename(
                sender_id,
                validity_value,
                case_has_serious_event,
                molecule_names_for_filename,
                fallback_stem=original_file_stem
            )
            renamed_xml_meta[idx] = {
                "xml_bytes": xml_bytes,
                "safety_id": sender_id,
                "is_serious": case_has_serious_event,
                "molecule_names": molecule_names_for_filename,
                "fallback_stem": original_file_stem,
            }

            all_rows_display.append({
                'SL No': idx,
                'Date': current_date,
                'Sender ID': sender_id,
                'Report Date': report_date_display,
                'Case Age (days)': case_age_days,
                'Reporter Qualification': reporter_qualification,
                'Patient Detail': patient_detail,
                'Product Detail': "\n ".join(product_details_list),
                'Event Details': event_details_combined_display,
                'Listedness': ('' if is_non_valid_case else event_wise_listedness_display),
                'Narrative': narrative_full,
                'Validity': validity_value,
                'Comment': comment_value,
                'Reportability': reportability,
                'Receipt Date': current_date,
                'Referance ID': sender_id,
                'IRD': td_disp,
                'Suspect Product': suspect_product_display,
                'Seriousness': seriousness_value,
                'Safety Report ID': 'WIP',
                'Tracker Comment': tracker_comment_value,
                'Suggested XML File Name': suggested_xml_filename,
                'Parsing Warnings': "; ".join(warnings) if warnings else ""
            })

            parsed_rows += 1
            progress.progress(idx / total_files)

        st.success(f"Parsing complete \u2705 — Files processed: {total_files}, Rows created: {parsed_rows}")

# -------------------------------- UI: Export & Edit ---------------------------
with tab2:
    st.markdown("### \U0001F4CB Parsed Data Table \U0001F4C3")
    if all_rows_display:
        df_display = pd.DataFrame(all_rows_display)

        show_full_narrative = st.checkbox("Show full narrative (may be long)", value=True)
        if not show_full_narrative:
            df_display['Narrative'] = df_display['Narrative'].astype(str).str.slice(0, 1000)

        preferred_order = [
            'SL No','Date','Sender ID','Report Date','Case Age (days)','Reporter Qualification',
            'Patient Detail','Product Detail','Event Details','Listedness','Narrative',
            'Validity','Comment','Reportability','Suggested XML File Name','Parsing Warnings'
        ]
        df_display = df_display[[c for c in preferred_order if c in df_display.columns]]

        editable_columns = [c for c in df_display.columns if c != 'Validity']
        edited_df = st.data_editor(
            df_display,
            num_rows="dynamic",
            use_container_width=True,
            disabled=editable_columns,
            key="parsed_table_editor"
        )

        # Recompute suggested XML file name after manual edits to Validity.
        if 'Suggested XML File Name' in edited_df.columns:
            used_edited_names: Set[str] = set()
            recomputed_names = []
            for _, row in edited_df.iterrows():
                sl_no = row.get('SL No')
                meta = renamed_xml_meta.get(sl_no)
                if meta:
                    new_name = build_suggested_xml_filename(
                        meta['safety_id'],
                        row.get('Validity', ''),
                        meta['is_serious'],
                        meta['molecule_names'],
                        fallback_stem=meta['fallback_stem']
                    )
                    new_name = make_unique_filename(new_name, used_edited_names)
                else:
                    new_name = row.get('Suggested XML File Name', '')
                recomputed_names.append(new_name)
            edited_df['Suggested XML File Name'] = recomputed_names

        excel_buffer = io.BytesIO()
        with pd.ExcelWriter(excel_buffer, engine='openpyxl') as writer:
            edited_df.to_excel(writer, index=False, sheet_name="Parsed Data")
        st.download_button("⬇️ Download Excel", excel_buffer.getvalue(), "parsed_data.xlsx")

        if not edited_df.empty and renamed_xml_meta:
            zip_buffer = io.BytesIO()
            with zipfile.ZipFile(zip_buffer, mode="w", compression=zipfile.ZIP_DEFLATED) as zip_file:
                for _, row in edited_df.iterrows():
                    sl_no = row.get('SL No')
                    meta = renamed_xml_meta.get(sl_no)
                    if not meta:
                        continue
                    file_name = row.get('Suggested XML File Name', '')
                    if not file_name:
                        file_name = build_suggested_xml_filename(
                            meta['safety_id'],
                            row.get('Validity', ''),
                            meta['is_serious'],
                            meta['molecule_names'],
                            fallback_stem=meta['fallback_stem']
                        )
                    zip_file.writestr(file_name, meta['xml_bytes'])
            st.download_button("⬇️ Download Renamed XML Files (.zip)", zip_buffer.getvalue(), "renamed_xml_files.zip", mime="application/zip")
    else:
        st.info("No data available yet. Please upload files in the first tab.")

with tab3:
    st.markdown("### 📋 Tracker Copy Format")
    if all_rows_display:
        tracker_source_df = pd.DataFrame(all_rows_display)

        if edited_df is not None and not edited_df.empty and 'SL No' in edited_df.columns and 'SL No' in tracker_source_df.columns:
            tracker_source_df = tracker_source_df.set_index('SL No')
            edited_updates = edited_df.set_index('SL No')
            for col in ['Validity', 'Suggested XML File Name']:
                if col in tracker_source_df.columns and col in edited_updates.columns:
                    tracker_source_df[col] = edited_updates[col]
            tracker_source_df = tracker_source_df.reset_index()

        if 'IRD' not in tracker_source_df.columns and 'Report Date' in tracker_source_df.columns:
            tracker_source_df['IRD'] = tracker_source_df['Report Date'].apply(extract_td_value)
        if 'Referance ID' not in tracker_source_df.columns and 'Sender ID' in tracker_source_df.columns:
            tracker_source_df['Referance ID'] = tracker_source_df['Sender ID']
        if 'Receipt Date' not in tracker_source_df.columns and 'Date' in tracker_source_df.columns:
            tracker_source_df['Receipt Date'] = tracker_source_df['Date']
        if 'Tracker Comment' not in tracker_source_df.columns:
            tracker_source_df['Tracker Comment'] = tracker_source_df.apply(
                lambda row: build_tracker_comment(row.get('Validity', ''), row.get('Comment', '')),
                axis=1,
            )
        if 'Seriousness' not in tracker_source_df.columns:
            tracker_source_df['Seriousness'] = ''
        if 'Suspect Product' not in tracker_source_df.columns:
            tracker_source_df['Suspect Product'] = ''

        tracker_df = pd.DataFrame({
            'Receipt Date': tracker_source_df['Receipt Date'].fillna(''),
            'Source': 'MHRA',
            'Referance ID': tracker_source_df['Referance ID'].fillna(''),
            'IRD': tracker_source_df['IRD'].fillna(''),
            'Suspect Product': tracker_source_df['Suspect Product'].apply(simplify_tracker_product).fillna(''),
            'Validity': tracker_source_df['Validity'].apply(simplify_tracker_validity).fillna(''),
            'Seriousness': tracker_source_df['Validity'].apply(lambda x: 'NA' if clean_value(x).startswith('Non-Valid') else '').where(
                tracker_source_df['Validity'].apply(lambda x: clean_value(x).startswith('Non-Valid')),
                tracker_source_df['Seriousness'].fillna('')
            ),
            'Safety Report ID': tracker_source_df['Validity'].apply(lambda x: 'NA' if clean_value(x).startswith('Non-Valid') else 'WIP'),
            'Comment': tracker_source_df['Tracker Comment'].fillna(''),
        })

        st.dataframe(tracker_df, use_container_width=True, hide_index=True)

        clipboard_ready_df = tracker_df.astype(str).replace({'nan': '', 'None': ''}, regex=False)
        copy_text = '\n'.join(
            '\t'.join(value.replace('\n', ' ').strip() for value in row)
            for row in clipboard_ready_df.itertuples(index=False, name=None)
        )

        st.caption('Copy output below is headerless and tab-separated, ready to paste into the tracker.')
        st.text_area('Preview (without header)', copy_text, height=220, key='tracker_copy_preview')

        copy_html = f"""
        <textarea id="tracker-copy-source" style="position:absolute; left:-9999px; top:-9999px;">{html.escape(copy_text)}</textarea>
        <button onclick="navigator.clipboard.writeText(document.getElementById('tracker-copy-source').value).then(function(){{document.getElementById('tracker-copy-status').innerText='Copied to clipboard!';}}).catch(function(){{document.getElementById('tracker-copy-status').innerText='Copy failed. Please copy from the preview box.';}});" style="padding:0.5rem 0.9rem; border:1px solid #d0d0d0; border-radius:0.4rem; background:#f4f4f4; cursor:pointer;">Copy without header</button>
        <div id="tracker-copy-status" style="margin-top:0.5rem; font-family:sans-serif; color:#333;"></div>
        """
        st.components.v1.html(copy_html, height=90)
    else:
        st.info("No data available yet. Please upload files in the first tab.")

st.markdown("""
**Developed by Jagamohan**
_Disclaimer: App is in developmental stage, validate before using the data._
""", unsafe_allow_html=True)


with tab4:
    st.markdown("### 📨 ACK XML Upload, Rename & Download")
    st.caption("Upload ACK XML files. The app uses a single Case ID (Sender ID / Safety Report ID treated as the same), renames files as CaseID_Ack.xml, and maps validity from the first tab using the same Case ID.")

    first_tab_validity_lookup = build_first_tab_validity_lookup(all_rows_display)
    if not first_tab_validity_lookup:
        st.info("Tip: Upload and parse the main case XMLs in the first tab first if you want ACK validity to exactly follow the first-tab categorization.")

    ack_uploaded_files = st.file_uploader(
        "Upload ACK XML files",
        type=["xml"],
        accept_multiple_files=True,
        help="Upload one or more acknowledgement XML files.",
        key=f"ack_xml_uploader_{st.session_state.get('uploader_version', 0)}"
    )

    if ack_uploaded_files:
        ack_rows: List[Dict[str, str]] = []
        ack_download_items: List[Dict[str, object]] = []
        used_ack_names: Set[str] = set()

        for ack_idx, ack_file in enumerate(ack_uploaded_files, start=1):
            ack_bytes = ack_file.getvalue()
            ack_meta = parse_ack_xml_metadata(ack_bytes, ack_file.name)
            case_id = clean_value(ack_meta.get('Case ID', ''))
            mapped_validity = first_tab_validity_lookup.get(case_id, '')
            zip_group = ack_zip_group_from_first_tab_validity(mapped_validity)

            unique_name = make_unique_filename(ack_meta['Renamed File Name'], used_ack_names)
            ack_meta['Renamed File Name'] = unique_name
            ack_meta['Mapped First-Tab Validity'] = mapped_validity or 'Case ID not found in first tab'
            ack_meta['ZIP Group'] = zip_group
            ack_meta['SL No'] = ack_idx
            ack_rows.append(ack_meta)
            ack_download_items.append({
                'SL No': ack_idx,
                'file_name': unique_name,
                'xml_bytes': ack_bytes,
                'zip_group': zip_group,
                'case_id': case_id,
                'mapped_validity': mapped_validity,
                'ack_status': ack_meta['ACK Derived Status'],
            })

        ack_df = pd.DataFrame(ack_rows)
        ack_display_order = [
            'SL No', 'Original File Name', 'Case ID', 'Mapped First-Tab Validity',
            'ACK Code', 'ACK Derived Status', 'Error Comment', 'Renamed File Name'
        ]
        ack_df = ack_df[[c for c in ack_display_order if c in ack_df.columns]]
        st.dataframe(ack_df, use_container_width=True, hide_index=True)

        valid_items = [item for item in ack_download_items if item['zip_group'] == 'Valid']
        non_valid_items = [item for item in ack_download_items if item['zip_group'] == 'Non-Valid']
        other_items = [item for item in ack_download_items if item['zip_group'] == 'Other']

        summary_col1, summary_col2, summary_col3, summary_col4 = st.columns(4)
        summary_col1.metric('Total ACK XMLs', len(ack_download_items))
        summary_col2.metric('Valid (from Tab 1)', len(valid_items))
        summary_col3.metric('Non-Valid (from Tab 1)', len(non_valid_items))
        summary_col4.metric('Not Mapped / Manual', len(other_items))

        zip_col1, zip_col2 = st.columns(2)

        valid_zip_buffer = io.BytesIO()
        with zipfile.ZipFile(valid_zip_buffer, mode='w', compression=zipfile.ZIP_DEFLATED) as zip_file:
            for item in valid_items:
                zip_file.writestr(item['file_name'], item['xml_bytes'])
        zip_col1.download_button(
            '⬇️ Download Valid ACK XMLs (.zip)',
            valid_zip_buffer.getvalue(),
            'valid_ack_xmls.zip',
            mime='application/zip',
            disabled=(len(valid_items) == 0),
            key='download_valid_ack_zip'
        )

        non_valid_zip_buffer = io.BytesIO()
        with zipfile.ZipFile(non_valid_zip_buffer, mode='w', compression=zipfile.ZIP_DEFLATED) as zip_file:
            for item in non_valid_items:
                zip_file.writestr(item['file_name'], item['xml_bytes'])
        zip_col2.download_button(
            '⬇️ Download Non-Valid ACK XMLs (.zip)',
            non_valid_zip_buffer.getvalue(),
            'non_valid_ack_xmls.zip',
            mime='application/zip',
            disabled=(len(non_valid_items) == 0),
            key='download_non_valid_ack_zip'
        )

        if other_items:
            st.caption('Files marked as Not Mapped / Manual are still available through the individual Download buttons below, but are excluded from the Valid and Non-Valid zip downloads.')

        st.markdown('#### Individual File Downloads')
        for item in ack_download_items:
            row_col1, row_col2, row_col3, row_col4 = st.columns([1.3, 1.8, 2.6, 1.2])
            row_col1.write(f"**{item['zip_group']}**")
            row_col2.write(f"**Case ID:** {item['case_id'] or 'Not found'}")
            row_col3.write(f"**Tab 1 Validity:** {item['mapped_validity'] or 'Case ID not found in first tab'}")
            row_col4.download_button(
                'Download',
                item['xml_bytes'],
                item['file_name'],
                mime='application/xml',
                key=f"download_ack_{item['SL No']}"
            )
    else:
        st.info('Upload ACK XML files to generate CaseID_Ack filenames and download options.')



# Password-protected PT listedness maintenance
try:
    _triage_products = list(locals().get("product_norm_to_pretty", {}).values())
    _triage_pts = locals().get("event_pts", [])
    render_listedness_updater("triage_listedness", _triage_products, _triage_pts)
except Exception as _listedness_ui_error:
    st.warning(f"Listedness updater unavailable: {_listedness_ui_error}")
