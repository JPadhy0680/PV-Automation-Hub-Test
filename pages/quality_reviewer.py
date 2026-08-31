# qc_twofile_compare_tabular.py
import streamlit as st
from shared_context import init_shared_context
from listedness_service import dataframe_to_index, assess, has_exact_match, render_missing_listedness_update
init_shared_context()
import pandas as pd
import xml.etree.ElementTree as ET
from datetime import datetime, date
import io, re, calendar, zipfile
from typing import Optional, Dict, Any, List, Tuple, Set
from pathlib import Path

# ---------------- UI setup ----------------
st.markdown('<h1 class="app-title">XML Qualty Reviewer - Automation</h1>', unsafe_allow_html=True)

# ---------------- Utilities ----------------
NS = {'hl7': 'urn:hl7-org:v3', 'xsi': 'http://www.w3.org/2001/XMLSchema-instance'}
UNKNOWN_TOKENS = {"unk", "asku", "unknown"}
CELIX_PRODUCTS = {
    "abiraterone", "apixaban", "apremilast", "bexarotene", "brivaracetam", "clobazam", "clonazepam",
    "cyanocobalamin", "dabigatran", "dapagliflozin", "dimethyl fumarate", "edoxaban", "empagliflozin",
    "famotidine", "fesoterodine", "icatibant", "itraconazole", "linagliptin", "linagliptin + metformin",
    "metformin", "letrozole", "nintedanib", "pirfenidone", "raltegravir", "ranolazine", "rivaroxaban",
    "safinamide", "saxagliptin", "sitagliptin", "sacubritril + valsartan", "sacubitril + valsartan",
    "sacubritril", "sacubitril", "valsartan", "tamsulosin + solifenacin", "tapentadol", "ticagrelor",
    "tamsulosin", "solifenacin"
}

def celix_product_name(product):
    """Return the Celix molecule matched from a suspect product, else blank."""
    if "suspect" not in normalize_text(product.get("Type", "")):
        return ""
    candidates = [product.get("Ingredient", ""), product.get("Drug Name", ""), product.get("Drug", "")]
    combined = normalize_text(" ".join(str(v or "") for v in candidates))
    for molecule in sorted(CELIX_PRODUCTS, key=len, reverse=True):
        if normalize_text(molecule) in combined:
            return molecule.upper()
    return ""


# Admin identifiers
SENDER_ID_OID = "2.16.840.1.113883.3.989.2.1.3.1"  # Sender ID
WWID_OID = "2.16.840.1.113883.3.989.2.1.3.2"       # WWID
FIRST_SENDER_OID = "2.16.840.1.113883.3.989.2.1.1.3"  # First sender of case (1=Regulator, 2=Other)
FIRST_SENDER_MAP = {"1": "Regulator", "2": "Other"}

# Reporter qualification OID
REPORTER_QUAL_OID = "2.16.840.1.113883.3.989.2.1.1.6"
REPORTER_MAP = {
    "1": "Physician",
    "2": "Pharmacist",
    "3": "Other health professional",
    "4": "Lawyer",
    "5": "Consumer or other non-health professional",
}

# Reporter SOURCE anchor OID
REPORT_SOURCE_OID = "2.16.840.1.113883.3.989.2.1.1.22"  # displayName="sourceReport"

# Amendment / Nullification investigation-characteristic OID
NULLIFICATION_AMENDMENT_OID = "2.16.840.1.113883.3.989.2.1.1.23"

# Patient OIDs
AGE_OID = "2.16.840.1.113883.3.989.2.1.1.19"
PATIENT_RECORD_OID = "2.16.840.1.113883.3.989.2.1.3.7"

# ---- Action Taken ----
ACTION_TAKEN_OID = "2.16.840.1.113883.3.989.2.1.1.15"
ACTION_TAKEN_MAP = {
    "1": "Drug withdrawn",
    "2": "Dose reduced",
    "3": "Dose increased",
    "4": "Dose not changed",
    "0": "Unknown",
    "9": "Not applicable",
}

# MedDRA / Clinical section OIDs
MEDDRA_LLT_OID = "2.16.840.1.113883.6.163"  # LLT codes in observations
MH_SECTION_OID = "2.16.840.1.113883.3.989.2.1.1.20"  # clinical sections
STATUS_OID = "2.16.840.1.113883.3.989.2.1.1.19"      # status & flags (causality/intervention/…)
INTERVENTION_CHAR_CODE = "20"
CAUSALITY_CODE = "39"

# TD priority paths (for Day Zero: Source=TD, Processed=LRD)
TD_PATHS = [
    './/hl7:transmissionWrapper/hl7:creationTime',
    './/hl7:ControlActProcess/hl7:effectiveTime',
    './/hl7:ClinicalDocument/hl7:effectiveTime',
    './/hl7:creationTime',
]

# --- UI styling ---
BOX_CSS = """
<style>
.block-container { padding-top: 2rem; }
.app-title { font-size: 34px; line-height: 1.35; font-weight: 700; color: #1f2937; margin: 0 0 1.4rem 0; padding: 0; letter-spacing: 0.01em; }
.qc-card { border: 1px solid #e5e7eb; border-radius: 12px; padding: 14px 16px; background: #ffffff; box-shadow: 0 1px 3px rgba(0,0,0,0.04); margin-bottom: 12px; }
.qc-kpi { border-radius: 12px; padding: 12px 14px; border: 1px solid #e5e7eb; background: #f8fbff; }
.qc-kpi-label { color: #6b7280; font-size: 12px; text-transform: uppercase; letter-spacing: .04em; }
.qc-kpi-value { color: #111827; font-size: 16px; font-weight: 700; margin-top: 4px; }
.qc-result { border-radius: 10px; padding: 16px 18px; font-size: 20px; font-weight: 800; text-align:center; border: 1px solid #e5e7eb; }
.qc-result-certain { background:#dcfce7; color:#166534; border-color:#86efac; }
.qc-result-probable { background:#e0f2fe; color:#075985; border-color:#7dd3fc; }
.qc-result-possible { background:#fef9c3; color:#854d0e; border-color:#fde047; }
.qc-result-unlikely { background:#ffedd5; color:#9a3412; border-color:#fdba74; }
.qc-result-unassessable { background:#f3f4f6; color:#374151; border-color:#d1d5db; }
.qc-result-empty { background:#fafafa; color:#6b7280; border-color:#e5e7eb; }
</style>
"""
st.markdown(BOX_CSS, unsafe_allow_html=True)

# ---------------- Small helpers ----------------
def _digits_only(s: str) -> str:
    return re.sub(r"\D", "", (s or "").strip())


def format_date(date_str: str) -> str:
    if not date_str:
        return ""
    digits = _digits_only(date_str)
    try:
        if len(digits) >= 8:
            return datetime.strptime(digits[:8], "%Y%m%d").strftime("%d-%b-%Y")
        elif len(digits) >= 6:
            return datetime.strptime(digits[:6], "%Y%m").strftime("%b-%Y")
        elif len(digits) >= 4:
            return digits[:4]
    except Exception:
        pass
    return ""


def parse_date_obj(date_str: str) -> Optional[date]:
    if not date_str:
        return None
    digits = _digits_only(date_str)
    try:
        if len(digits) >= 8:
            return datetime.strptime(digits[:8], "%Y%m%d").date()
        elif len(digits) >= 6:
            y, m = int(digits[:4]), int(digits[4:6])
            last = calendar.monthrange(y, m)[1]
            return date(y, m, last)
        elif len(digits) >= 4:
            y = int(digits[:4])
            return date(y, 12, 31)
    except Exception:
        pass
    return None


def clean_value(v: Any) -> str:
    if v is None:
        return ""
    s = str(v).strip()
    return "" if (not s or s.lower() in UNKNOWN_TOKENS) else s


def normalize_text(s: str) -> str:
    s = (s or "").lower()
    s = re.sub(r'[^a-z0-9\s+\-]', ' ', s)
    return re.sub(r'\s+', ' ', s).strip()


def map_gender(code: str) -> str:
    return {"1": "Male", "2": "Female", "M": "Male", "F": "Female"}.get(code, "Unknown")


def local_name(tag: str) -> str:
    return tag.split('}')[-1] if '}' in tag else tag


def get_text(elem) -> str:
    return (elem.text or "").strip() if (elem is not None and elem.text) else ""


def read_text_or_mask(elem: Optional[ET.Element]) -> str:
    if elem is None:
        return ""
    if elem.attrib.get('nullFlavor') == 'MSK':
        return "Masked"
    return (elem.text or "").strip()

# ✅ Simple finders with fixed namespace
def find_first(root, xpath, ns=None) -> Optional[ET.Element]:
    return root.find(xpath, NS)


def findall(root, xpath, ns=None) -> List[ET.Element]:
    return root.findall(xpath, NS)


def mismatch_marker(a: Any, b: Any, is_date=False) -> str:
    if is_date:
        da, db = parse_date_obj(a or ""), parse_date_obj(b or "")
        if da == db and da is not None:
            return ""
    return " 🔴" if (str(a) or "") != (str(b) or "") else ""


def has_value(x: str) -> bool:
    return bool((x or "").strip())


def safe_disp(v: str) -> str:
    return v if v else "—"

# 🔒 Convert any data to readable text for display tables
def _textify(x: Any) -> str:
    """Convert any value (list/dict/number/None) to a readable string for display."""
    if x is None:
        return ""
    if isinstance(x, (list, tuple, set)):
        return "; ".join(str(i) for i in x if i is not None)
    if isinstance(x, dict):
        return "; ".join(f"{k}: {v}" for k, v in x.items())
    return str(x)

# ---------------- Dependency-free XLSX reader ----------------
def _col_letters_to_index(col_letters: str) -> int:
    res = 0
    for ch in col_letters:
        if not ch.isalpha():
            break
        res = res * 26 + (ord(ch.upper()) - ord('A') + 1)
    return res - 1


def _parse_sheet_xml(sheet_xml_bytes: bytes, shared_strings: List[str]) -> pd.DataFrame:
    from xml.etree.ElementTree import fromstring

    root = fromstring(sheet_xml_bytes)
    ns = {'a': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
    rows: List[List[str]] = []
    max_col = 0

    for row in root.findall('.//a:sheetData/a:row', ns):
        values: Dict[int, str] = {}
        for c in row.findall('a:c', ns):
            r = c.attrib.get('r', '')
            col_letters = ''.join([ch for ch in r if ch.isalpha()]) or 'A'
            col_idx = _col_letters_to_index(col_letters)
            if col_idx > max_col:
                max_col = col_idx
            value = ""
            t = c.attrib.get('t', '')
            v = c.find('a:v', ns)
            is_node = c.find('a:is', ns)
            if t == 's':
                if v is not None and v.text and v.text.isdigit():
                    ss_idx = int(v.text)
                    if 0 <= ss_idx < len(shared_strings):
                        value = shared_strings[ss_idx]
            elif t == 'inlineStr' and is_node is not None:
                tnode = is_node.find('a:t', ns)
                value = (tnode.text or '') if tnode is not None else ''
            else:
                value = (v.text or '') if v is not None else ''
            values[col_idx] = value
        if values:
            row_list = ["" for _ in range(max_col + 1)]
            for idx, val in values.items():
                if idx <= max_col:
                    row_list[idx] = val
            rows.append(row_list)

    if not rows:
        return pd.DataFrame()
    header = rows[0]
    data = rows[1:] if len(rows) > 1 else []
    header = [h if h else f"col_{i+1}" for i, h in enumerate(header)]
    return pd.DataFrame(data, columns=header)


def _read_xlsx_no_openpyxl(uploaded_file) -> pd.DataFrame:
    data = uploaded_file.read()
    zf = zipfile.ZipFile(io.BytesIO(data))

    shared_strings: List[str] = []
    try:
        sst = zf.read('xl/sharedStrings.xml')
        sroot = ET.fromstring(sst)
        s_ns = {'a': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
        for si in sroot.findall('.//a:si', s_ns):
            parts: List[str] = []
            for tnode in si.findall('.//a:t', s_ns):
                parts.append(tnode.text or '')
            shared_strings.append(''.join(parts))
    except KeyError:
        shared_strings = []

    sheet_bytes = None
    try:
        sheet_bytes = zf.read('xl/worksheets/sheet1.xml')
    except KeyError:
        for name in zf.namelist():
            if name.startswith('xl/worksheets/') and name.endswith('.xml'):
                sheet_bytes = zf.read(name)
                break
    if sheet_bytes is None:
        raise ValueError("No worksheet XML found in XLSX.")

    return _parse_sheet_xml(sheet_bytes, shared_strings)

# ---- MedDRA mapping loader ----
def load_meddra_mapping(uploaded_file) -> Dict[str, Dict[str, str]]:
    if not uploaded_file:
        return {}
    fname = (uploaded_file.name or "").lower()
    try:
        if fname.endswith(".csv"):
            df = pd.read_csv(uploaded_file)
        elif fname.endswith(".xlsx"):
            df = _read_xlsx_no_openpyxl(uploaded_file)
        else:
            st.error("Unsupported MedDRA file format. Please upload .xlsx or .csv")
            return {}
    except Exception as e:
        st.error(f"Could not read MedDRA mapping: {e}")
        return {}

    cols = {c.strip().lower(): c for c in df.columns}
    required = ['llt code', 'llt term', 'pt code', 'pt term']
    if not all(k in cols for k in required):
        st.error("MedDRA mapping must contain columns: LLT Code, LLT Term, PT Code, PT Term")
        return {}

    mapping: Dict[str, Dict[str, str]] = {}
    for _, row in df.iterrows():
        try:
            llt_code = str(row[cols['llt code']]).strip()
            if not llt_code or llt_code.lower() in {"nan", "none"}:
                continue
            mapping[llt_code] = {
                "LLT Term": str(row[cols['llt term']]).strip() if pd.notna(row[cols['llt term']]) else "",
                "PT Code": str(row[cols['pt code']]).strip() if pd.notna(row[cols['pt code']]) else "",
                "PT Term": str(row[cols['pt term']]).strip() if pd.notna(row[cols['pt term']]) else "",
            }
        except Exception:
            continue
    st.success(f"MedDRA Mapping Loaded — {len(mapping):,} LLT rows")
    return mapping

# ---- Value helpers ----
def read_numeric_with_unit(value_node: Optional[ET.Element]) -> str:
    if value_node is None:
        return ""
    v = (value_node.attrib.get('value') or '').strip()
    u = (value_node.attrib.get('unit') or '').strip()
    if v or u:
        return f"{v} {u}".strip()
    center = value_node.find('.//hl7:center', NS)
    if center is not None:
        cv = (center.attrib.get('value') or '').strip()
        cu = (center.attrib.get('unit') or '').strip()
        if cv or cu:
            return f"{cv} {cu}".strip()
    low = value_node.find('.//hl7:low', NS)
    high = value_node.find('.//hl7:high', NS)
    lv = (low.attrib.get('value') or '').strip() if low is not None else ''
    lu = (low.attrib.get('unit') or '').strip() if low is not None else ''
    hv = (high.attrib.get('value') or '').strip() if high is not None else ''
    hu = (high.attrib.get('unit') or '').strip() if high is not None else ''
    if lv or hv:
        lo = f"{lv} {lu}".strip() if (lv or lu) else ""
        hi = f"{hv} {hu}".strip() if (hv or hu) else ""
        return f"{lo} – {hi}".strip(' –')
    return get_text(value_node)

# ---------------- Admin extraction ----------------
def extract_id_by_oid(root: ET.Element, oid: str) -> str:
    e = find_first(root, f'.//hl7:id[@root="{oid}"]')
    return clean_value(e.attrib.get('extension', '')) if e is not None else ""


def extract_sender_id(root: ET.Element) -> str:
    return extract_id_by_oid(root, SENDER_ID_OID)


def extract_wwid(root: ET.Element) -> str:
    return extract_id_by_oid(root, WWID_OID)


def extract_first_sender_type(root: ET.Element) -> str:
    for el in root.iter():
        if local_name(el.tag) == 'code' and el.attrib.get('codeSystem') == FIRST_SENDER_OID:
            raw = (el.attrib.get('code') or "").strip()
            return FIRST_SENDER_MAP.get(raw, raw or "")
    return ""


def extract_td_frd_lrd(root: ET.Element) -> Dict[str, str]:
    out = {"TD_raw": "", "TD": "", "FRD_raw": "", "FRD": "", "LRD_raw": "", "LRD": ""}
    # TD
    for p in TD_PATHS:
        e = find_first(root, p)
        if e is not None:
            val = e.attrib.get('value') or get_text(e)
            if val:
                out["TD_raw"] = val
                out["TD"] = format_date(val)
                break
    # LRD
    for el in root.iter():
        if local_name(el.tag) == 'availabilityTime':
            v = el.attrib.get('value')
            if v:
                out["LRD_raw"] = v
                out["LRD"] = format_date(v)
                break
    # FRD (earliest <low/>)
    lows: List[str] = []
    for el in root.iter():
        if local_name(el.tag) == 'low':
            v = el.attrib.get('value')
            if v:
                lows.append(v)
    if lows:
        pairs = [(parse_date_obj(v), v) for v in lows if parse_date_obj(v)]
        if pairs:
            pairs.sort(key=lambda t: t[0])
            out["FRD_raw"] = pairs[0][1]
            out["FRD"] = format_date(pairs[0][1])
    return out

# ---------------- Patient extraction ----------------
def get_pq_value_by_code(root: ET.Element, display_name: Optional[str] = None, code_system_oid: Optional[str] = None) -> Tuple[str, str]:
    for obs in root.findall('.//hl7:observation', NS):
        code_el = obs.find('hl7:code', NS)
        if code_el is None:
            continue
        ok = False
        if display_name and (code_el.attrib.get('displayName') or '').strip().lower() == display_name.lower():
            ok = True
        if (not ok) and code_system_oid and (code_el.attrib.get('codeSystem') == code_system_oid):
            ok = True
        if not ok:
            continue
        val_el = obs.find('hl7:value', NS)
        if val_el is None:
            continue
        v = (val_el.attrib.get('value') or '').strip()
        u = (val_el.attrib.get('unit') or '').strip()
        return v, u
    return "", ""


def find_mask_aware_id_by_root(root: ET.Element, oid: str) -> str:
    for el in root.iter():
        if local_name(el.tag) != 'id':
            continue
        if el.attrib.get('root') == oid:
            if el.attrib.get('nullFlavor') == 'MSK':
                return "Masked"
            ext = (el.attrib.get('extension') or '').strip()
            return ext
    return ""

def extract_patient(root: ET.Element) -> Dict[str, str]:
    # Gender
    gender_elem = find_first(root, './/hl7:administrativeGenderCode')
    gender_code = gender_elem.attrib.get('code', '') if gender_elem is not None else ''
    gender = clean_value(map_gender(gender_code))

    # Age
    age_val, age_unit_raw = get_pq_value_by_code(root, display_name="age", code_system_oid=AGE_OID)
    unit_map = {'a': 'year', 'b': 'month'}
    age_unit_label = unit_map.get((age_unit_raw or '').lower(), age_unit_raw or '')
    age = ""
    if clean_value(age_val):
        age = clean_value(age_val)
        if clean_value(age_unit_label):
            age = f"{age} {age_unit_label}"

    # Age Group
    age_group_map = {
        "0": "Foetus", "1": "Neonate", "2": "Infant", "3": "Child",
        "4": "Adolescent", "5": "Adult", "6": "Elderly"
    }
    ag_elem = find_first(root, './/hl7:code[@displayName="ageGroup"]/../hl7:value')
    age_group = ""
    if ag_elem is not None:
        c = ag_elem.attrib.get('code', '')
        nf = ag_elem.attrib.get('nullFlavor', '')
        age_group = age_group_map.get(
            c,
            "[Masked/Unknown]" if (c in ["MSK", "UNK", "ASKU", "NI"] or nf in ["MSK", "UNK", "ASKU", "NI"]) else ""
        )

    # Weight
    w_el = find_first(root, './/hl7:code[@displayName="bodyWeight"]/../hl7:value')
    w_val = w_el.attrib.get('value', '') if w_el is not None else ''
    w_unit = w_el.attrib.get('unit', '') if w_el is not None else ''
    if not (w_val or w_unit):
        for obs in root.findall('.//hl7:observation', NS):
            val = obs.find('hl7:value', NS)
            if val is None:
                continue
            u = (val.attrib.get('unit') or '').strip().lower()
            if u in {'kg', 'lb', 'lbs'}:
                w_val = (val.attrib.get('value') or '').strip()
                w_unit = (val.attrib.get('unit') or '').strip()
                if w_val:
                    break
    weight = ""
    if clean_value(w_val):
        weight = clean_value(w_val)
        if clean_value(w_unit):
            weight = f"{weight} {w_unit}"

    # Height
    h_el = find_first(root, './/hl7:code[@displayName="height"]/../hl7:value')
    h_val = h_el.attrib.get('value', '') if h_el is not None else ''
    h_unit = h_el.attrib.get('unit', '') if h_el is not None else ''
    if not (h_val or h_unit):
        for obs in root.findall('.//hl7:observation', NS):
            val = obs.find('hl7:value', NS)
            if val is None:
                continue
            u = (val.attrib.get('unit') or '').strip().lower()
            if u in {'cm', 'm', 'in'}:
                h_val = (val.attrib.get('value') or '').strip()
                h_unit = (val.attrib.get('unit') or '').strip()
                if h_val:
                    break
    height = ""
    if clean_value(h_val):
        height = clean_value(h_val)
        if clean_value(h_unit):
            height = f"{height} {h_unit}"

    # Initials (mask-aware)
    initials = ""
    nm = find_first(root, './/hl7:player1/hl7:name')
    if nm is not None:
        if nm.attrib.get('nullFlavor') == 'MSK':
            initials = "Masked"
        else:
            parts: List[str] = []
            for g in nm.findall('hl7:given', NS):
                if g.text and g.text.strip():
                    parts.append(g.text.strip()[0].upper())
            fam = nm.find('hl7:family', NS)
            if fam is not None and fam.text and fam.text.strip():
                parts.append(fam.text.strip()[0].upper())
            initials = "".join(parts) or clean_value(get_text(nm))

    # DOB and DOD
    dob_raw = ""
    dob_el = find_first(root, './/hl7:birthTime')
    if dob_el is not None:
        dob_raw = (dob_el.attrib.get('value') or '').strip()
    dob = format_date(dob_raw)

    dod_raw = ""
    dod_el = find_first(root, './/hl7:deceasedTime')
    if dod_el is not None:
        dod_raw = (dod_el.attrib.get('value') or '').strip()
    dod = format_date(dod_raw)

    # Patient Record Number (mask-aware, by OID)
    patient_record_no = find_mask_aware_id_by_root(root, PATIENT_RECORD_OID)

    return {
        "Gender": clean_value(gender),
        "Age": clean_value(age),
        "Age Group": clean_value(age_group),
        "Height": clean_value(height),
        "Weight": clean_value(weight),
        "Initials": clean_value(initials),
        "Patient Record Number": clean_value(patient_record_no),
        "DOB": clean_value(dob),
        "DOD": clean_value(dod),
    }

# ---------------- Helper: parent map ----------------
def build_parent_map(root: ET.Element) -> Dict[ET.Element, ET.Element]:
    return {c: p for p in root.iter() for c in list(p)}

# ---------------- Reaction map: RID -> LLT term ----------------
def build_reaction_id_to_term(root: ET.Element, meddra_map: Optional[Dict[str, Dict[str, str]]] = None) -> Dict[str, str]:
    out: Dict[str, str] = {}
    for obs in findall(root, './/hl7:observation'):
        code_el = obs.find('hl7:code', NS)
        if code_el is None or (code_el.attrib.get('displayName') or '').strip().lower() != 'reaction':
            continue
        id_el = find_first(obs, './/hl7:id')
        rid_root = (id_el.attrib.get('root') or '').strip() if id_el is not None else ''
        rid_ext = (id_el.attrib.get('extension') or '').strip() if id_el is not None else ''
        llt_term = ""
        val_el = obs.find('hl7:value', NS)
        llt_code = (val_el.attrib.get('code') or '').strip() if val_el is not None else ''
        if meddra_map and llt_code in meddra_map:
            llt_term = (meddra_map[llt_code].get("LLT Term") or '').strip()
        if not llt_term and val_el is not None:
            llt_term = (val_el.attrib.get('displayName') or '').strip()
        if not llt_term and val_el is not None:
            ot = val_el.find('hl7:originalText', NS)
            if ot is not None and (ot.text or '').strip():
                llt_term = ot.text.strip()
        if llt_term:
            if rid_root:
                out[rid_root] = llt_term
            if rid_ext:
                out[rid_ext] = llt_term
    return out

# ---------------- Medical History extraction ----------------
def extract_medical_history(root: ET.Element, meddra_map: Optional[Dict[str, Dict[str, str]]] = None) -> List[Dict[str, Any]]:
    items: List[Dict[str, Any]] = []
    pmap = build_parent_map(root)
    anchors: List[ET.Element] = []
    for code in root.findall('.//hl7:code', NS):
        if (code.attrib.get('codeSystem') or '').strip() != MH_SECTION_OID:
            continue
        c = (code.attrib.get('code') or '').strip()
        disp = (code.attrib.get('displayName') or '').strip().lower()
        if c == "1" or disp == "relevantmedicalhistoryandconcurrentconditions":
            anchors.append(code)

    for anc in anchors:
        container = pmap.get(anc, None) or root
        for obs in container.findall('.//hl7:observation', NS):
            code = obs.find('hl7:code', NS)
            if code is None:
                continue
            if (code.attrib.get('codeSystem') or '').strip() != MEDDRA_LLT_OID:
                continue
            llt_code = (code.attrib.get('code') or '').strip()
            llt_term, pt_code, pt_term = "", "", ""
            if meddra_map and llt_code in meddra_map:
                m = meddra_map[llt_code]
                llt_term = (m.get("LLT Term") or '').strip()
                pt_code = (m.get("PT Code") or '').strip()
                pt_term = (m.get("PT Term") or '').strip()
            if not llt_term:
                llt_term = (code.attrib.get('displayName') or '').strip()
            if not llt_term:
                ot = code.find('hl7:originalText', NS)
                if ot is not None and (ot.text or '').strip():
                    llt_term = ot.text.strip()

            # status flags (BL true) under STATUS_OID
            statuses: List[str] = []
            for inb in obs.findall('.//hl7:inboundRelationship/hl7:observation', NS):
                scode = inb.find('hl7:code', NS)
                val = inb.find('hl7:value', NS)
                if scode is not None and (scode.attrib.get('codeSystem') or '').strip() == STATUS_OID:
                    if val is not None and ((val.attrib.get('value') or '').strip().lower() == 'true'):
                        lbl = (scode.attrib.get('displayName') or scode.attrib.get('code') or '').strip()
                        if lbl and lbl not in statuses:
                            statuses.append(lbl)

            # Continue + Comment (STATUS_OID)
            mh_continue = ""
            mh_comment = ""
            for inb2 in obs.findall('.//hl7:inboundRelationship/hl7:observation', NS):
                sc2 = inb2.find('hl7:code', NS)
                val2 = inb2.find('hl7:value', NS)
                if sc2 is None:
                    continue
                cs2 = (sc2.attrib.get('codeSystem') or '').strip()
                cd2 = (sc2.attrib.get('code') or '').strip()
                dn2 = (sc2.attrib.get('displayName') or '').strip().lower()
                if cs2 == STATUS_OID and (cd2 == '10' or dn2 == 'comment'):
                    if val2 is not None:
                        mh_comment = (val2.text or val2.attrib.get('value') or '').strip()
                if cs2 == STATUS_OID and (cd2 == '13' or dn2 == 'continuing'):
                    if val2 is not None:
                        raw = (val2.attrib.get('value') or '').strip().lower()
                        mh_continue = 'Yes' if raw in {'true', '1', 'yes', 'y'} else 'No' if raw in {'false', '0', 'no', 'n'} else (raw or '')

            # Dates
            low = obs.find('.//hl7:effectiveTime/hl7:low', NS)
            high = obs.find('.//hl7:effectiveTime/hl7:high', NS)
            sd_raw = (low.attrib.get('value') or '').strip() if low is not None else ''
            ed_raw = (high.attrib.get('value') or '').strip() if high is not None else ''
            sd = format_date(sd_raw)
            ed = format_date(ed_raw)

            key = llt_code or normalize_text(llt_term)
            if not key:
                continue
            items.append({
                "LLT Code": clean_value(llt_code),
                "LLT Term": clean_value(llt_term),
                "Status": ", ".join(statuses) if statuses else "",
                "Status (Continue)": clean_value(mh_continue),
                "Comment": clean_value(mh_comment),
                "Start Date (raw)": sd_raw,
                "Start Date": sd,
                "End Date (raw)": ed_raw,
                "End Date": ed,
                "_key": key,
            })
    return items

# ---------------- Lab Details extraction ----------------
def extract_labs(root: ET.Element, meddra_map: Optional[Dict[str, Dict[str, str]]] = None) -> List[Dict[str, Any]]:
    items: List[Dict[str, Any]] = []
    pmap = build_parent_map(root)
    anchors: List[ET.Element] = []
    for code in root.findall('.//hl7:code', NS):
        if (code.attrib.get('codeSystem') or '').strip() != MH_SECTION_OID:
            continue
        c = (code.attrib.get('code') or '').strip()
        disp = (code.attrib.get('displayName') or '').strip().lower()
        if c == "3" or disp == "testsandproceduresrelevanttotheinvestigation":
            anchors.append(code)

    for anc in anchors:
        container = pmap.get(anc, None) or root
        for obs in container.findall('.//hl7:observation', NS):
            code = obs.find('hl7:code', NS)
            if code is None or (code.attrib.get('codeSystem') or '').strip() != MEDDRA_LLT_OID:
                continue
            llt_code = (code.attrib.get('code') or '').strip()
            llt_term = ""
            if meddra_map and llt_code in meddra_map:
                m = meddra_map[llt_code]
                llt_term = (m.get("LLT Term") or '').strip()
            if not llt_term:
                llt_term = (code.attrib.get('displayName') or '').strip()
            if not llt_term:
                ot = code.find('hl7:originalText', NS)
                if ot is not None and (ot.text or '').strip():
                    llt_term = ot.text.strip()

            # Result
            value_node = obs.find('hl7:value', NS)
            result = read_numeric_with_unit(value_node)

            # Result Date
            date_val = ""
            eff = obs.find('hl7:effectiveTime', NS)
            if eff is not None:
                v = (eff.attrib.get('value') or '').strip()
                if v:
                    date_val = format_date(v)
                else:
                    low = eff.find('hl7:low', NS)
                    high = eff.find('hl7:high', NS)
                    lv = (low.attrib.get('value') or '').strip() if low is not None else ''
                    hv = (high.attrib.get('value') or '').strip() if high is not None else ''
                    date_val = format_date(lv or hv)

            key = llt_code or normalize_text(llt_term)
            if not key:
                continue
            items.append({
                "LLT Code": clean_value(llt_code),
                "LLT Term": clean_value(llt_term),
                "Result": clean_value(result),
                "Result Date": clean_value(date_val),
                "_key": key,
            })
    return items

# ---------------- Causality extraction (relaxed + improved assessor) ----------------
def _iter_components_in_doc_order(root: ET.Element) -> List[ET.Element]:
    return findall(root, './/hl7:component[@typeCode="COMP"]')


def _resolve_intervention_label(val_node: Optional[ET.Element]) -> str:
    if val_node is None:
        return ""
    dsn = (val_node.attrib.get('displayName') or '').strip()
    if dsn:
        return dsn
    code = (val_node.attrib.get('code') or '').strip()
    if code:
        return f"code:{code}"
    ot = val_node.find('hl7:originalText', NS)
    if ot is not None and (ot.text or '').strip():
        return ot.text.strip()
    return get_text(val_node)


def _extract_assessor_label(node: ET.Element) -> str:
    cand_texts: List[str] = []
    for xp in [
        './/hl7:author//hl7:assignedEntity//hl7:code/hl7:originalText',
        './/hl7:author//hl7:assignedAuthor//hl7:code/hl7:originalText',
    ]:
        el = find_first(node, xp)
        if el is not None and (el.text or '').strip():
            cand_texts.append(el.text.strip())
    for xp in [
        './/hl7:author//hl7:assignedEntity//hl7:code',
        './/hl7:author//hl7:assignedAuthor//hl7:code',
    ]:
        el = find_first(node, xp)
        if el is not None:
            for attr in ('displayName', 'code'):
                v = (el.attrib.get(attr) or '').strip()
                if v:
                    cand_texts.append(v)
    for t in cand_texts:
        low = t.lower()
        if 'company' in low:
            return "Company"
        if 'reporter' in low:
            return "Reporter"
    nm = find_first(node, './/hl7:author//hl7:assignedEntity//hl7:name')
    if nm is not None and get_text(nm):
        return get_text(nm)
    ot = find_first(node, './/hl7:author//hl7:assignedEntity//hl7:originalText')
    if ot is not None and get_text(ot):
        return get_text(ot)
    return cand_texts[0] if cand_texts else ""


def extract_causality(
    root: ET.Element,
    product_id_to_name: Optional[Dict[str, str]] = None,
    reaction_id_to_term: Optional[Dict[str, str]] = None,
) -> List[Dict[str, Any]]:
    out: List[Dict[str, Any]] = []
    try:
        processed_nodes: Set[int] = set()
        seen_signatures: Set[Tuple[str, str, str, str, str, str]] = set()

        for comp in _iter_components_in_doc_order(root):
            current_intervention = ""  # resets per component
            ca_nodes = comp.findall('.//hl7:causalityAssessment', NS)
            for node in ca_nodes:
                if id(node) in processed_nodes:
                    continue
                processed_nodes.add(id(node))

                ccode = node.find('hl7:code', NS)
                if ccode is None:
                    continue
                cs = (ccode.attrib.get('codeSystem') or '').strip()
                cd = (ccode.attrib.get('code') or '').strip()
                dsn = (ccode.attrib.get('displayName') or '').strip().lower()
                if cs != STATUS_OID:
                    continue

                # Intervention sentinel
                if dsn == 'interventioncharacterization' or cd == INTERVENTION_CHAR_CODE:
                    current_intervention = _resolve_intervention_label(find_first(node, './/hl7:value'))
                    continue

                # Causality rows
                if not (cd == CAUSALITY_CODE or dsn == 'causality'):
                    continue

                # Assessment (xsi:ST text or @value)
                val = find_first(node, './/hl7:value')
                assessment = ""
                if val is not None:
                    assessment = (val.attrib.get('value') or "").strip()
                    if not assessment:
                        assessment = (val.text or "").strip()
                method = get_text(find_first(node, './/hl7:methodCode/hl7:originalText'))
                assessor = _extract_assessor_label(node)

                # IDs -> names
                evt_id = ""
                prd_id = ""
                evt = find_first(node, './/hl7:subject1//hl7:adverseEffectReference//hl7:id')
                if evt is not None:
                    evt_id = (evt.attrib.get('root') or '').strip() or (evt.attrib.get('extension') or '').strip()
                prd = find_first(node, './/hl7:subject2//hl7:productUseReference//hl7:id')
                if prd is not None:
                    prd_id = (prd.attrib.get('root') or '').strip() or (prd.attrib.get('extension') or '').strip()

                reaction_name = reaction_id_to_term.get(evt_id, "") if (reaction_id_to_term and evt_id) else ""
                drug_name = product_id_to_name.get(prd_id, "") if (product_id_to_name and prd_id) else ""
                if not drug_name:
                    sa = comp.find('.//hl7:substanceAdministration', NS)
                    if sa is not None:
                        nm_txt = _resolve_drug_name(sa).strip()
                        if nm_txt:
                            drug_name = nm_txt

                # Pairing key (internal only)
                key = (evt_id or "") + "::" + (prd_id or "")
                if not key.strip(":"):
                    key = "assess::" + normalize_text(assessment or "") + "::" + normalize_text(method or "")

                # Content-level de-duplication
                row_sig = (
                    (evt_id or "").strip().lower(),
                    (prd_id or "").strip().lower(),
                    normalize_text(reaction_name),
                    normalize_text(drug_name),
                    normalize_text(assessor),
                    normalize_text((method or "") + "\n" + (assessment or "")),
                )
                if row_sig in seen_signatures:
                    continue
                seen_signatures.add(row_sig)

                out.append({
                    "Assessment": clean_value(assessment),
                    "Method": clean_value(method),
                    "Assessor": clean_value(assessor),
                    "Reaction": clean_value(reaction_name),
                    "Drug": clean_value(drug_name),
                    "_key": key,
                    "_evt_id": evt_id,
                    "_prd_id": prd_id,
                })
    except Exception as e:
        st.warning(f"Causality parse error: {e}")
    return out

# ---------------- Products extraction (IDs helpers and NEW categorization map) ----------------
def extract_suspect_ids(root: ET.Element) -> Set[str]:
    out: Set[str] = set()
    for c in findall(root, './/hl7:causalityAssessment'):
        sid = find_first(c, './/hl7:subject2/hl7:productUseReference/hl7:id')
        if sid is not None:
            rid = sid.attrib.get('root', '')
            if rid:
                out.add(rid)
    return out


def extract_interacting_ids(root: ET.Element) -> Set[str]:
    ids: Set[str] = set()
    for obs in findall(root, './/hl7:observation'):
        code = obs.find('hl7:code', NS)
        disp = (code.attrib.get('displayName') or '').strip().lower() if code is not None else ''
        if 'interact' in disp:
            ref = obs.find('.//hl7:subject2/hl7:productUseReference/hl7:id', NS)
            if ref is not None:
                ids.add(ref.attrib.get('root', ''))
    return ids


def extract_treatment_ids(root: ET.Element) -> Set[str]:
    ids: Set[str] = set()
    for obs in findall(root, './/hl7:observation'):
        code = obs.find('hl7:code', NS)
        disp = (code.attrib.get('displayName') or '').strip().lower() if code is not None else ''
        if ('treat' in disp) or ('therapeutic' in disp):
            ref = obs.find('.//hl7:subject2/hl7:productUseReference/hl7:id', NS)
            if ref is not None:
                ids.add(ref.attrib.get('root', ''))
    return ids

# ---- Product categorization via causalityAssessment/interventionCharacterization ----
PRODUCT_TYPE_MAP: Dict[str, str] = {
    "1": "Suspect",
    "2": "Concomitant",
    "3": "Interacting",
    "4": "Drug Not Administered",
}

def build_product_type_by_pid(root: ET.Element) -> Dict[str, str]:
    """
    Build a map of product-id-root -> product type label based on:
    <causalityAssessment>
      <code code="20" codeSystem=STATUS_OID ... displayName="interventionCharacterization"/>
      <value code="1|2|3|4" .../>
      <subject2><productUseReference><id root="..."/></subject2>
    """
    result: Dict[str, str] = {}
    for node in findall(root, './/hl7:causalityAssessment'):
        ccode = node.find('hl7:code', NS)
        if ccode is None:
            continue
        cs = (ccode.attrib.get('codeSystem') or '').strip()
        cd = (ccode.attrib.get('code') or '').strip()
        dsn = (ccode.attrib.get('displayName') or '').strip().lower()
        if cs != STATUS_OID:
            continue
        if not (cd == INTERVENTION_CHAR_CODE or dsn == 'interventioncharacterization'):
            continue
        val = find_first(node, './/hl7:value')
        vcode = (val.attrib.get('code') or '').strip() if val is not None else ''
        if not vcode:
            continue
        pid_el = find_first(node, './/hl7:subject2//hl7:productUseReference//hl7:id')
        pid = (pid_el.attrib.get('root') or '').strip() if pid_el is not None else ''
        if not pid:
            continue
        label = PRODUCT_TYPE_MAP.get(vcode, '')
        if label:
            result[pid] = label
    return result


def _resolve_drug_name(admin: ET.Element) -> str:
    nm = find_first(admin, './/hl7:kindOfProduct/hl7:name')
    if nm is not None:
        t = (nm.text or '').strip()
        if t:
            return t
        disp = (nm.attrib.get('displayName') or '').strip()
        if disp:
            return disp
        ot = nm.find('hl7:originalText', NS)
        if ot is not None and (ot.text or '').strip():
            return ot.text.strip()
    alt = find_first(admin, './/hl7:manufacturedProduct/hl7:name')
    if alt is not None and (alt.text or '').strip():
        return alt.text.strip()
    mm = find_first(admin, './/hl7:manufacturedMaterial/hl7:name')
    if mm is not None and (mm.text or '').strip():
        return mm.text.strip()
    mm_code = find_first(admin, './/hl7:manufacturedMaterial/hl7:code')
    if mm_code is not None:
        disp = (mm_code.attrib.get('displayName') or '').strip()
        if disp:
            return disp
    amp = find_first(admin, './/hl7:asManufacturedProduct//hl7:name')
    if amp is not None and (amp.text or '').strip():
        return amp.text.strip()
    return ""


def _resolve_ingredient_names(admin: ET.Element) -> str:
    """Return active ingredient substance name(s) for a drug product.

    E2B(R3) drug XML can contain both:
      • kindOfProduct/name = medicinal product name, usually including strength
      • ingredient/ingredientSubstance/name = active ingredient/substance name

    Keep these separate so the Drug section can display both values.
    """
    ingredients: List[str] = []

    # Preferred E2B(R3) path for active ingredient substance names
    for nm in admin.findall('.//hl7:ingredient[@classCode="ACTI"]/hl7:ingredientSubstance/hl7:name', NS):
        _add_unique(ingredients, get_text(nm))

    # Fallback: capture any ingredient substance name even if classCode is missing/varies
    if not ingredients:
        for nm in admin.findall('.//hl7:ingredientSubstance/hl7:name', NS):
            _add_unique(ingredients, get_text(nm))

    return "\n".join(ingredients)

# ---------------- Drug History extraction (RELAXED + robust dates) ----------------
def extract_drug_history(root: ET.Element, meddra_map: Optional[Dict[str, Dict[str, str]]] = None) -> List[Dict[str, Any]]:
    """
    Collect prior/concomitant drug history items from the clinical section that is
    tagged as 'drugHistory'. Anchors on either:
      • codeSystem == MH_SECTION_OID AND code == '2'
      • displayName.lower() == 'drughistory'   (regardless of codeSystem)
    Uses robust date extraction (supports SXPR/IVL and mask).
    """
    items: List[Dict[str, Any]] = []
    pmap = build_parent_map(root)

    # ---- Find anchors by either codeSystem/code or by displayName='drugHistory'
    anchors: List[ET.Element] = []
    for code in root.findall('.//hl7:code', NS):
        disp = (code.attrib.get('displayName') or '').strip().lower()
        cs   = (code.attrib.get('codeSystem')  or '').strip()
        cd   = (code.attrib.get('code')        or '').strip()
        is_expected_section = (cs == MH_SECTION_OID and cd == '2')
        is_display_drughistory = (disp == 'drughistory')
        if is_expected_section or is_display_drughistory:
            anchors.append(code)

    if not anchors:
        return items

    def _map_llt(llt_code: str, fallback: str = "") -> str:
        if meddra_map and llt_code in meddra_map:
            term = (meddra_map[llt_code].get('LLT Term') or '').strip()
            if term:
                return term
        return fallback or llt_code

    for anc in anchors:
        container = pmap.get(anc, None) or root

        for sa in container.findall('.//hl7:substanceAdministration[@moodCode="EVN"][@classCode="SBADM"]', NS):
            drug = clean_value(_resolve_drug_name(sa))

            # Dates (robust)
            sd, ed = _extract_sa_dates(sa)

            # Indications / Reactions attached to this SA via STATUS_OID
            indications: List[str] = []
            reactions:  List[str] = []

            for ob in sa.findall('.//hl7:outboundRelationship2/hl7:observation', NS):
                c  = ob.find('hl7:code', NS)
                if c is None:
                    continue
                cs = (c.attrib.get('codeSystem') or '').strip()
                cd = (c.attrib.get('code') or '').strip()
                dn = (c.attrib.get('displayName') or '').strip().lower()
                if cs != STATUS_OID:
                    continue

                v = ob.find('hl7:value', NS)
                llt_code = (v.attrib.get('code') or '').strip() if v is not None else ''
                llt_disp = (v.attrib.get('displayName') or '').strip() if (v is not None and (v.attrib.get('displayName') or '').strip()) else ''

                if dn == 'indication' or cd == '19':
                    term = _map_llt(llt_code, llt_disp)
                    if has_value(term) and term not in indications:
                        indications.append(term)
                if dn == 'reaction' or cd == '29':
                    term = _map_llt(llt_code, llt_disp)
                    if has_value(term) and term not in reactions:
                        reactions.append(term)

            key = (normalize_text(drug) or f"sa::{sd}::{ed}")
            if not key:
                continue

            items.append({
                'Drug': drug,
                'Indication': "\n".join(indications),
                'Reaction':  "\n".join(reactions),
                'Start Date': sd,
                'End Date': ed,
                '_key': key,
            })

    return items

# ---------------- Death Details extraction (by displayName) ----------------
def extract_death_details(root: ET.Element, meddra_map: Optional[Dict[str, Dict[str, str]]] = None) -> List[Dict[str, str]]:
    """
    Fetch all death-related details by displayName and preserve their order.
      • reportedCauseOfDeath : value may be CE with MedDRA LLT and/or originalText
      • autopsy              : BL true/false -> 'Yes'/'No'

    Returns a list so multiple death-detail entries can be displayed in the UI.
    If counts differ, the nth cause/autopsy values are paired by position.
    """
    causes: List[str] = []
    autopsies: List[str] = []

    # reportedCauseOfDeath (collect every occurrence)
    for obs in root.findall('.//hl7:observation', NS):
        c = obs.find('hl7:code', NS)
        if c is None or (c.attrib.get('displayName') or '').strip() != 'reportedCauseOfDeath':
            continue
        val = obs.find('hl7:value', NS)
        cause_term = ""
        if val is not None:
            llt_code = (val.attrib.get('code') or '').strip()
            if meddra_map and llt_code and llt_code in meddra_map:
                cause_term = (meddra_map[llt_code].get('LLT Term') or '').strip()
            if not cause_term:
                cause_term = (val.attrib.get('displayName') or '').strip()
            if not cause_term:
                ot = val.find('hl7:originalText', NS)
                if ot is not None and (ot.text or '').strip():
                    cause_term = ot.text.strip()
            if not cause_term and llt_code:
                cause_term = llt_code
        cause_term = clean_value(cause_term)
        if cause_term:
            causes.append(cause_term)

    # autopsy (collect every occurrence)
    for obs in root.findall('.//hl7:observation', NS):
        c = obs.find('hl7:code', NS)
        if c is None or (c.attrib.get('displayName') or '').strip() != 'autopsy':
            continue
        val = obs.find('hl7:value', NS)
        if val is None:
            continue
        raw = (val.attrib.get('value') or val.text or '').strip()
        raw_low = raw.lower()
        if raw_low in {'true', '1', 'yes', 'y'}:
            autopsy_value = 'Yes'
        elif raw_low in {'false', '0', 'no', 'n'}:
            autopsy_value = 'No'
        else:
            autopsy_value = raw
        autopsy_value = clean_value(autopsy_value)
        if autopsy_value:
            autopsies.append(autopsy_value)

    count = max(len(causes), len(autopsies))
    out: List[Dict[str, str]] = []
    for i in range(count):
        rec = {
            'Reported Cause of Death': causes[i] if i < len(causes) else '',
            'Autopsy': autopsies[i] if i < len(autopsies) else '',
            '_key': f'death_{i + 1}',
        }
        if rec['Reported Cause of Death'] or rec['Autopsy']:
            out.append(rec)
    return out

# ---------------- Iter helpers and product aggregation ----------------
def _iter_drug_components(root: ET.Element) -> List[ET.Element]:
    comps: List[ET.Element] = []
    for comp in root.findall('.//hl7:component[@typeCode="COMP"]', NS):
        sas = comp.findall('.//hl7:substanceAdministration', NS)
        if sas:
            comps.append(comp)
    return comps

def _add_unique(acc_list: List[str], value: str):
    v = clean_value(value)
    if not v:
        return
    if v not in acc_list:
        acc_list.append(v)

# ✅ Robust SA date extractor (handles SXPR/IVL and mask)
def _extract_sa_dates(sa: ET.Element) -> Tuple[str, str]:
    """
    Extract (Start, Stop) for a substanceAdministration with robust handling:
      • <effectiveTime value="YYYYMMDD">
      • <effectiveTime><low/><high/></effectiveTime>
      • <effectiveTime>//comp//low|high (SXPR_TS / IVL_TS / PIVL_TS)
      • nullFlavor="MSK" -> "Masked"
    Returns display strings already formatted via format_date / "Masked".
    """
    et_nodes = sa.findall('.//hl7:effectiveTime', NS)

    def _pick_first(nodes: List[ET.Element], tag: str) -> Tuple[str, bool]:
        # direct child
        for node in nodes:
            child = node.find(f'hl7:{tag}', NS)
            if child is not None:
                nf = (child.attrib.get('nullFlavor') or '').strip().upper()
                if nf == 'MSK':
                    return ("Masked", True)
                val = (child.attrib.get('value') or '').strip()
                if val:
                    return (val, False)
        # any descendant (comp/IVL_TS/PIVL_TS)
        for node in nodes:
            for desc in node.findall(f'.//hl7:{tag}', NS):
                nf = (desc.attrib.get('nullFlavor') or '').strip().upper()
                if nf == 'MSK':
                    return ("Masked", True)
                val = (desc.attrib.get('value') or '').strip()
                if val:
                    return (val, False)
        return ("", False)

    low_val, low_mask = _pick_first(et_nodes, 'low')
    high_val, high_mask = _pick_first(et_nodes, 'high')

    # Instant @value as fallback
    if not low_val:
        for et in et_nodes:
            v = (et.attrib.get('value') or '').strip()
            if v:
                low_val = v
                break

    start_disp = "Masked" if low_mask else format_date(low_val)
    end_disp   = "Masked" if high_mask else format_date(high_val)
    return (start_disp or "", end_disp or "")

def extract_all_products(root: ET.Element, meddra_map: Optional[Dict[str, Dict[str, str]]] = None) -> List[Dict[str, Any]]:
    """
    Anchor-windowed extraction. Create exactly ONE product row for each direct-child
    <substanceAdministration SBADM/EVN> that has a non-empty <id@root>.

    For a given anchor SA, aggregate content from the current SA position up to
    (but not including) the next anchor SA within the SAME <component>.

    Now includes nested SAs and robust date parsing (SXPR/IVL).
    """
    suspects = extract_suspect_ids(root)
    interact = extract_interacting_ids(root)
    treatments = extract_treatment_ids(root)

    product_type_by_pid = build_product_type_by_pid(root)

    out: List[Dict[str, Any]] = []
    comps = _iter_drug_components(root)
    for cidx, comp in enumerate(comps, start=1):
        comp_children = list(comp)

        # Identify direct-child SAs at this component level
        sa_positions: List[Tuple[int, ET.Element]] = []
        for i, child in enumerate(comp_children):
            if local_name(child.tag) == 'substanceAdministration' and \
               child.attrib.get('moodCode') == 'EVN' and child.attrib.get('classCode') == 'SBADM':
                sa_positions.append((i, child))

        # Build list of ANCHORS = SA that have an id@root
        anchors: List[Tuple[int, ET.Element, str]] = []
        for pos, sa in sa_positions:
            id_el = find_first(sa, './/hl7:id')
            pid = (id_el.attrib.get('root') or '').strip() if id_el is not None else ''
            if pid:
                anchors.append((pos, sa, pid))

        if not anchors:
            continue
        anchors.sort(key=lambda t: t[0])

        for a_idx, (pos, sa_anchor, pid) in enumerate(anchors, start=1):
            start_pos = pos
            end_pos = anchors[a_idx][0] if a_idx < len(anchors) else len(comp_children)
            window_nodes = comp_children[start_pos:end_pos]

            def win_findall(xpath: str) -> List[ET.Element]:
                acc: List[ET.Element] = []
                for wn in window_nodes:
                    acc.extend(wn.findall(xpath, NS))
                return acc

            # Product name and active ingredient from ANCHOR SA only
            title = clean_value(_resolve_drug_name(sa_anchor)) or pid
            ingredient_txt = clean_value(_resolve_ingredient_names(sa_anchor))

            # Type from categorization map; else heuristic
            type_disp = product_type_by_pid.get(pid, "")
            if not type_disp:
                tags: Set[str] = set()
                if pid in suspects:
                    tags.add('Suspect')
                if pid in interact:
                    tags.add('Interacting')
                if pid in treatments:
                    tags.add('Treatment')
                if not tags:
                    tags.add('Concomitant')
                type_disp = ', '.join(sorted(tags))

            # Aggregate fields across ALL SAs within the window (including nested)
            dosage_texts: List[str] = []
            dose_vals: List[str] = []
            dose_units: List[str] = []
            start_dates: List[str] = []
            stop_dates: List[str] = []
            routes: List[str] = []
            forms: List[str] = []
            lots: List[str] = []
            mahs: List[str] = []

            for wn in window_nodes:
                # include self if SA, plus any nested SAs
                if local_name(wn.tag) == 'substanceAdministration':
                    sas = [wn] + wn.findall('.//hl7:substanceAdministration[@moodCode="EVN"][@classCode="SBADM"]', NS)
                else:
                    sas = wn.findall('.//hl7:substanceAdministration[@moodCode="EVN"][@classCode="SBADM"]', NS)

                for sa in sas:
                    # dosage text
                    txt = get_text(find_first(sa, './/hl7:text'))
                    _add_unique(dosage_texts, txt)

                    # dose
                    dq = find_first(sa, './/hl7:doseQuantity')
                    if dq is not None:
                        _add_unique(dose_vals, (dq.attrib.get('value') or '').strip())
                        _add_unique(dose_units, (dq.attrib.get('unit') or '').strip())

                    # dates (robust)
                    sd, ed = _extract_sa_dates(sa)
                    _add_unique(start_dates, sd)
                    _add_unique(stop_dates, ed)

                    # route
                    rtxt = get_text(find_first(sa, './/hl7:routeCode/hl7:originalText'))
                    if not rtxt:
                        rc = find_first(sa, './/hl7:routeCode')
                        rtxt = (rc.attrib.get('displayName') or '').strip() if rc is not None else ''
                    _add_unique(routes, rtxt)

                    # form
                    form = get_text(find_first(sa, './/hl7:formCode/hl7:originalText'))
                    _add_unique(forms, form)

                    # lot
                    lot = get_text(find_first(sa, './/hl7:lotNumberText'))
                    _add_unique(lots, lot)

                    # MAH
                    mah = ''
                    for xp in [
                        './/hl7:playingOrganization/hl7:name',
                        './/hl7:manufacturerOrganization/hl7:name',
                        './/hl7:asManufacturedProduct/hl7:manufacturerOrganization/hl7:name',
                    ]:
                        node = find_first(sa, xp)
                        if node is not None and get_text(node):
                            mah = get_text(node)
                            break
                    _add_unique(mahs, mah)

            # Window-scoped observations (Action Taken, Obtain Country, Indication)
            action_taken_vals: List[str] = []
            for act_code in win_findall('.//hl7:act[@classCode="ACT"][@moodCode="EVN"]/hl7:code'):
                if (act_code.attrib.get('codeSystem') or '').strip() == ACTION_TAKEN_OID:
                    c = (act_code.attrib.get('code') or '').strip()
                    label = ACTION_TAKEN_MAP.get(c, c or '')
                    _add_unique(action_taken_vals, label)
            action_taken = '\n'.join([v for v in action_taken_vals if has_value(v)])

            obtain_countries: List[str] = []
            for cn in win_findall('.//hl7:country'):
                val = (cn.text or '').strip()
                _add_unique(obtain_countries, val)
            obtain_country = '\n'.join([v for v in obtain_countries if has_value(v)])

            indications: List[str] = []
            for ind_obs in win_findall('.//hl7:observation'):
                code_el = ind_obs.find('hl7:code', NS)
                if code_el is None:
                    continue
                cs = (code_el.attrib.get('codeSystem') or '').strip()
                cd = (code_el.attrib.get('code') or '').strip()
                dn = (code_el.attrib.get('displayName') or '').strip().lower()
                if cs == STATUS_OID and (dn == 'indication' or cd == '19'):
                    val = ind_obs.find('hl7:value', NS)
                    if val is not None:
                        llt_code = (val.attrib.get('code') or '').strip()
                        rrt = ''
                        ot = val.find('hl7:originalText', NS)
                        if ot is not None and (ot.text or '').strip():
                            rrt = ot.text.strip()
                        llt_display = ''
                        if meddra_map and llt_code in meddra_map:
                            llt_display = (meddra_map[llt_code].get('LLT Term') or '').strip()
                        if not llt_display:
                            llt_display = llt_code
                        frag = f'Indication: RRT: {rrt}; LLT: {llt_display}'.strip()
                        _add_unique(indications, frag)
            indication_txt = '\n'.join([v for v in indications if has_value(v)])

            def join_vals(lst: List[str]) -> str:
                return "\n".join([v for v in lst if has_value(v)])

            out.append({
                'Drug': title,
                'Drug Name': title,
                'Ingredient': ingredient_txt,
                'Type': type_disp,
                'Dosage Text': join_vals(dosage_texts),
                'Dose Value': join_vals(dose_vals),
                'Dose Unit': join_vals(dose_units),
                'Start Date': join_vals(start_dates),
                'Stop Date': join_vals(stop_dates),
                'Route': join_vals(routes),
                'Formulation': join_vals(forms),
                'Lot No': join_vals(lots),
                'MAH': join_vals(mahs),
                'Action Taken': action_taken,
                'Drug Obtain Country': obtain_country,
                'Indication': indication_txt,
                '_gid': f"pid::{pid.lower()}",
                '_pid': pid,
            })
    return out

# ---------------- Reporter extraction (strict: sourceReport branches) ----------------
def find_all_source_report_containers(root: ET.Element) -> List[ET.Element]:
    code_nodes: List[ET.Element] = []
    for el in root.iter():
        if local_name(el.tag) == 'code' and el.attrib.get('codeSystem') == REPORT_SOURCE_OID:
            if (el.attrib.get('displayName') or '').strip().lower() == 'sourcereport':
                code_nodes.append(el)
    if not code_nodes:
        return []

    parent = build_parent_map(root)

    def ancestors(node: ET.Element) -> List[ET.Element]:
        acc: List[ET.Element] = []
        cur = node
        while cur in parent:
            cur = parent[cur]
            acc.append(cur)
        return acc

    containers: List[ET.Element] = []
    for code_el in code_nodes:
        for anc in ancestors(code_el):
            lname = local_name(anc.tag)
            if lname in {'relatedInvestigation', 'subjectOf2', 'controlActEvent'}:
                for xp in [
                    './/hl7:author/hl7:assignedEntity',
                    './/hl7:author/hl7:assignedAuthor',
                    './/hl7:informant/hl7:assignedEntity',
                ]:
                    cand = anc.find(xp, NS)
                    if cand is not None:
                        containers.append(cand)
                        break
                break

    seen: Set[int] = set()
    uniq: List[ET.Element] = []
    for el in containers:
        if id(el) not in seen:
            seen.add(id(el))
            uniq.append(el)
    return uniq


def extract_reporter_from_container(node: ET.Element) -> Dict[str, str]:
    result = {
        "Reporter Qualification": "",
        "Reporter IDs": "",
        "Reporter Title": "",
        "Reporter Given Name(s)": "",
        "Reporter Family Name": "",
        "Reporter Organization": "",
        "Reporter Street": "",
        "Reporter City/Town": "",
        "Reporter State/Province": "",
        "Reporter Postal Code": "",
        "Reporter Country": "",
        "Reporter Phone(s)": "",
        "Reporter Email(s)": "",
        "Reporter Fax(es)": "",
    }

    # IDs
    ids: List[str] = []
    for id_el in node.findall('.//hl7:id', NS):
        ext = (id_el.attrib.get('extension') or '').strip()
        rt = (id_el.attrib.get('root') or '').strip()
        if ext and rt:
            ids.append(f"{ext} ({rt})")
        elif ext:
            ids.append(ext)
        elif rt:
            ids.append(rt)
    if ids:
        result["Reporter IDs"] = "; ".join(dict.fromkeys(ids))

    # Qualification
    qual = ""
    for code_el in node.iter():
        if local_name(code_el.tag) == 'code' and code_el.attrib.get('codeSystem') == REPORTER_QUAL_OID:
            c = (code_el.attrib.get('code') or '').strip()
            qual = REPORTER_MAP.get(c, c)
            break
    result["Reporter Qualification"] = qual

    # Name parts (mask-aware)
    name_el = node.find('.//hl7:assignedPerson/hl7:name', NS) or node.find('.//hl7:name', NS)
    title_vals: List[str] = []
    given_vals: List[str] = []
    family_val = ""
    if name_el is not None:
        for pfx in name_el.findall('hl7:prefix', NS):
            v = read_text_or_mask(pfx)
            if v:
                title_vals.append(v)
        for g in name_el.findall('hl7:given', NS):
            v = read_text_or_mask(g)
            if v:
                given_vals.append(v)
        fam_el = name_el.find('hl7:family', NS)
        family_val = read_text_or_mask(fam_el)
    result["Reporter Title"] = "; ".join(title_vals) if title_vals else ""
    result["Reporter Given Name(s)"] = "; ".join(given_vals) if given_vals else ""
    result["Reporter Family Name"] = family_val

    # Organization
    for xp in [
        './/hl7:assignedEntity/hl7:representedOrganization/hl7:name',
        './/hl7:representedOrganization/hl7:name',
        './/hl7:scopingOrganization/hl7:name',
    ]:
        el = node.find(xp, NS)
        if el is not None:
            txt = read_text_or_mask(el)
            if txt:
                result["Reporter Organization"] = txt
                break

    # Address
    addr = node.find('.//hl7:addr', NS)
    streets: List[str] = []
    city = state = postal = country = ""
    if addr is not None:
        for sl in addr.findall('hl7:streetAddressLine', NS):
            val = read_text_or_mask(sl)
            if val:
                streets.append(val)
        city = read_text_or_mask(addr.find('hl7:city', NS))
        state = read_text_or_mask(addr.find('hl7:state', NS))
        postal = read_text_or_mask(addr.find('hl7:postalCode', NS))
        country = read_text_or_mask(addr.find('hl7:country', NS))
        if not country:
            loc = node.find('.//hl7:asLocatedEntity/hl7:location/hl7:code', NS)
            if loc is not None and loc.attrib.get('code'):
                country = loc.attrib.get('code').strip()
    result["Reporter Street"] = ", ".join(streets)
    result["Reporter City/Town"] = city
    result["Reporter State/Province"] = state
    result["Reporter Postal Code"] = postal
    result["Reporter Country"] = country

    # Telecoms
    phones: List[str] = []
    emails: List[str] = []
    faxes: List[str] = []
    for tel in node.findall('.//hl7:telecom', NS):
        raw = (tel.attrib.get('value') or '').strip()
        use = (tel.attrib.get('use') or '').upper()
        if not raw:
            continue
        low = raw.lower()
        if low.startswith('mailto:'):
            emails.append(raw.split(':', 1)[1])
        elif 'FAX' in use or low.startswith('fax:'):
            faxes.append(raw.split(':', 1)[-1] if ':' in raw else raw)
        elif low.startswith('tel:') or low.startswith('tel;'):
            phones.append(raw.split(':', 1)[1] if ':' in raw else raw)
        else:
            if '@' in raw:
                emails.append(raw.replace('mailto:', ''))
            else:
                digits = re.sub(r'\D', '', raw)
                if len(digits) >= 7:
                    phones.append(raw)
                else:
                    phones.append(raw)
    if phones:
        result["Reporter Phone(s)"] = "; ".join(dict.fromkeys(phones))
    if emails:
        result["Reporter Email(s)"] = "; ".join(dict.fromkeys(emails))
    if faxes:
        result["Reporter Fax(es)"] = "; ".join(dict.fromkeys(faxes))

    return result


def extract_reporters_from_sourceReport(root: ET.Element) -> List[Dict[str, str]]:
    containers = find_all_source_report_containers(root)
    reporters: List[Dict[str, str]] = []
    for node in containers:
        rep = extract_reporter_from_container(node)
        if any(clean_value(v) for v in rep.values()):
            reporters.append(rep)
    return reporters

# ---------------- Events extraction ----------------
def extract_events(root: ET.Element, meddra_map: Optional[Dict[str, Dict[str, str]]] = None) -> List[Dict[str, Any]]:
    seriousness_map = {
        "resultsInDeath": "Death",
        "isLifeThreatening": "LT",
        "requiresInpatientHospitalization": "Hospital",
        "resultsInPersistentOrSignificantDisability": "Disability",
        "congenitalAnomalyBirthDefect": "Congenital",
        "otherMedicallyImportantCondition": "IME",
    }
    outcome_map = {
        "1": "Recovered/Resolved",
        "2": "Recovering/Resolving",
        "3": "Not recovered/Ongoing",
        "4": "Recovered with sequelae",
        "5": "Fatal",
        "0": "Unknown",
    }

    out: List[Dict[str, Any]] = []
    try:
        rxns = findall(root, './/hl7:observation')
        for rxn in rxns:
            code_el = rxn.find('hl7:code', NS)
            if code_el is None or (code_el.attrib.get('displayName') or '').strip().lower() != 'reaction':
                continue

            # Raw LLT code (from value/@code)
            val_el = rxn.find('hl7:value', NS)
            llt_code = (val_el.attrib.get('code') or '').strip() if val_el is not None else ''

            # Decide what to display for the Event term
            event_term = ""
            pt_term = ""
            if meddra_map and llt_code in meddra_map:
                event_term = (meddra_map[llt_code].get("LLT Term") or '').strip()
                pt_term = (meddra_map[llt_code].get("PT Term") or '').strip()
            if not event_term and val_el is not None:
                event_term = (val_el.attrib.get('displayName') or '').strip() or event_term
            if not event_term and val_el is not None:
                ot = val_el.find('hl7:originalText', NS)
                if ot is not None and (ot.text or '').strip():
                    event_term = ot.text.strip()
            if not event_term:
                event_term = llt_code

            # RRT
            rrt_term = ""
            if val_el is not None:
                ot = val_el.find('hl7:originalText', NS)
                if ot is not None and (ot.text or '').strip():
                    rrt_term = ot.text.strip()

            # Seriousness
            flags: List[str] = []
            for crit, label in seriousness_map.items():
                crit_el = rxn.find(f'.//hl7:code[@displayName="{crit}"]/../hl7:value', NS)
                if crit_el is not None and (crit_el.attrib.get('value') or '').strip().lower() == 'true':
                    flags.append(label)
            seriousness_disp = "Non-serious" if not flags else ", ".join(sorted(set(flags)))

            # Outcome
            outcome_el = rxn.find('.//hl7:code[@displayName="outcome"]/../hl7:value', NS)
            outcome_code = (outcome_el.attrib.get('code') or '').strip() if outcome_el is not None else ''
            outcome = outcome_map.get(outcome_code, "Unknown" if outcome_code else "")

            # Dates
            low = rxn.find('.//hl7:effectiveTime/hl7:low', NS)
            high = rxn.find('.//hl7:effectiveTime/hl7:high', NS)
            start_raw = (low.attrib.get('value') or '').strip() if low is not None else ''
            end_raw = (high.attrib.get('value') or '').strip() if high is not None else ''
            start_disp = format_date(start_raw)
            end_disp = format_date(end_raw)

            # Country
            country = ""
            loc_code = rxn.find('.//hl7:location//hl7:locatedPlace//hl7:code', NS)
            if loc_code is not None and (loc_code.attrib.get('code') or '').strip():
                country = loc_code.attrib.get('code').strip()

            # Translation term
            translation_term = ""
            for ob in rxn.findall('.//hl7:outboundRelationship2[@typeCode="PERT"]/hl7:observation', NS):
                c = ob.find('hl7:code', NS)
                if c is None:
                    continue
                if (c.attrib.get('codeSystem') or '').strip() == STATUS_OID and (
                    (c.attrib.get('displayName') or '').strip().lower() == 'reactionfortranslation' or (c.attrib.get('code') or '').strip() == '30'
                ):
                    v = ob.find('hl7:value', NS)
                    if v is not None and (v.text or '').strip():
                        translation_term = v.text.strip()
                        break

            # Highlighted by reporter
            highlighted = ""
            for ob in rxn.findall('.//hl7:outboundRelationship2[@typeCode="PERT"]/hl7:observation', NS):
                c = ob.find('hl7:code', NS)
                if c is None:
                    continue
                if (c.attrib.get('codeSystem') or '').strip() == STATUS_OID and (
                    (c.attrib.get('displayName') or '').strip().lower() == 'termhighlightedbyreporter' or (c.attrib.get('code') or '').strip() == '37'
                ):
                    v = ob.find('hl7:value', NS)
                    code_val = (v.attrib.get('code') or '').strip() if v is not None else ''
                    if code_val == '1':
                        highlighted = "Yes"
                    elif code_val == '0':
                        highlighted = "No"
                    else:
                        highlighted = code_val
                    break

            # Medically confirmed (medicalConfirmationByHealthProfessional)
            medically_confirmed = ""
            for ob in rxn.findall('.//hl7:outboundRelationship2[@typeCode="PERT"]/hl7:observation', NS):
                c = ob.find('hl7:code', NS)
                if c is None:
                    continue
                if (c.attrib.get('codeSystem') or '').strip() == STATUS_OID and (
                    (c.attrib.get('displayName') or '').strip().lower() == 'medicalconfirmationbyhealthprofessional' or (c.attrib.get('code') or '').strip() == '24'
                ):
                    v = ob.find('hl7:value', NS)
                    raw_bool = ((v.attrib.get('value') if v is not None else '') or (v.text if v is not None and v.text else '')).strip().lower()
                    if raw_bool == 'true':
                        medically_confirmed = "Yes"
                    elif raw_bool == 'false':
                        medically_confirmed = "NO"
                    else:
                        medically_confirmed = ""
                    break

            # Stable key
            key = normalize_text(event_term) or normalize_text(rrt_term) or clean_value(llt_code)
            if not key:
                continue

            out.append({
                "Event Term": clean_value(event_term),
                "PT": clean_value(locals().get("pt_term", "") or event_term),
                "RRT": clean_value(rrt_term),
                "Country": clean_value(country),
                "Translation Term": clean_value(translation_term),
                "Highlighted by Reporter": clean_value(highlighted),
                "Medically Confirmed": medically_confirmed,
                "Seriousness": seriousness_disp,
                "Outcome": clean_value(outcome),
                "Event Start (raw)": start_raw,
                "Event Start": start_disp,
                "Event End (raw)": end_raw,
                "Event End": end_disp,
                "_key": key,
            })
        return out
    except Exception as e:
        st.warning(f"Events parse error: {e}")
        return out

# ---------------- Amendment / Nullification extraction ----------------
def _extract_inv_char_value(value_el: Optional[ET.Element], prefer_original_text: bool = False) -> str:
    if value_el is None:
        return ""
    if value_el.attrib.get("nullFlavor") == "MSK":
        return "Masked"

    def _original_text_text(node: ET.Element) -> str:
        ot = node.find("hl7:originalText", NS)
        if ot is None:
            return ""
        return clean_value(" ".join(t.strip() for t in ot.itertext() if (t or "").strip()))

    original_text = _original_text_text(value_el)
    ordered_attrs = ("displayName", "code", "value")
    if prefer_original_text:
        if original_text:
            return original_text
        for attr in ordered_attrs:
            val = clean_value(value_el.attrib.get(attr, ""))
            if val:
                return val
    else:
        for attr in ordered_attrs:
            val = clean_value(value_el.attrib.get(attr, ""))
            if val:
                return val
        if original_text:
            return original_text

    return clean_value(get_text(value_el))


def extract_amendment_nullification(root: ET.Element) -> Dict[str, str]:
    out = {
        "Nullification/Amendment Code": "",
        "Nullification/Amendment Reason": "",
    }

    for inv_char in findall(root, './/hl7:investigationCharacteristic'):
        code_el = inv_char.find('hl7:code', NS)
        if code_el is None:
            continue
        if (code_el.attrib.get('codeSystem') or '').strip() != NULLIFICATION_AMENDMENT_OID:
            continue

        disp = (code_el.attrib.get('displayName') or '').strip().lower()
        value_el = inv_char.find('hl7:value', NS)

        if disp == 'nullificationamendmentcode' and not out["Nullification/Amendment Code"]:
            out["Nullification/Amendment Code"] = _extract_inv_char_value(value_el, prefer_original_text=False)
        elif disp == 'nullificationamendmentreason' and not out["Nullification/Amendment Reason"]:
            out["Nullification/Amendment Reason"] = _extract_inv_char_value(value_el, prefer_original_text=True)

    return out

# ---------------- Narrative extraction ----------------
def extract_narrative(root: ET.Element) -> str:
    narrative_elem = root.find('.//hl7:code[@code="PAT_ADV_EVNT"]/../hl7:text', NS)
    txt = narrative_elem.text if narrative_elem is not None else ''
    return clean_value(txt)

# ---------------- Model builder ----------------
def extract_model(xml_bytes: bytes, meddra_map: Optional[Dict[str, Dict[str, str]]] = None) -> Dict[str, Any]:
    try:
        root = ET.fromstring(xml_bytes)
    except Exception as e:
        return {"_error": f"XML parse error: {e}"}

    model: Dict[str, Any] = {}
    model["Sender ID"] = extract_sender_id(root)
    model["WWID"] = extract_wwid(root)
    model["First Sender Type"] = extract_first_sender_type(root)
    model.update(extract_td_frd_lrd(root))
    model["Reporters"] = extract_reporters_from_sourceReport(root)
    model["Patient"] = extract_patient(root)
    model["MedicalHistory"] = extract_medical_history(root, meddra_map=meddra_map)
    model["LabDetails"] = extract_labs(root, meddra_map=meddra_map)
    model["DrugHistory"] = extract_drug_history(root, meddra_map=meddra_map)
    model["DeathDetails"] = extract_death_details(root, meddra_map=meddra_map)

    products = extract_all_products(root, meddra_map=meddra_map)
    model["Products"] = products
    model["Events"] = extract_events(root, meddra_map=meddra_map)

    product_id_to_name = {
        (p.get("_pid") or '').strip(): (p.get("Drug") or '').strip()
        for p in products if (p.get("_pid") or '').strip()
    }
    reaction_id_to_term = build_reaction_id_to_term(root, meddra_map=meddra_map)

    model["Causality"] = extract_causality(
        root,
        product_id_to_name=product_id_to_name,
        reaction_id_to_term=reaction_id_to_term,
    )
    model["Amendment/Nullification"] = extract_amendment_nullification(root)
    model["Narrative"] = extract_narrative(root)
    return model

# --------------- Table builders ----------------
def compare_table(rows: List[Tuple[str, Any, Any]], treat_as_dates: bool = False) -> pd.DataFrame:
    disp: List[Dict[str, str]] = []
    for field, s, p in rows:
        s_txt = _textify(s).strip()
        p_txt = _textify(p).strip()
        if not s_txt and not p_txt:
            continue
        marker = mismatch_marker(s_txt, p_txt, is_date=treat_as_dates)
        status = "Mismatch" if marker else ("Missing" if (not s_txt or not p_txt) else "Match")
        disp.append({"Field": field, "Source": safe_disp(s_txt), "Processed": safe_disp(p_txt) + marker, "Status": status})
    return pd.DataFrame(disp) if disp else pd.DataFrame(columns=["Field", "Source", "Processed", "Status"])

def _is_comparison_df(df: pd.DataFrame) -> bool:
    return isinstance(df, pd.DataFrame) and {"Field", "Source", "Processed"}.issubset(set(df.columns))

def _filtered_table(df: pd.DataFrame) -> pd.DataFrame:
    if df is None or df.empty:
        return df
    out = df.copy()
    if _is_comparison_df(out):
        if "Status" not in out.columns:
            out["Status"] = out.apply(lambda r: "Mismatch" if "🔴" in str(r.get("Processed", "")) else ("Missing" if (str(r.get("Source", "")) == "—" or str(r.get("Processed", "")).replace(" 🔴", "") == "—") else "Match"), axis=1)
        if st.session_state.get("show_only_mismatches", False):
            out = out[out["Status"].isin(["Mismatch", "Missing"])]
    return out

def _style_comparison_df(df: pd.DataFrame):
    statuses = df.get("Status", pd.Series([""] * len(df), index=df.index))
    display_df = df.drop(columns=["Status"], errors="ignore")

    def row_style(row):
        status = str(statuses.loc[row.name]) if row.name in statuses.index else ""
        if status == "Mismatch":
            return ["background-color: #fee2e2"] * len(row)
        if status == "Missing":
            return ["background-color: #fef3c7"] * len(row)
        return [""] * len(row)

    return display_df.style.apply(row_style, axis=1)

def render_table(df: pd.DataFrame, use_container_width: bool = True):
    df2 = _filtered_table(df)
    if df2 is None or df2.empty:
        st.info("No rows to display for the current filter.")
        return
    if _is_comparison_df(df2):
        st.dataframe(_style_comparison_df(df2), use_container_width=use_container_width, hide_index=True)
    else:
        st.dataframe(df2.drop(columns=["Status"], errors="ignore"), use_container_width=use_container_width, hide_index=True)

def count_mismatch_rows(df: pd.DataFrame) -> int:
    if df is None or df.empty:
        return 0
    if "Status" in df.columns:
        return int(df["Status"].isin(["Mismatch", "Missing"]).sum())
    if "Processed" in df.columns:
        return int(df["Processed"].astype(str).str.contains("🔴", regex=False).sum())
    return 0

def section_header(title: str, anchor_id: str):
    st.markdown(f'<a id="{anchor_id}"></a>', unsafe_allow_html=True)
    st.subheader(title)

def _result_class(value: str) -> str:
    val = (value or "").strip().lower().replace("/", "")
    if val == "certain": return "qc-result-certain"
    if val == "probable": return "qc-result-probable"
    if val == "possible": return "qc-result-possible"
    if val == "unlikely": return "qc-result-unlikely"
    if val == "unassessable": return "qc-result-unassessable"
    return "qc-result-empty"


def make_amendment_nullification_table(src: Dict[str, str], prc: Dict[str, str]) -> pd.DataFrame:
    rows = [
        ("Nullification/Amendment Code", src.get("Nullification/Amendment Code", ""), prc.get("Nullification/Amendment Code", "")),
        ("Nullification/Amendment Reason", src.get("Nullification/Amendment Reason", ""), prc.get("Nullification/Amendment Reason", "")),
    ]
    return compare_table(rows, treat_as_dates=False)


def make_admin_table(src: Dict[str, Any], prc: Dict[str, Any]) -> pd.DataFrame:
    rows: List[Tuple[str, str, str]] = []
    rows.append(("Sender ID", src.get("Sender ID", ""), prc.get("Sender ID", "")))
    rows.append(("WWID", src.get("WWID", ""), prc.get("WWID", "")))
    rows.append(("First Sender Type", src.get("First Sender Type", ""), prc.get("First Sender Type", "")))
    src_td_disp = src.get("TD", "") or format_date(src.get("TD_raw", ""))
    prc_lrd_disp = prc.get("LRD", "") or format_date(prc.get("LRD_raw", ""))
    rows.append(("Day Zero", src_td_disp, prc_lrd_disp))

    parts = [
        compare_table([rows[0]], treat_as_dates=False),
        compare_table([rows[1]], treat_as_dates=False),
        compare_table([rows[2]], treat_as_dates=False),
        compare_table([rows[3]], treat_as_dates=True),
    ]
    parts = [df for df in parts if not df.empty]
    return pd.concat(parts, ignore_index=True) if parts else pd.DataFrame(columns=["Field", "Source", "Processed"])

# ✅ Reporter table builder
def make_reporter_pair_table(src_rep: Dict[str, Any], prc_rep: Dict[str, Any]) -> pd.DataFrame:
    fields_in_order = [
        "Reporter IDs",
        "Reporter Qualification",
        "Reporter Title",
        "Reporter Given Name(s)",
        "Reporter Family Name",
        "Reporter Organization",
        "Reporter Street",
        "Reporter City/Town",
        "Reporter State/Province",
        "Reporter Postal Code",
        "Reporter Country",
        "Reporter Phone(s)",
        "Reporter Email(s)",
        "Reporter Fax(es)",
    ]
    rows = [(label, src_rep.get(label, ""), prc_rep.get(label, "")) for label in fields_in_order]
    return compare_table(rows, treat_as_dates=False)

# ✅ Patient table in requested order
def make_patient_table(src_pat: Dict[str, str], prc_pat: Dict[str, str]) -> pd.DataFrame:
    fields = [
        "Initials",
        "DOB",
        "Age",
        "Age Group",
        "Height",
        "Weight",
        "Gender",
        "Patient Record Number",
        "DOD",
    ]
    rows = [(f, src_pat.get(f, ''), prc_pat.get(f, '')) for f in fields]
    return compare_table(rows, treat_as_dates=False)

# ------ Drug UI helpers ------
def drug_base_token(title: str) -> str:
    """Return first token of normalized drug name (e.g., 'Apixaban 5 mg' -> 'apixaban')."""
    norm = normalize_text(title or '')
    if not norm:
        return ''
    return norm.split()[0]

def _drug_name_key(rec: Dict[str, Any]) -> str:
    """Name-level match key used to group products.

    Primary key is the base token of the drug name (e.g., 'Apixaban 5 mg' -> 'apixaban').
    Multiple product rows can share the same base token (e.g., same drug recorded as Suspect and Concomitant).

    We therefore group at *name* level first, and then pair within a name-group using product Type.
    """
    title = (rec.get('Drug') or '').strip()
    if title:
        base = drug_base_token(title)
        if base:
            return f'name::{base}'

    # Fallbacks when drug name is missing
    pid = (rec.get('_pid') or '').strip().lower()
    if pid:
        return f'pid::{pid}'
    gid = (rec.get('_gid') or '').strip().lower()
    return gid or 'unknown'


def _norm_drug_type(rec: Dict[str, Any]) -> str:
    """Normalized product Type for pairing within the same name-group."""
    return normalize_text(clean_value(rec.get('Type') or '')) or 'unknown'


def group_drugs_by_name(products: List[Dict[str, Any]]) -> Tuple[Dict[str, List[Dict[str, Any]]], List[str]]:
    """Return (groups, order) where groups[name_key] = list of records in original order."""
    groups: Dict[str, List[Dict[str, Any]]] = {}
    order: List[str] = []
    for rec in products or []:
        k = _drug_name_key(rec)
        if k not in groups:
            groups[k] = []
            order.append(k)
        groups[k].append(rec)
    return groups, order


def pair_drug_records(
    src_list: List[Dict[str, Any]],
    prc_list: List[Dict[str, Any]],
) -> List[Tuple[Dict[str, Any], Dict[str, Any]]]:
    """Pair drug records inside one name-group.

    Rules:
    1) If both sides have exactly one record, pair them directly so Type changes show as a mismatch.
    2) Otherwise, pair by normalized Type first (Suspect with Suspect, Concomitant with Concomitant, etc.),
       preserving original order within each Type bucket. Leftovers are paired with empty dicts.

    Fixes the bug where Suspect and Concomitant entries for the same drug name collapse into one comparison row
    and another row gets omitted.
    """
    src_list = src_list or []
    prc_list = prc_list or []

    if len(src_list) == 1 and len(prc_list) == 1:
        return [(src_list[0], prc_list[0])]

    # Stable Type order based on first appearance (Source first, then Processed)
    type_order: List[str] = []
    def _add(t: str):
        if t not in type_order:
            type_order.append(t)

    for r in src_list:
        _add(_norm_drug_type(r))
    for r in prc_list:
        _add(_norm_drug_type(r))

    src_b: Dict[str, List[Dict[str, Any]]] = {t: [] for t in type_order}
    prc_b: Dict[str, List[Dict[str, Any]]] = {t: [] for t in type_order}

    for r in src_list:
        src_b[_norm_drug_type(r)].append(r)
    for r in prc_list:
        prc_b[_norm_drug_type(r)].append(r)

    pairs: List[Tuple[Dict[str, Any], Dict[str, Any]]] = []
    for t in type_order:
        s_q = src_b.get(t, [])
        p_q = prc_b.get(t, [])
        n = min(len(s_q), len(p_q))

        for i in range(n):
            pairs.append((s_q[i], p_q[i]))
        for i in range(n, len(s_q)):
            pairs.append((s_q[i], {}))
        for i in range(n, len(p_q)):
            pairs.append(({}, p_q[i]))

    return pairs


def make_drug_compare_table(src_rec: Dict[str, Any], prc_rec: Dict[str, Any]) -> pd.DataFrame:
    fields = [
        "Drug Name", "Ingredient", "Type", "Dosage Text", "Dose Value", "Dose Unit", "Start Date", "Stop Date",
        "Route", "Formulation", "Lot No", "MAH", "Action Taken", "Drug Obtain Country", "Indication",
    ]
    rows = [(f, src_rec.get(f, ''), prc_rec.get(f, '')) for f in fields]
    return compare_table(rows, treat_as_dates=False)


# ---- Permanent MedDRA master loader ----
from shared_context import BASE_DIR, DATA_DIR
MEDDRA_MASTER_BASENAME = "MedDRA"

def find_master_meddra_file() -> Optional[Path]:
    allowed_suffixes = (".xlsx", ".xlsm", ".csv")
    if not DATA_DIR.exists():
        return None
    for suffix in allowed_suffixes:
        candidate = DATA_DIR / f"{MEDDRA_MASTER_BASENAME}{suffix}"
        if candidate.exists():
            return candidate
    wanted = {f"{MEDDRA_MASTER_BASENAME}{suffix}".lower() for suffix in allowed_suffixes}
    for candidate in DATA_DIR.iterdir():
        if candidate.is_file() and candidate.name.lower() in wanted:
            return candidate
    return None

def meddra_cache_key(file_path: Path) -> Tuple[str, int, int]:
    stat = file_path.stat()
    return (str(file_path.resolve()), int(stat.st_mtime_ns), int(stat.st_size))

class _NamedBytesIO(io.BytesIO):
    def __init__(self, data: bytes, name: str):
        super().__init__(data)
        self.name = name

@st.cache_data(show_spinner=False)
def load_permanent_meddra_mapping_cached(file_path_text: str, modified_time_ns: int, file_size: int) -> Dict[str, Dict[str, str]]:
    file_path = Path(file_path_text)
    file_obj = _NamedBytesIO(file_path.read_bytes(), file_path.name)
    return load_meddra_mapping(file_obj)

def load_permanent_meddra_mapping() -> Dict[str, Dict[str, str]]:
    file_path = find_master_meddra_file()
    if not file_path:
        st.warning("Permanent MedDRA file not found. Add data/MedDRA.xlsx or data/MedDRA.csv to GitHub.")
        return {}
    try:
        file_path_text, modified_time_ns, file_size = meddra_cache_key(file_path)
        mapping = load_permanent_meddra_mapping_cached(file_path_text, modified_time_ns, file_size)
        st.caption(f"Using permanent MedDRA master: {file_path.name} ({len(mapping):,} LLT rows)")
        return mapping
    except Exception as e:
        st.error(f"Failed to read permanent MedDRA file ({file_path.name}): {e}")
        return {}

# ---------------- Factor-based Causality Assessment helpers ----------------
PHARMACOLOGICALLY_OPTIONS = ["", "Yes", "No"]
RC_OPTIONS = ["", "+Ve", "-Ve", "Unk", "NA"]
DECHALLENGE_OPTIONS = ["", "+Ve", "-Ve", "Unk", "NA"]
CONFOUNDING_OPTIONS = ["", "Yes", "No", "Plausible"]
TIME_RELATIONSHIP_OPTIONS = ["", "Yes", "Unk", "Improbable"]

def calculate_factor_based_causality(pharmacologically: str, rechallenge: str, response_to_dc: str, confounding_factor: str, time_relationship: str) -> str:
    """Calculate causality using the latest criteria/factor matrix."""
    pharm = str(pharmacologically or "").strip().lower()
    rc = str(rechallenge or "").strip().lower()
    dc = str(response_to_dc or "").strip().lower()
    conf = str(confounding_factor or "").strip().lower()
    time_rel = str(time_relationship or "").strip().lower()

    if not any([pharm, rc, dc, conf, time_rel]):
        return ""

    rc_positive_or_na = rc in {"+ve", "positive", "na", "not applicable", "+ve/na", "+ve/ na"}
    rc_negative_or_unknown = rc in {"-ve", "negative", "unk", "unknown", "- ve", "-ve/unk", "-ve/ unk"}
    rc_blank = rc == ""

    dc_positive = dc in {"+ve", "positive"}
    dc_negative_unknown_na = dc in {"-ve", "negative", "unk", "unknown", "na", "not applicable", "- ve", "-ve/unk", "-ve/ unk"}
    dc_blank = dc == ""

    pharm_yes = pharm == "yes"
    pharm_no = pharm == "no"
    pharm_blank = pharm == ""

    conf_yes = conf == "yes"
    conf_no = conf == "no"
    conf_blank = conf == ""
    conf_no_or_blank = conf in {"no", ""}
    conf_plausible = conf == "plausible"

    time_yes = time_rel == "yes"
    time_improbable = time_rel in {"improbable", "improbabe"}
    time_unknown = time_rel in {"unk", "unknown"}

    # Certain: Pharm Yes + RC +Ve/NA + DC +Ve + Confounding No/Blank + Time Yes
    if pharm_yes and rc_positive_or_na and dc_positive and conf_no_or_blank and time_yes:
        return "Certain"

    # Probable 1: Pharm No + RC Any Value + DC +Ve + Confounding No/Blank + Time Yes
    if pharm_no and dc_positive and conf_no_or_blank and time_yes:
        return "Probable"

    # Probable 2: Pharm Any Value + RC -Ve/Unk + DC +Ve + Confounding No/Blank + Time Yes
    if rc_negative_or_unknown and dc_positive and conf_no_or_blank and time_yes:
        return "Probable"

    # Probable 3: Pharm Blank + RC Blank + DC +Ve + Confounding No/Blank + Time Yes
    if pharm_blank and rc_blank and dc_positive and conf_no_or_blank and time_yes:
        return "Probable"

    # Probable 4: Pharm Blank + RC +Ve + DC +Ve + Confounding No/Blank + Time Yes
    if pharm_blank and rc_positive_or_na and dc_positive and conf_no_or_blank and time_yes:
        return "Probable"

    # Possible 1: Pharm Any Value + RC Any Value + DC -Ve/Unk/NA + Confounding No + Time Yes
    if dc_negative_unknown_na and conf_no and time_yes:
        return "Possible"

    # Possible 2: Pharm Any Value + RC Any Value + DC Any Value + Confounding Yes + Time Yes
    if conf_yes and time_yes:
        return "Possible"

    # Possible 3: Pharm Blank + RC Blank + DC Blank + Confounding Blank + Time Yes
    if pharm_blank and rc_blank and dc_blank and conf_blank and time_yes:
        return "Possible"

    # Possible 4: Pharm Blank + RC Blank + DC Blank + Confounding No + Time Yes
    if pharm_blank and rc_blank and dc_blank and conf_no and time_yes:
        return "Possible"

    # Unlikely 1: Time Improbable, all other factors Any Value
    if time_improbable:
        return "Unlikely"

    # Unlikely 2/3: Confounding Plausible, all other factors Any Value
    if conf_plausible:
        return "Unlikely"

    # Unassessable: Time Unk, unless already matched above
    if time_unknown:
        return "Unassessable"

    return "Unassessable"

def build_causality_assessment_events(src_events: List[Dict[str, Any]], prc_events: List[Dict[str, Any]]) -> List[str]:
    labels: List[str] = []
    seen: Set[str] = set()
    for event_list in (src_events or [], prc_events or []):
        for idx, event in enumerate(event_list, start=1):
            term = clean_value(event.get("Event Term", "")) or clean_value(event.get("RRT", "")) or clean_value(event.get("LLT Term", "")) or clean_value(event.get("LLT Code", "")) or f"Event {idx}"
            key = normalize_text(term)
            if key and key not in seen:
                seen.add(key)
                labels.append(term)
    return labels

# --------------- UI: Upload & Parse ----------------
st.markdown("### Upload the XML files to compare")
st.info("MedDRA is loaded automatically from the GitHub repository data folder. Upload only the Source and Processed XML files here.")
col1, col2 = st.columns(2)
with col1:
    src_file = st.file_uploader("Source XML", type=["xml"], key="src_xml")
with col2:
    prc_file = st.file_uploader("Processed XML", type=["xml"], key="prc_xml")

meddra_map = load_permanent_meddra_mapping()
_listedness_path = find_master_meddra_file()
_listedness_candidates = [DATA_DIR / "Listedness_CX.xlsx", DATA_DIR / "Listedness_CX.csv"]
listedness_index = {}
for _lp in _listedness_candidates:
    if _lp.exists():
        _ldf = pd.read_csv(_lp) if _lp.suffix.lower() == ".csv" else pd.read_excel(_lp, engine="openpyxl")
        listedness_index = dataframe_to_index(_ldf)
        break

if not (src_file and prc_file):
    st.info("Please upload **both** Source and Processed XML files to view the tabular comparison.")
    st.stop()

src_bytes = src_file.read()
prc_bytes = prc_file.read()

with st.spinner("Parsing Source..."):
    src = extract_model(src_bytes, meddra_map=meddra_map)
with st.spinner("Parsing Processed..."):
    prc = extract_model(prc_bytes, meddra_map=meddra_map)

if src.get("_error") or prc.get("_error"):
    st.error(f"Source error: {src.get('_error', '-')}\nProcessed error: {prc.get('_error', '-')}")
    st.stop()

# ---------------- UI controls and navigation ----------------
st.sidebar.title("QC Navigation")
st.sidebar.checkbox("Show only mismatches / missing", key="show_only_mismatches")
st.sidebar.markdown("""
- [Admin](#admin)
- [Reporter](#reporter)
- [Patient](#patient)
- [Drug History](#drug-history)
- [Medical History](#medical-history)
- [Drug](#drug)
- [Event](#event)
- [Amendment / Nullification](#amendment-nullification)
- [Lab](#lab)
- [Narrative](#narrative)
- [Causality](#causality)
""")

# ==========================================================
# DISPLAY — ORDER YOU REQUESTED
# ==========================================================
# 1) Admin
section_header("Admin", "admin")
admin_df = make_admin_table(src, prc)
if not admin_df.empty:
    render_table(admin_df)
else:
    st.markdown('<div style="color:#888">No header/admin values present.</div>', unsafe_allow_html=True)

# 2) Reporter
section_header("Reporter", "reporter")
src_reps = src.get("Reporters", []) or []
prc_reps = prc.get("Reporters", []) or []
n_boxes = max(len(src_reps), len(prc_reps))
if n_boxes == 0:
    st.markdown('<div style="color:#888">No reporters (sourceReport) present.</div>', unsafe_allow_html=True)
else:
    for i in range(n_boxes):
        srep = src_reps[i] if i < len(src_reps) else {}
        prep = prc_reps[i] if i < len(prc_reps) else {}
        st.markdown(f'<h6 style="margin-top:0.5rem;margin-bottom:0.25rem;">Reporter {i+1}</h6><hr/>', unsafe_allow_html=True)
        try:
            r_df = make_reporter_pair_table(srep, prep)
            if not r_df.empty:
                render_table(r_df)
            else:
                st.markdown('<div style="color:#888">No values for this reporter.</div>', unsafe_allow_html=True)
        except Exception as e:
            st.error(f"Reporter {i+1} render error: {e}")
        st.markdown('<div style="height:0.5rem;"></div>', unsafe_allow_html=True)

# 3) Patient
section_header("Patient", "patient")
pat_df = make_patient_table(src.get("Patient", {}), prc.get("Patient", {}))
if not pat_df.empty:
    render_table(pat_df)
else:
    st.markdown('<div style="color:#888">No patient values present.</div>', unsafe_allow_html=True)

# 4) Drug History
section_header("Drug History", "drug-history")
src_dh = src.get("DrugHistory", []) or []
prc_dh = prc.get("DrugHistory", []) or []

def _idx_dh(lst: List[Dict[str, Any]]) -> Dict[str, Dict[str, Any]]:
    return {e.get("_key", ""): e for e in lst if e.get("_key", "")}

src_dh_idx = _idx_dh(src_dh)
prc_dh_idx = _idx_dh(prc_dh)
all_dh_keys = sorted(set(src_dh_idx) | set(prc_dh_idx))

def make_drughist_box_for_ui(src_rec: Dict[str, Any], prc_rec: Dict[str, Any], title: str):
    st.markdown(f'<h6 style="margin-top:0.5rem;margin-bottom:0.25rem;">Drug History: {title}</h6><hr/>', unsafe_allow_html=True)
    rows = [
        ("Drug", src_rec.get("Drug", ""), prc_rec.get("Drug", "")),
        ("Indication", src_rec.get("Indication", ""), prc_rec.get("Indication", "")),
        ("Reaction", src_rec.get("Reaction", ""), prc_rec.get("Reaction", "")),
        ("Start Date", src_rec.get("Start Date", ""), prc_rec.get("Start Date", "")),
        ("End Date", src_rec.get("End Date", ""), prc_rec.get("End Date", "")),
    ]
    df = compare_table(rows, treat_as_dates=True)
    if not df.empty:
        render_table(df)
    else:
        st.markdown('<div style="color:#888">No values for this drug-history item.</div>', unsafe_allow_html=True)
    st.markdown('<div style="height:0.5rem;"></div>', unsafe_allow_html=True)

if not all_dh_keys:
    st.markdown('<div style="color:#888">No drug history found.</div>', unsafe_allow_html=True)
else:
    for key in all_dh_keys:
        se = src_dh_idx.get(key, {})
        pe = prc_dh_idx.get(key, {})
        title = (se.get("Drug") or pe.get("Drug") or "(Unnamed drug)")
        make_drughist_box_for_ui(se, pe, title)

# 5) Medical History (after Drug History)
section_header("Medical History", "medical-history")

# Death Details mini-box at top of this section
src_dd_raw = src.get("DeathDetails", [])
prc_dd_raw = prc.get("DeathDetails", [])

def _death_details_as_list(value: Any) -> List[Dict[str, Any]]:
    if isinstance(value, list):
        return [v for v in value if isinstance(v, dict)]
    if isinstance(value, dict):
        return [value] if any(str(v).strip() for k, v in value.items() if not str(k).startswith('_')) else []
    return []

src_dd_list = _death_details_as_list(src_dd_raw)
prc_dd_list = _death_details_as_list(prc_dd_raw)

def make_death_details_table(src_dd: Dict[str, Any], prc_dd: Dict[str, Any]) -> pd.DataFrame:
    rows = [
        ("Reported Cause of Death", src_dd.get("Reported Cause of Death", ""), prc_dd.get("Reported Cause of Death", "")),
        ("Autopsy", src_dd.get("Autopsy", ""), prc_dd.get("Autopsy", "")),
    ]
    rows = [r for r in rows if (str(r[1]).strip() or str(r[2]).strip())]
    return compare_table(rows, treat_as_dates=False)

n_dd_boxes = max(len(src_dd_list), len(prc_dd_list))
any_dd_displayed = False
if n_dd_boxes:
    for i in range(n_dd_boxes):
        sdd = src_dd_list[i] if i < len(src_dd_list) else {}
        pdd = prc_dd_list[i] if i < len(prc_dd_list) else {}
        dd_df = make_death_details_table(sdd, pdd)
        if not dd_df.empty:
            any_dd_displayed = True
            st.markdown(f'<h6 style="margin-top:0.5rem;margin-bottom:0.25rem;">Death Details {i+1}</h6><hr/>', unsafe_allow_html=True)
            render_table(dd_df)
            st.markdown('<div style="height:0.5rem;"></div>', unsafe_allow_html=True)

# Existing Medical History items
src_mh = src.get("MedicalHistory", []) or []
prc_mh = prc.get("MedicalHistory", []) or []

def _idx_mh(lst: List[Dict[str, Any]]) -> Dict[str, Dict[str, Any]]:
    return {e.get("_key", ""): e for e in lst if e.get("_key", "")}

src_mh_idx = _idx_mh(src_mh)
prc_mh_idx = _idx_mh(prc_mh)
all_mh_keys = sorted(set(src_mh_idx) | set(prc_mh_idx))

def make_mh_box_for_ui(src_rec: Dict[str, Any], prc_rec: Dict[str, Any], title: str):
    st.markdown(f'<h6 style="margin-top:0.5rem;margin-bottom:0.25rem;">Medical History: {title}</h6><hr/>', unsafe_allow_html=True)
    llc = src_rec.get("LLT Code", "") or ""
    llt = src_rec.get("LLT Term", "") or ""
    plc = prc_rec.get("LLT Code", "") or ""
    plt = prc_rec.get("LLT Term", "") or ""
    if llt:
        llc = ""
    if plt:
        plc = ""
    pairs = [
        ("LLT", (llt or llc), (plt or plc)),
        ("Status", src_rec.get("Status", ""), prc_rec.get("Status", "")),
        ("Status (Continue)", src_rec.get("Status (Continue)", ""), prc_rec.get("Status (Continue)", "")),
        ("Comment", src_rec.get("Comment", ""), prc_rec.get("Comment", "")),
        (
            "Start Date",
            src_rec.get("Start Date", "") or format_date(src_rec.get("Start Date (raw)", "")),
            prc_rec.get("Start Date", "") or format_date(prc_rec.get("Start Date (raw)", "")),
        ),
        (
            "End Date",
            src_rec.get("End Date", "") or format_date(src_rec.get("End Date (raw)", "")),
            prc_rec.get("End Date", "") or format_date(prc_rec.get("End Date (raw)", "")),
        ),
    ]
    mh_df = compare_table(pairs, treat_as_dates=True)
    if not mh_df.empty:
        render_table(mh_df)
    else:
        st.markdown('<div style="color:#888">No values for this item.</div>', unsafe_allow_html=True)
    st.markdown('<div style="height:0.5rem;"></div>', unsafe_allow_html=True)

if not all_mh_keys and not any_dd_displayed:
    st.markdown('<div style="color:#888">No medical history found.</div>', unsafe_allow_html=True)
else:
    for key in all_mh_keys:
        se = src_mh_idx.get(key, {})
        pe = prc_mh_idx.get(key, {})
        title = se.get("LLT Term") or pe.get("LLT Term") or (se.get("LLT Code") or pe.get("LLT Code") or "(Unnamed history)")
        make_mh_box_for_ui(se, pe, title)

# 6) Drug
section_header("Drug", "drug")
src_prods = src.get("Products", [])
prc_prods = prc.get("Products", [])

src_groups, src_order = group_drugs_by_name(src_prods)
prc_groups, prc_order = group_drugs_by_name(prc_prods)

# Preserve display order: first-seen in Source, then any new drug-name groups from Processed
ordered_name_keys: List[str] = []
seen_name_keys: Set[str] = set()
for k in (src_order + prc_order):
    if k and k not in seen_name_keys:
        seen_name_keys.add(k)
        ordered_name_keys.append(k)

if not ordered_name_keys:
    st.markdown('<div style="color:#888">No products found in either file.</div>', unsafe_allow_html=True)
else:
    for name_key in ordered_name_keys:
        s_list = src_groups.get(name_key, [])
        p_list = prc_groups.get(name_key, [])
        pairs = pair_drug_records(s_list, p_list)
        if not pairs:
            continue

        for idx, (srec, prec) in enumerate(pairs, start=1):
            # Build a clear title. Include Type and an index if multiple rows exist for the same drug.
            drug_title = (
                (srec.get("Drug") or '').strip()
                or (prec.get("Drug") or '').strip()
                or (srec.get("_pid") or '').strip()
                or (prec.get("_pid") or '').strip()
                or "(Unnamed drug)"
            )
            type_title = (
                (srec.get("Type") or '').strip()
                or (prec.get("Type") or '').strip()
            )
            suffix = f" #{idx}" if len(pairs) > 1 else ""
            title = f"{drug_title} ({type_title}){suffix}" if type_title else f"{drug_title}{suffix}"

            st.markdown(
                f'<h6 style="margin-top:0.5rem;margin-bottom:0.25rem;">Drug: {title}</h6><hr/>',
                unsafe_allow_html=True
            )
            d_df = make_drug_compare_table(srec or {}, prec or {})
            if not d_df.empty:
                render_table(d_df)
            else:
                st.markdown('<div style="color:#888">No values to display for this drug.</div>', unsafe_allow_html=True)
            st.markdown('<div style="height:0.5rem;"></div>', unsafe_allow_html=True)

# 7) Event
section_header("Event", "event")
src_evts = src.get("Events", []) or []
prc_evts = prc.get("Events", []) or []

def _idx_events(lst: List[Dict[str, Any]]) -> Dict[str, Dict[str, Any]]:
    return {e.get("_key", ""): e for e in lst if e.get("_key", "")}

src_evt_idx = _idx_events(src_evts)
prc_evt_idx = _idx_events(prc_evts)
all_evt_keys = sorted(set(src_evt_idx) | set(prc_evt_idx))

def make_event_box_for_ui(src_rec: Dict[str, Any], prc_rec: Dict[str, Any], title: str):
    st.markdown(f'<h6 style="margin-top:0.5rem;margin-bottom:0.25rem;">Event: {title}</h6><hr/>', unsafe_allow_html=True)
    fields = [
        ("Event Term", "text"),
        ("RRT", "text"),
        ("Country", "text"),
        ("Translation Term", "text"),
        ("Highlighted by Reporter", "text"),
        ("Medically Confirmed", "text"),
        ("Seriousness", "text"),
        ("Outcome", "text"),
        ("Event Start", "date"),
        ("Event End", "date"),
    ]
    rows: List[Tuple[str, str, str]] = []
    for field, kind in fields:
        s_val = src_rec.get(field, "")
        p_val = prc_rec.get(field, "")
        if kind == "date":
            s_val = s_val or format_date(src_rec.get(field + " (raw)", ""))
            p_val = p_val or format_date(prc_rec.get(field + " (raw)", ""))
        rows.append((field, s_val, p_val))
    e_df = compare_table(rows, treat_as_dates=True)
    if not e_df.empty:
        render_table(e_df)
    else:
        st.markdown('<div style="color:#888">No values to display for this event.</div>', unsafe_allow_html=True)
    st.markdown('<div style="height:0.5rem;"></div>', unsafe_allow_html=True)

if not all_evt_keys:
    st.markdown('<div style="color:#888">No events found in either file.</div>', unsafe_allow_html=True)
else:
    for key in all_evt_keys:
        se = src_evt_idx.get(key, {})
        pe = prc_evt_idx.get(key, {})
        title = se.get("Event Term") or pe.get("Event Term") or se.get("RRT") or pe.get("RRT") or "(Unnamed event)"
        make_event_box_for_ui(se, pe, title)

# 8) Amendment / Nullification
section_header("Amendment / Nullification", "amendment-nullification")
src_amend = src.get("Amendment/Nullification", {}) or {}
prc_amend = prc.get("Amendment/Nullification", {}) or {}
amend_df = make_amendment_nullification_table(src_amend, prc_amend)
if not amend_df.empty:
    render_table(amend_df)
else:
    st.markdown('<div style="color:#888">No amendment/nullification values present in either file.</div>', unsafe_allow_html=True)

# 9) Lab
section_header("Lab", "lab")
src_lab = src.get("LabDetails", []) or []
prc_lab = prc.get("LabDetails", []) or []

def _group_lab_by_key(lst: List[Dict[str, Any]]) -> Tuple[Dict[str, List[Dict[str, Any]]], List[str]]:
    groups: Dict[str, List[Dict[str, Any]]] = {}
    order: List[str] = []
    for e in lst:
        key = e.get("_key", "") or normalize_text(e.get("LLT Term", "") or e.get("LLT Code", ""))
        if not key:
            key = f"__lab_{len(order)}"
        if key not in groups:
            groups[key] = []
            order.append(key)
        groups[key].append(e)
    return groups, order

def make_lab_box_for_ui(src_rec: Dict[str, Any], prc_rec: Dict[str, Any], title: str):
    st.markdown(f'<h6 style="margin-top:0.5rem;margin-bottom:0.25rem;">Lab: {title}</h6><hr/>', unsafe_allow_html=True)
    ll_s = src_rec.get("LLT Term", "") or ""
    lc_s = src_rec.get("LLT Code", "") or ""
    ll_p = prc_rec.get("LLT Term", "") or ""
    lc_p = prc_rec.get("LLT Code", "") or ""
    if ll_s:
        lc_s = ""  # show only term if available
    if ll_p:
        lc_p = ""
    display_ll_s = ll_s if ll_s else lc_s
    display_ll_p = ll_p if ll_p else lc_p

    rows = [
        ("LLT", display_ll_s, display_ll_p),
        ("Result", src_rec.get("Result", ""), prc_rec.get("Result", "")),
        ("Result Date", src_rec.get("Result Date", ""), prc_rec.get("Result Date", "")),
    ]
    df = compare_table(rows, treat_as_dates=True)
    if not df.empty:
        render_table(df)
    else:
        st.markdown('<div style="color:#888">No values for this lab item.</div>', unsafe_allow_html=True)
    st.markdown('<div style="height:0.5rem;"></div>', unsafe_allow_html=True)

src_lab_groups, src_lab_order = _group_lab_by_key(src_lab)
prc_lab_groups, prc_lab_order = _group_lab_by_key(prc_lab)
ordered_lab_keys: List[str] = []
seen_lab_keys: Set[str] = set()
for k in (src_lab_order + prc_lab_order):
    if k and k not in seen_lab_keys:
        seen_lab_keys.add(k)
        ordered_lab_keys.append(k)

if not ordered_lab_keys:
    st.markdown('<div style="color:#888">No lab details found.</div>', unsafe_allow_html=True)
else:
    for key in ordered_lab_keys:
        s_group = src_lab_groups.get(key, [])
        p_group = prc_lab_groups.get(key, [])
        n = max(len(s_group), len(p_group))
        for i in range(n):
            se = s_group[i] if i < len(s_group) else {}
            pe = p_group[i] if i < len(p_group) else {}
            base_title = (se.get("LLT Term") or pe.get("LLT Term") or se.get("LLT Code") or pe.get("LLT Code") or "(Unnamed lab)")
            title = f"{base_title} ({i+1})" if n > 1 else base_title
            make_lab_box_for_ui(se, pe, title)

# 10) Narrative
section_header("Narrative", "narrative")
src_narr_full = src.get("Narrative", "") or ""
prc_narr_full = prc.get("Narrative", "") or ""
if not has_value(src_narr_full) and not has_value(prc_narr_full):
    st.markdown('<div style="color:#888">No narrative present in either file.</div>', unsafe_allow_html=True)
else:
    st.markdown('<h6>Source</h6>', unsafe_allow_html=True)
    st.markdown(f'<div style="white-space:pre-wrap">{src_narr_full if src_narr_full else "—"}</div>', unsafe_allow_html=True)
    st.markdown('<div style="height:0.5rem;"></div>', unsafe_allow_html=True)
    st.markdown('<h6>Processed</h6>', unsafe_allow_html=True)
    st.markdown(f'<div style="white-space:pre-wrap">{prc_narr_full if prc_narr_full else "—"}</div>', unsafe_allow_html=True)
    st.markdown('<div style="height:0.5rem;"></div>', unsafe_allow_html=True)
    ncol1, ncol2 = st.columns(2)
    ncol1.download_button("Download Source Narrative", data=src_narr_full.encode("utf-8"), file_name="source_narrative.txt", mime="text/plain")
    ncol2.download_button("Download Processed Narrative", data=prc_narr_full.encode("utf-8"), file_name="processed_narrative.txt", mime="text/plain")

# 11) Causality — SINGLE CONSOLIDATED TABLE
section_header("Causality", "causality")

# Dechallenge reminder: show once in the Causality section when any Source or
# Processed product has Action Taken recorded as Drug withdrawn.
def _has_drug_withdrawn(products: List[Dict[str, Any]]) -> bool:
    for product in products or []:
        action_taken = normalize_text(str(product.get("Action Taken", "") or ""))
        if "drug withdrawn" in action_taken:
            return True
    return False

if _has_drug_withdrawn(src.get("Products", []) or []) or _has_drug_withdrawn(prc.get("Products", []) or []):
    st.warning("Kindly Check DC", icon="⚠️")


def _caus_df(lst: List[Dict[str, Any]]) -> pd.DataFrame:
    if not lst:
        return pd.DataFrame(columns=["Drug", "Reaction", "Assessor", "Method", "Assessment"])
    rows: List[Dict[str, str]] = []
    for r in lst:
        rows.append({
            "Drug": r.get("Drug", ""),
            "Reaction": r.get("Reaction", ""),
            "Assessor": r.get("Assessor", ""),
            "Method": r.get("Method", ""),
            "Assessment": r.get("Assessment", ""),
        })
    return pd.DataFrame(rows)

st.markdown("#### Source")
src_caus_df = _caus_df(src.get("Causality", []) or [])
if not src_caus_df.empty:
    st.dataframe(src_caus_df, use_container_width=True)
else:
    st.markdown('<div style="color:#888">No causality rows in Source.</div>', unsafe_allow_html=True)

st.markdown("#### Processed")
prc_caus_df = _caus_df(prc.get("Causality", []) or [])
if not prc_caus_df.empty:
    st.dataframe(prc_caus_df, use_container_width=True)
else:
    st.markdown('<div style="color:#888">No causality rows in Processed.</div>', unsafe_allow_html=True)


st.markdown("---")
if st.button("🧪 Causality Assessment", key="open_factor_causality_assessment"):
    st.session_state["show_factor_causality_assessment"] = True

if st.session_state.get("show_factor_causality_assessment", False):
    st.markdown("#### Factor-based Causality Assessment")
    st.caption("Select factors in one row. The result appears below.")

    event_labels = build_causality_assessment_events(src.get("Events", []) or [], prc.get("Events", []) or [])
    if not event_labels:
        st.info("No events found for causality assessment.")
    else:
        selected_event = st.selectbox(
            "Event",
            options=event_labels,
            key="causality_single_event",
        )

        # Sequential validation: each next factor is enabled only after the previous factor is selected.
        factor_state_keys = [
            "causality_time_relationship",
            "causality_confounding_factor",
            "causality_response_to_dc",
            "causality_rc",
            "causality_pharmacologically",
        ]
        for state_key in factor_state_keys:
            if state_key not in st.session_state:
                st.session_state[state_key] = ""

        # If an earlier factor is cleared, automatically clear all downstream factors.
        if not st.session_state.get("causality_time_relationship"):
            for state_key in factor_state_keys[1:]:
                st.session_state[state_key] = ""
        elif not st.session_state.get("causality_confounding_factor"):
            for state_key in factor_state_keys[2:]:
                st.session_state[state_key] = ""
        elif not st.session_state.get("causality_response_to_dc"):
            for state_key in factor_state_keys[3:]:
                st.session_state[state_key] = ""
        elif not st.session_state.get("causality_rc"):
            st.session_state["causality_pharmacologically"] = ""

        c1, c2, c3, c4, c5 = st.columns(5)
        with c1:
            time_value = st.selectbox(
                "Time Relationship",
                TIME_RELATIONSHIP_OPTIONS,
                key="causality_time_relationship",
            )
        with c2:
            conf_value = st.selectbox(
                "Confounding Factor",
                CONFOUNDING_OPTIONS,
                key="causality_confounding_factor",
                disabled=not bool(time_value),
            )
        with c3:
            dc_value = st.selectbox(
                "Response to DC",
                DECHALLENGE_OPTIONS,
                key="causality_response_to_dc",
                disabled=not bool(conf_value),
            )
        with c4:
            rc_value = st.selectbox(
                "RC",
                RC_OPTIONS,
                key="causality_rc",
                disabled=not bool(dc_value),
            )
        with c5:
            pharm_value = st.selectbox(
                "Pharmacologically",
                PHARMACOLOGICALLY_OPTIONS,
                key="causality_pharmacologically",
                disabled=not bool(rc_value),
            )

        calculated_causality = calculate_factor_based_causality(
            pharm_value,
            rc_value,
            dc_value,
            conf_value,
            time_value,
        )

        st.markdown("##### Calculated Causality")
        rclass = _result_class(calculated_causality)
        result_text = calculated_causality if calculated_causality else "Select factor values to calculate causality."
        st.markdown(f"<div class='qc-result {rclass}'>{result_text}</div>", unsafe_allow_html=True)
        def _reset_causality_factor_values():
            for key in ["causality_time_relationship", "causality_confounding_factor", "causality_response_to_dc", "causality_rc", "causality_pharmacologically"]:
                st.session_state[key] = ""

        st.button(
            "Reset Causality Factors",
            key="reset_causality_factors",
            on_click=_reset_causality_factor_values,
        )


# 12) Listedness, displayed after Causality
section_header("Listedness", "listedness")
_all_products = (src.get("Products", []) or []) + (prc.get("Products", []) or [])
_all_events = (src.get("Events", []) or []) + (prc.get("Events", []) or [])

# Only products that are both Suspect and included in the Celix product list.
_ingredients = []
for _product in _all_products:
    _celix_name = celix_product_name(_product)
    if _celix_name and _celix_name not in _ingredients:
        _ingredients.append(_celix_name)

_pts = []
for _event in _all_events:
    _pt = clean_value(_event.get("PT", "")) or clean_value(_event.get("Event Term", ""))
    if _pt and _pt not in _pts:
        _pts.append(_pt)

_rows = []
_missing_rows = []
for _ingredient in _ingredients:
    for _pt in _pts:
        _expectedness, _comment = assess(listedness_index, _ingredient, _pt)
        _rows.append({
            "Active Ingredients": _ingredient,
            "PT": _pt,
            "Expectedness": _expectedness,
            "Comment": _comment,
        })
        if not has_exact_match(listedness_index, _ingredient, _pt):
            _missing_rows.append((_ingredient, _pt))

if _rows:
    st.dataframe(pd.DataFrame(_rows), use_container_width=True, hide_index=True)
else:
    st.info("No Celix suspect product/event PT combinations are available for listedness assessment.")

# Update controls appear in the same Listedness section, only for unmatched rows.
for _idx, (_ingredient, _pt) in enumerate(_missing_rows, start=1):
    with st.expander(f"Update missing match: {_ingredient} | {_pt}", expanded=False):
        render_missing_listedness_update(f"quality_missing_{_idx}", _ingredient, _pt)
