import io
import re
import unicodedata

import pandas as pd
import streamlit as st
from rapidfuzz import fuzz
from openpyxl import load_workbook
from openpyxl.styles import PatternFill

import phonenumbers
from phonenumbers.phonenumberutil import country_code_for_region


# ------------------ UI / Page config ------------------
st.set_page_config(page_title="Lead Quality Checker", page_icon="✅", layout="wide")

HIGHLIGHT_FILL = PatternFill(start_color="FFF59D", end_color="FFF59D", fill_type="solid")  # soft yellow


# ------------------ Normalisation helpers ------------------
def norm_text(s) -> str:
    """Robust normalizer: trims, casefolds, collapses whitespace, removes NBSP, fixes unicode."""
    if s is None or (isinstance(s, float) and pd.isna(s)):
        return ""
    s = str(s)
    s = unicodedata.normalize("NFKC", s)
    s = s.replace("\u00A0", " ")
    s = re.sub(r"\s+", " ", s).strip()
    return s.casefold()


# ------------------ Country standardisation ------------------
COUNTRY_ALIASES_TO_CANON = {
    # UK nations/variants → united kingdom
    "uk": "united kingdom",
    "u.k.": "united kingdom",
    "gb": "united kingdom",
    "great britain": "united kingdom",
    "britain": "united kingdom",
    "england": "united kingdom",
    "scotland": "united kingdom",
    "wales": "united kingdom",
    "northern ireland": "united kingdom",

    # US variants → united states
    "us": "united states",
    "u.s.": "united states",
    "usa": "united states",
    "u.s.a.": "united states",
    "america": "united states",
    "united states of america": "united states",

    # UAE variants → united arab emirates
    "uae": "united arab emirates",
    "u.a.e.": "united arab emirates",
    "dubai": "united arab emirates",
    "abu dhabi": "united arab emirates",

    # Common safe alternates
    "holland": "netherlands",
    "czech republic": "czechia",
    "russian federation": "russia",
    "viet nam": "vietnam",
    "republic of korea": "south korea",
    "korea, south": "south korea",
}

def canon_country_key(s) -> str:
    x = norm_text(s)
    return COUNTRY_ALIASES_TO_CANON.get(x, x)


# For phone validation we map canonical country -> phonenumbers region.
# Add more if your campaigns include them.
COUNTRY_TO_REGIONS = {
    "united kingdom": ["GB"],
    "ireland": ["IE"],
    "united states": ["US"],
    "canada": ["CA"],
    "australia": ["AU"],
    "new zealand": ["NZ"],
    "singapore": ["SG"],
    "malaysia": ["MY"],
    "thailand": ["TH"],
    "indonesia": ["ID"],
    "philippines": ["PH"],
    "hong kong": ["HK"],
    "taiwan": ["TW"],
    "japan": ["JP"],
    "south korea": ["KR"],
    "china": ["CN"],
    "india": ["IN"],
    "vietnam": ["VN"],
    "france": ["FR"],
    "germany": ["DE"],
    "spain": ["ES"],
    "italy": ["IT"],
    "netherlands": ["NL"],
    "sweden": ["SE"],
    "norway": ["NO"],
    "denmark": ["DK"],
    "finland": ["FI"],
    "switzerland": ["CH"],
    "austria": ["AT"],
    "belgium": ["BE"],
    "portugal": ["PT"],
}


# ------------------ Company ↔ domain matching helpers ------------------
SUFFIXES = {
    "ltd","limited","co","company","corp","corporation","inc","incorporated",
    "plc","public","llc","lp","llp","ulc","pc","pllc","sa","ag","nv","se","bv",
    "oy","ab","aps","as","kft","zrt","rt","sarl","sas","spa","gmbh","ug",
    "kg","kgaa","pte","pty","sdn","bhd","kk","k.k.","co.","ltd.","inc.","plc.",
    "holdings","holding","group"
}

def _normalize_tokens(text: str) -> str:
    if not isinstance(text, str):
        return ""
    text = re.sub(r"[^a-zA-Z0-9\s]", " ", text.lower())
    parts = [w for w in text.split() if w and w not in SUFFIXES]
    return " ".join(parts).strip()

def _clean_domain(domain: str) -> str:
    if not isinstance(domain, str):
        return ""
    domain = domain.lower().strip()
    domain = re.sub(r"^https?://", "", domain)
    domain = re.sub(r"/.*$", "", domain)
    domain = domain.replace("www.", "")
    return domain.strip()

def compare_company_domain(company, domain) -> tuple[str, int, str]:
    c_raw = company if isinstance(company, str) else ""
    d_raw = domain if isinstance(domain, str) else ""
    c = _normalize_tokens(c_raw)
    d = _clean_domain(d_raw)
    if not c or not d:
        return "Unsure – Please Check", 0, "missing values"

    d_base = d.split(".")[0] if "." in d else d
    d_base = re.sub(r"[^a-zA-Z0-9]", "", d_base)
    c_compact = re.sub(r"[^a-zA-Z0-9]", "", c)

    if c_compact and c_compact in d_base:
        return "Likely Match", 95, "company contained in domain"
    if d_base and d_base in c_compact:
        return "Likely Match", 90, "domain contained in company"

    # fuzzy
    score = int(max(
        fuzz.token_sort_ratio(c, d_base),
        fuzz.partial_ratio(c, d_base)
    ))
    if score >= 85:
        return "Likely Match", score, "strong fuzzy"
    if score >= 70:
        return "Unsure – Please Check", score, "weak fuzzy"
    return "Likely NOT Match", score, "low similarity"


# ------------------ Seniority parsing ------------------
def parse_seniority(title):
    if not isinstance(title, str):
        return "Entry", "no title"
    t = title.lower().strip()
    if re.search(r"\bchief\b|\bcio\b|\bcto\b|\bceo\b|\bcfo\b|\bcoo\b|\bchro\b|\bpresident\b", t):
        return "C Suite", "c-level"
    if re.search(r"\bvice president\b|\bvp\b|\bsvp\b", t):
        return "VP", "vp"
    if re.search(r"\bhead\b", t):
        return "Head", "head"
    if re.search(r"\bdirector\b", t):
        return "Director", "director"
    if re.search(r"\bmanager\b", t):
        return "Manager", "manager"
    return "Entry", "default"


# ------------------ Phone validation ------------------
def normalise_phone_for_check(raw_phone: str, region: str) -> str:
    """
    Clean just for validation (does NOT change output phone column):
    - keeps digits and '+'
    - converts leading 00 to +
    - if phone starts with the region calling code (e.g. 44) but no '+', add '+'
    """
    phone = "" if raw_phone is None else str(raw_phone).strip()
    if not phone:
        return ""
    phone = re.sub(r"[^\d+]", "", phone)

    if phone.startswith("00"):
        phone = "+" + phone[2:]

    if phone.startswith("+"):
        return phone

    try:
        cc = str(country_code_for_region(region))
    except Exception:
        cc = ""

    if cc and phone.startswith(cc):
        return "+" + phone

    return phone

def phone_country_check(raw_phone, country_label) -> tuple[str, str]:
    """
    Returns (Status, Reason)
    Status: Match | Warning | Unsure
    """
    if raw_phone is None or (isinstance(raw_phone, float) and pd.isna(raw_phone)) or str(raw_phone).strip() == "":
        return "Unsure", "Missing phone"
    if country_label is None or (isinstance(country_label, float) and pd.isna(country_label)) or str(country_label).strip() == "":
        return "Unsure", "Missing country"

    canon = canon_country_key(country_label)
    regions = COUNTRY_TO_REGIONS.get(canon, [])
    if not regions:
        return "Unsure", f"No phone region mapping for '{country_label}'"

    for region in regions:
        phone_for_check = normalise_phone_for_check(raw_phone, region)
        try:
            num = phonenumbers.parse(phone_for_check, region)
            actual_region = phonenumbers.region_code_for_number(num) or ""

            if actual_region and actual_region not in regions:
                return "Warning", f"Parsed region {actual_region} does not match expected {regions}"

            if not phonenumbers.is_possible_number(num):
                return "Warning", "Number not possible for country"
            if not phonenumbers.is_valid_number(num):
                return "Warning", "Number not valid for country"

            inferred = ""
            s = str(raw_phone).strip()
            digits_only = re.sub(r"[^\d]", "", s)
            try:
                cc = str(country_code_for_region(region))
            except Exception:
                cc = ""
            if s.startswith("00") or (cc and digits_only.startswith(cc) and not s.startswith("+")):
                inferred = " (international prefix inferred)"

            return "Match", "Valid for country" + inferred

        except Exception:
            continue

    return "Unsure", "Could not parse phone for supplied country"


# ------------------ Main processing ------------------
def run_matching(master_bytes: bytes, picklist_bytes: bytes, highlight_changes: bool, progress_cb=None) -> bytes:
    df_master = pd.read_excel(io.BytesIO(master_bytes))
    df_picklist = pd.read_excel(io.BytesIO(picklist_bytes))

    df_out = df_master.copy()
    corrected_cells = set()

    # Build canonical -> exact picklist label (source of truth formatting)
    picklist_country_label_by_canon = {}
    for col in ("lead_country", "c_country", "country"):
        if col in df_picklist.columns:
            for v in df_picklist[col].dropna().astype(str):
                v = v.strip()
                if v:
                    picklist_country_label_by_canon[canon_country_key(v)] = v

    # Pairs where we try to match master values to picklist values (case/space-insensitive)
    EXACT_PAIRS = [
        ("companyname", "companyname"),
        ("c_country", "c_country"),
        ("c_state", "c_state"),
        ("lead_country", "lead_country"),
        ("departments", "departments"),
        ("c_industry", "c_industry"),
        ("asset_title", "asset_title"),
    ]

    if progress_cb:
        progress_cb(0.15, "Running picklist checks...")

    for master_col, pick_col in EXACT_PAIRS:
        out_col = f"Match_{master_col}"
        if master_col not in df_master.columns or pick_col not in df_picklist.columns:
            df_out[out_col] = "Column Missing"
            continue

        # picklist map: normalized -> exact label
        pick_map = {norm_text(v): str(v).strip() for v in df_picklist[pick_col].dropna().astype(str)}

        matches = []
        new_vals = []

        for i, raw_val in enumerate(df_master[master_col].fillna("").astype(str)):
            v = raw_val

            # If this is a country column, standardise to picklist label first (England -> United Kingdom)
            if master_col.strip().lower() in {"lead_country", "c_country", "country"}:
                canon = canon_country_key(v)
                if canon in picklist_country_label_by_canon:
                    v = picklist_country_label_by_canon[canon]

            key = norm_text(v)
            if key in pick_map:
                matches.append("Yes")
                desired = pick_map[key]
                new_vals.append(desired)

                # Only mark as corrected if the visible value changed
                if str(desired).strip() != str(raw_val).strip():
                    corrected_cells.add((master_col, i + 2))  # excel row index (+ header)
            else:
                matches.append("No")
                new_vals.append(v)

        df_out[out_col] = matches
        df_out[master_col] = new_vals

    # Country audit columns (using whichever is present)
    if progress_cb:
        progress_cb(0.30, "Standardising countries...")

    country_base_col = None
    for c in ["lead_country", "c_country", "country"]:
        if c in df_out.columns:
            country_base_col = c
            break

    if country_base_col:
        std_vals, notes = [], []
        for raw in df_out[country_base_col].fillna("").astype(str):
            canon = canon_country_key(raw)
            std = picklist_country_label_by_canon.get(canon, raw.strip())
            std_vals.append(std)
            notes.append(f"{raw} → {std}" if norm_text(raw) and norm_text(std) and norm_text(raw) != norm_text(std) else "")
        df_out["Country_Standardised"] = std_vals
        df_out["Country_Change_Note"] = notes

    # Seniority
    if progress_cb:
        progress_cb(0.45, "Parsing seniority...")

    if "jobtitle" in df_out.columns:
        parsed = df_out["jobtitle"].apply(parse_seniority)
        df_out["Parsed_Seniority"] = parsed.apply(lambda x: x[0])
        df_out["Seniority_Logic"] = parsed.apply(lambda x: x[1])
    else:
        df_out["Parsed_Seniority"] = ""
        df_out["Seniority_Logic"] = "jobtitle column not found"

    # Company/domain check
    if progress_cb:
        progress_cb(0.60, "Checking company vs domain...")

    company_col = None
    for c in df_master.columns:
        if c.strip().lower() in {"companyname", "company", "company name", "company_name"}:
            company_col = c
            break

    domain_col = None
    for c in df_master.columns:
        if c.strip().lower() in {"website", "domain", "email domain", "email_domain", "company_domain", "company domain"}:
            domain_col = c
            break

    email_col = None
    for c in df_master.columns:
        if c.strip().lower() == "email" or "email" in c.lower():
            email_col = c
            break

    if company_col:
        statuses, scores, reasons = [], [], []
        for i in range(len(df_master)):
            comp = df_master.at[i, company_col]
            dom = None

            if domain_col and pd.notna(df_master.at[i, domain_col]):
                dom = df_master.at[i, domain_col]
            elif email_col and pd.notna(df_master.at[i, email_col]):
                em = str(df_master.at[i, email_col])
                if "@" in em:
                    dom = em.split("@", 1)[1].strip()

            status, score, reason = compare_company_domain(comp, dom)
            statuses.append(status)
            scores.append(score)
            reasons.append(reason)

        df_out["Company_Domain_Status"] = statuses
        df_out["Company_Domain_Score"] = scores
        df_out["Company_Domain_Reason"] = reasons
    else:
        df_out["Company_Domain_Status"] = "company column not found"
        df_out["Company_Domain_Score"] = 0
        df_out["Company_Domain_Reason"] = "company column not found"

    # Phone/country check (does not modify phone values)
    if progress_cb:
        progress_cb(0.78, "Checking phone vs country...")

    country_for_phone = "Country_Standardised" if "Country_Standardised" in df_out.columns else country_base_col

    # Detect phone columns
    phone_cols = []
    for c in df_master.columns:
        cl = c.strip().lower()
        if cl in {
            "phone", "phone_main", "phonemain", "phone main", "phone number", "phonenumber",
            "mobile", "mobilephone", "phone_mobile", "phonemobile", "phone mobile"
        }:
            phone_cols.append(c)

    if country_for_phone and phone_cols:
        for pc in phone_cols:
            st_col = f"{pc}_PhoneCountry_Status"
            rs_col = f"{pc}_PhoneCountry_Reason"
            out_status, out_reason = [], []

            for i in range(len(df_master)):
                raw_phone = df_master.at[i, pc]
                ctry = df_out.at[i, country_for_phone]
                s, r = phone_country_check(raw_phone, ctry)
                out_status.append(s)
                out_reason.append(r)

            df_out[st_col] = out_status
            df_out[rs_col] = out_reason
    else:
        df_out["PhoneCountry_Status"] = "phone or country column not found"
        df_out["PhoneCountry_Reason"] = ""

    # Required fields completeness (only if those columns exist)
    if progress_cb:
        progress_cb(0.90, "Checking required fields...")

    required_candidates = ["email", "firstname", "lastname", "jobtitle", "companyname", "lead_country"]
    existing_required = [c for c in required_candidates if c in df_out.columns]

    if existing_required:
        missing_list = []
        for i in range(len(df_out)):
            missing = []
            for c in existing_required:
                v = df_out.at[i, c]
                if v is None or (isinstance(v, float) and pd.isna(v)) or str(v).strip() == "":
                    missing.append(c)
            missing_list.append(", ".join(missing))
        df_out["Missing_Required_Fields"] = missing_list

    # Write to Excel
    if progress_cb:
        progress_cb(0.95, "Writing Excel output...")

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df_out.to_excel(writer, index=False, sheet_name="Results")
    output.seek(0)

    # Highlight corrected cells (slow on huge files, but optional)
    if highlight_changes and corrected_cells:
        wb = load_workbook(output)
        ws = wb["Results"]
        header = [cell.value for cell in ws[1]]
        col_index = {str(v): i + 1 for i, v in enumerate(header)}

        for col_name, row_num in corrected_cells:
            if col_name in col_index:
                ws.cell(row=row_num, column=col_index[col_name]).fill = HIGHLIGHT_FILL

        out2 = io.BytesIO()
        wb.save(out2)
        out2.seek(0)
        return out2.read()

    return output.read()


# ------------------ Streamlit UI ------------------
st.markdown("## Lead Quality Checker")
st.caption(
    "Upload the Lead Master + Picklist. Click **Run matching** to generate the output. "
    "Countries are standardised to the picklist format. Phone checks run in the background "
    "(the phone value in your output stays unchanged)."
)

col1, col2 = st.columns(2)
with col1:
    master_file = st.file_uploader("Upload Lead Master (.xlsx)", type=["xlsx"], key="master")
with col2:
    picklist_file = st.file_uploader("Upload Picklist (.xlsx)", type=["xlsx"], key="picklist")

# Highlighting can slow big files — default OFF for speed
highlight = st.toggle("Highlight corrected cells (yellow)", value=False)

# Keep last result in session so toggling doesn't re-run automatically
if "last_output_bytes" not in st.session_state:
    st.session_state.last_output_bytes = None

run_btn = st.button(
    "▶ Run matching",
    type="primary",
    use_container_width=True,
    disabled=not (master_file and picklist_file),
)

if run_btn:
    progress = st.progress(0.0, text="Starting...")

    def prog(p, text=""):
        progress.progress(min(max(float(p), 0.0), 1.0), text=text)

    try:
        output_bytes = run_matching(
            master_file.read(),
            picklist_file.read(),
            highlight_changes=highlight,
            progress_cb=prog
        )
        st.session_state.last_output_bytes = output_bytes
        st.success("Processing complete.")
    except Exception as e:
        st.session_state.last_output_bytes = None
        st.error(f"Error: {e}")

if st.session_state.last_output_bytes:
    st.download_button(
        label="⬇ Download Processed File",
        data=st.session_state.last_output_bytes,
        file_name="Full_Check_Results.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True
    )
