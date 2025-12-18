import io
import re
import unicodedata
from decimal import Decimal, InvalidOperation

import pandas as pd
import streamlit as st
from rapidfuzz import fuzz
from openpyxl import load_workbook
from openpyxl.styles import PatternFill

import phonenumbers
from phonenumbers.phonenumberutil import country_code_for_region


# ------------------ UI / Page config ------------------
st.set_page_config(page_title="Lead Quality Checker", page_icon="✅", layout="wide")

# Header highlight for "new" columns
HEADER_YELLOW = PatternFill(start_color="FFF59D", end_color="FFF59D", fill_type="solid")

# Cell colour coding
CELL_GREEN = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")  # light green
CELL_BLUE = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")   # light blue


# ------------------ Normalisation helpers ------------------
def norm_text(s) -> str:
    """Robust normalizer: trims, casefolds, collapses whitespace, removes NBSP, fixes unicode."""
    if s is None or (isinstance(s, float) and pd.isna(s)):
        return ""
    s = str(s)
    s = unicodedata.normalize("NFKC", s)
    s = s.replace("\u00A0", " ")            # NBSP -> space
    s = re.sub(r"\s+", " ", s).strip()      # collapse whitespace
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


# ------------------ Company ↔ domain helpers ------------------
SUFFIXES = {
    "ltd","limited","co","company","corp","corporation","inc","incorporated",
    "plc","public","llc","lp","llp","ulc","pc","pllc","sa","ag","nv","se","bv",
    "oy","ab","aps","as","kft","zrt","rt","sarl","sas","spa","gmbh","ug",
    "kg","kgaa","pte","pty","sdn","bhd","kk","k.k.","co.","ltd.","inc.","plc.",
    "holdings","holding","group"
}

def _clean_domain(domain: str) -> str:
    if not isinstance(domain, str):
        return ""
    domain = domain.lower().strip()
    domain = re.sub(r"^https?://", "", domain)
    domain = re.sub(r"/.*$", "", domain)
    domain = domain.replace("www.", "")
    return domain.strip()

def _domain_base(domain: str) -> str:
    d = _clean_domain(domain)
    if not d:
        return ""
    base = d.split(".")[0]
    base = re.sub(r"[^a-z0-9]", "", base.lower())
    return base

def _company_tokens(company: str) -> list[str]:
    if not isinstance(company, str):
        return []
    s = company.strip()
    s = re.sub(r"^\s*uk\s*[-–—:]\s*", "", s, flags=re.IGNORECASE)  # strip leading "UK -"
    s = re.sub(r"[^A-Za-z0-9\s]", " ", s)
    toks = [t.lower() for t in s.split() if t.strip()]
    toks = [t for t in toks if t not in SUFFIXES and t not in {"of","and","the","for","to","a"}]
    return toks

def _company_acronym(company: str) -> str:
    toks = _company_tokens(company)
    if not toks:
        return ""
    return "".join(t[0] for t in toks if t).upper()

def _is_subsequence(short: str, long: str) -> bool:
    it = iter(long)
    return all(ch in it for ch in short)

def compare_company_domain(company, domain) -> tuple[str, int, str]:
    """
    Returns: (Status, Score, Reason)
    Status: Likely Match | Unsure – Please Check | Likely NOT Match
    """
    c_raw = company if isinstance(company, str) else ""
    d_raw = domain if isinstance(domain, str) else ""

    if not c_raw or not d_raw:
        return "Unsure – Please Check", 0, "missing values"

    dbase = _domain_base(d_raw)
    if not dbase:
        return "Unsure – Please Check", 0, "missing/invalid domain"

    # 1) Electronic Arts -> EA -> ea.com etc.
    acr = _company_acronym(c_raw)
    if acr and dbase == acr.lower():
        return "Likely Match", 99, "company acronym equals domain"

    # 2) Company is short abbreviation (DLG) and appears as subsequence in domain base
    c_compact = re.sub(r"[^A-Za-z0-9]", "", c_raw).upper()
    if 2 <= len(c_compact) <= 6 and c_compact.isalpha():
        if _is_subsequence(c_compact.lower(), dbase):
            return "Likely Match", 92, "company abbreviation is subsequence of domain"

    # 3) Token containment
    toks = _company_tokens(c_raw)
    joined = "".join(toks)
    if joined and joined in dbase:
        return "Likely Match", 95, "company tokens contained in domain"

    # 4) Fuzzy
    token_str = " ".join(toks)
    score = int(max(
        fuzz.token_sort_ratio(token_str, dbase),
        fuzz.partial_ratio(token_str, dbase)
    ))
    if score >= 85:
        return "Likely Match", score, "strong fuzzy"
    if score >= 70:
        return "Unsure – Please Check", score, "weak fuzzy"
    return "Likely NOT Match", score, "low similarity"


# ------------------ Seniority parsing (NO Seniority_Logic column) ------------------
def parse_seniority(title):
    if not isinstance(title, str):
        return "Entry"
    t = title.lower().strip()
    if re.search(r"\bchief\b|\bcio\b|\bcto\b|\bceo\b|\bcfo\b|\bcoo\b|\bchro\b|\bpresident\b", t):
        return "C Suite"
    if re.search(r"\bvice president\b|\bvp\b|\bsvp\b", t):
        return "VP"
    if re.search(r"\bhead\b", t):
        return "Head"
    if re.search(r"\bdirector\b", t):
        return "Director"
    if re.search(r"\bmanager\b", t):
        return "Manager"
    return "Entry"


# ------------------ Phone cleaning + validation ------------------
def phone_to_string(raw_phone) -> str:
    """
    Converts phone values that may be floats/scientific notation into a clean digit string.
    Examples:
      4.4788E+11 -> 447880000000  (approx, depending on original)
      447880000000.0 -> 447880000000
    """
    if raw_phone is None:
        return ""
    s = str(raw_phone).strip()
    if s == "":
        return ""

    # Scientific notation as string
    if "e" in s.lower():
        try:
            s = format(int(Decimal(s)), "d")
        except (InvalidOperation, ValueError):
            pass

    # Float-like with .0
    if re.fullmatch(r"\d+\.0+", s):
        s = s.split(".", 1)[0]

    return s

def normalise_phone_for_check(raw_phone: str, region: str) -> str:
    """
    Clean just for validation (does NOT change output phone column):
    - converts scientific notation to digits when possible
    - keeps digits and '+'
    - converts leading 00 to +
    - if phone starts with region calling code but no '+', add '+'
    """
    phone = phone_to_string(raw_phone)

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
                return "Warning", f"Parsed region {actual_region} != expected {regions}"

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
def run_matching(master_bytes: bytes, picklist_bytes: bytes, apply_colours: bool, progress_cb=None) -> bytes:
    # IMPORTANT: dtype=str prevents scientific notation + keeps phone numbers intact
    df_master = pd.read_excel(io.BytesIO(master_bytes), dtype=str, keep_default_na=False)
    df_picklist = pd.read_excel(io.BytesIO(picklist_bytes), dtype=str, keep_default_na=False)

    df_out = df_master.copy()
    added_cols = []  # track new columns for yellow header highlight

    # Build canonical -> exact picklist label (source of truth formatting)
    picklist_country_label_by_canon = {}
    for col in ("lead_country", "c_country", "country"):
        if col in df_picklist.columns:
            for v in df_picklist[col]:
                v = str(v).strip()
                if v:
                    picklist_country_label_by_canon[canon_country_key(v)] = v

    # Exact pairs (master col, picklist col)
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
        added_cols.append(out_col)

        if master_col not in df_master.columns or pick_col not in df_picklist.columns:
            df_out[out_col] = "Column Missing"
            continue

        pick_map = {norm_text(v): str(v).strip() for v in df_picklist[pick_col] if str(v).strip() != ""}

        matches = []
        new_vals = []

        for raw_val in df_master[master_col]:
            raw_val = str(raw_val)
            v = raw_val

            # Country standardisation to picklist label (England -> United Kingdom)
            if master_col.strip().lower() in {"lead_country", "c_country", "country"}:
                canon = canon_country_key(v)
                if canon in picklist_country_label_by_canon:
                    v = picklist_country_label_by_canon[canon]

            key = norm_text(v)

            if key in pick_map:
                matches.append("Yes")
                new_vals.append(pick_map[key])  # exact picklist label
            else:
                matches.append("No")
                new_vals.append(v)

        df_out[out_col] = matches
        df_out[master_col] = new_vals

    # Country audit columns
    if progress_cb:
        progress_cb(0.30, "Standardising countries...")

    country_base_col = None
    for c in ["lead_country", "c_country", "country"]:
        if c in df_out.columns:
            country_base_col = c
            break

    if country_base_col:
        std_vals, notes = [], []
        for raw in df_out[country_base_col]:
            raw = str(raw)
            canon = canon_country_key(raw)
            std = picklist_country_label_by_canon.get(canon, raw.strip())
            std_vals.append(std)
            notes.append(f"{raw} → {std}" if norm_text(raw) and norm_text(std) and norm_text(raw) != norm_text(std) else "")
        df_out["Country_Standardised"] = std_vals
        df_out["Country_Change_Note"] = notes
        added_cols += ["Country_Standardised", "Country_Change_Note"]

    # Seniority (no Seniority_Logic column)
    if progress_cb:
        progress_cb(0.45, "Parsing seniority...")

    if "jobtitle" in df_out.columns:
        df_out["Parsed_Seniority"] = df_out["jobtitle"].apply(parse_seniority)
    else:
        df_out["Parsed_Seniority"] = ""
    added_cols.append("Parsed_Seniority")

    # Company/domain check
    if progress_cb:
        progress_cb(0.60, "Checking company vs domain...")

    # detect company column
    company_col = None
    for c in df_master.columns:
        if c.strip().lower() in {"companyname", "company", "company name", "company_name"}:
            company_col = c
            break

    # detect website/domain column
    domain_col = None
    for c in df_master.columns:
        if any(k in c.strip().lower() for k in ["website", "domain", "company_domain", "company domain", "web"]):
            domain_col = c
            break

    # email fallback
    email_col = None
    for c in df_master.columns:
        if "email" in c.lower():
            email_col = c
            break

    if company_col:
        statuses, scores, reasons = [], [], []
        for i in range(len(df_master)):
            comp = df_master.at[i, company_col]
            dom = ""

            if domain_col and str(df_master.at[i, domain_col]).strip():
                dom = df_master.at[i, domain_col]
            elif email_col and str(df_master.at[i, email_col]).strip():
                em = str(df_master.at[i, email_col])
                if "@" in em:
                    dom = em.split("@", 1)[1].strip()

            status, score, reason = compare_company_domain(str(comp), str(dom))
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

    added_cols += ["Company_Domain_Status", "Company_Domain_Score", "Company_Domain_Reason"]

    # Phone/country check — dynamic detection (telephone/tel/mobile/phone etc.)
    if progress_cb:
        progress_cb(0.75, "Checking phone vs country...")

    country_for_phone = "Country_Standardised" if "Country_Standardised" in df_out.columns else country_base_col

    phone_cols = []
    for c in df_master.columns:
        cl = c.strip().lower()
        if any(k in cl for k in ["phone", "telephone", "tel", "mobile", "cell", "contact number", "contactno", "phone number"]):
            if "phonetic" in cl:
                continue
            phone_cols.append(c)

    if country_for_phone and phone_cols:
        for pc in phone_cols:
            st_col = f"{pc}_PhoneCountry_Status"
            rs_col = f"{pc}_PhoneCountry_Reason"
            added_cols += [st_col, rs_col]

            out_status, out_reason = [], []
            for i in range(len(df_master)):
                raw_phone = df_master.at[i, pc]
                ctry = df_out.at[i, country_for_phone] if country_for_phone in df_out.columns else ""
                s, r = phone_country_check(raw_phone, ctry)
                out_status.append(s)
                out_reason.append(r)

            df_out[st_col] = out_status
            df_out[rs_col] = out_reason
    else:
        df_out["PhoneCountry_Status"] = "phone or country column not found"
        df_out["PhoneCountry_Reason"] = ""
        added_cols += ["PhoneCountry_Status", "PhoneCountry_Reason"]

    # Required fields completeness (if exists)
    if progress_cb:
        progress_cb(0.85, "Checking required fields...")

    required_candidates = ["email", "firstname", "lastname", "jobtitle", "companyname", "lead_country"]
    existing_required = [c for c in required_candidates if c in df_out.columns]

    if existing_required:
        missing_list = []
        for i in range(len(df_out)):
            missing = []
            for c in existing_required:
                v = df_out.at[i, c]
                if v is None or str(v).strip() == "":
                    missing.append(c)
            missing_list.append(", ".join(missing))
        df_out["Missing_Required_Fields"] = missing_list
        added_cols.append("Missing_Required_Fields")

    # Overall status (Warnings trigger REVIEW; Unsure triggers REVIEW; Unsure company triggers REVIEW)
    if progress_cb:
        progress_cb(0.92, "Building overall status...")

    match_cols = [c for c in df_out.columns if c.startswith("Match_")]
    phone_status_cols = [c for c in df_out.columns if c.endswith("_PhoneCountry_Status")] + (
        ["PhoneCountry_Status"] if "PhoneCountry_Status" in df_out.columns else []
    )
    good_company_status = {"Likely Match", "Match"}

    overall_status = []
    overall_issues = []

    for i in range(len(df_out)):
        issues = []

        for c in match_cols:
            if norm_text(df_out.at[i, c]) != "yes":
                issues.append(c)

        if "Company_Domain_Status" in df_out.columns:
            if str(df_out.at[i, "Company_Domain_Status"]).strip() not in good_company_status:
                issues.append("Company_Domain_Status")

        for c in phone_status_cols:
            if str(df_out.at[i, c]).strip() != "Match":
                issues.append(c)

        if "Missing_Required_Fields" in df_out.columns:
            if str(df_out.at[i, "Missing_Required_Fields"]).strip():
                issues.append("Missing_Required_Fields")

        if issues:
            overall_status.append("REVIEW")
            overall_issues.append("; ".join(issues))
        else:
            overall_status.append("PASS")
            overall_issues.append("")

    df_out["Overall_Status"] = overall_status
    df_out["Overall_Issues"] = overall_issues
    added_cols += ["Overall_Status", "Overall_Issues"]

    # Write Excel
    if progress_cb:
        progress_cb(0.96, "Writing Excel output...")

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df_out.to_excel(writer, index=False, sheet_name="Results")
    output.seek(0)

    # Apply formatting
    if apply_colours:
        wb = load_workbook(output)
        ws = wb["Results"]

        headers = [cell.value for cell in ws[1]]
        col_index = {str(v): i + 1 for i, v in enumerate(headers)}
        max_row = ws.max_row

        # Yellow header for added columns
        for col_name in set(added_cols):
            if col_name in col_index:
                ws.cell(row=1, column=col_index[col_name]).fill = HEADER_YELLOW

        def fill_column(col_name: str, green_fn):
            if col_name not in col_index:
                return
            cidx = col_index[col_name]
            for r in range(2, max_row + 1):
                val = ws.cell(row=r, column=cidx).value
                if val is None:
                    continue
                if green_fn(str(val)):
                    ws.cell(row=r, column=cidx).fill = CELL_GREEN
                else:
                    ws.cell(row=r, column=cidx).fill = CELL_BLUE

        # Match_*: green if Yes
        for mc in match_cols:
            fill_column(mc, lambda v: norm_text(v) == "yes")

        # Company status: green if Likely Match/Match
        fill_column("Company_Domain_Status", lambda v: norm_text(v) in {"likely match", "match"})

        # Phone statuses: green if Match
        for pc in phone_status_cols:
            fill_column(pc, lambda v: norm_text(v) == "match")

        # Missing required fields: green if blank
        if "Missing_Required_Fields" in col_index:
            cidx = col_index["Missing_Required_Fields"]
            for r in range(2, max_row + 1):
                val = ws.cell(row=r, column=cidx).value
                if val is None or str(val).strip() == "":
                    ws.cell(row=r, column=cidx).fill = CELL_GREEN
                else:
                    ws.cell(row=r, column=cidx).fill = CELL_BLUE

        # Overall Status: PASS green else blue
        fill_column("Overall_Status", lambda v: norm_text(v) == "pass")

        # Overall Issues: blank green else blue
        if "Overall_Issues" in col_index:
            cidx = col_index["Overall_Issues"]
            for r in range(2, max_row + 1):
                val = ws.cell(row=r, column=cidx).value
                if val is None or str(val).strip() == "":
                    ws.cell(row=r, column=cidx).fill = CELL_GREEN
                else:
                    ws.cell(row=r, column=cidx).fill = CELL_BLUE

        out2 = io.BytesIO()
        wb.save(out2)
        out2.seek(0)
        return out2.read()

    return output.read()


# ------------------ Streamlit UI ------------------
st.markdown("## Lead Quality Checker")
st.caption(
    "Upload the Lead Master + Picklist, then click **Run matching**.\n\n"
    "🟩 Green = OK / match, 🟦 Blue = review.\n"
    "📞 Phone checking works even if phone is in scientific notation (e.g. 4.48E+11) and does NOT change the phone value."
)

col1, col2 = st.columns(2)
with col1:
    master_file = st.file_uploader("Upload Lead Master (.xlsx)", type=["xlsx"], key="master")
with col2:
    picklist_file = st.file_uploader("Upload Picklist (.xlsx)", type=["xlsx"], key="picklist")

apply_colours = st.toggle("Colour-code results (green/blue) + yellow headers for new columns", value=True)

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
            apply_colours=apply_colours,
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
