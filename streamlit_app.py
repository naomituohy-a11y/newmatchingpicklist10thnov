import io
import os
import re
import unicodedata
import pandas as pd
import streamlit as st
from rapidfuzz import fuzz
from openpyxl import load_workbook
from openpyxl.styles import PatternFill

# Optional: phone validation
import phonenumbers
from phonenumbers.phonenumberutil import country_code_for_region

# ------------------ page config (UI) ------------------
st.set_page_config(page_title="Lead Quality Checker", page_icon="✅", layout="wide")

# ------------------ constants ------------------

# Common legal suffixes / noise words to ignore in company matching
SUFFIXES = {
    "ltd","limited","co","company","corp","corporation","inc","incorporated",
    "plc","public","llc","lp","llp","ulc","pc","pllc","sa","ag","nv","se","bv",
    "oy","ab","aps","as","kft","zrt","rt","sarl","sas","spa","gmbh","ug","bvba",
    "kg","kgaa","pte","pty","sdn","bhd","kk","k.k.","co.","ltd.","inc.","plc.",
    "holdings","holding","group"  # we will still treat "group" specially in fuzzy step
}

THRESHOLD_WEAK = 70
THRESHOLD_STRONG = 85

# Alias map: raw data variants -> canonical key (lowercase)
# NOTE: Keep conservative. Anything politically sensitive should be explicitly approved.
COUNTRY_ALIASES_TO_CANON = {
    # United Kingdom nations/variants -> united kingdom
    "uk":"united kingdom",
    "u.k.":"united kingdom",
    "great britain":"united kingdom",
    "britain":"united kingdom",
    "gb":"united kingdom",
    "england":"united kingdom",
    "scotland":"united kingdom",
    "wales":"united kingdom",
    "northern ireland":"united kingdom",

    # United States variants -> united states
    "us":"united states",
    "u.s.":"united states",
    "usa":"united states",
    "u.s.a.":"united states",
    "america":"united states",
    "united states of america":"united states",

    # UAE variants -> united arab emirates
    "uae":"united arab emirates",
    "u.a.e.":"united arab emirates",
    "dubai":"united arab emirates",
    "abu dhabi":"united arab emirates",

    # Common safe alternates
    "holland":"netherlands",
    "czech republic":"czechia",
    "russian federation":"russia",
    "viet nam":"vietnam",
    "republic of korea":"south korea",
    "korea, south":"south korea",
}

# Country (canonical key) -> phone region code(s) for phonenumbers parsing/validation.
# Add more as your campaigns expand.
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

HIGHLIGHT_FILL = PatternFill(start_color="FFF59D", end_color="FFF59D", fill_type="solid")  # soft yellow

# ------------------ helpers ------------------

def norm_text(s) -> str:
    """Robust normalizer: fixes case, whitespace, and invisible unicode (NBSP etc.)."""
    if s is None or (isinstance(s, float) and pd.isna(s)):
        return ""
    s = str(s)
    s = unicodedata.normalize("NFKC", s)
    s = s.replace("\u00A0", " ")                  # NBSP -> space
    s = re.sub(r"\s+", " ", s).strip()            # collapse whitespace
    return s.casefold()                            # stronger than lower()

def canon_country_key(s) -> str:
    x = norm_text(s)
    return COUNTRY_ALIASES_TO_CANON.get(x, x)

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
    """Heuristic check: does company name plausibly match domain?"""
    c_raw = company if isinstance(company, str) else ""
    d_raw = domain if isinstance(domain, str) else ""

    c = _normalize_tokens(c_raw)
    d = _clean_domain(d_raw)
    if not c or not d:
        return "Unsure – Please Check", 0, "missing values"

    d_base = d.split(".")[0] if "." in d else d
    d_base = re.sub(r"[^a-zA-Z0-9]", "", d_base)

    # Strong containment checks
    c_compact = re.sub(r"[^a-zA-Z0-9]", "", c)
    if c_compact and c_compact in d_base:
        return "Likely Match", 95, "company contained in domain"
    if d_base and d_base in c_compact:
        return "Likely Match", 90, "domain contained in company"

    # Token containment
    for token in c.split():
        if len(token) >= 4 and token in d_base:
            score = fuzz.partial_ratio(c, d_base)
            if score >= 70:
                return "Likely Match", score, "token containment"

    # Brand-ish suffix logic (incl. group/holdings etc.)
    BRAND_TERMS = {"tx","bio","pharma","therapeutics","labs","health","med","rx","group","holdings"}
    if any(t in c.split() for t in BRAND_TERMS) and any(t in d for t in BRAND_TERMS):
        sc = fuzz.partial_ratio(c, d_base)
        if sc >= 70:
            return "Likely Match", max(80, sc), "brand term overlap"

    score_full = fuzz.token_sort_ratio(c, d_base)
    score_partial = fuzz.partial_ratio(c, d_base)
    score = int(max(score_full, score_partial))

    if score >= THRESHOLD_STRONG:
        return "Likely Match", score, "strong fuzzy"
    elif score >= THRESHOLD_WEAK:
        return "Unsure – Please Check", score, "weak fuzzy"
    else:
        return "Likely NOT Match", score, "low similarity"

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

def normalise_phone_for_check(raw_phone: str, region: str) -> str:
    """
    Clean number just for validation:
    - strips spaces/brackets/etc.
    - converts leading 00 to +
    - if number starts with the region's calling code but lacks '+', add '+'
    IMPORTANT: this does NOT modify the stored/original value in the output.
    """
    phone = "" if raw_phone is None else str(raw_phone)
    phone = phone.strip()
    if not phone:
        return ""
    phone = re.sub(r"[^\d+]", "", phone)          # keep digits and leading +
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

def phone_country_check(raw_phone, country_standardised_label) -> tuple[str, str]:
    """
    Returns (Status, Reason)
    Status: Match | Warning | Unsure
    """
    if raw_phone is None or (isinstance(raw_phone, float) and pd.isna(raw_phone)) or str(raw_phone).strip() == "":
        return "Unsure", "Missing phone"
    if country_standardised_label is None or (isinstance(country_standardised_label, float) and pd.isna(country_standardised_label)) or str(country_standardised_label).strip() == "":
        return "Unsure", "Missing country"

    canon = canon_country_key(country_standardised_label)
    regions = COUNTRY_TO_REGIONS.get(canon, [])
    if not regions:
        return "Unsure", f"No phone region mapping for '{country_standardised_label}'"

    # Try each possible region for that country (usually 1)
    for region in regions:
        phone_for_check = normalise_phone_for_check(raw_phone, region)
        try:
            num = phonenumbers.parse(phone_for_check, region)

            # If the number includes an international code, ensure region matches expectation where possible
            actual_region = phonenumbers.region_code_for_number(num) or ""
            if actual_region and actual_region not in regions:
                return "Warning", f"Parsed region {actual_region} does not match expected {regions}"

            if not phonenumbers.is_possible_number(num):
                return "Warning", "Number not possible for country"
            if not phonenumbers.is_valid_number(num):
                return "Warning", "Number not valid for country"

            inferred = ""
            s = str(raw_phone).strip()
            if (s.startswith("00") or (re.sub(r"[^\d]", "", s).startswith(str(country_code_for_region(region))) and not s.strip().startswith("+"))):
                inferred = " (international prefix inferred)"
            return "Match", "Valid for country" + inferred

        except Exception:
            continue

    return "Unsure", "Could not parse phone for supplied country"

# ------------------ main processing ------------------

def run_matching(master_bytes: bytes, picklist_bytes: bytes, highlight_changes: bool = True, progress_cb=None) -> bytes:
    df_master = pd.read_excel(io.BytesIO(master_bytes))
    df_picklist = pd.read_excel(io.BytesIO(picklist_bytes))

    # EXACT column pairs: (master, picklist)
    EXACT_PAIRS = [
        ("companyname","companyname"),
        ("c_country","c_country"),
        ("c_state","c_state"),
        ("lead_country","lead_country"),
        ("departments","departments"),
        ("c_industry","c_industry"),
        ("asset_title","asset_title"),
    ]

    df_out = df_master.copy()
    corrected_cells = set()

    # Build picklist canonical country -> exact picklist label (for target formatting)
    picklist_country_label_by_canon = {}
    for col in ("lead_country", "c_country"):
        if col in df_picklist.columns:
            for v in df_picklist[col].dropna().astype(str):
                v_str = v.strip()
                if not v_str:
                    continue
                picklist_country_label_by_canon[canon_country_key(v_str)] = v_str

    if progress_cb:
        progress_cb(0.35, text="Matching Master ↔ Picklist (exact checks)...")

    for master_col, picklist_col in EXACT_PAIRS:
        out_col = f"Match_{master_col}"
        if master_col in df_master.columns and picklist_col in df_picklist.columns:
            pick_map = {norm_text(v): str(v).strip() for v in df_picklist[picklist_col].dropna().astype(str)}
            matches, new_vals = [], []

            for i, val in enumerate(df_master[master_col].fillna("").astype(str)):
                raw_val = val
                key = norm_text(raw_val)

                # Country canonicalisation -> picklist spelling
                if master_col.strip().lower() in {"lead_country", "c_country", "country"}:
                    canon = canon_country_key(raw_val)
                    if canon in picklist_country_label_by_canon:
                        desired_label = picklist_country_label_by_canon[canon]
                        key = norm_text(desired_label)

                if key in pick_map:
                    matches.append("Yes")
                    new_val = pick_map[key]

                    if master_col.strip().lower() in {"lead_country","c_country","country"}:
                        canon = canon_country_key(raw_val)
                        new_val = picklist_country_label_by_canon.get(canon, new_val)

                    new_vals.append(new_val)
                    if str(new_val).strip() != str(raw_val).strip():
                        corrected_cells.add((master_col, i + 2))  # +2 because header is row 1
                else:
                    matches.append("No")
                    new_vals.append(raw_val)

            df_out[out_col] = matches
            df_out[master_col] = new_vals
        else:
            df_out[out_col] = "Column Missing"

    # Audit columns: standardised country + change note
    if "lead_country" in df_out.columns:
        std_vals, notes = [], []
        for v in df_out["lead_country"].fillna("").astype(str):
            canon = canon_country_key(v)
            std = picklist_country_label_by_canon.get(canon, str(v).strip())
            std_vals.append(std)
            notes.append(f"{v} → {std}" if norm_text(std) and norm_text(v) and norm_text(std) != norm_text(v) else "")
        df_out["Country_Standardised"] = std_vals
        df_out["Country_Change_Note"] = notes
    elif "c_country" in df_out.columns:
        std_vals, notes = [], []
        for v in df_out["c_country"].fillna("").astype(str):
            canon = canon_country_key(v)
            std = picklist_country_label_by_canon.get(canon, str(v).strip())
            std_vals.append(std)
            notes.append(f"{v} → {std}" if norm_text(std) and norm_text(v) and norm_text(std) != norm_text(v) else "")
        df_out["Country_Standardised"] = std_vals
        df_out["Country_Change_Note"] = notes

    # Seniority parse
    if progress_cb:
        progress_cb(0.55, text="Parsing seniority...")

    if "jobtitle" in df_master.columns:
        parsed = df_master["jobtitle"].apply(parse_seniority)
        df_out["Parsed_Seniority"] = parsed.apply(lambda x: x[0])
        df_out["Seniority_Logic"] = parsed.apply(lambda x: x[1])
    else:
        df_out["Parsed_Seniority"] = None
        df_out["Seniority_Logic"] = "jobtitle column not found"

    # Company ↔ domain validation
    if progress_cb:
        progress_cb(0.70, text="Validating company ↔ domain...")

    company_cols = [c for c in df_master.columns if c.strip().lower() in {"companyname","company","company name","company_name"}]
    domain_cols  = [c for c in df_master.columns if c.strip().lower() in {"website","domain","email domain","email_domain","company_domain","company domain"}]
    email_cols   = [c for c in df_master.columns if "email" in c.lower()]

    if company_cols:
        company_col = company_cols[0]
        domain_col  = domain_cols[0] if domain_cols else None
        email_col   = email_cols[0] if email_cols else None

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

    # Phone ↔ country validation (does NOT modify phone value)
    if progress_cb:
        progress_cb(0.82, text="Validating phone ↔ country...")

    country_for_phone_col = (
        "Country_Standardised" if "Country_Standardised" in df_out.columns
        else ("lead_country" if "lead_country" in df_out.columns
              else ("c_country" if "c_country" in df_out.columns else None))
    )

    phone_candidates = []
    for c in df_master.columns:
        cl = c.strip().lower()
        if cl in {"phone","phone_main","phonemain","phone main","phone number","phonenumber",
                  "mobile","mobilephone","phone_mobile","phonemobile","phone mobile"}:
            phone_candidates.append(c)

    if country_for_phone_col and phone_candidates:
        for phone_col in phone_candidates:
            st_col = f"{phone_col}_PhoneCountry_Status"
            rs_col = f"{phone_col}_PhoneCountry_Reason"
            out_status, out_reason = [], []

            for i in range(len(df_master)):
                raw_phone = df_master.at[i, phone_col]
                country_label = df_out.at[i, country_for_phone_col]
                s, r = phone_country_check(raw_phone, country_label)
                out_status.append(s)
                out_reason.append(r)

            df_out[st_col] = out_status
            df_out[rs_col] = out_reason
    else:
        df_out["PhoneCountry_Status"] = "phone or country column not found"
        df_out["PhoneCountry_Reason"] = ""

    # Write to Excel
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df_out.to_excel(writer, index=False, sheet_name="Results")

    output.seek(0)

    # Highlight corrected cells (only where we overwrote to picklist label)
    if highlight_changes and corrected_cells:
        wb = load_workbook(output)
        ws = wb["Results"]

        header = [cell.value for cell in ws[1]]
        col_index = {str(v): i+1 for i, v in enumerate(header)}

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
    "Upload the Lead Master + Picklist. The tool checks exact picklist matches, "
    "standardises countries to picklist format, validates company↔domain, and validates "
    "phone↔country (without changing phone values)."
)

col1, col2 = st.columns(2)
with col1:
    master_file = st.file_uploader("Upload Lead Master (.xlsx)", type=["xlsx"], key="master")
with col2:
    picklist_file = st.file_uploader("Upload Picklist (.xlsx)", type=["xlsx"], key="picklist")

highlight = st.toggle("Highlight corrected cells (yellow)", value=True)

if master_file and picklist_file:
    progress = st.progress(0.0, text="Ready")

    def prog(p, text=""):
        progress.progress(min(max(float(p), 0.0), 1.0), text=text)

    try:
        output_bytes = run_matching(
            master_file.read(),
            picklist_file.read(),
            highlight_changes=highlight,
            progress_cb=prog
        )
        st.success("Processing complete.")
        st.download_button(
            label="Download Processed File",
            data=output_bytes,
            file_name="Full_Check_Results.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )
    except Exception as e:
        st.error(f"Error: {e}")
