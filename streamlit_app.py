import io
import os
import re
import tempfile
import pandas as pd
from rapidfuzz import fuzz
from openpyxl import load_workbook
from openpyxl.styles import PatternFill
import streamlit as st

# Optional dependency for phone validation (added in requirements.txt).
try:
    import phonenumbers
except Exception:  # pragma: no cover
    phonenumbers = None


# ------------------ constants ------------------

SUFFIXES = {
    "ltd","limited","co","company","corp","corporation","inc","incorporated",
    "plc","public","llc","lp","llp","ulc","pc","pllc","sa","ag","nv","se","bv",
    "oy","ab","aps","as","kft","zrt","rt","sarl","sas","spa","gmbh","ug","bvba",
    "cvba","nvsa","pte","pty","bhd","sdn","kabushiki","kaisha","kk","godo","dk",
    "dmcc","pjsc","psc","jsc","ltda","srl","s.r.l","group","holdings","limitedpartnership"
}

COUNTRY_EQUIVALENTS = {
    "uk":"united kingdom","u.k.":"united kingdom","england":"united kingdom",
    "great britain":"united kingdom","britain":"united kingdom",
    "usa":"united states","u.s.a.":"united states","us":"united states",
    "america":"united states","united states of america":"united states",
    "uae":"united arab emirates","u.a.e.":"united arab emirates",
    "south korea":"republic of korea","korea":"republic of korea",
    "north korea":"democratic people's republic of korea","russia":"russian federation",
    "czechia":"czech republic","côte d’ivoire":"ivory coast","cote d'ivoire":"ivory coast",
    "iran":"islamic republic of iran","venezuela":"bolivarian republic of venezuela",
    "taiwan":"republic of china","hong kong sar":"hong kong","macao sar":"macau","prc":"china"
}

THRESHOLD = 70

# Minimal country → ISO region mapping for phone validation.
# (We normalize using COUNTRY_EQUIVALENTS first, then map to a 2-letter region when possible.)
COUNTRY_TO_REGION2 = {
    "ireland": "IE",
    "united kingdom": "GB",
    "great britain": "GB",
    "united states": "US",
    "canada": "CA",
    "australia": "AU",
    "new zealand": "NZ",
    "singapore": "SG",
    "malaysia": "MY",
    "thailand": "TH",
    "indonesia": "ID",
    "philippines": "PH",
    "hong kong": "HK",
    "taiwan": "TW",
    "japan": "JP",
    "korea": "KR",
    "republic of korea": "KR",
    "china": "CN",
    "india": "IN",
    "vietnam": "VN",
    "france": "FR",
    "germany": "DE",
    "spain": "ES",
    "italy": "IT",
    "netherlands": "NL",
    "belgium": "BE",
    "sweden": "SE",
    "norway": "NO",
    "denmark": "DK",
    "finland": "FI",
    "switzerland": "CH",
    "austria": "AT",
    "portugal": "PT",
    "poland": "PL",
    "czech republic": "CZ",
    "united arab emirates": "AE",
    "saudi arabia": "SA",
}

def _normalize_country_name(country: str) -> str:
    if not isinstance(country, str):
        return ""
    c = country.strip().lower()
    return COUNTRY_EQUIVALENTS.get(c, c)

def _validate_phone_vs_country(phone: str, country: str):
    """Return (status, reason, e164) where status in {Match, Mismatch, Unsure}."""
    if not isinstance(phone, str) or not phone.strip():
        return "Unsure", "missing phone", ""
    c_norm = _normalize_country_name(country)
    region = COUNTRY_TO_REGION2.get(c_norm, "")
    raw = phone.strip()

    # If phonenumbers isn't installed, do a light-touch heuristic only.
    if phonenumbers is None:
        if raw.startswith("+") and region:
            return "Unsure", "phonenumbers not installed (cannot validate)", raw
        return "Unsure", "phonenumbers not installed (cannot validate)", raw

    try:
        # Parse with region if not explicitly international.
        parsed = phonenumbers.parse(raw, region or None)
        if not phonenumbers.is_possible_number(parsed):
            return "Mismatch", "number not possible", ""
        if not phonenumbers.is_valid_number(parsed):
            return "Mismatch", "number not valid", ""
        e164 = phonenumbers.format_number(parsed, phonenumbers.PhoneNumberFormat.E164)

        # If we don't know the expected region, we can still output the parsed region.
        parsed_region = (phonenumbers.region_code_for_number(parsed) or "").upper()

        if not region:
            return "Unsure", f"unknown country mapping for '{country}' (parsed region {parsed_region})", e164

        if parsed_region == region:
            return "Match", f"matches country {region}", e164

        return "Mismatch", f"country {region} but phone parses to {parsed_region}", e164
    except Exception as e:
        return "Unsure", f"parse error: {e}", ""

def _best_picklist_match(value: str, pick_values, min_score: int = 90):
    """Return (matched_value, score) for the best fuzzy match to pick_values."""
    if not isinstance(value, str) or not value.strip():
        return "", 0
    v_norm = _normalize_tokens(value)
    best_val, best_score = "", 0
    for p in pick_values:
        p_str = str(p).strip()
        if not p_str:
            continue
        p_norm = _normalize_tokens(p_str)
        s = fuzz.token_sort_ratio(v_norm, p_norm)
        if s > best_score:
            best_score = s
            best_val = p_str
    return (best_val, best_score) if best_score >= min_score else ("", best_score)


# ------------------ helpers ------------------

def _normalize_tokens(text: str) -> str:
    if not isinstance(text, str):
        return ""
    text = re.sub(r"[^a-zA-Z0-9\s]", " ", text.lower())
    parts = [w for w in text.split() if w not in SUFFIXES]
    return " ".join(parts).strip()

def _clean_domain(domain: str) -> str:
    if not isinstance(domain, str):
        return ""
    domain = domain.lower().strip()
    domain = re.sub(r"^https?://", "", domain)
    domain = re.sub(r"/.*$", "", domain)
    domain = re.sub(r"^www\.", "", domain)
    parts = domain.split(".")
    return parts[-2] if len(parts) >= 2 else domain

def _extract_domain_from_email(email: str) -> str:
    if not isinstance(email, str) or "@" not in email:
        return ""
    domain = email.split("@")[-1].lower().strip()
    domain = re.sub(r"^www\.", "", domain)
    domain = re.sub(r"/.*$", "", domain)
    return domain

def compare_company_domain(company: str, domain: str):
    if not isinstance(company, str) or not isinstance(domain, str):
        return "Unsure – Please Check", 0, "missing input"

    c = _normalize_tokens(company)
    d_raw = domain.lower().strip()
    d = _clean_domain(d_raw)

    if d in c.replace(" ", "") or c.replace(" ", "") in d:
        return "Likely Match", 100, "direct containment"

    if any(word in c for word in d.split()) or any(word in d for word in c.split()):
        score = fuzz.partial_ratio(c, d)
        if score >= 70:
            return "Likely Match", score, "token containment"

    BRAND_TERMS = {"tx","bio","pharma","therapeutics","labs","health","med","rx","group","holdings"}
    if any(t in c.split() for t in BRAND_TERMS) and any(t in d for t in BRAND_TERMS):
        if fuzz.partial_ratio(c, d) >= 70:
            return "Likely Match", 90, "brand suffix match"

    score_full = fuzz.token_sort_ratio(c, d)
    score_partial = fuzz.partial_ratio(c, d)
    score = max(score_full, score_partial)

    if score >= 85:
        return "Likely Match", score, "strong fuzzy"
    elif score >= THRESHOLD:
        return "Unsure – Please Check", score, "weak fuzzy"
    else:
        return "Likely NOT Match", score, "low similarity"

def parse_seniority(title):
    if not isinstance(title, str): return "Entry", "no title"
    t = title.lower().strip()
    if re.search(r"\bchief\b|\bcio\b|\bcto\b|\bceo\b|\bcfo\b|\bciso\b|\bcpo\b|\bcso\b|\bcoo\b|\bchro\b|\bpresident\b", t): return "C Suite", "c-level"
    if re.search(r"\bvice president\b|\bvp\b|\bsvp\b", t): return "VP", "vp"
    if re.search(r"\bhead\b", t): return "Head", "head"
    if re.search(r"\bdirector\b", t): return "Director", "director"
    if re.search(r"\bmanager\b|\bmgr\b", t): return "Manager", "manager"
    if re.search(r"\bsenior\b|\bsr\b|\blead\b|\bprincipal\b", t): return "Senior", "senior"
    if re.search(r"\bintern\b|\btrainee\b|\bassistant\b|\bgraduate\b", t): return "Entry", "entry"
    return "Entry", "none"

# ------------------ main matching function ------------------

def run_matching(master_bytes: bytes, picklist_bytes: bytes, highlight_changes=True, progress_cb=None):
    # read Excel from in-memory bytes
    df_master = pd.read_excel(io.BytesIO(master_bytes))
    df_picklist = pd.read_excel(io.BytesIO(picklist_bytes))

    if progress_cb: progress_cb(0.2, text="Preparing data...")

    EXACT_PAIRS = [
        ("c_industry","c_industry"),
        ("asset_title","asset_title"),
        ("lead_country","lead_country"),
        ("departments","departments"),
        ("c_state","c_state")
    ]
    df_out = df_master.copy()
    corrected_cells = set()

    if progress_cb: progress_cb(0.4, text="Matching Master ↔ Picklist...")

    for master_col, picklist_col in EXACT_PAIRS:
        out_col = f"Match_{master_col}"
        if master_col in df_master.columns and picklist_col in df_picklist.columns:
            pick_map = {v.strip().lower(): v.strip() for v in df_picklist[picklist_col].dropna().astype(str)}
            matches, new_vals = [], []
            for i, val in enumerate(df_master[master_col].fillna("").astype(str)):
                val_norm = val.strip().lower()
                val_norm_eq = COUNTRY_EQUIVALENTS.get(val_norm, val_norm) if master_col.lower() in ["lead_country","country","c_country"] else val_norm
                if val_norm_eq in pick_map:
                    matches.append("Yes")
                    new_val = pick_map[val_norm_eq]
                    new_vals.append(new_val)
                    if new_val != val:
                        corrected_cells.add((master_col, i + 2))
                else:
                    matches.append("No")
                    new_vals.append(val)
            df_out[out_col] = matches
            df_out[master_col] = new_vals
        else:
            df_out[out_col] = "Column Missing"

    # dynamic question columns (q1, q2, question 3, etc.)
    q_cols = [c for c in df_picklist.columns if re.match(r"(?i)q0*\d+|question\s*\d+", c)]
    for qc in q_cols:
        out_col = f"Match_{qc}"
        if qc in df_master.columns:
            valid_answers = set(df_picklist[qc].dropna().astype(str).str.strip().str.lower())
            matches = []
            for val in df_master[qc].fillna("").astype(str):
                val_norm = val.strip().lower()
                if val_norm in valid_answers:
                    matches.append("Yes")
                elif val_norm == "":
                    matches.append("Blank")
                else:
                    matches.append("No")
            df_out[out_col] = matches
        else:
            df_out[out_col] = "Column Missing"

    # seniority parse
    if "jobtitle" in df_master.columns:
        parsed = df_master["jobtitle"].apply(parse_seniority)
        df_out["Parsed_Seniority"] = parsed.apply(lambda x: x[0])
        df_out["Seniority_Logic"] = parsed.apply(lambda x: x[1])
    else:
        df_out["Parsed_Seniority"] = None
        df_out["Seniority_Logic"] = "jobtitle column not found"

    # company ↔ domain validation
    if progress_cb: progress_cb(0.6, text="Validating company ↔ domain...")


    # --- Company Name mapping (fuzzy) ---
    if "companyname" in df_master.columns and "companyname" in df_picklist.columns:
        pick_companies = df_picklist["companyname"].dropna().astype(str).tolist()
        comp_status, comp_match, comp_score = [], [], []
        for i, val in enumerate(df_master["companyname"].fillna("").astype(str)):
            if not val.strip():
                comp_status.append("Missing"); comp_match.append(""); comp_score.append(0)
                continue

            # First try normalized exact match (handles "Group" / suffixes).
            v_norm = _normalize_tokens(val)
            found = ""
            for p in pick_companies:
                if _normalize_tokens(p) == v_norm:
                    found = str(p).strip()
                    break

            if found:
                comp_status.append("Yes"); comp_match.append(found); comp_score.append(100)
                if found != val:
                    df_out.at[i, "companyname"] = found
                    corrected_cells.add(("companyname", i + 2))
            else:
                best, score = _best_picklist_match(val, pick_companies, min_score=90)
                if best:
                    comp_status.append("Yes"); comp_match.append(best); comp_score.append(score)
                    if best != val:
                        df_out.at[i, "companyname"] = best
                        corrected_cells.add(("companyname", i + 2))
                else:
                    comp_status.append("No"); comp_match.append(""); comp_score.append(score)

        df_out["CompanyName_Match_Status"] = comp_status
        df_out["CompanyName_Matched_Value"] = comp_match
        df_out["CompanyName_Match_Score"] = comp_score
    else:
        df_out["CompanyName_Match_Status"] = "Column Missing"

    # --- Phone vs Country validation ---
    phone_cols = [c for c in df_master.columns if c.strip().lower() in ["telephone","phone","phone_number","phone number","mobile","mobilephone"]]
    country_cols = [c for c in df_master.columns if c.strip().lower() in ["lead_country","country","c_country"]]
    if phone_cols and country_cols:
        phone_col = phone_cols[0]
        country_col = country_cols[0]
        p_status, p_reason, p_e164 = [], [], []
        for i in range(len(df_master)):
            phone = str(df_master.at[i, phone_col]) if pd.notna(df_master.at[i, phone_col]) else ""
            country = str(df_master.at[i, country_col]) if pd.notna(df_master.at[i, country_col]) else ""
            status, reason, e164 = _validate_phone_vs_country(phone, country)
            p_status.append(status); p_reason.append(reason); p_e164.append(e164)
        df_out["Phone_Country_Status"] = p_status
        df_out["Phone_Country_Reason"] = p_reason
        df_out["Phone_E164"] = p_e164
    else:
        df_out["Phone_Country_Status"] = "No phone/country columns found"

    # --- Required field completeness check (easy to extend) ---
    default_required = ["email","companyname","firstname","lastname","jobtitle","lead_country"]
    present_required = [c for c in default_required if c in df_master.columns]
    if present_required:
        missing_list = []
        for i in range(len(df_master)):
            missing = [c for c in present_required if not str(df_master.at[i, c]).strip() or pd.isna(df_master.at[i, c])]
            missing_list.append(", ".join(missing))
        df_out["Missing_Required_Fields"] = missing_list

    company_cols = [c for c in df_master.columns if c.strip().lower() in ["companyname","company","company name","company_name"]]
    domain_cols  = [c for c in df_master.columns if c.strip().lower() in ["website","domain","email domain","email_domain"]]
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
                dom = _extract_domain_from_email(df_master.at[i, email_col])
            status, score, reason = compare_company_domain(comp, dom)
            statuses.append(status); scores.append(score); reasons.append(reason)

        df_out["Domain_Check_Status"] = statuses
        df_out["Domain_Check_Score"]  = scores
        df_out["Domain_Check_Reason"] = reasons
    else:
        df_out["Domain_Check_Status"] = "No company/domain columns found"
        df_out["Domain_Check_Score"]  = None
        df_out["Domain_Check_Reason"] = None

    # save to temp file and apply formatting with openpyxl
    if progress_cb: progress_cb(0.85, text="Saving results...")

    with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as tmp:
        out_path = tmp.name
    df_out.to_excel(out_path, index=False)

    wb = load_workbook(out_path)
    ws = wb.active
    yellow = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
    green  = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
    red    = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
    blue   = PatternFill(start_color="ADD8E6", end_color="ADD8E6", fill_type="solid")

    # color Match_* columns
    for col_idx, col in enumerate(df_out.columns, start=1):
        if str(col).startswith("Match_"):
            for row in range(2, ws.max_row + 1):
                val = str(ws.cell(row=row, column=col_idx).value).strip().lower()
                if val == "yes":
                    ws.cell(row=row, column=col_idx).fill = green
                elif val == "no":
                    ws.cell(row=row, column=col_idx).fill = red
                else:
                    ws.cell(row=row, column=col_idx).fill = yellow

    # highlight changed cells
    for col_name, row in corrected_cells:
        if col_name in df_out.columns:
            col_idx = list(df_out.columns).index(col_name) + 1
            ws.cell(row=row, column=col_idx).fill = blue

    wb.save(out_path)

    # read back as bytes for download
    with open(out_path, "rb") as f:
        out_bytes = f.read()
    os.remove(out_path)

    if progress_cb: progress_cb(1.0, text="Done!")

    return out_bytes

# ------------------ UI ------------------

st.set_page_config(page_title="Master–Picklist + Domain Matching Tool", layout="wide")
st.title("📊 Master–Picklist + Domain Matching Tool")
st.write("Upload MASTER & PICKLIST Excel files to auto-match, validate domains, map questions, and optionally highlight changed values.")

col1, col2 = st.columns(2)
with col1:
    master_file = st.file_uploader("Upload MASTER Excel file (.xlsx)", type=["xlsx"], key="master")
with col2:
    picklist_file = st.file_uploader("Upload PICKLIST Excel file (.xlsx)", type=["xlsx"], key="picklist")

highlight = st.checkbox("Highlight changed values (blue)", value=True)

run = st.button("Run Matching")

if run:
    if not master_file or not picklist_file:
        st.error("Please upload both MASTER and PICKLIST files.")
    else:
        progress = st.progress(0.0, text="Starting...")
        def prog(p, text=""):
            progress.progress(min(max(p, 0.0), 1.0), text=text)

        try:
            output_bytes = run_matching(master_file.read(), picklist_file.read(), highlight_changes=highlight, progress_cb=prog)
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
