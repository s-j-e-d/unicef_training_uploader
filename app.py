from __future__ import annotations

import json
import re
import uuid
from typing import Any, Optional, Tuple, Dict, List
import os
import io

import pandas as pd
import streamlit as st
import requests
import warnings

from openpyxl import load_workbook
from openpyxl.utils.cell import range_boundaries

# (optional) silence harmless openpyxl data-validation warnings
warnings.filterwarnings("ignore", message="Data Validation extension is not supported and will be removed")

# -------------------------------
# HARD-CODED DESTINATION & TOKEN
# -------------------------------
KOBO_TOKEN = os.environ.get("KOBO_TOKEN") or st.secrets.get("KOBO_TOKEN", "")
if not KOBO_TOKEN:
    st.error("Missing KOBO_TOKEN. Add it in Streamlit Cloud → App → Settings → Secrets.")
    st.stop()

ASSET_UID = "aS5UNu8wuynhJN87JYopFQ"                   # exact form asset uid (matches your new XLSForm)
KC_BASE = "https://kc-eu.kobotoolbox.org"             # EU KoboCAT
SUBMIT_URL = f"{KC_BASE}/api/v1/submissions.json"     # JSON submissions

FORMS = {
    "Training Attendance (EU)": {
        "asset_uid": ASSET_UID,
        "landing": "https://eu.kobotoolbox.org/#/forms/aS5UNu8wuynhJN87JYopFQ/landing",
    }
}

# ---------------------------------
# Expected headers in the Excel table (TblAttend)
# ---------------------------------
L = {
    # Personal
    "first_name": "First Name",
    "second_name": "Second Name",
    "phone": "Phone Number",
    "email": "Email",
    "gender": "Gender",
    "works_hf": "Are you working in Health Facility (Yes or No) / \nМісце роботи - Заклад Охорони Здоров'я (Так чи ні)",
    "position": "Position / \nПосада",

    # Location (Parent PCHF / Ambulatory)
    "oblast_parent": "Oblast of Parent PCHF",
    "rayon_parent": "Rayon of Parent PCHF",
    "hromada_parent": "Hromada of Parent PCHF",
    "parent_pchf_name": "Name of Parent PCHF",
    "ambulatory_name": "Name of Ambulatory",

    # Optional helper & other provider
    "helper": "Helper",
    "service_provider_other": "Other Service Provider or Facility / Інше Місце надання послуг чи Заклад",

    # Training meta (same each row)
    "score": "Post-Training Test Score %",
    "start_date_lbl": "Start date (dd-mm-yyyy)",  # if your sheet changed text, alias below will handle
    "end_date_lbl": "End date (dd-mm-yyyy)",
    "training_name": "Name of Training",
    "place": "Place where the training is conducted",
}

# -------------------------------
# Kobo keys (must match the XLSForm schema)
# -------------------------------
K = {
    "first_name": "personal_info/first_name",
    "second_name": "personal_info/second_name",
    "phone": "personal_info/phone",
    "email": "personal_info/email",
    "gender": "personal_info/gender",
    "works_hf": "personal_info/works_in_hf",

    "position": "work_location/position",
    "oblast": "work_location/oblast",
    "rayon": "work_location/rayon",
    "hromada": "work_location/hromada",
    "settlement": "work_location/settlement",
    "facility_code": "work_location/facility_code",          # PCHF (Parent)
    "service_provider": "work_location/service_provider",    # Ambulatory (name)
    "helper": "work_location/helper",
    "service_provider_other": "work_location/service_provider_other",

    "score": "training_score/training_score_pct",
    "start_date": "metadata/start_date",
    "end_date": "metadata/end_date",
    "training_name": "metadata/training_name",
    "place": "metadata/place",
}

# ---------------------------------
# Helpers: normalization & parsing
# ---------------------------------
def parse_name(full_name: Optional[str]) -> Tuple[Optional[str], Optional[str]]:
    if full_name is None:
        return None, None
    s = str(full_name).strip()
    if not s or s.lower() == "nan":
        return None, None
    parts = s.split()
    return (parts[0], " ".join(parts[1:])) if len(parts) > 1 else (parts[0], "")

def normalize_yes_no(v: Any) -> Optional[str]:
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return None
    s = str(v).strip().lower()
    if s in {"yes", "y", "true", "1", "так"}:
        return "yes"
    if s in {"no", "n", "false", "0", "ні"}:
        return "no"
    return None

def normalize_gender(v: Any) -> Optional[str]:
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return None
    s = str(v).strip().lower()
    if s.startswith("f"): return "female"
    if s.startswith("m"): return "male"
    if "other" in s: return "other"
    if "prefer" in s: return "prefer_not_to_say"
    return None

def parse_training_date(v: Any) -> Optional[str]:
    """
    Return ISO 'YYYY-MM-DD' for:
      - dd-mm-yyyy / dd/mm/yyyy
      - mm/yyyy   -> yyyy-mm-01
      - native Excel dates / pandas timestamps
    """
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return None

    # pandas can parse a lot of things, try first
    ts = pd.to_datetime(v, errors="coerce", dayfirst=True)
    if pd.notna(ts):
        return ts.strftime("%Y-%m-%d")

    # explicit regex paths
    s = str(v).strip()
    m = re.match(r"^(\d{1,2})[/-](\d{1,2})[/-](\d{4})$", s)  # dd-mm-yyyy
    if m:
        dd, mm, yyyy = int(m.group(1)), int(m.group(2)), int(m.group(3))
        try:
            return pd.Timestamp(year=yyyy, month=mm, day=dd).strftime("%Y-%m-%d")
        except Exception:
            return None

    m2 = re.match(r"^(\d{1,2})\s*[-/.]\s*(\d{4})$", s)       # mm/yyyy
    if m2:
        mm, yyyy = int(m2.group(1)), int(m2.group(2))
        try:
            return pd.Timestamp(year=yyyy, month=mm, day=1).strftime("%Y-%m-%d")
        except Exception:
            return None

    return None

def parse_score_int(v: Any) -> Optional[int]:
    """Accept 0..100 as whole numbers (also tolerates '82%' or 82.0)."""
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return None
    s = str(v).strip().replace("%", "")
    try:
        f = float(s)
    except Exception:
        return None
    n = int(round(f))
    return n if 0 <= n <= 100 else None

def prune(d: Dict[str, Any]) -> Dict[str, Any]:
    """Recursively remove empty/None/NaN/'' values but keep non-empty groups."""
    out: Dict[str, Any] = {}
    for k, v in d.items():
        if isinstance(v, dict):
            pruned = prune(v)
            if pruned:
                out[k] = pruned
        else:
            if v is None:
                continue
            if isinstance(v, float) and pd.isna(v):
                continue
            if isinstance(v, str) and v.strip() == "":
                continue
            out[k] = v
    return out

def post_submission_json(asset_uid: str, submission: Dict[str, Any]) -> requests.Response:
    headers = {
        "Authorization": f"Token {KOBO_TOKEN}",
        "Accept": "application/json",
        "Content-Type": "application/json",
    }
    body = {"id": asset_uid, "submission": submission}
    return requests.post(SUBMIT_URL, headers=headers, data=json.dumps(body), timeout=60)

# ---------------------------------
# Excel loader (prefer named table TblAttend on 'Attend')
# ---------------------------------
def load_attendance_df(file_like, sheet_fallback: str = "Attend", table_name: str = "TblAttend") -> pd.DataFrame:
    """Read the named table TblAttend if present; otherwise fallback to header-scan."""
    raw_bytes = file_like.read()

    # 1) Try openpyxl to locate the table
    try:
        wb = load_workbook(io.BytesIO(raw_bytes), data_only=True)
        if sheet_fallback in wb.sheetnames and table_name in wb[sheet_fallback].tables:
            ws = wb[sheet_fallback]
            ref = ws.tables[table_name].ref  # e.g., "A9:S999"
            min_col, min_row, max_col, max_row = range_boundaries(ref)
            data = []
            for row in ws.iter_rows(min_row=min_row, max_row=max_row, min_col=min_col, max_col=max_col, values_only=True):
                data.append(list(row))
            if not data:
                raise ValueError("TblAttend has no rows")
            headers = [str(h).strip() if h is not None else "" for h in data[0]]
            rows = data[1:]
            df = pd.DataFrame(rows, columns=headers)
            return df.dropna(how="all")
    except Exception:
        pass

    # 2) Fallback to old detection in the provided sheet
    buf = io.BytesIO(raw_bytes)
    raw = pd.read_excel(buf, sheet_name=sheet_fallback, header=None)
    header_idx: Optional[int] = None
    probe = [L["first_name"], L["phone"], L["email"]]
    for i in range(min(80, len(raw))):
        vals = [str(x).strip() if not pd.isna(x) else "" for x in list(raw.iloc[i, :].values)]
        if all(h in vals for h in probe):
            header_idx = i
            break
    if header_idx is None:
        raise ValueError("Could not locate the header row. Please confirm the sheet structure.")
    buf2 = io.BytesIO(raw_bytes)
    df = pd.read_excel(buf2, sheet_name=sheet_fallback, header=header_idx).dropna(how="all")
    return df

# ---------------------------------
# Build one submission (NESTED to match XLSForm groups/fields)
# ---------------------------------
def build_submission_nested(row: pd.Series) -> Dict[str, Any]:
    # Score (0..100 integer)
    score_val = parse_score_int(row.get(L["score"]))

    submission = {
        "metadata": {
            "training_name": None if pd.isna(row.get(L["training_name"])) else str(row.get(L["training_name"])).strip(),
            "start_date": parse_training_date(row.get(L["start_date_lbl"])),
            "end_date": parse_training_date(row.get(L["end_date_lbl"])),
            "place": None if pd.isna(row.get(L["place"])) else str(row.get(L["place"])).strip(),
        },
        "personal_info": {
            "first_name": None if pd.isna(row.get(L["first_name"])) else str(row.get(L["first_name"])).strip(),
            "second_name": None if pd.isna(row.get(L["second_name"])) else str(row.get(L["second_name"])).strip(),
            "phone": None if pd.isna(row.get(L["phone"])) else str(row.get(L["phone"])).strip(),
            "email": None if pd.isna(row.get(L["email"])) else str(row.get(L["email"])).strip(),
            "gender": normalize_gender(row.get(L["gender"])),
            "works_in_hf": normalize_yes_no(row.get(L["works_hf"])),
        },
        "work_location": {
            "position": None if pd.isna(row.get(L["position"])) else str(row.get(L["position"])).strip(),
            "oblast": None if pd.isna(row.get(L["oblast_parent"])) else str(row.get(L["oblast_parent"])).strip(),
            "rayon": None if pd.isna(row.get(L["rayon_parent"])) else str(row.get(L["rayon_parent"])).strip(),
            "hromada": None if pd.isna(row.get(L["hromada_parent"])) else str(row.get(L["hromada_parent"])).strip(),
            "settlement": None,  # not in the new template
            "facility_code": None if pd.isna(row.get(L["parent_pchf_name"])) else str(row.get(L["parent_pchf_name"])).strip(),  # PCHF name/code
            "service_provider": None if pd.isna(row.get(L["ambulatory_name"])) else str(row.get(L["ambulatory_name"])).strip(), # Ambulatory name
            "helper": None if pd.isna(row.get(L["helper"])) else str(row.get(L["helper"])).strip(),
            "service_provider_other": None if pd.isna(row.get(L["service_provider_other"])) else str(row.get(L["service_provider_other"])).strip(),
        },
        "training_score": {
            "training_score_pct": score_val,
        },
        "meta": {"instanceID": f"uuid:{uuid.uuid4()}"},  # KoboCAT requires nested meta
    }
    return prune(submission)

# ---------------------------------
# Validation (required fields & formats)
# ---------------------------------
def is_blank(v: Any) -> bool:
    if v is None: return True
    if isinstance(v, float) and pd.isna(v): return True
    s = str(v).strip()
    return s == "" or s.lower() == "nan" or s == "NaT"

# Required: everything except helper + service_provider_other
REQUIRED_LABELS = {
    L["first_name"], L["second_name"], L["phone"], L["email"], L["gender"], L["works_hf"],
    L["position"], L["oblast_parent"], L["rayon_parent"], L["hromada_parent"],
    L["score"], L["start_date_lbl"], L["end_date_lbl"], L["training_name"], L["place"]
}

def validate_dataframe(df: pd.DataFrame) -> tuple[pd.DataFrame, int, int]:
    issues: List[Dict[str, Any]] = []
    for idx, row in df.iterrows():
        errors: List[str] = []
        warnings_: List[str] = []

        # Requiredness (direct)
        for lbl in REQUIRED_LABELS:
            if lbl in df.columns and is_blank(row.get(lbl)):
                errors.append(f"Missing: {lbl}")
            elif lbl not in df.columns:
                errors.append(f"Missing column: {lbl}")

        # At least one of PCHF or Ambulatory (facility fields)
        pchf = row.get(L["parent_pchf_name"]) if L["parent_pchf_name"] in df.columns else None
        amb = row.get(L["ambulatory_name"]) if L["ambulatory_name"] in df.columns else None
        if is_blank(pchf) and is_blank(amb):
            errors.append(f"Facility required: provide at least one of '{L['parent_pchf_name']}' or '{L['ambulatory_name']}'")

        # Score: integer 0..100
        sv = row.get(L["score"]) if L["score"] in df.columns else None
        if not is_blank(sv) and parse_score_int(sv) is None:
            errors.append("Score must be a whole number 0–100")

        # Dates: parseable
        for k in ("start_date_lbl", "end_date_lbl"):
            if k in L and L[k] in df.columns:
                v = row.get(L[k])
                if not is_blank(v) and parse_training_date(v) is None:
                    errors.append(f"{L[k]}: invalid date (use dd-mm-yyyy or a proper Excel date)")

        # Yes/No + Gender: normalization check (non-blocking)
        v = row.get(L["works_hf"]) if L["works_hf"] in df.columns else None
        if not is_blank(v) and normalize_yes_no(v) is None:
            warnings_.append("Works in HF should be Yes/No (Так/Ні)")

        v = row.get(L["gender"]) if L["gender"] in df.columns else None
        if not is_blank(v) and normalize_gender(v) is None:
            warnings_.append("Gender not recognized")

        if errors or warnings_:
            issues.append({
                "Row #": idx + 1,
                "Name": row.get(L["first_name"]) if L["first_name"] in df.columns else None,
                "Errors": "; ".join(errors),
                "Warnings": "; ".join(warnings_),
            })

    issues_df = pd.DataFrame(issues) if issues else pd.DataFrame(columns=["Row #", "Name", "Errors", "Warnings"])
    n_errors = sum(1 for i in issues if i["Errors"])
    n_warnings = sum(1 for i in issues if i["Warnings"])
    return issues_df, n_errors, n_warnings

# ---------------------------------
# UI
# ---------------------------------
st.set_page_config(page_title="Kobo Uploader (EU)", page_icon="⬆️", layout="centered")
st.title("⬆️ Excel → Kobo submission (EU)")

form_choice = st.selectbox("Destination form", list(FORMS.keys()))
asset_uid = FORMS[form_choice]["asset_uid"]
st.caption(f"Form: {FORMS[form_choice]['landing']}")

uploaded = st.file_uploader("Upload UNICEF Attendance Sheet (.xlsx)", type=["xlsx"])
sheet_name = st.text_input("Sheet name (fallback if no TblAttend)", value="Attend")

# one-time init for upload lock
if "uploading" not in st.session_state:
    st.session_state["uploading"] = False

if uploaded:
    try:
        df = load_attendance_df(uploaded, sheet_fallback=sheet_name, table_name="TblAttend")
    except Exception as e:
        st.error(str(e))
        st.stop()

    st.subheader("Preview (first 10 rows)")
    st.dataframe(df.head(10), use_container_width=True)

    # ---- Validation section ----
    st.subheader("Validation")
    st.caption(
        "Rules: (1) All fields are required **except** 'Helper' and "
        "'Other Service Provider or Facility'. (2) Score must be a **whole number 0–100**. "
        "(3) Provide at least one of **Parent PCHF** OR **Ambulatory**. "
        "(4) Dates must be valid (dd-mm-yyyy or Excel date)."
    )

    issues_df, n_errors, n_warnings = validate_dataframe(df)

    if n_errors > 0:
        st.error(f"{n_errors} row(s) have errors. Fix and re-upload.")
    else:
        st.success("All rows passed required checks.")

    if n_warnings > 0:
        st.warning(f"{n_warnings} row(s) have warnings (non-blocking).")

    if not issues_df.empty:
        st.dataframe(issues_df, use_container_width=True)
        csv = issues_df.to_csv(index=False).encode("utf-8-sig")
        st.download_button("Download validation report (CSV)", data=csv, file_name="validation_report.csv", mime="text/csv")

    # Build submissions (nested)
    submissions = [build_submission_nested(row) for _, row in df.iterrows()]
    st.write(f"Prepared {len(submissions)} submission(s)")

    # Submit button with progress & lockout (disabled if errors)
    uploading = st.session_state.get("uploading", False)
    submit_disabled = (not submissions) or uploading or (n_errors > 0)

    submit_btn = st.button(
        "🚀 Submit to Kobo",
        type="primary",
        use_container_width=True,
        disabled=submit_disabled,
        key="submit_btn",
    )

    if submit_btn:
        st.session_state["uploading"] = True
        progress = st.progress(0)
        status_box = st.empty()

        total = len(submissions)
        successes, failures, results = 0, 0, []

        for i, sub in enumerate(submissions, start=1):
            try:
                r = post_submission_json(asset_uid, sub)
                ctype = (r.headers.get("content-type", "") or "")
                data = r.json() if "application/json" in ctype else {
                    "status": r.status_code,
                    "text": r.text[:800],
                }
                ok = r.status_code in (200, 201, 202)
            except Exception as e:
                ok = False
                data = {"error": str(e)}
                r = None

            results.append({"ok": ok, "status": getattr(r, "status_code", "n/a"), "response": data})
            successes += 1 if ok else 0
            failures += 0 if ok else 1

            progress.progress(i / total)
            status_box.write(f"Uploaded {i}/{total}… (ok: {successes}, failed: {failures})")

        st.session_state["uploading"] = False

        st.write({"successes": successes, "failures": failures})
        st.json(results[:5])
        if failures == 0:
            st.success("All submissions sent successfully.")
        else:
            st.warning("Some submissions failed. See responses above.")
else:
    st.info("Upload an Excel file to begin.")
