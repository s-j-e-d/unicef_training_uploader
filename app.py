from __future__ import annotations

import json
import re
import uuid
from typing import Any, Optional, Tuple, Dict, List
import os 

import pandas as pd
import streamlit as st
import requests
import warnings

# (optional) silence harmless openpyxl data-validation warnings
warnings.filterwarnings("ignore", message="Data Validation extension is not supported and will be removed")

# -------------------------------
# HARD-CODED DESTINATION & TOKEN
# -------------------------------
KOBO_TOKEN = os.environ.get("KOBO_TOKEN") or st.secrets.get("KOBO_TOKEN", "")
if not KOBO_TOKEN:
    st.error("Missing KOBO_TOKEN. Add it in Streamlit Cloud → App → Settings → Secrets.")
    st.stop()
ASSET_UID = "aS5UNu8wuynhJN87JYopFQ"                     # exact form asset uid
KC_BASE = "https://kc-eu.kobotoolbox.org"               # EU KoboCAT
SUBMIT_URL = f"{KC_BASE}/api/v1/submissions.json"       # JSON submissions

FORMS = {
    "Training Attendance (EU)": {
        "asset_uid": ASSET_UID,
        "landing": "https://eu.kobotoolbox.org/#/forms/aS5UNu8wuynhJN87JYopFQ/landing",
    }
}

# ---------------------------------
# UNICEF Attendance headers (source)
# ---------------------------------
L = {
    "name": "Name",
    "phone": "Phone Number",
    "email": "Email",
    "gender": "Gender",
    "works_hf": "Are you working in Health Facility (Yes or No) / \nМісце роботи - Заклад Охорони Здоров'я (Так чи ні)",
    "position": "Position / \nПосада",
    "oblast": "Oblast / \nОбласть",
    "rayon": "Rayon / \nРайон",
    "hromada": "Hromada / Громада",
    "settlement": "Settlement / Населений пункт",
    "facility": "Health Facility/Pcode/Care_type / \nМедичний заклад/ЄДРПОУ/Тип допомоги",
    "service_provider": "Service Provider / Місце надання послуг",
    "service_provider_other": "Other Service Provider or Facility / Інше Місце надання послуг чи Заклад",
    "score": "Post-Training Test Score %",
    "start_mm_yyyy": "Training start date\n mm/yyyy",
    "end_mm_yyyy": "Training end date\n mm/yyyy",
    "training_name": "Name of Training",
    "place": "Place where the training is conducted",
}

# -------------------------------
# Helpers
# -------------------------------
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

def parse_mm_yyyy(v: Any) -> Optional[str]:
    """Accept 'mm/yyyy' or an Excel date; return 'yyyy-mm-01' (ISO) or None."""
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return None
    ts = pd.to_datetime(v, errors="coerce")
    if pd.notna(ts):
        return f"{ts.year:04d}-{ts.month:02d}-01"
    s = str(v).strip()
    m = re.match(r"^(\d{1,2})\s*[-/.]\s*(\d{4})$", s)
    return f"{m.group(2)}-{int(m.group(1)):02d}-01" if m else None

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

# -------------------------------
# Build one submission (NESTED)
# -------------------------------
def build_submission_nested(row: pd.Series) -> Dict[str, Any]:
    first, second = parse_name(row.get(L["name"]))

    score_val = None
    if L["score"] in row and pd.notna(row.get(L["score"])):
        score_val = pd.to_numeric(row.get(L["score"]), errors="coerce")
        score_val = int(score_val) if pd.notna(score_val) else None

    submission = {
        "metadata": {
            "training_name": None if pd.isna(row.get(L["training_name"])) else str(row.get(L["training_name"])).strip(),
            "start_date": parse_mm_yyyy(row.get(L["start_mm_yyyy"])),
            "end_date": parse_mm_yyyy(row.get(L["end_mm_yyyy"])),
            "place": None if pd.isna(row.get(L["place"])) else str(row.get(L["place"])).strip(),
        },
        "personal_info": {
            "first_name": first,
            "second_name": second,
            "phone": None if pd.isna(row.get(L["phone"])) else str(row.get(L["phone"])),
            "email": None if pd.isna(row.get(L["email"])) else str(row.get(L["email"])).strip(),
            "gender": normalize_gender(row.get(L["gender"])),
            "works_in_hf": normalize_yes_no(row.get(L["works_hf"])),
        },
        "work_location": {
            "position": None if pd.isna(row.get(L["position"])) else str(row.get(L["position"])).strip(),
            "oblast": None if pd.isna(row.get(L["oblast"])) else str(row.get(L["oblast"])).strip(),
            "rayon": None if pd.isna(row.get(L["rayon"])) else str(row.get(L["rayon"])).strip(),
            "hromada": None if pd.isna(row.get(L["hromada"])) else str(row.get(L["hromada"])).strip(),
            "settlement": None if pd.isna(row.get(L["settlement"])) else str(row.get(L["settlement"])).strip(),
            "facility_code": None if pd.isna(row.get(L["facility"])) else str(row.get(L["facility"])).strip(),
            "service_provider": None if pd.isna(row.get(L["service_provider"])) else str(row.get(L["service_provider"])).strip(),
            "service_provider_other": None if pd.isna(row.get(L["service_provider_other"])) else str(row.get(L["service_provider_other"])).strip(),
        },
        "training_score": {
            "training_score_pct": score_val,
        },
        # KoboCAT requires nested meta
        "meta": {"instanceID": f"uuid:{uuid.uuid4()}"},
    }
    return prune(submission)

# -------------------------------
# UI
# -------------------------------
st.set_page_config(page_title="Kobo Uploader (EU)", page_icon="⬆️", layout="centered")
st.title("⬆️ Excel → Kobo submission (EU)")

form_choice = st.selectbox("Destination form", list(FORMS.keys()))
asset_uid = FORMS[form_choice]["asset_uid"]
st.caption(f"Form: {FORMS[form_choice]['landing']}")

uploaded = st.file_uploader("Upload UNICEF Attendance Sheet (.xlsx)", type=["xlsx"])
sheet_name = st.text_input("Sheet name", value="Attendance Sheet")

# one-time init for upload lock
if "uploading" not in st.session_state:
    st.session_state["uploading"] = False

if uploaded:
    # Find header row by scanning first 80 rows for key headers
    raw = pd.read_excel(uploaded, sheet_name=sheet_name, header=None)
    header_idx: Optional[int] = None
    for i in range(min(80, len(raw))):
        vals = [str(x).strip() if not pd.isna(x) else "" for x in list(raw.iloc[i, :].values)]
        if all(h in vals for h in [L["name"], L["phone"], L["email"]]):
            header_idx = i
            break

    if header_idx is None:
        st.error("Could not locate the header row. Please confirm the sheet structure.")
    else:
        df = pd.read_excel(uploaded, sheet_name=sheet_name, header=header_idx).dropna(how="all")

        st.subheader("Preview (first 10 rows)")
        st.dataframe(df.head(10))

        submissions = [build_submission_nested(row) for _, row in df.iterrows()]
        st.write(f"Prepared {len(submissions)} submission(s)")

        # Submit button with progress & lockout
        uploading = st.session_state.get("uploading", False)
        submit_disabled = (not submissions) or uploading

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
