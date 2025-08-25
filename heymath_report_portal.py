# pip install streamlit requests pandas openpyxl
import io
import json
import os
import re
import subprocess
import zipfile
from datetime import date
from pathlib import Path
from datetime import datetime
import pandas as pd
import streamlit as st
import sys
import platform, tempfile
from pathlib import Path
OUTDIR_BASE = Path(os.getenv("HM_OUTDIR", "/tmp/report_outputs"))
OUTDIR_BASE.mkdir(parents=True, exist_ok=True)


st.set_page_config(page_title="HeyMath Reports Config", page_icon="📊", layout="centered")
st.title("HeyMath! Reports — Setup")

st.caption({"OS": platform.platform(),
            "TMP": tempfile.gettempdir(),
            "Can write /tmp": os.access("/tmp", os.W_OK)})
# ========= Config =========
EXCEL_PATH = "School_Details_filled_with_subjects_final.xlsx"  # adjust path if needed
SUBJECT_MAX = 10  # SubjectCode 1..SubjectCode 10

# ========= Helpers =========
#@st.cache_data(show_spinner=False)
def load_master(path: str) -> pd.DataFrame:
    if not Path(path).exists():
        st.error(f"Master Excel not found: {path}")
        return pd.DataFrame()
    df = pd.read_excel(path)
    df.columns = [str(c).strip() for c in df.columns]
    return df

def ddmmyyyy(d: date) -> str:
    return d.strftime("%d/%m/%Y")

def subject_for_grade_wide(school_df: pd.DataFrame, grade_label: str) -> str:
    """
    Your sheet is 'wide': columns are SubjectCode 1..SubjectCode N.
    We extract the number from grade_label (e.g., 'Grade 2' -> 2),
    then read that SubjectCode column. If multiple rows exist for school,
    take the first non-empty value in that column.
    """
    m = re.search(r"(\d+)", str(grade_label))
    if not m:
        return ""
    idx = int(m.group(1))
    col = f"SubjectCode {idx}"
    if col not in school_df.columns:
        return ""
    vals = school_df[col].dropna().astype(str).str.strip()
    return vals.iloc[0] if not vals.empty and vals.iloc[0] else ""

def safe_date_str(d: date) -> str:
    return d.strftime("%d-%m-%Y")

def slug(s: str) -> str:
    return re.sub(r"[^A-Za-z0-9._-]+", "", (s or "").replace(" ", "_"))

# ========= Load Excel =========
if st.button("🔄 Reload Excel Data"):
    st.cache_data.clear()
    st.rerun()
master_df = load_master(EXCEL_PATH)
if master_df.empty:
    st.stop()

# Validate required columns quickly
required = {"SchoolName", "ShortCode", "GradeLabel"}
missing_cols = required - set(master_df.columns)
if missing_cols:
    st.error(f"Missing required columns in Excel: {', '.join(sorted(missing_cols))}")
    st.stop()

# Build school options
schools_df = master_df[["SchoolName", "ShortCode"]].dropna().drop_duplicates()
school_options = [f"{row.ShortCode} — {row.SchoolName}" for _, row in schools_df.iterrows()]
school_options = sorted(school_options, key=lambda s: s.split(" — ", 1)[-1].lower())

# --- School selection (reactive) ---
school_choice = st.selectbox("School", ["— Select a school —"] + school_options, index=0, key="school_select")

if school_choice == "— Select a school —":
    st.info("Pick a school to load Classes (Levels) and continue.")
    st.stop()

# Resolve selected school fields
short_code = school_choice.split(" — ", 1)[0].strip()
school_row = schools_df[schools_df["ShortCode"] == short_code].iloc[0]
school_name = str(school_row.SchoolName).strip()

# Restrict to this school’s rows
school_rows = master_df[master_df["ShortCode"] == short_code].copy()
school_rows.columns = [str(c).strip() for c in school_rows.columns]

# ---------- Build grade labels by scanning SubjectCode columns ----------
base_label = (
    school_rows["GradeLabel"].dropna().astype(str).str.strip().iloc[0]
    if "GradeLabel" in school_rows.columns and not school_rows["GradeLabel"].dropna().empty
    else "Grade"
)

indices_with_data = []
for i in range(1, SUBJECT_MAX + 1):
    col = f"SubjectCode {i}"
    if col in school_rows.columns:
        vals = school_rows[col].dropna().astype(str).str.strip()
        if not vals.empty and any(v for v in vals if v and v.lower() != "nan"):
            indices_with_data.append(i)

if indices_with_data:
    start_idx = min(indices_with_data)
    end_idx   = max(indices_with_data)
else:
    start_idx, end_idx = 1, 12  # fallback

grade_labels = [f"{base_label} {i}" for i in range(start_idx, end_idx + 1)]
label_to_code = {f"{base_label} {i}": f"{i:02d}" for i in range(start_idx, end_idx + 1)}

st.caption(f"Detected levels for {short_code}: {start_idx} → {end_idx} ({base_label})")

# ========= Form =========
with st.form("cfg"):
    # Dates...
    col1, col2 = st.columns(2)
    with col1:
        start = st.date_input("Start date", value=date.today().replace(day=1))
    with col2:
        end = st.date_input("End date", value=date.today())

    st.markdown("### Classes (Levels)")
    chosen_labels = st.multiselect("Select grades", grade_labels, default=grade_labels)
    levels = [{"code": label_to_code[lbl], "name": lbl} for lbl in chosen_labels]

    need_class_reports = st.checkbox("Need Class & Student Reports", value=False)  # no sidebar

    
    # Tokens + UA...
    st.markdown("### Session tokens (keep private)")
    jsessionid = st.text_input("JSESSIONID", type="password")
    auth_token = st.text_input("authToken", type="password")
    ua = st.text_input(
        "User-Agent",
        value="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/138.0.0.0 Safari/537.36"
    )

    submitted = st.form_submit_button("Save config.json")

# ========= Save config =========
if submitted:
    if not levels:
        st.error("Please select at least one grade.")
        st.stop()
    if not jsessionid or not auth_token:
        st.error("Please enter JSESSIONID and authToken.")
        st.stop()

    headers = {
        "accept": "*/*",
        "accept-language": "en-US,en;q=0.9",
        "content-type": "application/x-www-form-urlencoded;charset=UTF-8",
        "user-agent": ua,
        "referer": "https://report.heymath.com/reports/reports.action",
        "x-requested-with": "XMLHttpRequest",
        "cache-control": "no-cache",
        "pragma": "no-cache",
        "cookie": f"JSESSIONID={jsessionid}; authToken={auth_token}"
    }

    subject_map_by_level = {}
    missing = []
    for lv in levels:
        lbl = lv["name"]
        code = lv["code"]
        uuid = subject_for_grade_wide(school_rows, lbl)
        if uuid:
            subject_map_by_level[code] = uuid
        else:
            missing.append(f"{lbl} ({code})")

    if missing:
        st.warning("No SubjectCode found for: " + ", ".join(missing) + ". They will be skipped when running.")

    cfg = {
        "schoolName": school_name,                # display name
        "schoolShortCode": short_code,            # short code (e.g., AAIS)
        "levels": levels,                         # [{code, name}]
        "subjectMapByLevel": subject_map_by_level,# {"01": "<uuid>", ...}
        "gradeLabelMap": {lv["code"]: lv["name"] for lv in levels},  # optional
        "dateRange": {"start": ddmmyyyy(start), "end": ddmmyyyy(end)},
        "headers": headers
    }
    cfg["needClassReports"] = bool(need_class_reports)


    with open("config.json", "w", encoding="utf-8") as f:
        json.dump(cfg, f, indent=2)
    st.success("config.json saved.")
    st.code(json.dumps(cfg, indent=2), language="json")

# ========= Run + Download THIS run =========

# Build a run_id the same way the runner will if not provided
# run_id = f"{short_code}_{safe_date_str(start)}_{safe_date_str(end)}_{datetime.now().strftime('%Y%m%d_%H%M%S')}"

# if st.button("Run now"):
    # try:
        # result = subprocess.run(
            # ["python", "test_runner_all.py",
             # "--config", "config.json",
             # "--outdir", "report_outputs",
             # "--run-id", run_id],
            # capture_output=True, text=True, check=False
        # )
        # st.text_area("Output", result.stdout + "\n" + result.stderr, height=320)
        # if result.returncode == 0:
            # st.success(f"Finished. Run ID: {run_id}")
        # else:
            # st.error(f"Script exited with code {result.returncode}.")
    # except Exception as e:
        # st.error(f"Failed to run: {e}")
# ========= Run + Download THIS run =========
st.divider()
st.markdown("### Run the reports")

# Persist run_id across reruns so downloads keep working
if "run_id" not in st.session_state:
    st.session_state["run_id"] = ""

# Compute a candidate run id (used when the user clicks Run now)
candidate_run_id = f"{short_code}_{safe_date_str(start)}_{safe_date_str(end)}_{datetime.now().strftime('%Y%m%d_%H%M%S')}"

if st.button("Run now"):
    # Lock the run_id for this execution and future downloads
    st.session_state["run_id"] = candidate_run_id
    try:
        # This portion works for windows and ubuntu but not streamlit cloud
        # result = subprocess.run(
            # ["python", "test_runner_all.py",
             # "--config", "config.json",
             # "--outdir", "report_outputs",
             # "--run-id", st.session_state["run_id"]],
            # capture_output=True, text=True, check=False
        # )
        outdir = str(OUTDIR_BASE)  # /tmp/report_outputs on Cloud

        result = subprocess.run(
            [sys.executable, "-u", "test_runner_all.py",
             "--config", "config.json",
             "--outdir", outdir,
             "--run-id", st.session_state["run_id"]],
            capture_output=True, text=True, check=False
        )

        st.text_area("Output", result.stdout + "\n" + result.stderr, height=320)
        if result.returncode == 0:
            st.success(f"Finished. Run ID: {st.session_state['run_id']}")
        else:
            st.error(f"Script exited with code {result.returncode}.")
    except Exception as e:
        st.error(f"Failed to run: {e}")

# --- Downloads for THIS run only ---
st.divider()
st.markdown("### Download current run")

# Use the last successful run if available; otherwise show the candidate
run_id = st.session_state.get("run_id") or candidate_run_id

# This portion works for windows and ubuntu but not streamlit cloud
#run_folder = Path("report_outputs") / run_id

run_folder = OUTDIR_BASE / run_id

if run_folder.exists():
    files = sorted(list(run_folder.glob("*.csv")) + list(run_folder.glob("*.xls")) + list(run_folder.glob("*.xlsx")))
    if not files:
        st.info("No files were generated in this run.")
    else:
        # A) Per-file buttons with proper MIME
        for p in files:
            ext = p.suffix.lower()
            mime = "text/csv" if ext == ".csv" else ("application/vnd.ms-excel" if ext == ".xls" else "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            with p.open("rb") as fh:
                st.download_button(
                    label=f"Download {p.name}",
                    data=fh.read(),
                    file_name=p.name,
                    mime=mime,
                    key=f"dl-{p.name}"
                )

        # B) All-in-one ZIP for this run (CSV + XLS/XLSX)
        mem = io.BytesIO()
        with zipfile.ZipFile(mem, "w", zipfile.ZIP_DEFLATED) as zf:
            for p in files:
                zf.write(str(p), arcname=p.name)

        mem.seek(0)
        st.download_button(
            "Download THIS run as ZIP",
            data=mem,
            file_name=f"{run_id}.zip",
            mime="application/zip",
            key="zip-current-run"
        )
else:
    st.info("Run the reports to enable downloads for this run.")


