# test_runner_all.py
# ---------------------------------------------------------
# HeyMath! report runner (CQR removed)
# - Loads report_def_final.json + config.json
# - Discovers school sections + friendly labels from Excel
# - Runs table/teacher/class reports
# - SQR: uses site academic start -> today (per school)
# - Writes CSV and .meta.json sidecars; keeps a run manifest
# ---------------------------------------------------------

import argparse
import csv
import json
import os
import re
import time
from datetime import datetime

import pandas as pd
import requests

# =========================
# Logging + small utilities
# =========================

def _supports_utf8() -> bool:
    enc = (getattr(__import__('sys').stdout, "encoding", "") or "").lower()
    return "utf" in enc

OK   = "✅" if _supports_utf8() else "[OK]"
WARN = "⚠️" if _supports_utf8() else "[WARN]"
ERR  = "❌" if _supports_utf8() else "[ERR]"
ARROW = "→" if _supports_utf8() else "->"   # for logs only

CLASS_CODE_RE = re.compile(r"\b\d{2}[A-Z]\d\b")  # e.g., 04A0

def sanitize_filename(name: str) -> str:
    """Make a filename safe for most filesystems (keep spaces/dashes)."""
    return (name.replace("/", "-").replace("\\", "-").replace(":", "-")
                .replace("*", "-").replace("?", "-").replace('"', "'")
                .replace("<", "(").replace(">", ")").replace("|", "-"))

def add_school_suffix(file_name: str, school: str) -> str:
    """Append school name for clarity, if present."""
    if not school:
        return file_name
    root, ext = os.path.splitext(file_name)
    return f"{root}_{sanitize_filename(school)}{ext}"

def safe_date_for_name(d: str) -> str:
    """Convert DD/MM/YYYY → DD-MM-YYYY (for filenames)."""
    return d.replace("/", "-").replace("\\", "-").replace(":", "-")

def strip_double_pipes(s: str) -> str:
    return s.strip().replace("||", " ").strip().strip("|")

def excel_text_guard(s: str) -> str:
    """Prevent Excel scientific notation like 01E0 or 1.2E+05 by forcing text."""
    if re.fullmatch(r"\d+(?:\.\d+)?[Ee][+-]?\d+", s):
        return f'="{s}"'
    return s

def clean_value(v):
    """Normalize strings for Excel & remove mojibake."""
    if isinstance(v, str):
        s = v.strip()
        if "Ã" in s and "�" not in s:  # repair common mojibake when possible
            try:
                s = s.encode("latin-1").decode("utf-8")
            except Exception:
                pass
        s = strip_double_pipes(s)
        s = excel_text_guard(s)
        return s
    return v

def clean_row(row: dict) -> dict:
    return {k: clean_value(v) for k, v in row.items()}

def decode_maybe_json(value):
    """If a value looks JSON-encoded (stringified), keep json.loads-ing until it's a dict/list."""
    while isinstance(value, str):
        try:
            value = json.loads(value)
        except json.JSONDecodeError:
            break
    return value

# ============================================
# Excel: discover sections + friendly label map
# ============================================

def load_sections_and_labels(xlsx_path: str, school_short_code: str):
    """
    Returns (codes, labels_map) where labels_map maps '04A0' -> 'Grade 4 TPP Group 1'.
    Parses 'Class_Sections' for the given ShortCode. Tolerates separators and " - "/":".
    """
    try:
        df = pd.read_excel(xlsx_path, sheet_name=0)
    except Exception as e:
        print(f"{WARN} Could not read sections Excel '{xlsx_path}': {e}")
        return [], {}
    df.columns = [str(c).strip() for c in df.columns]
    if "ShortCode" not in df.columns or "Class_Sections" not in df.columns:
        print(f"{WARN} Excel must have columns 'ShortCode' and 'Class_Sections'. Found: {df.columns.tolist()}")
        return [], {}
    row = df[df["ShortCode"].astype(str).str.upper() == str(school_short_code).upper()]
    if row.empty:
        print(f"{WARN} No row for school '{school_short_code}' in sections Excel.")
        return [], {}

    raw = str(row.iloc[0]["Class_Sections"] or "")
    parts = re.split(r"[;,\n]+", raw)
    labels, codes, seen = {}, [], set()
    for part in parts:
        part = part.strip()
        if not part:
            continue
        m = CLASS_CODE_RE.search(part)
        if not m:
            continue
        code = m.group(0)
        label = part[m.end():].lstrip(" :-–—").strip() or f"Class {code}"
        labels[code] = label
        if code not in seen:
            seen.add(code); codes.append(code)
    if codes:
        print(f"Sections discovered: {', '.join(codes)}")
    else:
        print(f"{WARN} No sections discovered for this school.")
    return codes, labels

# ==================================================
# Filename templating: inject class label & run dates
# ==================================================

SECTION_LABEL_MAP = {}  # filled at runtime

def inject_class_and_dates(template_name: str, class_code: str, start_label: str, end_label: str,
                           grade_label_map: dict = None) -> str:
    """
    Replace <class>, <start_date>, <end_date> in template_name with friendly values.
    Also supports legacy <SAFE_START>/<SAFE_END> and (SAFE_START)/(SAFE_END).
    Class label preference:
      1) SECTION_LABEL_MAP['04A0'] → 'Grade 4 TPP Group 1'
      2) grade_label_map[class_code] or grade_label_map[:2]
      3) 'Class 04A0'
    """
    grade_label_map = grade_label_map or {}
    if class_code:
        human = (SECTION_LABEL_MAP.get(class_code)
                 or grade_label_map.get(class_code)
                 or grade_label_map.get(class_code[:2]))
        class_label = human if human else f"Class {class_code}"
    else:
        class_label = ""

    safe_name = (template_name
                 .replace("<class>", class_label)
                 .replace("(class)", (class_label + "_") if class_label else "")
                 .replace("<start_date>", start_label).replace("(start_date)", start_label)
                 .replace("<end_date>",   end_label).  replace("(end_date)",   end_label)
                 .replace("<SAFE_START>", start_label).replace("(SAFE_START)", start_label)
                 .replace("<SAFE_END>",   end_label).  replace("(SAFE_END)",   end_label))

    # Readability tweaks
    if class_label and f"{class_label}{start_label}" in safe_name:
        safe_name = safe_name.replace(f"{class_label}{start_label}", f"{class_label}_{start_label}")
    if start_label + end_label in safe_name:
        safe_name = safe_name.replace(start_label + end_label, f"{start_label}_{end_label}")
    while "__" in safe_name:
        safe_name = safe_name.replace("__", "_")
    return safe_name

# ==========================
# Extract/parsing helpers
# ==========================

def extract_values(report_name: str, response_obj):
    """Convert server payload into a list[dict] when possible, per EXTRACT_MAP."""
    mode = EXTRACT_MAP.get(report_name)
    if mode is None:
        return None

    obj = decode_maybe_json(response_obj)
    # Treat explicit empty markers as no rows
    if obj in (None, False) or (isinstance(obj, str) and obj.strip().lower() in ("false", "null", "")):
        return []

    if mode == "direct_list":
        if isinstance(obj, list):
            return obj
        if isinstance(obj, dict):
            for k in ("data", "TABLE_DATA", "tableData", "tableJSON", "lessoninfo"):
                if k in obj:
                    return decode_maybe_json(obj[k])
        return None

    # generic: pick the named key but accept common aliases
    if isinstance(obj, dict):
        if mode not in obj:
            for k in ("data", "TABLE_DATA", "tableData", "tableJSON", "lessoninfo"):
                if k in obj:
                    mode = k
                    break
        return decode_maybe_json(obj.get(mode))

    return None

def ensure_rows(values):
    """Validate we have list[dict]; return (all_fieldnames, rows) or None."""
    if not isinstance(values, list):
        return None
    union_keys, fixed_rows = set(), []
    for row in values:
        if isinstance(row, dict):
            union_keys.update(row.keys())
            fixed_rows.append(row)
        else:
            return None
    return list(union_keys), fixed_rows

def build_params(report: dict, class_code: str, subject_code: str) -> dict:
    """
    Expand JSON-like params string in report definition into a dict with placeholders filled.
    Also backfills SQR's assessmentType if missing.
    """
    param_str = report.get("params", "")
    param_str = (param_str.replace("<start_date>", START_DATE)
                           .replace("<end_date>", END_DATE)
                           .replace("<class>", class_code or "")
                           .replace("<subject>", subject_code or ""))
    params = json.loads("{" + param_str + "}") if param_str else {}

    # Safety: SQR requires assessmentType=1 in many deployments
    if report.get("name") == "Student Quiz Performance Report" and "assessmentType" not in params:
        params["assessmentType"] = "1"

    return params

# =====================
# CLI & config loading
# =====================

parser = argparse.ArgumentParser()
parser.add_argument("--config", default="config.json")
parser.add_argument("--outdir",  default="report_outputs")
parser.add_argument("--run-id",  default="")
args = parser.parse_args()

# Load report definitions (list or {"reports":[...]})
with open("report_def_final.json", "r", encoding="utf-8") as f:
    defs = json.load(f)
reports = defs["reports"] if isinstance(defs, dict) and "reports" in defs else defs

# Load portal config
with open(args.config, "r", encoding="utf-8") as f:
    config = json.load(f)

school_name        = (config.get("schoolName") or "").strip()
school_short_code  = (config.get("schoolShortCode") or "").strip()
grade_label_map    = config.get("gradeLabelMap", {})       # {"04A0":"Grade 4 ...", "04":"Grade 4", ...}
subject_map_by_lvl = config.get("subjectMapByLevel", {})   # {"04": "<uuid>", ...}
NEED_CLASS         = bool(config.get("needClassReports", False))

# Build headers + cookies (honor 'cookie' header if provided)
headers = dict(config.get("headers", {}))
cookies = {}
for hk in list(headers.keys()):
    if hk.lower() == "cookie":
        cookie_items = str(headers.pop(hk)).split("; ")
        for item in cookie_items:
            if "=" in item:
                key, value = item.split("=", 1)
                cookies[key.strip()] = value.strip()

# Dates
date_cfg   = config["dateRange"]
START_DATE = date_cfg["start"]  # DD/MM/YYYY
END_DATE   = date_cfg["end"]
SAFE_START = safe_date_for_name(START_DATE)
SAFE_END   = safe_date_for_name(END_DATE)

# Run folder
timestamp   = datetime.now().strftime("%Y%m%d_%H%M%S")
RUN_ID      = args.run_id or f"{(school_short_code or school_name)}_{SAFE_START}_{SAFE_END}_{timestamp}"
output_dir  = os.path.join(args.outdir, RUN_ID)
os.makedirs(output_dir, exist_ok=True)

print(f"Output folder: {output_dir}")
print(f"RUN_ID: {RUN_ID}")

# ==================================
# Section discovery & grade filtering
# ==================================

levels = config.get("levels", []) or []
selected_levels = {
    str((lv or {}).get("code", "")).zfill(2)
    for lv in levels
    if (lv or {}).get("code")
}

SECTION_CODES, SECTION_LABEL_MAP = ([], {})
if NEED_CLASS:
    codes, labels = load_sections_and_labels("School_Details_filled_with_sections_final.xlsx", school_short_code)
    SECTION_LABEL_MAP = labels
    # if user selected specific grade levels, filter sections by first 2 digits
    SECTION_CODES = [c for c in codes if (not selected_levels or c[:2] in selected_levels)]
else:
    print("[INFO] Class reports disabled in config; skipping per-class sections.")

print("Selected grade codes:", ", ".join(sorted(selected_levels)) or "(none)")
print("Sections taken for Class reports:", ", ".join(SECTION_CODES) or "(none)")

# ==================================
# Report lists & extract map
# ==================================

REPORT_NAME_SECTIONS = "Level Lessons Usage Report for each Class"  # must match JSON (if used)

TABLE_DATA_REPORTS = [
    "School Logins Report",
    "School Lessons Usage Report",
    "School Assignments Usage Report",
    REPORT_NAME_SECTIONS,
]

OTHER_REPORTS = [
    "Teacher Assignment Log Quiz",
    "Teacher Assignment Log Worksheet",
    "Teacher Assignment Log Prasso",
    "Teacher Assignment Log Reading",
    "All Teachers Usage Logins",
    "All Teachers Usage Lesson Accessed",
    "All Teachers Usage Assignments assigned",
    "Class Login Report",
    "Class Lessons Report",
    "Class Assignment Report",
    "Student Quiz Performance Report",
    # "Consolidated Quiz Report",  # removed
]

# Auto-include any custom "Class ..." reports defined in JSON
CLASS_REPORTS = [r["name"] for r in reports if str(r.get("name","")).startswith("Class ")]
OTHER_REPORTS += [n for n in CLASS_REPORTS if n not in OTHER_REPORTS]

# Where to extract rows from the JSON payload
EXTRACT_MAP = {
    "School Logins Report": "TABLE_DATA",
    "School Lessons Usage Report": "TABLE_DATA",
    "School Assignments Usage Report": "TABLE_DATA",
    REPORT_NAME_SECTIONS: "TABLE_DATA",

    "Teacher Assignment Log Quiz": "direct_list",
    "Teacher Assignment Log Worksheet": "direct_list",
    "Teacher Assignment Log Prasso": "direct_list",
    "Teacher Assignment Log Reading": "direct_list",

    "All Teachers Usage Logins": "direct_list",
    "All Teachers Usage Lesson Accessed": "lessoninfo",
    "All Teachers Usage Assignments assigned": "tableJSON",

    "Class Login Report": "data",
    "Class Lessons Report": "direct_list",
    "Class Assignment Report": "data",

    # SQR returns {"data":"[ {...},{...} ]"} ; 'data' then decode_maybe_json handles the inner list
    "Student Quiz Performance Report": "data",
}

# ========================
# Run-plan construction
# ========================

def runs_teacher_reports_from_map():
    """Map each selected level to subject UUID for the 4 teacher logs."""
    runs = []
    for lvl in levels:
        code = (lvl or {}).get("code", "")
        subj_uuid = subject_map_by_lvl.get(code, "")
        if code and subj_uuid:
            runs.append(({"code": code}, {"code": subj_uuid}))
    return runs or [({}, {})]

RUN_PLAN = {
    # single-run school reports
    "School Logins Report": [({}, {})],
    "School Lessons Usage Report": [({}, {})],
    "School Assignments Usage Report": [({}, {})],

    # per-class table report (if present in JSON)
    REPORT_NAME_SECTIONS: [({ "code": lvl["code"] }, {}) for lvl in levels] or [({}, {})],

    # teacher reports (per selected level + subject mapping)
    "Teacher Assignment Log Quiz": runs_teacher_reports_from_map(),
    "Teacher Assignment Log Worksheet": runs_teacher_reports_from_map(),
    "Teacher Assignment Log Prasso": runs_teacher_reports_from_map(),
    "Teacher Assignment Log Reading": runs_teacher_reports_from_map(),

    # usage rollups (single run)
    "All Teachers Usage Logins": [({}, {})],
    "All Teachers Usage Lesson Accessed": [({}, {})],
    "All Teachers Usage Assignments assigned": [({}, {})],

    # class reports (explicit per section)
    "Class Login Report":      [({"code": c}, {}) for c in (SECTION_CODES or [])] or [({}, {})],
    "Class Lessons Report":    [({"code": c}, {}) for c in (SECTION_CODES or [])] or [({}, {})],
    "Class Assignment Report": [({"code": c}, {}) for c in (SECTION_CODES or [])] or [({}, {})],
    "Student Quiz Performance Report": [({"code": c}, {}) for c in (SECTION_CODES or [])] or [({}, {})],
    # "Consolidated Quiz Report":        [...],  # removed
}

# Include any custom "Class ..." reports from JSON
for name in [r for r in [r["name"] for r in reports] if str(r).startswith("Class ")]:
    RUN_PLAN[name] = [({"code": c}, {}) for c in (SECTION_CODES or [])] or [({}, {})]

# ================================
# Manifest helpers (sidecar + run)
# ================================

RUN_MANIFEST = {
    "runId": RUN_ID,
    "school": school_name,
    "schoolShortCode": school_short_code,
    "startDate": START_DATE,
    "endDate": END_DATE,
    "outputDir": output_dir,
    "generatedAt": datetime.now().isoformat(timespec="seconds"),
    "files": []
}

def write_sidecar_meta(path: str, meta: dict):
    sidecar = path + ".meta.json"
    try:
        with open(sidecar, "w", encoding="utf-8") as f:
            json.dump(meta, f, indent=2)
    except Exception as e:
        print(f"{WARN} Could not write sidecar for {path}: {e}")

def add_to_run_manifest(entry: dict):
    RUN_MANIFEST["files"].append(entry)

def finalise_run_manifest():
    try:
        manifest_path = os.path.join(output_dir, "run_manifest.json")
        with open(manifest_path, "w", encoding="utf-8") as f:
            json.dump(RUN_MANIFEST, f, indent=2, ensure_ascii=False)

        # print without emoji (Windows safe)
        print(f"[OK] Wrote run manifest: {manifest_path}")

    except Exception as e:
        print(f"{WARN} Could not write run_manifest.json: {e}")

# ============
# Main runner
# ============

all_report_names = [r["name"] for r in reports]
def _close_matches(name, all_names):
    import difflib as _dl
    return _dl.get_close_matches(name, all_names, n=3, cutoff=0.5)

for report_name in TABLE_DATA_REPORTS + OTHER_REPORTS:
    # Gate all "Class ..." reports behind the checkbox
    if report_name in CLASS_REPORTS and not NEED_CLASS:
        print(f"[INFO] Skipping {report_name} because needClassReports=False")
        continue
    # Gate SQR too (CQR removed)
    if report_name in {"Student Quiz Performance Report"} and not NEED_CLASS:
        print(f"[INFO] Skipping {report_name} because needClassReports=False")
        continue

    report = next((r for r in reports if r["name"] == report_name), None)
    if not report:
        print(f"{ERR} Report not found in config: {report_name}")
        suggestions = _close_matches(report_name, all_report_names)
        if suggestions:
            print("   Did you mean one of:", suggestions)
        continue

    runs = RUN_PLAN.get(report_name, [({}, {})])

    for class_level, subject in runs:
        try:
            class_code   = (class_level or {}).get("code", "")
            subject_code = (subject or {}).get("code", "")

            # Build params with placeholders replaced (adds SQR assessmentType if missing)
            try:
                params = build_params(report, class_code, subject_code)
            except json.JSONDecodeError as e:
                print(f"{WARN} JSON decode error in params for [{report_name}]: {e}")
                continue

            # Filename date labels (SQR may override with site window)
            use_start_label = SAFE_START
            use_end_label   = SAFE_END

            # --- SQR: preflight to get site academic start + today, then override params + labels
            if report_name == "Student Quiz Performance Report":
                base = "https://report.heymath.com/reports/generateReport.action"
                today = datetime.now()
                # Probe a wide window (Jan 1 → today) so the server returns the school’s academic start
                probe_params = {
                    "timePeriod": "predefined",
                    "startDate": f"01/01/{today.strftime('%Y')}",
                    "endDate": today.strftime("%d/%m/%Y"),
                    "levelSection": class_code,
                }
                try:
                    probe = requests.get(base, params=probe_params, headers=headers, cookies=cookies, timeout=20)
                    obj = decode_maybe_json(probe.json() if (probe.headers.get("Content-Type","").startswith("application/json")) else probe.text)
                    data = obj.get("data") if isinstance(obj, dict) else obj
                    payload = decode_maybe_json(data)
                    items = []
                    if isinstance(payload, dict):
                        for bucket in ("newData","oldData"):
                            b = payload.get(bucket)
                            if isinstance(b, dict):
                                items.extend(list(b.values()))
                    if items:
                        items.sort(key=lambda it: int(it.get("createdDate", 0)), reverse=True)
                        ftd = str(items[0].get("fromToDate") or "")
                        m = re.match(r"\s*(\d{2}/\d{2}/\d{2,4})\s*-\s*(\d{2}/\d{2}/\d{2,4})\s*", ftd)
                        if m:
                            site_start = m.group(1)  # academic start as the site uses it
                            today_str = today.strftime("%d/%m/%Y")
                            params["startDate"] = site_start
                            params["endDate"]   = today_str
                            use_start_label     = safe_date_for_name(site_start)
                            use_end_label       = safe_date_for_name(today_str)
                except Exception:
                    pass
                print(f"[SQR] {class_code}: using site window {params.get('startDate')} , {params.get('endDate')}")

            # --- Make the request for the report itself
            response = requests.request(
                method=report["method"],
                url=report["url"],
                headers=headers,
                cookies=cookies,
                params=params if report["method"].upper() == "GET" else None,
                timeout=30,
            )

            print(f"{OK} {report_name} [{class_code or '-'} | {subject_code or '-'}]: Status {response.status_code}")
            if response.status_code != 200:
                print(f"{WARN} Response content:", response.text[:240])
                continue

            # --- Robust JSON parse + normalize double-encoded payloads
            try:
                resp_obj = response.json()
            except Exception:
                raw = response.content
                try:
                    text = raw.decode("utf-8", errors="strict")
                except UnicodeDecodeError:
                    response.encoding = "utf-8"
                    text = response.text
                if "Ã" in text and "�" not in text:
                    try:
                        text = text.encode("latin-1").decode("utf-8", errors="ignore")
                    except Exception:
                        pass
                try:
                    resp_obj = json.loads(text)
                except Exception:
                    print(f"{ERR} Could not parse response as JSON at all.")
                    continue
            resp_obj = decode_maybe_json(resp_obj)

            # --- Generic extraction (CSV writers for all non-CQR)
            values = extract_values(report_name, resp_obj)
            values = decode_maybe_json(values)
            ensured = ensure_rows(values)

            # header-only CSV if server returns explicit empty list
            if not ensured:
                if isinstance(values, list) and len(values) == 0:
                    # Avoid NameError if EXPECTED_HEADERS is not defined in this file
                    fieldnames = (globals().get("EXPECTED_HEADERS", {}) or {}).get(report_name, [])
                    template = report.get("outputFile") or report.get("outfile") or f"{report_name}.csv"
                    safe_name = sanitize_filename(template)
                    safe_name = inject_class_and_dates(safe_name, class_code, use_start_label, use_end_label, grade_label_map)
                    safe_name = add_school_suffix(safe_name, school_name)
                    file_path = os.path.join(output_dir, safe_name)
                    with open(file_path, "w", newline="", encoding="utf-8-sig") as f:
                        writer = csv.DictWriter(f, fieldnames=fieldnames)
                        writer.writeheader()
                    print(f"{WARN} Saved header-only CSV to {file_path}")

                    meta = {
                        "reportName": report_name,
                        "fileName": os.path.basename(file_path),
                        "filePath": file_path,
                        "url": report["url"], "method": report["method"],
                        "classCode": class_code, "subjectCode": subject_code,
                        "startDate": START_DATE, "endDate": END_DATE,
                        "savedAt": datetime.now().isoformat(timespec="seconds"),
                        "rowCount": 0,
                        "classLabel": SECTION_LABEL_MAP.get(class_code) or
                                      grade_label_map.get(class_code) or
                                      grade_label_map.get(class_code[:2]) or
                                      (f"Class {class_code}" if class_code else "")
                    }
                    # Reflect site window in meta for SQR
                    if report_name == "Student Quiz Performance Report":
                        meta["startDate"] = params.get("startDate", START_DATE)
                        meta["endDate"]   = params.get("endDate", END_DATE)

                    write_sidecar_meta(file_path, meta)
                    add_to_run_manifest(meta)
                    continue

                # otherwise nothing usable
                keys = list(resp_obj.keys()) if isinstance(resp_obj, dict) else type(resp_obj)
                print(f"{WARN} No usable rows. Keys/type: {keys}")
                print(" Sample:", (str(resp_obj)[:240] if not isinstance(resp_obj, dict)
                                  else str({k: str(resp_obj[k])[:120] for k in list(resp_obj)[:2]})))
                continue

            # normal CSV with rows
            fieldnames, rows = ensured
            rows = [clean_row(r) for r in rows]

            template = report.get("outputFile") or report.get("outfile") or f"{report_name}.csv"
            safe_name = sanitize_filename(template)
            safe_name = inject_class_and_dates(safe_name, class_code, use_start_label, use_end_label, grade_label_map)
            safe_name = add_school_suffix(safe_name, school_name)
            file_path = os.path.join(output_dir, safe_name)

            with open(file_path, "w", newline="", encoding="utf-8-sig") as f:
                writer = csv.DictWriter(f, fieldnames=list(fieldnames))
                writer.writeheader()
                for row in rows:
                    writer.writerow(row)

            print(f"{OK} Saved to {file_path}")

            meta = {
                "reportName": report_name,
                "fileName": os.path.basename(file_path),
                "filePath": file_path,
                "url": report["url"], "method": report["method"],
                "classCode": class_code, "subjectCode": subject_code,
                "startDate": START_DATE, "endDate": END_DATE,
                "savedAt": datetime.now().isoformat(timespec="seconds"),
                "rowCount": len(rows),
                "classLabel": SECTION_LABEL_MAP.get(class_code) or
                              grade_label_map.get(class_code) or
                              grade_label_map.get(class_code[:2]) or
                              (f"Class {class_code}" if class_code else "")
            }
            if report_name == "Student Quiz Performance Report":
                meta["startDate"] = params.get("startDate", START_DATE)
                meta["endDate"]   = params.get("endDate", END_DATE)

            write_sidecar_meta(file_path, meta)
            add_to_run_manifest(meta)

        except Exception as e:
            print(f"{ERR} Error in report [{report_name}]: {e}")

# Final manifest
finalise_run_manifest()
print(f"{OK} Done.")
