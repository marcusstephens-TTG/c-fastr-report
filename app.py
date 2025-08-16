#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import os, sys, io, json, subprocess, datetime
import streamlit as st
import re

# Filenames expected by your generator script
SURVEY_CSV = "CFastR_Survey_Data.csv"
MAPPING_CSV = "CFASTR_Category_Mapping_V1.csv"
THRESHOLDS_CSV = "category_copy_thresholds.csv"
TEMPLATE_DOCX = "client_report_template.docx"
OUTPUT_DOCX = "Client_CFASTR_Report_Generated.docx"
SCRIPT = "generate_client_report_PATCHED.py"
JSON_PATH = "consultant_inputs.json"

st.set_page_config(page_title="C FASTR Report Builder", layout="centered")

# Helper: build a nice, Windows-safe download name
def build_download_name(company: str) -> str:
    safe_company = re.sub(r"[^\w\s\-]", "", (company or "").strip()) or "Unknown Company"
    date_str = datetime.date.today().isoformat()  # YYYY-MM-DD
    time_str = datetime.datetime.now().strftime("%I-%M %p").lstrip("0")  # e.g., 3-07 PM
    return f"TTG - {safe_company} - {date_str} - {time_str}.docx"

# --- UI helpers & status checks (unchanged) ---
def file_status(path):
    exists = os.path.exists(path)
    return f"✅ {path}" if exists else f"❌ {path} (missing)"

def build_notice(company):
    cn = company.strip() if company.strip() else "[Company name]"
    return (
        f"This is a demo notice for {cn}. "
        f"The resulting file will be named using the TTG convention at download time."
    )

st.title("C FASTR Report Builder")

st.markdown(
    "Upload the required inputs or keep the preloaded files next to this app. "
    "When you click **Generate**, the script runs and produces a report which you can download."
)

with st.expander("Required files (must be in the same folder as app.py)", expanded=True):
    st.write(f"{file_status(SURVEY_CSV)}")
    st.write(f"{file_status(MAPPING_CSV)}")
    st.write(f"{file_status(THRESHOLDS_CSV)}")
    st.write(f"{file_status(TEMPLATE_DOCX)}")
    st.write(f"{file_status(SCRIPT)}")

st.divider()

st.subheader("Consultant inputs")

company = st.text_input("Company Name", value="", placeholder="Acme Corp")
consultant = st.text_input("Consultant Name", value="", placeholder="Your name")
date_range = st.text_input("Date Range (optional)", value="", placeholder="e.g., Q2 2025")

# Optional preview/notice
if st.checkbox("Show preview notice", value=False):
    st.info(build_notice(company))

st.divider()

st.subheader("Generate report")

# Validate presence of required on-disk files before enabling generation
missing = [p for p in (SURVEY_CSV, MAPPING_CSV, THRESHOLDS_CSV, TEMPLATE_DOCX, SCRIPT) if not os.path.exists(p)]
if missing:
    st.error(
        "The following required files are missing. Please upload them to the repository root (same place as `app.py`):\n\n"
        + "\n".join(f"- {m}" for m in missing)
    )
    st.stop()

# Write the JSON that the generator expects (non-invasive; keeps same key names)
payload = {
    "company_name": company.strip(),
    "consultant_name": consultant.strip(),
    "date_range": date_range.strip(),
    # You can add more fields here if your generator uses them
}

try:
    with open(JSON_PATH, "w", encoding="utf-8") as f:
        json.dump(payload, f, indent=2, ensure_ascii=False)
except Exception as e:
    st.error(f"Failed to write {JSON_PATH}: {e}")
    st.stop()

if st.button("🚀 Generate Report", type="primary"):
    st.write("Running report generation…")
    try:
        result = subprocess.run(
            [sys.executable, SCRIPT],
            capture_output=True, text=True, check=False
        )
    except Exception as e:
        st.error(f"Failed to launch script: {e}")
        st.stop()

    # Show script output
    if result.stdout:
        st.code(result.stdout, language="bash")
    if result.stderr:
        st.code(result.stderr, language="bash")

    # Offer the generated DOCX to download
    if os.path.exists(OUTPUT_DOCX):
        with open(OUTPUT_DOCX, "rb") as f:
            data = f.read()
        st.success("Report generated! Download below:")
        st.download_button(
            "⬇️ Download Report",
            data=data,
            file_name=build_download_name(company),
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        )
    else:
        st.error("The report file was not created. Check the logs above for errors.")
