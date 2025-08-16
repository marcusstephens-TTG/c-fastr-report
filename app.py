#!/usr/bin/env python3
# -*- coding: utf-8 -*-

# Streamlit front-end for the C FASTR™ report generator.
# Restores full consultant inputs (Title page, Exec Summary, Conclusion),
# writes them to consultant_inputs.json in the schema the generator expects,
# runs generate_client_report_PATCHED.py, and offers a properly named download.

import os
import sys
import json
import subprocess
import datetime
import re
import streamlit as st

# ---------- Files your generator expects (leave as-is unless you change the script) ----------
SURVEY_CSV     = "CFastR_Survey_Data.csv"
MAPPING_CSV    = "CFASTR_Category_Mapping_V1.csv"
THRESHOLDS_CSV = "category_copy_thresholds.csv"
TEMPLATE_DOCX  = "client_report_template.docx"
OUTPUT_DOCX    = "Client_CFASTR_Report_Generated.docx"
SCRIPT         = "generate_client_report_PATCHED.py"
JSON_PATH      = "consultant_inputs.json"

st.set_page_config(page_title="C FASTR Report Builder", layout="centered")

# ---------- Helpers ----------
def build_download_name(company: str) -> str:
    """
    Return TTG - <Company> - YYYY-MM-DD - H-MM AM/PM.docx
    (Windows-safe: no weird chars, no seconds)
    """
    safe_company = re.sub(r"[^\w\s\-]", "", (company or "").strip())
    safe_company = re.sub(r"\s+", " ", safe_company).strip() or "Unknown Company"
    date_str = datetime.date.today().isoformat()
    time_str = datetime.datetime.now().strftime("%I-%M %p").lstrip("0")  # e.g., 3-07 PM
    return f"TTG - {safe_company} - {date_str} - {time_str}.docx"

def file_ok(path: str) -> bool:
    return os.path.exists(path)

# ---------- UI ----------
st.title("C FASTR Report Builder")
st.markdown(
    "When you click **Generate Report**, the app writes your inputs to "
    "`consultant_inputs.json`, runs the generator, and gives you a download."
)

with st.expander("Required files on the server (must sit next to app.py)", expanded=False):
    cols = st.columns(2)
    cols[0].write(("✅ " if file_ok(SURVEY_CSV) else "❌ ") + SURVEY_CSV)
    cols[1].write(("✅ " if file_ok(MAPPING_CSV) else "❌ ") + MAPPING_CSV)
    cols[0].write(("✅ " if file_ok(THRESHOLDS_CSV) else "❌ ") + THRESHOLDS_CSV)
    cols[1].write(("✅ " if file_ok(TEMPLATE_DOCX) else "❌ ") + TEMPLATE_DOCX)
    st.write(("✅ " if file_ok(SCRIPT) else "❌ ") + SCRIPT)

st.divider()
st.subheader("Consultant Inputs")

# Use a form to avoid reruns while typing.
with st.form("inputs_form", clear_on_submit=False):
    st.markdown("### Title Page")
    c1, c2 = st.columns(2)
    company_name = c1.text_input("Company Name *", placeholder="Acme Corp")
    client_contact = c2.text_input("Client Contact (optional)", placeholder="Jane Doe, CHRO")
    report_date = c1.text_input("Report Date (optional)", placeholder="e.g., 2025-08-16")
    confidentiality_notice = c2.text_area(
        "Confidentiality Notice (optional)",
        placeholder="Leave blank to use the default notice"
    )

    st.markdown("### Executive Summary")
    engagement_summary = st.text_area("Engagement Summary (optional)", height=120)
    results_summary    = st.text_area("Results Summary (optional)", height=120)
    suggested_actions  = st.text_area("Suggested Actions (optional)", height=120)

    st.markdown("### Conclusion")
    conc_overview           = st.text_area("Overview (optional)", height=120)
    conc_3090               = st.text_area("30/60/90 (optional)", height=120)
    conc_next_30            = st.text_area("Next 30 Days (optional)", height=100)
    conc_days_31_60         = st.text_area("Days 31–60 (optional)", height=100)
    conc_days_61_90         = st.text_area("Days 61–90 (optional)", height=100)
    conc_metrics_quarterly  = st.text_area("Metrics (Quarterly) (optional)", height=100)
    conc_closing_thoughts   = st.text_area("Closing Thoughts (optional)", height=100)

    submitted = st.form_submit_button("🚀 Generate Report", type="primary")

# Stop until they submit.
if not submitted:
    st.stop()

# ---------- Validate required fields ----------
errors = []
if not (company_name or "").strip():
    errors.append("Please enter a Company Name.")
if errors:
    st.error("\\n".join(errors))
    st.stop()

# ---------- Verify required on-disk files before running ----------
missing = [p for p in (SURVEY_CSV, MAPPING_CSV, THRESHOLDS_CSV, TEMPLATE_DOCX, SCRIPT) if not file_ok(p)]
if missing:
    st.error(
        "These required files are missing on the server (next to `app.py`). "
        "Add them and try again:\\n\\n" + "\\n".join(f"- {m}" for m in missing)
    )
    st.stop()

# ---------- Write consultant_inputs.json in the expected schema ----------
payload = {
    "title_page": {
        "company_name": (company_name or "").strip(),
        "client_contact": (client_contact or "").strip(),
        "report_date": (report_date or "").strip(),
        "confidentiality_notice": (confidentiality_notice or "").strip(),
    },
    "exec_summary": {
        "engagement_summary": (engagement_summary or "").strip(),
        "results_summary": (results_summary or "").strip(),
        "suggested_actions": (suggested_actions or "").strip(),
    },
    "conclusion": {
        "overview": (conc_overview or "").strip(),
        "thirty_sixty_ninety": (conc_3090 or "").strip(),
        "next_30_days": (conc_next_30 or "").strip(),
        "days_31_60": (conc_days_31_60 or "").strip(),
        "days_61_90": (conc_days_61_90 or "").strip(),
        "metrics_quarterly": (conc_metrics_quarterly or "").strip(),
        "closing_thoughts": (conc_closing_thoughts or "").strip(),
    },
    "meta": {
        "generated_at": datetime.datetime.now().isoformat(),
        "app_version": "ui-restore-1",
    },
}

try:
    with open(JSON_PATH, "w", encoding="utf-8") as f:
        json.dump(payload, f, indent=2, ensure_ascii=False)
except Exception as e:
    st.error(f"Failed to write {JSON_PATH}: {e}")
    st.stop()

# ---------- Run the generator ----------
st.write("Running report generation…")
try:
    result = subprocess.run(
        [sys.executable, SCRIPT],
        capture_output=True, text=True, check=False
    )
except Exception as e:
    st.error(f"Failed to launch script: {e}")
    st.stop()

# Show logs (helpful when troubleshooting)
if result.stdout:
    st.code(result.stdout, language="bash")
if result.stderr:
    st.code(result.stderr, language="bash")

# ---------- Offer the generated file for download ----------
if os.path.exists(OUTPUT_DOCX):
    with open(OUTPUT_DOCX, "rb") as f:
        data = f.read()
    st.success("Report generated! Download below:")
    st.download_button(
        "⬇️ Download Report",
        data=data,
        file_name=build_download_name(company_name),
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    )
else:
    st.error("The report file was not created. Check the logs above for errors.")
