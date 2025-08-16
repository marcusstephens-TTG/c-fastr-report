#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import os, sys, io, json, subprocess, datetime
import streamlit as st

# Filenames expected by your generator script
SURVEY_CSV = "CFastR_Survey_Data.csv"
MAPPING_CSV = "CFASTR_Category_Mapping_V1.csv"
THRESHOLDS_CSV = "category_copy_thresholds.csv"
TEMPLATE_DOCX = "client_report_template.docx"
OUTPUT_DOCX = "Client_CFASTR_Report_Generated.docx"
SCRIPT = "generate_client_report_PATCHED.py"
JSON_PATH = "consultant_inputs.json"

st.set_page_config(page_title="C FASTR Report Builder", layout="centered")

st.title("C FASTR — Report Builder")

st.write(
    "Fill in the consultant sections below, then click **Generate Report**. "
    "Your patched script will run, and you can download the completed Word report."
)

def file_status(name):
    return "✅" if os.path.exists(name) else "⚠️"

with st.expander("Required files (must be in this folder)"):
    st.write(f"{file_status(SURVEY_CSV)} `{SURVEY_CSV}`")
    st.write(f"{file_status(MAPPING_CSV)} `{MAPPING_CSV}`")
    st.write(f"{file_status(THRESHOLDS_CSV)} `{THRESHOLDS_CSV}`")
    st.write(f"{file_status(TEMPLATE_DOCX)} `{TEMPLATE_DOCX}`")
    st.write(f"{file_status(SCRIPT)} `{SCRIPT}`")
    st.caption("Tip: put this app and all inputs in the same directory (e.g., `~/Downloads`).")

# --- Title page ---
st.header("Title Page")

def build_notice(company):
    cn = company.strip() if company.strip() else "[Company name]"
    return (f"This report contains confidential and proprietary information intended solely for {cn}. "
            f"Do not distribute, reproduce, or share without written permission from The Transformation Guild and {cn}.")

# Session state for notice so user edits aren't overwritten
if "notice_user_set" not in st.session_state:
    st.session_state.notice_user_set = False
if "notice" not in st.session_state:
    st.session_state.notice = build_notice("")

def on_notice_change():
    st.session_state.notice_user_set = True

col1, col2 = st.columns(2)
with col1:
    company = st.text_input("Company Name", value="", placeholder="Acme Corp")
    contact = st.text_input("Client Contact", value="", placeholder="Jane Doe")
with col2:
    date_val = st.date_input("Date of Report", value=datetime.date.today())
    date_str = date_val.isoformat()

# auto-fill notice unless user has edited
if not st.session_state.notice_user_set:
    st.session_state.notice = build_notice(company)
notice = st.text_area("Confidentiality Notice", value=st.session_state.notice, key="notice", on_change=on_notice_change)

st.divider()

# --- Executive Summary ---
st.header("Executive Summary")
exec_eng = st.text_area("Engagement Summary", height=140, placeholder="What we did, who we engaged, scope/timing…")
exec_res = st.text_area("Results Summary", height=140, placeholder="Key outcomes, signals, highlights…")
exec_act = st.text_area("Suggested Actions", height=140, placeholder="Top 3–5 recommendations…")

st.divider()

# --- Conclusion (auto-drafted; optional overrides) ---
st.header("Conclusion — Optional Overrides")
st.caption("These are auto-drafted from the data. Only fill these if you want to override the auto text.")

conc_ov   = st.text_area("Overview (optional override)", height=160, placeholder="Leave blank to use auto-generated Overview")
conc_3090 = st.text_area("30/60/90 (optional override)", height=160, placeholder="Leave blank to use auto-generated 30/60/90")

colA, colB, colC = st.columns(3)
with colA:
    next_30 = st.text_area("Next 30 days (optional)", height=140, placeholder="")
with colB:
    days_31_60 = st.text_area("Days 31–60 (optional)", height=140, placeholder="")
with colC:
    days_61_90 = st.text_area("Days 61–90 (optional)", height=140, placeholder="")

metrics_q = st.text_area("What to measure (quarterly) (optional)", height=140, placeholder="")
closing   = st.text_area("Closing thoughts (optional)", height=140, placeholder="")

st.divider()

# --- Generate button ---
btn = st.button("🚀 Generate Report", type="primary")

def save_json():
    data = {
        "title_page": {
            "company_name": company.strip(),
            "client_contact": contact.strip(),
            "report_date": date_str.strip(),
            "confidentiality_notice": notice.strip(),
        },
        "exec_summary": {
            "engagement_summary": exec_eng.strip(),
            "results_summary": exec_res.strip(),
            "suggested_actions": exec_act.strip(),
        },
        "conclusion": {
            "overview": conc_ov.strip(),
            "thirty_sixty_ninety": conc_3090.strip(),
            "next_30_days": next_30.strip(),
            "days_31_60": days_31_60.strip(),
            "days_61_90": days_61_90.strip(),
            "metrics_quarterly": metrics_q.strip(),
            "closing_thoughts": closing.strip(),
        },
    }
    with open(JSON_PATH, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
    return data

if btn:
    # Basic validations
    missing = []
    for f in [SURVEY_CSV, MAPPING_CSV, THRESHOLDS_CSV, TEMPLATE_DOCX, SCRIPT]:
        if not os.path.exists(f):
            missing.append(f)
    if missing:
        st.error("Missing required files:\n" + "\n".join(f"- {m}" for m in missing))
        st.stop()

    # Save consultant inputs to JSON for the generator to consume
    _ = save_json()
    st.info("Saved consultant_inputs.json")

    # Run the generator script with the same Python interpreter as Streamlit
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
        st.download_button("⬇️ Download Client_CFASTR_Report_Generated.docx", data=data, file_name=OUTPUT_DOCX, mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
    else:
        st.error("The report file was not created. Check the logs above for errors.")
