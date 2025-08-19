# app.py — Single-file Streamlit app with simple, explicit diagnostics
from __future__ import annotations

import os, json, datetime as dt, inspect
from pathlib import Path
from typing import Dict, Any, Tuple

import streamlit as st

# Charts
import matplotlib
matplotlib.use("Agg")  # headless
import matplotlib.pyplot as plt

# Word report
from docxtpl import DocxTemplate, InlineImage
from docx.shared import Mm

# ---------------------------
# Minimal diagnostic helpers
# ---------------------------
LOGFILE = Path("cfastr_run.log").resolve()

def _now() -> str:
    return dt.datetime.now().isoformat(timespec="seconds")

def log(msg: str, data: dict | None = None):
    payload = {"ts": _now(), "msg": msg, "data": data or {}}
    LOGFILE.parent.mkdir(parents=True, exist_ok=True)
    with LOGFILE.open("a", encoding="utf-8") as f:
        f.write(json.dumps(payload) + "\n")
    # Also echo to Streamlit for immediate visibility
    st.caption(f"[{payload['ts']}] {msg} — {payload['data']}")

# ---------------------------
# Config: where is the template?
# Default: same folder as app.py, filename you specified
# ---------------------------
DEFAULT_TEMPLATE_FILENAME = "client_report_template.docx"

def get_template_path() -> Path:
    env = os.getenv("CFASTR_TEMPLATE")
    if env:
        p = Path(env).expanduser().resolve()
    else:
        # EXACTLY next to app.py (no templates/ folder)
        p = (Path(__file__).parent / DEFAULT_TEMPLATE_FILENAME).resolve()
    return p

# ---------------------------
# Deterministic color logic
# ---------------------------
PALETTE = {
    "poor": "#D32F2F",  # <60%
    "mid":  "#FBC02D",  # 60–79%
    "good": "#388E3C",  # 80%+
}

def pick_band_color(score_pct: float) -> Tuple[str, str]:
    if score_pct < 60:
        return "poor", PALETTE["poor"]
    if score_pct < 80:
        return "mid", PALETTE["mid"]
    return "good", PALETTE["good"]

# ---------------------------
# Chart generation (logs the path it writes)
# ---------------------------
def save_band_bar(category: str, score_pct: float, out_path: Path):
    band, color = pick_band_color(score_pct)
    fig, ax = plt.subplots(figsize=(4, 2))
    ax.bar([category], [score_pct], color=color)
    ax.set_ylim(0, 100)
    ax.set_ylabel("% positive")
    ax.set_title(f"{category} — {score_pct:.0f}%")
    for s in ("top", "right"):
        ax.spines[s].set_visible(False)
    fig.tight_layout()
    out_path.parent.mkdir(parents=True, exist_ok=True)
    fig.savefig(out_path, dpi=150)
    plt.close(fig)
    log("CHART_SAVED", {"category": category, "score_pct": score_pct, "band": band, "file": str(out_path)})

# ---------------------------
# Report generation (prints where it READS template and WRITES docx)
# ---------------------------
def generate_client_report(context: Dict[str, Any], out_path: Path) -> Path:
    # Confirm which file is running & from where
    module_file = Path(inspect.getsourcefile(generate_client_report)).resolve()
    st.caption(f"Running code from: `{module_file}`")

    # Resolve template right before use, then LOG IT
    tpl = get_template_path()
    if not tpl.exists():
        log("TEMPLATE_NOT_FOUND", {"attempted": str(tpl)})
        raise FileNotFoundError(
            f"Template not found at {tpl}. "
            f"Place '{DEFAULT_TEMPLATE_FILENAME}' next to app.py or set CFASTR_TEMPLATE to an absolute path."
        )
    log("TEMPLATE_OPEN", {"path": str(tpl)})

    # Build per-category charts, logging each saved file
    out_path = out_path.resolve()
    chart_dir = out_path.parent / "charts"
    categories = [
        ("Collusion", "collusion_pct", "collusion_bar"),
        ("Feedback", "feedback_pct", "feedback_bar"),
        ("Accountability", "accountability_pct", "accountability_bar"),
        ("Sensitivity", "sensitivity_pct", "sensitivity_bar"),
        ("Trust", "trust_pct", "trust_bar"),
        ("Relationships", "relationships_pct", "relationships_bar"),
    ]

    charts: Dict[str, str] = {}
    for label, pct_key, bar_key in categories:
        score = float(context.get(pct_key, 0.0))
        img_path = chart_dir / f"{bar_key}.png"
        save_band_bar(label, score, img_path)
        charts[bar_key] = str(img_path)

    # Prepare template context (inject image paths; docxtpl will embed them)
    base_context = dict(context)
    for _, _, bar_key in categories:
        base_context[bar_key] = charts[bar_key]

    # Render and WRITE the .docx (log exact path)
    doc = DocxTemplate(str(tpl))
    # Convert image paths to InlineImage
    for k, v in list(base_context.items()):
        if isinstance(v, str) and v.lower().endswith((".png", ".jpg", ".jpeg")):
            base_context[k] = InlineImage(doc, v, width=Mm(90))

    out_path.parent.mkdir(parents=True, exist_ok=True)
    log("RENDER_BEGIN", {"template": str(tpl), "output": str(out_path)})
    doc.render(base_context)
    doc.save(str(out_path))
    log("REPORT_WRITTEN", {"output": str(out_path)})

    return out_path

# ---------------------------
# Streamlit UI (required)
# ---------------------------
st.set_page_config(page_title="C FASTR Report Generator — Simple Diagnostics", layout="centered")
st.title("C FASTR Report Generator — Simple Diagnostics")

with st.expander("How diagnostics work", expanded=True):
    st.markdown(f"""
- The **template is expected next to `app.py`** by default (file: **`{DEFAULT_TEMPLATE_FILENAME}`**).  
  You can override with `CFASTR_TEMPLATE=/absolute/path/to/file.docx`.
- When the app **opens the template**, it logs the **exact path**.
- When the app **saves each chart** and the **final .docx**, it logs those **exact paths**.
- Logs are printed here and written to `cfastr_run.log`.
""")

with st.form("inputs"):
    col1, col2 = st.columns(2)
    with col1:
        collusion = st.number_input("Collusion % positive", 0, 100, 55)
        feedback = st.number_input("Feedback % positive", 0, 100, 62)
        accountability = st.number_input("Accountability % positive", 0, 100, 47)
    with col2:
        sensitivity = st.number_input("Sensitivity % positive", 0, 100, 73)
        trust = st.number_input("Trust % positive", 0, 100, 66)
        relationships = st.number_input("Relationships % positive", 0, 100, 81)
    out_name = st.text_input("Output filename", "cfastr_report.docx")
    submitted = st.form_submit_button("Generate Report")

if submitted:
    ctx = {
        "collusion_pct": collusion,
        "feedback_pct": feedback,
        "accountability_pct": accountability,
        "sensitivity_pct": sensitivity,
        "trust_pct": trust,
        "relationships_pct": relationships,
    }
    try:
        # Show the path we expect BEFORE generation
        expected_tpl = get_template_path()
        st.caption(f"Expected template location: `{expected_tpl}`")

        out_file = Path("out") / out_name
        result = generate_client_report(ctx, out_file)
        st.success(f"Report written to: {result}")
        st.code(str(result), language="bash")

        # Tail the last few log lines for quick visibility
        try:
            with LOGFILE.open("r", encoding="utf-8") as f:
                lines = f.readlines()[-20:]
            st.text_area("Recent log output", value="".join(lines), height=220)
        except Exception:
            st.info("No log file yet.")
    except Exception as e:
        st.error(str(e))
        try:
            with LOGFILE.open("r", encoding="utf-8") as f:
                lines = f.readlines()[-40:]
            st.text_area("Recent log output", value="".join(lines), height=260)
        except Exception:
            st.info("No log file yet.")
