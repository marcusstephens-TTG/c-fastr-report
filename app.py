# app.py — C FASTR: survey + mapping (Question Number, C FASTR Category, Polarity)
# Builds vertical breakdown charts with categorical colors (no risk bands)
from __future__ import annotations

import os, json, datetime as dt, inspect, csv, re
from pathlib import Path
from typing import Dict, Any, Tuple, List, Optional

import streamlit as st

# Charts
import matplotlib
matplotlib.use("Agg")  # headless
import matplotlib.pyplot as plt

# Word report
from docxtpl import DocxTemplate, InlineImage
from docx.shared import Mm

# =========================
# Diagnostics
# =========================
LOGFILE = Path("cfastr_run.log").resolve()
def _now() -> str:
    return dt.datetime.now().isoformat(timespec="seconds")
def log(msg: str, data: dict | None = None):
    payload = {"ts": _now(), "msg": msg, "data": data or {}}
    LOGFILE.parent.mkdir(parents=True, exist_ok=True)
    with LOGFILE.open("a", encoding="utf-8") as f:
        f.write(json.dumps(payload) + "\n")
    st.caption(f"[{payload['ts']}] {msg} — {payload['data']}")

# =========================
# Config
# =========================
DEFAULT_TEMPLATE_FILENAME = "client_report_template.docx"
DEFAULT_SURVEY_CSV        = os.getenv("CFASTR_SURVEY", "CFastR_Survey_Data.csv")
DEFAULT_MAPPING_CSV       = os.getenv("CFASTR_MAPPING", "CFASTR_Category_Mapping_V1.csv")

DISPLAY_TO_KEY = {
    "Collusion": "collusion",
    "Feedback, Receiving": "feedback_receiving",
    "Feedback, Giving": "feedback_giving",
    "Accountability": "accountability",
    "Sensitivity": "sensitivity",
    "Trust": "trust",
    "Relationship Focus": "relationship_focus",
    "Relationships": "relationship_focus",  # alias
}
CATEGORY_ORDER = [
    "collusion","feedback_receiving","feedback_giving",
    "accountability","sensitivity","trust","relationship_focus",
]

DEFAULT_FUNCTION_COL = "Business Function"
DEFAULT_LEVEL_COL    = "Job Level"  # or Title if you don’t have a level column

def get_template_path() -> Path:
    env = os.getenv("CFASTR_TEMPLATE")
    return Path(env).expanduser().resolve() if env else (Path(__file__).parent / DEFAULT_TEMPLATE_FILENAME).resolve()

def here(name: str) -> Path:
    return (Path(__file__).parent / name).resolve()

# =========================
# Colors (categorical only)
# =========================
CAT_PALETTE   = list(plt.get_cmap("tab10").colors)   # for single-value bars per category
FUNC_PALETTE  = list(plt.get_cmap("tab20").colors)   # functions
LEVEL_PALETTE = list(plt.get_cmap("Set3").colors)    # levels

CATEGORY_COLOR_MAP: Dict[str, Tuple[float,float,float]] = {
    cat: CAT_PALETTE[i % len(CAT_PALETTE)] for i, cat in enumerate(CATEGORY_ORDER)
}
FUNCTION_COLOR_MAP: Dict[str, Tuple[float,float,float]] = {}
LEVEL_COLOR_MAP: Dict[str, Tuple[float,float,float]] = {}
def color_for_function(name: str) -> Tuple[float,float,float]:
    k = (name or "").strip()
    if k not in FUNCTION_COLOR_MAP:
        FUNCTION_COLOR_MAP[k] = FUNC_PALETTE[len(FUNCTION_COLOR_MAP) % len(FUNC_PALETTE)]
    return FUNCTION_COLOR_MAP[k]
def color_for_level(name: str) -> Tuple[float,float,float]:
    k = (name or "").strip()
    if k not in LEVEL_COLOR_MAP:
        LEVEL_COLOR_MAP[k] = LEVEL_PALETTE[len(LEVEL_COLOR_MAP) % len(LEVEL_PALETTE)]
    return LEVEL_COLOR_MAP[k]

# =========================
# Normalize helpers
# =========================
def norm_cat_to_key(raw: str) -> Optional[str]:
    if not raw: return None
    if raw in DISPLAY_TO_KEY: return DISPLAY_TO_KEY[raw]
    r = re.sub(r"[^a-z]+", "", raw.lower())
    for disp, key in DISPLAY_TO_KEY.items():
        if re.sub(r"[^a-z]+", "", disp.lower()) == r:
            return key
    if raw.lower() in CATEGORY_ORDER: return raw.lower()
    return None

def val_to_num(v: str) -> Optional[float]:
    if v is None or v == "": return None
    try: return float(v)
    except: return None

# =========================
# Load survey
# =========================
def load_survey(path: Path) -> List[Dict[str,str]]:
    if not path.exists():
        raise FileNotFoundError(f"Survey file not found at {path}")
    rows: List[Dict[str,str]] = []
    with path.open("r", encoding="utf-8", newline="") as f:
        reader = csv.DictReader(f)
        for raw in reader:
            rows.append({ (k or "").strip(): (v or "").strip() for k,v in raw.items() })
    log("SURVEY_ROWS", {"count": len(rows)})
    return rows

# =========================
# Mapping resolution
# =========================
def resolve_field_from_question(qnum: str, headers: List[str]) -> Optional[str]:
    """
    Map a mapping “Question Number” to a survey header.
    Accepts exact match, case-insensitive, numeric → Q<num>/Question <num>, and prefix like 'Q12 - ...'
    """
    if not qnum: return None
    q = qnum.strip()

    # exact / case-insensitive
    for h in headers:
        if h == q or h.lower() == q.lower():
            return h

    # numeric candidates
    m = re.search(r"\d+", q)
    num = m.group(0) if m else None
    candidates = []
    if num:
        candidates += [f"Q{num}", f"Q{num}.", f"Q{num}_", f"Question {num}", f"Q{num} -", f"Q{num}—", f"Q{num} –"]
    # prefix variant
    q_prefix = re.sub(r"\s+[-–—:].*$", "", q)
    if q_prefix:
        candidates.append(q_prefix)

    lc_headers = [(h, h.lower()) for h in headers]
    for cand in candidates:
        lc = cand.lower()
        for h, hl in lc_headers:
            if hl == lc: return h
        for h, hl in lc_headers:
            if hl.startswith(lc): return h

    r_q = re.sub(r"[^a-z0-9]+", "", q.lower())
    for h in headers:
        r_h = re.sub(r"[^a-z0-9]+", "", h.lower())
        if r_h.startswith(r_q) or r_q.startswith(r_h):
            return h
    return None

def parse_polarity_to_good_when(polarity_raw: str) -> Optional[str]:
    """
    Convert mapping Polarity to good_when based on 1..5 scale (1=Strongly Agree):
      positive  -> agreement is good  -> low values good  -> 'low'
      negative  -> agreement is bad   -> high values good -> 'high'
    Accepts: 'positive'/'pos', 'negative'/'neg', and numeric encodings '1', '1.0', '+1' (positive) / '-1', '-1.0' (negative).
    """
    if polarity_raw is None: return None
    p = polarity_raw.strip().lower().replace(" ", "")
    if p in {"positive","pos","p","+","+1","1","+1.0","1.0","true","t","lowgood","goodlow"}:
        return "low"
    if p in {"negative","neg","n","-","-1","-1.0","false","f","highgood","goodhigh"}:
        return "high"
    return None

def load_and_resolve_mapping(path: Path, survey_headers: List[str]) -> List[Dict[str,str]]:
    """
    Mapping CSV columns (exact, per spec):
      - Question Number
      - C FASTR Category
      - Polarity           (positive/negative or 1/-1)
      - Special Interest   (optional)
    Returns rows with resolved survey field names and polarity converted to good_when ('low' or 'high').
    """
    if not path.exists():
        raise FileNotFoundError(f"Mapping file not found at {path}")

    raw_rows: List[Dict[str,str]] = []
    with path.open("r", encoding="utf-8", newline="") as f:
        reader = csv.DictReader(f)
        for r in reader:
            raw_rows.append({ (k or "").strip(): (v or "").strip() for k,v in r.items() })
    log("MAPPING_ROWS_FOUND", {"count": len(raw_rows)})

    resolved: List[Dict[str,str]] = []
    unresolved_samples: List[Dict[str,str]] = []

    for row in raw_rows:
        qnum      = row.get("Question Number", "")
        cat_disp  = row.get("C FASTR Category", "")
        polarity  = row.get("Polarity", "")
        special   = row.get("Special Interest", "")

        cat_key = norm_cat_to_key(cat_disp)
        good_when = parse_polarity_to_good_when(polarity)

        if not (qnum and cat_key and good_when):
            reason = "missing-field" if not (qnum and cat_key and polarity) else "unknown-polarity"
            unresolved_samples.append({"qn": qnum, "cat": cat_disp, "polarity": polarity, "reason": reason})
            continue

        field = resolve_field_from_question(qnum, survey_headers)
        if not field:
            unresolved_samples.append({"qn": qnum, "cat": cat_disp, "polarity": polarity, "reason":"no-matching-survey-column"})
            continue

        resolved.append({
            "question_number": qnum,
            "field": field,
            "category_key": cat_key,
            "good_when": good_when,
            "special_interest": special,
        })

    log("MAPPING_FIELDS_RESOLVED", {
        "resolved": len(resolved),
        "unresolved": len(unresolved_samples),
        "unresolved_sample": unresolved_samples[:10],
    })
    return resolved

# =========================
# Aggregation
# =========================
def compute_aggregates(
    survey_rows: List[Dict[str,str]],
    mapping: List[Dict[str,str]],
    function_col: str,
    level_col: str,
) -> Tuple[Dict[str,float], Dict[str, Dict[str, List[Tuple[str,float]]]]]:
    overall = {cat: {"good":0, "total":0} for cat in CATEGORY_ORDER}
    by_func: Dict[str, Dict[str, Dict[str,int]]] = {cat:{} for cat in CATEGORY_ORDER}
    by_level: Dict[str, Dict[str, Dict[str,int]]] = {cat:{} for cat in CATEGORY_ORDER}

    for row in survey_rows:
        func = row.get(function_col) or row.get("Business Function") or "Unknown"
        lvl  = row.get(level_col) or row.get("Job Level") or row.get("Title") or "Unknown"

        for m in mapping:
            field = m["field"]
            cat   = m["category_key"]
            gw    = m["good_when"]  # 'low' or 'high'
            val   = val_to_num(row.get(field))
            if val is None:
                continue

            # 1..5 scale (1=Strongly Agree):
            # good_when='low'  -> 1 or 2 are good
            # good_when='high' -> 4 or 5 are good
            is_good = (val <= 2.0) if gw == "low" else (val >= 4.0)

            overall[cat]["total"] += 1
            overall[cat]["good"]  += 1 if is_good else 0

            fslot = by_func[cat].setdefault(func, {"good":0, "total":0})
            fslot["total"] += 1; fslot["good"] += 1 if is_good else 0

            lslot = by_level[cat].setdefault(lvl, {"good":0, "total":0})
            lslot["total"] += 1; lslot["good"] += 1 if is_good else 0

    topline_pct: Dict[str,float] = {
        cat: ((c["good"]/c["total"])*100.0 if c["total"]>0 else 0.0)
        for cat, c in overall.items()
    }

    def pct_list(d: Dict[str, Dict[str,int]]) -> List[Tuple[str,float]]:
        out: List[Tuple[str,float]] = []
        for label, c in d.items():
            if c["total"]>0:
                out.append((label, (c["good"]/c["total"])*100.0))
        out.sort(key=lambda x: x[0].lower())
        return out

    breakdowns: Dict[str, Dict[str, List[Tuple[str,float]]]] = {
        cat: {"function": pct_list(by_func[cat]), "level": pct_list(by_level[cat])}
        for cat in CATEGORY_ORDER
    }

    log("AGG_DONE", {"topline_nonzero": {k: round(v,1) for k,v in topline_pct.items() if v>0}})
    return topline_pct, breakdowns

# =========================
# Charting
# =========================
def save_single_value_bar(category_key: str, label: str, score_pct: float, out_path: Path):
    color = CATEGORY_COLOR_MAP.get(category_key, CAT_PALETTE[0])
    fig, ax = plt.subplots(figsize=(4, 2))
    ax.bar([label], [score_pct], color=[color])
    ax.set_ylim(0, 100); ax.set_ylabel("% positive"); ax.set_title(f"{label} — {score_pct:.0f}%")
    for s in ("top","right"): ax.spines[s].set_visible(False)
    fig.tight_layout(); out_path.parent.mkdir(parents=True, exist_ok=True)
    fig.savefig(out_path, dpi=150); plt.close(fig)
    log("CHART_SAVED", {"type":"single","category":label,"score_pct":round(score_pct,1),"file":str(out_path)})

def save_breakdown_vertical(title: str, items: List[Tuple[str, float]], out_path: Path, mode: str):
    if not items: return
    labels, values = zip(*items)
    colors = [color_for_function(l) if mode=="function" else color_for_level(l) for l in labels]
    fig, ax = plt.subplots(figsize=(max(6.5, len(labels)*0.9), 4.2))
    ax.bar(list(range(len(labels))), list(values), color=colors)
    ax.set_ylim(0, 100); ax.set_ylabel("% positive"); ax.set_title(title)
    ax.set_xticks(list(range(len(labels))))
    ax.set_xticklabels(labels, rotation=20, ha="right")
    for s in ("top","right"): ax.spines[s].set_visible(False)
    fig.tight_layout(); out_path.parent.mkdir(parents=True, exist_ok=True)
    fig.savefig(out_path, dpi=150); plt.close(fig)
    log("CHART_SAVED", {"type":f"by_{mode}","title":title,"count":len(labels),"file":str(out_path)})

# =========================
# Report generation
# =========================
def generate_client_report(
    topline_pct: Dict[str,float],
    breakdown: Dict[str, Dict[str, List[Tuple[str, float]]]],
    out_path: Path,
) -> Path:
    st.caption(f"Running code from: `{Path(inspect.getsourcefile(generate_client_report)).resolve()}`")
    tpl = get_template_path()
    if not tpl.exists():
        log("TEMPLATE_NOT_FOUND", {"attempted": str(tpl)})
        raise FileNotFoundError(
            f"Template not found at {tpl}. Place '{DEFAULT_TEMPLATE_FILENAME}' next to app.py "
            f"or set CFASTR_TEMPLATE to an absolute path."
        )
    log("TEMPLATE_OPEN", {"path": str(tpl)})

    out_path = out_path.resolve()
    chart_dir = out_path.parent / "charts"

    singles_spec = [
        ("collusion","Collusion", topline_pct.get("collusion",0.0), "collusion_bar"),
        ("feedback_receiving","Feedback", topline_pct.get("feedback_receiving",0.0), "feedback_bar"),
        ("accountability","Accountability", topline_pct.get("accountability",0.0), "accountability_bar"),
        ("sensitivity","Sensitivity", topline_pct.get("sensitivity",0.0), "sensitivity_bar"),
        ("trust","Trust", topline_pct.get("trust",0.0), "trust_bar"),
        ("relationship_focus","Relationships", topline_pct.get("relationship_focus",0.0), "relationships_bar"),
    ]
    ctx: Dict[str, Any] = {}
    for cat_key, label, pct, bar_key in singles_spec:
        img = chart_dir / f"{bar_key}.png"
        save_single_value_bar(cat_key, label, float(pct), img)
        ctx[bar_key] = str(img)

    for cat_key in CATEGORY_ORDER:
        bucket = breakdown.get(cat_key, {})
        if bucket.get("function"):
            imgf = chart_dir / f"{cat_key}_by_function_chart.png"
            save_breakdown_vertical(f"{cat_key.replace('_',' ').title()}: by Business Function",
                                    bucket["function"], imgf, mode="function")
            ctx[f"{cat_key}_by_function_chart"] = str(imgf)
        if bucket.get("level"):
            imgl = chart_dir / f"{cat_key}_by_level_chart.png"
            save_breakdown_vertical(f"{cat_key.replace('_',' ').title()}: by Job Level",
                                    bucket["level"], imgl, mode="level")
            ctx[f"{cat_key}_by_level_chart"] = str(imgl)

    doc = DocxTemplate(str(tpl))
    present = set()
    try:
        present = set(doc.get_undeclared_template_variables() or [])
        log("TEMPLATE_VARS_FOUND", {"count": len(present), "sample": sorted(list(present))[:20]})
    except Exception:
        pass

    for k, v in topline_pct.items():
        ctx[f"{k}_pct"] = round(float(v),1)

    fin: Dict[str, Any] = {}
    for k, v in ctx.items():
        if isinstance(v, str) and v.lower().endswith((".png",".jpg",".jpeg")):
            fin[k] = InlineImage(doc, v, width=Mm(120 if "by_" in k else 90))
        else:
            fin[k] = v

    try:
        unresolved = sorted([v for v in present if v not in fin])
        log("UNRESOLVED_PLACEHOLDERS", {"count": len(unresolved), "vars": unresolved})
    except Exception:
        pass

    out_path.parent.mkdir(parents=True, exist_ok=True)
    log("RENDER_BEGIN", {"template": str(tpl), "output": str(out_path)})
    doc.render(fin); doc.save(str(out_path))
    log("REPORT_WRITTEN", {"output": str(out_path)})
    return out_path

# =========================
# Streamlit UI
# =========================
st.set_page_config(page_title="C FASTR — Survey-driven charts", layout="centered")
st.title("C FASTR — Survey-driven charts (vertical; categorical colors)")

with st.expander("How this works", expanded=True):
    st.markdown("""
- Reads **CFastR_Survey_Data.csv** (one row per respondent) and **CFASTR_Category_Mapping_V1.csv**.
- **Mapping columns (exact names):**
  - `Question Number` — your question ID (e.g., `Q12`, `12`, or `Q12 – text`)
  - `C FASTR Category` — Collusion / Feedback, Receiving / Feedback, Giving / Accountability / Sensitivity / Trust / Relationship Focus
  - `Polarity` — accepts `positive`/`pos` **or** numeric encodings `1`, `1.0`, `+1` (positive); `negative`/`neg` **or** `-1`, `-1.0` (negative).
    *Positive → 1/2 count good; Negative → 4/5 count good* (scale: 1=Strongly Agree … 5=Strongly Disagree).
  - `Special Interest` — optional; parsed, not used yet.
- The app resolves each `Question Number` to a **survey column header** (exact match, `Q<num>`, `Question <num>`, or prefix like `Q12 - ...`).
- Generates **category single bars** and **vertical breakdown charts** (by Function & by Level) with label-consistent colors.
- Template placeholders expected (examples):
  - `{{ collusion_by_function_chart }}`, `{{ collusion_by_level_chart }}`
  - `{{ trust_by_function_chart }}`, `{{ trust_by_level_chart }}`
""")

with st.form("inputs"):
    colA, colB = st.columns(2)
    with colA:
        survey_path  = st.text_input("Survey CSV path", value=DEFAULT_SURVEY_CSV)
        mapping_path = st.text_input("Mapping CSV path", value=DEFAULT_MAPPING_CSV)
    with colB:
        function_col = st.text_input("Function column name", value=DEFAULT_FUNCTION_COL)
        level_col    = st.text_input("Level column name", value=DEFAULT_LEVEL_COL)
    out_name = st.text_input("Output filename", "cfastr_report.docx")
    submitted = st.form_submit_button("Generate Report")

if submitted:
    try:
        survey_rows  = load_survey(here(survey_path))
        headers = list(survey_rows[0].keys()) if survey_rows else []
        mapping_rows = load_and_resolve_mapping(here(mapping_path), headers)

        if not mapping_rows:
            st.error("No mapping rows resolved. Check 'Question Number', 'C FASTR Category', and 'Polarity' values.")
            st.info("See Recent log output below for unresolved samples (polarity/field mismatches).")
        else:
            st.success(f"{len(mapping_rows)} mapping rows resolved to survey columns.")

        topline_pct, breakdowns = compute_aggregates(
            survey_rows, mapping_rows, function_col=function_col, level_col=level_col
        )

        expected_tpl = get_template_path()
        st.caption(f"Expected template location: `{expected_tpl}`")

        out_file = Path("out") / out_name
        result = generate_client_report(topline_pct, breakdowns, out_file)
        st.success(f"Report written to: {result}")
        st.code(str(result), language="bash")

        try:
            with LOGFILE.open("r", encoding="utf-8") as f:
                lines = f.readlines()[-80:]
            st.text_area("Recent log output", value="".join(lines), height=360)
        except Exception:
            st.info("No log file yet.")
    except Exception as e:
        st.error(str(e))
        try:
            with LOGFILE.open("r", encoding="utf-8") as f:
                lines = f.readlines()[-80:]
            st.text_area("Recent log output", value="".join(lines), height=360)
        except Exception:
            st.info("No log file yet.")
