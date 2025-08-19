# app.py — C FASTR: compute from survey + mapping; vertical breakdown charts; categorical colors
from __future__ import annotations

import os, json, datetime as dt, inspect, csv, io, re
from pathlib import Path
from typing import Dict, Any, Tuple, List, Iterable

import streamlit as st

# Charts
import matplotlib
matplotlib.use("Agg")  # headless for servers
import matplotlib.pyplot as plt

# Word report
from docxtpl import DocxTemplate, InlineImage
from docx.shared import Mm

# =========================
# Diagnostics helpers
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
# Config & constants
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
KEY_TO_DISPLAY = {v: k for k, v in DISPLAY_TO_KEY.items()}

CATEGORY_ORDER = [
    "collusion",
    "feedback_receiving",
    "feedback_giving",
    "accountability",
    "sensitivity",
    "trust",
    "relationship_focus",
]

# Default survey column names (override in UI if needed)
DEFAULT_FUNCTION_COL = "Business Function"
DEFAULT_LEVEL_COL    = "Job Level"   # if absent, try "Title" heuristics, but better to have a column

# =========================
# Paths
# =========================
def get_template_path() -> Path:
    env = os.getenv("CFASTR_TEMPLATE")
    if env:
        p = Path(env).expanduser().resolve()
    else:
        p = (Path(__file__).parent / DEFAULT_TEMPLATE_FILENAME).resolve()
    return p

def get_repo_path(name: str) -> Path:
    """Resolve a file in the same folder as app.py (Streamlit Cloud mount is /mount/src)."""
    return (Path(__file__).parent / name).resolve()

# =========================
# Colors (categorical only)
# =========================
CAT_PALETTE  = list(plt.get_cmap("tab10").colors)   # top-line category color
FUNC_PALETTE = list(plt.get_cmap("tab20").colors)   # many distinct function colors
LEVEL_PALETTE= list(plt.get_cmap("Set3").colors)    # distinct level colors

CATEGORY_COLOR_MAP: Dict[str, Tuple[float,float,float]] = {
    cat: CAT_PALETTE[i % len(CAT_PALETTE)] for i, cat in enumerate(CATEGORY_ORDER)
}
FUNCTION_COLOR_MAP: Dict[str, Tuple[float,float,float]] = {}
LEVEL_COLOR_MAP: Dict[str, Tuple[float,float,float]] = {}

def color_for_function(name: str) -> Tuple[float,float,float]:
    key = (name or "").strip()
    if key not in FUNCTION_COLOR_MAP:
        idx = len(FUNCTION_COLOR_MAP) % len(FUNC_PALETTE)
        FUNCTION_COLOR_MAP[key] = FUNC_PALETTE[idx]
    return FUNCTION_COLOR_MAP[key]

def color_for_level(name: str) -> Tuple[float,float,float]:
    key = (name or "").strip()
    if key not in LEVEL_COLOR_MAP:
        idx = len(LEVEL_COLOR_MAP) % len(LEVEL_PALETTE)
        LEVEL_COLOR_MAP[key] = LEVEL_PALETTE[idx]
    return LEVEL_COLOR_MAP[key]

# =========================
# Utility: normalize
# =========================
def norm_cat_to_key(raw: str) -> str | None:
    if not raw:
        return None
    # direct
    if raw in DISPLAY_TO_KEY:
        return DISPLAY_TO_KEY[raw]
    # relaxed
    r = re.sub(r"[^a-z]+", "", raw.lower())
    for disp, key in DISPLAY_TO_KEY.items():
        if re.sub(r"[^a-z]+", "", disp.lower()) == r:
            return key
    # already a key?
    if raw.lower() in CATEGORY_ORDER:
        return raw.lower()
    return None

def pick_from(row: Dict[str,str], candidates: Iterable[str]) -> str | None:
    for c in candidates:
        if c in row and row[c] != "":
            return row[c]
    for c in candidates:
        # case-insensitive fallback
        for rk in row:
            if rk.strip().lower() == c.strip().lower() and row[rk] != "":
                return row[rk]
    return None

# =========================
# Load mapping & survey
# =========================
def load_mapping(path: Path) -> List[Dict[str,str]]:
    """
    Expected columns (case-insensitive):
      - field (or: question, column)  -> exact column name in survey CSV
      - category                      -> display name or key (see DISPLAY_TO_KEY)
      - good_when (or: polarity/direction/good_on) -> 'high' or 'low'
    """
    if not path.exists():
        raise FileNotFoundError(f"Mapping file not found at {path}")
    out: List[Dict[str,str]] = []
    with path.open("r", encoding="utf-8", newline="") as f:
        reader = csv.DictReader(f)
        for raw in reader:
            row = { (k or "").strip(): (v or "").strip() for k,v in raw.items() }
            field = pick_from(row, ["field","question","column","survey_field","survey_column"])
            cat_raw = pick_from(row, ["category","cat"])
            good_when = (pick_from(row, ["good_when","good_on","polarity","direction"]) or "").lower()
            if not (field and cat_raw and good_when in {"high","low"}):
                continue
            cat_key = norm_cat_to_key(cat_raw)
            if not cat_key:
                continue
            out.append({"field": field, "category_key": cat_key, "good_when": good_when})
    log("MAPPING_ROWS", {"count": len(out)})
    return out

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
# Compute good% aggregates
# =========================
def val_to_num(v: str) -> float | None:
    if v is None or v == "":
        return None
    try:
        return float(v)
    except:
        return None

def compute_aggregates(
    survey_rows: List[Dict[str,str]],
    mapping: List[Dict[str,str]],
    function_col: str,
    level_col: str,
) -> tuple[Dict[str,float], Dict[str, Dict[str, List[Tuple[str,float]]]]]:
    """
    Returns:
      topline_pct: {category_key -> pct_good}
      breakdowns:  {category_key -> {"function":[(label,pct_good)...], "level":[(label,pct_good)...]}}
    """
    # counters
    overall = {cat: {"good":0, "total":0} for cat in CATEGORY_ORDER}
    by_func: Dict[str, Dict[str, Dict[str,int]]] = {cat:{} for cat in CATEGORY_ORDER}  # cat -> func -> {good,total}
    by_level: Dict[str, Dict[str, Dict[str,int]]] = {cat:{} for cat in CATEGORY_ORDER} # cat -> lvl  -> {good,total}

    # iterate rows
    for row in survey_rows:
        func = row.get(function_col) or row.get("Business Function") or "Unknown"
        lvl  = row.get(level_col) or row.get("Job Level") or row.get("Title") or "Unknown"

        for m in mapping:
            field = m["field"]
            cat   = m["category_key"]
            pol   = m["good_when"]  # 'high' or 'low'
            val   = val_to_num(row.get(field))
            if val is None:
                continue
            # good if high (4/5) or low (1/2)
            is_good = (val >= 4.0) if pol == "high" else (val <= 2.0)

            overall[cat]["total"] += 1
            overall[cat]["good"]  += 1 if is_good else 0

            fslot = by_func[cat].setdefault(func, {"good":0, "total":0})
            fslot["total"] += 1; fslot["good"] += 1 if is_good else 0

            lslot = by_level[cat].setdefault(lvl, {"good":0, "total":0})
            lslot["total"] += 1; lslot["good"] += 1 if is_good else 0

    topline_pct = {
        cat: ( (overall[cat]["good"] / overall[cat]["total"])*100.0 if overall[cat]["total"]>0 else 0.0 )
        for cat in CATEGORY_ORDER
    }

    def pct_list(d: Dict[str, Dict[str,int]]) -> List[Tuple[str,float]]:
        items = []
        for label, c in d.items():
            if c["total"]>0:
                items.append((label, (c["good"]/c["total"])*100.0))
        # Sort by label for consistent output
        items.sort(key=lambda x: x[0].lower())
        return items

    breakdowns: Dict[str, Dict[str, List[Tuple[str,float]]]] = {}
    for cat in CATEGORY_ORDER:
        breakdowns[cat] = {
            "function": pct_list(by_func[cat]),
            "level":    pct_list(by_level[cat]),
        }
    log("AGG_DONE", {"topline_nonzero": {k:round(v,1) for k,v in topline_pct.items() if v>0}})
    return topline_pct, breakdowns

# =========================
# Charting
# =========================
def save_single_value_bar(category_key: str, label: str, score_pct: float, out_path: Path):
    color = CATEGORY_COLOR_MAP.get(category_key, CAT_PALETTE[0])
    fig, ax = plt.subplots(figsize=(4, 2))
    ax.bar([label], [score_pct], color=[color])
    ax.set_ylim(0, 100)
    ax.set_ylabel("% positive")
    ax.set_title(f"{label} — {score_pct:.0f}%")
    for s in ("top", "right"):
        ax.spines[s].set_visible(False)
    fig.tight_layout()
    out_path.parent.mkdir(parents=True, exist_ok=True)
    fig.savefig(out_path, dpi=150)
    plt.close(fig)
    log("CHART_SAVED", {"type":"single", "category": label, "score_pct": round(score_pct,1), "file": str(out_path)})

def save_breakdown_vertical(title: str, items: List[Tuple[str, float]], out_path: Path, mode: str):
    if not items:
        return
    labels, values = zip(*items)
    colors = [color_for_function(l) if mode=="function" else color_for_level(l) for l in labels]

    fig, ax = plt.subplots(figsize=(max(6.5, len(labels)*0.9), 4.2))
    ax.bar(list(range(len(labels))), list(values), color=colors)
    ax.set_ylim(0, 100)
    ax.set_ylabel("% positive")
    ax.set_title(title)
    ax.set_xticks(list(range(len(labels))), labels, rotation=20, ha="right")
    for s in ("top", "right"):
        ax.spines[s].set_visible(False)
    fig.tight_layout()
    out_path.parent.mkdir(parents=True, exist_ok=True)
    fig.savefig(out_path, dpi=150)
    plt.close(fig)
    log("CHART_SAVED", {"type": f"by_{mode}", "title": title, "count": len(labels), "file": str(out_path)})

# =========================
# Report generation
# =========================
def generate_client_report(
    topline_pct: Dict[str,float],
    breakdown: Dict[str, Dict[str, List[Tuple[str, float]]]],
    out_path: Path,
) -> Path:
    module_file = Path(inspect.getsourcefile(generate_client_report)).resolve()
    st.caption(f"Running code from: `{module_file}`")

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

    # --- single-value bars (category-colored) ---
    singles_spec = [
        ("collusion",           "Collusion",     topline_pct.get("collusion", 0.0),            "collusion_bar"),
        ("feedback_receiving",  "Feedback",      topline_pct.get("feedback_receiving", 0.0),   "feedback_bar"),
        ("accountability",      "Accountability",topline_pct.get("accountability", 0.0),       "accountability_bar"),
        ("sensitivity",         "Sensitivity",   topline_pct.get("sensitivity", 0.0),          "sensitivity_bar"),
        ("trust",               "Trust",         topline_pct.get("trust", 0.0),                "trust_bar"),
        ("relationship_focus",  "Relationships", topline_pct.get("relationship_focus", 0.0),   "relationships_bar"),
    ]
    base_context: Dict[str, Any] = {}
    for cat_key, label, pct, bar_key in singles_spec:
        img = chart_dir / f"{bar_key}.png"
        save_single_value_bar(cat_key, label, float(pct), img)
        base_context[bar_key] = str(img)

    # --- breakdown charts (vertical; categorical colors per label) ---
    for cat_key in CATEGORY_ORDER:
        bucket = breakdown.get(cat_key, {})
        if bucket.get("function"):
            imgf = chart_dir / f"{cat_key}_by_function_chart.png"
            title = f"{cat_key.replace('_',' ').title()}: by Business Function"
            save_breakdown_vertical(title, bucket["function"], imgf, mode="function")
            base_context[f"{cat_key}_by_function_chart"] = str(imgf)
        if bucket.get("level"):
            imgl = chart_dir / f"{cat_key}_by_level_chart.png"
            title = f"{cat_key.replace('_',' ').title()}: by Job Level"
            save_breakdown_vertical(title, bucket["level"], imgl, mode="level")
            base_context[f"{cat_key}_by_level_chart"] = str(imgl)

    # load template
    doc = DocxTemplate(str(tpl))

    # Try logging unresolved placeholders
    present = set()
    try:
        present = set(doc.get_undeclared_template_variables() or [])
        log("TEMPLATE_VARS_FOUND", {"count": len(present), "sample": sorted(list(present))[:20]})
    except Exception:
        pass

    # Expose topline % (optional if template uses them)
    for k, v in topline_pct.items():
        base_context[f"{k}_pct"] = round(float(v), 1)

    # Convert image paths to InlineImage
    final_context: Dict[str, Any] = {}
    for k, v in base_context.items():
        if isinstance(v, str) and v.lower().endswith((".png", ".jpg", ".jpeg")):
            width = 120 if "by_" in k else 90
            final_context[k] = InlineImage(doc, v, width=Mm(width))
        else:
            final_context[k] = v

    # Log unresolved after we filled what we can
    try:
        unresolved = sorted([v for v in present if v not in final_context])
        log("UNRESOLVED_PLACEHOLDERS", {"count": len(unresolved), "vars": unresolved})
    except Exception:
        pass

    out_path.parent.mkdir(parents=True, exist_ok=True)
    log("RENDER_BEGIN", {"template": str(tpl), "output": str(out_path)})
    doc.render(final_context)
    doc.save(str(out_path))
    log("REPORT_WRITTEN", {"output": str(out_path)})
    return out_path

# =========================
# Streamlit UI
# =========================
st.set_page_config(page_title="C FASTR Diagnostics", layout="centered")
st.title("C FASTR Diagnostics — Survey-driven charts (vertical, categorical colors)")

with st.expander("How this works", expanded=True):
    st.markdown("""
- Reads **CFastR_Survey_Data.csv** (one row per respondent) and **CFASTR_Category_Mapping_V1.csv** (which question maps to which category, and whether high/low is "good").
- Computes **good %** as:
  - If `good_when = high`: answers **4 or 5** count as good
  - If `good_when = low`:  answers **1 or 2** count as good
- Aggregates **top-line** and **breakdowns** by **Business Function** and **Job Level** (no risk colors; colors are categorical).
- Requires your Word template to include placeholders like `{{ collusion_by_function_chart }}` etc.
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
        # Load data
        survey_rows  = load_survey(get_repo_path(survey_path))
        mapping_rows = load_mapping(get_repo_path(mapping_path))

        # Compute aggregates
        topline_pct, breakdowns = compute_aggregates(
            survey_rows, mapping_rows, function_col=function_col, level_col=level_col
        )

        # Show path we expect BEFORE generation
        expected_tpl = get_template_path()
        st.caption(f"Expected template location: `{expected_tpl}`")

        # Build report
        out_file = Path("out") / out_name
        result = generate_client_report(topline_pct, breakdowns, out_file)
        st.success(f"Report written to: {result}")
        st.code(str(result), language="bash")

        # Tail logs
        try:
            with LOGFILE.open("r", encoding="utf-8") as f:
                lines = f.readlines()[-60:]
            st.text_area("Recent log output", value="".join(lines), height=320)
        except Exception:
            st.info("No log file yet.")
    except Exception as e:
        st.error(str(e))
        try:
            with LOGFILE.open("r", encoding="utf-8") as f:
                lines = f.readlines()[-60:]
            st.text_area("Recent log output", value="".join(lines), height=320)
        except Exception:
            st.info("No log file yet.")
