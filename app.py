# app.py — C FASTR Report Generator (single file)
from __future__ import annotations

import os, io, json, re, datetime as dt
from pathlib import Path
from typing import Dict, Any, List, Tuple, Optional

import streamlit as st
import pandas as pd
import numpy as np

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt

from docxtpl import DocxTemplate, InlineImage
from docx.shared import Mm

# ----------------------
# Constants / Config
# ----------------------
APP_TITLE = "C FASTR — Report Generator (single-file)"
DEFAULT_TEMPLATE = "client_report_template.docx"
DEFAULT_SURVEY_CSV = "CFastR_Survey_Data.csv"
DEFAULT_MAPPING_CSV = "CFASTR_Category_Mapping_V1.csv"
DEFAULT_THRESHOLDS_CSV = "category_copy_thresholds.csv"

# GitHub defaults (raw) — update if your repo/branch moves
GITHUB_BASE = "https://raw.githubusercontent.com/marcusstephens-TTG/c-fastr-report/main/"
GH_SURVEY     = GITHUB_BASE + "CFastR_Survey_Data.csv"
GH_MAPPING    = GITHUB_BASE + "CFASTR_Category_Mapping_V1.csv"
GH_THRESHOLDS = GITHUB_BASE + "category_copy_thresholds.csv"

CATEGORY_DISPLAY_TO_KEY = {
    "Collusion": "collusion",
    "Feedback, Receiving": "feedback_receiving",
    "Feedback, Giving": "feedback_giving",
    "Accountability": "accountability",
    "Sensitivity": "sensitivity",
    "Trust": "trust",
    "Relationship Focus": "relationship_focus",
    "Relationships": "relationship_focus",  # alias
}
CATEGORY_KEYS_IN_ORDER = [
    "collusion", "feedback_receiving", "feedback_giving",
    "accountability", "sensitivity", "trust", "relationship_focus"
]

DEFAULT_FUNCTION_COL = "Business Function"
DEFAULT_LEVEL_COL    = "Title"  # use Title for “Job Level” per spec

LOGFILE = Path("cfastr_run.log").resolve()

def _now() -> str:
    return dt.datetime.now().isoformat(timespec="seconds")

def log(msg: str, data: Optional[dict]=None):
    payload = {"ts": _now(), "msg": msg, "data": data or {}}
    LOGFILE.parent.mkdir(parents=True, exist_ok=True)
    with LOGFILE.open("a", encoding="utf-8") as f:
        f.write(json.dumps(payload) + "\n")
    st.caption(f"[{payload['ts']}] {msg} — {payload['data']}")

# ----------------------
# Utilities
# ----------------------
def norm_cat_to_key(label: str) -> Optional[str]:
    if not label: return None
    if label in CATEGORY_DISPLAY_TO_KEY:
        return CATEGORY_DISPLAY_TO_KEY[label]
    r = re.sub(r"[^a-z]+", "", label.lower())
    for disp, key in CATEGORY_DISPLAY_TO_KEY.items():
        if re.sub(r"[^a-z]+", "", disp.lower()) == r:
            return key
    if label.lower() in CATEGORY_KEYS_IN_ORDER:
        return label.lower()
    return None

def parse_polarity_to_good_when(val: str) -> Optional[str]:
    """Return 'low' if 1/2 are good; 'high' if 4/5 are good; None if unknown."""
    if val is None: return None
    p = str(val).strip().lower().replace(" ", "")
    # Normalize fancy dashes early
    p = p.replace("−", "-").replace("–", "-").replace("—", "-")
    if p in {"positive","pos","p","+","+1","1","+1.0","1.0","true","t"}:
        return "low"
    if p in {"negative","neg","n","-","-1","-1.0","false","f"}:
        return "high"
    return None

def map_question_to_column(qnum: str, headers: List[str]) -> Optional[str]:
    """Map 'Question Number' to an actual survey header (robust)."""
    if not qnum: return None
    q = qnum.strip()
    # Exact/ci
    for h in headers:
        if h == q or h.lower() == q.lower():
            return h
    # Extract numeric core and try variants
    m = re.search(r"\d+", q)
    num = m.group(0) if m else None
    candidates = []
    if num:
        candidates += [
            f"Q{num}", f"Q{num}.", f"Q{num}_", f"Question {num}",
            f"{num}", f"Q{num} -", f"Q{num} –", f"Q{num}—"
        ]
    # Prefix like 'Q12 - prompt'
    q_prefix = re.sub(r"\s+[-–—:].*$", "", q)
    if q_prefix:
        candidates.append(q_prefix)
    # compare CI equals or startswith
    lc_headers = [(h, h.lower()) for h in headers]
    for cand in candidates:
        lc = cand.lower()
        for h, hl in lc_headers:
            if hl == lc:
                return h
        for h, hl in lc_headers:
            if hl.startswith(lc):
                return h
    # Relaxed alnum compare
    rq = re.sub(r"[^a-z0-9]+", "", q.lower())
    for h in headers:
        rh = re.sub(r"[^a-z0-9]+", "", h.lower())
        if rh.startswith(rq) or rq.startswith(rh):
            return h
    return None

def ensure_numeric_1_5(df: pd.DataFrame, cols: List[str]) -> pd.DataFrame:
    for c in cols:
        if c in df.columns:
            df[c] = pd.to_numeric(df[c], errors="coerce")
    return df

# ----------------------
# Data loading
# ----------------------
def read_csv_auto(path_or_url: str) -> pd.DataFrame:
    return pd.read_csv(path_or_url)

def load_survey(path_or_url: str) -> pd.DataFrame:
    df = read_csv_auto(path_or_url)
    log("SURVEY_ROWS", {"count": int(df.shape[0])})
    return df

def load_mapping(path_or_url: str, survey_headers: List[str]) -> Tuple[pd.DataFrame, pd.DataFrame]:
    raw = read_csv_auto(path_or_url)
    log("MAPPING_ROWS_FOUND", {"count": int(raw.shape[0])})
    # Flexible header lookup
    cols = {c.lower().strip(): c for c in raw.columns}
    get = lambda name: raw[cols[name]].astype(str) if name in cols else pd.Series("", index=raw.index)
    qnum = get("question number")
    cat  = get("c fastr category")
    pol  = get("polarity")
    spec = get("special interest") if "special interest" in cols else pd.Series("", index=raw.index)
    # Resolve
    out_rows, unresolved = [], []
    for i in range(len(raw)):
        q = qnum.iloc[i].strip()
        cdisp = cat.iloc[i].strip()
        pval  = pol.iloc[i].strip()
        skey = norm_cat_to_key(cdisp)
        gw   = parse_polarity_to_good_when(pval)
        if not q or not skey or not gw:
            unresolved.append({"row": i+1, "qn": q, "cat": cdisp, "polarity": pval, "reason": "missing-or-unknown"})
            continue
        field = map_question_to_column(q, survey_headers)
        if not field:
            unresolved.append({"row": i+1, "qn": q, "cat": cdisp, "polarity": pval, "reason": "no-matching-survey-column"})
            continue
        out_rows.append({
            "question_number": q,
            "field": field,
            "category_key": skey,
            "good_when": gw,
            "special_interest": spec.iloc[i]
        })
    res = pd.DataFrame(out_rows)
    log("MAPPING_FIELDS_RESOLVED", {
        "resolved": int(res.shape[0]),
        "unresolved": len(unresolved),
        "unresolved_sample": unresolved[:10]
    })
    return res, pd.DataFrame(unresolved)

# ----------------------
# Aggregation
# ----------------------
def compute_percent_goods(survey: pd.DataFrame, mapping: pd.DataFrame, by_col: str) -> Dict[str, List[Tuple[str, float]]]:
    """
    Return dict: category_key -> list of (group_label, pct_good), sorted by label.
    Aggregates across all mapped questions within a category.
    """
    needed_fields = sorted(set(mapping["field"].unique().tolist()))
    survey = ensure_numeric_1_5(survey, needed_fields)
    counts: Dict[str, Dict[str, List[int]]] = {cat: {} for cat in CATEGORY_KEYS_IN_ORDER}
    for _, m in mapping.iterrows():
        field = m["field"]; cat = m["category_key"]; gw = m["good_when"]
        if field not in survey.columns: 
            continue
        vals = survey[field]
        for idx, val in vals.items():
            if pd.isna(val): 
                continue
            group = str(survey.at[idx, by_col]) if by_col in survey.columns else "Unknown"
            if not group: group = "Unknown"
            is_good = (val <= 2.0) if gw == "low" else (val >= 4.0)
            g = counts[cat].setdefault(group, [0,0])
            g[1] += 1
            if is_good: g[0] += 1
    out: Dict[str, List[Tuple[str, float]]] = {}
    for cat in CATEGORY_KEYS_IN_ORDER:
        pairs = []
        for grp, (good, total) in counts[cat].items():
            if total > 0:
                pairs.append((grp, 100.0*good/total))
        pairs.sort(key=lambda x: x[0].lower())
        out[cat] = pairs
    return out

def compute_topline(survey: pd.DataFrame, mapping: pd.DataFrame) -> Dict[str, float]:
    needed_fields = sorted(set(mapping["field"].unique().tolist()))
    survey = ensure_numeric_1_5(survey, needed_fields)
    counts = {cat: [0,0] for cat in CATEGORY_KEYS_IN_ORDER}
    for _, m in mapping.iterrows():
        field = m["field"]; cat = m["category_key"]; gw = m["good_when"]
        if field not in survey.columns: 
            continue
        for val in survey[field].dropna().astype(float):
            is_good = (val <= 2.0) if gw == "low" else (val >= 4.0)
            counts[cat][1] += 1; counts[cat][0] += (1 if is_good else 0)
    return {cat: (100.0*c[0]/c[1] if c[1]>0 else 0.0) for cat, c in counts.items()}

# ----------------------
# Colors & charts
# ----------------------
FUNC_PALETTE  = list(plt.get_cmap("tab20").colors)
LEVEL_PALETTE = list(plt.get_cmap("Set3").colors)

def assign_colors(labels: List[str], palette: List[Tuple[float,float,float]]) -> Dict[str, Tuple[float,float,float]]:
    mapping = {}
    i = 0
    for lab in sorted(set([str(x) for x in labels]), key=lambda x: x.lower()):
        mapping[lab] = palette[i % len(palette)]
        i += 1
    return mapping

def save_bar_chart_vertical(title: str, items: List[Tuple[str, float]], color_map: Dict[str, Tuple[float,float,float]], out_path: Path):
    if not items: 
        return
    labels, values = zip(*items)
    colors = [color_map.get(str(l), (0.5,0.5,0.5)) for l in labels]
    fig, ax = plt.subplots(figsize=(max(6.5, len(labels)*0.9), 4.2))
    ax.bar(list(range(len(labels))), list(values), color=colors)
    ax.set_ylim(0, 100); ax.set_ylabel("% positive responses"); ax.set_title(title)
    ax.set_xticks(list(range(len(labels)))); ax.set_xticklabels(labels, rotation=20, ha="right")
    for s in ("top","right"): ax.spines[s].set_visible(False)
    fig.tight_layout()
    out_path.parent.mkdir(parents=True, exist_ok=True)
    fig.savefig(out_path, dpi=150)
    plt.close(fig)
    log("CHART_SAVED", {"title": title, "count": len(labels), "file": str(out_path)})

# ----------------------
# Copy threshold selection
# ----------------------
def pick_copy_for_category(thresholds: pd.DataFrame, category_display: str, pct_good: float) -> Tuple[str, str]:
    """
    Return (band_label, copy_text). Expected columns:
      - Category (or 'C FASTR Category')
      - Lower Threshold Value
      - Upper Threshold Value
      - Copy for Summary Section
    Picks the row where lower <= pct < upper. If none, falls back to nearest lower band.
    """
    if thresholds is None or thresholds.empty:
        return ("", "")
    cols = {c.lower().strip(): c for c in thresholds.columns}
    cat_col  = cols.get("category") or cols.get("c fastr category")
    low_col  = cols.get("lower threshold value")
    up_col   = cols.get("upper threshold value")
    copy_col = cols.get("copy for summary section")
    if not (cat_col and low_col and up_col and copy_col):
        return ("", "")
    mask = thresholds[cat_col].astype(str).str.strip().str.lower() == category_display.strip().lower()
    sub = thresholds[mask].copy()
    if sub.empty:
        return ("", "")
    sub[low_col] = pd.to_numeric(sub[low_col], errors="coerce")
    sub[up_col]  = pd.to_numeric(sub[up_col], errors="coerce")
    # find band
    row = None
    for _, r in sub.iterrows():
        lo = float(r[low_col]) if not pd.isna(r[low_col]) else -np.inf
        hi = float(r[up_col])  if not pd.isna(r[up_col])  else  np.inf
        if pct_good >= lo and pct_good < hi:
            row = r; break
    if row is None:
        sub = sub.sort_values(by=[low_col], ascending=False)
        row = sub.iloc[0]
    text = str(row[copy_col]) if row is not None else ""
    band = (
        f"{int(row[low_col])}-{int(row[up_col])}"
        if row is not None and not (pd.isna(row[low_col]) or pd.isna(row[up_col]))
        else ""
    )
    return (band, text)

# ----------------------
# DOCX generation
# ----------------------
def render_report(template_path: str,
                  out_path: Path,
                  topline: Dict[str, float],
                  by_func: Dict[str, List[Tuple[str, float]]],
                  by_level: Dict[str, List[Tuple[str, float]]],
                  thresholds_df: Optional[pd.DataFrame],
                  function_colors: Dict[str, Tuple[float,float,float]],
                  level_colors: Dict[str, Tuple[float,float,float]],
                  header_text: Dict[str, str]):
    tpl = DocxTemplate(template_path)
    ctx: Dict[str, Any] = {}
    chart_dir = out_path.parent / "charts"
    chart_dir.mkdir(parents=True, exist_ok=True)

    # Charts, topline %s, summary copy
    for cat_key in CATEGORY_KEYS_IN_ORDER:
        disp = next((d for d,k in CATEGORY_DISPLAY_TO_KEY.items() if k == cat_key), cat_key.title())
        func_items = by_func.get(cat_key, [])
        lvl_items  = by_level.get(cat_key, [])
        img_func = chart_dir / f"{cat_key}_by_function_chart.png"
        img_lvl  = chart_dir / f"{cat_key}_by_level_chart.png"
        save_bar_chart_vertical(f"{disp}: by Business Function", func_items, function_colors, img_func)
        save_bar_chart_vertical(f"{disp}: by Title",           lvl_items,  level_colors,    img_lvl)
        ctx[f"{cat_key}_by_function_chart"] = InlineImage(tpl, str(img_func), width=Mm(120))
        ctx[f"{cat_key}_by_level_chart"]    = InlineImage(tpl, str(img_lvl),  width=Mm(120))
        # topline % (optional)
        ctx[f"{cat_key}_pct"] = round(float(topline.get(cat_key, 0.0)), 1)
        # summary copy from thresholds
        _, text = pick_copy_for_category(thresholds_df, disp, ctx[f"{cat_key}_pct"])
        ctx[f"{cat_key}_summary_copy"]   = text or ""
        ctx[f"{cat_key}_good_copy"]      = text or ""
        ctx[f"{cat_key}_attention_copy"] = text or ""

    # Header/meta text placeholders (best-effort)
    ctx.update({
        "title_company_name": header_text.get("company_name", ""),
        "title_client_contact": header_text.get("company_contact", ""),
        "title_date": dt.date.today().strftime("%B %d, %Y"),
        "exec_engagement_summary": header_text.get("engagement_overview", ""),
        "exec_results_summary": header_text.get("results_summary", ""),
        "exec_suggested_actions": header_text.get("actions_summary", ""),
        "conclusion_overview": header_text.get("report_conclusion", ""),
        "generated_ts": _now(),
    })

    # Try to log undeclared variables (optional)
    try:
        present = set(tpl.get_undeclared_template_variables() or [])
        log("TEMPLATE_VARS_FOUND", {"count": len(present), "sample": sorted(list(present))[:40]})
    except Exception:
        pass

    out_path.parent.mkdir(parents=True, exist_ok=True)
    log("RENDER_BEGIN", {"template": str(template_path), "output": str(out_path)})
    tpl.render(ctx)
    tpl.save(str(out_path))
    log("REPORT_WRITTEN", {"output": str(out_path)})

# ----------------------
# Streamlit UI
# ----------------------
st.set_page_config(page_title=APP_TITLE, layout="centered")
st.title(APP_TITLE)

with st.expander("How this works", expanded=True):
    st.markdown("""
- Point to the three CSVs (Survey, Mapping, Thresholds) and the Word template.
- Computes **% positive responses** per C FASTR category by **Business Function** and **Title**.
- Colors are **categorical and consistent** across charts (not risk bands).
- Template placeholders expected (examples):
  `{{ collusion_by_function_chart }}`, `{{ collusion_by_level_chart }}`, `{{ collusion_summary_copy }}` (repeat for each category).
- Output: `out/cfastr_report.docx`.
""")

st.subheader("1) Consultant & Client Info")
col1, col2 = st.columns(2)
with col1:
    consultant_name = st.text_input("Consultant name *", "")
    company_name    = st.text_input("Company name *", "")
    company_contact = st.text_input("Company contact *", "")
with col2:
    engagement_overview = st.text_area("Engagement Overview *", height=100)
    results_summary     = st.text_area("Results Summary *", height=100)
    actions_summary     = st.text_area("Suggested Actions Summary *", height=100)
    report_conclusion   = st.text_area("Report Conclusion", height=100)

st.subheader("2) Data sources")
use_repo_defaults = st.checkbox("Use GitHub defaults (public repo)", value=True)
survey_path     = st.text_input("Survey CSV path or URL",     GH_SURVEY if use_repo_defaults else DEFAULT_SURVEY_CSV)
mapping_path    = st.text_input("Mapping CSV path or URL",    GH_MAPPING if use_repo_defaults else DEFAULT_MAPPING_CSV)
thresholds_path = st.text_input("Thresholds CSV path or URL", GH_THRESHOLDS if use_repo_defaults else DEFAULT_THRESHOLDS_CSV)
template_path   = st.text_input("Word template path", DEFAULT_TEMPLATE)

st.subheader("3) Columns")
function_col = st.text_input("Function column name", DEFAULT_FUNCTION_COL)
level_col    = st.text_input("Level column name (Title recommended)", DEFAULT_LEVEL_COL)

out_name = st.text_input("Output filename", "cfastr_report.docx")
go = st.button("Generate Report")

if go:
    try:
        # Load inputs
        survey = load_survey(survey_path)
        mapping_resolved, unresolved = load_mapping(mapping_path, survey.columns.tolist())
        if mapping_resolved.empty:
            st.error("No mapping rows resolved. Check 'Question Number', 'C FASTR Category', and 'Polarity' (1/-1 or positive/negative).")
            if not unresolved.empty:
                st.dataframe(unresolved.head(15))
            st.stop()
        thresholds_df = None
        try:
            thresholds_df = read_csv_auto(thresholds_path)
        except Exception as e:
            st.warning(f"Could not read thresholds CSV: {e}")

        # Compute aggregates
        topline     = compute_topline(survey, mapping_resolved)
        by_function = compute_percent_goods(survey, mapping_resolved, by_col=function_col)
        by_level    = compute_percent_goods(survey, mapping_resolved, by_col=level_col)

        # Colors consistent across the whole doc
        all_functions = sorted({g for cat in CATEGORY_KEYS_IN_ORDER for (g, _) in by_function.get(cat, [])})
        all_levels    = sorted({g for cat in CATEGORY_KEYS_IN_ORDER for (g, _) in by_level.get(cat, [])})
        function_colors = assign_colors(all_functions, FUNC_PALETTE)
        level_colors    = assign_colors(all_levels, LEVEL_PALETTE)

        # Render DOCX
        out_file = Path("out") / out_name
        tpath = Path(template_path).resolve()
        if not tpath.exists():
            st.error(f"Template not found at {tpath}. Place '{DEFAULT_TEMPLATE}' next to app.py or provide an absolute path.")
            st.stop()

        header_text = {
            "consultant_name": consultant_name,
            "company_name": company_name,
            "company_contact": company_contact,
            "engagement_overview": engagement_overview,
            "results_summary": results_summary,
            "actions_summary": actions_summary,
            "report_conclusion": report_conclusion,
        }

        render_report(
            str(tpath), out_file, topline, by_function, by_level,
            thresholds_df, function_colors, level_colors, header_text
        )

        st.success(f"Report written to: {out_file}")
        try:
            with open(out_file, "rb") as f:
                st.download_button(
                    "Download report",
                    f,
                    file_name=out_file.name,
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                )
        except Exception:
            pass

        # Show recent log tail
        try:
            with LOGFILE.open("r", encoding="utf-8") as f:
                lines = f.readlines()[-100:]
            st.text_area("Recent log output", value="".join(lines), height=360)
        except Exception:
            st.info("No log file yet.")
    except Exception as e:
        st.error(str(e))
        try:
            with LOGFILE.open("r", encoding="utf-8") as f:
                lines = f.readlines()[-100:]
            st.text_area("Recent log output", value="".join(lines), height=360)
        except Exception:
            st.info("No log file yet.")
