# generate_client_report_PATCHED.py
# ---------------------------------
# CONSISTENT per-function colors across ALL charts + visible debug stamp.
# Column used for functions: "Business Function".

from __future__ import annotations

import io
import os
from typing import Dict, List, Sequence, Tuple

import matplotlib
matplotlib.use("Agg")  # headless backend for servers/Streamlit
import matplotlib.pyplot as plt
from matplotlib.colors import to_hex
import pandas as pd

# ========= VERSION / DEBUG =========
__version__ = "CF-ColorMap v0.3"
DEBUG_OVERLAY = os.getenv("CFASTR_DEBUG", "1") == "1"  # default ON so you can see it now
DEBUG_FILE = "CF_COLOR_DEBUG.txt"

# ========= CONFIG =========
FUNC_COL = "Business Function"

# Distinct, readable palette (repeats if you have more functions than colors)
PALETTE: List[str] = [
    "#1f77b4", "#ff7f0e", "#2ca02c", "#d62728", "#9467bd",
    "#8c564b", "#e377c2", "#7f7f7f", "#bcbd22", "#17becf",
    "#4e79a7", "#f28e2b", "#76b7b2", "#59a14f", "#edc949",
    "#af7aa1", "#ff9da7", "#9c755f", "#bab0ab", "#86bc86"
]

# ========= GLOBAL COLOR MAP (auto-filled) =========
FUNCTION_COLOR: Dict[str, str] = {}

def _norm_label(x: str) -> str:
    return str(x).strip()

def _log_debug_line(line: str) -> None:
    try:
        with open(DEBUG_FILE, "a", encoding="utf-8") as f:
            f.write(line.rstrip() + "\n")
    except Exception:
        pass
    try:
        print("[CF-DEBUG]", line)
    except Exception:
        pass

def _ensure_colors_for(labels: List[str]) -> None:
    """
    Assign colors to any function labels not yet in the global map (deterministic).
    Sorted for determinism; palette repeats if needed.
    """
    global FUNCTION_COLOR
    seen = {_norm_label(l) for l in labels if pd.notna(l)}
    new_labels = sorted([l for l in seen if l not in FUNCTION_COLOR], key=str.lower)
    if not new_labels:
        return

    start_idx = len(FUNCTION_COLOR)
    needed = start_idx + len(new_labels)
    reps = (needed + len(PALETTE) - 1) // len(PALETTE)
    palette = (PALETTE * max(1, reps))

    for i, lab in enumerate(new_labels):
        FUNCTION_COLOR[lab] = palette[start_idx + i]

    sample = ", ".join(f"{lab}→{FUNCTION_COLOR[lab]}" for lab in new_labels[:5])
    _log_debug_line(f"{__version__}: assigned colors for {len(new_labels)} new functions | {sample}")

def _prep_series_for_plot(
    df: pd.DataFrame,
    func_col: str,
    value_col: str,
    dropna: bool = True,
) -> Tuple[List[str], List[float]]:
    """Return (labels, values) sorted ascending by value (for horizontal bars)."""
    work = df.copy()
    if dropna:
        work = work[pd.notna(work[value_col])]
    work[value_col] = pd.to_numeric(work[value_col], errors="coerce")
    work = work.dropna(subset=[value_col])
    work = work.sort_values(value_col, ascending=True)
    labels = work[func_col].astype(str).tolist()
    values = work[value_col].astype(float).tolist()
    return labels, values

def _annotate_debug(ax, labels: List[str], colors: List[str]) -> None:
    if not DEBUG_OVERLAY:
        return
    uniq = []
    for l in labels:
        k = _norm_label(l)
        if k not in uniq:
            uniq.append(k)
    samples = []
    for name in uniq[:2]:
        col = FUNCTION_COLOR.get(name, "#6B7280")
        try:
            samples.append(f"{name}→{to_hex(col)}")
        except Exception:
            samples.append(f"{name}")
    text = f"{__version__} • {len(uniq)} funcs • " + (", ".join(samples) if samples else "")
    ax.text(
        0.995, -0.12, text,
        transform=ax.transAxes, ha="right", va="top",
        fontsize=7, color="#555555"
    )

def plot_hbar_by_function(
    df: pd.DataFrame,
    func_col: str,
    value_col: str,
    title: str,
    xlabel: str,
    figsize: Tuple[float, float] = (10.8, 3.2),
) -> io.BytesIO:
    """Horizontal bar chart with consistent per-function colors. Returns PNG bytes."""
    labels, values = _prep_series_for_plot(df, func_col, value_col)
    _ensure_colors_for(labels)
    colors = [FUNCTION_COLOR.get(_norm_label(lbl), "#6B7280") for lbl in labels]

    fig, ax = plt.subplots(figsize=figsize)
    ax.barh(labels, values, color=colors, edgecolor="none")
    ax.set_title(title, pad=8)
    ax.set_xlabel(xlabel)
    ax.set_ylabel("")
    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)
    ax.grid(axis="x", linestyle=":", linewidth=0.6, alpha=0.5)
    _annotate_debug(ax, labels, colors)

    fig.tight_layout()
    buf = io.BytesIO()
    fig.savefig(buf, format="png", dpi=160, bbox_inches="tight")
    plt.close(fig)
    buf.seek(0)
    return buf

# ========= SECTION WRAPPERS =========

def chart_relationship_focus_by_function(df_relationship: pd.DataFrame) -> io.BytesIO:
    return plot_hbar_by_function(
        df_relationship,
        func_col=FUNC_COL,
        value_col="avg_score",  # 1–5
        title="Relationship Focus: Avg Score by Business Function",
        xlabel="Average Score (1–5)",
    )

def chart_collusion_by_function(df_collusion: pd.DataFrame) -> io.BytesIO:
    return plot_hbar_by_function(
        df_collusion,
        func_col=FUNC_COL,
        value_col="percent_positive",  # 0–100
        title="Collusion: % Positive by Business Function",
        xlabel="Percent Positive",
    )

def chart_trust_by_function(df_trust: pd.DataFrame) -> io.BytesIO:
    return plot_hbar_by_function(
        df_trust,
        func_col=FUNC_COL,
        value_col="percent_positive",
        title="Trust: % Positive by Business Function",
        xlabel="Percent Positive",
    )

def chart_accountability_by_function(df_accountability: pd.DataFrame) -> io.BytesIO:
    return plot_hbar_by_function(
        df_accountability,
        func_col=FUNC_COL,
        value_col="percent_positive",
        title="Accountability: % Positive by Business Function",
        xlabel="Percent Positive",
    )

def chart_feedback_giving_by_function(df_feedback_giving: pd.DataFrame) -> io.BytesIO:
    return plot_hbar_by_function(
        df_feedback_giving,
        func_col=FUNC_COL,
        value_col="percent_positive",
        title="Feedback (Giving): % Positive by Business Function",
        xlabel="Percent Positive",
    )

def chart_feedback_receiving_by_function(df_feedback_receiving: pd.DataFrame) -> io.BytesIO:
    return plot_hbar_by_function(
        df_feedback_receiving,
        func_col=FUNC_COL,
        value_col="percent_positive",
        title="Feedback (Receiving): % Positive by Business Function",
        xlabel="Percent Positive",
    )

def chart_sensitivity_by_function(df_sensitivity: pd.DataFrame) -> io.BytesIO:
    return plot_hbar_by_function(
        df_sensitivity,
        func_col=FUNC_COL,
        value_col="percent_positive",
        title="Sensitivity: % Positive by Business Function",
        xlabel="Percent Positive",
    )

def chart_generic_by_function(
    df: pd.DataFrame,
    value_col: str,
    title: str,
    xlabel: str,
    width: float = 10.8,
    height: float = 3.2,
) -> io.BytesIO:
    return plot_hbar_by_function(
        df,
        func_col=FUNC_COL,
        value_col=value_col,
        title=title,
        xlabel=xlabel,
        figsize=(width, height),
    )
