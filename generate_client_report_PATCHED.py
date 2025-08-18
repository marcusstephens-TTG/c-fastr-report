# generate_client_report_PATCHED.py
# ---------------------------------
# Generates charts/tables for the C FASTR report with CONSISTENT per-function colors.
# Each Business Function gets one color, reused across every chart in this run.
# No external init calls required.

from __future__ import annotations

import io
from typing import Dict, Iterable, List, Sequence, Tuple

import matplotlib
matplotlib.use("Agg")  # headless backend for Streamlit/servers
import matplotlib.pyplot as plt
import pandas as pd

# ========= CONFIG =========
# Column that holds business function labels across your dataframes:
FUNC_COL = "Business Function"

# A long, readable palette; repeated if you have more functions than colors here.
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

def _ensure_colors_for(labels: List[str]) -> None:
    """Assign colors to any function labels not yet in the global map (deterministic)."""
    global FUNCTION_COLOR
    # Normalize and filter
    seen = {_norm_label(l) for l in labels if pd.notna(l)}
    # Add only the ones not mapped yet, in sorted order for determinism
    new_labels = sorted([l for l in seen if l not in FUNCTION_COLOR], key=str.lower)
    if not new_labels:
        return
    # Extend palette if needed
    start_idx = len(FUNCTION_COLOR)
    needed = start_idx + len(new_labels)
    reps = (needed + len(PALETTE) - 1) // len(PALETTE)
    palette = (PALETTE * max(1, reps))
    for i, lab in enumerate(new_labels):
        FUNCTION_COLOR[lab] = palette[start_idx + i]

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

def plot_hbar_by_function(
    df: pd.DataFrame,
    func_col: str,
    value_col: str,
    title: str,
    xlabel: str,
    figsize: Tuple[float, float] = (10.0, 3.0),
) -> io.BytesIO:
    """
    Horizontal bar chart with consistent per-function colors.
    Returns a PNG image in-memory (BytesIO).
    """
    labels, values = _prep_series_for_plot(df, func_col, value_col)

    # Ensure we have colors for every label seen anywhere so far
    _ensure_colors_for(labels)
    colors = [FUNCTION_COLOR.get(_norm_label(lbl), "#6B7280") for lbl in labels]

    fig, ax = plt.subplots(figsize=figsize)
    ax.barh(labels, values, color=colors, edgecolor="none")
    ax.set_title(title, pad=8)
    ax.set_xlabel(xlabel)
    ax.set_ylabel("")  # cleaner
    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)
    ax.grid(axis="x", linestyle=":", linewidth=0.6, alpha=0.5)

    fig.tight_layout()
    buf = io.BytesIO()
    fig.savefig(buf, format="png", dpi=200, bbox_inches="tight")
    plt.close(fig)
    buf.seek(0)
    return buf

# ========= WRAPPERS FOR YOUR SECTIONS =========
# Adjust the value column names if yours differ.

def chart_relationship_focus_by_function(df_relationship: pd.DataFrame) -> io.BytesIO:
    """
    Expects columns:
      - "Business Function"
      - "avg_score" (1–5)
    """
    return plot_hbar_by_function(
        df_relationship,
        func_col=FUNC_COL,
        value_col="avg_score",
        title="Relationship Focus: Avg Score by Business Function",
        xlabel="Average Score (1–5)",
        figsize=(10.5, 3.2),
    )

def chart_collusion_by_function(df_collusion: pd.DataFrame) -> io.BytesIO:
    """
    Expects columns:
      - "Business Function"
      - "percent_positive" (0–100)
    """
    return plot_hbar_by_function(
        df_collusion,
        func_col=FUNC_COL,
        value_col="percent_positive",
        title="Collusion: % Positive by Business Function",
        xlabel="Percent Positive",
        figsize=(10.5, 3.2),
    )

def chart_trust_by_function(df_trust: pd.DataFrame) -> io.BytesIO:
    return plot_hbar_by_function(
        df_trust,
        func_col=FUNC_COL,
        value_col="percent_positive",
        title="Trust: % Positive by Business Function",
        xlabel="Percent Positive",
        figsize=(10.5, 3.2),
    )

def chart_accountability_by_function(df_accountability: pd.DataFrame) -> io.BytesIO:
    return plot_hbar_by_function(
        df_accountability,
        func_col=FUNC_COL,
        value_col="percent_positive",
        title="Accountability: % Positive by Business Function",
        xlabel="Percent Positive",
        figsize=(10.5, 3.2),
    )

def chart_feedback_giving_by_function(df_feedback_giving: pd.DataFrame) -> io.BytesIO:
    return plot_hbar_by_function(
        df_feedback_giving,
        func_col=FUNC_COL,
        value_col="percent_positive",
        title="Feedback (Giving): % Positive by Business Function",
        xlabel="Percent Positive",
        figsize=(10.5, 3.2),
    )

def chart_feedback_receiving_by_function(df_feedback_receiving: pd.DataFrame) -> io.BytesIO:
    return plot_hbar_by_function(
        df_feedback_receiving,
        func_col=FUNC_COL,
        value_col="percent_positive",
        title="Feedback (Receiving): % Positive by Business Function",
        xlabel="Percent Positive",
        figsize=(10.5, 3.2),
    )

def chart_sensitivity_by_function(df_sensitivity: pd.DataFrame) -> io.BytesIO:
    return plot_hbar_by_function(
        df_sensitivity,
        func_col=FUNC_COL,
        value_col="percent_positive",
        title="Sensitivity: % Positive by Business Function",
        xlabel="Percent Positive",
        figsize=(10.5, 3.2),
    )

# Generic helper if you need a one-off with a different value column/title.
def chart_generic_by_function(
    df: pd.DataFrame,
    value_col: str,
    title: str,
    xlabel: str,
    width: float = 10.5,
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
