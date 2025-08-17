#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
generate_cfastr_test_data.py

CFastR test data generator that respects your official mapping:
- Reads categories (and optional Polarity) from CFASTR_Category_Mapping_V1.csv
- Lets you choose row count, company name, and per-category profiles
  (terrific / mixed / poor) for: Collusion, Feedback, Accountability,
  Sensitivity, Trust, Relationship.
- Produces **CFastR_Survey_Test_Data.csv** by default with columns:
  ["Title", "Business Function", <every 'Question Number' from the mapping in order>]

Run as a Streamlit UI:
    streamlit run generate_cfastr_test_data.py

Run as a CLI:
    python generate_cfastr_test_data.py \
      --rows 200 \
      --company "Acme" \
      --collusion terrific \
      --feedback poor \
      --accountability mixed \
      --sensitivity terrific \
      --trust mixed \
      --relationship poor \
      --out CFastR_Survey_Test_Data.csv
"""

import os
import re
import random
import argparse
from dataclasses import dataclass
from typing import Dict, List

import pandas as pd

# ---- Files/columns ----
MAPPING_PATH = "CFASTR_Category_Mapping_V1.csv"
DEFAULT_OUT = "CFastR_Survey_Test_Data.csv"  # <- new default name

TITLE_COL = "Title"
FUNC_COL  = "Business Function"

# ---- Light demographic pools ----
TITLES = ["Executive", "Director", "Manager", "IC"]
FUNCTIONS = ["HR", "Finance", "IT", "Operations", "Sales", "Marketing", "Engineering", "Legal"]

BASE_CATS = ["Collusion", "Feedback", "Accountability", "Sensitivity", "Trust", "Relationship"]


# ---------------- Mapping & scoring helpers ----------------

def load_mapping(path: str = MAPPING_PATH) -> pd.DataFrame:
    """Load the official mapping CSV. Must include 'Question Number' and 'C FASTR Category'.
       Polarity is optional; defaults to +1 where missing."""
    if not os.path.exists(path):
        raise FileNotFoundError(f"Mapping file not found: {path}")

    df = pd.read_csv(path)
    # normalize header artifacts
    df.columns = [str(c).replace("\ufeff", "").strip() for c in df.columns]

    for col in ("Question Number", "C FASTR Category"):
        if col not in df.columns:
            raise KeyError(f"Mapping CSV must include '{col}'")

    if "Polarity" not in df.columns:
        df["Polarity"] = 1
    df["Polarity"] = pd.to_numeric(df["Polarity"], errors="coerce").fillna(1).astype(int)

    # normalize categories into the six base buckets where applicable
    def normalize_category(cat: str) -> str:
        c = (cat or "").strip().lower()
        c = re.sub(r"\s+", " ", c)

        # funnel common variants into the six bases
        synonyms = {
            "feedback giving": "feedback",
            "feedback receiving": "feedback",
            "relationship focus": "relationship",
        }
        base = synonyms.get(c, c)
        base = re.sub(r"[^a-z]+", "", base)

        mapping = {
            "collusion": "Collusion",
            "feedback": "Feedback",
            "accountability": "Accountability",
            "sensitivity": "Sensitivity",
            "trust": "Trust",
            "relationship": "Relationship",
            "overall": "Overall",
        }
        return mapping.get(base, cat)  # if it’s not a base, keep original

    df["_BaseCategory"] = df["C FASTR Category"].map(normalize_category)
    return df


def is_likert(label: str) -> bool:
    """Heuristic: Likert questions contain '(1-5)' in the mapping label."""
    return "(1-5)" in str(label)


def sample_score(profile: str, polarity: int) -> int:
    """
    Return a 1–5 Likert score based on profile & polarity.
      terrific: high adjusted score overall
          polarity +1 -> raw high (mostly 5/4)
          polarity -1 -> raw low  (mostly 1/2)
      poor:     low adjusted score overall
          polarity +1 -> raw low
          polarity -1 -> raw high
      mixed:    2–4 random
    """
    prof = (profile or "mixed").strip().lower()
    if prof == "terrific":
        if polarity >= 0:
            return random.choices([5, 4, 3], weights=[70, 25, 5], k=1)[0]
        else:
            return random.choices([1, 2, 3], weights=[70, 25, 5], k=1)[0]
    if prof == "poor":
        if polarity >= 0:
            return random.choices([1, 2, 3], weights=[70, 25, 5], k=1)[0]
        else:
            return random.choices([5, 4, 3], weights=[70, 25, 5], k=1)[0]
    # mixed
    return random.choice([2, 3, 4])


@dataclass
class Profiles:
    Collusion: str = "mixed"
    Feedback: str = "mixed"
    Accountability: str = "mixed"
    Sensitivity: str = "mixed"
    Trust: str = "mixed"
    Relationship: str = "mixed"

    def for_category(self, cat: str) -> str:
        if cat in ("Collusion", "Feedback", "Accountability", "Sensitivity", "Trust", "Relationship"):
            return getattr(self, cat)
        return "mixed"


# ---------------- Generation ----------------

def generate_rows(n_rows: int, company: str, mapping: pd.DataFrame, profiles: Profiles) -> pd.DataFrame:
    """Create a DataFrame with required demographics and all mapped question columns."""
    q_labels: List[str] = mapping["Question Number"].astype(str).tolist()
    rows: List[Dict[str, object]] = []

    for _ in range(n_rows):
        row: Dict[str, object] = {}
        row[TITLE_COL] = random.choice(TITLES)
        row[FUNC_COL]  = random.choice(FUNCTIONS)

        for _, m in mapping.iterrows():
            q_label = str(m["Question Number"])
            base_cat = str(m["_BaseCategory"])
            pol = int(m.get("Polarity", 1))

            if is_likert(q_label):
                prof = profiles.for_category(base_cat)
                row[q_label] = sample_score(prof, pol)
            else:
                # Non-Likert “Overall” style items — drop in light text
                # (your scoring code typically ignores non-numeric)
                if "Workplace Health Assessment" in q_label:
                    row[q_label] = f"{company or 'Company'} health: {random.choice(['Strong','Average','Needs focus'])}"
                elif "Strengths/Barriers" in q_label:
                    row[q_label] = random.choice([
                        "Strengths: collaboration, speed. Barriers: resourcing.",
                        "Strengths: trust, clarity. Barriers: silos.",
                        "Strengths: innovation. Barriers: feedback culture.",
                        "Strengths: accountability. Barriers: communication gaps.",
                    ])
                elif "Workplace Improvement" in q_label:
                    row[q_label] = "Top 3 improvements: 1) clarity 2) coaching 3) cross-team rituals"
                else:
                    row[q_label] = ""

        rows.append(row)

    df = pd.DataFrame(rows)
    df = df.reindex(columns=[TITLE_COL, FUNC_COL] + q_labels)  # demographics first, then questions
    return df


def save_csv(df: pd.DataFrame, out_path: str = DEFAULT_OUT) -> str:
    df.to_csv(out_path, index=False)
    return os.path.abspath(out_path)


# ---------------- CLI ----------------

def build_arg_parser() -> argparse.ArgumentParser:
    ap = argparse.ArgumentParser(description="Generate CFastR test data using the official mapping CSV.")
    ap.add_argument("--rows", type=int, default=100, help="Number of rows to generate (default 100)")
    ap.add_argument("--company", type=str, default="Acme Corp", help="Company name (used in some text fields)")
    for base in BASE_CATS:
        ap.add_argument(
            f"--{base.lower()}",
            type=str,
            choices=["terrific", "mixed", "poor"],
            default="mixed",
            help=f"{base} profile (terrific/mixed/poor)",
        )
    ap.add_argument("--out", type=str, default=DEFAULT_OUT, help=f"Output CSV path (default {DEFAULT_OUT})")
    return ap


def main_cli():
    args = build_arg_parser().parse_args()
    mapping = load_mapping(MAPPING_PATH)
    profs = Profiles(
        Collusion=args.collusion,
        Feedback=args.feedback,
        Accountability=args.accountability,
        Sensitivity=args.sensitivity,
        Trust=args.trust,
        Relationship=args.relationship,
    )
    df = generate_rows(args.rows, args.company, mapping, profs)
    out = save_csv(df, args.out)
    print(f"Wrote {out}")


# ---------------- Streamlit UI ----------------

def _run_streamlit():
    import streamlit as st  # local import so CLI doesn't need it

    st.set_page_config(page_title="CFastR Test Data Generator (Mapped)", layout="centered")
    st.title("CFastR Test Data Generator")
    st.caption("Generates CFastR_Survey_Test_Data.csv using the official category mapping file.")

    c1, c2 = st.columns(2)
    rows = c1.number_input("How many rows?", min_value=1, max_value=100_000, value=100, step=50)
    company = c2.text_input("Company Name", value="Acme Corp")

    st.markdown("#### Category Profiles (affect 1–5 items; reversed where mapping Polarity = −1)")
    col1, col2, col3 = st.columns(3)
    p_collusion      = col1.selectbox("Collusion",      ["terrific", "mixed", "poor"], index=1)
    p_feedback       = col2.selectbox("Feedback",       ["terrific", "mixed", "poor"], index=1)
    p_accountability = col3.selectbox("Accountability", ["terrific", "mixed", "poor"], index=1)
    p_sensitivity    = col1.selectbox("Sensitivity",    ["terrific", "mixed", "poor"], index=1)
    p_trust          = col2.selectbox("Trust",          ["terrific", "mixed", "poor"], index=1)
    p_relationship   = col3.selectbox("Relationship",   ["terrific", "mixed", "poor"], index=1)

    if st.button("🚀 Generate CSV", type="primary"):
        try:
            mapping = load_mapping(MAPPING_PATH)
        except Exception as e:
            st.error(f"Could not read mapping ({MAPPING_PATH}): {e}")
            return

        profs = Profiles(
            Collusion=p_collusion,
            Feedback=p_feedback,
            Accountability=p_accountability,
            Sensitivity=p_sensitivity,
            Trust=p_trust,
            Relationship=p_relationship,
        )

        with st.spinner("Generating data…"):
            df = generate_rows(int(rows), company, mapping, profs)

        st.success(f"Generated {len(df):,} rows. Preview below and download as CSV.")
        st.dataframe(df.head(20), use_container_width=True)
        st.download_button(
            "⬇️ Download CFastR_Survey_Test_Data.csv",
            data=df.to_csv(index=False).encode("utf-8"),
            file_name="CFastR_Survey_Test_Data.csv",  # <- new name in UI download
            mime="text/csv",
        )


# ---------------- Entry point ----------------

if __name__ == "__main__":
    # If launched via Streamlit, runtime usually sets STREAMLIT_RUNTIME=1
    if os.environ.get("STREAMLIT_RUNTIME") == "1" or "STREAMLIT_SERVER_ENABLED" in os.environ:
        _run_streamlit()
    else:
        main_cli()
