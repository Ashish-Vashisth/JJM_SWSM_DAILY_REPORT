import streamlit as st
import pandas as pd
import re
from datetime import datetime
import os

st.set_page_config(page_title="JJMUP Report Generator", layout="centered")
st.title("💧 JJMUP Water Supply Report Generator")

st.markdown("Upload your JJMUP Excel file (.xls or .xlsx) to generate the daily water supply report.")

uploaded_file = st.file_uploader("Choose Excel file", type=["xls", "xlsx"])

if uploaded_file:
    try:
        # Detect file extension
        file_extension = os.path.splitext(uploaded_file.name)[1].lower()

        # Try reading Excel normally
        if file_extension == ".xlsx":
            df = pd.read_excel(uploaded_file, engine="openpyxl")
        elif file_extension == ".xls":
            try:
                df = pd.read_excel(uploaded_file, engine="xlrd")
            except Exception:
                content = uploaded_file.read().decode("utf-8", errors="ignore")
                tables = pd.read_html(content)
                df = max(tables, key=lambda t: t.shape[0])
        else:
            raise ValueError("Unsupported file format")

        # Flatten MultiIndex columns if present
        if isinstance(df.columns, pd.MultiIndex):
            df.columns = [" ".join([str(x) for x in tup if pd.notna(x)]).strip() for tup in df.columns]

        # Show available columns for debugging
        st.write("📋 Available columns:", list(df.columns))

        # Normalize column names
        norm = {c: re.sub(r"\s+", "", str(c)).strip().lower() for c in df.columns}

        def find_col_contains_any(*needles: str) -> str:
            for c, cn in norm.items():
                if any(n in cn for n in needles):
                    return c
            raise KeyError(f"Missing column with any of: {needles}")

        scheme_id_col = find_col_contains_any("schemeid")
        scheme_name_col = find_col_contains_any("schemename")
        daily_demand_col = find_col_contains_any("waterdemand", "meter3", "daily", "demand")
        yest_prod_col = find_col_contains_any("oht", "watersupply", "meter3", "yesterday", "supply")
        today_prod_col = find_col_contains_any("today", "waterproduction", "meter3", "production")

        # Build working DataFrame
        work_df = df[[scheme_id_col, scheme_name_col, daily_demand_col, yest_prod_col, today_prod_col]].copy()
        work_df.columns = [
            "Scheme Id",
            "Scheme Name",
            "Daily Water Demand (m^3)",
            "Yesterday Water Production (m^3)",
            "Today Water Production (m^3)"
        ]

        # Coerce numerics
        for c in ["Daily Water Demand (m^3)", "Yesterday Water Production (m^3)", "Today Water Production (m^3)"]:
            work_df[c] = pd.to_numeric(work_df[c], errors="coerce")

        # Calculate percentage
        work_df["Percentage"] = (work_df["Yesterday Water Production (m^3)"] / work_df["Daily Water Demand (m^3)"]) * 100

        # Sheet 1: <75%
        less75_df = work_df[work_df["Percentage"].fillna(0) < 75].copy()
        less75_df["Supplied Water Percentage"] = "<75%"
        less75_df.insert(0, "SR.No.", range(1, len(less75_df) + 1))
        less75_df = less75_df.drop(columns=["Today Water Production (m^3)"])

        # Sheet 2: ZERO/INACTIVE
        zero_df = work_df[
            (work_df["Yesterday Water Production (m^3)"].fillna(0) == 0) &
            (work_df["Today Water Production (m^3)"].fillna(0) == 0)
        ].copy()
        zero_df = zero_df[["Scheme Id", "Scheme Name", "Yesterday Water Production (m^3)", "Today Water Production (m^3)"]]
        zero_df["Site Status"] = "ZERO/INACTIVE SITE"
        zero_df.insert(0, "SR.No.", range(1, len(zero_df) + 1))

        # Save Excel
        out_name = f"ZERO & LESS THAN 75 SITES {datetime.now().strftime('%Y-%m-%d')}.xlsx"
        with pd.ExcelWriter(out_name, engine="openpyxl") as w:
            less75_df.to_excel(w, sheet_name="SUPPLIED WATER LESS THAN 75", index=False)
            zero_df.to_excel(w, sheet_name="ZERO(INACTIVE SITES)", index=False)

        # Download button
        with open(out_name, "rb") as f:
            st.success("✅ Report generated successfully!")
            st.download_button("📥 Download Processed Excel", f, file_name=out_name)

    except Exception as e:
        st.error(f"❌ Error processing file: {e}")
