import streamlit as st
import pandas as pd
import re
from datetime import datetime

st.set_page_config(page_title="JJMUP Report Generator", layout="centered")
st.title("💧 JJMUP Water Supply Report Generator")

st.markdown("Upload your JJMUP Excel file (.xls or .xlsx) to generate the daily report.")

uploaded_file = st.file_uploader("Choose Excel file", type=["xls", "xlsx"])

if uploaded_file:
    try:
        # Read Excel file
        df = pd.read_excel(uploaded_file)

        # Flatten MultiIndex columns if present
        if isinstance(df.columns, pd.MultiIndex):
            df.columns = [" ".join([str(x) for x in tup if pd.notna(x)]).strip() for tup in df.columns]

        # Normalize column names
        norm = {c: re.sub(r"\s+", "", str(c)).strip().lower() for c in df.columns}
        def find_col_contains(*needles: str) -> str:
            for c, cn in norm.items():
                if all(n in cn for n in needles):
                    return c
            raise KeyError(f"Missing column with fragments: {needles}")

        # Identify required columns
        scheme_id_col = find_col_contains("schemeid")
        scheme_name_col = find_col_contains("schemename")
        daily_demand_col = find_col_contains("waterdemand", "meter3", "daily")
        yest_prod_col = [c for c, cn in norm.items() if "oht" in cn and "watersupply" in cn and "meter3" in cn and "yesterday" in cn][0]
        today_prod_col = find_col_contains("today", "waterproduction", "meter3")

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

