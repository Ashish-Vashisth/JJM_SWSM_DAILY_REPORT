import streamlit as st
import pandas as pd
import re
from datetime import datetime

st.title("JJMUP Water Supply Report Generator")

uploaded_file = st.file_uploader("Upload JJMUP Excel file", type=["xls", "xlsx"])
if uploaded_file:
    def read_source(src_path) -> pd.DataFrame:
        try:
            return pd.read_excel(src_path, engine="openpyxl")
        except Exception:
            pass
        try:
            return pd.read_excel(src_path, engine="xlrd")
        except Exception:
            pass
        html = uploaded_file.read().decode("utf-8", errors="ignore")
        tables = pd.read_html(html)
        return max(tables, key=lambda t: t.shape[0])

    df = read_source(uploaded_file)

    if isinstance(df.columns, pd.MultiIndex):
        df.columns = [" ".join([str(x) for x in tup if pd.notna(x)]).strip() for tup in df.columns]

    norm = {c: re.sub(r"\s+", "", str(c)).strip().lower() for c in df.columns}
    def find_col_contains(*needles: str) -> str:
        for c, cn in norm.items():
            if all(n in cn for n in needles):
                return c
        raise KeyError(f"Missing column with fragments: {needles}")

    scheme_id_col = find_col_contains("schemeid")
    scheme_name_col = find_col_contains("schemename")
    daily_demand_col = find_col_contains("waterdemand","meter3","daily")
    yest_prod_col = [c for c,cn in norm.items() if "oht" in cn and "watersupply" in cn and "meter3" in cn and "yesterday" in cn][0]
    today_prod_col = find_col_contains("today","waterproduction","meter3")

    work_df = df[[scheme_id_col, scheme_name_col, daily_demand_col, yest_prod_col, today_prod_col]].copy()
    work_df.columns = ["Scheme Id","Scheme Name","Daily Water Demand (m^3)","Yesterday Water Production (m^3)","Today Water Production (m^3)"]

    for c in ["Daily Water Demand (m^3)","Yesterday Water Production (m^3)","Today Water Production (m^3)"]:
        work_df[c] = pd.to_numeric(work_df[c], errors="coerce")

    work_df["Percentage"] = (work_df["Yesterday Water Production (m^3)"] / work_df["Daily Water Demand (m^3)"]) * 100

    # Sheet 1: <75%
    less75_df = work_df[work_df["Percentage"].fillna(0) < 75].copy()
    less75_df["Supplied Water Percentage"] = "<75%"
    less75_df.insert(0,"SR.No.",range(1,len(less75_df)+1))
    less75_df = less75_df.drop(columns=["Today Water Production (m^3)"])

    # Sheet 2: ZERO/INACTIVE
    zero_df = work_df[(work_df["Yesterday Water Production (m^3)"].fillna(0)==0)&(work_df["Today Water Production (m^3)"].fillna(0)==0)].copy()
    zero_df = zero_df[["Scheme Id","Scheme Name","Yesterday Water Production (m^3)","Today Water Production (m^3)"]]
    zero_df["Site Status"] = "ZERO/INACTIVE SITE"
    zero_df.insert(0,"SR.No.",range(1,len(zero_df)+1))

    out_name = f"ZERO & LESS THAN 75 SITES {datetime.now().strftime('%Y-%m-%d')}.xlsx"
    with pd.ExcelWriter(out_name, engine="openpyxl") as w:
        less75_df.to_excel(w, sheet_name="SUPPLIED WATER LESS THAN 75", index=False)
        zero_df.to_excel(w, sheet_name="ZERO(INACTIVE SITES)", index=False)

    with open(out_name, "rb") as f:
        st.download_button("Download Processed Excel", f, file_name=out_name)
