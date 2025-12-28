import streamlit as st
import pandas as pd
import re
from datetime import datetime
import os
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font, Border, Side, PatternFill
from openpyxl.utils import get_column_letter

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
                if all(n in cn for n in needles):
                    return c
            raise KeyError(f"Missing column with fragments: {needles}")

        # Identify required columns
        scheme_id_col = find_col_contains_any("schemeid")
        scheme_name_col = "Scheme Name" if "Scheme Name" in df.columns else find_col_contains_any("schemename")
        daily_demand_col = find_col_contains_any("waterdemand", "meter3", "daily", "demand")

        # Yesterday production — exact match
        yest_prod_col = next((c for c in df.columns if str(c).strip().lower() == "oht water supply (meter3)".lower()), None)
        if yest_prod_col is None:
            raise KeyError("Could not find column: 'OHT Water Supply (Meter3)'")

        today_prod_col = find_col_contains_any("today", "waterproduction", "meter3", "production")
        last_date_col = "Last Data Receive Date" if "Last Data Receive Date" in df.columns else find_col_contains_any("lastdatareceivedate")

        # Build working DataFrame
        work_df = df[[scheme_id_col, scheme_name_col, daily_demand_col, yest_prod_col, today_prod_col, last_date_col]].copy()
        work_df.columns = [
            "Scheme Id",
            "Scheme Name",
            "Daily Water Demand (m^3)",
            "Yesterday Water Production (m^3)",
            "Today Water Production (m^3)",
            "Last Data Receive Date"
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

        # Sheet 2: ZERO/INACTIVE — keep Scheme Name + Last Data Receive Date
        zero_df = work_df[
            (work_df["Yesterday Water Production (m^3)"].fillna(0) == 0) &
            (work_df["Today Water Production (m^3)"].fillna(0) == 0)
        ][["Scheme Id", "Scheme Name", "Yesterday Water Production (m^3)", "Today Water Production (m^3)", "Last Data Receive Date"]].copy()
        zero_df["Site Status"] = "ZERO/INACTIVE SITE"
        zero_df.insert(0, "SR.No.", range(1, len(zero_df) + 1))

        # Debug preview
        st.write("🔍 Zero & Inactive preview:", zero_df.head())

        # Save Excel
        out_name = f"ZERO & LESS THAN 75 SITES {datetime.now().strftime('%Y-%m-%d')}.xlsx"
        with pd.ExcelWriter(out_name, engine="openpyxl") as w:
            less75_df.to_excel(w, sheet_name="Supplied Water <75%", index=False)
            zero_df.to_excel(w, sheet_name="Zero & Inactive Sites", index=False)

        # Apply formatting
        wb = load_workbook(out_name)
        thin = Side(style="thin", color="000000")
        border_all = Border(left=thin, right=thin, top=thin, bottom=thin)
        align_center = Alignment(horizontal="center", vertical="center", wrap_text=False)
        align_left = Alignment(horizontal="left", vertical="center", wrap_text=False)
        header_font = Font(bold=True, color="000000")
        header_fill = PatternFill("solid", fgColor="5B9BD5")

        def format_sheet(ws):
            for cell in ws[1]:
                cell.font = header_font
                cell.fill = header_fill
                cell.alignment = align_center
                cell.border = border_all
            maxlen = {}
            for r in ws.iter_rows(min_row=2, max_row=ws.max_row, min_col=1, max_col=ws.max_column):
                for cell in r:
                    cell.border = border_all
                    if cell.column == 3:  # Scheme Name column -> left
                        cell.alignment = align_left
                    else:
                        cell.alignment = align_center
                    val = "" if cell.value is None else str(cell.value)
                    maxlen[cell.column] = max(maxlen.get(cell.column, 0), len(val))
            for c in range(1, ws.max_column + 1):
                header_val = str(ws.cell(row=1, column=c).value or "")
                maxlen[c] = max(maxlen.get(c, 0), len(header_val))
                width = maxlen.get(c, 0)
                ws.column_dimensions[get_column_letter(c)].width = max(10, min(60, int(width * 1.2) + 2))

        format_sheet(wb["Supplied Water <75%"])
        format_sheet(wb["Zero & Inactive Sites"])
        wb.save(out_name)

        # Download button
        with open(out_name, "rb") as f:
            st.success("✅ Report generated successfully!")
            st.download_button("📥 Download Processed Excel", f, file_name=out_name)

    except Exception as e:
        st.error(f"❌ Error processing file: {e}")
