
import re
from io import BytesIO, StringIO
from datetime import datetime

import pandas as pd
import streamlit as st

from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font, Border, Side, PatternFill
from openpyxl.utils import get_column_letter


# ---------------------------
# Reading the uploaded file
# ---------------------------
def read_source(uploaded_file) -> pd.DataFrame:
    """
    Robust reader:
    1) Try Excel via openpyxl (works for .xlsx/.xlsm)
    2) If fails, try HTML-table fallback (common for JJMUP .xls exports that are HTML)
    """
    raw = uploaded_file.getvalue()

    # Try normal Excel first
    try:
        return pd.read_excel(BytesIO(raw), engine="openpyxl")
    except Exception:
        pass

    # HTML fallback (your uploaded .xls looks like HTML)  [1](https://voltasworld-my.sharepoint.com/personal/80002819_voltasworld_com/_layouts/15/Doc.aspx?sourcedoc=%7BD49513F1-E3AE-4C81-AD79-D83F0323FF10%7D&file=jjmup%20(32).xls&action=default&mobileredirect=true).xls&action=default&mobileredirect=true).xls&action=default&mobileredirect=true).xls&action=default&mobileredirect=true).xls&action=default&mobileredirect=true).xls&action=default&mobileredirect=true).xls&action=default&mobileredirect=true).xls&action=default&mobileredirect=true).xls&action=default&mobileredirect=true).xls&action=default&mobileredirect=true).xls&action=default&mobileredirect=true).xls&action=default&mobileredirect=true).xls&action=default&mobileredirect=true).xls&action=default&mobileredirect=true).xls&action=default&mobileredirect=true).xls&action=default&mobileredirect=true).xls&action=default&mobileredirect=true).xls&action=default&mobileredirect=true).xls&action=default&mobileredirect=true).xls&action=default&mobileredirect=true).xls&action=default&mobileredirect=true).xls&action=default&mobileredirect=true).xls&action=default&mobileredirect=true).xls&action=default&mobileredirect=true).xls&action=default&mobileredirect=true).xls&action=default&mobileredirect=true).xls&action=default&mobileredirect=true).xls&action=default&mobileredirect=true).xls&action=default&mobileredirect=true).xls&action=default&mobileredirect=true).xls&action=default&mobileredirect=true).xls&action=default&mobileredirect=true).xls&action=default&mobileredirect=true).xls&action=default&mobileredirect=true).xls&action=default&mobileredirect=true).xls&action=default&mobileredirect=true).xls&action=default&mobileredirect=true).xls&action=default&mobileredirect=true).xls&action=default&mobileredirect=true).xls&action=default&mobileredirect=true).xls&action=default&mobileredirect=true).xls&action=default&mobileredirect=true).xls&action=default&mobileredirect=true).xls&action=default&mobileredirect=true).xls&action=default&mobileredirect=true).xls&action=default&mobileredirect=true).xls&action=default&mobileredirect=true).xls&action=default&mobileredirect=true).xls&action=default&mobileredirect=true).xls&action=default&mobileredirect=true).xls&action=default&mobileredirect=true).xls&action=default&mobileredirect=true).xls&action=default&mobileredirect=true).xls&action=default&mobileredirect=true).xls&action=default&mobileredirect=true).xls&action=default&mobileredirect=true).xls&action=default&mobileredirect=true).xls&action=default&mobileredirect=true).xls&action=default&mobileredirect=true).xls&action=default&mobileredirect=true).xls&action=default&mobileredirect=true).xls&action=default&mobileredirect=true).xls&action=default&mobileredirect=true).xls&action=default&mobileredirect=true).xls&action=default&mobileredirect=true).xls&action=default&mobileredirect=true).xls&action=default&mobileredirect=true).xls&action=default&mobileredirect=true).xls&action=default&mobileredirect=true).xls&action=default&mobileredirect=true).xls&action=default&mobileredirect=true).xls&action=default&mobileredirect=true).xls&action=default&mobileredirect=true).xls&action=default&mobileredirect=true).xls&action=default&mobileredirect=true).xls&action=default&mobileredirect=true).xls&action=default&mobileredirect=true).xls&action=default&mobileredirect=true).xls&action=default&mobileredirect=true).xls&action=default&mobileredirect=true).xls&action=default&mobileredirect=true).xls&action=default&mobileredirect=true).xls&action=default&mobileredirect=true).xls&action=default&mobileredirect=true).xls&action=default&mobileredirect=true).xls&action=default&mobileredirect=true).xls&action=default&mobileredirect=true).xls&action=default&mobileredirect=true).xls&action=default&mobileredirect=true).xls&action=default&mobileredirect=true).xls&action=default&mobileredirect=true).xls&action=default&mobileredirect=true).xls&action=default&mobileredirect=true).xls&action=default&mobileredirect=true).xls&action=default&mobileredirect=true).xls&action=default&mobileredirect=true).xls&action=default&mobileredirect=true).xls&action=default&mobileredirect=true).xls&action=default&mobileredirect=true).xls&action=default&mobileredirect=true).xls&action=default&mobileredirect=true).xls&action=default&mobileredirect=true).xls&action=default&mobileredirect=true).xls&action=default&mobileredirect=true).xls&action=default&mobileredirect=true).xls&action=default&mobileredirect=true).xls&action=default&mobileredirect=true)
    html = raw.decode("utf-8", errors="ignore")
    tables = pd.read_html(StringIO(html))
    df = max(tables, key=lambda t: t.shape[0])  # largest table
    return df


def flatten_columns(df: pd.DataFrame) -> pd.DataFrame:
    if isinstance(df.columns, pd.MultiIndex):
        df.columns = [
            " ".join([str(x) for x in tup if pd.notna(x)]).strip()
            for tup in df.columns
        ]
    return df


def normalize_columns(df: pd.DataFrame) -> dict:
    return {c: re.sub(r"\s+", "", str(c)).strip().lower() for c in df.columns}


def find_col_contains(norm_map: dict, *needles: str) -> str:
    for c, cn in norm_map.items():
        if all(n in cn for n in needles):
            return c
    raise KeyError(f"Missing column with fragments: {needles}")


# ---------------------------
# Business logic (your report)
# ---------------------------
def build_report(df: pd.DataFrame, threshold: float = 75.0):
    df = flatten_columns(df)
    norm = normalize_columns(df)

    scheme_id_col = find_col_contains(norm, "schemeid")
    scheme_name_col = find_col_contains(norm, "schemename")
    daily_demand_col = find_col_contains(norm, "waterdemand", "meter3", "daily")

    # Yesterday production (more specific)
    yest_prod_col = None
    for c, cn in norm.items():
        if ("oht" in cn) and ("watersupply" in cn) and ("meter3" in cn) and ("yesterday" in cn):
            yest_prod_col = c
            break
    if yest_prod_col is None:
        raise KeyError("Could not find 'OHT Water Supply (Meter3) Yesterday' column.")

    today_prod_col = find_col_contains(norm, "today", "waterproduction", "meter3")
    last_date_col = find_col_contains(norm, "lastdatareceivedate")

    work_df = df[[scheme_id_col, scheme_name_col, daily_demand_col, yest_prod_col, today_prod_col]].copy()
    work_df.columns = [
        "Scheme Id",
        "Scheme Name",
        "Daily Water Demand (m^3)",
        "Yesterday Water Production (m^3)",
        "Today Water Production (m^3)",
    ]

    for c in ["Daily Water Demand (m^3)", "Yesterday Water Production (m^3)", "Today Water Production (m^3)"]:
        work_df[c] = pd.to_numeric(work_df[c], errors="coerce")

    work_df["Percentage"] = (work_df["Yesterday Water Production (m^3)"] / work_df["Daily Water Demand (m^3)"]) * 100

    # Sheet 1 (< threshold)
    less_df = work_df[work_df["Percentage"].fillna(0) < threshold].copy()
    less_df["Supplied Water Percentage"] = f"<{threshold:g}%"
    less_df.insert(0, "SR.No.", range(1, len(less_df) + 1))
    less_df = less_df.drop(columns=["Today Water Production (m^3)"])

    # Sheet 2 (ZERO/INACTIVE)
    zero_df = work_df[
        (work_df["Yesterday Water Production (m^3)"].fillna(0) == 0)
        & (work_df["Today Water Production (m^3)"].fillna(0) == 0)
    ].copy()

    zero_df = zero_df[["Scheme Id", "Scheme Name", "Yesterday Water Production (m^3)", "Today Water Production (m^3)"]]
    zero_df["Last Data Receive Date"] = df[last_date_col]
    zero_df["Site Status"] = "ZERO/INACTIVE SITE"
    zero_df.insert(0, "SR.No.", range(1, len(zero_df) + 1))

    return less_df, zero_df


# ---------------------------
# Excel writing + formatting
# ---------------------------
def apply_formatting(xlsx_bytes: bytes) -> bytes:
    wb = load_workbook(BytesIO(xlsx_bytes))

    thin = Side(style="thin", color="000000")
    border_all = Border(left=thin, right=thin, top=thin, bottom=thin)

    align_center = Alignment(horizontal="center", vertical="center", wrap_text=False)
    align_left = Alignment(horizontal="left", vertical="center", wrap_text=False)

    header_font = Font(bold=True, color="000000")
    header_fill = PatternFill("solid", fgColor="5B9BD5")

    def format_sheet(ws):
        # Header row
        for cell in ws[1]:
            cell.font = header_font
            cell.fill = header_fill
            cell.alignment = align_center
            cell.border = border_all

        # Body + compute widths
        maxlen = {}
        for r in ws.iter_rows(min_row=2, max_row=ws.max_row, min_col=1, max_col=ws.max_column):
            for cell in r:
                cell.border = border_all
                # Column 3: Scheme Name (SR.No. + Scheme Id + Scheme Name)
                if cell.column == 3:
                    cell.alignment = align_left
                else:
                    cell.alignment = align_center

                val = "" if cell.value is None else str(cell.value)
                maxlen[cell.column] = max(maxlen.get(cell.column, 0), len(val))

        # Auto width
        for c in range(1, ws.max_column + 1):
            header_val = str(ws.cell(row=1, column=c).value or "")
            maxlen[c] = max(maxlen.get(c, 0), len(header_val))
            width = maxlen.get(c, 0)
            ws.column_dimensions[get_column_letter(c)].width = max(10, min(60, int(width * 1.2) + 2))

    format_sheet(wb["SUPPLIED WATER LESS THAN 75"])
    format_sheet(wb["ZERO(INACTIVE SITES)"])

    out = BytesIO()
    wb.save(out)
    return out.getvalue()


def create_output_excel(less_df: pd.DataFrame, zero_df: pd.DataFrame) -> tuple[str, bytes]:
    date_str = datetime.now().strftime("%Y-%m-%d")
    out_name = f"ZERO & LESS THAN 75 SITES {date_str}.xlsx"

    buffer = BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as w:
        less_df.to_excel(w, sheet_name="SUPPLIED WATER LESS THAN 75", index=False)
        zero_df.to_excel(w, sheet_name="ZERO(INACTIVE SITES)", index=False)

    styled = apply_formatting(buffer.getvalue())
    return out_name, styled


# ---------------------------
# Streamlit UI
# ---------------------------
st.set_page_config(page_title="JJM SWSM Daily Report", layout="wide")
st.title("JJM SWSM Daily Report Generator")
st.write("Upload JJMUP export (.xls/.xlsx) → Download the formatted report Excel. [2](https://voltasworld-my.sharepoint.com/personal/80002819_voltasworld_com/_layouts/15/Doc.aspx?sourcedoc=%7BBA94A5E6-312D-47C8-85BB-AADB3673F10D%7D&file=ZERO%20%26%20LESS%20THAN%2075%20SITES%202025-12-28__.xlsx&action=default&mobileredirect=true)[1](https://voltasworld-my.sharepoint.com/personal/80002819_voltasworld_com/_layouts/15/Doc.aspx?sourcedoc=%7BD49513F1-E3AE-4C81-AD79-D83F0323FF10%7D&file=jjmup%20(32).xls&action=default&mobileredirect=true).xls&action=default&mobileredirect=true).xls&action=default&mobileredirect=true).xls&action=default&mobileredirect=true).xls&action=default&mobileredirect=true).xls&action=default&mobileredirect=true).xls&action=default&mobileredirect=true).xls&action=default&mobileredirect=true).xls&action=default&mobileredirect=true).xls&action=default&mobileredirect=true).xls&action=default&mobileredirect=true).xls&action=default&mobileredirect=true).xls&action=default&mobileredirect=true).xls&action=default&mobileredirect=true).xls&action=default&mobileredirect=true).xls&action=default&mobileredirect=true).xls&action=default&mobileredirect=true).xls&action=default&mobileredirect=true).xls&action=default&mobileredirect=true).xls&action=default&mobileredirect=true).xls&action=default&mobileredirect=true).xls&action=default&mobileredirect=true).xls&action=default&mobileredirect=true).xls&action=default&mobileredirect=true).xls&action=default&mobileredirect=true).xls&action=default&mobileredirect=true).xls&action=default&mobileredirect=true)")

threshold = st.number_input("Threshold (%) for LESS THAN list", min_value=1.0, max_value=100.0, value=75.0, step=1.0)

uploaded = st.file_uploader("Upload JJMUP file", type=["xls", "xlsx", "xlsm"])

if uploaded:
    st.info(f"Uploaded: {uploaded.name}")

    if st.button("Generate Report", type="primary"):
        try:
            df = read_source(uploaded)
            less_df, zero_df = build_report(df, threshold=threshold)
            out_name, out_bytes = create_output_excel(less_df, zero_df)

            st.success(f"Created: {out_name}")
            c1, c2 = st.columns(2)
            c1.metric("Rows < threshold", len(less_df))
            c2.metric("ZERO/INACTIVE rows", len(zero_df))

            with st.expander("Preview: SUPPLIED WATER LESS THAN 75"):
                st.dataframe(less_df, use_container_width=True)

            with st.expander("Preview: ZERO(INACTIVE SITES)"):
                st.dataframe(zero_df, use_container_width=True)

            st.download_button(
                "⬇️ Download Excel Report",
                data=out_bytes,
                file_name=out_name,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

        except Exception as e:
            st.error("Error while generating report. Please check the uploaded file format/columns.")
            st.exception(e)
else:
    st.warning("Please upload the JJMUP export file to proceed.")
