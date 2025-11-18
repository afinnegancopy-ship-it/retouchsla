import streamlit as st
import pandas as pd
import numpy as np
import datetime as dt
import tempfile
import os
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from openpyxl.styles import Alignment

# ----------------------------
# Configuration
# ----------------------------
COLS_TO_DELETE = ["A", "C", "D", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "S", "X", "AB", "AC", "AD", "AE", "AF", "AG"]
SLA_DAYS = {"STILLS": 2, "MODEL": 2, "MANNEQUIN": 2}

# ----------------------------
# Utilities
# ----------------------------
def excel_col_to_index(col):
    col = col.strip().upper()
    expn, num = 0, 0
    for char in reversed(col):
        num += (ord(char) - 64) * (26 ** expn)
        expn += 1
    return num - 1

def working_days_diff(start, end):
    if pd.isna(start) or pd.isna(end):
        return np.nan
    return np.busday_count(start, end)

# ----------------------------
# Streamlit App
# ----------------------------
st.title("Retouch SLA Checker")

uploaded_file = st.file_uploader("Upload your Excel file (.xls or .xlsx)", type=["xls", "xlsx"])
today_input = st.date_input("Today's date", dt.date.today())

if uploaded_file:
    # Save uploaded file to temporary location
    with tempfile.NamedTemporaryFile(delete=False, suffix=os.path.splitext(uploaded_file.name)[1]) as tmp_file:
        tmp_file.write(uploaded_file.read())
        tmp_file_path = tmp_file.name

    # Read Excel safely
    try:
        if tmp_file_path.endswith(".xls"):
            df = pd.read_excel(tmp_file_path, engine="xlrd")  # pandas 1.5.3 + xlrd 1.2.0 supports .xls
        else:
            df = pd.read_excel(tmp_file_path, engine="openpyxl")
        st.success(f"Loaded {len(df)} rows and {len(df.columns)} columns.")
    except Exception as e:
        st.error(f"Failed to read Excel file: {e}")
        st.stop()

    # Drop unwanted columns
    indices = sorted({excel_col_to_index(l) for l in COLS_TO_DELETE})
    names_to_drop = [df.columns[i] for i in indices if i < len(df.columns)]
    df.drop(columns=names_to_drop, inplace=True, errors="ignore")

    # Scan In Date
    scan_col = next((c for c in df.columns if "scan" in str(c).lower() and "in" in str(c).lower()), None)
    if not scan_col:
        st.error("Could not find a 'Scan In Date' column.")
        st.stop()
    df[scan_col] = pd.to_datetime(df[scan_col], errors="coerce")
    df = df[~df[scan_col].isna()]

    # Scan Out Date
    scan_out_col = next((c for c in df.columns if "scan" in str(c).lower() and "out" in str(c).lower()), None)
    if scan_out_col:
        df[scan_out_col] = pd.to_datetime(df[scan_out_col], errors="coerce")

    # Add new columns
    new_cols = [
        "Stills Out of SLA", "Day(s) out of SLA - STILLS",
        "Model Out of SLA", "Day(s) out of SLA - MODEL",
        "Mannequin Out of SLA", "Day(s) out of SLA - MANNEQUIN",
        "Notes", "Days in Studio"
    ]
    for col in new_cols:
        df[col] = ""

    # Convert date columns
    date_cols = [c for c in df.columns if "date" in c.lower()]
    for c in date_cols:
        df[c] = pd.to_datetime(df[c], errors="coerce")

    # SLA logic
    for prefix, (photo_col, upload_col) in {
        "Stills": ("Photo Still Date", "Still Upload Date"),
        "Model": ("Photo Model Date", "Model Upload Date"),
        "Mannequin": ("Photo Mannequin Date", "Mannequin Upload Date")
    }.items():
        if photo_col in df.columns:
            for i, row in df.iterrows():
                start = row.get(photo_col)
                end = row.get(upload_col)
                if pd.isna(start):
                    continue
                days = working_days_diff(start.date(), today_input if pd.isna(end) else end.date())
                if pd.isna(end) and days > SLA_DAYS[prefix.upper()]:
                    df.at[i, f"{prefix} Out of SLA"] = "LATE"
                    df.at[i, f"Day(s) out of SLA - {prefix.upper()}"] = days - SLA_DAYS[prefix.upper()]
                elif not pd.isna(end):
                    total = working_days_diff(start.date(), end.date())
                    if total > SLA_DAYS[prefix.upper()]:
                        df.at[i, f"{prefix} Out of SLA"] = "LATE"
                        df.at[i, f"Day(s) out of SLA - {prefix.upper()}"] = total - SLA_DAYS[prefix.upper()]

    # Notes: Awaiting model shot
    if "Photo Still Date" in df.columns and scan_out_col:
        for i, row in df.iterrows():
            still_date = row.get("Photo Still Date")
            scan_out = row.get(scan_out_col)
            if pd.isna(scan_out) and not pd.isna(still_date):
                diff = working_days_diff(still_date.date(), today_input)
                if diff > 2:
                    df.at[i, "Notes"] = "Awaiting model shot"

    # Days in Studio
    def compute_days_in_studio(row):
        scan_in = row.get(scan_col)
        scan_out = row.get(scan_out_col) if scan_out_col else None
        shot_upload_cols = [
            "Photo Still Date", "Photo Model Date", "Photo Mannequin Date",
            "Still Upload Date", "Model Upload Date", "Mannequin Upload Date"
        ]
        all_blank = all(pd.isna(row.get(c)) for c in shot_upload_cols)

        if pd.notna(scan_in) and pd.notna(scan_out) and all_blank:
            return "SCANNED OUT AND NEVER SHOT"
        elif pd.notna(scan_out):
            return "SCANNED OUT"
        elif pd.notna(scan_in):
            return working_days_diff(scan_in.date(), today_input)
        else:
            return ""

    df["Days in Studio"] = df.apply(compute_days_in_studio, axis=1)

    # SLA Status Summary
    sla_cols = ["Stills Out of SLA", "Model Out of SLA", "Mannequin Out of SLA"]
    df["SLA status"] = df[sla_cols].apply(lambda row: "LATE" if "LATE" in row.values else "", axis=1)

    # Move SLA status to column U (optional)
    cols = df.columns.tolist()
    if "SLA status" in cols:
        target_index = 20  # Column U
        cols.insert(target_index, cols.pop(cols.index("SLA status")))
        df = df[cols]

    # Save processed Excel
    output_path = os.path.join(tempfile.gettempdir(), "check_retouch_processed.xlsx")
    df.to_excel(output_path, index=False)

    # Excel formatting
    wb = load_workbook(output_path)
    ws = wb.active

    # Format date columns
    for idx, col in enumerate(df.columns, start=1):
        if "date" in col.lower():
            for row_cells in ws.iter_rows(min_row=2, min_col=idx, max_col=idx):
                for cell in row_cells:
                    cell.number_format = "mm/dd/yyyy"

    # Auto-fit column widths
    for col in ws.columns:
        max_length = 0
        col_letter = get_column_letter(col[0].column)
        for cell in col:
            try:
                cell_length = len(str(cell.value)) if cell.value is not None else 0
                if cell_length > max_length:
                    max_length = cell_length
            except:
                pass
        ws.column_dimensions[col_letter].width = max_length + 2

    # Left-align headers
    for cell in ws[1]:
        cell.alignment = Alignment(horizontal="left")

    wb.save(output_path)

    # Download link
    with open(output_path, "rb") as f:
        st.download_button(
            label="Download Processed Excel",
            data=f,
            file_name="check_retouch_processed.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    st.success("✅ Processing complete!")
