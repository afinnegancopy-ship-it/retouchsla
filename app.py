import streamlit as st
import pandas as pd
import numpy as np
import datetime as dt
import os
import time
import win32com.client as win32
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
# Robust Date Parsing
# ----------------------------
UK_DATE_FORMATS = [
    "%d/%m/%Y", "%d/%m/%y",
    "%d-%m-%Y", "%d-%m-%y",
    "%d.%m.%Y", "%d.%m.%y",
    "%d %b %Y", "%d %B %Y",
    "%Y-%m-%d",
]

def parse_date_uk(value):
    if pd.isna(value):
        return pd.NaT
    if isinstance(value, (dt.datetime, dt.date)):
        return pd.to_datetime(value)
    text = str(value).strip()
    try:
        return pd.to_datetime(text, dayfirst=True, errors="raise")
    except:
        pass
    for fmt in UK_DATE_FORMATS:
        try:
            return dt.datetime.strptime(text, fmt)
        except:
            continue
    if " " in text:
        try:
            date_part = text.split(" ")[0]
            return pd.to_datetime(date_part, dayfirst=True, errors="raise")
        except:
            pass
    return pd.NaT

# ----------------------------
# Excel Reader
# ----------------------------
def read_excel_safely(path):
    try:
        return pd.read_excel(path)
    except:
        excel = win32.Dispatch("Excel.Application")
        excel.DisplayAlerts = False
        excel.Visible = False
        wb = excel.Workbooks.Open(path)
        new_path = os.path.splitext(path)[0] + "_converted.xlsx"
        if os.path.exists(new_path):
            os.remove(new_path)
            time.sleep(1)
        wb.SaveAs(new_path, FileFormat=51)
        wb.Close(SaveChanges=False)
        excel.Quit()
        return pd.read_excel(new_path)

# ----------------------------
# Streamlit App
# ----------------------------
st.title("Retouch SLA Checker")
uploaded = st.file_uploader("Upload Excel File", type=["xlsx", "xls"])
today = st.date_input("Today's Date", dt.date.today())

if uploaded and st.button("Process File"):
    df = pd.read_excel(uploaded)

    indices = sorted({excel_col_to_index(l) for l in COLS_TO_DELETE})
    names_to_drop = [df.columns[i] for i in indices if i < len(df.columns)]
    df.drop(columns=names_to_drop, inplace=True, errors="ignore")

    scan_col = next((c for c in df.columns if "scan" in c.lower() and "in" in c.lower()), None)
    scan_out_col = next((c for c in df.columns if "scan" in c.lower() and "out" in c.lower()), None)

    df[scan_col] = df[scan_col].apply(parse_date_uk)
    df = df[~df[scan_col].isna()]

    if scan_out_col:
        df[scan_out_col] = df[scan_out_col].apply(parse_date_uk)

    new_cols = [
        "Stills Out of SLA", "Day(s) out of SLA - STILLS",
        "Model Out of SLA", "Day(s) out of SLA - MODEL",
        "Mannequin Out of SLA", "Day(s) out of SLA - MANNEQUIN",
        "Notes", "Days in Studio"
    ]
    for col in new_cols:
        df[col] = ""

    for c in df.columns:
        if "date" in c.lower():
            df[c] = df[c].apply(parse_date_uk)

    sla_map = {
        "Stills": ("Photo Still Date", "Still Upload Date"),
        "Model": ("Photo Model Date", "Model Upload Date"),
        "Mannequin": ("Photo Mannequin Date", "Mannequin Upload Date"),
    }

    for prefix, (photo_col, upload_col) in sla_map.items():
        if photo_col not in df.columns:
            continue
        for i, row in df.iterrows():
            start = row.get(photo_col)
            end = row.get(upload_col)
            if pd.isna(start):
                continue
            if pd.isna(end):
                days = working_days_diff(start.date(), today)
                if days > SLA_DAYS[prefix.upper()]:
                    df.at[i, f"{prefix} Out of SLA"] = "LATE"
                    df.at[i, f"Day(s) out of SLA - {prefix.upper()}"] = days - SLA_DAYS[prefix.upper()]
            else:
                total = working_days_diff(start.date(), end.date())
                if total > SLA_DAYS[prefix.upper()]:
                    df.at[i, f"{prefix} Out of SLA"] = "LATE"
                    df.at[i, f"Day(s) out of SLA - {prefix.upper()}"] = total - SLA_DAYS[prefix.upper()]

    if "Photo Still Date" in df.columns and scan_out_col:
        for i, row in df.iterrows():
            still = row.get("Photo Still Date")
            out = row.get(scan_out_col)
            if pd.isna(out) and pd.notna(still):
                diff = working_days_diff(still.date(), today)
                if diff > 2:
                    df.at[i, "Notes"] = "Awaiting model shot"

    def compute_days(row):
        scan_in = row.get(scan_col)
        scan_out = row.get(scan_out_col)
        shot_cols = ["Photo Still Date", "Photo Model Date", "Photo Mannequin Date",
                     "Still Upload Date", "Model Upload Date", "Mannequin Upload Date"]
        if pd.notna(scan_in) and pd.notna(scan_out) and all(pd.isna(row.get(c)) for c in shot_cols):
            return "SCANNED OUT AND NEVER SHOT"
        elif pd.notna(scan_out):
            return "SCANNED OUT"
        elif pd.notna(scan_in):
            return working_days_diff(scan_in.date(), today)
        return ""

    df["Days in Studio"] = df.apply(compute_days, axis=1)

    sla_cols = ["Stills Out of SLA", "Model Out of SLA", "Mannequin Out of SLA"]
    df["SLA status"] = df[sla_cols].apply(lambda r: "LATE" if "LATE" in r.values else "", axis=1)

    st.success("Processing complete!")
    st.dataframe(df)

    output = df.to_excel(index=False)
    st.download_button("Download Processed File", data=output, file_name="processed.xlsx")
