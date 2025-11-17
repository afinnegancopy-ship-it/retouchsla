import streamlit as st
import pandas as pd
import numpy as np
import datetime as dt
from io import BytesIO
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from openpyxl.styles import Alignment

# ----------------------------
# Configuration
# ----------------------------
COLS_TO_DELETE = [A, C, D, H, I, J, K, L, M, N, O, P, Q, S, X, AB, AC, AD, AE, AF, AG]
SLA_DAYS = {STILLS 2, MODEL 2, MANNEQUIN 2}

# ----------------------------
# Utilities
# ----------------------------
def excel_col_to_index(col)
    col = col.strip().upper()
    expn, num = 0, 0
    for char in reversed(col)
        num += (ord(char) - 64)  (26  expn)
        expn += 1
    return num - 1

def working_days_diff(start, end)
    if pd.isna(start) or pd.isna(end)
        return np.nan
    return np.busday_count(start, end)

# ----------------------------
# Streamlit UI
# ----------------------------
st.title(📊 Retouch SLA Checker)

uploaded_file = st.file_uploader(Upload your Excel file, type=[xls, xlsx])

today = st.date_input(Select today's date, dt.date.today())

if uploaded_file
    try
        df = pd.read_excel(uploaded_file)
        st.success(f✅ Loaded {len(df)} rows and {len(df.columns)} columns.)
    except Exception as e
        st.error(f❌ Failed to read Excel file {e})
        st.stop()

    # ----------------------------
    # Drop columns by Excel letter
    # ----------------------------
    indices = sorted({excel_col_to_index(l) for l in COLS_TO_DELETE})
    names_to_drop = [df.columns[i] for i in indices if i  len(df.columns)]
    df.drop(columns=names_to_drop, inplace=True, errors=ignore)

    # ----------------------------
    # Find Scan InOut Dates
    # ----------------------------
    scan_col = next((c for c in df.columns if scan in str(c).lower() and in in str(c).lower()), None)
    if not scan_col
        st.error(❌ Could not find a 'Scan In Date' column.)
        st.stop()
    df[scan_col] = pd.to_datetime(df[scan_col], errors=coerce)
    df = df[~df[scan_col].isna()]

    scan_out_col = next((c for c in df.columns if scan in str(c).lower() and out in str(c).lower()), None)
    if scan_out_col
        df[scan_out_col] = pd.to_datetime(df[scan_out_col], errors=coerce)

    # ----------------------------
    # Add new columns
    # ----------------------------
    new_cols = [
        Stills Out of SLA, Day(s) out of SLA - STILLS,
        Model Out of SLA, Day(s) out of SLA - MODEL,
        Mannequin Out of SLA, Day(s) out of SLA - MANNEQUIN,
        Notes, Days in Studio
    ]
    for col in new_cols
        df[col] = 

    # Convert date columns if exist
    date_cols = [c for c in df.columns if date in c.lower()]
    for c in date_cols
        df[c] = pd.to_datetime(df[c], errors=coerce)

    # ----------------------------
    # SLA logic
    # ----------------------------
    for prefix, (photo_col, upload_col) in {
        Stills (Photo Still Date, Still Upload Date),
        Model (Photo Model Date, Model Upload Date),
        Mannequin (Photo Mannequin Date, Mannequin Upload Date)
    }.items()
        if photo_col in df.columns
            for i, row in df.iterrows()
                start = row.get(photo_col)
                end = row.get(upload_col)
                if pd.isna(start)
                    continue
                days = working_days_diff(start.date(), today if pd.isna(end) else end.date())
                if pd.isna(end) and days  SLA_DAYS[prefix.upper()]
                    df.at[i, f{prefix} Out of SLA] = LATE
                    df.at[i, fDay(s) out of SLA - {prefix.upper()}] = days - SLA_DAYS[prefix.upper()]
                elif not pd.isna(end)
                    total = working_days_diff(start.date(), end.date())
                    if total  SLA_DAYS[prefix.upper()]
                        df.at[i, f{prefix} Out of SLA] = LATE
                        df.at[i, fDay(s) out of SLA - {prefix.upper()}] = total - SLA_DAYS[prefix.upper()]

    # Awaiting model shot
    if Photo Still Date in df.columns and scan_out_col
        for i, row in df.iterrows()
            still_date = row.get(Photo Still Date)
            scan_out = row.get(scan_out_col)
            if pd.isna(scan_out) and not pd.isna(still_date)
                diff = working_days_diff(still_date.date(), today)
                if diff  2
                    df.at[i, Notes] = Awaiting model shot

    # Days in Studio
    def compute_days_in_studio(row)
        scan_in = row.get(scan_col)
        scan_out = row.get(scan_out_col) if scan_out_col else None
        shot_upload_cols = [
            Photo Still Date, Photo Model Date, Photo Mannequin Date,
            Still Upload Date, Model Upload Date, Mannequin Upload Date
        ]
        all_blank = all(pd.isna(row.get(c)) for c in shot_upload_cols)
        if pd.notna(scan_in) and pd.notna(scan_out) and all_blank
            return SCANNED OUT AND NEVER SHOT
        elif pd.notna(scan_out)
            return SCANNED OUT
        elif pd.notna(scan_in)
            return working_days_diff(scan_in.date(), today)
        else
            return 
    df[Days in Studio] = df.apply(compute_days_in_studio, axis=1)

    # SLA Status Summary
    sla_cols = [Stills Out of SLA, Model Out of SLA, Mannequin Out of SLA]
    df[SLA status] = df[sla_cols].apply(lambda row LATE if LATE in row.values else , axis=1)

    # Move SLA status to column U (optional, not needed for display)
    
    # ----------------------------
    # Display and Download
    # ----------------------------
    st.subheader(Processed Data)
    st.dataframe(df)

    # Download button
    output = BytesIO()
    df.to_excel(output, index=False, engine='openpyxl')
    st.download_button(
        label=📥 Download Processed Excel,
        data=output.getvalue(),
        file_name=fcheck_retouch_processed_{today}.xlsx,
        mime=applicationvnd.openxmlformats-officedocument.spreadsheetml.sheet
    )
