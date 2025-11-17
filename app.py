import streamlit as st
import pandas as pd
import numpy as np
import datetime as dt
from io import BytesIO
import pyexcel as p

# ----------------------------
# Configuration
# ----------------------------
COLS_TO_DELETE = ['A', 'C', 'D', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'S', 'X', 'AB', 'AC', 'AD', 'AE', 'AF', 'AG']
SLA_DAYS = {'STILLS': 2, 'MODEL': 2, 'MANNEQUIN': 2}


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
# Streamlit UI
# ----------------------------
st.title("📊 Retouch SLA Checker")

uploaded_file = st.file_uploader("Upload your Excel file", type=['xls', 'xlsx', 'csv'])
today = st.date_input("Select today's date", dt.date.today())

if uploaded_file:

    # --------------------------------------------
    # BULLET-PROOF FILE LOADER
    # --------------------------------------------
    file_ext = uploaded_file.name.split(".")[-1].lower()

    def load_excel_safely(uploaded_file, file_ext):
        """
        Attempts:
        1. openpyxl (xlsx)
        2. csv fallback
        3. pyexcel fallback for malformed xlsx
        4. xlrd for xls
        """
        if file_ext == "xlsx":
            # Try normal XLSX
            try:
                return pd.read_excel(uploaded_file, engine="openpyxl"), "xlsx"
            except Exception:
                pass

            # Try CSV fallback
            try:
                uploaded_file.seek(0)
                return pd.read_csv(uploaded_file), "csv-fallback"
            except Exception:
                pass

            # Try pyexcel final fallback
            try:
                uploaded_file.seek(0)
                book = p.get_book(file_type="xlsx", file_content=uploaded_file.read())
                sheet = book.sheet_by_index(0)
                df = pd.DataFrame(sheet.to_array())
                df.columns = df.iloc[0]
                df = df[1:]
                return df, "pyexcel"
            except Exception:
                pass

            raise Exception("Unrecoverable XLSX corruption")

        elif file_ext == "xls":
            return pd.read_excel(uploaded_file, engine="xlrd"), "xls"

        elif file_ext == "csv":
            return pd.read_csv(uploaded_file), "csv"

        else:
            raise Exception("Unsupported file type")

    try:
        df, load_method = load_excel_safely(uploaded_file, file_ext)

        if load_method == "pyexcel":
            st.warning("⚠️ File is corrupted XLSX but was successfully loaded using a recovery method.")
        elif load_method == "csv-fallback":
            st.warning("⚠️ File is not a real XLSX. Loaded as CSV.")

        st.success(f"✅ File loaded using: {load_method}")
        st.info(f"Rows: {len(df)} | Columns: {len(df.columns)}")

    except Exception as e:
        st.error(f"❌ Failed to read file: {e}")
        st.stop()

    # ----------------------------
    # Drop columns by Excel letter
    # ----------------------------
    indices = sorted({excel_col_to_index(l) for l in COLS_TO_DELETE})
    names_to_drop = [df.columns[i] for i in indices if i < len(df.columns)]
    df.drop(columns=names_to_drop, inplace=True, errors='ignore')

    # ----------------------------
    # Identify scan columns
    # ----------------------------
    scan_col = next((c for c in df.columns if 'scan' in c.lower() and 'in' in c.lower()), None)
    if not scan_col:
        st.error("❌ Could not find a 'Scan In Date' column.")
        st.stop()

    df[scan_col] = pd.to_datetime(df[scan_col], errors='coerce')
    df = df[~df[scan_col].isna()]

    scan_out_col = next((c for c in df.columns if 'scan' in c.lower() and 'out' in c.lower()), None)
    if scan_out_col:
        df[scan_out_col] = pd.to_datetime(df[scan_out_col], errors='coerce')

    # ----------------------------
    # Add SLA columns
    # ----------------------------
    new_cols = [
        'Stills Out of SLA', 'Day(s) out of SLA - STILLS',
        'Model Out of SLA', 'Day(s) out of SLA - MODEL',
        'Mannequin Out of SLA', 'Day(s) out of SLA - MANNEQUIN',
        'Notes', 'Days in Studio'
    ]
    for col in new_cols:
        df[col] = np.nan

    # Convert any date-like columns
    for c in df.columns:
        if "date" in c.lower():
            df[c] = pd.to_datetime(df[c], errors='coerce')

    # ----------------------------
    # SLA Logic
    # ----------------------------
    sla_mapping = {
        'Stills': ('Photo Still Date', 'Still Upload Date'),
        'Model': ('Photo Model Date', 'Model Upload Date'),
        'Mannequin': ('Photo Mannequin Date', 'Mannequin Upload Date')
    }

    for prefix, (photo_col, upload_col) in sla_mapping.items():
        if photo_col in df.columns:

            start = df[photo_col]
            end = df[upload_col] if upload_col in df.columns else pd.NaT

            effective_end = end.fillna(today)

            days_diff = [
                working_days_diff(
                    s.date() if pd.notna(s) else np.nan,
                    e.date() if pd.notna(e) else np.nan
                )
                for s, e in zip(start, effective_end)
            ]

            out_days = np.maximum(np.array(days_diff) - SLA_DAYS[prefix.upper()], 0)

            df[f"Day(s) out of SLA - {prefix.upper()}"] = out_days
            df[f"{prefix} Out of SLA"] = np.where(out_days > 0, "LATE", "")

    # ----------------------------
    # Awaiting model shot note
    # ----------------------------
    if "Photo Still Date" in df.columns and scan_out_col:
        mask = df[scan_out_col].isna() & df["Photo Still Date"].notna()
        diff = df.loc[mask, "Photo Still Date"].apply(lambda x: working_days_diff(x.date(), today))
        df.loc[mask & (diff > 2), "Notes"] = "Awaiting model shot"

    # ----------------------------
    # Days in studio
    # ----------------------------
    def compute_days_in_studio(row):
        scan_in = row.get(scan_col)
        scan_out = row.get(scan_out_col) if scan_out_col else None

        shot_cols = [
            'Photo Still Date', 'Photo Model Date', 'Photo Mannequin Date',
            'Still Upload Date', 'Model Upload Date', 'Mannequin Upload Date'
        ]
        all_blank = all(pd.isna(row.get(c)) for c in shot_cols)

        if pd.notna(scan_in) and pd.notna(scan_out) and all_blank:
            return "SCANNED OUT AND NEVER SHOT"
        elif pd.notna(scan_out):
            return "SCANNED OUT"
        elif pd.notna(scan_in):
            return working_days_diff(scan_in.date(), today)

        return np.nan

    df['Days in Studio'] = df.apply(compute_days_in_studio, axis=1)

    # ----------------------------
    # SLA summary column
    # ----------------------------
    sla_cols = ['Stills Out of SLA', 'Model Out of SLA', 'Mannequin Out of SLA']
    df["SLA status"] = df[sla_cols].apply(lambda r: "LATE" if "LATE" in r.values else "", axis=1)

    # ----------------------------
    # Preview + Download
    # ----------------------------
    st.subheader("Processed Data")
    st.dataframe(df)

    output = BytesIO()
    df.to_excel(output, index=False, engine="openpyxl")

    st.download_button(
        "📥 Download Processed Excel",
        output.getvalue(),
        file_name=f"check_retouch_processed_{today}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
