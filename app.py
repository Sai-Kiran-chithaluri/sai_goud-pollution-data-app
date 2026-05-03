import streamlit as st
import pandas as pd
import numpy as np
import io
from openpyxl import load_workbook
from openpyxl.styles import Alignment

st.set_page_config(page_title="Sai Goud Report Processor", layout="wide")

def find_header_row(file_bytes):
    df_preview = pd.read_excel(file_bytes, header=None, nrows=20)
    for i, row in df_preview.iterrows():
        row_values = [str(val).lower() for val in row.values]
        if any("sl no." in val or "time" in val for val in row_values):
            return i
    return 0

def process_pollution_report(file_bytes):
    # 1. Detect Header
    file_bytes.seek(0)
    header_idx = find_header_row(file_bytes)
    
    # 2. Load the actual workbook to preserve STYLES (Bold, Center, Colors)
    file_bytes.seek(0)
    wb = load_workbook(file_bytes)
    ws = wb.active
    
    # 3. Read data for processing logic
    file_bytes.seek(0)
    df_raw = pd.read_excel(file_bytes, header=header_idx)
    df_raw.columns = [str(c).strip() for c in df_raw.columns]
    
    u_cols = [c for c in df_raw.columns if c.endswith('_U')]
    
    for u_col in u_cols:
        col_idx_in_excel = df_raw.columns.get_loc(u_col) + 1 # openpyxl is 1-indexed
        base_name = u_col[:-2]
        l_col = base_name + '_L'
        
        # Merge logic
        if l_col in df_raw.columns:
            merged = df_raw[u_col].fillna(df_raw[l_col])
        else:
            merged = df_raw[u_col]
        
        # Strict Gap Filling (Preserve 0s)
        numeric_series = pd.to_numeric(merged, errors='coerce')
        is_shutdown = merged.astype(str).str.contains("Site Shutdown", case=False, na=False)
        is_missing = numeric_series.isna() & ~is_shutdown
        
        if is_missing.any():
            global_mean = numeric_series.dropna().mean() or 20.0
            for idx in df_raw[is_missing].index:
                window = numeric_series.iloc[max(0, idx-50):idx].dropna()
                if window.empty:
                    window = numeric_series.iloc[idx+1:idx+51].dropna()
                
                val = np.random.choice(window.values) if not window.empty else global_mean
                merged.iloc[idx] = np.round(val * np.random.uniform(0.98, 1.02), 2)
        
        # 4. WRITE DATA INTO THE TEMPLATE (Preserves Formatting)
        # Data starts at header_idx + 2
        for row_idx, value in enumerate(merged):
            cell = ws.cell(row=header_idx + row_idx + 2, column=col_idx_in_excel)
            cell.value = value
            # Ensure new data stays centered like the original
            cell.alignment = Alignment(horizontal='center')

    output = io.BytesIO()
    wb.save(output)
    return output.getvalue()

# --- STREAMLIT UI (With your custom wording) ---
st.title("Sai Goud Report Processor")
st.info("Welcome to my page! Upload your pollution report in Excel format")

file = st.sidebar.file_uploader("Upload Pollution Excel", type="xlsx")

if file:
    if st.button("press for Formatting"):
        with st.spinner("Processing report..."):
            # This now returns the full formatted Excel bytes
            processed_file_bytes = process_pollution_report(file)
            
            st.success("Processing complete! Thank you for Using Sai Goud's Website, Please Download the Processed Report Below and Come Back for More!")
            
            st.download_button(
                label="Download Report",
                data=processed_file_bytes,
                file_name="Universal_Processed_Report.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
