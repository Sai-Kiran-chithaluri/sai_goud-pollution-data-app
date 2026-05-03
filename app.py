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
    file_bytes.seek(0)
    # 1. Load the workbook
    wb = load_workbook(file_bytes)
    ws = wb.active
    
    # 2. Find the header row accurately
    header_row_num = 1
    for row in ws.iter_rows(min_row=1, max_row=20):
        for cell in row:
            if cell.value and "sl no." in str(cell.value).lower():
                header_row_num = cell.row
                break
        else: continue
        break

    # 3. Read data for processing logic
    file_bytes.seek(0)
    df_raw = pd.read_excel(file_bytes, header=header_row_num - 1)
    df_raw.columns = [str(c).strip() for c in df_raw.columns]
    
    u_cols = [c for c in df_raw.columns if c.endswith('_U')]
    
    for u_col in u_cols:
        # Find which column index this is in Excel (1-based)
        col_idx_in_excel = df_raw.columns.get_loc(u_col) + 1
        
        base_name = u_col[:-2]
        l_col = base_name + '_L'
        
        # Merge U and L
        if l_col in df_raw.columns:
            merged = df_raw[u_col].fillna(df_raw[l_col])
        else:
            merged = df_raw[u_col]
        
        # Process Gaps
        numeric_series = pd.to_numeric(merged, errors='coerce')
        is_missing = numeric_series.isna()
        
        if is_missing.any():
            global_mean = numeric_series.dropna().mean() or 20.0
            
            # 4. WRITE TO EXCEL
            for idx, value in enumerate(merged):
                # Calculate absolute excel row: 
                # Header is 'header_row_num', first data is 'header_row_num + 1'
                excel_row = header_row_num + idx + 1
                
                target_cell = ws.cell(row=excel_row, column=col_idx_in_excel)
                
                # If it was NA or empty, fill it!
                if target_cell.value is None or str(target_cell.value).strip().lower() in ['nan', 'na', '']:
                    # Use the filled value logic
                    val = numeric_series.iloc[idx]
                    if pd.isna(val):
                        # Simple fill for testing - you can use your random logic here
                        val = round(global_mean * np.random.uniform(0.98, 1.02), 2)
                    
                    target_cell.value = val
                
                # Keep it centered
                target_cell.alignment = Alignment(horizontal='center')

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
