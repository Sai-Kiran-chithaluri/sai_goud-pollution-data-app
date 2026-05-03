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
    wb = load_workbook(file_bytes)
    ws = wb.active
    
    # 1. Find Header Row
    header_row_num = 7 
    for row in ws.iter_rows(min_row=1, max_row=20):
        for cell in row:
            if cell.value and "sl no." in str(cell.value).lower():
                header_row_num = cell.row
                break

    # 2. Map Columns (U and L pairs)
    headers = [str(cell.value).strip() if cell.value else "" for cell in ws[header_row_num]]
    u_cols = {i+1: h[:-2] for i, h in enumerate(headers) if h.endswith('_U')}
    
    # 3. Process every row
    for row_idx in range(header_row_num + 1, ws.max_row + 1):
        for col_idx, base_name in u_cols.items():
            u_cell = ws.cell(row=row_idx, column=col_idx)
            
            # Find the matching _L column for this base name
            l_col_name = base_name + "_L"
            l_col_idx = None
            if l_col_name in headers:
                l_col_idx = headers.index(l_col_name) + 1
            
            # Logic: Use U, if U is empty use L, if both empty use random
            val_u = u_cell.value
            val_l = ws.cell(row=row_idx, column=l_col_idx).value if l_col_idx else None
            
            def is_empty(v):
                s = str(v).strip().lower()
                return v is None or s in ['nan', 'na', '', 'n/a']

            if is_empty(val_u):
                if not is_empty(val_l):
                    # Fill U with L's value
                    u_cell.value = val_l
                else:
                    # Both empty - Fill with random pollution data
                    u_cell.value = round(np.random.uniform(18.5, 24.5), 2)
                
                # Style the new cell
                u_cell.alignment = Alignment(horizontal='center')
                
            # Optional: Rename the header from 'XYZ_U' to 'XYZ'
            ws.cell(row=header_row_num, column=col_idx).value = base_name

    # 4. Save
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
