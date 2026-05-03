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
    
    # 1. Find the exact row where data starts
    header_row_num = 7 
    for row in ws.iter_rows(min_row=1, max_row=20):
        for cell in row:
            if cell.value and "sl no." in str(cell.value).lower():
                header_row_num = cell.row
                break

    # 2. Identify all columns ending in _U
    u_col_indices = []
    for cell in ws[header_row_num]:
        val = str(cell.value).strip() if cell.value else ""
        if val.endswith('_U'):
            u_col_indices.append(cell.column)

    # 3. FORCE FILL every single non-numeric or empty cell
    for row_idx in range(header_row_num + 1, ws.max_row + 1):
        for col_idx in u_col_indices:
            cell = ws.cell(row=row_idx, column=col_idx)
            
            # Get the value and clean it
            raw_val = cell.value
            clean_val = str(raw_val).strip().lower() if raw_val is not None else ""
            
            # If it's NOT a number, we replace it
            # This covers None, 'nan', 'na', ' ', and 'site shutdown'
            try:
                # Try to see if it's already a valid number
                float(clean_val)
                is_numeric = True
            except:
                is_numeric = False

            if not is_numeric or clean_val in ['nan', '']:
                # Generate the new data
                new_value = round(np.random.uniform(19.5, 23.5), 2)
                
                # Overwrite the cell
                cell.value = new_value
                
                # Force center alignment
                cell.alignment = Alignment(horizontal='center')

    # 4. Final Save
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
