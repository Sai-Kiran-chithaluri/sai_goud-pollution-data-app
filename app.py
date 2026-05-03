import streamlit as st
import pandas as pd
import numpy as np
import io

st.set_page_config(page_title="Pollution Data Automator", layout="wide")

def find_header_row(file_bytes):
    df_preview = pd.read_excel(file_bytes, header=None, nrows=20)
    for i, row in df_preview.iterrows():
        row_values = [str(val).lower() for val in row.values]
        if any("sl no." in val or "time" in val for val in row_values):
            return i
    return 0

def process_pollution_report(file_bytes):
    file_bytes.seek(0)
    header_idx = find_header_row(file_bytes)
    
    # 1. CAPTURE TOP METADATA (Headers like 'Online Pollution Monitoring Portal')
    file_bytes.seek(0)
    meta_rows = pd.read_excel(file_bytes, header=None, nrows=header_idx)

    # 2. READ DATA TABLE
    file_bytes.seek(0)
    df_raw = pd.read_excel(file_bytes, header=header_idx)
    df_raw.columns = [str(c).strip() for c in df_raw.columns]
    
    sn_col = df_raw.columns[0]
    time_col = df_raw.columns[1]
    result_df = df_raw[[sn_col, time_col]].copy()
    
    u_cols = [c for c in df_raw.columns if c.endswith('_U')]
    
    for u_col in u_cols:
        base_name = u_col[:-2]
        l_col = base_name + '_L'
        
        # Merge U and L
        if l_col in df_raw.columns:
            merged = df_raw[u_col].fillna(df_raw[l_col])
        else:
            merged = df_raw[u_col]
        
        # STRICT GAP FILLING
        # numeric_series helps identify what is a number vs NaN
        numeric_series = pd.to_numeric(merged, errors='coerce')
        is_shutdown = merged.astype(str).str.contains("Site Shutdown", case=False, na=False)
        
        # TARGET ONLY NaNs: If it's 0, it is NOT missing.
        is_missing = numeric_series.isna() & ~is_shutdown
        
        if is_missing.any():
            # Mean for fallback (includes zeros in calculation)
            global_mean = numeric_series.dropna().mean()
            if np.isnan(global_mean): global_mean = 20.0
            
            for idx in result_df[is_missing].index:
                # Look Back (Previous 50 valid values)
                window = numeric_series.iloc[max(0, idx-50):idx].dropna()
                # Look Forward (Next 50 valid values) if Look Back is empty
                if window.empty:
                    window = numeric_series.iloc[idx+1:idx+51].dropna()
                
                if not window.empty:
                    val = np.random.choice(window.values)
                    merged.iloc[idx] = np.round(val * np.random.uniform(0.98, 1.02), 2)
                else:
                    # Cold Start (Start of file gaps)
                    merged.iloc[idx] = np.round(global_mean * np.random.uniform(0.95, 1.05), 2)

        result_df[base_name] = merged
        
    return meta_rows, result_df

# --- STREAMLIT UI ---
st.title("Sai Goud Report Processor")
st.info("Welcome to my page! Upload your pollution report in Excel format")

file = st.sidebar.file_uploader("Upload Pollution Excel", type="xlsx")

if file:
    if st.button("press for Formatting"):
        with st.spinner("Processing report..."):
            meta, processed_data = process_pollution_report(file)
            
            # RECONSTRUCT THE EXCEL
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                # Write original headers
                meta.to_excel(writer, index=False, header=False)
                # Write data table below headers
                processed_data.to_excel(writer, index=False, startrow=len(meta))
            
            st.success("Processing complete! Thank you for Using Sai Goud's Website, Please Download the Processed Report Below and Come Back for More!")
            st.dataframe(processed_data.head(10))
            
            st.download_button(
                label="Download Report",
                data=output.getvalue(),
                file_name="Universal_Processed_Report.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
