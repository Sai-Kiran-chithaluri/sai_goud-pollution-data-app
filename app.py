import streamlit as st
import pandas as pd
import numpy as np
import io

st.set_page_config(page_title="Universal Data Merger", layout="wide")

def find_header_row(file_bytes):
    """Automatically finds which row contains 'Sl No.' or 'Time'"""
    # Read just the first 20 rows to locate the header
    df_preview = pd.read_excel(file_bytes, header=None, nrows=20)
    for i, row in df_preview.iterrows():
        row_values = [str(val).lower() for val in row.values]
        if any("sl no." in val or "time" in val for val in row_values):
            return i
    return 0  # Default to row 0 if keywords aren't found

def process_pollution_report(df_raw):
    # 1. Clean column names
    df_raw.columns = [str(c).strip() for c in df_raw.columns]
    
    # 2. Identify Metadata
    sn_col = df_raw.columns[0]
    time_col = df_raw.columns[1]
    
    result_df = df_raw[[sn_col, time_col]].copy()
    
    # 3. Find parameter bases (columns ending in _U)
    u_cols = [c for c in df_raw.columns if c.endswith('_U')]
    
    for u_col in u_cols:
        base_name = u_col[:-2]
        l_col = base_name + '_L'
        
        # Merge Logic
        if l_col in df_raw.columns:
            merged = df_raw[u_col].fillna(df_raw[l_col])
        else:
            merged = df_raw[u_col]
        
        # 4. Universal Gap Filling (Ignoring 'Site Shutdown')
        numeric_series = pd.to_numeric(merged, errors='coerce')
        is_shutdown = merged.astype(str).str.contains("Site Shutdown", case=False, na=False)
        is_missing = numeric_series.isna() & ~is_shutdown
        
        if is_missing.any():
            for idx in result_df[is_missing].index:
                # Look at previous 50 rows for valid numbers to sample from
                window = numeric_series.iloc[max(0, idx-50):idx].dropna()
                
                if not window.empty:
                    val = np.random.choice(window.values)
                    merged.iloc[idx] = np.round(val * np.random.uniform(0.98, 1.02), 2)
                else:
                    # Look forward if backward window is empty
                    window_future = numeric_series.iloc[idx+1:idx+51].dropna()
                    if not window_future.empty:
                        val = np.random.choice(window_future.values)
                        merged.iloc[idx] = np.round(val * np.random.uniform(0.98, 1.02), 2)

        result_df[base_name] = merged
        
    return result_df

# --- STREAMLIT UI ---
st.title("📊 Smart Pollution Data Merger")
st.markdown("This tool automatically detects headers, merges `_U/_L` columns, and fills gaps while preserving **'Site Shutdown'** labels.")

file = st.sidebar.file_uploader("Upload Excel File", type="xlsx")

if file:
    # --- AUTOMATIC HEADER DETECTION ---
    header_index = find_header_row(file)
    # Reset file pointer to read again from the start
    file.seek(0) 
    
    # Read the data using the detected header index
    df = pd.read_excel(file, header=header_index)
    
    st.write(f"✅ Auto-detected headers at row {header_index + 1}")
    
    if st.button("🚀 Process Entire Sheet"):
        with st.spinner("Analyzing structure and filling gaps..."):
            final_df = process_pollution_report(df)
            
            st.success("Processing complete!")
            st.dataframe(final_df)
            
            # Excel Download
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                final_df.to_excel(writer, index=False)
            
            st.download_button(
                label="📥 Download Processed Excel",
                data=output.getvalue(),
                file_name="Universal_Processed_Report.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
