import streamlit as st
import pandas as pd
import numpy as np
import io

# --- CORE LOGIC FUNCTION ---
def process_data(df_raw):
    # Clean headers
    df_raw.columns = [str(c).strip().replace('\n', ' ') for c in df_raw.columns]
    
    # Smart Merge (_L and _U)
    l_cols = [c for c in df_raw.columns if c.endswith('_L')]
    merged_data = df_raw.iloc[:, :2].copy()
    meta_cols = merged_data.columns.tolist()
    sn_col = meta_cols[0]

    for l_col in l_cols:
        base_name = l_col[:-2]
        u_col = base_name + '_U'
        if u_col in df_raw.columns:
            merged_data[base_name] = df_raw[l_col].combine_first(df_raw[u_col])
        else:
            merged_data[base_name] = df_raw[l_col]

    # Gap Filling
    merged_data[sn_col] = pd.to_numeric(merged_data[sn_col], errors='coerce')
    params = [c for c in merged_data.columns if c not in meta_cols]

    for col in params:
        # Fill specific gap 634-769 based on 500-633
        gap_indices = merged_data[(merged_data[sn_col] >= 634) & (merged_data[sn_col] <= 769)].index
        ref_v = pd.to_numeric(merged_data[(merged_data[sn_col] >= 500) & (merged_data[sn_col] < 634)][col], errors='coerce').dropna().values
        
        if len(ref_v) > 0:
            samples = np.random.choice(ref_v, size=len(gap_indices))
            merged_data.loc[gap_indices, col] = np.round(samples * np.random.uniform(0.98, 1.02, size=len(gap_indices)), 2)

        # Fill remaining stray NaNs
        for idx in merged_data[merged_data[col].isnull()].index:
            sn_val = merged_data.loc[idx, sn_col]
            window = merged_data[(merged_data[sn_col] < sn_val) & (merged_data[sn_col] >= sn_val - 50)][col].dropna().values
            if len(window) > 0:
                merged_data.loc[idx, col] = np.round(np.random.choice(window) * np.random.uniform(0.98, 1.02), 2)
    
    return merged_data

# --- STREAMLIT UI ---
st.set_page_config(page_title="Pollution Report Cleaner", layout="centered")
st.title("📊 Pollution Report Merger & Filler")
st.write("Upload your raw Excel file to merge L/U columns and fill missing data.")

uploaded_file = st.file_uploader("Choose an Excel file", type="xlsx")

if uploaded_file:
    # Preview the data
    df = pd.read_excel(uploaded_file, header=1)
    st.success("File uploaded successfully!")
    st.subheader("Raw Data Preview (First 5 rows)")
    st.write(df.head())

    if st.button("✨ Process and Fill Data"):
        with st.spinner('Merging and simulating data...'):
            final_df = process_data(df)
            
            # Convert dataframe to Excel in memory
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                final_df.to_excel(writer, index=False)
            
            st.balloons()
            st.subheader("✅ Processing Complete!")
            
            # Download button
            st.download_button(
                label="📥 Download Cleaned Excel File",
                data=output.getvalue(),
                file_name="Cleaned_Pollution_Report.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )