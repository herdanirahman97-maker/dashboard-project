import streamlit as st
import pandas as pd

st.set_page_config(page_title="Technician Dashboard", layout="wide")
st.title("🛠️ Technician Report Dashboard")

@st.cache_data(ttl=10)
def load_data():
    file_path = "Technician Report - Updated (1).xlsx"
    df = pd.read_excel(file_path, sheet_name="Service report", header=2)
    df = df.dropna(subset=['SN'])
    df['TANGGAL_PARSED'] = pd.to_datetime(df['TANGGAL'], errors='coerce')
    df = df.sort_values(by='TANGGAL_PARSED', ascending=False)
    return df

df = load_data()

st.subheader("🔍 Find Equipment Status")

# Create THREE columns so all search boxes sit neatly side-by-side
col1, col2, col3 = st.columns(3)

with col1:
    # 1. Hospital Search Box (NAMA CUSTOMER)
    rs_list = ["All"] + df['NAMA CUSTOMER'].dropna().astype(str).unique().tolist()
    search_rs = st.selectbox("1. Filter by Hospital (RS):", options=rs_list, index=0)

# Shrink the data based on Hospital selection
if search_rs != "All":
    df_filtered_rs = df[df['NAMA CUSTOMER'].astype(str) == search_rs]
else:
    df_filtered_rs = df

with col2:
    # 2. Modality Search Box (NAMA ALAT) 
    modality_list = ["All"] + df_filtered_rs['NAMA ALAT'].dropna().astype(str).unique().tolist()
    search_modality = st.selectbox("2. Filter by Modality:", options=modality_list, index=0)

# Shrink the data again based on Modality selection
if search_modality != "All":
    df_filtered_modality = df_filtered_rs[df_filtered_rs['NAMA ALAT'].astype(str) == search_modality]
else:
    df_filtered_modality = df_filtered_rs

with col3:
    # 3. SN Search Box 
    sn_list = [""] + df_filtered_modality['SN'].dropna().astype(str).unique().tolist()
    search_sn = st.selectbox("3. Select Serial Number (SN):", options=sn_list, index=0)

# Display Results
if search_sn:
    # If they picked an SN, show the detailed report
    sn_data = df[df['SN'].astype(str) == search_sn]
    if not sn_data.empty:
        latest = sn_data.iloc[0]
        st.success(f"Latest Update Found for SN: **{search_sn}**")
        
        info_col1, info_col2, info_col3 = st.columns(3)
        with info_col1:
            st.metric("Status", str(latest.get('STATUS', 'N/A')))
            st.write(f"**Technician:** {latest.get('PELAKSANA', 'N/A')}")
        with info_col2:
            st.write(f"**Date:** {latest.get('TANGGAL', 'N/A')}")
            st.write(f"**Tool:** {latest.get('NAMA ALAT', 'N/A')}")
        with info_col3:
            st.write(f"**Customer:** {latest.get('NAMA CUSTOMER', 'N/A')}")
            st.write(f"**Job Type:** {latest.get('JENIS PEKERJAAN', 'N/A')}")

        st.info(f"**Job Description:**\n\n{latest.get('URAIAN PEKERJAAN', 'No description.')}")
        
        # Show older history if it exists
        if len(sn_data) > 1:
            with st.expander(f"View History ({len(sn_data) - 1} older updates)"):
                st.dataframe(sn_data.drop(columns=['TANGGAL_PARSED']).iloc[1:])

elif search_rs != "All" or search_modality != "All":
    # If they picked a Hospital or Modality but haven't picked an SN yet, show a summary table!
    st.info("Showing machines based on your filters (Select an SN above to see full job details)")
    
    # Create a clean summary table of just the latest status
    summary_table = df_filtered_modality[['SN', 'NAMA CUSTOMER', 'NAMA ALAT', 'STATUS', 'TANGGAL']].drop_duplicates(subset=['SN'])
    st.dataframe(summary_table, use_container_width=True)
