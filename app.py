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

# Create two columns so the search boxes sit side-by-side
col1, col2 = st.columns(2)

with col1:
    # 1. Modality Search Box (NAMA ALAT)
    modality_list = ["All"] + df['NAMA ALAT'].dropna().astype(str).unique().tolist()
    search_modality = st.selectbox("1. Filter by Modality Name (Optional):", options=modality_list, index=0)

# Filter the data so the SN list only shows SNs for the chosen Modality
if search_modality != "All":
    filtered_df = df[df['NAMA ALAT'].astype(str) == search_modality]
else:
    filtered_df = df

with col2:
    # 2. SN Search Box
    sn_list = [""] + filtered_df['SN'].dropna().astype(str).unique().tolist()
    search_sn = st.selectbox("2. Search or select a Serial Number (SN):", options=sn_list, index=0)

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

elif search_modality != "All":
    # If they picked a Modality but haven't picked an SN yet, show a summary table!
    st.info(f"Showing all machines for Modality: **{search_modality}** (Please select an SN above for full details)")
    
    # Create a clean summary table of just the latest status for all machines in that modality
    summary_table = filtered_df[['SN', 'NAMA CUSTOMER', 'STATUS', 'TANGGAL']].drop_duplicates(subset=['SN'])
    st.dataframe(summary_table, use_container_width=True)
