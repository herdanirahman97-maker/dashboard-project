import streamlit as st
import pandas as pd
from datetime import datetime

st.set_page_config(
    page_title="Mandaya Royal Hospital - Radiology Quality System",
    page_icon="🏥",
    layout="wide"
)

# Custom Styling
st.markdown("""
    <style>
    .main-header {
        font-size: 24px;
        font-weight: bold;
        color: #002B5B;
    }
    .sub-header {
        font-size: 14px;
        color: #64748B;
    }
    </style>
""", unsafe_allow_html=True)

# Sidebar Navigation
st.sidebar.title("🏥 Mandaya Radiology")
st.sidebar.markdown("---")
menu = st.sidebar.radio(
    "Pilih Menu Utama:",
    [
        "📊 Dashboard",
        "📝 Input Harian QC",
        "📅 Live Preview & Download All",
        "📈 PMI (Pemantapan Mutu Internal)",
        "🛡️ PME (Pemantapan Mutu Eksternal)",
        "⚙️ Sistem & Revisi"
    ]
)

# Header Title
st.markdown('<p class="main-header">MANDAYA ROYAL HOSPITAL — RADIOLOGY QUALITY SYSTEM</p>', unsafe_allow_html=True)
st.markdown('<p class="sub-header">Departemen Radiologi | Enterprise Dashboard & Quality Assurance 2026</p>', unsafe_allow_html=True)
st.markdown("---")

# 1. DASHBOARD
if menu == "📊 Dashboard":
    st.subheader("📊 Ringkasan Mutu Departemen Radiologi")
    col1, col2, col3, col4 = st.columns(4)
    with col1:
        st.metric("Total Modality Aktif", "12 Unit", "100% Normal")
    with col2:
        st.metric("Compliance PMI", "98.5%", "+1.2%")
    with col3:
        st.metric("Jadwal PME & Kalibrasi", "3 Menunggu", "Valid")
    with col4:
        st.metric("Status Akreditasi", "Paripurna", "SNARS / JCI")
    
    st.info("💡 Gunakan menu navigasi di sebelah kiri untuk mengakses input harian, preview langsung, serta portal mandiri PMI dan PME.")

# 2. INPUT HARIAN QC
elif menu == "📝 Input Harian QC":
    st.subheader("📝 Form Input Harian Quality Control & Suhu Ruangan")
    with st.form("input_harian_form"):
        tanggal = st.date_input("Tanggal Pencatatan", datetime.today())
        ruangan = st.selectbox("Pilih Ruangan / Modality", [
            "MRI 3T (IQON / Philips)",
            "CT Scan 256 Slices",
            "X-Ray Digital Diagnost C50",
            "Fluoroscopy CombiDiagnost R90",
            "Mammography FDR",
            "USG Philips Epic Elite",
            "Ruang Baca & Konsultasi"
        ])
        
        col_a, col_b = st.columns(2)
        with col_a:
            suhu = st.number_input("Suhu Ruangan (°C) [Standar: 18-23°C]", value=20.5)
            kelembapan = st.number_input("Kelembapan (%) [Standar: 40-60%]", value=50.0)
        with col_b:
            status_qc = st.selectbox("Hasil QC Harian Alat", ["Normal / Layak", "Perlu Perhatian", "Maintenance / Rusak"])
            petugas = st.text_input("Nama Petugas Radiografer / Fisikawan")
            
        catatan = st.text_area("Catatan Tambahan / Tindak Lanjut")
        submitted = st.form_submit_button("💾 Simpan Data Harian")
        if submitted:
            st.success(f"Data untuk {ruangan} tanggal {tanggal} berhasil disimpan ke database sistem!")

# 3. LIVE PREVIEW & DOWNLOAD ALL
elif menu == "📅 Live Preview & Download All":
    st.subheader("📅 Live Preview & Download Center")
    tab_p1, tab_p2 = st.tabs(["🌡️ Rekap Suhu & Kelembapan", "🔬 Rekap Daily QC Modality"])
    
    with tab_p1:
        st.markdown("### Pengaturan Preview & Unduh Suhu Ruangan")
        pilih_ruangan = st.selectbox("Filter Ruangan:", ["🌐 Pilih Semua Ruangan (Download All)", "MRI 3T", "CT Scan", "X-Ray C50", "Fluoroscopy R90"])
        
        df_suhu = pd.DataFrame({
            "Tanggal": ["01/10/2026", "02/10/2026", "03/10/2026", "04/10/2026"],
            "Ruangan": ["MRI 3T", "MRI 3T", "CT Scan", "CT Scan"],
            "Suhu (°C)": [20.1, 20.4, 19.8, 20.2],
            "Kelembapan (%)": [52, 50, 48, 51],
            "Status": ["OK", "OK", "OK", "OK"]
        })
        st.dataframe(df_suhu, use_container_width=True)
        
        if st.button("🖨️ Print / Save as PDF (Suhu Ruangan)"):
            st.success("File PDF Rekap Suhu berhasil disiapkan untuk diunduh.")
            
    with tab_p2:
        st.markdown("### Pengaturan Preview & Unduh Daily QC")
        pilih_mod = st.selectbox("Filter Modality:", ["🌐 Pilih Semua Modality (Download All)", "Philips DigitalDiagnost", "IQON Spectral CT", "CombiDiagnost R90"])
        
        df_qc = pd.DataFrame({
            "No": [1, 2, 3],
            "Modality": ["DigitalDiagnost C50", "IQON Spectral CT", "CombiDiagnost R90"],
            "Parameter Uji": ["Suhu & Kelayakan Kaset", "CT Number & Uniformity", "Akurasi Tegangan kVp"],
            "Output Target": ["Checklist OK", "Parameter Phantom OK", "Sesuai Standar"],
            "Status": ["Selesai", "Selesai", "Selesai"]
        })
        st.dataframe(df_qc, use_container_width=True)
        
        if st.button("🖨️ Print / Save as PDF (Daily QC All)"):
            st.success("File PDF Rekap Daily QC All berhasil disiapkan.")

# 4. PMI (PEMANTAPAN MUTU INTERNAL)
elif menu == "📈 PMI (Pemantapan Mutu Internal)":
    st.subheader("📈 Pemantapan Mutu Internal (PMI) Departemen Radiologi")
    pilih_kategori_pmi = st.selectbox("Pilih Kategori Periode PMI:", ["Harian", "Mingguan", "Bulanan", "Tahunan"])
    
    if pilih_kategori_pmi == "Harian":
        st.markdown("#### Parameter PMI Harian")
        st.table(pd.DataFrame({
            "No": [1, 2, 3],
            "Modality / Ruangan": ["Semua Ruangan Pemeriksaan", "Philips DigitalDiagnost C50", "Kamar Gelap / CR Reader"],
            "Item Uji & Parameter": ["Pengecekan Suhu (18-23C) & Kelembapan (40-60%)", "Checklist Kesiapan Alat & Ekposur", "Kualitas Kaset & Penghapusan Sinar"],
            "Target Mutu": ["Sesuai Standar Ruang", "Normal Tanpa Error", "Bersih dari Artifact"]
        }))
    elif pilih_kategori_pmi == "Mingguan":
        st.markdown("#### Parameter PMI Mingguan")
        st.table(pd.DataFrame({
            "No": [1, 2],
            "Modality": ["IQON Spectral CT", "MRI 3T Philips"],
            "Item Uji": ["CT Number, Uniformity, & Noise", "Signal to Noise Ratio (SNR) & Image Artifact"],
            "Target": ["Toleransi ± 4 HU", "Sesuai Baseline Pabrikan"]
        }))
    elif pilih_kategori_pmi == "Bulanan":
        st.markdown("#### Parameter PMI Bulanan")
        st.table(pd.DataFrame({
            "No": [1, 2],
            "Modality": ["Fluoroscopy CombiDiagnost R90", "Mammography FDR"],
            "Item Uji": ["Akurasi Pergerakan Meja & Kolimasi", "Compression Force & Phantom Imaging"],
            "Target": ["Berjalan Lancar Tanpa Hambatan", "Sesuai Standar Akreditasi"]
        }))
    else:
        st.markdown("#### Parameter PMI Tahunan")
        st.table(pd.DataFrame({
            "No": [1, 2],
            "Modality": ["Seluruh Ruangan Radiologi & OT", "Apron & Alat Pelindung Diri (APD)"],
            "Item Uji": ["Uji Kebocoran Tabung X-Ray & Ruangan", "Fluoroskopi Uji Integritas Apron Timbal"],
            "Target": ["Batas Aman Paparan Radiasi", "Tidak Ada Retak/Kebocoran Timbal"]
        }))

# 5. PME (PEMANTAPAN MUTU EKSTERNAL)
elif menu == "🛡️ PME (Pemantapan Mutu Eksternal)":
    st.subheader("🛡️ Pemantapan Mutu Eksternal (PME) & Kalibrasi")
    pilih_kategori_pme = st.selectbox("Pilih Kategori PME / Layanan Eksternal:", ["Preventive Maintenance (PM Philips)", "Uji Kesesuaian (UKES) BAPETEN", "Kalibrasi Alat Ukur Radiasi"])
    
    if pilih_kategori_pme == "Preventive Maintenance (PM Philips)":
        st.markdown("#### Jadwal Pemeliharaan Berkala oleh Vendor Resmi Philips")
        st.table(pd.DataFrame({
            "No": [1, 2, 3],
            "Modality": ["IQON Spectral CT", "Philips DigitalDiagnost C50", "CombiDiagnost R90"],
            "Jenis Layanan": ["PM Berkala 4 Bulan Sekali", "PM Berkala 3 Bulan Sekali", "PM Berkala 6 Bulan Sekali"],
            "Jadwal Berikutnya": ["Mei 2026", "Juni 2026", "Juli 2026"],
            "Status": ["Terjadwal", "Terjadwal", "Terjadwal"]
        }))
    elif pilih_kategori_pme == "Uji Kesesuaian (UKES) BAPETEN":
        st.markdown("#### Sertifikasi Uji Kesesuaian Instalasi Radiodiagnostik (4 Tahun Sekali)")
        st.table(pd.DataFrame({
            "No": [1, 2],
            "Alat / Pesawat": ["X-Ray General R/F", "CT Scan 256 Slices"],
            "Lembaga Penguji": ["LKT Berizin BAPETEN", "LKT Berizin BAPETEN"],
            "Masa Berlaku SIO": ["2024 s.d. 2028", "2025 s.d. 2029"],
            "Status": ["Aktif", "Aktif"]
        }))
    else:
        st.markdown("#### Kalibrasi Alat Ukur Radiasi (1 Tahun Sekali)")
        st.table(pd.DataFrame({
            "No": [1, 2],
            "Instrumen / Alat Ukur": ["Survey Meter (Ratemeter)", "Dosimeter Perorangan (TLD / OSL)"],
            "Laboratorium Kalibrasi": ["Standard ASTM / BPFK", "Pusat Kalibrasi Terakreditasi KAN"],
            "Periode Kalibrasi": ["Januari 2026", "Desember 2025"],
            "Status Selesai": ["Sertifikat Valid", "Sertifikat Valid"]
        }))

# 6. SISTEM & REVISI
else:
    st.subheader("⚙️ Sistem, Log, & Catatan Revisi")
    st.markdown("### Histori Pembaruan Sistem (Changelog)")
    st.markdown("""
    * **Versi 3.5 (April 2026):** Pemisahan total menu independen antara PMI (Internal) dan PME (Eksternal) sesuai permintaan user.
    * **Versi 3.0 (Maret 2026):** Penambahan fitur *Download All* untuk rekapitulasi suhu ruangan dan daily QC modality.
    * **Versi 2.0 (Februari 2026):** Integrasi standar parameter rumah sakit korporat dan layout enterprise.
    """)
    st.success("Sistem berjalan dengan stabil dan siap diimplementasikan.")
