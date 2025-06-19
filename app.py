import streamlit as st
import pandas as pd
import sys
import subprocess
import pkg_resources
from io import BytesIO

# ==============================================
# FUNGSI UTAMA DAN DEPENDENCIES
# ==============================================

def check_dependencies():
    """Memeriksa dan menginstall dependencies yang diperlukan"""
    required = {
        'google-api-python-client',
        'gspread', 
        'oauth2client',
        'pandas',
        'streamlit'
    }
    
    installed = {pkg.key for pkg in pkg_resources.working_set}
    missing = required - installed
    
    if missing:
        st.warning(f"Modul yang diperlukan belum terinstall: {missing}")
        if st.button("Install Modul yang Hilang"):
            for package in missing:
                subprocess.check_call([sys.executable, "-m", "pip", "install", package])
            st.success("Modul berhasil diinstall! Silakan refresh halaman.")
            st.stop()

# Panggil fungsi cek dependencies
check_dependencies()

# Import modul setelah dependencies terpenuhi
try:
    import gspread
    from oauth2client.service_account import ServiceAccountCredentials
    from googleapiclient.discovery import build
    from googleapiclient.errors import HttpError
except ImportError as e:
    st.error(f"Gagal mengimport modul: {str(e)}")
    st.stop()

# ==============================================
# KONFIGURASI AWAL
# ==============================================

st.set_page_config("Form Input BOQ", layout="centered")
st.title("📋 Form Input BOQ Otomatis")

# Inisialisasi session state
def reset_form():
    """Reset semua nilai form ke default"""
    st.session_state.update({
        'sumber': "ODC",
        'lop_name': "",
        'kabel_12': 0.0,
        'kabel_24': 0.0,
        'odp_8': 0,
        'odp_16': 0,
        'tiang_new': 0,
        'tiang_existing': 0,
        'tikungan': 0,
        'izin': "",
        'downloaded': False
    })

if 'sumber' not in st.session_state:
    reset_form()

# ==============================================
# FORM INPUT
# ==============================================

with st.form("boq_form"):
    st.subheader("🔹 Data Proyek")
    col1, col2 = st.columns(2)
    with col1:
        st.radio(
            "Sumber Data:", 
            ["ODC", "ODP"],
            index=0,
            key='sumber'
        )
    with col2:
        st.text_input(
            "Nama LOP (untuk nama file):",
            key='lop_name'
        )

    st.subheader("🔹 Input Kabel")
    col1, col2 = st.columns(2)
    with col1:
        st.number_input(
            "Panjang Kabel 12 Core (meter):",
            min_value=0.0,
            value=st.session_state.kabel_12,
            key='kabel_12'
        )
    with col2:
        st.number_input(
            "Panjang Kabel 24 Core (meter):",
            min_value=0.0,
            value=st.session_state.kabel_24,
            key='kabel_24'
        )

    st.subheader("🔹 Input ODP")
    col1, col2 = st.columns(2)
    with col1:
        st.number_input(
            "ODP 8 Port:",
            min_value=0,
            value=st.session_state.odp_8,
            key='odp_8'
        )
    with col2:
        st.number_input(
            "ODP 16 Port:",
            min_value=0,
            value=st.session_state.odp_16,
            key='odp_16'
        )

    st.subheader("🔹 Input Pendukung")
    st.number_input(
        "Tiang Baru:",
        min_value=0,
        value=st.session_state.tiang_new,
        key='tiang_new'
    )
    st.number_input(
        "Tiang Existing:",
        min_value=0,
        value=st.session_state.tiang_existing,
        key='tiang_existing'
    )
    st.number_input(
        "Jumlah Tikungan:",
        min_value=0,
        value=st.session_state.tikungan,
        key='tikungan'
    )
    st.text_input(
        "Nilai Izin (jika ada):",
        value=st.session_state.izin,
        key='izin'
    )

    col1, col2 = st.columns(2)
    with col1:
        submitted = st.form_submit_button("🚀 Proses BOQ")
    with col2:
        reset_clicked = st.form_submit_button("🔄 Reset Form")

# ==============================================
# PENANGANAN FORM SUBMIT
# ==============================================

if reset_clicked:
    reset_form()
    st.rerun()

if submitted:
    # Validasi input
    if not st.session_state.lop_name:
        st.warning("Harap masukkan Nama LOP terlebih dahulu!")
        st.stop()
    
    try:
        # Hitung total ODP
        total_odp = st.session_state.odp_8 + st.session_state.odp_16
        
        # 1. Perhitungan Volume Kabel
        if st.session_state.sumber == "ODC":
            vol_kabel_12 = round((st.session_state.kabel_12 * 1.02) + total_odp) if st.session_state.kabel_12 > 0 else 0
            vol_kabel_24 = round((st.session_state.kabel_24 * 1.02) + total_odp) if st.session_state.kabel_24 > 0 else 0
        else:  # ODP
            vol_kabel_12 = round(st.session_state.kabel_12 * 1.02) if st.session_state.kabel_12 > 0 else 0
            vol_kabel_24 = round(st.session_state.kabel_24 * 1.02) if st.session_state.kabel_24 > 0 else 0
        
        # 2. Perhitungan PU-AS
        vol_puas = (total_odp * 2 - 1) if total_odp > 1 else (1 if total_odp == 1 else 0)
        vol_puas += st.session_state.tiang_new + st.session_state.tiang_existing + st.session_state.tikungan

        # 3. Perhitungan OS
        if st.session_state.sumber == "ODC":
            os_odc = (12 if st.session_state.kabel_12 > 0 else 24 if st.session_state.kabel_24 > 0 else 0) + total_odp
            os_odp = 0
        else:  # ODP
            os_odc = 0
            os_odp = total_odp * 2
        
        os_total = os_odc + os_odp

        # 4. Perhitungan PC
        pc_upc = (total_odp - 1) // 4 + 1 if total_odp > 0 else 0
        pc_apc = 18 if pc_upc == 1 else (pc_upc * 2 if pc_upc > 1 else 0)

        # 5. Perhitungan Lainnya
        tc02 = 1 if st.session_state.sumber == "ODC" else 0
        dd40 = 6 if st.session_state.sumber == "ODC" else 0
        bc06 = 6 if st.session_state.sumber == "ODC" else 0
        ps_odc = (total_odp - 1) // 4 + 1 if st.session_state.sumber == "ODC" and total_odp > 0 else 0

        # Membuat DataFrame hasil
        designators = []
        volumes = []
        
        def add_item(designator, volume):
            if volume > 0 or (designator == "Preliminary Project HRB/Kawasan Khusus" and st.session_state.izin):
                designators.append(designator)
                volumes.append(volume)
        
        # Tambahkan item ke dataframe
        if st.session_state.kabel_12 > 0:
            add_item("AC-OF-SM-12-SC_O_STOCK", vol_kabel_12)
        if st.session_state.kabel_24 > 0:
            add_item("AC-OF-SM-24-SC_O_STOCK", vol_kabel_24)
        if st.session_state.odp_8 > 0:
            add_item("ODP Solid-PB-8 AS", st.session_state.odp_8)
        if st.session_state.odp_16 > 0:
            add_item("ODP Solid-PB-16 AS", st.session_state.odp_16)
        
        add_item("PU-S7.0-400NM", st.session_state.tiang_new)
        add_item("PU-AS", vol_puas)
        
        if st.session_state.sumber == "ODC":
            add_item("OS-SM-1-ODC", os_odc)
            add_item("TC-02-ODC", tc02)
            add_item("DD-HDPE-40-1", dd40)
            add_item("BC-TR-0.6", bc06)
            add_item("PS-1-4-ODC", ps_odc)
        else:
            add_item("OS-SM-1-ODP", os_odp)
        
        add_item("OS-SM-1", os_total)
        add_item("PC-UPC-652-2", pc_upc)
        add_item("PC-APC/UPC-652-A1", pc_apc)
        
        if st.session_state.izin:
            add_item("Preliminary Project HRB/Kawasan Khusus", 1)
        
        df = pd.DataFrame({"Designator": designators, "Volume": volumes})

        # Tampilkan hasil
        st.success("✅ Perhitungan BOQ Berhasil!")
        st.subheader("📊 Hasil Perhitungan BOQ")
        st.dataframe(df.style.highlight_max(axis=0), use_container_width=True)

        # Tampilkan Total
        st.subheader("📌 Ringkasan")
        col1, col2, col3 = st.columns(3)
        with col1:
            st.metric("Total Kabel (m)", f"{vol_kabel_12 + vol_kabel_24:,}m")
        with col2:
            st.metric("Total ODP", f"{total_odp:,} unit")
        with col3:
            st.metric("Total PU-AS", f"{vol_puas:,} unit")

        # Simpan hasil di session state
        st.session_state.df_result = df
        st.session_state.show_download = True

    except Exception as e:
        st.error(f"Terjadi kesalahan saat memproses: {str(e)}")

# ==============================================
# DOWNLOAD OPTIONS
# ==============================================

if st.session_state.get('show_download', False):
    st.subheader("💾 Download Options")
    
    tab1, tab2 = st.tabs(["Download Hasil BOQ", "Download Full RAB"])

    with tab1:
        # Download hasil BOQ sebagai Excel
        output_boq = BytesIO()
        with pd.ExcelWriter(output_boq, engine='openpyxl') as writer:
            st.session_state.df_result.to_excel(writer, index=False, sheet_name='BOQ')
        output_boq.seek(0)
        
        st.download_button(
            label="⬇️ Download Hasil BOQ (Excel)",
            data=output_boq,
            file_name=f"BOQ_{st.session_state.lop_name}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    with tab2:
        st.info("Download seluruh file RAB spreadsheet dari Google Sheets")
        
        if st.button("⬇️ Generate Full RAB Spreadsheet"):
            with st.spinner("Mempersiapkan file. Harap tunggu..."):
                try:
                    # Inisialisasi Google Drive API
                    scope = ['https://www.googleapis.com/auth/drive']
                    creds = ServiceAccountCredentials.from_json_keyfile_dict(
                        st.secrets["gcp_service_account"],
                        scopes=scope
                    )
                    drive_service = build('drive', 'v3', credentials=creds)
                    
                    # Download file
                    request = drive_service.files().export_media(
                        fileId='1Zl0txYzsqslXjGV4Y4mcpVMB-vikTDCauzcLOfbbD5c',
                        mimeType='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
                    )
                    
                    output_rab = BytesIO()
                    downloader = MediaIoBaseDownload(output_rab, request)
                    
                    progress_bar = st.progress(0)
                    done = False
                    while not done:
                        status, done = downloader.next_chunk()
                        progress_bar.progress(int(status.progress() * 100))
                    
                    output_rab.seek(0)
                    
                    st.session_state.downloaded = True
                    st.success("File siap diunduh!")
                    
                    st.download_button(
                        label="💾 Klik untuk Download RAB Lengkap",
                        data=output_rab,
                        file_name=f"RAB_Lengkap_{st.session_state.lop_name}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                    
                except HttpError as error:
                    st.error(f"Google API error: {error}")
                except Exception as e:
                    st.error(f"Terjadi kesalahan: {str(e)}")

# ==============================================
# RESET FORM SETELAH DOWNLOAD
# ==============================================

if st.session_state.get('downloaded', False):
    st.success("🎉 File telah berhasil diunduh!")
    if st.button("🔄 Buat Input Baru"):
        reset_form()
        st.session_state.show_download = False
        st.rerun()
