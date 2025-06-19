import streamlit as st
import pandas as pd
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload
from io import BytesIO
from urllib.parse import urlparse

# Konfigurasi Aplikasi
st.set_page_config("Form Input BOQ", layout="centered")
st.title("📋 Form Input BOQ Otomatis")

# 1. FUNGSI INISIALISASI ======================================================

def init_google_services():
    """Inisialisasi koneksi ke Google Sheets dan Drive API"""
    try:
        # Scope yang diperlukan
        scope = [
            'https://spreadsheets.google.com/feeds',
            'https://www.googleapis.com/auth/drive',
            'https://www.googleapis.com/auth/spreadsheets'
        ]
        
        # Load credentials dari secrets Streamlit
        creds_dict = st.secrets["gcp_service_account"]
        creds = ServiceAccountCredentials.from_json_keyfile_dict(creds_dict, scope)
        
        # Inisialisasi klien
        gc = gspread.authorize(creds)
        drive_service = build('drive', 'v3', credentials=creds)
        
        return gc, drive_service
    except Exception as e:
        st.error(f"Gagal menginisialisasi Google Services: {str(e)}")
        st.stop()

# Inisialisasi koneksi
gc, drive_service = init_google_services()

# 2. FUNGSI UTILITAS =========================================================

def extract_sheet_id(url):
    """Ekstrak ID spreadsheet dari URL"""
    try:
        parsed = urlparse(url)
        if 'docs.google.com' in parsed.netloc:
            path_parts = parsed.path.split('/')
            if 'd' in path_parts:
                return path_parts[path_parts.index('d')+1]
        return url.split('/d/')[1].split('/')[0]
    except:
        st.error("Format URL tidak valid")
        return None

def reset_form():
    """Reset semua nilai form ke default"""
    st.session_state.update({
        'downloaded': False,
        'sumber': "ODC",
        'kabel_12': 0.0,
        'kabel_24': 0.0,
        'odp_8': 0,
        'odp_16': 0,
        'tiang_new': 0,
        'tiang_existing': 0,
        'tikungan': 0,
        'izin': "",
        'lop_name': ""
    })

# 3. KONFIGURASI SPREADSHEET =================================================

SPREADSHEET_URL = "https://docs.google.com/spreadsheets/d/1Zl0txYzsqslXjGV4Y4mcpVMB-vikTDCauzcLOfbbD5c/edit"
SPREADSHEET_ID = extract_sheet_id(SPREADSHEET_URL)

# 4. FORM INPUT ==============================================================

# Inisialisasi session state
if 'sumber' not in st.session_state:
    reset_form()

with st.form("boq_form"):
    st.subheader("🔹 Data Proyek")
    col1, col2 = st.columns(2)
    with col1:
        st.session_state.sumber = st.radio(
            "Sumber Data:", 
            ["ODC", "ODP"],
            index=0,
            key='sumber'
        )
    with col2:
        st.session_state.lop_name = st.text_input(
            "Nama LOP (untuk nama file):",
            key='lop_name'
        )
    
    st.subheader("🔹 Input Kabel")
    col1, col2 = st.columns(2)
    with col1:
        st.session_state.kabel_12 = st.number_input(
            "Panjang Kabel 12 Core (meter):",
            min_value=0.0,
            value=st.session_state.kabel_12,
            key='kabel_12'
        )
    with col2:
        st.session_state.kabel_24 = st.number_input(
            "Panjang Kabel 24 Core (meter):",
            min_value=0.0,
            value=st.session_state.kabel_24,
            key='kabel_24'
        )
    
    st.subheader("🔹 Input ODP")
    col1, col2 = st.columns(2)
    with col1:
        st.session_state.odp_8 = st.number_input(
            "ODP 8 Port:",
            min_value=0,
            value=st.session_state.odp_8,
            key='odp_8'
        )
    with col2:
        st.session_state.odp_16 = st.number_input(
            "ODP 16 Port:",
            min_value=0,
            value=st.session_state.odp_16,
            key='odp_16'
        )
    
    st.subheader("🔹 Input Pendukung")
    st.session_state.tiang_new = st.number_input(
        "Tiang Baru:",
        min_value=0,
        value=st.session_state.tiang_new,
        key='tiang_new'
    )
    st.session_state.tiang_existing = st.number_input(
        "Tiang Existing:",
        min_value=0,
        value=st.session_state.tiang_existing,
        key='tiang_existing'
    )
    st.session_state.tikungan = st.number_input(
        "Jumlah Tikungan:",
        min_value=0,
        value=st.session_state.tikungan,
        key='tikungan'
    )
    st.session_state.izin = st.text_input(
        "Nilai Izin (jika ada):",
        value=st.session_state.izin,
        key='izin'
    )
    
    col1, col2 = st.columns(2)
    with col1:
        submitted = st.form_submit_button("🚀 Proses BOQ")
    with col2:
        if st.form_submit_button("🔄 Reset Form"):
            reset_form()
            st.rerun()

# 5. PROSES PERHITUNGAN =====================================================

if submitted and not st.session_state.downloaded:
    if not st.session_state.lop_name:
        st.warning("Harap masukkan Nama LOP terlebih dahulu!")
        st.stop()
    
    try:
        # Buka spreadsheet
        spreadsheet = gc.open_by_url(SPREADSHEET_URL)
        sheet = spreadsheet.sheet1
        
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

        # Update Google Sheet
        values = sheet.get_all_values()
        
        for i in range(8, len(values)):
            designator = values[i][1]
            if designator in df["Designator"].values:
                volume = df[df["Designator"] == designator]["Volume"].values[0]
                if designator == "Preliminary Project HRB/Kawasan Khusus":
                    sheet.update_cell(i+1, 6, int(volume))  # Kolom F
                    sheet.update_cell(i+1, 7, 1)           # Kolom G
                else:
                    sheet.update_cell(i+1, 7, int(volume))

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

        # Set session state
        st.session_state.df_result = df
        st.session_state.show_download = True

    except Exception as e:
        st.error(f"Terjadi kesalahan saat memproses: {str(e)}")

# 6. DOWNLOAD OPTIONS ========================================================

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
        # Download full spreadsheet menggunakan Drive API
        st.info("Download seluruh file RAB spreadsheet dari Google Sheets")
        
        if st.button("⬇️ Generate Full RAB Spreadsheet"):
            with st.spinner("Mempersiapkan file. Harap tunggu..."):
                try:
                    request = drive_service.files().export_media(
                        fileId=SPREADSHEET_ID,
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
                    
                except Exception as e:
                    st.error(f"Gagal mendownload spreadsheet: {str(e)}")

# 7. RESET FORM =============================================================

if st.session_state.get('downloaded', False):
    st.success("🎉 File telah berhasil diunduh!")
    if st.button("🔄 Buat Input Baru"):
        # Reset Google Sheet values
        try:
            spreadsheet = gc.open_by_url(SPREADSHEET_URL)
            sheet = spreadsheet.sheet1
            values = sheet.get_all_values()
            
            for i in range(8, len(values)):
                if values[i][1] == "Preliminary Project HRB/Kawasan Khusus":
                    sheet.update_cell(i+1, 6, "0")
                sheet.update_cell(i+1, 7, "0")
                
            reset_form()
            st.session_state.show_download = False
            st.rerun()
            
        except Exception as e:
            st.error(f"Gagal mereset spreadsheet: {str(e)}")
