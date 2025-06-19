import streamlit as st
import pandas as pd
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from io import BytesIO

# Konfigurasi Aplikasi
st.set_page_config("Form Input BOQ", layout="centered")
st.title("📋 Form Input BOQ Otomatis")

# Fungsi untuk koneksi Google Sheets
def connect_to_gsheet():
    scope = [
        "https://spreadsheets.google.com/feeds",
        "https://www.googleapis.com/auth/drive"
    ]
    creds = ServiceAccountCredentials.from_json_keyfile_dict(st.secrets["gcp_service_account"], scope)
    client = gspread.authorize(creds)
    return client

# Inisialisasi session state
if 'submitted' not in st.session_state:
    st.session_state.submitted = False
    st.session_state.download_ready = False

# URL spreadsheet Google Sheets
SPREADSHEET_URL = "https://docs.google.com/spreadsheets/d/1Zl0txYzsqslXjGV4Y4mcpVMB-vikTDCauzcLOfbbD5c/edit"

# Form Input
with st.form("boq_form"):
    st.subheader("🔹 Data Proyek")
    col1, col2 = st.columns(2)
    with col1:
        sumber = st.radio("Sumber Data:", ["ODC", "ODP"], index=0)
    with col2:
        lop_name = st.text_input("Nama LOP (untuk nama file):")

    st.subheader("🔹 Input Kabel")
    col1, col2 = st.columns(2)
    with col1:
        kabel_12 = st.number_input("Panjang Kabel 12 Core (meter):", min_value=0.0, value=0.0)
    with col2:
        kabel_24 = st.number_input("Panjang Kabel 24 Core (meter):", min_value=0.0, value=0.0)

    st.subheader("🔹 Input ODP")
    col1, col2 = st.columns(2)
    with col1:
        odp_8 = st.number_input("ODP 8 Port:", min_value=0, value=0)
    with col2:
        odp_16 = st.number_input("ODP 16 Port:", min_value=0, value=0)

    st.subheader("🔹 Input Pendukung")
    tiang_new = st.number_input("Tiang Baru:", min_value=0, value=0)
    tiang_existing = st.number_input("Tiang Existing:", min_value=0, value=0)
    tikungan = st.number_input("Jumlah Tikungan:", min_value=0, value=0)
    izin = st.text_input("Nilai Izin (jika ada):", value="")

    submitted = st.form_submit_button("🚀 Proses dan Update RAB")

# Proses setelah form disubmit
if submitted:
    if not lop_name:
        st.warning("Harap masukkan Nama LOP terlebih dahulu!")
        st.stop()
    
    try:
        # Hitung total ODP
        total_odp = odp_8 + odp_16
        
        # 1. Perhitungan Volume Kabel
        if sumber == "ODC":
            vol_kabel_12 = round((kabel_12 * 1.02) + total_odp) if kabel_12 > 0 else 0
            vol_kabel_24 = round((kabel_24 * 1.02) + total_odp) if kabel_24 > 0 else 0
        else:  # ODP
            vol_kabel_12 = round(kabel_12 * 1.02) if kabel_12 > 0 else 0
            vol_kabel_24 = round(kabel_24 * 1.02) if kabel_24 > 0 else 0
        
        # 2. Perhitungan PU-AS
        vol_puas = (total_odp * 2 - 1) if total_odp > 1 else (1 if total_odp == 1 else 0)
        vol_puas += tiang_new + tiang_existing + tikungan

        # 3. Perhitungan OS
        if sumber == "ODC":
            os_odc = (12 if kabel_12 > 0 else 24 if kabel_24 > 0 else 0) + total_odp
            os_odp = 0
        else:  # ODP
            os_odc = 0
            os_odp = total_odp * 2
        
        os_total = os_odc + os_odp

        # 4. Perhitungan PC
        pc_upc = (total_odp - 1) // 4 + 1 if total_odp > 0 else 0
        pc_apc = 18 if pc_upc == 1 else (pc_upc * 2 if pc_upc > 1 else 0)

        # 5. Perhitungan Lainnya
        tc02 = 1 if sumber == "ODC" else 0
        dd40 = 6 if sumber == "ODC" else 0
        bc06 = 6 if sumber == "ODC" else 0
        ps_odc = (total_odp - 1) // 4 + 1 if sumber == "ODC" and total_odp > 0 else 0

        # Mapping designator dengan volume
        items = {
            "AC-OF-SM-12-SC_O_STOCK": vol_kabel_12 if kabel_12 > 0 else None,
            "AC-OF-SM-24-SC_O_STOCK": vol_kabel_24 if kabel_24 > 0 else None,
            "ODP Solid-PB-8 AS": odp_8 if odp_8 > 0 else None,
            "ODP Solid-PB-16 AS": odp_16 if odp_16 > 0 else None,
            "PU-S7.0-400NM": tiang_new if tiang_new > 0 else None,
            "PU-AS": vol_puas,
            "OS-SM-1-ODC": os_odc if sumber == "ODC" else None,
            "TC-02-ODC": tc02 if sumber == "ODC" else None,
            "DD-HDPE-40-1": dd40 if sumber == "ODC" else None,
            "BC-TR-0.6": bc06 if sumber == "ODC" else None,
            "PS-1-4-ODC": ps_odc if sumber == "ODC" else None,
            "OS-SM-1-ODP": os_odp if sumber == "ODP" else None,
            "OS-SM-1": os_total,
            "PC-UPC-652-2": pc_upc,
            "PC-APC/UPC-652-A1": pc_apc,
            "Preliminary Project HRB/Kawasan Khusus": 1 if izin else None
        }

        # Koneksi ke Google Sheets
        client = connect_to_gsheet()
        spreadsheet = client.open_by_url(SPREADSHEET_URL)
        sheet = spreadsheet.sheet1
        
        # Update nilai di Google Sheets
        records = sheet.get_all_records()
        
        for i, row in enumerate(records, start=2):  # Mulai dari baris 2 (indeks 1)
            designator = row['Designator']
            if designator in items and items[designator] is not None:
                # Update kolom Volume (asumsi kolom volume adalah kolom 2)
                sheet.update_cell(i, 2, items[designator])
        
        st.success("✅ Data berhasil diupdate di RAB Spreadsheet!")
        
        # Download file
        output = BytesIO()
        spreadsheet.export(format='xlsx', output=output)
        output.seek(0)
        
        st.session_state.download_data = output
        st.session_state.download_ready = True
        st.session_state.lop_name = lop_name
        
    except Exception as e:
        st.error(f"Terjadi kesalahan: {str(e)}")

# Tampilkan tombol download jika sudah siap
if st.session_state.get('download_ready', False):
    st.subheader("💾 Download RAB Terupdate")
    st.download_button(
        label="⬇️ Download RAB Excel",
        data=st.session_state.download_data,
        file_name=f"RAB_{st.session_state.lop_name}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
    
    if st.button("🔄 Buat Input Baru"):
        st.session_state.submitted = False
        st.session_state.download_ready = False
        st.rerun()
