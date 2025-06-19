import streamlit as st
import pandas as pd
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from io import BytesIO

# Load credentials from secrets
scope = ["https://spreadsheets.google.com/feeds", 'https://www.googleapis.com/auth/drive']
creds_dict = st.secrets["gcp_service_account"]
creds = ServiceAccountCredentials.from_json_keyfile_dict(creds_dict, scope)
client = gspread.authorize(creds)
spreadsheet = client.open_by_url("https://docs.google.com/spreadsheets/d/1Zl0txYzsqslXjGV4Y4mcpVMB-vikTDCauzcLOfbbD5c/edit")

st.set_page_config("Form Input BOQ", layout="centered")
st.title("📋 Form Input BOQ Otomatis")

# Initialize session state
if 'downloaded' not in st.session_state:
    st.session_state.downloaded = False

def reset_form():
    st.session_state.downloaded = False
    st.session_state.kabel_12 = 0.0
    st.session_state.kabel_24 = 0.0 
    st.session_state.odp_8 = 0
    st.session_state.odp_16 = 0
    st.session_state.tiang_new = 0
    st.session_state.tiang_existing = 0
    st.session_state.tikungan = 0
    st.session_state.izin = ""
    st.session_state.lop_name = ""

# Form input
with st.form("boq_form"):
    st.subheader("Input Kabel")
    kabel_12 = st.number_input("Panjang Kabel 12 Core (meter)", min_value=0.0, value=0.0, key="kabel_12")
    kabel_24 = st.number_input("Panjang Kabel 24 Core (meter)", min_value=0.0, value=0.0, key="kabel_24")
    
    st.subheader("Input ODP")
    odp_8 = st.number_input("Jumlah ODP 8", min_value=0, value=0, key="odp_8")
    odp_16 = st.number_input("Jumlah ODP 16", min_value=0, value=0, key="odp_16")
    
    st.subheader("Input Lainnya")
    tiang_new = st.number_input("Total Tiang Baru", min_value=0, value=0, key="tiang_new")
    tiang_existing = st.number_input("Total Tiang Existing", min_value=0, value=0, key="tiang_existing")
    tikungan = st.number_input("Jumlah Tikungan", min_value=0, value=0, key="tikungan")
    izin = st.text_input("Nilai Izin (isi jika ada)", key="izin")
    lop_name = st.text_input("Nama LOP (untuk nama file export)", key="lop_name")
    
    submitted = st.form_submit_button("Proses BOQ")

if submitted and not st.session_state.downloaded:
    if not lop_name:
        st.warning("Harap masukkan Nama LOP terlebih dahulu!")
        st.stop()
    
    # Hitung total kabel dan ODP
    total_kabel = kabel_12 + kabel_24
    total_odp = odp_8 + odp_16
    
    # Hitung volume
    vol_kabel_12 = round((kabel_12 * 1.02) + odp_8 + odp_16) if kabel_12 > 0 else 0
    vol_kabel_24 = round((kabel_24 * 1.02) + odp_8 + odp_16) if kabel_24 > 0 else 0
    
    # PU-AS Logic
    if total_odp == 0:
        vol_puas = 0
    elif total_odp == 1:
        vol_puas = 1
    else:
        vol_puas = (total_odp * 2 - 1)
    vol_puas += tiang_new + tiang_existing + tikungan

    # Buat dataframe hasil
    data = {
        "Designator": [
            "AC-OF-SM-12-SC_O_STOCK" if kabel_12 > 0 else "",
            "AC-OF-SM-24-SC_O_STOCK" if kabel_24 > 0 else "",
            "ODP Solid-PB-8 AS" if odp_8 > 0 else "",
            "ODP Solid-PB-16 AS" if odp_16 > 0 else "",
            "PU-S7.0-400NM",
            "PU-AS",
            "Preliminary Project HRB/Kawasan Khusus" if izin else ""
        ],
        "Volume": [
            vol_kabel_12,
            vol_kabel_24,
            odp_8,
            odp_16,
            tiang_new,
            vol_puas,
            1 if izin else 0
        ]
    }
    
    df = pd.DataFrame(data)
    df = df[df["Designator"] != ""]  # Hapus baris kosong

    # Update Google Sheet
    sheet = spreadsheet.sheet1
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
    st.subheader("Hasil Perhitungan BOQ")
    st.dataframe(df)

    # Download Options
    st.subheader("Download Options")
    
    # 1. Download Hasil BOQ
    output_boq = BytesIO()
    with pd.ExcelWriter(output_boq, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='BOQ')
    output_boq.seek(0)
    
    st.download_button(
        label="⬇️ Download Hasil BOQ",
        data=output_boq,
        file_name=f"BOQ_{lop_name}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
    
    # 2. Download Seluruh Spreadsheet RAB
    st.markdown("### Download RAB Lengkap")
    st.warning("File ini akan mendownload seluruh spreadsheet RAB sebagai Excel")
    
    if st.button("⬇️ Download Full RAB Spreadsheet"):
        # Download seluruh spreadsheet
        output_rab = BytesIO()
        spreadsheet.export(format='xlsx', output=output_rab)
        output_rab.seek(0)
        
        st.download_button(
            label="Klik untuk Download RAB",
            data=output_rab,
            file_name=f"RAB_Lengkap_{lop_name}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            on_click=lambda: setattr(st.session_state, 'downloaded', True)
        )

# Reset after download
if st.session_state.downloaded:
    st.success("File telah berhasil diunduh!")
    if st.button("🔁 Buat Input Baru"):
        # Reset Google Sheet values
        sheet = spreadsheet.sheet1
        values = sheet.get_all_values()
        for i in range(8, len(values)):
            if values[i][1] == "Preliminary Project HRB/Kawasan Khusus":
                sheet.update_cell(i+1, 6, "0")
            sheet.update_cell(i+1, 7, "0")
        reset_form()
        st.rerun()
