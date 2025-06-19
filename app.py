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
    st.session_state.sumber = "ODC"
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
    # Pilihan sumber data
    sumber = st.radio("Sumber Data", ["ODC", "ODP"], key="sumber")
    
    st.subheader("Input Kabel")
    col1, col2 = st.columns(2)
    with col1:
        kabel_12 = st.number_input("Panjang Kabel 12 Core (meter)", min_value=0.0, value=0.0, key="kabel_12")
    with col2:
        kabel_24 = st.number_input("Panjang Kabel 24 Core (meter)", min_value=0.0, value=0.0, key="kabel_24")
    
    st.subheader("Input ODP")
    col1, col2 = st.columns(2)
    with col1:
        odp_8 = st.number_input("Jumlah ODP 8", min_value=0, value=0, key="odp_8")
    with col2:
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
    
    # Hitung volume kabel dengan sumber data ODC/ODP
    if sumber == "ODC":
        vol_kabel_12 = round((kabel_12 * 1.02) + total_odp) if kabel_12 > 0 else 0
        vol_kabel_24 = round((kabel_24 * 1.02) + total_odp) if kabel_24 > 0 else 0
    else:  # ODP
        vol_kabel_12 = round(kabel_12 * 1.02) if kabel_12 > 0 else 0
        vol_kabel_24 = round(kabel_24 * 1.02) if kabel_24 > 0 else 0
    
    # PU-AS Logic
    if total_odp == 0:
        vol_puas = 0
    elif total_odp == 1:
        vol_puas = 1
    else:
        vol_puas = (total_odp * 2 - 1)
    vol_puas += tiang_new + tiang_existing + tikungan

    # Hitung OS berdasarkan sumber data
    if sumber == "ODC":
        os_odc = (12 if kabel_12 > 0 else 24 if kabel_24 > 0 else 0) + total_odp
        os_odp = 0
    else:  # ODP
        os_odc = 0
        os_odp = total_odp * 2
    
    os_total = os_odc + os_odp

    # Buat dataframe hasil
    designators = []
    volumes = []
    
    def add_designator(designator, volume):
        if volume > 0 or (designator == "Preliminary Project HRB/Kawasan Khusus" and izin):
            designators.append(designator)
            volumes.append(volume)
    
    if kabel_12 > 0:
        add_designator("AC-OF-SM-12-SC_O_STOCK", vol_kabel_12)
    if kabel_24 > 0:
        add_designator("AC-OF-SM-24-SC_O_STOCK", vol_kabel_24)
    if odp_8 > 0:
        add_designator("ODP Solid-PB-8 AS", odp_8)
    if odp_16 > 0:
        add_designator("ODP Solid-PB-16 AS", odp_16)
    
    add_designator("PU-S7.0-400NM", tiang_new)
    add_designator("PU-AS", vol_puas)
    
    if sumber == "ODC":
        add_designator("OS-SM-1-ODC", os_odc)
    else:
        add_designator("OS-SM-1-ODP", os_odp)
    
    add_designator("OS-SM-1", os_total)
    
    if izin:
        add_designator("Preliminary Project HRB/Kawasan Khusus", 1)
    
    df = pd.DataFrame({"Designator": designators, "Volume": volumes})

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
    
    col1, col2 = st.columns(2)
    
    with col1:
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
    
    with col2:
        # 2. Download Seluruh Spreadsheet RAB
        st.warning("Download seluruh RAB spreadsheet")
        
        if st.button("⬇️ Download Full RAB"):
            output_rab = BytesIO()
            spreadsheet.export(format='xlsx', output=output_rab)
            output_rab.seek(0)
            
            st.download_button(
                label="Klik untuk Download",
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
