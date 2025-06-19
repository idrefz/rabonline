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

# Initialize all fields to empty/zero if not already in session state
if 'sumber' not in st.session_state:
    reset_form()

# Form input
with st.form("boq_form"):
    # Section 1: Sumber Data (ODC/ODP)
    st.subheader("🔹 Sumber Data")
    sumber = st.radio("Pilih Sumber Data:", ["ODC", "ODP"], key="sumber", 
                     help="Pilihan ini akan mempengaruhi perhitungan kabel, OS, dan komponen lainnya")
    
    # Section 2: Input Kabel
    st.subheader("🔹 Input Kabel")
    col1, col2 = st.columns(2)
    with col1:
        kabel_12 = st.number_input("Panjang Kabel 12 Core (meter)", min_value=0.0, value=st.session_state.kabel_12, key="kabel_12")
    with col2:
        kabel_24 = st.number_input("Panjang Kabel 24 Core (meter)", min_value=0.0, value=st.session_state.kabel_24, key="kabel_24")
    
    # Section 3: Input ODP
    st.subheader("🔹 Input ODP")
    col1, col2 = st.columns(2)
    with col1:
        odp_8 = st.number_input("Jumlah ODP 8", min_value=0, value=st.session_state.odp_8, key="odp_8")
    with col2:
        odp_16 = st.number_input("Jumlah ODP 16", min_value=0, value=st.session_state.odp_16, key="odp_16")
    
    # Section 4: Input Pendukung
    st.subheader("🔹 Input Pendukung")
    tiang_new = st.number_input("Total Tiang Baru", min_value=0, value=st.session_state.tiang_new, key="tiang_new")
    tiang_existing = st.number_input("Total Tiang Existing", min_value=0, value=st.session_state.tiang_existing, key="tiang_existing")
    tikungan = st.number_input("Jumlah Tikungan", min_value=0, value=st.session_state.tikungan, key="tikungan")
    izin = st.text_input("Nilai Izin (isi jika ada)", value=st.session_state.izin, key="izin")
    lop_name = st.text_input("Nama LOP (untuk nama file export)", value=st.session_state.lop_name, key="lop_name")
    
    submitted = st.form_submit_button("🚀 Proses BOQ")
    reset_button = st.form_submit_button("🔄 Reset Form")

if reset_button:
    reset_form()
    st.rerun()

if submitted and not st.session_state.downloaded:
    if not lop_name:
        st.warning("Harap masukkan Nama LOP terlebih dahulu!")
        st.stop()
    
    # Hitung total ODP
    total_odp = odp_8 + odp_16
    
    # 1. PERHITUNGAN VOLUME KABEL
    if sumber == "ODC":
        vol_kabel_12 = round((kabel_12 * 1.02) + total_odp) if kabel_12 > 0 else 0
        vol_kabel_24 = round((kabel_24 * 1.02) + total_odp) if kabel_24 > 0 else 0
    else:  # ODP
        vol_kabel_12 = round(kabel_12 * 1.02) if kabel_12 > 0 else 0
        vol_kabel_24 = round(kabel_24 * 1.02) if kabel_24 > 0 else 0
    
    # 2. PERHITUNGAN PU-AS
    vol_puas = (total_odp * 2 - 1) if total_odp > 1 else (1 if total_odp == 1 else 0)
    vol_puas += tiang_new + tiang_existing + tikungan

    # 3. PERHITUNGAN OS
    if sumber == "ODC":
        os_odc = (12 if kabel_12 > 0 else 24 if kabel_24 > 0 else 0) + total_odp
        os_odp = 0
    else:  # ODP
        os_odc = 0
        os_odp = total_odp * 2
    
    os_total = os_odc + os_odp

    # 4. PERHITUNGAN PC
    pc_upc = (total_odp - 1) // 4 + 1 if total_odp > 0 else 0
    pc_apc = 18 if pc_upc == 1 else (pc_upc * 2 if pc_upc > 1 else 0)

    # 5. PERHITUNGAN LAINNYA
    tc02 = 1 if sumber == "ODC" else 0
    dd40 = 6 if sumber == "ODC" else 0
    bc06 = 6 if sumber == "ODC" else 0
    ps_odc = (total_odp - 1) // 4 + 1 if sumber == "ODC" and total_odp > 0 else 0

    # Membuat DataFrame hasil
    designators = []
    volumes = []
    
    def add_item(designator, volume):
        if volume > 0 or (designator == "Preliminary Project HRB/Kawasan Khusus" and izin):
            designators.append(designator)
            volumes.append(volume)
    
    # Tambahkan semua item ke dataframe
    if kabel_12 > 0:
        add_item("AC-OF-SM-12-SC_O_STOCK", vol_kabel_12)
    if kabel_24 > 0:
        add_item("AC-OF-SM-24-SC_O_STOCK", vol_kabel_24)
    if odp_8 > 0:
        add_item("ODP Solid-PB-8 AS", odp_8)
    if odp_16 > 0:
        add_item("ODP Solid-PB-16 AS", odp_16)
    
    add_item("PU-S7.0-400NM", tiang_new)
    add_item("PU-AS", vol_puas)
    
    if sumber == "ODC":
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
    
    if izin:
        add_item("Preliminary Project HRB/Kawasan Khusus", 1)
    
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

    # Download Options
    st.subheader("💾 Download")
    
    tab1, tab2 = st.tabs(["Download Hasil BOQ", "Download Full RAB"])
    
    with tab1:
        output_boq = BytesIO()
        with pd.ExcelWriter(output_boq, engine='openpyxl') as writer:
            df.to_excel(writer, index=False, sheet_name='BOQ')
        output_boq.seek(0)
        
        st.download_button(
            label="⬇️ Download Hasil BOQ (Excel)",
            data=output_boq,
            file_name=f"BOQ_{lop_name}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            help="Download hasil perhitungan BOQ dalam format Excel"
        )
    
    with tab2:
        st.info("Download seluruh file RAB spreadsheet dari Google Sheets")
        if st.button("⬇️ Generate Full RAB Spreadsheet"):
            output_rab = BytesIO()
            spreadsheet.export(format='xlsx', output=output_rab)
            output_rab.seek(0)
            
            st.download_button(
                label="💾 Klik untuk Download RAB Lengkap",
                data=output_rab,
                file_name=f"RAB_Lengkap_{lop_name}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                on_click=lambda: setattr(st.session_state, 'downloaded', True)
            )

# Reset after download
if st.session_state.downloaded:
    st.success("🎉 File telah berhasil diunduh!")
    if st.button("🔄 Buat Input Baru"):
        # Reset Google Sheet values
        sheet = spreadsheet.sheet1
        values = sheet.get_all_values()
        for i in range(8, len(values)):
            if values[i][1] == "Preliminary Project HRB/Kawasan Khusus":
                sheet.update_cell(i+1, 6, "0")
            sheet.update_cell(i+1, 7, "0")
        reset_form()
        st.rerun()
