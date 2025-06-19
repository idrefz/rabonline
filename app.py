import streamlit as st
import pandas as pd
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import json

scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
creds_dict = st.secrets["gcp_service_account"]
creds = ServiceAccountCredentials.from_json_keyfile_dict(creds_dict, scope)
client = gspread.authorize(creds)

sheet = client.open_by_url("https://docs.google.com/spreadsheets/d/1Zl0txYzsqslXjGV4Y4mcpVMB-vikTDCauzcLOfbbD5c/edit").sheet1


st.set_page_config("Form Input BOQ", layout="centered")
st.title("📋 Form Input BOQ Otomatis")

# INPUT
sumber = st.radio("Sumber Data", ["ODC", "ODP"])
panjang_kabel = st.number_input("Panjang Kabel (meter)", min_value=0.0)
jenis_kabel = st.selectbox("Jenis Kabel", ["12 Core", "24 Core"])
total_odp = st.selectbox("Jumlah ODP", [0, 8, 16])
tiang_new = st.number_input("Total Tiang Baru", min_value=0)
tiang_existing = st.number_input("Total Tiang Existing", min_value=0)
tikungan = st.number_input("Tikungan", min_value=0)
izin = st.text_input("Nilai Izin (isi jika ada)")
lop_name = st.text_input("Nama LOP (untuk nama file export)")

if st.button("Proses BOQ"):
    vol_kabel = round((panjang_kabel * 1.02) + total_odp, 2)
    kabel_designator = "AC-OF-SM-12-SC_O_STOCK" if jenis_kabel == "12 Core" else "AC-OF-SM-24-SC_O_STOCK"
    odp_designator = "ODP Solid-PB-8 AS" if total_odp == 8 else ("ODP Solid-PB-16 AS" if total_odp == 16 else None)
    if total_odp == 0:
        vol_puas = 0
    elif total_odp == 1:
        vol_puas = 1
    else:
        vol_puas = (total_odp * 2 - 1)
    vol_puas += tiang_new + tiang_existing + tikungan
    os_odc = (12 if jenis_kabel == "12 Core" else 24) + total_odp if sumber == "ODC" else 0
    os_odp = total_odp * 2 if sumber == "ODP" else 0
    os_total = os_odc + os_odp
    pc_upc = (total_odp - 1) // 4 + 1 if total_odp else 0
    pc_apc = 18 if pc_upc == 1 else pc_upc * 2 if pc_upc > 1 else 0
    ps_odc = (total_odp - 1) // 4 + 1 if total_odp else 0
    tc02 = 1 if sumber == "ODC" else 0
    dd40 = 6 if sumber == "ODC" else 0
    bc06 = 6 if sumber == "ODC" else 0
    izin_val = 1 if izin else 0
    izin_designator = "Preliminary Project HRB/Kawasan Khusus" if izin else None

    data = {"Designator": [], "Volume": []}
    def add(designator, volume):
        if volume and designator:
            data["Designator"].append(designator)
            data["Volume"].append(volume)

    add(kabel_designator, vol_kabel)
    add(odp_designator, total_odp)
    add("PU-S7.0-400NM", tiang_new)
    add("PU-AS", vol_puas)
    add("OS-SM-1-ODC", os_odc)
    add("OS-SM-1-ODP", os_odp)
    add("OS-SM-1", os_total)
    add("PC-UPC-652-2", pc_upc)
    add("PC-APC/UPC-652-A1", pc_apc)
    add("PS-1-4-ODC", ps_odc)
    add("TC-02-ODC", tc02)
    add("DD-HDPE-40-1", dd40)
    add("BC-TR-0.6", bc06)
    add(izin_designator, izin_val)

    df = pd.DataFrame(data)
    values = sheet.get_all_values()
    for i in range(8, len(values)):
        designator_cell = values[i][1]
        match = df[df["Designator"] == designator_cell]
        if not match.empty:
            volume = match["Volume"].values[0]
            sheet.update_cell(i+1, 7, volume)

    st.success("Data berhasil diproses dan dikirim ke Google Sheet ✅")
    st.download_button("⬇️ Download Excel", df.to_excel(index=False, engine='openpyxl'), file_name=f"{lop_name}.xlsx")
    for i in range(8, len(values)):
        sheet.update_cell(i+1, 7, "0")
