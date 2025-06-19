import streamlit as st
import pandas as pd
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from io import BytesIO
import json

# Load credentials dari secrets
scope = ["https://spreadsheets.google.com/feeds", 'https://www.googleapis.com/auth/drive']
creds_dict = st.secrets["gcp_service_account"]
creds = ServiceAccountCredentials.from_json_keyfile_dict(creds_dict, scope)
client = gspread.authorize(creds)
sheet = client.open_by_url("https://docs.google.com/spreadsheets/d/1Zl0txYzsqslXjGV4Y4mcpVMB-vikTDCauzcLOfbbD5c/edit").sheet1

st.set_page_config("Form Input BOQ", layout="centered")
st.title("📋 Form Input BOQ Otomatis")

# Pilihan sumber data (ODC atau ODP)
sumber = st.radio("Sumber Data", ["ODC", "ODP"])

# Pilihan Jenis Kabel dan Input Panjang Kabel
jenis_kabel = st.radio("Pilih Jenis Kabel", ["12 Core", "24 Core"])
panjang_kabel = st.number_input("Masukkan Panjang Kabel (meter)", min_value=0.0)

# Pilihan ODP dan Totalnya
jenis_odp = st.radio("Pilih Jenis ODP", ["ODP 8", "ODP 16"])
total_odp = st.number_input("Masukkan Total ODP", min_value=0)

tiang_new = st.number_input("Total Tiang Baru", min_value=0)
tiang_existing = st.number_input("Total Tiang Existing", min_value=0)
tikungan = st.number_input("Jumlah Tikungan", min_value=0)
izin = st.text_input("Nilai Izin (isi jika ada)")
lop_name = st.text_input("Nama LOP (untuk nama file export)")

if st.button("Proses BOQ"):
    vol_kabel = round((panjang_kabel * 1.02) + total_odp)
    kabel_designator = "AC-OF-SM-12-SC_O_STOCK" if jenis_kabel == "12 Core" else "AC-OF-SM-24-SC_O_STOCK"
    odp_designator = "ODP Solid-PB-8 AS" if jenis_odp == "ODP 8" else "ODP Solid-PB-16 AS"

    # PU-AS Logic
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
            data["Volume"].append(round(volume))

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

    # Simpan ke Google Sheet
    values = sheet.get_all_values()
    for i in range(8, len(values)):
        designator_cell = values[i][1]
        match = df[df["Designator"] == designator_cell]
        if not match.empty:
            volume = match["Volume"].values[0]
            if designator_cell == "Preliminary Project HRB/Kawasan Khusus":
                sheet.update_cell(i + 1, 6, int(volume))  # Kolom F
                sheet.update_cell(i + 1, 7, 1)            # Kolom G
            else:
                sheet.update_cell(i + 1, 7, int(volume))

    # Tampilkan tabel hasil BOQ
    st.subheader("Hasil Perhitungan BOQ")
    st.dataframe(df)

    # Total material dan jasa dari Google Sheet
    total_material = sheet.acell("G289").value
    total_jasa = sheet.acell("G290").value
    total_all = sheet.acell("G291").value

    st.markdown(f"**Total Material:** Rp {total_material}")
    st.markdown(f"**Total Jasa:** Rp {total_jasa}")
    st.markdown(f"**Total Keseluruhan:** Rp {total_all}")

    # CPP dan Perizinan
    total_boq_volume = df["Volume"].sum()
    cpp = (total_odp * 8) / total_boq_volume if total_boq_volume else 0

    izin_rows = sheet.col_values(6)[8:]
    izin_count = izin_rows.count("1")
    izin_persen = (izin_count / len(izin_rows)) * 100 if izin_rows else 0

    st.markdown(f"**CPP:** {cpp:.2f}")
    st.markdown(f"**Perizinan:** {izin_persen:.2f}%")

    # Proteksi download
    password = st.text_input("Masukkan password untuk download", type="password")
    if password == "sdibisa":
        output = BytesIO()
        df.to_excel(output, index=False, engine='openpyxl')
        output.seek(0)
        st.download_button("⬇️ Download Excel", output, file_name=f"{lop_name}.xlsx")

        # Reset semua kolom G dan F untuk Preliminary Project HRB/Kawasan Khusus
        for i in range(8, len(values)):
            if values[i][1] == "Preliminary Project HRB/Kawasan Khusus":
                sheet.update_cell(i + 1, 6, "0")
            sheet.update_cell(i + 1, 7, "0")

        st.success("Download selesai dan volume berhasil di-reset.")

    # Tombol reset
    if st.button("🔁 Input Lagi (Reset Volume ke 0)"):
        for i in range(8, len(values)):
            sheet.update_cell(i + 1, 7, "0")
        st.success("Volume berhasil di-reset ke 0.")
