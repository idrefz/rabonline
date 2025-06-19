import streamlit as st
import pandas as pd
from io import BytesIO
import openpyxl

# Konfigurasi Aplikasi
st.set_page_config("Form Input BOQ", layout="centered")
st.title("📋 Form Input BOQ Otomatis")

# Inisialisasi session state
if 'df_result' not in st.session_state:
    st.session_state.df_result = None
if 'download_ready' not in st.session_state:
    st.session_state.download_ready = False

# Form Input
with st.form("boq_form"):
    st.subheader("🔹 Data Proyek")
    col1, col2 = st.columns(2)
    with col1:
        sumber = st.radio("Sumber Data:", ["ODC", "ODP"], index=0)
    with col2:
        lop_name = st.text_input("Nama LOP (untuk nama file):")
        project_name = st.text_input("Nama Project:")
        sto_code = st.text_input("Kode STO:")

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

    # Upload file RAB Excel template
    uploaded_file = st.file_uploader("Upload Template RAB Excel", type=["xlsx", "xls"])

    submitted = st.form_submit_button("🚀 Proses BOQ")

# Fungsi untuk mapping designator
def map_designator(kode):
    mapping = {
        "AC-OF-SM-12-SC_O_STOCK": "DC-01-01-1111",
        "AC-OF-SM-24-SC_O_STOCK": "DC-01-04-1100",
        "ODP Solid-PB-8 AS": "DC-01-08-4400",
        "ODP Solid-PB-16 AS": "DC-01-04-0400",
        "PU-S7.0-400NM": "DC-01-04-0410",
        "PU-AS": "DC-01-08-4280",
        "OS-SM-1-ODC": "AC-01-04-1100",
        "TC-02-ODC": "AC-01-04-2400",
        "DD-HDPE-40-1": "AC-01-04-0500",
        "BC-TR-0.6": "DC-01-04-1420",
        "PS-1-4-ODC": "DC-01-04-2420",
        "OS-SM-1-ODP": "DC-01-04-2460",
        "OS-SM-1": "DC-01-04-2480",
        "PC-UPC-652-2": "DC-01-04-2490",
        "PC-APC/UPC-652-A1": "DC-01-04-2500"
    }
    return mapping.get(kode, "")

# Proses setelah form disubmit
if submitted:
    if not lop_name or not project_name or not sto_code:
        st.warning("Harap lengkapi data proyek (Nama LOP, Nama Project, Kode STO)!")
        st.stop()
    
    if not uploaded_file:
        st.warning("Harap upload template RAB Excel terlebih dahulu!")
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

        # Membuat DataFrame hasil
        items = [
            {"Designator": "AC-OF-SM-12-SC_O_STOCK", "Volume": vol_kabel_12 if kabel_12 > 0 else None},
            {"Designator": "AC-OF-SM-24-SC_O_STOCK", "Volume": vol_kabel_24 if kabel_24 > 0 else None},
            {"Designator": "ODP Solid-PB-8 AS", "Volume": odp_8 if odp_8 > 0 else None},
            {"Designator": "ODP Solid-PB-16 AS", "Volume": odp_16 if odp_16 > 0 else None},
            {"Designator": "PU-S7.0-400NM", "Volume": tiang_new if tiang_new > 0 else None},
            {"Designator": "PU-AS", "Volume": vol_puas},
            {"Designator": "OS-SM-1-ODC", "Volume": os_odc if sumber == "ODC" else None},
            {"Designator": "TC-02-ODC", "Volume": tc02 if sumber == "ODC" else None},
            {"Designator": "DD-HDPE-40-1", "Volume": dd40 if sumber == "ODC" else None},
            {"Designator": "BC-TR-0.6", "Volume": bc06 if sumber == "ODC" else None},
            {"Designator": "PS-1-4-ODC", "Volume": ps_odc if sumber == "ODC" else None},
            {"Designator": "OS-SM-1-ODP", "Volume": os_odp if sumber == "ODP" else None},
            {"Designator": "OS-SM-1", "Volume": os_total},
            {"Designator": "PC-UPC-652-2", "Volume": pc_upc},
            {"Designator": "PC-APC/UPC-652-A1", "Volume": pc_apc},
            {"Designator": "Preliminary Project HRB/Kawasan Khusus", "Volume": 1 if izin else None}
        ]

        # Baca template RAB
        wb = openpyxl.load_workbook(uploaded_file)
        ws = wb.active

        # Update header proyek
        ws['B1'] = "DATA MATERIAL SATUAN"
        ws['B2'] = f"PENGADAAN DAN PEMASANGAN GRANULAR MODERNIZATION"
        ws['B3'] = f"PROJECT : {project_name}"
        ws['B4'] = f"STO : {sto_code}"

        # Temukan baris data (mulai dari baris 8 berdasarkan contoh)
        for row in ws.iter_rows(min_row=8, max_row=ws.max_row):
            designator_cell = row[1]  # Kolom B (Designator)
            volume_cell = row[6]     # Kolom G (VOL)
            
            # Cari kode designator yang sesuai
            for item in items:
                if item["Volume"] is not None and str(designator_cell.value).strip() == map_designator(item["Designator"]):
                    volume_cell.value = item["Volume"]
                    break

        # Simpan ke BytesIO
        output = BytesIO()
        wb.save(output)
        output.seek(0)

        # Simpan hasil di session state
        st.session_state.download_data = output
        st.session_state.download_ready = True
        st.session_state.lop_name = lop_name

        # Tampilkan hasil perhitungan
        st.success("✅ Perhitungan BOQ Berhasil!")
        
        # Tampilkan tabel hasil perhitungan
        df_result = pd.DataFrame([item for item in items if item["Volume"] is not None])
        st.subheader("📊 Hasil Perhitungan BOQ")
        st.dataframe(df_result)

        # Tampilkan Total
        st.subheader("📌 Ringkasan")
        col1, col2, col3 = st.columns(3)
        with col1:
            st.metric("Total Kabel (m)", f"{vol_kabel_12 + vol_kabel_24:,}m")
        with col2:
            st.metric("Total ODP", f"{total_odp:,} unit")
        with col3:
            st.metric("Total PU-AS", f"{vol_puas:,} unit")

    except Exception as e:
        st.error(f"Terjadi kesalahan saat memproses: {str(e)}")

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
        st.session_state.df_result = None
        st.session_state.download_ready = False
        st.rerun()
