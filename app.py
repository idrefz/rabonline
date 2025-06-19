import streamlit as st
import pandas as pd
from io import BytesIO
import openpyxl

# Konfigurasi Aplikasi
st.set_page_config("Form Input BOQ", layout="centered")
st.title("📋 Form Input BOQ Otomatis")

# Inisialisasi session state dengan nilai default
if 'download_ready' not in st.session_state:
    st.session_state.update({
        'download_ready': False,
        'download_data': None,
        'lop_name': "",
        'debug_info': []  # Pastikan debug_info selalu ada sebagai list
    })

# Fungsi untuk memetakan designator ke kode RAB
def get_rab_code(designator):
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
        "PC-APC/UPC-652-A1": "DC-01-04-2500",
        "Preliminary Project HRB/Kawasan Khusus": "IZIN-KHUSUS-001"
    }
    return mapping.get(designator, "")

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

    uploaded_file = st.file_uploader("Upload Template RAB Excel", type=["xlsx", "xls"])

    submitted = st.form_submit_button("🚀 Proses BOQ")

if submitted:
    if not all([lop_name, project_name, sto_code]):
        st.warning("Harap lengkapi data proyek!")
        st.stop()
    
    if not uploaded_file:
        st.warning("Harap upload template RAB Excel!")
        st.stop()

    try:
        # Reset debug info
        st.session_state.debug_info = []
        
        # Hitung volume
        total_odp = odp_8 + odp_16
        vol_kabel_12 = round((kabel_12 * 1.02) + (total_odp if sumber == "ODC" else 0)) if kabel_12 > 0 else 0
        vol_kabel_24 = round((kabel_24 * 1.02) + (total_odp if sumber == "ODC" else 0)) if kabel_24 > 0 else 0
        vol_puas = (total_odp * 2 - 1) if total_odp > 1 else (1 if total_odp == 1 else 0)
        vol_puas += tiang_new + tiang_existing + tikungan

        # Daftar item dengan volume
        items = [
            {"designator": "AC-OF-SM-12-SC_O_STOCK", "volume": vol_kabel_12 if kabel_12 > 0 else None},
            {"designator": "AC-OF-SM-24-SC_O_STOCK", "volume": vol_kabel_24 if kabel_24 > 0 else None},
            {"designator": "ODP Solid-PB-8 AS", "volume": odp_8 if odp_8 > 0 else None},
            {"designator": "ODP Solid-PB-16 AS", "volume": odp_16 if odp_16 > 0 else None},
            {"designator": "PU-S7.0-400NM", "volume": tiang_new if tiang_new > 0 else None},
            {"designator": "PU-AS", "volume": vol_puas},
            {"designator": "OS-SM-1-ODC", "volume": (12 if kabel_12 > 0 else 24 if kabel_24 > 0 else 0) + total_odp if sumber == "ODC" else None},
            {"designator": "TC-02-ODC", "volume": 1 if sumber == "ODC" else None},
            {"designator": "DD-HDPE-40-1", "volume": 6 if sumber == "ODC" else None},
            {"designator": "BC-TR-0.6", "volume": 6 if sumber == "ODC" else None},
            {"designator": "PS-1-4-ODC", "volume": (total_odp - 1) // 4 + 1 if sumber == "ODC" and total_odp > 0 else None},
            {"designator": "OS-SM-1-ODP", "volume": total_odp * 2 if sumber == "ODP" else None},
            {"designator": "OS-SM-1", "volume": ((12 if kabel_12 > 0 else 24 if kabel_24 > 0 else 0) + total_odp) if sumber == "ODC" else (total_odp * 2)},
            {"designator": "PC-UPC-652-2", "volume": (total_odp - 1) // 4 + 1 if total_odp > 0 else 0},
            {"designator": "PC-APC/UPC-652-A1", "volume": 18 if ((total_odp - 1) // 4 + 1) == 1 else (((total_odp - 1) // 4 + 1) * 2 if ((total_odp - 1) // 4 + 1) > 1 else 0)},
            {"designator": "Preliminary Project HRB/Kawasan Khusus", "volume": 1 if izin else None}
        ]

        # Baca template
        wb = openpyxl.load_workbook(uploaded_file)
        ws = wb.active
        
        # Debug info
        st.session_state.debug_info.append(f"File loaded: {uploaded_file.name}")
        st.session_state.debug_info.append(f"Sheet name: {ws.title}")
        st.session_state.debug_info.append(f"Max row: {ws.max_row}, Max column: {ws.max_column}")

        # Update header proyek
        ws['B1'] = "DATA MATERIAL SATUAN"
        ws['B2'] = f"PENGADAAN DAN PEMASANGAN GRANULAR MODERNIZATION"
        ws['B3'] = f"PROJECT : {project_name}"
        ws['B4'] = f"STO : {sto_code}"

        # Temukan kolom VOL
        vol_col = None
        header_row = 7  # Asumsi header di row 7
        for col in range(1, ws.max_column + 1):
            cell_value = str(ws.cell(row=header_row, column=col).value).strip().upper()
            if "VOL" in cell_value:
                vol_col = col
                break
        
        if vol_col is None:
            vol_col = 7  # Fallback
            st.warning("Kolom VOL tidak ditemukan, menggunakan kolom 7 sebagai default")

        st.session_state.debug_info.append(f"Kolom VOL: {vol_col}")

        # Proses update volume
        updated_count = 0
        start_row = 8  # Asumsi data mulai dari row 8
        
        for row in range(start_row, ws.max_row + 1):
            designator_cell = ws.cell(row=row, column=2)  # Kolom B
            rab_code = str(designator_cell.value).strip() if designator_cell.value else ""
            
            for item in items:
                if item["volume"] > 0 and rab_code == get_rab_code(item["designator"]):
                    ws.cell(row=row, column=vol_col, value=item["volume"])
                    updated_count += 1
                    st.session_state.debug_info.append(
                        f"Updated row {row}: {item['designator']} = {item['volume']}"
                    )
                    break

        st.session_state.debug_info.append(f"Total updated: {updated_count} items")

        # Simpan ke BytesIO
        output = BytesIO()
        wb.save(output)
        output.seek(0)

        # Simpan untuk download
        st.session_state.download_data = output
        st.session_state.download_ready = True
        st.session_state.lop_name = lop_name

        st.success(f"✅ Berhasil mengupdate {updated_count} item volume!")
        
        # Tampilkan data yang diupdate
        st.subheader("Data yang Diupdate")
        st.table(pd.DataFrame([item for item in items if item['volume'] > 0]))

    except Exception as e:
        st.error(f"Error: {str(e)}")
        if 'debug_info' in st.session_state:
            st.session_state.debug_info.append(f"Error: {str(e)}")

# Tampilkan tombol download jika sudah siap
if st.session_state.get('download_ready', False):
    st.subheader("💾 Download RAB Terupdate")
    st.download_button(
        label="⬇️ Download RAB Excel",
        data=st.session_state.download_data,
        file_name=f"RAB_{st.session_state.lop_name}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

    # Tampilkan debug info dengan pengecekan
    if 'debug_info' in st.session_state and st.session_state.debug_info:
        with st.expander("Debug Information"):
            for info in st.session_state.debug_info:
                st.write(info)

    if st.button("🔄 Buat Input Baru"):
        st.session_state.update({
            'download_ready': False,
            'download_data': None,
            'debug_info': []
        })
        st.rerun()
