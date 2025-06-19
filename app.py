import streamlit as st
import pandas as pd
from io import BytesIO

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
# FORM INPUT DENGAN SUBMIT BUTTON YANG BENAR
# ==============================================

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

    # PERBAIKAN UTAMA: Gunakan st.form_submit_button() yang benar
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
