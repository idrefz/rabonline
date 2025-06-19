import streamlit as st
import pandas as pd
from io import BytesIO

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

    # Upload file RAB Excel
    uploaded_file = st.file_uploader("Upload File RAB Excel", type=["xlsx", "xls"])

    submitted = st.form_submit_button("🚀 Proses BOQ")

# Proses setelah form disubmit
if submitted:
    if not lop_name:
        st.warning("Harap masukkan Nama LOP terlebih dahulu!")
        st.stop()
    
    if not uploaded_file:
        st.warning("Harap upload file RAB Excel terlebih dahulu!")
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
        designators = []
        volumes = []
        
        def add_item(designator, volume):
            if volume > 0 or (designator == "Preliminary Project HRB/Kawasan Khusus" and izin):
                designators.append(designator)
                volumes.append(volume)
        
        # Tambahkan item ke dataframe
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
        
        df_result = pd.DataFrame({"Designator": designators, "Volume": volumes})

        # Baca file RAB yang diupload
        rab_df = pd.read_excel(uploaded_file)
        
        # Gabungkan dengan hasil perhitungan (contoh sederhana)
        # Anda mungkin perlu menyesuaikan logika penggabungan ini
        final_df = pd.merge(rab_df, df_result, on="Designator", how="left")
        
        # Simpan hasil di session state
        st.session_state.df_result = final_df
        st.session_state.download_ready = True
        st.session_state.lop_name = lop_name

        # Tampilkan hasil
        st.success("✅ Perhitungan BOQ Berhasil!")
        st.subheader("📊 Hasil Perhitungan BOQ")
        st.dataframe(df_result.style.highlight_max(axis=0), use_container_width=True)

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
    
    # Buat file Excel dari DataFrame
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        st.session_state.df_result.to_excel(writer, index=False, sheet_name='RAB')
    output.seek(0)
    
    st.download_button(
        label="⬇️ Download RAB Excel",
        data=output,
        file_name=f"RAB_{st.session_state.lop_name}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
    
    if st.button("🔄 Buat Input Baru"):
        st.session_state.df_result = None
        st.session_state.download_ready = False
        st.rerun()
