import streamlit as st
import pandas as pd
from io import BytesIO
import openpyxl
import math

# ======================
# 🎩 CONFIGURATION
# ======================
st.set_page_config("BOQ Generator", layout="centered")
st.markdown("""
    <style>
        .block-container {
            padding-top: 2rem;
            padding-bottom: 2rem;
            padding-left: 3rem;
            padding-right: 3rem;
        }
        .stRadio > div {
            flex-direction: row;
        }
        .metric {
            text-align: center;
        }
        .metric .st-emotion-cache-1xarl3l {
            font-size: 1.2rem !important;
        }
        .disabled-label {
            color: #6c757d;
        }
    </style>
""", unsafe_allow_html=True)

st.title("📊 BOQ Generator (Custom Rules)")

# ======================
# 🔄 STATE MANAGEMENT
# ======================
def initialize_session_state():
    """Initialize all session state variables"""
    if 'form_values' not in st.session_state:
        st.session_state.form_values = {
            'lop_name': "",
            'sumber': "ODC",
            'kabel_12': 0.0,
            'kabel_24': 0.0,
            'adss_12': 0.0,
            'adss_24': 0.0,
            'odp_8': 0,
            'odp_16': 0,
            'total_tiang': 0,
            'tiang_new': 0,
            'tiang_existing': 0,
            'tikungan': 0,
            'izin': "",
            'posisi_odp': [],
            'posisi_belokan': [],
            'jumlah_closure':0,
            'uploaded_file': None
        }
    
    if 'boq_state' not in st.session_state:
        st.session_state.boq_state = {
            'ready': False,
            'excel_data': None,
            'project_name': "",
            'updated_items': [],
            'summary': {}
        }

def reset_application():
    """Reset the entire application state"""
    st.session_state.form_values = {
        'lop_name': "",
        'sumber': "ODC",
        'kabel_12': 0.0,
        'kabel_24': 0.0,
        'adss_12': 0.0,
        'adss_24': 0.0,
        'odp_8': 0,
        'odp_16': 0,
        'total_tiang': 0,
        'tiang_new': 0,
        'tiang_existing': 0,
        'tikungan': 0,
        'izin': "",
        'posisi_odp': [],
        'posisi_belokan': [],
        'jumlah_closure':0,
        'uploaded_file': None
    }
    st.session_state.boq_state = {
        'ready': False,
        'excel_data': None,
        'project_name': "",
        'updated_items': [],
        'summary': {}
    }

initialize_session_state()

# ======================
# 🔧 CORE FUNCTIONS
# ======================
def hitung_puas_hl(n_tiang, source='ODC'):
    """Perhitungan khusus PU-AS-HL untuk ADSS"""
    return 16 if source == 'ODC' else 15

def hitung_puas_sc():
    """Perhitungan khusus PU-AS-SC untuk ADSS"""
    return 3

def calculate_volumes(inputs):
    """Calculate all required volumes based on input parameters"""
    total_odp = inputs['odp_8'] + inputs['odp_16']
    
    # Tiang calculation (same for both ADSS and non-ADSS)
    tiang_new = inputs['tiang_new']  # Always track new poles
    tiang_existing = inputs['tiang_existing']
    total_tiang = tiang_new + tiang_existing

    # Volume kabel
    vol_kabel_12 = round(inputs['kabel_12'] * 1.02) if inputs['kabel_12'] > 0 else 0
    vol_kabel_24 = round(inputs['kabel_24'] * 1.02) if inputs['kabel_24'] > 0 else 0
    vol_adss_12 = round(inputs['adss_12'] * 1.02) if inputs['adss_12'] > 0 else 0
    vol_adss_24 = round(inputs['adss_24'] * 1.02) if inputs['adss_24'] > 0 else 0

    # PU-AS atau PU-AS-HL/SC
    if inputs['adss_12'] > 0 or inputs['adss_24'] > 0:
        vol_puas_hl = hitung_puas_hl(total_tiang, inputs['sumber'])
        vol_puas_sc = hitung_puas_sc()
        vol_puas = 0
    else:
        vol_puas = max(0, (total_odp * 2) - 1 + total_tiang + inputs['tikungan'])
        vol_puas_hl = 0
        vol_puas_sc = 0

    # PS-1-4-ODC (1 untuk setiap 4 ODP, minimal 1 jika ada ODP)
    vol_ps_1_4_odc = max(1, math.ceil(total_odp / 4)) if total_odp > 0 and inputs['sumber'] == "ODC" else 0

    # Closure
    vol_closure = inputs.get('jumlah_closure', 0)

    # OS-SM-1
    vol_os_sm_1_odc = total_odp * 2 if inputs['sumber'] == "ODC" else 0
    vol_os_sm_1_odp = total_odp * 2 if inputs['sumber'] == "ODP" else 0
    vol_os_sm_1 = vol_os_sm_1_odc + vol_os_sm_1_odp

    # Base Tray
    vol_base_tray_odc = 0
    if inputs['sumber'] == "ODC":
        if inputs['kabel_12'] > 0:
            vol_base_tray_odc = 1
        elif inputs['kabel_24'] > 0:
            vol_base_tray_odc = 2

    # Connectors
    vol_pc_upc = max(1, math.ceil(total_odp / 4)) if total_odp > 0 else 0
    vol_pc_apc = 18 if vol_pc_upc == 1 else vol_pc_upc * 2 if vol_pc_upc > 1 else 0

    # Komponen lain
    vol_tc_02_odc = 1 if inputs['sumber'] == "ODC" else 0
    vol_dd_hdpe = 6 if inputs['sumber'] == "ODC" else 0
    vol_bc_tr = 3 if inputs['sumber'] == "ODC" else 0

    return [
        {"designator": "AC-OF-SM-12-SC_O_STOCK", "volume": vol_kabel_12},
        {"designator": "AC-OF-SM-24-SC_O_STOCK", "volume": vol_kabel_24},
        {"designator": "AC-OF-SM-ADSS-12D", "volume": vol_adss_12},
        {"designator": "AC-OF-SM-ADSS-24D", "volume": vol_adss_24},
        {"designator": "ODP Solid-PB-8 AS", "volume": inputs['odp_8']},
        {"designator": "ODP Solid-PB-16 AS", "volume": inputs['odp_16']},
        {"designator": "PU-S7.0-400NM", "volume": tiang_new},  # Always included
        {"designator": "PU-AS", "volume": vol_puas},
        {"designator": "PU-AS-HL", "volume": vol_puas_hl},
        {"designator": "PU-AS-SC", "volume": vol_puas_sc},
        {"designator": "PS-1-4-ODC", "volume": vol_ps_1_4_odc},
        {"designator": "OS-SM-1-ODC", "volume": vol_os_sm_1_odc},
        {"designator": "OS-SM-1-ODP", "volume": vol_os_sm_1_odp},
        {"designator": "OS-SM-1", "volume": vol_os_sm_1},
        {"designator": "PC-UPC-652-2", "volume": vol_pc_upc},
        {"designator": "PC-APC/UPC-652-A1", "volume": vol_pc_apc},
        {"designator": "TC-02-ODC", "volume": vol_tc_02_odc},
        {"designator": "DD-HDPE-40-1", "volume": vol_dd_hdpe},
        {"designator": "BC-TR-0.6", "volume": vol_bc_tr},
        {"designator": "Base Tray ODC", "volume": vol_base_tray_odc},
        {"designator": "SC-OF-SM-24", "volume": vol_closure},
        {"designator": "Preliminary Project HRB/Kawasan Khusus", 
         "volume": 1 if inputs['izin'] else 0, 
         "izin_value": float(inputs['izin']) if inputs['izin'] else 0}
    ]

def process_boq_template(uploaded_file, inputs, lop_name):
    """Process the BOQ template file and calculate all metrics"""
    try:
        wb = openpyxl.load_workbook(uploaded_file)
        ws = wb.active
        items = calculate_volumes(inputs)

        updated_count = 0
        for row in range(9, 289):
            designator = str(ws[f'B{row}'].value or "").strip()

            # Handle preliminary project entry
            if inputs['izin'] and designator == "" and "Preliminary Project HRB/Kawasan Khusus" not in [str(ws[f'B{r}'].value) for r in range(9, 289)]:
                ws[f'B{row}'] = "Preliminary Project HRB/Kawasan Khusus"
                ws[f'F{row}'] = float(inputs['izin'])
                ws[f'G{row}'] = 1
                updated_count += 1
                continue

            # Update existing items
            for item in items:
                if item["volume"] > 0 and designator == item["designator"]:
                    ws[f'G{row}'] = item["volume"]
                    if designator == "Preliminary Project HRB/Kawasan Khusus":
                        ws[f'F{row}'] = item.get("izin_value", 0)
                    updated_count += 1
                    break
          # Closure - ambil dari inputs bukan variabel langsung
        vol_closure = inputs['jumlah_closure']
        # Calculate material, jasa, and total costs
        material = jasa = 0.0
        for row in range(9, 289):
            try:
                h_mat = float(ws[f'E{row}'].value or 0)
                h_jasa = float(ws[f'F{row}'].value or 0)
                vol = float(ws[f'G{row}'].value or 0)
                material += h_mat * vol
                jasa += h_jasa * vol
            except:
                continue

        total = material + jasa
        total_odp = inputs['odp_8'] + inputs['odp_16']
        cpp = round((total / (total_odp * 8)), 2) if (total_odp * 8) > 0 else 0

        # Prepare output file
        output = BytesIO()
        wb.save(output)
        output.seek(0)

        return {
            'excel_data': output,
            'updated_count': updated_count,
            'summary': {
                'material': material,
                'jasa': jasa,
                'total': total,
                'cpp': cpp,
                'total_odp': total_odp,
                'total_ports': total_odp * 8
            },
            'updated_items': [item for item in items if item['volume'] > 0]
        }

    except Exception as e:
        st.error(f"Error processing template: {str(e)}")
        return None

# ======================
# 🖥️ FORM UI
# ======================
with st.form("boq_form"):
    st.subheader("📁 Informasi Proyek")
    col1, col2 = st.columns([2, 1])
    with col1:
        lop_name = st.text_input(
            "Nama LOP*",
            value=st.session_state.form_values['lop_name'],
            help="Contoh: LOP_JAKARTA_123"
        )
    with col2:
        sumber = st.radio(
            "Sumber*",
            ["ODC", "ODP"],
            index=0 if st.session_state.form_values['sumber'] == "ODC" else 1,
            horizontal=True
        )

    st.subheader("📦 Kebutuhan Kabel")
    col1, col2 = st.columns(2)
    with col1:
        kabel_12 = st.number_input(
            "12 Core STOCK (meter)",
            min_value=0.0,
            value=st.session_state.form_values['kabel_12'],
            step=1.0,
            format="%.1f"
        )
        adss_12 = st.number_input(
            "ADSS 12 Core (meter)",
            min_value=0.0,
            value=st.session_state.form_values['adss_12'],
            step=1.0,
            format="%.1f"
        )
    with col2:
        kabel_24 = st.number_input(
            "24 Core STOCK (meter)",
            min_value=0.0,
            value=st.session_state.form_values['kabel_24'],
            step=1.0,
            format="%.1f"
        )
        adss_24 = st.number_input(
            "ADSS 24 Core (meter)",
            min_value=0.0,
            value=st.session_state.form_values['adss_24'],
            step=1.0,
            format="%.1f"
        )

    st.subheader("📊 ODP & Tiang")
    col1, col2 = st.columns(2)
    with col1:
        odp_8 = st.number_input(
            "ODP 8 Port*",
            min_value=0,
            value=st.session_state.form_values['odp_8']
        )
    with col2:
        odp_16 = st.number_input(
            "ODP 16 Port*",
            min_value=0,
            value=st.session_state.form_values['odp_16']
        )

    st.subheader("📍 Konfigurasi Tiang")
    col1, col2 = st.columns(2)
    with col1:
        total_tiang = st.number_input(
            "Total Tiang (ADSS)*",
            min_value=0,
            value=st.session_state.form_values['total_tiang'],
            help="Digunakan untuk perhitungan ADSS"
        )
        
        pos_odp_raw = st.text_input(
            "Posisi ODP (contoh: 5,9,14)", 
            value=",".join(map(str, st.session_state.form_values['posisi_odp'])),
            help="Wajib diisi untuk ADSS"
        )
    
    with col2:
        st.markdown('<p class="disabled-label">Untuk Kabel STOCK:</p>', unsafe_allow_html=True)
        tiang_new = st.number_input(
            "Tiang Baru",
            min_value=0,
            value=st.session_state.form_values['tiang_new'],
            help="Digunakan untuk kabel STOCK"
        )
        tiang_existing = st.number_input(
            "Tiang Eksisting",
            min_value=0,
            value=st.session_state.form_values['tiang_existing'],
            help="Digunakan untuk kabel STOCK"
        )
        
        pos_belokan_raw = st.text_input(
            "Posisi Tikungan (contoh: 7,13)", 
            value=",".join(map(str, st.session_state.form_values['posisi_belokan'])),
            help="Wajib diisi untuk ADSS"
        )

    st.subheader("⚙️ Konfigurasi Tambahan")
    tikungan = st.number_input(
        "Jumlah Tikungan*",
        min_value=0,
        value=st.session_state.form_values['tikungan']
    )
    jumlah_closure = st.number_input(
        "Jumlah Closure",
        min_value=0,
        value=st.session_state.form_values.get('jumlah_closure', 0),
        help="Jumlah closure yang digunakan"
    )
    izin = st.text_input(
        "Preliminary Project (Rp)",
        value=st.session_state.form_values['izin'],
        help="Contoh: 500000"
    )
    st.subheader("📤 Template File")
    uploaded_file = st.file_uploader(
        "Unggah Template BOQ*",
        type=["xlsx"],
        help="Format file harus .xlsx"
    )

    submitted = st.form_submit_button("🚀 Generate BOQ", use_container_width=True)

# ======================
# 🚀 FORM PROCESSING
# ======================
if submitted:
    # Validasi input
    if not uploaded_file:
        st.error("Harap unggah file template BOQ!")
        st.stop()
    if not lop_name:
        st.error("Harap isi nama LOP!")
        st.stop()
    # Validasi tiang baru
    if tiang_new < 0:
        st.error("Jumlah tiang baru tidak boleh negatif!")
        st.stop()
    
    # Validasi closure
    if jumlah_closure < 0:
        st.error("Jumlah closure tidak boleh negatif!")
        st.stop()
    # Validasi kabel
    is_adss = adss_12 > 0 or adss_24 > 0
    is_stock = kabel_12 > 0 or kabel_24 > 0
    
    if is_adss and is_stock:
        st.error("Pilih hanya satu jenis kabel (STOCK atau ADSS)!")
        st.stop()
    if not is_adss and not is_stock:
        st.error("Harap pilih minimal satu jenis kabel!")
        st.stop()
    
    # Validasi khusus ADSS
    if is_adss:
        if total_tiang == 0:
            st.error("Untuk ADSS, total tiang harus diisi!")
            st.stop()
        try:
            posisi_odp = [int(x.strip()) for x in pos_odp_raw.split(',') if x.strip().isdigit()]
            posisi_belokan = [int(x.strip()) for x in pos_belokan_raw.split(',') if x.strip().isdigit()]
        except:
            st.error("Format posisi tidak valid! Gunakan format contoh: 5,9,14")
            st.stop()
    else:
        posisi_odp = []
        posisi_belokan = []
        if (tiang_new + tiang_existing) == 0:
            st.error("Untuk STOCK, harap isi tiang baru atau eksisting!")
            st.stop()

    # Update session state
    st.session_state.form_values = {
        'lop_name': lop_name,
        'sumber': sumber,
        'kabel_12': kabel_12,
        'kabel_24': kabel_24,
        'adss_12': adss_12,
        'adss_24': adss_24,
        'odp_8': odp_8,
        'odp_16': odp_16,
        'total_tiang': total_tiang,
        'tiang_new': tiang_new,
        'tiang_existing': tiang_existing,
        'tikungan': tikungan,
        'izin': izin,
        'posisi_odp': posisi_odp,
        'posisi_belokan': posisi_belokan,
        'jumlah_closure': jumlah_closure,
        'uploaded_file': uploaded_file
    }

    # Proses BOQ
    result = process_boq_template(uploaded_file, st.session_state.form_values, lop_name)
    
    if result:
        st.session_state.boq_state = {
            'ready': True,
            'excel_data': result['excel_data'],
            'project_name': lop_name,
            'updated_items': result['updated_items'],
            'summary': result['summary']
        }
        st.success(f"✅ BOQ berhasil digenerate! {result['updated_count']} item diupdate.")

# ======================
# 📊 RESULTS DISPLAY
# ======================
if st.session_state.boq_state.get('ready', False):
    st.divider()
    st.subheader("📥 Download BOQ File")
    st.download_button(
        label="⬇️ Download BOQ",
        data=st.session_state.boq_state['excel_data'],
        file_name=f"BOQ_{st.session_state.boq_state['project_name']}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True
    )

    # Project Summary
    st.subheader("📊 Project Summary")
    summary = st.session_state.boq_state['summary']
    
    col1, col2, col3 = st.columns(3)
    with col1:
        st.metric("Total ODP", summary['total_odp'])
        st.metric("Total Port", summary['total_ports'])
    with col2:
        st.metric("Material", f"Rp {summary['material']:,.0f}")
        st.metric("Jasa", f"Rp {summary['jasa']:,.0f}")
    with col3:
        st.metric("Total Biaya", f"Rp {summary['total']:,.0f}")
        st.metric("CPP (Cost Per Port)", f"Rp {summary['cpp']:,.0f}")

    # Updated Items
    st.subheader("📋 Item yang Diupdate")
    df_items = pd.DataFrame(st.session_state.boq_state['updated_items'])
    st.dataframe(df_items, hide_index=True, use_container_width=True)

    # Reset button
    if st.button("🔄 Buat BOQ Baru", on_click=reset_application, use_container_width=True):
        st.rerun()
else:
    st.info("ℹ️ Silakan isi form dan unggah template BOQ untuk memulai.")

# Footer
st.divider()
st.caption("BOQ Generator v2.0 | © 2024 Telkom Indonesia")
