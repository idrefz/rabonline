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
        .disabled-input {
            opacity: 0.7;
        }
    </style>
""", unsafe_allow_html=True)

st.title("📊 BOQ Generator (Custom Rules)")

# ======================
# 🔄 STATE MANAGEMENT
# ======================
def initialize_session_state():
    """Initialize all session state variables with proper defaults"""
    defaults = {
        'lop_name': "",
        'sumber': "ODC",
        'kabel_12': 0.0,
        'kabel_24': 0.0,
        'adss_12': 0.0,
        'adss_24': 0.0,
        'odp_8': 0,
        'odp_16': 0,
        'tiang_new': 0,
        'tiang_existing': 0,
        'tikungan': 0,
        'izin': "",
        'posisi_odp': [],
        'posisi_belokan': [],
        'uploaded_file': None,
        'cable_type': "STOCK"  # Tambahkan state untuk jenis kabel
    }
    
    if 'form_values' not in st.session_state:
        st.session_state.form_values = defaults.copy()
    
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
    initialize_session_state()  # Reinitialize with defaults
    st.rerun()

# Initialize the application
initialize_session_state()

# ======================
# 🔧 CORE FUNCTIONS
# ======================
def hitung_puas_hl(n_tiang, source='ODC', posisi_odp=[]):
    """Calculate PU-AS-HL count"""
    puas_hl = 0
    counter = 0
    for i in range(n_tiang):
        idx = i + 1
        if idx == 1:
            puas_hl += 2 if source == 'ODC' else 1
            counter = 1
        elif idx in posisi_odp:
            puas_hl += 1
            counter = 1
        elif counter == 5:
            puas_hl += 2
            counter = 1
        else:
            counter += 1
    return puas_hl

def hitung_puas_sc(posisi_odp, posisi_belokan):
    """Calculate PU-AS-SC count"""
    return sum([2 if i in posisi_belokan else 3 for i in posisi_odp])

def calculate_volumes(inputs):
    """Calculate all required volumes based on input parameters"""
    total_odp = inputs['odp_8'] + inputs['odp_16']
    total_tiang = inputs['tiang_new'] + inputs['tiang_existing']
    
    # Determine cable type
    is_stock = inputs['cable_type'] == "STOCK"
    is_adss = inputs['cable_type'] == "ADSS"
    
    # Volume kabel
    vol_kabel_12 = round(inputs['kabel_12'] * 1.02) if is_stock and inputs['kabel_12'] > 0 else 0
    vol_kabel_24 = round(inputs['kabel_24'] * 1.02) if is_stock and inputs['kabel_24'] > 0 else 0
    vol_adss_12 = round(inputs['adss_12'] * 1.02) if is_adss and inputs['adss_12'] > 0 else 0
    vol_adss_24 = round(inputs['adss_24'] * 1.02) if is_adss and inputs['adss_24'] > 0 else 0

    # PU-AS atau PU-AS-HL/SC
    if is_adss:
        vol_puas_hl = hitung_puas_hl(total_tiang, inputs['sumber'], inputs['posisi_odp'])
        vol_puas_sc = hitung_puas_sc(inputs['posisi_odp'], inputs['posisi_belokan'])
        vol_puas = 0
    else:
        vol_puas = max(0, (total_odp * 2) - 1 + total_tiang + inputs['tikungan'])
        vol_puas_hl = 0
        vol_puas_sc = 0

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
    vol_pc_upc = ((total_odp - 1) // 4) + 1 if total_odp > 0 else 0
    vol_pc_apc = 18 if vol_pc_upc == 1 else vol_pc_upc * 2 if vol_pc_upc > 1 else 0
    vol_ps_1_4_odc = ((total_odp - 1) // 4) + 1 if inputs['sumber'] == "ODC" and total_odp > 0 else 0

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
        {"designator": "PU-S7.0-400NM", "volume": inputs['tiang_new']},
        {"designator": "PU-AS", "volume": vol_puas},
        {"designator": "PU-AS-HL", "volume": vol_puas_hl},
        {"designator": "PU-AS-SC", "volume": vol_puas_sc},
        {"designator": "OS-SM-1-ODC", "volume": vol_os_sm_1_odc},
        {"designator": "OS-SM-1-ODP", "volume": vol_os_sm_1_odp},
        {"designator": "OS-SM-1", "volume": vol_os_sm_1},
        {"designator": "PC-UPC-652-2", "volume": vol_pc_upc},
        {"designator": "PC-APC/UPC-652-A1", "volume": vol_pc_apc},
        {"designator": "PS-1-4-ODC", "volume": vol_ps_1_4_odc},
        {"designator": "TC-02-ODC", "volume": vol_tc_02_odc},
        {"designator": "DD-HDPE-40-1", "volume": vol_dd_hdpe},
        {"designator": "BC-TR-0.6", "volume": vol_bc_tr},
        {"designator": "Base Tray ODC", "volume": vol_base_tray_odc},
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
# 🗅️ FORM UI
# ======================
with st.form("boq_form"):
    # Common sections
    st.subheader("📁 Informasi Proyek")
    col1, col2 = st.columns([2, 1])
    with col1:
        lop_name = st.text_input(
            "Nama LOP*",
            value=st.session_state.form_values.get('lop_name', ''),
            key='lop_name_input',
            help="Masukkan nama LOP (contoh: LOP_JAKARTA_123)"
        )
    with col2:
        sumber = st.radio(
            "Sumber*",
            ["ODC", "ODP"],
            index=0 if st.session_state.form_values.get('sumber', 'ODC') == "ODC" else 1,
            key='sumber_input',
            horizontal=True
        )
    
    # Jenis Kabel Selection
    cable_type = st.radio(
        "Jenis Kabel*",
        ["STOCK", "ADSS"],
        index=0 if st.session_state.form_values.get('cable_type', 'STOCK') == "STOCK" else 1,
        key='cable_type_input',
        horizontal=True
    )

    # Create tabs
    tab1, tab2 = st.tabs(["📦 STOCK", "📡 ADSS"])
    
    # Initialize common fields
    common_fields = {
        'odp_8': st.session_state.form_values.get('odp_8', 0),
        'odp_16': st.session_state.form_values.get('odp_16', 0),
        'tiang_new': st.session_state.form_values.get('tiang_new', 0),
        'tiang_existing': st.session_state.form_values.get('tiang_existing', 0),
        'tikungan': st.session_state.form_values.get('tikungan', 0),
        'izin': st.session_state.form_values.get('izin', ''),
        'cable_type': cable_type
    }
    
    with tab1:
        # STOCK MATERIALS
        st.subheader("Kebutuhan Material Stock")
        col1, col2 = st.columns(2)
        with col1:
            kabel_12 = st.number_input(
                "12 Core Cable (meter)*",
                min_value=0.0,
                value=st.session_state.form_values.get('kabel_12', 0.0),
                key='kabel_12_input',
                step=1.0,
                format="%.1f",
                disabled=cable_type == "ADSS"
            )
            odp_8 = st.number_input(
                "ODP 8 Port*",
                min_value=0,
                value=st.session_state.form_values.get('odp_8', 0),
                key='odp_8_input'
            )
            tiang_new = st.number_input(
                "Tiang Baru*",
                min_value=0,
                value=st.session_state.form_values.get('tiang_new', 0),
                key='tiang_new_input'
            )
        with col2:
            kabel_24 = st.number_input(
                "24 Core Cable (meter)*",
                min_value=0.0,
                value=st.session_state.form_values.get('kabel_24', 0.0),
                key='kabel_24_input',
                step=1.0,
                format="%.1f",
                disabled=cable_type == "ADSS"
            )
            odp_16 = st.number_input(
                "ODP 16 Port*",
                min_value=0,
                value=st.session_state.form_values.get('odp_16', 0),
                key='odp_16_input'
            )
            tiang_existing = st.number_input(
                "Tiang Eksisting*",
                min_value=0,
                value=st.session_state.form_values.get('tiang_existing', 0),
                key='tiang_existing_input'
            )
        
        tikungan = st.number_input(
            "Tikungan*",
            min_value=0,
            value=st.session_state.form_values.get('tikungan', 0),
            key='tikungan_input'
        )
        izin = st.text_input(
            "Preliminary (isi nominal jika ada)",
            value=st.session_state.form_values.get('izin', ''),
            key='izin_input',
            help="Masukkan nilai dalam rupiah (contoh: 500000)"
        )

        # Update common fields
        common_fields.update({
            'odp_8': odp_8,
            'odp_16': odp_16,
            'tiang_new': tiang_new,
            'tiang_existing': tiang_existing,
            'tikungan': tikungan,
            'izin': izin
        })

    with tab2:
        # ADSS MATERIALS
        st.subheader("Kebutuhan Material ADSS")
        col1, col2 = st.columns(2)
        with col1:
            adss_12 = st.number_input(
                "ADSS 12 Core (meter)*",
                min_value=0.0,
                value=st.session_state.form_values.get('adss_12', 0.0),
                key='adss_12_input',
                step=1.0,
                format="%.1f",
                disabled=cable_type == "STOCK"
            )
            odp_8 = st.number_input(
                "ODP 8 Port*",
                min_value=0,
                value=st.session_state.form_values.get('odp_8', 0),
                key='odp_8_adss_input'
            )
            tiang_new = st.number_input(
                "Tiang Baru*",
                min_value=0,
                value=st.session_state.form_values.get('tiang_new', 0),
                key='tiang_new_adss_input'
            )
            pos_odp_raw = st.text_input(
                "Posisi Tiang ODP (misal: 5,9,14)*", 
                value=",".join(map(str, st.session_state.form_values.get('posisi_odp', []))),
                key='pos_odp_input',
                disabled=cable_type == "STOCK"
            )
        with col2:
            adss_24 = st.number_input(
                "ADSS 24 Core (meter)*",
                min_value=0.0,
                value=st.session_state.form_values.get('adss_24', 0.0),
                key='adss_24_input',
                step=1.0,
                format="%.1f",
                disabled=cable_type == "STOCK"
            )
            odp_16 = st.number_input(
                "ODP 16 Port*",
                min_value=0,
                value=st.session_state.form_values.get('odp_16', 0),
                key='odp_16_adss_input'
            )
            tiang_existing = st.number_input(
                "Tiang Eksisting*",
                min_value=0,
                value=st.session_state.form_values.get('tiang_existing', 0),
                key='tiang_existing_adss_input'
            )
            pos_belokan_raw = st.text_input(
                "Posisi Tikungan (misal: 7,13)*", 
                value=",".join(map(str, st.session_state.form_values.get('posisi_belokan', []))),
                key='pos_belokan_input',
                disabled=cable_type == "STOCK"
            )
        
        tikungan = st.number_input(
            "Tikungan*",
            min_value=0,
            value=st.session_state.form_values.get('tikungan', 0),
            key='tikungan_adss_input'
        )
        izin = st.text_input(
            "Preliminary (isi nominal jika ada)",
            value=st.session_state.form_values.get('izin', ''),
            key='izin_adss_input',
            help="Masukkan nilai dalam rupiah (contoh: 500000)"
        )

        # Update common fields
        common_fields.update({
            'odp_8': odp_8,
            'odp_16': odp_16,
            'tiang_new': tiang_new,
            'tiang_existing': tiang_existing,
            'tikungan': tikungan,
            'izin': izin
        })

    # FILE UPLOAD
    st.subheader("📤 Template File")
    uploaded_file = st.file_uploader(
        "Unggah Template BOQ*",
        type=["xlsx"],
        key='uploaded_file_input',
        help="File template Excel format BOQ"
    )

    # SUBMIT BUTTON
    submitted = st.form_submit_button("🚀 Generate BOQ", use_container_width=True)

# ======================
# 🚀 FORM SUBMISSION
# ======================
if submitted:
    # Validate required fields
    if not uploaded_file:
        st.error("Silakan unggah file template BOQ!")
        st.stop()
    if not lop_name:
        st.error("Silakan isi nama LOP!")
        st.stop()
    
    # Validate izin is numeric if provided
    if common_fields['izin'] and not common_fields['izin'].replace(',', '').replace('.', '').isdigit():
        st.error("Nilai preliminary harus berupa angka!")
        st.stop()
    
    # Validate cable inputs based on selected type
    if cable_type == "STOCK":
        if kabel_12 <= 0 and kabel_24 <= 0:
            st.error("Untuk kabel STOCK, minimal salah satu core (12 atau 24) harus diisi!")
            st.stop()
    else:  # ADSS
        if adss_12 <= 0 and adss_24 <= 0:
            st.error("Untuk kabel ADSS, minimal salah satu core (12 atau 24) harus diisi!")
            st.stop()
        if not pos_odp_raw:
            st.error("Untuk kabel ADSS, posisi tiang ODP harus diisi!")
            st.stop()

    # Process position inputs for ADSS
    posisi_odp = []
    posisi_belokan = []
    if cable_type == "ADSS":
        try:
            posisi_odp = [int(x.strip()) for x in pos_odp_raw.split(',') if x.strip().isdigit()]
            if pos_belokan_raw:
                posisi_belokan = [int(x.strip()) for x in pos_belokan_raw.split(',') if x.strip().isdigit()]
            
            # Validate positions are positive and within range
            total_tiang = common_fields['tiang_new'] + common_fields['tiang_existing']
            if any(p <= 0 for p in posisi_odp + posisi_belokan):
                st.error("Posisi harus berupa angka positif")
                st.stop()
            if any(p > total_tiang for p in posisi_odp + posisi_belokan):
                st.error(f"Posisi tidak boleh melebihi jumlah tiang ({total_tiang})")
                st.stop()
        except Exception as e:
            st.error(f"Format posisi tidak valid: {str(e)}")
            st.stop()

    # Update session state with current form values
    st.session_state.form_values = {
        'lop_name': lop_name,
        'sumber': sumber,
        'kabel_12': kabel_12,
        'kabel_24': kabel_24,
        'adss_12': adss_12,
        'adss_24': adss_24,
        'odp_8': common_fields['odp_8'],
        'odp_16': common_fields['odp_16'],
        'tiang_new': common_fields['tiang_new'],
        'tiang_existing': common_fields['tiang_existing'],
        'tikungan': common_fields['tikungan'],
        'izin': common_fields['izin'],
        'posisi_odp': posisi_odp,
        'posisi_belokan': posisi_belokan,
        'uploaded_file': uploaded_file,
        'cable_type': cable_type
    }

    # Process the BOQ template
    with st.spinner("Sedang memproses BOQ..."):
        input_data = st.session_state.form_values.copy()
        result = process_boq_template(uploaded_file, input_data, lop_name)
    
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
# 📂 RESULTS SECTION
# ======================
if st.session_state.boq_state.get('ready', False):
    st.divider()
    st.subheader("📥 Download BOQ File")
    
    # Download button
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
        pass  # reset_application() will be called via on_click
else:
    st.info("ℹ️ Silakan isi form dan unggah template BOQ untuk memulai.")

# Footer
st.divider()
st.caption("BOQ Generator v1.0 | © 2024 Telkom Indonesia")
