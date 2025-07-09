import streamlit as st
import pandas as pd
import math
import openpyxl
from io import BytesIO
import xml.etree.ElementTree as ET
from geopy.distance import geodesic

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
            'jumlah_closure': 0,
            'uploaded_file': None,
            'kml_file': None
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
        'jumlah_closure': 0,
        'uploaded_file': None,
        'kml_file': None
    }
    st.session_state.boq_state = {
        'ready': False,
        'excel_data': None,
        'project_name': "",
        'updated_items': [],
        'summary': {}
    }

# Initialize the application
initialize_session_state()

# ======================
# 🔧 CORE FUNCTIONS
# ======================
def calculate_line_length(coord_text):
    """Calculate approximate line length in meters from coordinates"""
    try:
        points = []
        for coord in coord_text.split():
            lon, lat, _ = map(float, coord.split(','))
            points.append((lat, lon))
        
        total_distance = 0.0
        for i in range(len(points)-1):
            total_distance += geodesic(points[i], points[i+1]).meters
        
        return total_distance
    except:
        return 0.0

def parse_kml_file(kml_file):
    """Parse KML file to extract pole, ODP, cable, OTB, and closure positions with enhanced rules"""
    try:
        kml_data = kml_file.read().decode('utf-8')
        root = ET.fromstring(kml_data)
        
        results = {
            'tiang_list': [],
            'odp_positions': [],
            'cable_info': [],
            'otb_count': 0,
            'closure_count': 0,
            'total_tiang': 0,
            'tiang_new': 0,
            'tiang_existing': 0,
            'odp_8': 0,
            'odp_16': 0,
            'adss_12': 0.0,
            'adss_24': 0.0,
            'stock_12': 0.0,
            'stock_24': 0.0
        }

        for elem in root.iter():
            if 'Placemark' in elem.tag:
                name = elem.find('name').text if elem.find('name') is not None else ""
                desc = elem.find('description').text if elem.find('description') is not None else ""
                name_lower = name.lower()
                desc_lower = desc.lower()

                # Extract coordinates
                coords = elem.find('.//coordinates')
                coord_text = coords.text.strip() if coords is not None else ""

                # 1. ODP Identification
                if 'odp' in name_lower:
                    if 'new' in name_lower or 'baru' in name_lower:
                        if '8' in name or ('solid-pb-8' in desc_lower):
                            results['odp_8'] += 1
                            results['odp_positions'].append(f"ODP-8-{results['odp_8']}")
                        elif '16' in name or ('solid-pb-16' in desc_lower):
                            results['odp_16'] += 1
                            results['odp_positions'].append(f"ODP-16-{results['odp_16']}")

                # 2. Pole Identification
                elif any(x in name_lower for x in ['tn', 't7', 'tiang new']):
                    results['tiang_new'] += 1
                    results['tiang_list'].append({'type': 'new', 'position': name})
                
                elif any(x in name_lower for x in ['te', 'tiang exist']):
                    results['tiang_existing'] += 1
                    results['tiang_list'].append({'type': 'existing', 'position': name})

                # 3. Cable Identification (LineString)
                elif 'linestring' in str(elem.find('.//LineString')).lower():
                    if any(x in name_lower or x in desc_lower for x in ['dis new', 'distribusi', 'br']):
                        # New cable rules
                        if 'adss-12d' in desc_lower:
                            results['adss_12'] += calculate_line_length(coord_text)
                        elif 'adss-24d' in desc_lower:
                            results['adss_24'] += calculate_line_length(coord_text)
                        elif '12-sc_o_stock' in desc_lower:
                            results['stock_12'] += calculate_line_length(coord_text)
                        elif '24-sc_o_stock' in desc_lower:
                            results['stock_24'] += calculate_line_length(coord_text)
                    
                    elif any(x in name_lower or x in desc_lower for x in ['ds', 'dis existing']):
                        # Existing cable (track but don't calculate)
                        pass

                # 4. OTB Identification
                elif 'otb' in name_lower and ('new' in name_lower or 'baru' in name_lower):
                    results['otb_count'] += 1

                # 5. Closure Identification
                elif any(x in name_lower for x in ['cl', 'closure']) and ('new' in name_lower or 'baru' in name_lower):
                    results['closure_count'] += 1

        results['total_tiang'] = results['tiang_new'] + results['tiang_existing']
        return results

    except Exception as e:
        st.error(f"Error parsing KML file: {str(e)}")
        return None

def hitung_puas_hl(n_tiang, source='ODC'):
    """Perhitungan khusus PU-AS-HL untuk ADSS"""
    return 16 if source == 'ODC' else 15

def hitung_puas_sc():
    """Perhitungan khusus PU-AS-SC untuk ADSS"""
    return 3

def calculate_volumes(inputs):
    """Calculate all required volumes based on input parameters"""
    total_odp = inputs['odp_8'] + inputs['odp_16']
    total_tiang = inputs['tiang_new'] + inputs['tiang_existing']
    
    # Determine cable type
    is_adss = inputs['adss_12'] > 0 or inputs['adss_24'] > 0
    is_stock = inputs['kabel_12'] > 0 or inputs['kabel_24'] > 0
    
    # Volume kabel
    vol_kabel_12 = round(inputs['kabel_12'] * 1.02) if inputs['kabel_12'] > 0 else 0
    vol_kabel_24 = round(inputs['kabel_24'] * 1.02) if inputs['kabel_24'] > 0 else 0
    vol_adss_12 = round(inputs['adss_12'] * 1.02) if inputs['adss_12'] > 0 else 0
    vol_adss_24 = round(inputs['adss_24'] * 1.02) if inputs['adss_24'] > 0 else 0

    # PU-AS atau PU-AS-HL/SC
    if is_adss:
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
        {"designator": "PS-1-4-ODC", "volume": vol_ps_1_4_odc},
        {"designator": "OS-SM-1-ODC", "volume": vol_os_sm_1_odc},
        {"designator": "OS-SM-1-ODP", "volume": vol_os_sm_1_odp},
        {"designator": "PC-UPC-652-2", "volume": vol_pc_upc},
        {"designator": "PC-APC/UPC-652-A1", "volume": vol_pc_apc},
        {"designator": "TC-02-ODC", "volume": 1 if inputs['sumber'] == "ODC" else 0},
        {"designator": "DD-HDPE-40-1", "volume": 6 if inputs['sumber'] == "ODC" else 0},
        {"designator": "BC-TR-0.6", "volume": 3 if inputs['sumber'] == "ODC" else 0},
        {"designator": "Base Tray ODC", "volume": vol_base_tray_odc},
        {"designator": "SC-OF-SM-24", "volume": vol_closure},
        {"designator": "TC-SM-12", "volume": inputs.get('otb_count', 0)},
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
# 🖥️ FORM UI
# ======================
def show_manual_input_form():
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
            value=st.session_state.form_values['jumlah_closure']
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

        if submitted:
            # Process form submission
            posisi_odp = [int(x.strip()) for x in pos_odp_raw.split(',') if x.strip().isdigit()]
            posisi_belokan = [int(x.strip()) for x in pos_belokan_raw.split(',') if x.strip().isdigit()]
            
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

            # Validate and process
            if not uploaded_file:
                st.error("Harap unggah file template BOQ!")
                return
            
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

def show_kml_input_form():
    with st.form("kml_form"):
        st.subheader("📁 Upload File KML")
        kml_file = st.file_uploader(
            "Unggah File KML*",
            type=["kml"],
            help="Format file harus .kml"
        )
        
        st.subheader("⚙️ Konfigurasi Tambahan")
        sumber = st.radio(
            "Sumber*",
            ["ODC", "ODP"],
            index=0,
            horizontal=True
        )
        pos_belokan_raw = st.text_input(
            "Posisi Tikungan (contoh: 7,13)", 
            value="",
            help="Wajib diisi posisi tikungan"
        )
        izin = st.text_input(
            "Preliminary Project (Rp)",
            value="",
            help="Contoh: 500000"
        )
        
        st.subheader("📤 Template File")
        uploaded_file = st.file_uploader(
            "Unggah Template BOQ*",
            type=["xlsx"],
            help="Format file harus .xlsx"
        )

        submitted = st.form_submit_button("🚀 Generate BOQ dari KML", use_container_width=True)

        if submitted:
            if not kml_file or not uploaded_file:
                st.error("Harap unggah file KML dan template BOQ!")
                return
            
            # Parse KML file
            kml_data = parse_kml_file(kml_file)
            if not kml_data:
                st.error("Gagal memproses file KML!")
                return
            
            posisi_belokan = [int(x.strip()) for x in pos_belokan_raw.split(',') if x.strip().isdigit()]
            
            # Prepare inputs
            inputs = {
                'lop_name': "BOQ dari KML",
                'sumber': sumber,
                'kabel_12': kml_data.get('stock_12', 0),
                'kabel_24': kml_data.get('stock_24', 0),
                'adss_12': kml_data.get('adss_12', 0),
                'adss_24': kml_data.get('adss_24', 0),
                'odp_8': kml_data.get('odp_8', 0),
                'odp_16': kml_data.get('odp_16', 0),
                'total_tiang': kml_data.get('total_tiang', 0),
                'tiang_new': kml_data.get('tiang_new', 0),
                'tiang_existing': kml_data.get('tiang_existing', 0),
                'tikungan': len(posisi_belokan),
                'izin': izin,
                'posisi_odp': kml_data.get('odp_positions', []),
                'posisi_belokan': posisi_belokan,
                'jumlah_closure': kml_data.get('closure_count', 0),
                'otb_count': kml_data.get('otb_count', 0),
                'uploaded_file': uploaded_file
            }
            
            # Process BOQ
            result = process_boq_template(uploaded_file, inputs, inputs['lop_name'])
            
            if result:
                st.session_state.boq_state = {
                    'ready': True,
                    'excel_data': result['excel_data'],
                    'project_name': inputs['lop_name'],
                    'updated_items': result['updated_items'],
                    'summary': result['summary']
                }
                st.success(f"✅ BOQ berhasil digenerate dari KML! {result['updated_count']} item diupdate.")

# ======================
# 📊 MAIN APPLICATION
# ======================
def main():
    st.title("📊 BOQ Generator (KML + Manual)")
    
    tab1, tab2 = st.tabs(["📝 Manual Input", "🗺️ Input dari KML"])
    with tab1:
        show_manual_input_form()
    with tab2:
        show_kml_input_form()
    
    # Show results if available
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

if __name__ == "__main__":
    initialize_session_state()
    main()
