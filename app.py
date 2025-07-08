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
        .block-container { padding: 2rem 3rem; }
        .stRadio > div { flex-direction: row; }
        .metric { text-align: center; }
        .metric .st-emotion-cache-1xarl3l { font-size: 1.2rem !important; }
    </style>
""", unsafe_allow_html=True)

st.title("📊 BOQ Generator (Custom Rules)")

# ======================
# 🔄 STATE MANAGEMENT
# ======================
def initialize_session_state():
    if 'form_values' not in st.session_state:
        st.session_state.form_values = {
            # Standard BOQ
            'lop_name': "", 'sumber': "ODC",
            'kabel_12': 0.0, 'kabel_24': 0.0,
            'odp_8': 0, 'odp_16': 0,
            'tiang_new': 0, 'tiang_existing': 0,
            'tikungan': 0, 'izin': "",
            'uploaded_file': None,
            # ADSS BOQ
            'adss_proj': "", 'kabel_12d': 0.0, 'kabel_24d': 0.0,
            'adss_tiang_new': 0, 'adss_tiang_existing': 0,
            'adss_tikungan': 0, 'adss_dead_end': 0,
            'adss_uploaded_file': None
        }
    if 'boq_state' not in st.session_state:
        st.session_state.boq_state = {'ready': False, 'excel_data': None,
                                      'project_name': "", 'updated_items': [],
                                      'summary': {}}
    if 'adss_state' not in st.session_state:
        st.session_state.adss_state = {'ready': False, 'excel_data': None,
                                       'project_name': "", 'updated_items': [],
                                       'summary': {}}

def reset_application():
    initialize_session_state()
    for k in list(st.session_state.boq_state): st.session_state.boq_state[k] = False if k=='ready' else None
    for k in list(st.session_state.adss_state): st.session_state.adss_state[k] = False if k=='ready' else None

initialize_session_state()


# ======================
# 🔧 CORE FUNCTIONS
# ======================
def calculate_adss_volumes(inputs):
    total_poles = inputs['adss_tiang_new'] + inputs['adss_tiang_existing']
    # SC = total_poles - 2 (minimum 0)
    sc = max(0, total_poles - 2)
    # HL = max(ceil(poles/4), ceil(length/200)) + tikungan
    hl_from_poles = math.ceil(total_poles/4)
    hl_from_len12 = math.ceil(inputs['kabel_12d']/200)
    hl_from_len24 = math.ceil(inputs['kabel_24d']/200)
    hl = max(hl_from_poles, hl_from_len12, hl_from_len24) + inputs['adss_tikungan']
    return [
        {"designator": "AC-OF-SM-ADSS-12D", "volume": round(inputs['kabel_12d']*1.02)},
        {"designator": "AC-OF-SM-ADSS-24D", "volume": round(inputs['kabel_24d']*1.02)},
        {"designator": "PU-AS-SC",          "volume": sc},
        {"designator": "PU-AS-HL",          "volume": hl},
        {"designator": "PU-S7.0-400NM",     "volume": inputs['adss_tiang_new']},
    ]

def process_adss_template(uploaded_file, inputs, proj_name):
    wb = openpyxl.load_workbook(uploaded_file)
    ws = wb.active
    items = calculate_adss_volumes(inputs)
    for r in range(9, 289):
        designator = str(ws[f'B{r}'].value or "").strip()
        for itm in items:
            if itm["volume"] > 0 and designator == itm["designator"]:
                ws[f'G{r}'] = itm["volume"]
    # hitung material & jasa
    material = jasa = 0.0
    for r in range(9, 289):
        try:
            h_mat = float(ws[f'E{r}'].value or 0)
            h_jasa= float(ws[f'F{r}'].value or 0)
            vol  = float(ws[f'G{r}'].value or 0)
            material += h_mat*vol
            jasa     += h_jasa*vol
        except: pass

    total = material+jasa
    output = BytesIO()
    wb.save(output); output.seek(0)
    summary = {
        'material': material, 'jasa': jasa,
        'total': total,
        'total_poles': inputs['adss_tiang_new']+inputs['adss_tiang_existing'],
        'cable_length': inputs['kabel_12d']+inputs['kabel_24d']
    }
    return {'excel_data': output, 'updated_items': items, 'summary': summary}


# ======================
# 🗅️ FORM UI & SUBMISSION
# ======================
tab1, tab2 = st.tabs(["📝 Standard BOQ","📡 ADSS BOQ"])

with tab1:
    with st.form("boq_form"):
        st.subheader("Standard BOQ – Project Information")
        col1, col2 = st.columns([2,1])
        with col1:
            lop = st.text_input("Nama LOP*", key='lop_name_input')
        with col2:
            sumber = st.radio("Sumber*", ["ODC","ODP"], key='sumber_input', horizontal=True)

        st.subheader("Material Requirements")
        c1,c2 = st.columns(2)
        with c1:
            k12  = st.number_input("12-Core Cable (m)*", key='kabel_12_input', format="%.1f")
            odp8 = st.number_input("ODP 8 Port*", key='odp_8_input')
            tn   = st.number_input("Tiang Baru*", key='tiang_new_input')
            tik  = st.number_input("Tikungan*", key='tikungan_input')
        with c2:
            k24  = st.number_input("24-Core Cable (m)*", key='kabel_24_input', format="%.1f")
            odp16= st.number_input("ODP 16 Port*", key='odp_16_input')
            tex = st.number_input("Tiang Existing*", key='tiang_existing_input')
            izin= st.text_input("Preliminary (nominal)", key='izin_input')

        file_std = st.file_uploader("Unggah Template BOQ* (xlsx)", key='uploaded_file_input')
        submit_std = st.form_submit_button("🚀 Generate Standard BOQ")

    if submit_std:
        # validasi
        if not file_std or not lop:
            st.error("Nama LOP & template wajib diisi.")
        else:
            st.session_state.form_values.update({
                'lop_name': lop, 'sumber': sumber, 'kabel_12': k12, 'kabel_24': k24,
                'odp_8': odp8, 'odp_16': odp16, 'tiang_new': tn, 'tiang_existing': tex,
                'tikungan': tik, 'izin': izin,'uploaded_file': file_std
            })
            res = process_boq_template(file_std, st.session_state.form_values, lop)
            if res:
                st.session_state.boq_state = {
                    'ready': True,
                    'excel_data': res['excel_data'],
                    'project_name': lop,
                    'updated_items': res['updated_items'],
                    'summary': res['summary']
                }
                st.success("✅ Standard BOQ berhasil digenerate!")

with tab2:
    with st.form("adss_form"):
        st.subheader("ADSS BOQ – Project Information")
        proj = st.text_input("Nama Proyek*", key='adss_proj_input')

        st.subheader("ADSS Material Requirements")
        c1,c2 = st.columns(2)
        with c1:
            k12d = st.number_input("ADSS 12D Cable (m)*", key='kabel_12d_input', format="%.1f")
            tnd  = st.number_input("Tiang Baru*", key='adss_tiang_new_input')
            tikd = st.number_input("Tikungan*", key='adss_tikungan_input')
        with c2:
            k24d = st.number_input("ADSS 24D Cable (m)*", key='kabel_24d_input', format="%.1f")
            ted  = st.number_input("Tiang Existing*", key='adss_tiang_existing_input')
            dead = st.number_input("Dead End Clamp*", key='adss_dead_end_input',
                                   help="1 per 4 tiang atau per 200m + tikungan")

        file_adss = st.file_uploader("Unggah Template ADSS BOQ* (xlsx)", key='adss_uploaded_file')
        submit_adss = st.form_submit_button("🚀 Generate ADSS BOQ")

    if submit_adss:
        if not file_adss or not proj:
            st.error("Nama Proyek & template ADSS wajib diisi.")
        else:
            st.session_state.form_values.update({
                'adss_proj': proj,
                'kabel_12d': k12d, 'kabel_24d': k24d,
                'adss_tiang_new': tnd, 'adss_tiang_existing': ted,
                'adss_tikungan': tikd, 'adss_dead_end': dead
            })
            res = process_adss_template(file_adss, st.session_state.form_values, proj)
            if res:
                st.session_state.adss_state = {
                    'ready': True,
                    'excel_data': res['excel_data'],
                    'project_name': proj,
                    'updated_items': res['updated_items'],
                    'summary': res['summary']
                }
                st.success("✅ ADSS BOQ berhasil digenerate!")

# ======================
# 📂 RESULTS SECTION
# ======================
# Standard BOQ result
if st.session_state.boq_state['ready']:
    st.divider()
    st.subheader("📥 Download Standard BOQ")
    st.download_button("⬇️ Download Standard BOQ",
                       data=st.session_state.boq_state['excel_data'],
                       file_name=f"BOQ_{st.session_state.boq_state['project_name']}.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    # … tampilkan summary & tabel updated_items …

# ADSS BOQ result
if st.session_state.adss_state['ready']:
    st.divider()
    st.subheader("📥 Download ADSS BOQ")
    st.download_button("⬇️ Download ADSS BOQ",
                       data=st.session_state.adss_state['excel_data'],
                       file_name=f"ADSS_BOQ_{st.session_state.adss_state['project_name']}.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    # … tampilkan summary & tabel updated_items …
