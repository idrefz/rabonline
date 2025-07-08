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
            # Standard BOQ
            'lop_name': "", 'sumber': "ODC",
            'kabel_12': 0.0, 'kabel_24': 0.0,
            'odp_8': 0, 'odp_16': 0,
            'tiang_new': 0, 'tiang_existing': 0,
            'tikungan': 0, 'izin': "",
            'uploaded_file': None,
            # ADSS BOQ
            'adss_proj': "",
            'kabel_12d': 0.0, 'kabel_24d': 0.0,
            'adss_tiang_new': 0, 'adss_tiang_existing': 0,
            'adss_tikungan': 0, 'adss_dead_end': 0,
            'adss_uploaded_file': None
        }
    if 'boq_state' not in st.session_state:
        st.session_state.boq_state = {
            'ready': False, 'excel_data': None,
            'project_name': "", 'updated_items': [], 'summary': {}
        }
    if 'adss_state' not in st.session_state:
        st.session_state.adss_state = {
            'ready': False, 'excel_data': None,
            'project_name': "", 'updated_items': [], 'summary': {}
        }

def reset_application():
    """Reset the entire application state"""
    initialize_session_state()
    st.session_state.boq_state.update({
        'ready': False, 'excel_data': None,
        'project_name': "", 'updated_items': [], 'summary': {}
    })
    st.session_state.adss_state.update({
        'ready': False, 'excel_data': None,
        'project_name': "", 'updated_items': [], 'summary': {}
    })

initialize_session_state()

# ======================
# 🔧 CORE FUNCTIONS
# ======================
def calculate_volumes(inputs):
    """Calculate all required volumes based on input parameters"""
    total_odp = inputs['odp_8'] + inputs['odp_16']
    if inputs['kabel_12'] > 0 and inputs['kabel_24'] > 0:
        raise ValueError("Silakan pilih hanya satu jenis kabel (12-core ATAU 24-core)")
    vol_kabel_12 = round(inputs['kabel_12'] * 1.02) if inputs['kabel_12'] > 0 else 0
    vol_kabel_24 = round(inputs['kabel_24'] * 1.02) if inputs['kabel_24'] > 0 else 0
    vol_puas = max(0, (total_odp * 2) - 1 + inputs['tiang_new'] + inputs['tiang_existing'] + inputs['tikungan'])
    vol_os_sm_1_odc = total_odp * 2 if inputs['sumber']=="ODC" else 0
    vol_os_sm_1_odp = total_odp * 2 if inputs['sumber']=="ODP" else 0
    vol_base_tray_odc = 1 if inputs['sumber']=="ODC" and inputs['kabel_12']>0 else \
                        2 if inputs['sumber']=="ODC" and inputs['kabel_24']>0 else 0
    vol_pc_upc = ((total_odp-1)//4)+1 if total_odp>0 else 0
    vol_pc_apc = 18 if vol_pc_upc==1 else vol_pc_upc*2 if vol_pc_upc>1 else 0
    vol_ps_1_4_odc = ((total_odp-1)//4)+1 if inputs['sumber']=="ODC" and total_odp>0 else 0
    vol_tc_02_odc = 1 if inputs['sumber']=="ODC" else 0
    vol_dd_hdpe   = 6 if inputs['sumber']=="ODC" else 0
    vol_bc_tr     = 3 if inputs['sumber']=="ODC" else 0

    return [
        {"designator":"AC-OF-SM-12-SC_O_STOCK","volume":vol_kabel_12},
        {"designator":"AC-OF-SM-24-SC_O_STOCK","volume":vol_kabel_24},
        {"designator":"ODP Solid-PB-8 AS","volume":inputs['odp_8']},
        {"designator":"ODP Solid-PB-16 AS","volume":inputs['odp_16']},
        {"designator":"PU-S7.0-400NM","volume":inputs['tiang_new']},
        {"designator":"PU-AS","volume":vol_puas},
        {"designator":"OS-SM-1-ODC","volume":vol_os_sm_1_odc},
        {"designator":"OS-SM-1-ODP","volume":vol_os_sm_1_odp},
        {"designator":"OS-SM-1","volume":vol_os_sm_1_odc+vol_os_sm_1_odp},
        {"designator":"PC-UPC-652-2","volume":vol_pc_upc},
        {"designator":"PC-APC/UPC-652-A1","volume":vol_pc_apc},
        {"designator":"PS-1-4-ODC","volume":vol_ps_1_4_odc},
        {"designator":"TC-02-ODC","volume":vol_tc_02_odc},
        {"designator":"DD-HDPE-40-1","volume":vol_dd_hdpe},
        {"designator":"BC-TR-0.6","volume":vol_bc_tr},
        {"designator":"Base Tray ODC","volume":vol_base_tray_odc},
        {"designator":"Preliminary Project HRB/Kawasan Khusus",
         "volume":1 if inputs['izin'] else 0,
         "izin_value":float(inputs['izin']) if inputs['izin'] else 0}
    ]

def process_boq_template(uploaded_file, inputs, proj_name):
    """Process the BOQ template file and calculate all metrics"""
    wb = openpyxl.load_workbook(uploaded_file)
    ws = wb.active
    items = calculate_volumes(inputs)
    updated_count = 0

    # Fill volumes & preliminary row
    for r in range(9,289):
        cell = str(ws[f'B{r}'].value or "").strip()
        if inputs['izin'] and cell=="" and \
           "Preliminary Project HRB/Kawasan Khusus" not in [str(ws[f'B{i}'].value) for i in range(9,289)]:
            ws[f'B{r}']="Preliminary Project HRB/Kawasan Khusus"
            ws[f'F{r}']=inputs['izin_value']=float(inputs['izin'])
            ws[f'G{r}']=1
            updated_count+=1
            continue
        for itm in items:
            if itm["volume"]>0 and cell==itm["designator"]:
                ws[f'G{r}']=itm["volume"]
                if "Preliminary" in cell:
                    ws[f'F{r}']=itm.get("izin_value",0)
                updated_count+=1
                break

    # Sum costs
    material=jasa=0.0
    for r in range(9,289):
        try:
            m=float(ws[f'E{r}'].value or 0)
            j=float(ws[f'F{r}'].value or 0)
            v=float(ws[f'G{r}'].value or 0)
            material+=m*v
            jasa+=j*v
        except: pass

    total_odp=inputs['odp_8']+inputs['odp_16']
    total=material+jasa
    cpp=round(total/(total_odp*8),2) if total_odp>0 else 0

    output=BytesIO(); wb.save(output); output.seek(0)
    return {
        'excel_data': output,
        'updated_count':updated_count,
        'updated_items':[i for i in items if i['volume']>0],
        'summary':{
            'material':material,'jasa':jasa,'total':total,
            'cpp':cpp,'total_odp':total_odp,'total_ports':total_odp*8
        }
    }

def calculate_adss_volumes(inputs):
    """Calculate ADSS-specific volumes"""
    poles=inputs['adss_tiang_new']+inputs['adss_tiang_existing']
    sc=max(0,poles-2)
    hl_p=math.ceil(poles/4)
    hl_l12=math.ceil(inputs['kabel_12d']/200)
    hl_l24=math.ceil(inputs['kabel_24d']/200)
    hl=max(hl_p,hl_l12,hl_l24)+inputs['adss_tikungan']
    return [
        {"designator":"AC-OF-SM-ADSS-12D","volume":round(inputs['kabel_12d']*1.02)},
        {"designator":"AC-OF-SM-ADSS-24D","volume":round(inputs['kabel_24d']*1.02)},
        {"designator":"PU-AS-SC","volume":sc},
        {"designator":"PU-AS-HL","volume":hl},
        {"designator":"PU-S7.0-400NM","volume":inputs['adss_tiang_new']},
    ]

def process_adss_template(uploaded_file, inputs, proj_name):
    """Fill ADSS template"""
    wb=openpyxl.load_workbook(uploaded_file)
    ws=wb.active
    items=calculate_adss_volumes(inputs)
    for r in range(9,289):
        cell=str(ws[f'B{r}'].value or "").strip()
        for itm in items:
            if itm["volume"]>0 and cell==itm["designator"]:
                ws[f'G{r}']=itm["volume"]
    material=jasa=0.0
    for r in range(9,289):
        try:
            m=float(ws[f'E{r}'].value or 0)
            j=float(ws[f'F{r}'].value or 0)
            v=float(ws[f'G{r}'].value or 0)
            material+=m*v
            jasa+=j*v
        except: pass
    total=material+jasa
    output=BytesIO(); wb.save(output); output.seek(0)
    summary={
        'material':material,'jasa':jasa,'total':total,
        'total_poles':inputs['adss_tiang_new']+inputs['adss_tiang_existing'],
        'cable_length':inputs['kabel_12d']+inputs['kabel_24d']
    }
    return {'excel_data':output,'updated_items':items,'summary':summary}


# ======================
# 🗅️ FORM UI & SUBMISSION
# ======================
# Standard BOQ Form
with st.form("boq_form"):
    st.subheader("Project Information")
    c1,c2=st.columns([2,1])
    with c1:
        lop_name=st.text_input("Nama LOP*",key='lop_name_input')
    with c2:
        sumber=st.radio("Sumber*",["ODC","ODP"],key='sumber_input',horizontal=True)

    st.subheader("Material Requirements")
    a1,a2=st.columns(2)
    with a1:
        kabel_12=st.number_input("12 Core Cable (m)*",key='kabel_12_input',format="%.1f")
        odp_8=    st.number_input("ODP 8 Port*",key='odp_8_input')
        tiang_new=st.number_input("Tiang Baru*",key='tiang_new_input')
        tikungan= st.number_input("Tikungan*",key='tikungan_input')
    with a2:
        kabel_24=   st.number_input("24 Core Cable (m)*",key='kabel_24_input',format="%.1f")
        odp_16=     st.number_input("ODP 16 Port*",key='odp_16_input')
        tiang_existing=st.number_input("Tiang Eksisting*",key='tiang_existing_input')
        izin=       st.text_input("Preliminary (nominal)",key='izin_input')

    uploaded_file=st.file_uploader("Unggah Template BOQ* (xlsx)",type=["xlsx"],key='uploaded_file_input')
    submitted=st.form_submit_button("🚀 Generate BOQ",use_container_width=True)

if submitted:
    if not uploaded_file or not lop_name:
        st.error("Nama LOP dan template wajib diisi.")
    else:
        st.session_state.form_values.update({
            'lop_name':lop_name,'sumber':sumber,
            'kabel_12':kabel_12,'kabel_24':kabel_24,
            'odp_8':odp_8,'odp_16':odp_16,
            'tiang_new':tiang_new,'tiang_existing':tiang_existing,
            'tikungan':tikungan,'izin':izin,
            'uploaded_file':uploaded_file
        })
        res=process_boq_template(uploaded_file,st.session_state.form_values,lop_name)
        if res:
            st.session_state.boq_state.update({
                'ready':True,
                'excel_data':res['excel_data'],
                'project_name':lop_name,
                'updated_items':res['updated_items'],
                'summary':res['summary']
            })
            st.success(f"✅ BOQ berhasil digenerate! {res['updated_count']} item diupdate.")

# Standard BOQ Results
if st.session_state.boq_state['ready']:
    st.divider()
    st.subheader("📥 Download BOQ")
    st.download_button(
        "⬇️ Download BOQ",
        data=st.session_state.boq_state['excel_data'],
        file_name=f"BOQ_{st.session_state.boq_state['project_name']}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True
    )
    # you can also show summary & updated_items here…

# ADSS BOQ Form
st.markdown("---")
with st.form("adss_form"):
    st.subheader("📡 ADSS BOQ – Project Information")
    proj=st.text_input("Nama Proyek*",key='adss_proj_input')

    st.subheader("ADSS Material Requirements")
    b1,b2=st.columns(2)
    with b1:
        kabel_12d=st.number_input("ADSS 12D Cable (m)*",key='kabel_12d_input',format="%.1f")
        adss_tiang_new=st.number_input("Tiang Baru*",key='adss_tiang_new_input')
        adss_tikungan=st.number_input("Tikungan*",key='adss_tikungan_input')
    with b2:
        kabel_24d=st.number_input("ADSS 24D Cable (m)*",key='kabel_24d_input',format="%.1f")
        adss_tiang_existing=st.number_input("Tiang Existing*",key='adss_tiang_existing_input')
        adss_dead_end=st.number_input("Dead End Clamp*",key='adss_dead_end_input',
                                      help="1 per 4 tiang atau per 200m + tikungan")

    adss_uploaded_file=st.file_uploader("Unggah Template ADSS BOQ* (xlsx)",type=["xlsx"],key='adss_uploaded_file')
    submitted_adss=st.form_submit_button("🚀 Generate ADSS BOQ",use_container_width=True)

if submitted_adss:
    if not adss_uploaded_file or not proj:
        st.error("Nama proyek & template ADSS wajib diisi.")
    else:
        st.session_state.form_values.update({
            'adss_proj':proj,
            'kabel_12d':kabel_12d,'kabel_24d':kabel_24d,
            'adss_tiang_new':adss_tiang_new,'adss_tiang_existing':adss_tiang_existing,
            'adss_tikungan':adss_tikungan,'adss_dead_end':adss_dead_end,
            'adss_uploaded_file':adss_uploaded_file
        })
        res_adss=process_adss_template(adss_uploaded_file,st.session_state.form_values,proj)
        if res_adss:
            st.session_state.adss_state.update({
                'ready':True,
                'excel_data':res_adss['excel_data'],
                'project_name':proj,
                'updated_items':res_adss['updated_items'],
                'summary':res_adss['summary']
            })
            st.success("✅ ADSS BOQ berhasil digenerate!")

# ADSS BOQ Results
if st.session_state.adss_state['ready']:
    st.divider()
    st.subheader("📥 Download ADSS BOQ")
    st.download_button(
        "⬇️ Download ADSS BOQ",
        data=st.session_state.adss_state['excel_data'],
        file_name=f"ADSS_BOQ_{st.session_state.adss_state['project_name']}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True
    )
    # you can also show summary & updated_items here…

# Footer
st.divider()
st.caption("BOQ Generator v1.0 | © 2024 Telkom Indonesia")
