import streamlit as st
import pandas as pd
from io import BytesIO
import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill

# ======================
# 🎩 CONFIGURATION
# ======================
st.set_page_config("BOQ Generator", layout="centered")
st.title("📊 BOQ Generator (Custom Rules)")

# Initialize session state
if 'boq_state' not in st.session_state:
    st.session_state.boq_state = {
        'ready': False,
        'excel_data': None,
        'project_name': "",
        'updated_items': [],
        'summary': {}
    }

# ======================
# 🔧 CORE FUNCTIONS
# ======================
def calculate_volumes(inputs):
    total_odp = inputs['odp_8'] + inputs['odp_16']

    vol_kabel_12 = round(inputs['kabel_12'] * 1.02) if inputs['kabel_12'] > 0 else 0
    vol_kabel_24 = round(inputs['kabel_24'] * 1.02) if inputs['kabel_24'] > 0 else 0

    if total_odp == 0:
        vol_puas = 0
    elif total_odp == 1:
        vol_puas = 1
    else:
        vol_puas = (total_odp * 2) - 1
    vol_puas += inputs['tiang_new'] + inputs['tiang_existing'] + inputs['tikungan']

    if inputs['sumber'] == "ODC":
        if inputs['kabel_12'] > 0:
            vol_os_sm_1_odc = 12 + total_odp
        elif inputs['kabel_24'] > 0:
            vol_os_sm_1_odc = 24 + total_odp
        else:
            vol_os_sm_1_odc = 0
    else:
        vol_os_sm_1_odc = 0

    vol_os_sm_1_odp = total_odp * 2 if inputs['sumber'] == "ODP" else 0
    vol_os_sm_1 = vol_os_sm_1_odc + vol_os_sm_1_odp

    vol_pc_upc = ((total_odp - 1) // 4) + 1 if total_odp > 0 else 0

    if vol_pc_upc == 1:
        vol_pc_apc = 18
    elif vol_pc_upc > 1:
        vol_pc_apc = vol_pc_upc * 2
    else:
        vol_pc_apc = 0

    vol_ps_1_4_odc = ((total_odp - 1) // 4) + 1 if inputs['sumber'] == "ODC" and total_odp > 0 else 0

    vol_tc_02_odc = 1 if inputs['sumber'] == "ODC" else 0
    vol_dd_hdpe = 6 if inputs['sumber'] == "ODC" else 0
    vol_bc_tr = 6 if inputs['sumber'] == "ODC" else 0

    return [
        {"designator": "AC-OF-SM-12-SC_O_STOCK", "volume": vol_kabel_12},
        {"designator": "AC-OF-SM-24-SC_O_STOCK", "volume": vol_kabel_24},
        {"designator": "ODP Solid-PB-8 AS", "volume": inputs['odp_8']},
        {"designator": "ODP Solid-PB-16 AS", "volume": inputs['odp_16']},
        {"designator": "PU-S7.0-400NM", "volume": inputs['tiang_new']},
        {"designator": "PU-AS", "volume": vol_puas},
        {"designator": "OS-SM-1-ODC", "volume": vol_os_sm_1_odc},
        {"designator": "OS-SM-1-ODP", "volume": vol_os_sm_1_odp},
        {"designator": "OS-SM-1", "volume": vol_os_sm_1},
        {"designator": "PC-UPC-652-2", "volume": vol_pc_upc},
        {"designator": "PC-APC/UPC-652-A1", "volume": vol_pc_apc},
        {"designator": "PS-1-4-ODC", "volume": vol_ps_1_4_odc},
        {"designator": "TC-02-ODC", "volume": vol_tc_02_odc},
        {"designator": "DD-HDPE-40-1", "volume": vol_dd_hdpe},
        {"designator": "BC-TR-0.6", "volume": vol_bc_tr},
        {"designator": "Preliminary Project HRB/Kawasan Khusus", "volume": 1 if inputs['izin'] else 0, "izin_value": float(inputs['izin']) if inputs['izin'] else 0}
    ]

# ======================
# 💾 DOWNLOAD OUTPUT & SUMMARY
# ======================
if st.session_state.boq_state.get('ready', False):
    st.subheader("📥 Download Results")
    st.download_button(
        label="⬇️ Download BOQ File",
        data=st.session_state.boq_state['excel_data'],
        file_name=f"BOQ_{st.session_state.boq_state['project_name']}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

    # Ringkasan hasil
    st.subheader("📌 Ringkasan")
    summary = st.session_state.boq_state.get("summary", {})
    col1, col2, col3, col4 = st.columns(4)
    col1.metric("MATERIAL", f"Rp {summary.get('material', 0):,.0f}")
    col2.metric("JASA", f"Rp {summary.get('jasa', 0):,.0f}")
    col3.metric("TOTAL", f"Rp {summary.get('total', 0):,.0f}")
    col4.metric("CPP", f"{summary.get('cpp', 0):.4f}")

    # Updated Items Table
    st.subheader("📋 Tabel Updated Items")
    st.dataframe(pd.DataFrame(st.session_state.boq_state['updated_items']))

    if st.button("🔄 Create New BOQ"):
        for key in ["kabel_12", "kabel_24", "odp_8", "odp_16", "tiang_new", "tiang_existing", "tikungan", "izin", "uploaded_file"]:
            if key in st.session_state:
                del st.session_state[key]
        st.session_state.boq_state = {'ready': False}
        st.rerun()
