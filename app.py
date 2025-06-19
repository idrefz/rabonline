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
        'updated_items': []
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

    if total_odp == 0:
        vol_pc_upc = 0
    else:
        vol_pc_upc = ((total_odp - 1) // 4) + 1

    if vol_pc_upc == 1:
        vol_pc_apc = 18
    elif vol_pc_upc > 1:
        vol_pc_apc = vol_pc_upc * 2
    else:
        vol_pc_apc = 0

    if inputs['sumber'] == "ODC" and total_odp > 0:
        vol_ps_1_4_odc = ((total_odp - 1) // 4) + 1
    else:
        vol_ps_1_4_odc = 0

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
        {"designator": "Preliminary Project HRB/Kawasan Khusus", "volume": 1 if inputs['izin'] else 0}
    ]

# ======================
# 🖥️ USER INTERFACE
# ======================
with st.form("boq_form"):
    st.subheader("📋 Project Information")
    col1, col2 = st.columns(2)
    with col1:
        sumber = st.radio("Source Type:", ["ODC", "ODP"], index=0)
    with col2:
        lop_name = st.text_input("LOP Name:")
        project_name = st.text_input("Project Name:")
        sto_code = st.text_input("STO Code:")

    st.subheader("📡 Cable Inputs")
    col1, col2 = st.columns(2)
    with col1:
        kabel_12 = st.number_input("12 Core Cable (m):", min_value=0.0, value=0.0)
    with col2:
        kabel_24 = st.number_input("24 Core Cable (m):", min_value=0.0, value=0.0)

    st.subheader("🏗️ ODP Inputs")
    col1, col2 = st.columns(2)
    with col1:
        odp_8 = st.number_input("ODP 8 Port:", min_value=0, value=0)
    with col2:
        odp_16 = st.number_input("ODP 16 Port:", min_value=0, value=0)

    st.subheader("⚙️ Support Inputs")
    col1, col2, col3 = st.columns(3)
    with col1:
        tiang_new = st.number_input("New Poles:", min_value=0, value=0)
    with col2:
        tiang_existing = st.number_input("Existing Poles:", min_value=0, value=0)
    with col3:
        tikungan = st.number_input("Bends:", min_value=0, value=0)

    izin = st.text_input("Special Permit (if any):", value="")
    uploaded_file = st.file_uploader("Upload BOQ Template", type=["xlsx", "xls"])

    submitted = st.form_submit_button("🚀 Generate BOQ")

if submitted:
    if not all([lop_name, project_name, sto_code]):
        st.warning("Please complete all project information fields!")
        st.stop()

    if not uploaded_file:
        st.warning("Please upload a template file!")
        st.stop()

    try:
        wb = openpyxl.load_workbook(uploaded_file)
        ws = wb.active

        ws['B1'] = "DAFTAR HARGA SATUAN"
        ws['B2'] = "PENGADAAN DAN PEMASANGAN GRANULAR MODERNIZATION"
        ws['B3'] = f"PROJECT : {project_name}"
        ws['B4'] = f"STO : {sto_code}"

        input_data = {
            'sumber': sumber,
            'kabel_12': kabel_12,
            'kabel_24': kabel_24,
            'odp_8': odp_8,
            'odp_16': odp_16,
            'tiang_new': tiang_new,
            'tiang_existing': tiang_existing,
            'tikungan': tikungan,
            'izin': izin
        }
        items = calculate_volumes(input_data)

        updated_count = 0
        special_permit_added = False

        for row in range(9, 289):
            designator = str(ws[f'B{row}'].value or "").strip()

            if not special_permit_added and izin and row > 9 and not ws[f'B{row}'].value:
                ws[f'B{row}'] = "Preliminary Project HRB/Kawasan Khusus"
                ws[f'F{row}'] = izin
                ws[f'G{row}'] = 1
                special_permit_added = True
                updated_count += 1
                continue

            for item in items:
                if item["volume"] > 0 and designator == item["designator"]:
                    ws[f'G{row}'] = item["volume"]
                    updated_count += 1
                    break

        output = BytesIO()
        wb.save(output)
        output.seek(0)

        st.session_state.boq_state = {
            'ready': True,
            'excel_data': output,
            'project_name': lop_name,
            'updated_items': [item for item in items if item['volume'] > 0]
        }

        st.success(f"✅ Successfully updated {updated_count} items!")

        with st.expander("📋 Updated Items"):
            st.dataframe(pd.DataFrame(st.session_state.boq_state['updated_items']))

    except Exception as e:
        st.error(f"Error processing BOQ: {str(e)}")

if st.session_state.boq_state.get('ready', False):
    st.subheader("📥 Download Results")
    st.download_button(
        label="⬇️ Download BOQ File",
        data=st.session_state.boq_state['excel_data'],
        file_name=f"BOQ_{st.session_state.boq_state['project_name']}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

    if st.button("🔄 Create New BOQ"):
        st.session_state.boq_state = {'ready': False}
        st.rerun()
