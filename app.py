import streamlit as st
import pandas as pd
from io import BytesIO
import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill
from openpyxl.utils import get_column_letter

# ======================
# 🎩 CONFIGURATION
# ======================
st.set_page_config("BOQ Generator", layout="centered")
st.title("📊 BOQ Generator (Autofill VOL)")

# Initialize session state
if 'boq_state' not in st.session_state:
    st.session_state.boq_state = {
        'ready': False,
        'excel_data': None,
        'project_name': "",
        'updated_items': []
    }

# Designator to RAB Code Mapping (adjusted for your template)
RAB_MAP = {
    "DC-01-01-1111": "DC-01-01-1111",  # AC-OF-SM-12-SC_O_STOCK
    "DC-01-04-1100": "DC-01-04-1100",  # AC-OF-SM-24-SC_O_STOCK
    "DC-01-08-4400": "DC-01-08-4400",  # ODP Solid-PB-8 AS
    "DC-01-04-0400": "DC-01-04-0400",  # ODP Solid-PB-16 AS
    "DC-01-04-0410": "DC-01-04-0410",  # PU-S7.0-400NM
    "DC-01-08-4280": "DC-01-08-4280",  # PU-AS
    "AC-01-04-1100": "AC-01-04-1100",  # OS-SM-1-ODC
    "AC-01-04-2400": "AC-01-04-2400",  # TC-02-ODC
    "AC-01-04-0500": "AC-01-04-0500",  # DD-HDPE-40-1
    "DC-01-04-1420": "DC-01-04-1420",  # BC-TR-0.6
    "DC-01-04-2420": "DC-01-04-2420",  # PS-1-4-ODC
    "DC-01-04-2460": "DC-01-04-2460",  # OS-SM-1-ODP
    "DC-01-04-2480": "DC-01-04-2480",  # OS-SM-1
    "DC-01-04-2490": "DC-01-04-2490",  # PC-UPC-652-2
    "DC-01-04-2500": "DC-01-04-2500",  # PC-APC/UPC-652-A1
    "IZIN-KHUSUS-001": "IZIN-KHUSUS-001"  # Preliminary Project HRB/Kawasan Khusus
}

# ======================
# 🔧 CORE FUNCTIONS
# ======================
def calculate_volumes(inputs):
    """Calculate BOQ volumes based on inputs"""
    total_odp = inputs['odp_8'] + inputs['odp_16']
    
    # Cable calculations
    vol_kabel_12 = round((inputs['kabel_12'] * 1.02) + (total_odp if inputs['sumber'] == "ODC" else 0)) if inputs['kabel_12'] > 0 else 0
    vol_kabel_24 = round((inputs['kabel_24'] * 1.02) + (total_odp if inputs['sumber'] == "ODC" else 0)) if inputs['kabel_24'] > 0 else 0
    
    # Support calculations
    vol_puas = (total_odp * 2 - 1) if total_odp > 1 else (1 if total_odp == 1 else 0)
    vol_puas += inputs['tiang_new'] + inputs['tiang_existing'] + inputs['tikungan']
    
    return [
        {"designator": "DC-01-01-1111", "volume": vol_kabel_12},  # 12 Core
        {"designator": "DC-01-04-1100", "volume": vol_kabel_24},  # 24 Core
        {"designator": "DC-01-08-4400", "volume": inputs['odp_8']},  # ODP 8
        {"designator": "DC-01-04-0400", "volume": inputs['odp_16']},  # ODP 16
        {"designator": "DC-01-04-0410", "volume": inputs['tiang_new']},  # New Poles
        {"designator": "DC-01-08-4280", "volume": vol_puas},  # PU-AS
        {"designator": "AC-01-04-1100", "volume": (12 if inputs['kabel_12'] > 0 else 24 if inputs['kabel_24'] > 0 else 0) + total_odp if inputs['sumber'] == "ODC" else 0},  # OS-SM-1-ODC
        {"designator": "AC-01-04-2400", "volume": 1 if inputs['sumber'] == "ODC" else 0},  # TC-02-ODC
        {"designator": "AC-01-04-0500", "volume": 6 if inputs['sumber'] == "ODC" else 0},  # DD-HDPE-40-1
        {"designator": "DC-01-04-1420", "volume": 6 if inputs['sumber'] == "ODC" else 0},  # BC-TR-0.6
        {"designator": "DC-01-04-2420", "volume": (total_odp - 1) // 4 + 1 if inputs['sumber'] == "ODC" and total_odp > 0 else 0},  # PS-1-4-ODC
        {"designator": "DC-01-04-2460", "volume": total_odp * 2 if inputs['sumber'] == "ODP" else 0},  # OS-SM-1-ODP
        {"designator": "DC-01-04-2480", "volume": ((12 if inputs['kabel_12'] > 0 else 24 if inputs['kabel_24'] > 0 else 0) + total_odp) if inputs['sumber'] == "ODC" else (total_odp * 2)},  # OS-SM-1
        {"designator": "DC-01-04-2490", "volume": (total_odp - 1) // 4 + 1 if total_odp > 0 else 0},  # PC-UPC-652-2
        {"designator": "DC-01-04-2500", "volume": 18 if ((total_odp - 1) // 4 + 1) == 1 else (((total_odp - 1) // 4 + 1) * 2 if ((total_odp - 1) // 4 + 1) > 1 else 0)},  # PC-APC/UPC-652-A1
        {"designator": "IZIN-KHUSUS-001", "volume": 1 if inputs['izin'] else 0}  # Izin Khusus
    ]

# ======================
# 🖥️ USER INTERFACE
# ======================
with st.form("boq_form"):
    st.subheader("📋 Project Information")
    col1, col2 = st.columns(2)
    with col1:
        sumber = st.selectbox("Source Type:", ["ODC", "ODP"], index=0)
        project_name = st.text_input("Project Name:")
    with col2:
        lop_name = st.text_input("LOP Name:")
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
    
    izin = st.text_input("Special Permit (IZIN-KHUSUS-001):", value="")
    
    uploaded_file = st.file_uploader("Upload BOQ Template", type=["xlsx", "xls"])

    submitted = st.form_submit_button("🚀 Generate BOQ")

# ======================
# 🔄 PROCESSING
# ======================
if submitted:
    # Validate inputs
    if not all([lop_name, project_name, sto_code]):
        st.warning("Please complete all project information fields!")
        st.stop()
    
    if not uploaded_file:
        st.warning("Please upload a template file!")
        st.stop()

    try:
        # Load workbook
        wb = openpyxl.load_workbook(uploaded_file)
        ws = wb.active
        
        # Update project info (rows 1-4)
        ws['B1'] = "DAFTAR HARGA SATUAN"
        ws['B2'] = "PENGADAAN DAN PEMASANGAN GRANULAR MODERNIZATION"
        ws['B3'] = f"PROJECT : {project_name}"
        ws['B4'] = f"STO : {sto_code}"
        
        # Calculate volumes
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
        
        # Update volumes (B9:B288 to G9:G288)
        updated_count = 0
        special_permit_added = False
        
        for row in range(9, 289):  # From row 9 to 288
            designator = str(ws[f'B{row}'].value or "").strip()
            
            # Special handling for permit (add to first empty row if needed)
            if not special_permit_added and izin and row > 9 and not ws[f'B{row}'].value:
                ws[f'B{row}'] = "IZIN-KHUSUS-001"
                ws[f'F{row}'] = izin  # Special permit value in column F
                ws[f'G{row}'] = 1     # Volume 1 in column G
                special_permit_added = True
                updated_count += 1
                continue
                
            for item in items:
                if item["volume"] > 0 and designator == item["designator"]:
                    ws[f'G{row}'] = item["volume"]  # Update VOL in column G
                    updated_count += 1
                    break
        
        # Save output
        output = BytesIO()
        wb.save(output)
        output.seek(0)
        
        # Update session state
        st.session_state.boq_state = {
            'ready': True,
            'excel_data': output,
            'project_name': lop_name,
            'updated_items': [item for item in items if item['volume'] > 0]
        }
        
        st.success(f"✅ Successfully updated {updated_count} items in columns B9:G288!")
        
        # Show updated items
        with st.expander("📋 Updated Items Preview"):
            st.dataframe(pd.DataFrame(st.session_state.boq_state['updated_items']))
            
    except Exception as e:
        st.error(f"Error processing BOQ: {str(e)}")

# ======================
# 💾 DOWNLOAD OUTPUT
# ======================
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
