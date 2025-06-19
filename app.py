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
st.title("📊 BOQ Generator (Designator from B9)")

# Initialize session state
if 'boq_state' not in st.session_state:
    st.session_state.boq_state = {
        'ready': False,
        'excel_data': None,
        'project_name': "",
        'updated_items': []
    }

# Designator to RAB Code Mapping
RAB_MAP = {
    "AC-OF-SM-12-SC_O_STOCK": "DC-01-01-1111",
    "AC-OF-SM-24-SC_O_STOCK": "DC-01-04-1100",
    "ODP Solid-PB-8 AS": "DC-01-08-4400",
    "ODP Solid-PB-16 AS": "DC-01-04-0400",
    "PU-S7.0-400NM": "DC-01-04-0410",
    "PU-AS": "DC-01-08-4280",
    "OS-SM-1-ODC": "AC-01-04-1100",
    "TC-02-ODC": "AC-01-04-2400",
    "DD-HDPE-40-1": "AC-01-04-0500",
    "BC-TR-0.6": "DC-01-04-1420",
    "PS-1-4-ODC": "DC-01-04-2420",
    "OS-SM-1-ODP": "DC-01-04-2460",
    "OS-SM-1": "DC-01-04-2480",
    "PC-UPC-652-2": "DC-01-04-2490",
    "PC-APC/UPC-652-A1": "DC-01-04-2500",
    "Preliminary Project HRB/Kawasan Khusus": "IZIN-KHUSUS-001"
}

# ======================
# 🔧 CORE FUNCTIONS
# ======================
def validate_template(ws):
    """Validate template structure with designator starting at B9"""
    try:
        # Check if row 8 has headers
        if not any(cell.value for cell in ws[8]):
            return False, "Row 8 should contain column headers"
        
        # Check if B9 has first designator
        if not ws['B9'].value:
            return False, "First designator should be at cell B9"
        
        return True, ""
    except Exception as e:
        return False, f"Validation error: {str(e)}"

def apply_header_style(ws):
    """Apply styling to header row (row 8)"""
    header_fill = PatternFill(start_color="4B0082", end_color="4B0082", fill_type="solid")
    header_font = Font(bold=True, color="FFFFFF")
    
    for cell in ws[8]:
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center")
    return ws

def calculate_volumes(inputs):
    """Calculate BOQ volumes"""
    total_odp = inputs['odp_8'] + inputs['odp_16']
    
    # Cable calculations
    vol_kabel_12 = round((inputs['kabel_12'] * 1.02) + (total_odp if inputs['sumber'] == "ODC" else 0)) if inputs['kabel_12'] > 0 else 0
    vol_kabel_24 = round((inputs['kabel_24'] * 1.02) + (total_odp if inputs['sumber'] == "ODC" else 0)) if inputs['kabel_24'] > 0 else 0
    
    # Support calculations
    vol_puas = (total_odp * 2 - 1) if total_odp > 1 else (1 if total_odp == 1 else 0)
    vol_puas += inputs['tiang_new'] + inputs['tiang_existing'] + inputs['tikungan']
    
    return [
        {"designator": "AC-OF-SM-12-SC_O_STOCK", "volume": vol_kabel_12},
        {"designator": "AC-OF-SM-24-SC_O_STOCK", "volume": vol_kabel_24},
        {"designator": "ODP Solid-PB-8 AS", "volume": inputs['odp_8']},
        {"designator": "ODP Solid-PB-16 AS", "volume": inputs['odp_16']},
        {"designator": "PU-S7.0-400NM", "volume": inputs['tiang_new']},
        {"designator": "PU-AS", "volume": vol_puas},
        {"designator": "OS-SM-1-ODC", "volume": (12 if inputs['kabel_12'] > 0 else 24 if inputs['kabel_24'] > 0 else 0) + total_odp if inputs['sumber'] == "ODC" else 0},
        {"designator": "TC-02-ODC", "volume": 1 if inputs['sumber'] == "ODC" else 0},
        {"designator": "DD-HDPE-40-1", "volume": 6 if inputs['sumber'] == "ODC" else 0},
        {"designator": "BC-TR-0.6", "volume": 6 if inputs['sumber'] == "ODC" else 0},
        {"designator": "PS-1-4-ODC", "volume": (total_odp - 1) // 4 + 1 if inputs['sumber'] == "ODC" and total_odp > 0 else 0},
        {"designator": "OS-SM-1-ODP", "volume": total_odp * 2 if inputs['sumber'] == "ODP" else 0},
        {"designator": "OS-SM-1", "volume": ((12 if inputs['kabel_12'] > 0 else 24 if inputs['kabel_24'] > 0 else 0) + total_odp) if inputs['sumber'] == "ODC" else (total_odp * 2)},
        {"designator": "PC-UPC-652-2", "volume": (total_odp - 1) // 4 + 1 if total_odp > 0 else 0},
        {"designator": "PC-APC/UPC-652-A1", "volume": 18 if ((total_odp - 1) // 4 + 1) == 1 else (((total_odp - 1) // 4 + 1) * 2 if ((total_odp - 1) // 4 + 1) > 1 else 0)},
        {"designator": "Preliminary Project HRB/Kawasan Khusus", "volume": 1 if inputs['izin'] else 0}
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
    
    izin = st.text_input("Special Permit (if any):", value="")
    
    uploaded_file = st.file_uploader("Upload Template (Header row 8, Designator from B9)", 
                                   type=["xlsx", "xls"])

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
        
        # Validate template structure
        is_valid, validation_msg = validate_template(ws)
        if not is_valid:
            st.error(f"Template validation failed: {validation_msg}")
            st.info("""
            Your template should have:
            1. Column headers in row 8
            2. First designator in cell B9
            """)
            st.stop()
        
        # Find VOL column (assume it's in row 8)
        vol_col = None
        for cell in ws[8]:
            if cell.value and "VOL" in str(cell.value).upper():
                vol_col = cell.column
                break
        
        if not vol_col:
            st.error("Could not find VOL column in header row (row 8)")
            st.stop()
        
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
        
        # Update volumes starting from B9
        updated_count = 0
        for row in range(9, ws.max_row + 1):
            designator = str(ws.cell(row=row, column=2).value or "").strip()  # Column B
            
            for item in items:
                if item["volume"] > 0 and designator == RAB_MAP.get(item["designator"], ""):
                    ws.cell(row=row, column=vol_col, value=item["volume"])
                    updated_count += 1
                    break
        
        # Apply styling
        ws = apply_header_style(ws)
        
        # Auto-adjust column widths
        for col in ws.columns:
            max_length = max(
                len(str(cell.value or "")) for cell in col
            )
            adjusted_width = (max_length + 2) * 1.2
            ws.column_dimensions[get_column_letter(col[0].column)].width = adjusted_width
        
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
        
        st.success(f"✅ Successfully updated {updated_count} items!")
        
        # Show updated items
        with st.expander("📋 Updated Items"):
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
