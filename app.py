import streamlit as st
import pandas as pd
from io import BytesIO
import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill
from openpyxl.utils import get_column_letter

# ======================
# 🎩 MAGIC CONFIGURATION
# ======================
st.set_page_config("MAGIC BOQ Generator", layout="centered", page_icon="🧙")
st.title("🧙‍♂️ MAGIC BOQ GENERATOR")

# Initialize session state
if 'magic_state' not in st.session_state:
    st.session_state.magic_state = {
        'ready': False,
        'excel_data': None,
        'project_name': "",
        'updated_items': []
    }

# ======================
# 🔮 MAGIC MAPPINGS
# ======================
RAB_MAGIC_MAP = {
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
# 🪄 MAGIC FUNCTIONS
# ======================
def apply_magic_styles(ws):
    """Apply magical styles to worksheet"""
    header_fill = PatternFill(start_color="4B0082", end_color="4B0082", fill_type="solid")
    header_font = Font(bold=True, color="FFFFFF")
    center_align = Alignment(horizontal="center", vertical="center")
    
    # Style header row
    for cell in ws[7]:
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = center_align
    
    # Auto-adjust column widths
    for col in ws.columns:
        max_length = max(len(str(cell.value or "")) for cell in col)
        adjusted_width = (max_length + 2) * 1.2
        ws.column_dimensions[get_column_letter(col[0].column)].width = adjusted_width
    
    return ws

def calculate_magic_volumes(inputs):
    """Calculate magical volumes"""
    total_odp = inputs['odp_8'] + inputs['odp_16']
    
    # Cable calculations
    vol_kabel_12 = round((inputs['kabel_12'] * 1.02) + (total_odp if inputs['sumber'] == "ODC" else 0)) if inputs['kabel_12'] > 0 else 0
    vol_kabel_24 = round((inputs['kabel_24'] * 1.02) + (total_odp if inputs['sumber'] == "ODC" else 0)) if inputs['kabel_24'] > 0 else 0
    
    # Support calculations
    vol_puas = (total_odp * 2 - 1) if total_odp > 1 else (1 if total_odp == 1 else 0)
    vol_puas += inputs['tiang_new'] + inputs['tiang_existing'] + inputs['tikungan']
    
    # Prepare items
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

def validate_template(ws):
    """Validate Google Sheets template structure"""
    required_cells = {'B1': "DATA MATERIAL SATUAN"}
    required_headers = ["NO", "DESIGNATOR", "VOL"]
    
    errors = []
    
    # Check required cells
    for cell_ref, expected_value in required_cells.items():
        cell = ws[cell_ref]
        if not cell.value or str(cell.value).strip() != expected_value:
            errors.append(f"Cell {cell_ref} should contain '{expected_value}'")
    
    # Check required headers (row 7)
    header_row = [str(cell.value or "").strip().upper() for cell in ws[7]]
    for header in required_headers:
        if header not in " ".join(header_row):
            errors.append(f"Header '{header}' not found in row 7")
    
    return errors

# ======================
# ✨ MAGIC FORM
# ======================
with st.form("magic_form"):
    st.subheader("🔮 Project Details")
    col1, col2 = st.columns(2)
    with col1:
        sumber = st.radio("Source Type:", ["ODC", "ODP"], index=0, help="Select ODC for Central Office or ODP for Distribution Point")
    with col2:
        lop_name = st.text_input("LOP Name:", help="Will be used for output filename")
        project_name = st.text_input("Project Name:")
        sto_code = st.text_input("STO Code:")

    st.subheader("📡 Cable Inputs")
    col1, col2 = st.columns(2)
    with col1:
        kabel_12 = st.number_input("12 Core Cable (meters):", min_value=0.0, value=0.0)
    with col2:
        kabel_24 = st.number_input("24 Core Cable (meters):", min_value=0.0, value=0.0)

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
    
    izin = st.text_input("Special Permit (if any):", value="", help="Leave empty if no special permit required")

    uploaded_file = st.file_uploader("Upload Template File", type=["xlsx", "xls"], 
                                   help="Upload your BOQ template Excel file")

    submitted = st.form_submit_button("✨ GENERATE MAGIC BOQ")

# ======================
# 🎯 FORM PROCESSING
# ======================
if submitted:
    # Input validation
    if not all([lop_name, project_name, sto_code]):
        st.error("Please complete all project details!")
        st.stop()
    
    if not uploaded_file:
        st.error("Please upload template file!")
        st.stop()

    try:
        # Load and validate template
        wb = openpyxl.load_workbook(uploaded_file)
        ws = wb.active
        
        template_errors = validate_template(ws)
        if template_errors:
            st.error("Template validation failed!")
            with st.expander("See errors"):
                for error in template_errors:
                    st.error(f"• {error}")
            st.stop()
        
        # Calculate volumes
        magic_inputs = {
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
        magic_items = calculate_magic_volumes(magic_inputs)
        
        # Update template
        ws['B1'] = "DATA MATERIAL SATUAN"
        ws['B2'] = f"PENGADAAN DAN PEMASANGAN GRANULAR MODERNIZATION"
        ws['B3'] = f"PROJECT : {project_name}"
        ws['B4'] = f"STO : {sto_code}"
        
        # Find VOL column (column G)
        vol_col = 7
        updated_count = 0
        
        for row in range(8, ws.max_row + 1):
            designator = str(ws.cell(row=row, column=2).value or "").strip()
            
            for item in magic_items:
                if item["volume"] > 0 and designator == RAB_MAGIC_MAP.get(item["designator"], ""):
                    ws.cell(row=row, column=vol_col, value=item["volume"])
                    updated_count += 1
                    break
        
        # Apply styles
        ws = apply_magic_styles(ws)
        
        # Save to BytesIO
        output = BytesIO()
        wb.save(output)
        output.seek(0)
        
        # Update session state
        st.session_state.magic_state = {
            'ready': True,
            'excel_data': output,
            'project_name': lop_name,
            'updated_items': [item for item in magic_items if item['volume'] > 0]
        }
        
        st.success(f"Successfully updated {updated_count} items!")
        
        # Show updated items
        st.subheader("Updated Items")
        st.dataframe(pd.DataFrame(st.session_state.magic_state['updated_items']))
        
    except Exception as e:
        st.error(f"Magic failed: {str(e)}")

# ======================
# 💾 DOWNLOAD SECTION
# ======================
if st.session_state.magic_state.get('ready', False):
    st.subheader("📥 Download Results")
    st.download_button(
        label="⬇️ Download Magic BOQ",
        data=st.session_state.magic_state['excel_data'],
        file_name=f"BOQ_{st.session_state.magic_state['project_name']}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
    
    if st.button("🔄 Start New BOQ"):
        st.session_state.magic_state = {'ready': False}
        st.rerun()
