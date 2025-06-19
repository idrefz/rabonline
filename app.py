import streamlit as st
import pandas as pd
from io import BytesIO
import openpyxl
from openpyxl.utils import get_column_letter
from openpyxl.styles import Font, Alignment

# ✨ Magic Configuration
st.set_page_config("Magic BOQ Generator", layout="centered", page_icon="🧙")
st.title("🧙 Magic BOQ Generator")

# 🪄 Initialize session state
if 'magic_state' not in st.session_state:
    st.session_state.magic_state = {
        'ready': False,
        'data': None,
        'project_name': "",
        'debug_log': [],
        'style_applied': False
    }

# 🔮 Designator to RAB Code Mapping
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

# 🧙 Magic Functions
def apply_magic_styles(worksheet):
    """Apply magical styles to the worksheet"""
    header_font = Font(bold=True, color="FFFFFF")
    header_fill = openpyxl.styles.PatternFill(start_color="4B0082", end_color="4B0082", fill_type="solid")
    alignment = Alignment(horizontal="center", vertical="center")
    
    for row in worksheet.iter_rows(min_row=7, max_row=7):
        for cell in row:
            cell.font = header_font
            cell.fill = header_fill
            cell.alignment = alignment
    
    for column in worksheet.columns:
        max_length = 0
        column_letter = get_column_letter(column[0].column)
        for cell in column:
            try:
                if len(str(cell.value)) > max_length:
                    max_length = len(str(cell.value))
            except:
                pass
        adjusted_width = (max_length + 2) * 1.2
        worksheet.column_dimensions[column_letter].width = adjusted_width
    
    return worksheet

def calculate_magic_volumes(inputs):
    """Perform magical volume calculations"""
    total_odp = inputs['odp_8'] + inputs['odp_16']
    
    # ✨ Magic formula for cable calculations
    kabel_12 = inputs['kabel_12']
    kabel_24 = inputs['kabel_24']
    sumber = inputs['sumber']
    
    vol_kabel_12 = round((kabel_12 * 1.02) + (total_odp if sumber == "ODC" else 0)) if kabel_12 > 0 else 0
    vol_kabel_24 = round((kabel_24 * 1.02) + (total_odp if sumber == "ODC" else 0)) if kabel_24 > 0 else 0
    
    # � Magic formula for PU-AS
    vol_puas = (total_odp * 2 - 1) if total_odp > 1 else (1 if total_odp == 1 else 0)
    vol_puas += inputs['tiang_new'] + inputs['tiang_existing'] + inputs['tikungan']
    
    # 🪄 Prepare magic items
    magic_items = [
        {"designator": "AC-OF-SM-12-SC_O_STOCK", "volume": vol_kabel_12},
        {"designator": "AC-OF-SM-24-SC_O_STOCK", "volume": vol_kabel_24},
        {"designator": "ODP Solid-PB-8 AS", "volume": inputs['odp_8']},
        {"designator": "ODP Solid-PB-16 AS", "volume": inputs['odp_16']},
        {"designator": "PU-S7.0-400NM", "volume": inputs['tiang_new']},
        {"designator": "PU-AS", "volume": vol_puas},
        {"designator": "OS-SM-1-ODC", "volume": (12 if kabel_12 > 0 else 24 if kabel_24 > 0 else 0) + total_odp if sumber == "ODC" else 0},
        {"designator": "TC-02-ODC", "volume": 1 if sumber == "ODC" else 0},
        {"designator": "DD-HDPE-40-1", "volume": 6 if sumber == "ODC" else 0},
        {"designator": "BC-TR-0.6", "volume": 6 if sumber == "ODC" else 0},
        {"designator": "PS-1-4-ODC", "volume": (total_odp - 1) // 4 + 1 if sumber == "ODC" and total_odp > 0 else 0},
        {"designator": "OS-SM-1-ODP", "volume": total_odp * 2 if sumber == "ODP" else 0},
        {"designator": "OS-SM-1", "volume": ((12 if kabel_12 > 0 else 24 if kabel_24 > 0 else 0) + total_odp) if sumber == "ODC" else (total_odp * 2)},
        {"designator": "PC-UPC-652-2", "volume": (total_odp - 1) // 4 + 1 if total_odp > 0 else 0},
        {"designator": "PC-APC/UPC-652-A1", "volume": 18 if ((total_odp - 1) // 4 + 1) == 1 else (((total_odp - 1) // 4 + 1) * 2 if ((total_odp - 1) // 4 + 1) > 1 else 0)},
        {"designator": "Preliminary Project HRB/Kawasan Khusus", "volume": 1 if inputs['izin'] else 0}
    ]
    
    return magic_items

# 🔮 Magic Form
with st.form("magic_boq_form"):
    st.subheader("🔮 Project Details")
    col1, col2 = st.columns(2)
    with col1:
        sumber = st.radio("Source Type:", ["ODC", "ODP"], index=0)
    with col2:
        lop_name = st.text_input("LOP Name (for filename):")
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
    tiang_new = st.number_input("New Poles:", min_value=0, value=0)
    tiang_existing = st.number_input("Existing Poles:", min_value=0, value=0)
    tikungan = st.number_input("Number of Bends:", min_value=0, value=0)
    izin = st.text_input("Special Permit (if any):", value="")

    uploaded_file = st.file_uploader("Upload RAB Template", type=["xlsx", "xls"])

    submitted = st.form_submit_button("✨ Generate Magic BOQ")

if submitted:
    # 🧹 Input Validation
    if not all([lop_name, project_name, sto_code]):
        st.error("🦄 Please complete all project details!")
        st.stop()
    
    if not uploaded_file:
        st.error("📄 Please upload the RAB template file!")
        st.stop()

    try:
        # 📜 Prepare magic inputs
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
        
        # 🔢 Calculate magic volumes
        magic_items = calculate_magic_volumes(magic_inputs)
        
        # 📖 Load the magic template
        wb = openpyxl.load_workbook(uploaded_file)
        ws = wb.active
        
        # 🏰 Validate the castle... I mean template
        if ws['B1'].value != "DATA MATERIAL SATUAN":
            st.error("🧙‍♂️ This doesn't look like the right template! The magic won't work.")
            st.stop()

        # ✍️ Update project headers
        ws['B1'] = "DATA MATERIAL SATUAN"
        ws['B2'] = f"PENGADAAN DAN PEMASANGAN GRANULAR MODERNIZATION"
        ws['B3'] = f"PROJECT : {project_name}"
        ws['B4'] = f"STO : {sto_code}"

        # 🔍 Find the VOL column (magic detection)
        vol_col = None
        for col in range(1, ws.max_column + 1):
            cell_value = str(ws.cell(row=7, column=col).value).strip().upper()
            if "VOL" in cell_value:
                vol_col = col
                break
        
        vol_col = vol_col or 7  # Default to column 7 if not found
        
        # 🪄 Update volumes with magic
        updated_count = 0
        for row in range(8, ws.max_row + 1):
            designator_cell = ws.cell(row=row, column=2)
            rab_code = str(designator_cell.value).strip() if designator_cell.value else ""
            
            for item in magic_items:
                if item["volume"] > 0 and rab_code == RAB_MAGIC_MAP.get(item["designator"], ""):
                    ws.cell(row=row, column=vol_col, value=item["volume"])
                    updated_count += 1
                    break

        # 💅 Apply magical styling
        ws = apply_magic_styles(ws)
        
        # 💾 Save the magic
        output = BytesIO()
        wb.save(output)
        output.seek(0)

        # 🏆 Store the magic results
        st.session_state.magic_state = {
            'ready': True,
            'data': output,
            'project_name': lop_name,
            'updated_count': updated_count,
            'magic_items': magic_items,
            'style_applied': True
        }

        st.success(f"🎉 Successfully updated {updated_count} items with magic!")
        
        # 📊 Show the magic results
        st.subheader("✨ Magic Results")
        result_df = pd.DataFrame([item for item in magic_items if item['volume'] > 0])
        st.dataframe(result_df.style.highlight_max(axis=0, color='#E6E6FA'))

    except Exception as e:
        st.error(f"🧨 Magic failed: {str(e)}")

# 🎁 Display download button if magic is ready
if st.session_state.magic_state.get('ready', False):
    st.subheader("💾 Download Your Magic BOQ")
    st.download_button(
        label="⬇️ Download Magical RAB",
        data=st.session_state.magic_state['data'],
        file_name=f"MAGIC_RAB_{st.session_state.magic_state['project_name']}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

    if st.button("🔄 Start New Magic"):
        st.session_state.magic_state = {
            'ready': False,
            'data': None,
            'project_name': "",
            'debug_log': [],
            'style_applied': False
        }
        st.rerun()
