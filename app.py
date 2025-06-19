import streamlit as st
import pandas as pd
from io import BytesIO
import openpyxl

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
    </style>
""", unsafe_allow_html=True)

st.title("📊 BOQ Generator (Custom Rules)")

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
    vol_puas = max(0, (total_odp * 2) - 1 + inputs['tiang_new'] + inputs['tiang_existing'] + inputs['tikungan'])

    vol_os_sm_1_odc = 0
    if inputs['sumber'] == "ODC":
        if inputs['kabel_12'] > 0:
            vol_os_sm_1_odc = 12 + total_odp
        elif inputs['kabel_24'] > 0:
            vol_os_sm_1_odc = 24 + total_odp

    vol_os_sm_1_odp = total_odp * 2 if inputs['sumber'] == "ODP" else 0
    vol_os_sm_1 = vol_os_sm_1_odc + vol_os_sm_1_odp

    vol_pc_upc = ((total_odp - 1) // 4) + 1 if total_odp > 0 else 0
    vol_pc_apc = 18 if vol_pc_upc == 1 else vol_pc_upc * 2 if vol_pc_upc > 1 else 0
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
# 🗅️ FORM UI
# ======================
with st.form("boq_form"):
    col1, col2 = st.columns([2, 1])
    with col1:
        lop_name = st.text_input("Nama LOP")
    with col2:
        sumber = st.radio("Sumber", ["ODC", "ODP"], index=0)

    kabel_12 = st.number_input("12 Core Cable (m)", min_value=0.0, value=0.0, key="kabel_12")
    kabel_24 = st.number_input("24 Core Cable (m)", min_value=0.0, value=0.0, key="kabel_24")
    odp_8 = st.number_input("ODP 8 Port", min_value=0, value=0, key="odp_8")
    odp_16 = st.number_input("ODP 16 Port", min_value=0, value=0, key="odp_16")
    tiang_new = st.number_input("Tiang Baru", min_value=0, value=0, key="tiang_new")
    tiang_existing = st.number_input("Tiang Eksisting", min_value=0, value=0, key="tiang_existing")
    tikungan = st.number_input("Tikungan", min_value=0, value=0, key="tikungan")

    izin = st.text_input("Preliminary (isi nominal jika ada)", value="")
    uploaded_file = st.file_uploader("Unggah Template BOQ", type=["xlsx"])
    submitted = st.form_submit_button("🚀 Generate BOQ")

if submitted:
    if not uploaded_file or not lop_name:
        st.warning("Lengkapi Nama LOP dan unggah file template!")
        st.stop()

    try:
        wb = openpyxl.load_workbook(uploaded_file)
        ws = wb.active
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
        for row in range(9, 289):
            designator = str(ws[f'B{row}'].value or "").strip()

            if izin and designator == "" and "Preliminary Project HRB/Kawasan Khusus" not in [str(ws[f'B{r}'].value) for r in range(9, 289)]:
                ws[f'B{row}'] = "Preliminary Project HRB/Kawasan Khusus"
                ws[f'F{row}'] = float(izin)
                ws[f'G{row}'] = 1
                updated_count += 1
                continue

            for item in items:
                if item["volume"] > 0 and designator == item["designator"]:
                    ws[f'G{row}'] = item["volume"]
                    if designator == "Preliminary Project HRB/Kawasan Khusus":
                        ws[f'F{row}'] = item.get("izin_value", 0)
                    updated_count += 1
                    break

        material = jasa = 0.0
        for row in range(9, 289):
            try:
                h_mat = ws[f'E{row}'].value
                h_jasa = ws[f'F{row}'].value
                vol = ws[f'G{row}'].value

                if all(isinstance(v, (int, float)) or (isinstance(v, str) and v.replace('.', '', 1).isdigit()) for v in [h_mat, h_jasa, vol]):
                    h_mat = float(h_mat)
                    h_jasa = float(h_jasa)
                    vol = float(vol)
                    material += h_mat * vol
                    jasa += h_jasa * vol
            except:
                continue

        total = material + jasa
        total_odp = odp_8 + odp_16
        cpp = round((total_odp * 8 / total), 4) if total else 0

        summary = {'material': material, 'jasa': jasa, 'total': total, 'cpp': cpp}

        output = BytesIO()
        wb.save(output)
        output.seek(0)

        st.session_state.boq_state = {
            'ready': True,
            'excel_data': output,
            'project_name': lop_name,
            'updated_items': [item for item in items if item['volume'] > 0],
            'summary': summary
        }

        st.success(f"✅ Berhasil update {updated_count} item!")

    except Exception as e:
        st.error(f"Terjadi kesalahan saat memproses: {str(e)}")

# ======================
# 📂 DOWNLOAD OUTPUT & SUMMARY
# ======================
if st.session_state.boq_state.get('ready', False):
    st.subheader("📥 Unduh BOQ")
    st.download_button(
        label="⬇️ Download BOQ File",
        data=st.session_state.boq_state['excel_data'],
        file_name=f"BOQ-{st.session_state.boq_state['project_name']}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

    st.subheader("📌 Ringkasan")
    summary = st.session_state.boq_state.get("summary", {})
    col1, col2, col3, col4 = st.columns(4)
    col1.metric("MATERIAL", f"Rp {summary.get('material', 0):,.0f}")
    col2.metric("JASA", f"Rp {summary.get('jasa', 0):,.0f}")
    col3.metric("TOTAL", f"Rp {summary.get('total', 0):,.0f}")
    col4.metric("CPP", f"{summary.get('cpp', 0):.4f}")

    st.subheader("📋 Item Terupdate")
    st.dataframe(pd.DataFrame(st.session_state.boq_state['updated_items']))

    if st.button("🔄 Buat BOQ Baru"):
        for key in list(st.session_state.keys()):
            del st.session_state[key]
        st.rerun()
else:
    st.info("⬆️ Isi form dan unggah template untuk memulai.")
