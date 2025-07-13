# 📊 BOQ Generator App with ADSS Support (Full Version)

import streamlit as st
import pandas as pd
from io import BytesIO
import openpyxl
import xml.etree.ElementTree as ET
from geopy.distance import geodesic
from datetime import datetime
import math
import re

# ------------------- Session State ------------------- #
def initialize_session_state():
    if 'boq_form_values' not in st.session_state:
        st.session_state.boq_form_values = {
            'lop_name': "",
            'sumber': "ODC",
            'kabel_12': 0.0,
            'kabel_24': 0.0,
            'odp_8': 0,
            'odp_16': 0,
            'tiang_new': 0,
            'tiang_existing': 0,
            'tikungan': 0,
            'izin': "",
            'closure': 0,
            'otb_12': 0,
            'uploaded_file': None,
            'kml_file': None
        }

    if 'boq_state' not in st.session_state:
        st.session_state.boq_state = {
            'ready': False,
            'excel_data': None,
            'project_name': "",
            'updated_items': [],
            'summary': {},
            'modified_kml': None,
            'active_tab': "manual"
        }

def reset_boq_application():
    initialize_session_state()
    st.session_state.boq_form_values.update({
        'kabel_12': 0.0,
        'kabel_24': 0.0,
        'odp_8': 0,
        'odp_16': 0,
        'otb_12': 0,
        'tiang_new': 0,
        'tiang_existing': 0,
        'tikungan': 0,
        'izin': "",
        'closure': 0,
        'uploaded_file': None,
        'kml_file': None
    })
    st.session_state.boq_state.update({
        'ready': False,
        'excel_data': None,
        'project_name': "",
        'updated_items': [],
        'summary': {},
        'modified_kml': None
    })

# ... (keep existing parse_kml_file_adss, generate_modified_kml, calculate_volumes_adss functions as they are) ...

# ------------------- Excel Processing ------------------- #
def process_boq_template(template_file, inputs, lop_name, custom_items=None):
    try:
        wb = openpyxl.load_workbook(template_file)
        ws = wb.active
        items = custom_items if custom_items else calculate_volumes_adss(inputs)

        for row in range(9, 289):
            cell_value = str(ws[f'B{row}'].value or "").strip()
            for item in items:
                if cell_value == item["designator"] and item["volume"] > 0:
                    ws[f'G{row}'] = item["volume"]
                    if "Preliminary" in cell_value and "izin_value" in item:
                        ws[f'F{row}'] = item["izin_value"]

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
        total_ports = (total_odp * 8) + (1 if inputs.get('otb_12', 0) > 0 else 0) * 8
        cpp = round(total / total_ports, 2) if total_ports > 0 else 0

        output = BytesIO()
        wb.save(output)
        output.seek(0)

        return {
            'excel_data': output,
            'summary': {
                'material': material,
                'jasa': jasa,
                'total': total,
                'cpp': cpp,
                'total_odp': total_odp,
                'total_ports': total_ports
            },
            'updated_items': [item for item in items if item['volume'] > 0]
        }

    except Exception as e:
        st.error(f"Error processing BOQ template: {str(e)}")
        return None

# ------------------- Streamlit UI: ADSS Form ------------------- #
def adss_kml_form():
    initialize_session_state()

    with st.form("adss_kml_form"):
        st.subheader("⚡ ADSS Project Info")

        col1, col2 = st.columns([2, 1])
        with col1:
            st.session_state.boq_form_values['lop_name'] = st.text_input("Nama LOP*", value=st.session_state.boq_form_values.get('lop_name', ""))
        with col2:
            st.session_state.boq_form_values['sumber'] = st.radio("Sumber*", ["ODC", "ODP"], index=0 if st.session_state.boq_form_values.get('sumber') == "ODC" else 1, horizontal=True)

        st.subheader("📤 Upload KML")
        st.session_state.boq_form_values['kml_file'] = st.file_uploader("Upload KML ADSS*", type=["kml"])

        st.subheader("📎 Additional Inputs")
        st.session_state.boq_form_values['tikungan'] = st.number_input("Tikungan*", min_value=0, value=st.session_state.boq_form_values.get('tikungan', 0))
        st.session_state.boq_form_values['izin'] = st.text_input("Preliminary (nominal jika ada)", value=st.session_state.boq_form_values.get('izin', ""))

        st.subheader("📥 Template BOQ")
        st.session_state.boq_form_values['uploaded_file'] = st.file_uploader("Upload Template Excel*", type=["xlsx"])

        submitted = st.form_submit_button("🚀 Generate BOQ + KML")
        if submitted:
            if not st.session_state.boq_form_values['kml_file'] or not st.session_state.boq_form_values['uploaded_file'] or not st.session_state.boq_form_values['lop_name']:
                st.error("Semua input wajib diisi.")
                return

            with st.spinner("⏳ Memproses KML dan Excel..."):
                parsed = parse_kml_file_adss(st.session_state.boq_form_values['kml_file'], st.session_state.boq_form_values['sumber'])
                st.session_state.boq_form_values.update(parsed)
                items = calculate_volumes_adss(st.session_state.boq_form_values)

                result = process_boq_template(
                    st.session_state.boq_form_values['uploaded_file'],
                    st.session_state.boq_form_values,
                    st.session_state.boq_form_values['lop_name'],
                    custom_items=items
                )

                if result:
                    mod_kml = generate_modified_kml(
                        st.session_state.boq_form_values['kml_file'].read(),
                        parsed['puas_sc_coords']
                    )
                    st.session_state.boq_state.update({
                        'ready': True,
                        'excel_data': result['excel_data'],
                        'project_name': st.session_state.boq_form_values['lop_name'],
                        'updated_items': result['updated_items'],
                        'summary': result['summary'],
                        'modified_kml': mod_kml
                    })
                    st.success("✅ BOQ & KML berhasil digenerate!")

# ------------------- Main App ------------------- #
def show():
    initialize_session_state()

    st.title("📊 BOQ Generator (ADSS Support)")
    tab1, tab2 = st.tabs(["⚡ ADSS KML Upload", "📥 Output"])

    with tab1:
        adss_kml_form()

    if st.session_state.boq_state.get('ready', False):
        with tab2:
            summary = st.session_state.boq_state['summary']
            st.metric("Total ODP", summary['total_odp'])
            st.metric("Total Port", summary['total_ports'])
            st.metric("Material", f"Rp {summary['material']:,.0f}")
            st.metric("Jasa", f"Rp {summary['jasa']:,.0f}")
            st.metric("Total Biaya", f"Rp {summary['total']:,.0f}")
            st.metric("CPP", f"Rp {summary['cpp']:,.0f}")

            st.dataframe(pd.DataFrame(st.session_state.boq_state['updated_items']), use_container_width=True)

            col1, col2 = st.columns([1, 1])
            with col1:
                st.download_button("⬇️ Download BOQ", data=st.session_state.boq_state['excel_data'], file_name=f"BOQ-{st.session_state.boq_state['project_name']}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            with col2:
                st.download_button("⬇️ Download KML", data=st.session_state.boq_state['modified_kml'], file_name="Updated-ADSS-KML.kml", mime="application/vnd.google-earth.kml+xml")

            if st.button("🔁 Reset BOQ"):
                reset_boq_application()
                st.rerun()

# ------------------- Run App ------------------- #
def main():
    show()

if __name__ == '__main__':
    main()
