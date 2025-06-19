import streamlit as st
import pandas as pd
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from io import BytesIO

# Konfigurasi awal
st.set_page_config("Form Input BOQ", layout="centered")
st.title("📋 Form Input BOQ Otomatis")

# Fungsi untuk authorize Google Sheets
def authorize_google_sheets():
    scope = ["https://spreadsheets.google.com/feeds", 
             "https://www.googleapis.com/auth/drive"]
    creds_dict = st.secrets["gcp_service_account"]
    creds = ServiceAccountCredentials.from_json_keyfile_dict(creds_dict, scope)
    return gspread.authorize(creds)

# Fungsi untuk download spreadsheet
def download_spreadsheet(client, url, filename):
    try:
        spreadsheet = client.open_by_url(url)
        # Export as Excel
        output = BytesIO()
        spreadsheet.export(format='xlsx', output=output)
        output.seek(0)
        return output
    except Exception as e:
        st.error(f"Gagal mendownload spreadsheet: {str(e)}")
        return None

# Initialize session state
if 'downloaded' not in st.session_state:
    st.session_state.downloaded = False

# URL spreadsheet
SPREADSHEET_URL = "https://docs.google.com/spreadsheets/d/1Zl0txYzsqslXjGV4Y4mcpVMB-vikTDCauzcLOfbbD5c/edit"

# Authorize Google Sheets
try:
    client = authorize_google_sheets()
except Exception as e:
    st.error(f"Gagal mengotorisasi Google Sheets: {str(e)}")
    st.stop()

# Form input
with st.form("boq_form"):
    # ... (bagian form input lainnya tetap sama) ...
    
    submitted = st.form_submit_button("🚀 Proses BOQ")
    reset_button = st.form_submit_button("🔄 Reset Form")

# ... (bagian proses BOQ lainnya tetap sama) ...

# Download Options
st.subheader("💾 Download")
    
tab1, tab2 = st.tabs(["Download Hasil BOQ", "Download Full RAB"])

with tab1:
    # ... (kode download hasil BOQ tetap sama) ...

with tab2:
    st.info("Download seluruh file RAB spreadsheet dari Google Sheets")
    
    # Tombol untuk generate full RAB
    if st.button("⬇️ Generate Full RAB Spreadsheet", key="generate_rab"):
        with st.spinner("Mempersiapkan file RAB..."):
            try:
                output_rab = download_spreadsheet(client, SPREADSHEET_URL, f"RAB_Lengkap_{st.session_state.lop_name}.xlsx")
                
                if output_rab:
                    st.session_state.downloaded = True
                    st.success("File RAB siap diunduh!")
                    
                    st.download_button(
                        label="💾 Klik untuk Download RAB Lengkap",
                        data=output_rab,
                        file_name=f"RAB_Lengkap_{st.session_state.lop_name}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        key="download_rab"
                    )
            except Exception as e:
                st.error(f"Terjadi kesalahan: {str(e)}")
                st.error("Pastikan Anda memiliki akses ke spreadsheet dan koneksi internet stabil")

# Reset after download
if st.session_state.downloaded:
    st.success("🎉 File telah berhasil diunduh!")
    if st.button("🔄 Buat Input Baru"):
        reset_form()
        st.rerun()
