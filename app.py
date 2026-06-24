import streamlit as st
import os
import shutil
import uuid
from kml_parser import KMLParser
from boq_mapper import BOQMapper
from excel_generator import ExcelGenerator
from config import BASE_DIR, MAPPING_FILE

# Set page title and favicon
st.set_page_config(
    page_title="FTTH BOQ Generator - Telkom",
    page_icon="🔴",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom Premium CSS for UI styling (Matching Telkom brand and modern layout)
st.markdown("""
    <style>
        /* Headings & general font styling */
        h1, h2, h3, h4, h5, h6 {
            font-family: 'Plus Jakarta Sans', sans-serif !important;
            font-weight: 700 !important;
            color: #1A202C !important;
        }
        
        /* Main logo banner */
        .header-container {
            display: flex;
            align-items: center;
            gap: 20px;
            padding: 24px;
            background: white;
            border-radius: 16px;
            box-shadow: 0 10px 30px rgba(0,0,0,0.02);
            border: 1px solid #ECEFF5;
            margin-bottom: 30px;
        }
        .header-logo {
            width: 60px;
            height: 60px;
            object-fit: contain;
        }
        
        /* Sidebar layout styling */
        [data-testid="stSidebar"] {
            background: linear-gradient(180deg, #E31E24 0%, #B31015 100%) !important;
            color: white !important;
        }
        [data-testid="stSidebar"] h1, [data-testid="stSidebar"] h2, [data-testid="stSidebar"] h3, 
        [data-testid="stSidebar"] h4, [data-testid="stSidebar"] h5, [data-testid="stSidebar"] h6,
        [data-testid="stSidebar"] label, [data-testid="stSidebar"] p, [data-testid="stSidebar"] span {
            color: white !important;
        }
        
        /* Sidebar separator */
        .sidebar-divider {
            height: 1px;
            background: rgba(255,255,255,0.15);
            margin: 20px 0;
        }
        
        /* File uploader hover styling */
        [data-testid="stFileUploader"] {
            border: 2px dashed rgba(227, 30, 36, 0.3) !important;
            background-color: #FAFAFC !important;
            border-radius: 16px !important;
            padding: 20px !important;
            transition: all 0.3s ease;
        }
        [data-testid="stFileUploader"]:hover {
            border-color: #E31E24 !important;
            background-color: #FFF8F8 !important;
        }
        
        /* Primary button custom layout */
        div.stButton > button:first-child {
            background: linear-gradient(135deg, #E31E24 0%, #FF4D52 100%) !important;
            color: white !important;
            border: none !important;
            padding: 12px 24px !important;
            border-radius: 12px !important;
            font-weight: 700 !important;
            box-shadow: 0 4px 12px rgba(227,30,36,0.2) !important;
            transition: all 0.3s ease !important;
            width: 100% !important;
            height: 50px !important;
        }
        div.stButton > button:first-child:hover {
            transform: translateY(-2px) !important;
            box-shadow: 0 8px 20px rgba(227,30,36,0.3) !important;
        }
        
        /* Custom styles for instructions card */
        .info-card {
            background: white;
            padding: 30px;
            border-radius: 20px;
            border: 1px solid #ECEFF5;
            box-shadow: 0 10px 30px rgba(0,0,0,0.02);
            margin-top: 20px;
        }
        .info-title {
            display: flex;
            align-items: center;
            gap: 10px;
            font-size: 16px;
            font-weight: 700;
            color: #1A202C;
            margin-bottom: 15px;
        }
        .info-title i {
            color: #E31E24;
        }
    </style>
""", unsafe_allow_html=True)

# Define directories
UPLOAD_DIR = BASE_DIR / "uploads"
OUTPUT_DIR = BASE_DIR / "outputs"
UPLOAD_DIR.mkdir(exist_ok=True)
OUTPUT_DIR.mkdir(exist_ok=True)

# ----------------- SIDEBAR CONFIG -----------------
st.sidebar.markdown("""
    <div style='text-align: center; margin-bottom: 20px;'>
        <h3 style='margin: 0; font-weight: 800; letter-spacing: -0.5px;'>FTTH BOQ</h3>
        <p style='font-size: 12px; opacity: 0.8;'>PT TELKOM INDONESIA</p>
    </div>
    <div class='sidebar-divider'></div>
""", unsafe_allow_html=True)

st.sidebar.markdown("### ⚙️ Pengaturan")
network_type_display = st.sidebar.selectbox(
    "Konfigurasi Jaringan",
    options=["🏡 Perumahan / Komplek", "🏢 Ruko", "🚫 None / Tanpa HH PIT"]
)

# Convert display select to mapping parameter
network_type = "perumahan"
if "Ruko" in network_type_display:
    network_type = "ruko"
elif "None" in network_type_display:
    network_type = "none"

st.sidebar.markdown("<div class='sidebar-divider'></div>", unsafe_allow_html=True)
st.sidebar.markdown("""
    ### 📁 Format Berkas
    * **Masukan:** KML Google Earth Pro
    * **Keluaran:** Excel BOQ (.xlsx)
    
    <div class='sidebar-divider'></div>
    
    ### ℹ️ Versi Aplikasi
    * **Versi:** v1.0 (Streamlit Cloud)
    
    <div class='sidebar-divider'></div>
    <p style='font-size: 11px; opacity: 0.6; text-align: center; margin-top: 50px;'>
        PT Telkom Indonesia &copy; 2026
    </p>
""", unsafe_allow_html=True)

# ----------------- MAIN HEADER BANNER -----------------
st.markdown("""
    <div class="header-container">
        <div style="font-size: 40px; line-height: 1;">🔴</div>
        <div>
            <h1 style="margin: 0; font-size: 26px;">FTTH BOQ GENERATOR</h1>
            <p style="margin: 0; font-size: 14px; color: #718096;">Aplikasi Otomatisasi Perhitungan Bill of Quantities (BoQ) PT Telkom Indonesia</p>
        </div>
    </div>
""", unsafe_allow_html=True)

# ----------------- MAIN CARD / FILE UPLOADER -----------------
uploaded_file = st.file_uploader(
    "Seret & letakkan file KML Anda di sini, atau klik untuk memilih file",
    type=["kml"]
)

if uploaded_file is not None:
    # Create temp file path
    unique_id = str(uuid.uuid4())
    temp_kml_path = UPLOAD_DIR / f"{unique_id}_{uploaded_file.name}"
    
    # Save the file temporarily
    with open(temp_kml_path, "wb") as f:
        f.write(uploaded_file.getbuffer())
        
    st.success(f"Berkas '{uploaded_file.name}' sukses diunggah. Siap diproses!")

    # Action Button
    if st.button("🚀 GENERATE EXCEL BOQ"):
        with st.spinner("Sedang membaca KML dan menghitung volume..."):
            try:
                # 1. Parse KML
                parser = KMLParser(str(temp_kml_path))
                parsed_data = parser.parse()
                raw_objects = parsed_data["raw_objects"]
                cable_spans = parsed_data["cable_spans"]

                # 2. Map objects using Excel Master
                template_file = BASE_DIR / "FORMAT BOQ PLAN.xlsx"
                if not template_file.exists():
                    raise FileNotFoundError("Berkas master template 'FORMAT BOQ PLAN.xlsx' tidak ditemukan.")
                
                mapper = BOQMapper(
                    str(MAPPING_FILE),
                    "Master",
                    network_type=network_type,
                    template_path=str(template_file)
                )
                boq_result = mapper.map_objects(raw_objects)

                # 3. Write data to Excel
                generator = ExcelGenerator(str(template_file))
                generator.fill_boq(boq_result)
                generator.update_boq_header(uploaded_file.name, "Master")
                generator.create_data_kml_sheet(raw_objects, cable_spans, uploaded_file.name)

                # 4. Save Output
                clean_name = uploaded_file.name.replace(".kml", "").replace(" ", "_")
                output_xlsx_path = OUTPUT_DIR / f"BOQ_{clean_name}.xlsx"
                generator.save(str(output_xlsx_path))

                st.balloons()
                st.success("Perhitungan BoQ selesai! Klik tombol di bawah untuk mengunduh berkas Excel.")

                # Serve download button
                with open(output_xlsx_path, "rb") as file:
                    st.download_button(
                        label="📥 Unduh Hasil Excel BoQ",
                        data=file,
                        file_name=f"BOQ_{clean_name}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )

            except Exception as e:
                st.error(f"Gagal memproses file KML: {str(e)}")
                
            finally:
                # Clean up uploaded temp file
                if temp_kml_path.exists():
                    try:
                        temp_kml_path.unlink()
                    except Exception:
                        pass

# ----------------- INSTRUCTIONS CARD -----------------
st.markdown("""
    <div class="info-card">
        <div class="info-title">
            <span>ℹ️</span> <strong>Petunjuk Penggunaan</strong>
        </div>
        <ol style="margin-left: 20px; font-size: 13px; color: #4A5568; line-height: 1.6;">
            <li>Ekspor folder pekerjaan di Google Earth Pro menjadi berkas berkstensi <code>.kml</code>.</li>
            <li>Pilih jenis konfigurasi jaringan yang sesuai pada menu dropdown di panel kiri (Perumahan, Ruko, atau Tanpa HH PIT).</li>
            <li>Unggah berkas KML dengan menyeretnya ke area unggahan di atas atau memilihnya secara manual.</li>
            <li>Klik tombol <strong>Generate Excel BOQ</strong> untuk memulai proses pemetaan dan kalkulasi.</li>
            <li>Setelah proses selesai, klik tombol <strong>Unduh Hasil Excel BoQ</strong> untuk mendownload berkas Excel Anda.</li>
        </ol>
    </div>
""", unsafe_allow_html=True)
