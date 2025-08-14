# ==============================================================================
# Aplikasi Streamlit untuk Mengisi Laporan Inspeksi dan Rincian Biaya Dinas
# ==============================================================================
import io
import math
import json
from datetime import date
from typing import Dict, Any
from urllib.parse import quote_plus

import streamlit as st
from PIL import Image
from docx import Document
from docx.shared import Inches
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from docx.table import Table
import pandas as pd
from openpyxl import load_workbook
import requests
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders

# --- MongoDB Imports ---
from pymongo.mongo_client import MongoClient
from pymongo.server_api import ServerApi

# MongoDB Configuration
# Perbaikan: Menggunakan urllib.parse.quote_plus untuk meng-escape karakter khusus
# pada username dan password sesuai RFC 3986.
# Contoh: password "Rahasia100%" menjadi "Rahasia100%25"
username = quote_plus("laporanglss")
password = quote_plus("Kmzway87aa")
MONGODB_URI = f"mongodb+srv://{username}:{password}@cluster0.fbp5d0n.mongodb.net/?retryWrites=true&w=majority&appName=Cluster0"
DATABASE_NAME = "laporan_dinas"
COLLECTION_NAME = "rbd_trips"

# Initialize MongoDB connection
@st.cache_resource
def init_mongodb():
    """Initialize MongoDB connection with caching"""
    try:
        client = MongoClient(MONGODB_URI, server_api=ServerApi('1'))
        # Test the connection
        client.admin.command('ping')
        st.success("✅ Successfully connected to MongoDB!")
        return client[DATABASE_NAME][COLLECTION_NAME]
    except Exception as e:
        st.error(f"❌ Failed to connect to MongoDB: {e}")
        return None

# Initialize database
db_collection = init_mongodb()

# ==============================================================================
# KONFIGURASI APLIKASI
# ==============================================================================

# Perbaikan: Menggunakan URL raw yang benar dari GitHub
TEMPLATE_INSPEKSI_URL = "https://raw.githubusercontent.com/FajarGLS/Laporan-Dinas/main/INSPEKSI.docx"
TEMPLATE_RBD_URL = "https://raw.githubusercontent.com/FajarGLS/Laporan-Dinas/main/RBD.xlsx"

# Konfigurasi SMTP email
EMAIL_SENDER = "fajar@dpagls.my.id"
EMAIL_PASSWORD = "Rahasia100%"
SMTP_SERVER = "mail.dpagls.my.id"
SMTP_PORT = 465

# Daftar kapal
VESSEL_LISTS = {
    "Bulk Carrier": [
        "MV NAZIHA", "MV AMMAR", "MV NAMEERA", "MV ARFIANIE AYU", "MV SAMI",
        "MV KAREEM", "MV NADHIF", "MV KAYSAN", "MV ABDUL HAMID", "MV NASHALINA",
        "MV GUNALEILA", "MV NUR AWLIYA", "MV NATASCHA", "MV SARAH S", "MV MARIA NASHWAH",
        "MV ZALEHA FITRAT", "MV HAMMADA", "MV KAMADIYA", "MV MUBASYIR", "MV MUHASYIR",
        "MV MUNQIDZ", "MV MUHARRIK", "MV MUMTAZ", "MV UNITAMA LILY", "MV RIMBA EMPAT",
        "MV MUADZ", "MV MUNIF", "MV RAFA", "MV NOUR MUSHTOFA", "MV FEIZA",
        "MV MURSYID", "MV AFKAR", "MV. AMOLONGO EMRAN", "MV. NIMAOME EMRAN",
        "MV. SYABIL EMRAN"
    ],
    "Tanker": [
        "MT BIO EXPRESS", "MT KENCANA EXPRESS", "MT SIL EXPRESS", "MT SELUMA EXPRESS"
    ],
    "Tug&Barge": [
        "TB. AZALEA", "TB. GRENADA", "TB. MAGNOLIA", "TB. ZINNIA", "TB. JAZZY",
        "TB. AMARILLYS", "TB. FENNEL", "TB. LAUREL", "TB. JASMINE", "TB. CAMELIA",
        "TB. MULBERRY", "TB. FUSCHIA", "TB. IRIS", "TB. JUNIPER", "TB. SPEEDWELL",
        "TB. MIMOSA", "TB. IVY", "TB. SORREL", "TB. GERBERA", "TB. CLEMATIS",
        "TB. EUSTOMA", "TB. FEIHA", "TB. EHSAL", "TB. GMS CEMERLANG 1", "TB. GMS CEMERLANG 2",
        "TB. ALYSUM", "TB. CATALEYA", "TB. GMS CEMERLANG 3"
    ]
}

# Inisialisasi session state
if "dok_rows" not in st.session_state:
    st.session_state.dok_rows = 10
if "report_type" not in st.session_state:
    st.session_state.report_type = "Laporan Inspeksi"

# ==============================================================================
# HELPERS UNTUK MEMANIPULASI DOKUMEN WORD (.docx)
# ==============================================================================

def _replace_in_paragraph(paragraph, placeholder, value):
    if placeholder not in paragraph.text:
        return
    full_text = "".join(run.text for run in paragraph.runs)
    new_text = full_text.replace(placeholder, value)
    for _ in range(len(paragraph.runs)):
        paragraph.runs[0].clear()
        paragraph.runs[0].text = ""
    if len(paragraph.runs) == 0:
        paragraph.add_run(new_text)
    else:
        paragraph.runs[0].text = new_text

def replace_placeholder_everywhere(doc: Document, placeholder: str, value: str):
    for p in doc.paragraphs:
        _replace_in_paragraph(p, placeholder, value)
    for tbl in doc.tables:
        for row in tbl.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    _replace_in_paragraph(p, placeholder, value)

def find_cell_with_text(doc: Document, placeholder: str):
    for tbl in doc.tables:
        for r_idx, row in enumerate(tbl.rows):
            for c_idx, cell in enumerate(row.cells):
                for p in cell.paragraphs:
                    if placeholder in p.text:
                        return cell, tbl, r_idx, c_idx
    return None, None, None, None

def _get_table_grid_col_widths_in_inches(tbl: Table):
    try:
        grid = tbl._tbl.tblGrid
        if grid is None:
            return None
        cols = []
        for gcol in grid.gridCol_lst:
            twips = int(gcol.w)
            cols.append(twips / 1440.0)
        return cols
    except Exception:
        return None

def _get_page_usable_width_inches(doc_or_body) -> float:
    try:
        section = doc_or_body.sections[0]
        page_width_in = section.page_width.inches
        left_in = section.left_margin.inches
        right_in = section.right_margin.inches
        return max(0.1, page_width_in - left_in - right_in)
    except AttributeError:
        return 6.5

def _estimate_cell_width_inches(cell, tbl: Table):
    grid_cols = _get_table_grid_col_widths_in_inches(tbl)
    if grid_cols:
        for r in tbl.rows:
            if cell in r.cells:
                col_index = r.cells.index(cell)
                return max(0.1, grid_cols[col_index] - 0.05)
    usable = _get_page_usable_width_inches(tbl._parent)
    ncols = len(tbl.rows[0].cells) if tbl.rows and tbl.rows[0].cells else 2
    return max(0.1, (usable / ncols) - 0.05)

def insert_image_into_cell(cell, tbl: Table, image_bytes: bytes, sizing_mode='adaptive'):
    """Memasukkan gambar ke dalam sel dengan mode ukuran yang berbeda."""
    if not image_bytes:
        return
    try:
        with Image.open(io.BytesIO(image_bytes)) as img_check:
            img_check.verify()
    except Exception:
        cell.text = "[Invalid image]"
        return

    cell.text = ""
    par = cell.paragraphs[0] if cell.paragraphs else cell.add_paragraph()
    par.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER

    if sizing_mode == 'fixed':
        # Ukuran tetap untuk gambar DOKUMENTASI agar seragam
        width_in = Inches(3.0)
        height_in = Inches(2.25)
        run = par.add_run()
        run.add_picture(io.BytesIO(image_bytes), width=width_in, height=height_in)
    else:  # Mode 'adaptive' untuk FOTOHALUAN
        cell_width_in = _estimate_cell_width_inches(cell, tbl)
        run = par.add_run()
        run.add_picture(io.BytesIO(image_bytes), width=Inches(cell_width_in))

def find_paragraph_with_text(doc: Document, placeholder: str):
    for p in doc.paragraphs:
        if placeholder in p.text:
            return p
    for tbl in doc.tables:
        for row in tbl.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    if placeholder in p.text:
                        return p
    return None

def add_table_after_paragraph(doc: Document, paragraph, rows: int, cols: int) -> Table:
    temp_table = doc.add_table(rows=rows, cols=cols)
    tbl_element = temp_table._tbl
    paragraph._p.addnext(tbl_element)
    return Table(tbl_element, paragraph._parent)

def center_all_cells(tbl: Table):
    for row in tbl.rows:
        for cell in row.cells:
            for p in cell.paragraphs:
                p.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER

def set_equal_column_widths(tbl: Table, total_width_in: float):
    if len(tbl.columns) != 2:
        return
    col_w = total_width_in / 2.0
    for col in tbl.columns:
        for cell in col.cells:
            cell.width = Inches(col_w)

def build_dokumentasi_table_at_placeholder(doc: Document, placeholder: str, items):
    p = find_paragraph_with_text(doc, placeholder)
    if not p:
        p = doc.add_paragraph("")
    _replace_in_paragraph(p, placeholder, "")

    cleaned = []
    for it in items:
        if it.get("image_bytes") or (it.get("caption", "").strip()):
            cleaned.append(it)
    items = cleaned if cleaned else items

    n_items = len(items)
    grid_rows = math.ceil(n_items / 2) if n_items > 0 else 1
    doc_rows = grid_rows * 2
    tbl = add_table_after_paragraph(doc, p, rows=doc_rows, cols=2)
    tbl.autofit = True
    usable = _get_page_usable_width_inches(doc)
    set_equal_column_widths(tbl, usable)
    center_all_cells(tbl)

    idx = 0
    for r in range(grid_rows):
        image_row = tbl.rows[r * 2]
        caption_row = tbl.rows[r * 2 + 1]
        for c in range(2):
            if idx < n_items:
                item = items[idx]
                if item.get("image_bytes"):
                    insert_image_into_cell(image_row.cells[c], tbl, item["image_bytes"], sizing_mode='fixed')
                else:
                    image_row.cells[c].text = ""
                cap_text = (item.get("caption") or "").strip()
                caption_row.cells[c].text = cap_text
                for pcap in caption_row.cells[c].paragraphs:
                    pcap.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
            else:
                image_row.cells[c].text = ""
                caption_row.cells[c].text = ""
            idx += 1

# ==============================================================================
# HELPERS UNTUK MENGIRIM EMAIL
# ==============================================================================

def send_email_with_attachment(
    from_email, password, to_email, smtp_server, smtp_port, subject, body, attachments
):
    """Mengirim email dengan banyak lampiran."""
    try:
        msg = MIMEMultipart()
        msg["From"] = from_email
        msg["To"] = to_email
        msg["Subject"] = subject
        msg.attach(MIMEBase("text", "plain", payload=body))

        for attachment_bytes, attachment_filename in attachments:
            part = MIMEBase("application", "octet-stream")
            part.set_payload(attachment_bytes)
            encoders.encode_base64(part)
            part.add_header(
                "Content-Disposition",
                f"attachment; filename= {attachment_filename}",
            )
            msg.attach(part)

        server = smtplib.SMTP_SSL(smtp_server, int(smtp_port))
        server.login(from_email, password)
        text = msg.as_string()
        server.sendmail(from_email, to_email, text)
        server.quit()
        return True
    except Exception as e:
        st.error(f"Gagal mengirim email: {e}")
        return False

# ==============================================================================
# MONGODB FUNCTIONS
# ==============================================================================

def save_to_mongodb(data: Dict[str, Any]):
    """Menyimpan data ke MongoDB"""
    if not db_collection:
        st.error("❌ Koneksi MongoDB tidak tersedia")
        return False
    
    if not data.get("trip_id"):
        st.error("❌ Trip ID tidak boleh kosong untuk menyimpan data.")
        return False
    
    try:
        # Menggunakan upsert untuk update jika sudah ada, insert jika belum
        result = db_collection.replace_one(
            {"trip_id": data["trip_id"]}, 
            data, 
            upsert=True
        )
        
        if result.upserted_id:
            st.success(f"✅ Data perjalanan dinas '{data['trip_id']}' berhasil disimpan!")
        else:
            st.success(f"✅ Data perjalanan dinas '{data['trip_id']}' berhasil diperbarui!")
        return True
    except Exception as e:
        st.error(f"❌ Gagal menyimpan ke MongoDB: {e}")
        return False

def load_from_mongodb(trip_id: str) -> Dict[str, Any]:
    """Memuat data dari MongoDB"""
    if not db_collection:
        st.error("❌ Koneksi MongoDB tidak tersedia")
        return {}
    
    try:
        document = db_collection.find_one({"trip_id": trip_id})
        if document:
            # Hapus _id dari hasil karena tidak diperlukan
            document.pop('_id', None)
            st.success(f"✅ Data perjalanan dinas '{trip_id}' berhasil dimuat!")
            return document
        else:
            st.warning(f"⚠️ Data perjalanan dinas dengan ID '{trip_id}' tidak ditemukan.")
            return {}
    except Exception as e:
        st.error(f"❌ Gagal memuat dari MongoDB: {e}")
        return {}

def get_all_trip_ids() -> list:
    """Mendapatkan semua trip ID yang tersimpan"""
    if not db_collection:
        return []
    
    try:
        trip_ids = [doc["trip_id"] for doc in db_collection.find({}, {"trip_id": 1, "_id": 0})]
        return sorted(trip_ids)
    except Exception as e:
        st.error(f"❌ Gagal mengambil daftar trip ID: {e}")
        return []

# ==============================================================================
# CALLBACK FUNCTIONS
# ==============================================================================

def add_dok_rows():
    """Callback function untuk menambah rows dokumentasi"""
    st.session_state.dok_rows += 2

# ==============================================================================
# UTAMA APLIKASI STREAMLIT
# ==============================================================================

st.set_page_config(page_title="Multi-Report App", layout="wide")
st.title("Aplikasi Pembuat Laporan & RBD")

# --- PILIH JENIS LAPORAN DI SIDEBAR ---
with st.sidebar:
    st.header("Pilih Jenis Laporan")
    report_type = st.radio(
        "Pilih laporan yang akan dibuat:",
        ("Laporan Inspeksi", "Rincian Biaya Perjalanan Dinas")
    )
    st.session_state.report_type = report_type
    
    if st.session_state.report_type == "Laporan Inspeksi":
        if "dok_rows" not in st.session_state:
            st.session_state.dok_rows = 10
        st.button("➕ Tambah Row Dokumentasi", on_click=add_dok_rows)
    
    # Tampilkan status koneksi MongoDB
    st.markdown("---")
    st.subheader("Status Database")
    if db_collection:
        st.success("🟢 MongoDB Connected")
    else:
        st.error("🔴 MongoDB Disconnected")

# ==============================================================================
# BAGIAN 1: LAPORAN INSPEKSI
# ==============================================================================
if st.session_state.report_type == "Laporan Inspeksi":

    st.header("Laporan Inspeksi Kapal")

    st.markdown("Aplikasi ini akan secara otomatis mengambil template dari GitHub.")
    
    # Mengambil template dari URL
    try:
        with st.spinner('Mengambil template dari GitHub...'):
            response = requests.get(TEMPLATE_INSPEKSI_URL)
            response.raise_for_status()
        template_file = io.BytesIO(response.content)
    except requests.exceptions.RequestException as e:
        st.error(f"Gagal mengambil template dari GitHub: {e}. Pastikan URL-nya benar.")
        st.stop()

    # --- UI Input untuk Laporan Inspeksi ---
    st.subheader("FOTOHALUAN")
    foto_haluan_file = st.file_uploader(
        "Upload FOTOHALUAN image", type=["jpg", "jpeg", "png"], key="foto_haluan"
    )

    if foto_haluan_file is not None:
        st.session_state["foto_haluan_bytes"] = foto_haluan_file.getvalue()
    if "foto_haluan_bytes" in st.session_state:
        st.subheader("Preview FOTOHALUAN")
        st.image(st.session_state["foto_haluan_bytes"], use_container_width=True)

    st.markdown("---")

    st.subheader("Vessel Details")
    col1, col2 = st.columns(2, gap="large")
    with col1:
        ship_type = st.selectbox("Type (*TYPE*)", options=list(VESSEL_LISTS.keys()))
        vessel_options = VESSEL_LISTS.get(ship_type, [])
        vessel_name = st.selectbox("Vessel (*VESSEL*)", options=vessel_options)
        imo = st.text_input("IMO (*IMO*)")
    with col2:
        callsign = st.text_input("Callsign (*CALLSIGN*)")
        place = st.text_input("Place (*PLACEDATE*)", placeholder="e.g., Jakarta")

    survey_date = st.date_input("Date (*PLACEDATE*)", value=date.today())
    master = st.text_input("Master (*MASTER*)")
    surveyor = st.text_input("Surveyor (*SURVEYOR*)")

    st.markdown("---")

    st.subheader("DOKUMENTASI")
    dok_items = []
    row_pairs = [(i, i + 1) for i in range(0, st.session_state.dok_rows, 2)]

    def render_preview_50(file_bytes):
        try:
            with Image.open(io.BytesIO(file_bytes)) as img:
                img_resized = img.resize((400, 300))
            st.image(img_resized, use_container_width=False)
        except Exception:
            st.warning("Format gambar tidak valid.")

    for left_idx, right_idx in row_pairs:
        col_left, col_right = st.columns(2, gap="large")
        with col_left:
            img_key_left = f"dok_img_{left_idx}_0"
            cap_key_left = f"dok_cap_{left_idx}_0"
            file_left = st.file_uploader(
                f"Row {left_idx + 1} - Left Image",
                type=["jpg", "jpeg", "png"],
                key=img_key_left
            )
            if file_left is not None:
                st.session_state[img_key_left + "_bytes"] = file_left.getvalue()
            if st.session_state.get(img_key_left + "_bytes"):
                render_preview_50(st.session_state[img_key_left + "_bytes"])
            caption_left = st.text_input(f"Row {left_idx + 1} - Caption", key=cap_key_left)
            dok_items.append({
                "image_bytes": st.session_state.get(img_key_left + "_bytes", None),
                "caption": caption_left or ""
            })
        with col_right:
            if right_idx < st.session_state.dok_rows:
                img_key_right = f"dok_img_{right_idx}_1"
                cap_key_right = f"dok_cap_{right_idx}_1"
                file_right = st.file_uploader(
                    f"Row {right_idx + 1} - Right Image",
                    type=["jpg", "jpeg", "png"],
                    key=img_key_right
                )
                if file_right is not None:
                    st.session_state[img_key_right + "_bytes"] = file_right.getvalue()
                if st.session_state.get(img_key_right + "_bytes"):
                    render_preview_50(st.session_state[img_key_right + "_bytes"])
                caption_right = st.text_input(f"Row {right_idx + 1} - Caption", key=cap_key_right)
                dok_items.append({
                    "image_bytes": st.session_state.get(img_key_right + "_bytes", None),
                    "caption": caption_right or ""
                })

    # --- Tombol Generate dan Email ---
    st.markdown("---")
    st.subheader("Email Penerima")
    st.info("Masukkan email penerima di bawah ini. Laporan akan langsung dikirim setelah di-generate.")
    email_to_send = st.text_input("Email Penerima")

    if st.button("📝 Generate & Send Report"):
        if not email_to_send:
            st.error("Silakan masukkan alamat email penerima.")
            st.stop()
        
        try:
            doc = Document(template_file)
        except Exception:
            st.error("Template .docx tidak valid atau rusak.")
            st.stop()

        replace_placeholder_everywhere(doc, "*VESSEL*", vessel_name)
        replace_placeholder_everywhere(doc, "*IMO*", imo)
        replace_placeholder_everywhere(doc, "*TYPE*", ship_type)
        replace_placeholder_everywhere(doc, "*CALLSIGN*", callsign)
        replace_placeholder_everywhere(doc, "*PLACEDATE*", f"{place}, {survey_date.strftime('%d %B %Y')}")
        replace_placeholder_everywhere(doc, "*MASTER*", master)
        replace_placeholder_everywhere(doc, "*SURVEYOR*", surveyor)

        foto_bytes = st.session_state.get("foto_haluan_bytes")
        if foto_bytes:
            cell, tbl, _, _ = find_cell_with_text(doc, "*FOTOHALUAN*")
            if cell:
                insert_image_into_cell(cell, tbl, foto_bytes, sizing_mode='adaptive')
                replace_placeholder_everywhere(doc, "*FOTOHALUAN*", "")
            else:
                replace_placeholder_everywhere(doc, "*FOTOHALUAN*", "")

        build_dokumentasi_table_at_placeholder(doc, "*DOKUMENTASI*", dok_items)

        docx_buffer = io.BytesIO()
        doc.save(docx_buffer)
        docx_buffer.seek(0)
        
        base_filename = f"{survey_date.strftime('%Y.%m.%d')} {vessel_name} Inspection Report"
        docx_filename = f"{base_filename}.docx"

        st.write("Mengirim laporan melalui email...")
        attachments_list = [(docx_buffer.getvalue(), docx_filename)]
        success = send_email_with_attachment(
            EMAIL_SENDER, EMAIL_PASSWORD, email_to_send, SMTP_SERVER, SMTP_PORT,
            f"Laporan Inspeksi: {vessel_name}",
            f"Terlampir laporan inspeksi kapal {vessel_name} dalam format DOCX.",
            attachments_list
        )
        if success:
            st.success(f"Laporan berhasil dikirim ke {email_to_send}!")
        
        st.download_button(
            label="📄 Download Laporan (DOCX)",
            data=docx_buffer,
            file_name=docx_filename,
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )

# ==============================================================================
# BAGIAN 2: RINCIAN BIAYA PERJALANAN DINAS
# ==============================================================================
if st.session_state.report_type == "Rincian Biaya Perjalanan Dinas":
    st.header("Rincian Biaya Perjalanan Dinas (RBD)")

    # --- UI Input untuk RBD ---
    st.subheader("Detail Perjalanan Dinas")
    
    # Dropdown untuk memilih trip ID yang sudah ada
    col_trip1, col_trip2 = st.columns([2, 1])
    with col_trip1:
        trip_id = st.text_input("ID Perjalanan Dinas (untuk menyimpan)", help="Contoh: 'FAJAR-JAKARTA-2024-08-15'")
    with col_trip2:
        all_trip_ids = get_all_trip_ids()
        if all_trip_ids:
            selected_existing_trip = st.selectbox("Atau pilih yang sudah ada:", [""] + all_trip_ids)
            if selected_existing_trip:
                trip_id = selected_existing_trip
                st.rerun()
    
    col1, col2 = st.columns(2)
    with col1:
        start_date = st.date_input("Tanggal Mulai", value=date.today())
    with col2:
        end_date = st.date_input("Tanggal Selesai", value=date.today())
    
    trip_purpose = st.text_input("Tujuan Dinas", help="Akan masuk ke sel C13")
    vessel_code = st.text_input("Kapal Tujuan Dinas / Vessel Code", help="Akan masuk ke sel F13")

    st.subheader("Detail Biaya (Isi nilai 'Pemakaian' atau 'Rp' di sini)")
    hotel_cost = st.text_input("Akomodasi Hotel", value="0", help="Akan masuk ke sel N20")
    deposit = st.text_input("Deposit Hotel", value="0", help="Akan masuk ke sel N22")
    plane_cost = st.text_input("Pesawat", value="0", help="Akan masuk ke sel N24")
    miscellaneous = st.text_input("Miscellaneous Document Cargo", value="0", help="Akan masuk ke sel N26")
    airport_tax = st.text_input("Airport Tax", value="0", help="Akan masuk ke sel N28")
    ship_cost = st.text_input("Kapal Laut", value="0", help="Akan masuk ke sel N30")
    train_cost = st.text_input("Kereta Api", value="0", help="Akan masuk ke sel N33")
    bus_cost = st.text_input("Bis", value="0", help="Akan masuk ke sel N36")
    fuel_cost = st.text_input("Kendaraan Dinas (BBM)", value="0", help="Akan masuk ke sel N39")
    toll_cost = st.text_input("Nota Toll", value="0", help="Akan masuk ke sel N40")
    taxi_cost = st.text_input("Taksi / Bis", value="0", help="Akan masuk ke sel N42")
    local_transport = st.text_input("Transportasi di tempat dinas", value="0", help="Akan masuk ke sel N46")
    boat_jetty = st.text_input("Boat Jetty", value="0", help="Akan masuk ke sel N47")
    weekend_transport = st.text_input("Uang Transport di tanggal Merah", value="0", help="Akan masuk ke sel N52")

    # Tombol untuk menyimpan, memuat, dan generate
    col_buttons = st.columns(3)
    with col_buttons[0]:
        if st.button("💾 Simpan Data"):
            trip_data = {
                "trip_id": trip_id,
                "start_date": start_date.strftime("%Y-%m-%d"),
                "end_date": end_date.strftime("%Y-%m-%d"),
                "trip_purpose": trip_purpose,
                "vessel_code": vessel_code,
                "hotel_cost": hotel_cost,
                "deposit": deposit,
                "plane_cost": plane_cost,
                "miscellaneous": miscellaneous,
                "airport_tax": airport_tax,
                "ship_cost": ship_cost,
                "train_cost": train_cost,
                "bus_cost": bus_cost,
                "fuel_cost": fuel_cost,
                "toll_cost": toll_cost,
                "taxi_cost": taxi_cost,
                "local_transport": local_transport,
                "boat_jetty": boat_jetty,
                "weekend_transport": weekend_transport
            }

            if save_to_mongodb(trip_data):
                st.experimental_rerun()
    with col_buttons[1]:
        if st.button("📂 Muat Data"):
            if trip_id:
                loaded_data = load_from_mongodb(trip_id)
                if loaded_data:
                    st.session_state.update({
                        "trip_id": loaded_data.get("trip_id", ""),
                        "start_date": pd.to_datetime(loaded_data.get("start_date")).date(),
                        "end_date": pd.to_datetime(loaded_data.get("end_date")).date(),
                        "trip_purpose": loaded_data.get("trip_purpose", ""),
                        "vessel_code": loaded_data.get("vessel_code", ""),
                        "hotel_cost": loaded_data.get("hotel_cost", "0"),
                        "deposit": loaded_data.get("deposit", "0"),
                        "plane_cost": loaded_data.get("plane_cost", "0"),
                        "miscellaneous": loaded_data.get("miscellaneous", "0"),
                        "airport_tax": loaded_data.get("airport_tax", "0"),
                        "ship_cost": loaded_data.get("ship_cost", "0"),
                        "train_cost": loaded_data.get("train_cost", "0"),
                        "bus_cost": loaded_data.get("bus_cost", "0"),
                        "fuel_cost": loaded_data.get("fuel_cost", "0"),
                        "toll_cost": loaded_data.get("toll_cost", "0"),
                        "taxi_cost": loaded_data.get("taxi_cost", "0"),
                        "local_transport": loaded_data.get("local_transport", "0"),
                        "boat_jetty": loaded_data.get("boat_jetty", "0"),
                        "weekend_transport": loaded_data.get("weekend_transport", "0")
                    })
                    st.experimental_rerun()
            else:
                st.warning("⚠️ Silakan masukkan Trip ID untuk memuat data.")
    with col_buttons[2]:
        if st.button("📝 Generate Laporan & Send Email"):
            if not trip_id:
                st.error("Silakan masukkan Trip ID untuk melanjutkan.")
                st.stop()
            
            # Mendapatkan data terakhir yang diisi
            trip_data = {
                "trip_id": trip_id,
                "start_date": start_date,
                "end_date": end_date,
                "trip_purpose": trip_purpose,
                "vessel_code": vessel_code,
                "hotel_cost": hotel_cost,
                "deposit": deposit,
                "plane_cost": plane_cost,
                "miscellaneous": miscellaneous,
                "airport_tax": airport_tax,
                "ship_cost": ship_cost,
                "train_cost": train_cost,
                "bus_cost": bus_cost,
                "fuel_cost": fuel_cost,
                "toll_cost": toll_cost,
                "taxi_cost": taxi_cost,
                "local_transport": local_transport,
                "boat_jetty": boat_jetty,
                "weekend_transport": weekend_transport
            }

            st.write("Mengambil template dari GitHub...")
            try:
                response = requests.get(TEMPLATE_RBD_URL)
                response.raise_for_status()
                template_file = io.BytesIO(response.content)
            except requests.exceptions.RequestException as e:
                st.error(f"Gagal mengambil template RBD dari GitHub: {e}. Pastikan URL-nya benar.")
                st.stop()

            st.write("Membuat laporan RBD...")
            try:
                rbd_wb = load_workbook(template_file)
                ws = rbd_wb.active

                # Mengisi data ke spreadsheet
                ws['C13'] = trip_data["trip_purpose"]
                ws['F13'] = trip_data["vessel_code"]
                ws['K13'] = "DPA"
                ws['N20'] = float(trip_data["hotel_cost"])
                ws['N22'] = float(trip_data["deposit"])
                ws['N24'] = float(trip_data["plane_cost"])
                ws['N26'] = float(trip_data["miscellaneous"])
                ws['N28'] = float(trip_data["airport_tax"])
                ws['N30'] = float(trip_data["ship_cost"])
                ws['N33'] = float(trip_data["train_cost"])
                ws['N36'] = float(trip_data["bus_cost"])
                ws['N39'] = float(trip_data["fuel_cost"])
                ws['N40'] = float(trip_data["toll_cost"])
                ws['N42'] = float(trip_data["taxi_cost"])
                ws['N46'] = float(trip_data["local_transport"])
                ws['N47'] = float(trip_data["boat_jetty"])
                ws['N52'] = float(trip_data["weekend_transport"])
                
                # Menghitung durasi
                duration = (trip_data["end_date"] - trip_data["start_date"]).days + 1
                ws['I15'] = duration
                ws['I16'] = trip_data["start_date"].strftime("%d %B %Y")
                ws['I17'] = trip_data["end_date"].strftime("%d %B %Y")
                
                # Mengisi tanggal
                ws['L60'] = trip_data["end_date"]
                
                output_buffer = io.BytesIO()
                rbd_wb.save(output_buffer)
                output_buffer.seek(0)
                
                # Konfigurasi email
                st.subheader("Email Penerima")
                email_rbd = st.text_input("Masukkan email penerima untuk RBD")
                
                if email_rbd:
                    rbd_filename = f"RBD {trip_data['trip_id']}.xlsx"
                    attachments_list = [(output_buffer.getvalue(), rbd_filename)]
                    
                    st.write("Mengirim laporan melalui email...")
                    success = send_email_with_attachment(
                        EMAIL_SENDER, EMAIL_PASSWORD, email_rbd, SMTP_SERVER, SMTP_PORT,
                        f"Rincian Biaya Dinas: {trip_data['trip_id']}",
                        f"Terlampir laporan rincian biaya perjalanan dinas dengan ID {trip_data['trip_id']} dalam format XLSX.",
                        attachments_list
                    )
                    if success:
                        st.success(f"Laporan RBD berhasil dikirim ke {email_rbd}!")
                        
                    st.download_button(
                        label="📄 Download Laporan (XLSX)",
                        data=output_buffer,
                        file_name=rbd_filename,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                else:
                    st.warning("⚠️ Masukkan alamat email untuk mengirim RBD.")
            
            except Exception as e:
                st.error(f"Gagal membuat laporan RBD: {e}")
