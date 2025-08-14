# ==============================================================================
# Aplikasi Streamlit untuk Mengisi Laporan Inspeksi dan Rincian Biaya Dinas
# ==============================================================================
import io
import math
import json
from datetime import date
from typing import Dict, Any

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

# --- Firebase Imports ---
from firebase_admin import credentials, firestore, initialize_app
# Catatan: Kredensial Firebase harus disediakan di lingkungan deployment
# sebagai string JSON, misalnya melalui Streamlit secrets.
try:
    if not st.secrets.get("firebase_service_account"):
        st.error("Firebase credentials not found. Please add a firebase_service_account to your Streamlit secrets.")
        st.stop()
    if not initialize_app():
        cred = credentials.Certificate(st.secrets["firebase_service_account"])
        initialize_app(cred)
    db = firestore.client()
except Exception as e:
    st.error(f"Failed to initialize Firebase: {e}")
    st.info("Pastikan Anda telah menambahkan kredensial Firebase ke Streamlit secrets.")
    db = None

# ==============================================================================
# KONFIGURASI APLIKASI
# ==============================================================================

# Konfigurasi URL template .docx
TEMPLATE_INSPEKSI_URL = "https://github.com/FajarDPA/Laporan-Inspeksi/raw/main/INSPEKSI.docx"
# Konfigurasi URL template .xlsx
TEMPLATE_RBD_URL = "https://github.com/FajarDPA/Laporan-Inspeksi/raw/main/RBD.xlsx"

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
            for cell in cell.cells:
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
        st.button("➕ Tambah Row Dokumentasi", on_click=lambda: st.session_state.dok_rows += 2 and st.rerun())


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
    trip_id = st.text_input("ID Perjalanan Dinas (untuk menyimpan)", help="Contoh: 'FAJAR-JAKARTA-2024-08-15'")
    
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
    
    # --- Fungsi untuk Menyimpan Data ke Firestore ---
    def save_to_firestore(data: Dict[str, Any]):
        if not db: return
        if not data.get("trip_id"):
            st.error("Trip ID tidak boleh kosong untuk menyimpan data.")
            return
        doc_ref = db.collection("rbd_trips").document(data["trip_id"])
        doc_ref.set(data)
        st.success(f"Data perjalanan dinas '{data['trip_id']}' berhasil disimpan!")

    def load_from_firestore(trip_id: str) -> Dict[str, Any]:
        if not db: return {}
        doc_ref = db.collection("rbd_trips").document(trip_id)
        doc = doc_ref.get()
        if doc.exists:
            st.success(f"Data perjalanan dinas '{trip_id}' berhasil dimuat!")
            return doc.to_dict()
        else:
            st.warning(f"Data perjalanan dinas dengan ID '{trip_id}' tidak ditemukan.")
            return {}

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
            save_to_firestore(trip_data)

    with col_buttons[1]:
        if st.button("📥 Muat Data"):
            loaded_data = load_from_firestore(trip_id)
            if loaded_data:
                st.session_state["loaded_data"] = loaded_data

    if "loaded_data" in st.session_state and st.session_state.loaded_data:
        st.write("Menggunakan data yang dimuat. Anda bisa mengubahnya sebelum membuat laporan.")
        st.session_state.trip_id = st.session_state.loaded_data.get("trip_id")
        st.session_state.start_date = date.fromisoformat(st.session_state.loaded_data.get("start_date"))
        st.session_state.end_date = date.fromisoformat(st.session_state.loaded_data.get("end_date"))
        st.session_state.trip_purpose = st.session_state.loaded_data.get("trip_purpose")
        st.session_state.vessel_code = st.session_state.loaded_data.get("vessel_code")
        st.session_state.hotel_cost = st.session_state.loaded_data.get("hotel_cost")
        st.session_state.deposit = st.session_state.loaded_data.get("deposit")
        st.session_state.plane_cost = st.session_state.loaded_data.get("plane_cost")
        st.session_state.miscellaneous = st.session_state.loaded_data.get("miscellaneous")
        st.session_state.airport_tax = st.session_state.loaded_data.get("airport_tax")
        st.session_state.ship_cost = st.session_state.loaded_data.get("ship_cost")
        st.session_state.train_cost = st.session_state.loaded_data.get("train_cost")
        st.session_state.bus_cost = st.session_state.loaded_data.get("bus_cost")
        st.session_state.fuel_cost = st.session_state.loaded_data.get("fuel_cost")
        st.session_state.toll_cost = st.session_state.loaded_data.get("toll_cost")
        st.session_state.taxi_cost = st.session_state.loaded_data.get("taxi_cost")
        st.session_state.local_transport = st.session_state.loaded_data.get("local_transport")
        st.session_state.boat_jetty = st.session_state.loaded_data.get("boat_jetty")
        st.session_state.weekend_transport = st.session_state.loaded_data.get("weekend_transport")

    # --- Tombol Generate dan Download RBD ---
    if st.button("📝 Generate & Download RBD"):
        try:
            with st.spinner('Mengambil template RBD dari GitHub...'):
                response = requests.get(TEMPLATE_RBD_URL)
                response.raise_for_status()
            
            # Load the Excel file from the byte stream
            rbd_buffer = io.BytesIO(response.content)
            wb = load_workbook(rbd_buffer)
            ws = wb.active
            
            # Mengisi data ke sel-sel yang ditentukan
            ws['D11'] = start_date.strftime("%d-%m-%Y")
            ws['H11'] = end_date.strftime("%d-%m-%Y")
            ws['C13'] = trip_purpose
            ws['F13'] = vessel_code
            ws['N20'] = hotel_cost
            ws['N22'] = deposit
            ws['N24'] = plane_cost
            ws['N26'] = miscellaneous
            ws['N28'] = airport_tax
            ws['N30'] = ship_cost
            ws['N33'] = train_cost
            ws['N36'] = bus_cost
            ws['N39'] = fuel_cost
            ws['N40'] = toll_cost
            ws['N42'] = taxi_cost
            ws['N46'] = local_transport
            ws['N47'] = boat_jetty
            ws['N52'] = weekend_transport
            ws['A58'] = "Jakarta," + date.today().strftime("%d %B %Y")
            
            # Menyimpan kembali ke buffer
            output_rbd = io.BytesIO()
            wb.save(output_rbd)
            output_rbd.seek(0)
            
            rbd_filename = f"RBD_{vessel_code}_{start_date.strftime('%Y.%m.%d')}.xlsx"
            
            st.success("Laporan RBD berhasil dibuat!")
            st.download_button(
                label="📄 Download Rincian Biaya Dinas (XLSX)",
                data=output_rbd,
                file_name=rbd_filename,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            
        except requests.exceptions.RequestException as e:
            st.error(f"Gagal mengambil template RBD dari GitHub: {e}. Pastikan URL-nya benar.")
            st.stop()
        except Exception as e:
            st.error(f"Terjadi kesalahan saat memproses file Excel: {e}")
            st.stop()
