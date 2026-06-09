"""
Daily Report Generator — PT Nale System Integrator
Versi dengan:
- Upload foto tanpa batasan
- DOCX: otomatis menambah halaman & tabel foto sesuai jumlah foto (per 8 foto)
- PDF: 4 foto per halaman (grid 2x2)
- Persist hanya PIC & Team Support
"""

import streamlit as st
from docx import Document
from docx.shared import Cm, Pt
from docx.oxml.ns import qn
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from PIL import Image, ImageOps
from datetime import date
from pathlib import Path
from copy import deepcopy
import io
import json
import re

# Untuk PDF
from fpdf import FPDF
import tempfile
import os

# ═══════════════════════════════════════════════════════════════
# CONFIG
# ═══════════════════════════════════════════════════════════════
APP_DIR = Path(__file__).parent
TEMPLATE_PATH = APP_DIR / "template.docx"
PERSIST_FILE = APP_DIR / ".user_data.json"
FULL_DATA_FILE = APP_DIR / ".full_report_data.json"

BULAN_ID = [
    "Januari", "Februari", "Maret", "April", "Mei", "Juni",
    "Juli", "Agustus", "September", "Oktober", "November", "Desember"
]

st.set_page_config(
    page_title="Daily Report Generator — Nale SI",
    page_icon="📋",
    layout="wide",
    initial_sidebar_state="collapsed",
)

# ═══════════════════════════════════════════════════════════════
# STYLING (dark theme + file uploader fix)
# ═══════════════════════════════════════════════════════════════
st.markdown("""
<style>
    .stApp { background: #0e1233; }
    .main .block-container { padding-top: 1.5rem; max-width: 1100px; }
    h1, h2, h3, h4, p, label, span, div { color: #e2e8f0 !important; }
    .stTextInput input, .stTextArea textarea, .stDateInput input, .stSelectbox div[data-baseweb="select"] > div {
        background: #181e4a !important; color: #e2e8f0 !important; border: 1px solid rgba(255,255,255,0.07) !important;
    }
    .stTextInput input::placeholder, .stTextArea textarea::placeholder {
        color: #7a82a6 !important; font-style: italic; font-size: 0.85rem;
    }
    .stFileUploader {
        background: #1a1f4e;
        border: 2px dashed #f5a623;
        border-radius: 12px;
        padding: 10px;
    }
    .stFileUploader > div:first-child,
    .stFileUploader label,
    .stFileUploader span,
    .stFileUploader p {
        color: #f5a623 !important;
        font-weight: 500;
    }
    .stFileUploader button {
        background: #f5a623 !important;
        color: #141942 !important;
        border: none !important;
        font-weight: bold;
        border-radius: 6px;
    }
    .stFileUploader button:hover {
        background: #ffb940 !important;
    }
    .stFileUploader [data-testid="stFileUploaderDropzone"] {
        background: #0e1233;
        border-color: #f5a623;
    }
    .stFileUploader [data-testid="stFileUploaderDropzone"] div {
        color: #e2e8f0;
    }
    .stButton > button {
        background: linear-gradient(135deg, #f5a623, #e09000); color: #141942 !important;
        border: none; font-weight: 700; border-radius: 8px;
    }
    .stButton > button:hover { background: linear-gradient(135deg, #ffb940, #f5a623); }
    .stDownloadButton > button {
        background: linear-gradient(135deg, #22c55e, #16a34a); color: white !important;
        border: none; font-weight: 700; border-radius: 8px; padding: 10px 24px;
    }
    [data-testid="stHeader"] { background: transparent; }
    .brand-header {
        background: linear-gradient(135deg, #141942, #1c2255);
        border: 1px solid rgba(255,255,255,0.07);
        border-radius: 16px; padding: 22px 28px; margin-bottom: 20px;
        border-top: 3px solid #f5a623;
    }
    .brand-tag { font-size: 11px; font-weight: 700; letter-spacing: 2.5px; color: #f5a623; text-transform: uppercase; }
    .brand-title { font-size: 22px; font-weight: 800; color: white; margin: 4px 0 2px; }
    .brand-sub { font-size: 12px; color: #7a82a6; }
    .section-card {
        background: #181e4a;
        border: 1px solid rgba(255,255,255,0.07);
        border-radius: 12px; padding: 18px; margin-bottom: 14px;
    }
    .section-num {
        display: inline-flex; align-items: center; justify-content: center;
        width: 28px; height: 28px; border-radius: 7px;
        background: linear-gradient(135deg, #f5a623, #e09000);
        color: #141942; font-weight: 800; margin-right: 10px;
    }
    .persist-info {
        font-size: 11px; color: #22c55e; background: rgba(34,197,94,0.08);
        border: 1px solid rgba(34,197,94,0.15); padding: 4px 12px;
        border-radius: 20px; display: inline-block;
    }
</style>
""", unsafe_allow_html=True)


# ═══════════════════════════════════════════════════════════════
# PERSISTENCE
# ═══════════════════════════════════════════════════════════════
def load_persist():
    try:
        if PERSIST_FILE.exists():
            data = json.loads(PERSIST_FILE.read_text())
            return {
                "pic_project": data.get("pic_project", ""),
                "team_support": data.get("team_support", "")
            }
    except Exception:
        pass
    return {"pic_project": "", "team_support": ""}

def save_persist(data):
    to_save = {
        "pic_project": data.get("pic_project", ""),
        "team_support": data.get("team_support", "")
    }
    try:
        PERSIST_FILE.write_text(json.dumps(to_save, ensure_ascii=False, indent=2))
    except Exception:
        pass

def load_full_data():
    default = {
        "activities": [{"waktu": "", "uraian": "", "keterangan": "", "status": "Selesai"}],
        "incidents": [],
        "materials": [],
        "followups": [],
        "catatan": "",
    }
    try:
        if FULL_DATA_FILE.exists():
            data = json.loads(FULL_DATA_FILE.read_text())
            for k in default:
                if k not in data:
                    data[k] = default[k]
            return data
    except Exception:
        pass
    return default

def save_full_data(data):
    to_save = {k: v for k, v in data.items() if k != "photos"}
    try:
        FULL_DATA_FILE.write_text(json.dumps(to_save, ensure_ascii=False, indent=2))
    except Exception:
        pass

# Inisialisasi session state
if "full_data" not in st.session_state:
    st.session_state.full_data = load_full_data()
if "photos" not in st.session_state:
    st.session_state.photos = []

def auto_save_full_data():
    save_full_data(st.session_state.full_data)

# ═══════════════════════════════════════════════════════════════
# IMAGE HELPERS
# ═══════════════════════════════════════════════════════════════
def resize_image(img_bytes: bytes, max_width: int = 1200, quality: int = 85) -> bytes:
    img = Image.open(io.BytesIO(img_bytes))
    img = ImageOps.exif_transpose(img)
    if img.mode not in ("RGB", "L"):
        img = img.convert("RGB")
    if img.width > max_width:
        ratio = max_width / img.width
        new_height = int(img.height * ratio)
        img = img.resize((max_width, new_height), Image.Resampling.LANCZOS)
    out = io.BytesIO()
    img.save(out, format="JPEG", quality=quality, optimize=True)
    return out.getvalue()

def rotate_image(img_bytes: bytes, degrees: int) -> bytes:
    img = Image.open(io.BytesIO(img_bytes))
    img = ImageOps.exif_transpose(img)
    if degrees == 90:
        img = img.transpose(Image.ROTATE_90)
    elif degrees == -90:
        img = img.transpose(Image.ROTATE_270)
    if img.mode not in ("RGB", "L"):
        img = img.convert("RGB")
    out = io.BytesIO()
    img.save(out, format="JPEG", quality=85, optimize=True)
    return out.getvalue()

def fix_image_orientation(img_bytes: bytes) -> bytes:
    img = Image.open(io.BytesIO(img_bytes))
    img = ImageOps.exif_transpose(img)
    if img.mode not in ("RGB", "L"):
        img = img.convert("RGB")
    out = io.BytesIO()
    img.save(out, format="JPEG", quality=85, optimize=True)
    return out.getvalue()

# ═══════════════════════════════════════════════════════════════
# DOCX MANIPULATION HELPERS (tambah fungsi untuk clone tabel)
# ═══════════════════════════════════════════════════════════════
def set_cell_text(cell, text, *, bold=None, font_size_pt=None, align=None, preserve_format=True):
    existing_font = None
    existing_size = None
    existing_bold = None
    existing_color = None
    if preserve_format and cell.paragraphs and cell.paragraphs[0].runs:
        first_run = cell.paragraphs[0].runs[0]
        existing_font = first_run.font.name
        existing_size = first_run.font.size
        existing_bold = first_run.font.bold
        if first_run.font.color and first_run.font.color.rgb:
            existing_color = first_run.font.color.rgb
    for p in cell.paragraphs:
        for r in list(p.runs):
            r._element.getparent().remove(r._element)
    p = cell.paragraphs[0]
    if align is not None:
        p.alignment = align
    lines = str(text).split("\n") if text else [""]
    for i, line in enumerate(lines):
        if i > 0:
            run = p.add_run()
            run.add_break()
        run = p.add_run(line)
        if existing_font:
            run.font.name = existing_font
        if font_size_pt is not None:
            run.font.size = Pt(font_size_pt)
        elif existing_size:
            run.font.size = existing_size
        if bold is not None:
            run.font.bold = bold
        elif existing_bold is not None:
            run.font.bold = existing_bold
        if existing_color:
            run.font.color.rgb = existing_color

def set_cell_text_clean(cell, text, *, align=None):
    existing_font = None
    existing_size = None
    existing_bold = None
    if cell.paragraphs and cell.paragraphs[0].runs:
        first_run = cell.paragraphs[0].runs[0]
        existing_font = first_run.font.name
        existing_size = first_run.font.size
        existing_bold = first_run.font.bold
    first_p = cell.paragraphs[0]
    extras = list(cell.paragraphs)[1:]
    for p in extras:
        p._element.getparent().remove(p._element)
    for r in list(first_p.runs):
        r._element.getparent().remove(r._element)
    pPr = first_p._element.find(qn('w:pPr'))
    if pPr is not None:
        numPr = pPr.find(qn('w:numPr'))
        if numPr is not None:
            pPr.remove(numPr)
    if align is not None:
        first_p.alignment = align
    lines = str(text).split("\n") if text else [""]
    for i, line in enumerate(lines):
        if i == 0:
            p = first_p
        else:
            p = cell.add_paragraph()
            pPr2 = p._element.find(qn('w:pPr'))
            if pPr2 is not None:
                numPr2 = pPr2.find(qn('w:numPr'))
                if numPr2 is not None:
                    pPr2.remove(numPr2)
            if align is not None:
                p.alignment = align
        run = p.add_run(line)
        if existing_font:
            run.font.name = existing_font
        if existing_size:
            run.font.size = existing_size
        if existing_bold is not None:
            run.font.bold = existing_bold

def force_page_break_before(paragraph):
    from docx.oxml import OxmlElement
    pPr = paragraph._element.find(qn('w:pPr'))
    if pPr is None:
        pPr = OxmlElement('w:pPr')
        paragraph._element.insert(0, pPr)
    existing = pPr.find(qn('w:pageBreakBefore'))
    if existing is not None:
        pPr.remove(existing)
    pbb = OxmlElement('w:pageBreakBefore')
    pPr.insert(0, pbb)

def remove_element(elem):
    parent = elem.getparent()
    if parent is not None:
        parent.remove(elem)

def add_image_to_cell(cell, img_bytes, caption, width_cm=7.5):
    for p in cell.paragraphs:
        for r in list(p.runs):
            r._element.getparent().remove(r._element)
    paragraphs = list(cell.paragraphs)
    for p in paragraphs[1:]:
        p._element.getparent().remove(p._element)
    p_img = cell.paragraphs[0]
    p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p_img.add_run()
    run.add_picture(io.BytesIO(img_bytes), width=Cm(width_cm))
    p_cap = cell.add_paragraph()
    p_cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
    cap_run = p_cap.add_run(caption or "Foto Kegiatan")
    cap_run.font.size = Pt(9)

def fill_data_table(table, rows_data, col_extractors, center_cols=None):
    if center_cols is None:
        center_cols = []
    if len(table.rows) < 2:
        return
    tbl_element = table._tbl
    existing_rows = tbl_element.findall(qn('w:tr'))
    for tr in existing_rows[1:]:
        tbl_element.remove(tr)
    if rows_data:
        template_row = existing_rows[1] if len(existing_rows) > 1 else None
        if template_row is None:
            return
        for _ in rows_data:
            new_tr = deepcopy(template_row)
            tbl_element.append(new_tr)
    for i, row_data in enumerate(rows_data):
        if i+1 >= len(table.rows):
            break
        row = table.rows[i+1]
        for col_idx, extractor in enumerate(col_extractors):
            if col_idx >= len(row.cells):
                break
            value = extractor(i, row_data)
            align = WD_ALIGN_PARAGRAPH.CENTER if col_idx in center_cols else WD_ALIGN_PARAGRAPH.LEFT
            set_cell_text(row.cells[col_idx], value, align=align)

def create_photo_table(doc, template_table, rows=4, cols=2):
    """
    Clone template table (yang berisi 4x2 grid) dan tambahkan ke doc.
    template_table adalah tabel yang sudah ada (misal doc.tables[6]).
    """
    # Clone seluruh elemen tabel
    new_tbl = deepcopy(template_table._tbl)
    # Tambahkan ke body dokumen sebelum section break
    body = doc.element.body
    # Cari posisi terakhir setelah semua paragraf/tabel
    body.append(new_tbl)
    # Return tabel yang baru dibuat
    return doc.tables[-1]  # karena baru ditambahkan di akhir

def add_photo_page(doc, template_table):
    """
    Menambah halaman baru dengan heading "LAMPIRAN FOTO KEGIATAN" dan tabel foto (4x2).
    Return tabel baru.
    """
    # Tambah page break
    p = doc.add_paragraph()
    force_page_break_before(p)
    # Tambah heading
    heading = doc.add_paragraph()
    run = heading.add_run("LAMPIRAN FOTO KEGIATAN")
    run.bold = True
    run.underline = True
    heading.alignment = WD_ALIGN_PARAGRAPH.CENTER
    # Tambah tabel baru hasil clone dari template_table
    new_table = create_photo_table(doc, template_table)
    return new_table

# ═══════════════════════════════════════════════════════════════
# CORE BUILD REPORT (DOCX) - tanpa batasan foto
# ═══════════════════════════════════════════════════════════════
def format_tanggal(d):
    if isinstance(d, date):
        return f"{d.day} {BULAN_ID[d.month-1]} {d.year}"
    return str(d)

def build_report_docx(data, photos):
    doc = Document(str(TEMPLATE_PATH))
    
    # Table 0: Informasi Umum
    t_info = doc.tables[0]
    r0_tcs = t_info._tbl.findall(qn('w:tr'))[0].findall(qn('w:tc'))
    from docx.table import _Cell
    pic_value_cell = _Cell(r0_tcs[1], t_info)
    tanggal_value_cell = _Cell(r0_tcs[3], t_info)
    set_cell_text(pic_value_cell, data["pic_project"] or "-")
    set_cell_text(tanggal_value_cell, format_tanggal(data["tanggal"]))
    
    r1_tcs = t_info._tbl.findall(qn('w:tr'))[1].findall(qn('w:tc'))
    team_col1_cell = _Cell(r1_tcs[1], t_info)
    team_col2_cell = _Cell(r1_tcs[2], t_info)
    lokasi_value_cell = _Cell(r1_tcs[4], t_info)
    
    team_list = [t.strip() for t in re.split(r"[,;\n]", data["team_support"]) if t.strip()]
    if team_list:
        mid = (len(team_list)+1)//2
        col1 = team_list[:mid]
        col2 = team_list[mid:]
        col1_text = "\n".join(f"{i+1}. {name}" for i,name in enumerate(col1))
        col2_text = "\n".join(f"{mid+i+1}. {name}" for i,name in enumerate(col2))
        set_cell_text_clean(team_col1_cell, col1_text)
        set_cell_text_clean(team_col2_cell, col2_text)
    else:
        set_cell_text_clean(team_col1_cell, "-")
        set_cell_text_clean(team_col2_cell, "")
    set_cell_text(lokasi_value_cell, data["lokasi"] or "-")
    
    # Table 1: Kegiatan
    fill_data_table(doc.tables[1], data["activities"],
        [lambda i,a: str(i+1), lambda i,a: a["waktu"], lambda i,a: a["uraian"], lambda i,a: a["keterangan"], lambda i,a: a["status"]],
        center_cols=[0,4])
    
    # Table 2: Kendala
    if data["incidents"]:
        fill_data_table(doc.tables[2], data["incidents"],
            [lambda i,x: str(i+1), lambda i,x: x["masalah"], lambda i,x: x["tindakan"], lambda i,x: x["hasil"]],
            center_cols=[0])
    else:
        fill_data_table(doc.tables[2], [{"masalah":"","tindakan":"","hasil":""}],
            [lambda i,x: "", lambda i,x: x["masalah"], lambda i,x: x["tindakan"], lambda i,x: x["hasil"]],
            center_cols=[0])
    
    # Table 3: Follow Up
    followups = data.get("followups", [])
    if followups:
        fill_data_table(doc.tables[3], followups,
            [lambda i,f: str(i+1), lambda i,f: f["deskripsi"], lambda i,f: f["alasan"], lambda i,f: f["target"]],
            center_cols=[0])
    else:
        fill_data_table(doc.tables[3], [{"deskripsi":"","alasan":"","target":""}],
            [lambda i,f: "", lambda i,f: f["deskripsi"], lambda i,f: f["alasan"], lambda i,f: f["target"]],
            center_cols=[0])
    
    # Table 4: Catatan
    t_catatan = doc.tables[4]
    set_cell_text(t_catatan.rows[0].cells[0], data["catatan"] or " ")
    
    # Table 5: Material
    if data["materials"]:
        fill_data_table(doc.tables[5], data["materials"],
            [lambda i,m: str(i+1), lambda i,m: m["nama"], lambda i,m: m["qty"], lambda i,m: m["kondisi"], lambda i,m: m["keterangan"]],
            center_cols=[0,2,3])
    else:
        fill_data_table(doc.tables[5], [{"nama":"-","qty":"","kondisi":"","keterangan":""}],
            [lambda i,m: "", lambda i,m: m["nama"], lambda i,m: m["qty"], lambda i,m: m["kondisi"], lambda i,m: m["keterangan"]],
            center_cols=[0,1,2,3])
    
    # ========== PHOTO SECTION - TANPA BATASAN ==========
    if photos:
        # Template tabel foto asli (Table 6) yang memiliki 4 baris x 2 kolom = 8 slot
        template_table = doc.tables[6]
        # Hapus Table 6 dan Table 7 dari template karena kita akan membuat ulang semuanya
        # Kita akan hapus kedua tabel tersebut dan membuat halaman foto baru secara dinamis.
        # Cara: cari semua tabel setelah Table 5, lalu hapus (karena kita akan buat ulang)
        # Lebih mudah: simpan referensi template_table untuk cloning, lalu hapus semua tabel mulai dari indeks 6.
        tables_to_remove = []
        for i in range(6, len(doc.tables)):
            tables_to_remove.append(doc.tables[i])
        for tbl in tables_to_remove:
            tbl._element.getparent().remove(tbl._element)
        
        # Sekarang buat halaman foto pertama
        # Pastikan ada page break sebelum lampiran foto
        p_break = doc.add_paragraph()
        force_page_break_before(p_break)
        heading1 = doc.add_paragraph()
        run1 = heading1.add_run("LAMPIRAN FOTO KEGIATAN")
        run1.bold = True
        run1.underline = True
        heading1.alignment = WD_ALIGN_PARAGRAPH.CENTER
        
        # Buat tabel pertama dengan clone template_table
        new_table = create_photo_table(doc, template_table)
        
        # Isi foto per 8
        photos_per_page = 8
        total_photos = len(photos)
        photo_idx = 0
        
        # Fungsi untuk mengisi tabel tertentu dengan foto
        def fill_one_table(table, start_idx):
            row_idx = 0
            for row in table.rows:
                for cell in row.cells:
                    if start_idx + row_idx*2 + (row_idx==0?0:0) < len(photos):
                        p = photos[start_idx + row_idx*2 + (cell._element.getparent().index(cell._element) % 2)]
                        # Agak tricky, lebih mudah iterasi cell berurutan
                        pass
            # Cara simple: urutkan cell dari kiri ke kanan, atas ke bawah
            cells = []
            for row in table.rows:
                for cell in row.cells:
                    cells.append(cell)
            for i, cell in enumerate(cells):
                if start_idx + i < len(photos):
                    p = photos[start_idx + i]
                    add_image_to_cell(cell, p["bytes"], p["caption"])
                else:
                    # kosongkan cell
                    for para in cell.paragraphs:
                        for r in list(para.runs):
                            r._element.getparent().remove(r._element)
                    extras = list(cell.paragraphs)[1:]
                    for ex in extras:
                        ex._element.getparent().remove(ex._element)
        
        # Isi tabel pertama
        cells_first = []
        for row in new_table.rows:
            for cell in row.cells:
                cells_first.append(cell)
        for i, cell in enumerate(cells_first):
            if i < len(photos):
                p = photos[i]
                add_image_to_cell(cell, p["bytes"], p["caption"])
            else:
                for para in cell.paragraphs:
                    for r in list(para.runs):
                        r._element.getparent().remove(r._element)
                extras = list(cell.paragraphs)[1:]
                for ex in extras:
                    ex._element.getparent().remove(ex._element)
        
        # Jika foto lebih dari 8, tambah halaman baru setiap 8 foto
        remaining = photos[8:]
        page_num = 1
        for start in range(8, len(photos), 8):
            # Tambah halaman baru dengan heading dan tabel
            p_break_new = doc.add_paragraph()
            force_page_break_before(p_break_new)
            heading_new = doc.add_paragraph()
            run_new = heading_new.add_run("LAMPIRAN FOTO KEGIATAN (lanjutan)")
            run_new.bold = True
            run_new.underline = True
            heading_new.alignment = WD_ALIGN_PARAGRAPH.CENTER
            new_table_page = create_photo_table(doc, template_table)
            cells_new = []
            for row in new_table_page.rows:
                for cell in row.cells:
                    cells_new.append(cell)
            for i, cell in enumerate(cells_new):
                idx = start + i
                if idx < len(photos):
                    p = photos[idx]
                    add_image_to_cell(cell, p["bytes"], p["caption"])
                else:
                    for para in cell.paragraphs:
                        for r in list(para.runs):
                            r._element.getparent().remove(r._element)
                    extras = list(cell.paragraphs)[1:]
                    for ex in extras:
                        ex._element.getparent().remove(ex._element)
    else:
        # Tidak ada foto, hapus bagian foto di template asli
        # Hapus semua paragraf "LAMPIRAN FOTO KEGIATAN" dan tabel 6,7
        body = doc.element.body
        to_remove = []
        in_photo = False
        for child in list(body.iterchildren()):
            tag = child.tag.split('}')[-1]
            if tag == 'p':
                texts = [t.text for t in child.iter(qn('w:t')) if t.text]
                text = ' '.join(texts).upper()
                if 'LAMPIRAN FOTO KEGIATAN' in text:
                    in_photo = True
                    to_remove.append(child)
                    continue
            if in_photo and tag != 'sectPr':
                to_remove.append(child)
        for elem in to_remove:
            remove_element(elem)
    
    output = io.BytesIO()
    doc.save(output)
    output.seek(0)
    return output

# ═══════════════════════════════════════════════════════════════
# PDF GENERATION (4 foto per halaman, grid 2x2)
# ═══════════════════════════════════════════════════════════════
class PDF(FPDF):
    def header(self):
        self.set_font('Arial', 'B', 12)
        self.cell(0, 10, 'DAILY REPORT - NALE SYSTEM INTEGRATOR', 0, 1, 'C')
        self.set_font('Arial', '', 10)
        self.cell(0, 5, 'Managed Service Dinas Kominfo Kab. Mojokerto', 0, 1, 'C')
        self.ln(10)
    
    def footer(self):
        self.set_y(-15)
        self.set_font('Arial', 'I', 8)
        self.cell(0, 10, f'Halaman {self.page_no()}', 0, 0, 'C')
    
    def section_title(self, num, title):
        self.set_font('Arial', 'B', 11)
        self.cell(0, 8, f'{num}. {title}', 0, 1, 'L')
        self.ln(2)
    
    def add_table(self, headers, rows, col_widths=None):
        if not rows:
            return
        if not col_widths:
            col_widths = [40] * len(headers)
        self.set_font('Arial', 'B', 9)
        for i, header in enumerate(headers):
            self.cell(col_widths[i], 8, header, 1, 0, 'C')
        self.ln()
        self.set_font('Arial', '', 9)
        for row in rows:
            for i, cell in enumerate(row):
                self.cell(col_widths[i], 8, str(cell), 1, 0, 'L')
            self.ln()

def build_report_pdf(data, photos):
    pdf = PDF()
    pdf.add_page()
    
    # Informasi Umum
    pdf.section_title('1', 'Informasi Umum')
    pdf.set_font('Arial', '', 10)
    pdf.cell(40, 6, f"PIC Project: {data['pic_project']}", 0, 1)
    pdf.cell(40, 6, f"Tanggal: {format_tanggal(data['tanggal'])}", 0, 1)
    pdf.cell(40, 6, f"Team Support: {data['team_support']}", 0, 1)
    pdf.cell(40, 6, f"Lokasi: {data['lokasi']}", 0, 1)
    pdf.ln(5)
    
    # Kegiatan Harian
    pdf.section_title('2', 'Kegiatan Harian')
    kegiatan_rows = [[i+1, a['waktu'], a['uraian'], a['keterangan'], a['status']] 
                     for i, a in enumerate(data['activities'])]
    pdf.add_table(['No','Waktu','Uraian','Keterangan','Status'], kegiatan_rows, [15,30,70,40,25])
    
    # Kendala
    pdf.section_title('3', 'Kendala / Insiden')
    if data['incidents']:
        kendala_rows = [[i+1, inc['masalah'], inc['tindakan'], inc['hasil']] 
                        for i, inc in enumerate(data['incidents'])]
        pdf.add_table(['No','Masalah','Tindakan','Hasil'], kendala_rows, [15,65,55,55])
    else:
        pdf.cell(0, 6, "Tidak ada kendala.", 0, 1)
    
    # Follow Up
    pdf.section_title('4', 'Pekerjaan Belum Selesai / Follow Up')
    if data['followups']:
        fu_rows = [[i+1, fu['deskripsi'], fu['alasan'], fu['target']] 
                   for i, fu in enumerate(data['followups'])]
        pdf.add_table(['No','Deskripsi','Alasan/Kendala','Target'], fu_rows, [15,60,60,60])
    else:
        pdf.cell(0, 6, "Tidak ada follow up.", 0, 1)
    
    # Catatan Tambahan
    pdf.section_title('5', 'Catatan Tambahan')
    pdf.multi_cell(0, 6, data['catatan'] if data['catatan'] else "-")
    
    # Material
    pdf.section_title('6', 'Material yang Digunakan')
    if data['materials']:
        mat_rows = [[i+1, m['nama'], m['qty'], m['kondisi'], m['keterangan']] 
                    for i, m in enumerate(data['materials'])]
        pdf.add_table(['No','Nama','Qty','Kondisi','Keterangan'], mat_rows, [15,70,20,30,55])
    else:
        pdf.cell(0, 6, "Tidak ada material.", 0, 1)
    
    # ========== LAMPIRAN FOTO - 4 FOTO PER HALAMAN (GRID 2x2) ==========
    if photos:
        pdf.add_page()
        pdf.set_font('Arial', 'B', 11)
        pdf.cell(0, 8, 'LAMPIRAN FOTO KEGIATAN', 0, 1, 'C')
        pdf.ln(5)
        
        # Atur grid 2 kolom, 2 baris per halaman
        img_width = 80  # mm
        img_height = 60  # mm (proporsional, nanti disesuaikan)
        x_left = 20
        x_right = 110  # 20 + 80 + 10 = 110
        y_start = pdf.get_y()
        y_step = 70  # tinggi per baris foto + caption
        
        for idx, p in enumerate(photos):
            # Tentukan posisi baris dan kolom
            row = idx // 2
            col = idx % 2
            if row >= 2:  # lebih dari 2 baris, pindah halaman
                pdf.add_page()
                pdf.set_font('Arial', 'B', 11)
                pdf.cell(0, 8, 'LAMPIRAN FOTO KEGIATAN (lanjutan)', 0, 1, 'C')
                pdf.ln(5)
                y_start = pdf.get_y()
                row = 0
                # reset idx untuk perhitungan? kita lanjutkan dengan idx yang sama tapi row baru
                # Karena kita pindah halaman, kita set ulang row=0, col=idx%2
                # Kita perlu simpan offset halaman, lebih mudah pakai while
                pass
            # Karena logika di atas tidak sempurna, lebih baik pakai pendekatan per halaman:
            # Kita buat per 4 foto.
        
        # Pendekatan lebih sederhana: loop per 4 foto
        for page_start in range(0, len(photos), 4):
            if page_start > 0:
                pdf.add_page()
                pdf.set_font('Arial', 'B', 11)
                pdf.cell(0, 8, 'LAMPIRAN FOTO KEGIATAN (lanjutan)', 0, 1, 'C')
                pdf.ln(5)
            y_current = pdf.get_y()
            for i in range(4):
                idx = page_start + i
                if idx >= len(photos):
                    break
                p = photos[idx]
                col = i % 2
                row = i // 2
                x = x_left if col == 0 else x_right
                y = y_current + row * y_step
                try:
                    with tempfile.NamedTemporaryFile(delete=False, suffix='.jpg') as tmp:
                        tmp.write(p['bytes'])
                        tmp_path = tmp.name
                    pdf.image(tmp_path, x=x, y=y, w=img_width)
                    # Caption di bawah foto
                    pdf.set_xy(x, y + img_height)
                    pdf.set_font('Arial', 'I', 8)
                    pdf.cell(img_width, 5, f"Foto {idx+1}: {p['caption']}", 0, 1, 'C')
                    os.unlink(tmp_path)
                except Exception:
                    pdf.set_xy(x, y)
                    pdf.cell(img_width, img_height, f"Foto {idx+1}: Gagal memuat", 1, 1, 'C')
    
    output = io.BytesIO()
    pdf.output(output)
    output.seek(0)
    return output

# ═══════════════════════════════════════════════════════════════
# UI (sama seperti sebelumnya, hanya modifikasi caption di tab foto)
# ═══════════════════════════════════════════════════════════════

st.markdown("""
<div class="brand-header">
  <div>
    <div class="brand-tag">NALE SYSTEM INTEGRATOR</div>
    <div class="brand-title">📋 Daily Report Generator</div>
    <div class="brand-sub">Managed Service — Dinas Kominfo Kab. Mojokerto</div>
  </div>
</div>
""", unsafe_allow_html=True)

if not TEMPLATE_PATH.exists():
    st.error(f"❌ File template tidak ditemukan: `{TEMPLATE_PATH}`")
    st.stop()

# Load data
persist = load_persist()
full_data = st.session_state.full_data

# Tab
tab_info, tab_photos = st.tabs(["📋 Info & Kegiatan", "📸 Foto"])

# ═══════════════════════════════════════════════════════════════
# TAB INFO & KEGIATAN (sama seperti sebelumnya, disingkat)
# ═══════════════════════════════════════════════════════════════
with tab_info:
    # --- 1. Informasi Umum ---
    st.markdown('<div class="section-card">', unsafe_allow_html=True)
    col_a, col_b = st.columns([3,1])
    with col_a:
        st.markdown('<h3><span class="section-num">1</span>Informasi Umum</h3>', unsafe_allow_html=True)
    with col_b:
        st.markdown('<div class="persist-info">● Auto-saved (PIC & Team)</div>', unsafe_allow_html=True)
    
    pic_project = st.text_input("PIC Project", value=persist.get("pic_project", ""), placeholder="Nama PIC Project", key="pic_project")
    tanggal = st.date_input("Tanggal", value=date.today(), key="tanggal", format="DD/MM/YYYY")
    team_support = st.text_input("Team Support (pisahkan koma)", value=persist.get("team_support", ""), placeholder="Contoh: Andi, Budi, Citra", key="team_support")
    lokasi = st.text_input("Lokasi / Site", value="", placeholder="Nama lokasi kegiatan", key="lokasi")
    
    new_persist = {"pic_project": pic_project, "team_support": team_support}
    if new_persist != persist:
        save_persist(new_persist)
    st.markdown('</div>', unsafe_allow_html=True)
    
    # --- 2. Kegiatan Harian ---
    st.markdown('<div class="section-card">', unsafe_allow_html=True)
    st.markdown('<h3><span class="section-num">2</span>Kegiatan Harian</h3>', unsafe_allow_html=True)
    activities = full_data.setdefault("activities", [{"waktu":"","uraian":"","keterangan":"","status":"Selesai"}])
    for i, act in enumerate(activities):
        cols = st.columns([0.4,1.3,3,1.5,1.2,0.4])
        cols[0].markdown(f"<div style='padding-top:32px;text-align:center'>{i+1}</div>", unsafe_allow_html=True)
        act["waktu"] = cols[1].text_input("Waktu", act["waktu"], placeholder="08.00-10.00", key=f"act_w_{i}", label_visibility="collapsed")
        act["uraian"] = cols[2].text_input("Uraian", act["uraian"], placeholder="Uraian kegiatan", key=f"act_u_{i}", label_visibility="collapsed")
        act["keterangan"] = cols[3].text_input("Keterangan", act["keterangan"], placeholder="Keterangan (opsional)", key=f"act_k_{i}", label_visibility="collapsed")
        act["status"] = cols[4].selectbox("Status", ["Selesai","Proses","Pending"], index=["Selesai","Proses","Pending"].index(act["status"]) if act["status"] in ["Selesai","Proses","Pending"] else 0, key=f"act_s_{i}", label_visibility="collapsed")
        if cols[5].button("✕", key=f"act_del_{i}"):
            activities.pop(i)
            auto_save_full_data()
            st.rerun()
    if st.button("+ Tambah Kegiatan", key="add_act"):
        activities.append({"waktu":"","uraian":"","keterangan":"","status":"Selesai"})
        auto_save_full_data()
        st.rerun()
    st.markdown('</div>', unsafe_allow_html=True)
    
    # --- 3. Kendala ---
    st.markdown('<div class="section-card">', unsafe_allow_html=True)
    st.markdown('<h3><span class="section-num">3</span>Kendala / Insiden</h3>', unsafe_allow_html=True)
    incidents = full_data.setdefault("incidents", [])
    for i, inc in enumerate(incidents):
        cols = st.columns([0.4,3,3,3,0.4])
        cols[0].markdown(f"<div style='padding-top:32px'>{i+1}</div>", unsafe_allow_html=True)
        inc["masalah"] = cols[1].text_area("Masalah", inc["masalah"], placeholder="Deskripsi masalah / problem", key=f"inc_m_{i}", label_visibility="collapsed", height=70)
        inc["tindakan"] = cols[2].text_area("Tindakan", inc["tindakan"], placeholder="Tindakan yang dilakukan", key=f"inc_t_{i}", label_visibility="collapsed", height=70)
        inc["hasil"] = cols[3].text_area("Hasil", inc["hasil"], placeholder="Hasil dari tindakan", key=f"inc_h_{i}", label_visibility="collapsed", height=70)
        if cols[4].button("✕", key=f"inc_del_{i}"):
            incidents.pop(i)
            auto_save_full_data()
            st.rerun()
    if st.button("+ Tambah Insiden", key="add_inc"):
        incidents.append({"masalah":"","tindakan":"","hasil":""})
        auto_save_full_data()
        st.rerun()
    st.markdown('</div>', unsafe_allow_html=True)
    
    # --- 4. Follow Up ---
    st.markdown('<div class="section-card">', unsafe_allow_html=True)
    st.markdown('<h3><span class="section-num">4</span>Pekerjaan Belum Selesai / Follow Up</h3>', unsafe_allow_html=True)
    followups = full_data.setdefault("followups", [])
    for i, fu in enumerate(followups):
        cols = st.columns([0.4,3,3,3,0.4])
        cols[0].markdown(f"<div style='padding-top:32px'>{i+1}</div>", unsafe_allow_html=True)
        fu["deskripsi"] = cols[1].text_area("Deskripsi", fu.get("deskripsi",""), placeholder="Pekerjaan yang belum selesai", key=f"fu_desc_{i}", label_visibility="collapsed", height=70)
        fu["alasan"] = cols[2].text_area("Alasan/Kendala", fu.get("alasan",""), placeholder="Alasan atau kendala", key=f"fu_alasan_{i}", label_visibility="collapsed", height=70)
        fu["target"] = cols[3].text_area("Target", fu.get("target",""), placeholder="Target penyelesaian", key=f"fu_target_{i}", label_visibility="collapsed", height=70)
        if cols[4].button("✕", key=f"fu_del_{i}"):
            followups.pop(i)
            auto_save_full_data()
            st.rerun()
    if st.button("+ Tambah Follow Up", key="add_fu"):
        followups.append({"deskripsi":"","alasan":"","target":""})
        auto_save_full_data()
        st.rerun()
    st.markdown('</div>', unsafe_allow_html=True)
    
    # --- 5. Catatan ---
    st.markdown('<div class="section-card">', unsafe_allow_html=True)
    st.markdown('<h3><span class="section-num">5</span>Catatan Tambahan</h3>', unsafe_allow_html=True)
    catatan = st.text_area("Catatan", value=full_data.get("catatan",""), placeholder="Catatan tambahan (opsional)...", key="catatan", height=80, label_visibility="collapsed")
    full_data["catatan"] = catatan
    st.markdown('</div>', unsafe_allow_html=True)
    
    # --- 6. Material ---
    st.markdown('<div class="section-card">', unsafe_allow_html=True)
    st.markdown('<h3><span class="section-num">6</span>Material yang Digunakan</h3>', unsafe_allow_html=True)
    materials = full_data.setdefault("materials", [])
    for i, mat in enumerate(materials):
        cols = st.columns([0.4,3,0.8,1.2,2.5,0.4])
        cols[0].markdown(f"<div style='padding-top:32px'>{i+1}</div>", unsafe_allow_html=True)
        mat["nama"] = cols[1].text_input("Nama", mat.get("nama",""), placeholder="Nama material/alat", key=f"mat_n_{i}", label_visibility="collapsed")
        mat["qty"] = cols[2].text_input("Qty", mat.get("qty",""), placeholder="Jumlah", key=f"mat_q_{i}", label_visibility="collapsed")
        mat["kondisi"] = cols[3].selectbox("Kondisi", ["Baik","Rusak","Habis"], index=["Baik","Rusak","Habis"].index(mat.get("kondisi","Baik")), key=f"mat_k_{i}", label_visibility="collapsed")
        mat["keterangan"] = cols[4].text_input("Keterangan", mat.get("keterangan",""), placeholder="Keterangan (opsional)", key=f"mat_ket_{i}", label_visibility="collapsed")
        if cols[5].button("✕", key=f"mat_del_{i}"):
            materials.pop(i)
            auto_save_full_data()
            st.rerun()
    if st.button("+ Tambah Material", key="add_mat"):
        materials.append({"nama":"","qty":"","kondisi":"Baik","keterangan":""})
        auto_save_full_data()
        st.rerun()
    st.markdown('</div>', unsafe_allow_html=True)
    
    # Tombol simpan manual
    st.markdown("---")
    col_save1, col_save2 = st.columns(2)
    with col_save1:
        if st.button("💾 Simpan Draft (Semua Data)", use_container_width=True):
            auto_save_full_data()
            st.success("Data disimpan ke file. Tidak akan hilang meskipun refresh.")
    with col_save2:
        if st.button("🔄 Muat Ulang Draft", use_container_width=True):
            st.session_state.full_data = load_full_data()
            st.rerun()

# ═══════════════════════════════════════════════════════════════
# TAB FOTO (tanpa batasan)
# ═══════════════════════════════════════════════════════════════
with tab_photos:
    st.markdown('<div class="section-card">', unsafe_allow_html=True)
    st.markdown('<h3><span class="section-num">📸</span>Lampiran Foto Kegiatan</h3>', unsafe_allow_html=True)
    st.caption("Gunakan tombol ↻ untuk rotate gambar. **Tidak ada batasan jumlah foto.** Semua foto akan masuk ke laporan (PDF: 4 foto/halaman, DOCX: 8 foto/halaman).")
    
    uploaded_files = st.file_uploader(
        "Pilih foto (bisa pilih banyak, tanpa batasan jumlah)",
        type=["jpg","jpeg","png","webp"],
        accept_multiple_files=True,
        key="photo_uploader_multiple"
    )
    
    col_add, col_info = st.columns([1, 2])
    with col_add:
        if st.button("➕ Tambahkan Foto ke Daftar", use_container_width=True):
            if uploaded_files:
                existing_names = {p["name"] for p in st.session_state.photos}
                new_added = 0
                for f in uploaded_files:
                    if f.name not in existing_names:
                        raw = f.read()
                        fixed = fix_image_orientation(raw)
                        resized = resize_image(fixed, max_width=1200, quality=85)
                        st.session_state.photos.append({
                            "bytes": resized,
                            "caption": "",
                            "name": f.name,
                        })
                        new_added += 1
                if new_added > 0:
                    st.success(f"{new_added} foto ditambahkan!")
                    st.rerun()
                else:
                    st.warning("Tidak ada foto baru yang ditambahkan (semua sudah ada).")
            else:
                st.warning("Silakan pilih file foto terlebih dahulu.")
    
    with col_info:
        st.caption(f"Total foto: {len(st.session_state.photos)}")
    
    if st.session_state.photos:
        st.markdown(f"**Daftar Foto ({len(st.session_state.photos)}):**")
        for i, photo in enumerate(st.session_state.photos):
            cols = st.columns([1.5, 3, 0.8, 0.8, 0.5])
            with cols[0]:
                st.image(photo["bytes"], use_container_width=True)
            with cols[1]:
                new_caption = st.text_area(f"Caption {i+1}", value=photo["caption"], placeholder="Deskripsi foto kegiatan...", key=f"photo_cap_{i}", height=100, label_visibility="collapsed")
                st.session_state.photos[i]["caption"] = new_caption
            with cols[2]:
                if st.button("↻ Kiri", key=f"rot_left_{i}"):
                    rotated = rotate_image(photo["bytes"], -90)
                    st.session_state.photos[i]["bytes"] = rotated
                    st.rerun()
            with cols[3]:
                if st.button("↻ Kanan", key=f"rot_right_{i}"):
                    rotated = rotate_image(photo["bytes"], 90)
                    st.session_state.photos[i]["bytes"] = rotated
                    st.rerun()
            with cols[4]:
                if st.button("✕", key=f"del_photo_{i}"):
                    st.session_state.photos.pop(i)
                    st.rerun()
    else:
        st.info("Belum ada foto. Pilih file lalu klik 'Tambahkan Foto ke Daftar'.")
    
    st.markdown('</div>', unsafe_allow_html=True)

# ═══════════════════════════════════════════════════════════════
# EXPORT BUTTON
# ═══════════════════════════════════════════════════════════════
st.markdown("---")
st.markdown("<h4 style='text-align:center'>📥 Download Laporan</h4>", unsafe_allow_html=True)

col_format, col_gen, col_dummy = st.columns([1, 2, 1])
with col_format:
    format_option = st.radio("Pilih format:", ["Microsoft Word (.docx)", "PDF (.pdf)"], horizontal=True, key="format_radio")
with col_gen:
    if st.button("📄 Generate & Download", type="primary", use_container_width=True):
        if not pic_project.strip():
            st.error("⚠️ PIC Project harus diisi.")
        elif not any(a["uraian"].strip() for a in full_data.get("activities", [])):
            st.error("⚠️ Minimal harus ada 1 kegiatan harian dengan uraian.")
        else:
            with st.spinner("Membuat laporan..."):
                data_report = {
                    "pic_project": pic_project,
                    "tanggal": tanggal,
                    "team_support": team_support,
                    "lokasi": lokasi,
                    "activities": [a for a in full_data.get("activities", []) if a["uraian"].strip()],
                    "incidents": [i for i in full_data.get("incidents", []) if i["masalah"].strip()],
                    "followups": [f for f in full_data.get("followups", []) if f["deskripsi"].strip()],
                    "materials": [m for m in full_data.get("materials", []) if m["nama"].strip()],
                    "catatan": full_data.get("catatan", ""),
                }
                try:
                    tgl_str = format_tanggal(tanggal)
                    if format_option == "Microsoft Word (.docx)":
                        output = build_report_docx(data_report, st.session_state.photos)
                        ext = ".docx"
                    else:
                        output = build_report_pdf(data_report, st.session_state.photos)
                        ext = ".pdf"
                    filename = f"Daily Report_Managed Service_Kominfo Mojokerto - {tgl_str}{ext}"
                    st.session_state.last_output = output.getvalue()
                    st.session_state.last_filename = filename
                    st.success("✅ Laporan siap di-download!")
                except Exception as e:
                    st.error(f"❌ Error: {e}")
                    import traceback
                    st.code(traceback.format_exc())

if "last_output" in st.session_state:
    col_l, col_c, col_r = st.columns([1,2,1])
    with col_c:
        st.download_button(
            "⬇️ Download",
            data=st.session_state.last_output,
            file_name=st.session_state.last_filename,
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document" if ".docx" in st.session_state.last_filename else "application/pdf",
            use_container_width=True,
        )
