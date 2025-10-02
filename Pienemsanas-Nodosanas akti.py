import sys
import os
import os.path

import json
import tempfile
from datetime import datetime
from decimal import Decimal, InvalidOperation
from PyQt5.QtWidgets import QAbstractItemView
from PySide6.QtGui import QColor
from dataclasses import dataclass, asdict, field
import shutil
import requests # Jauns imports

# Pārliecināties, ka šīs ir importētas no PySide6.QtWidgets
from PySide6.QtWidgets import (
    QApplication, QWidget, QMainWindow, QLabel, QLineEdit, QTextEdit, QPushButton,
    QFileDialog, QSpinBox, QDoubleSpinBox, QTableWidget, QTableWidgetItem, QHeaderView,
    QToolButton, QTabWidget, QFormLayout, QVBoxLayout, QHBoxLayout, QMessageBox, QCheckBox,
    QListWidget, QListWidgetItem, QGroupBox, QComboBox, QInputDialog, QSplitter, QScrollArea, QDateEdit, QAbstractItemView
)

from PIL import Image
from PIL.ImageQt import ImageQt # JAUNS IMPORTS
from PySide6.QtGui import QPainter # JAUNS IMPORTS

from pdf2image import convert_from_path
from PySide6.QtGui import QPixmap

from PySide6.QtCore import Qt, QSize, QSettings, QStandardPaths, QUrl
from PySide6.QtGui import QAction, QIcon
from PySide6.QtWidgets import (
    QApplication, QWidget, QMainWindow, QLabel, QLineEdit, QTextEdit, QPushButton,
    QFileDialog, QSpinBox, QDoubleSpinBox, QTableWidget, QTableWidgetItem, QHeaderView,
    QToolButton, QTabWidget, QFormLayout, QVBoxLayout, QHBoxLayout, QMessageBox, QCheckBox,
    QListWidget, QListWidgetItem, QGroupBox, QComboBox, QInputDialog, QSplitter, QScrollArea, QDateEdit, QAbstractItemView
)
from PySide6.QtPrintSupport import QPrinter, QPrintDialog, QPrintPreviewDialog # JAUNS IMPORTS



# Import for WebEngine
from PySide6.QtWebEngineWidgets import QWebEngineView
from PySide6.QtWebEngineCore import QWebEngineUrlRequestInterceptor # For intercepting URL changes

# ReportLab imports (unchanged)
from reportlab.lib.pagesizes import A4, landscape, portrait, letter, legal, A3, A5
from reportlab.pdfgen import canvas
from reportlab.lib import colors
from reportlab.lib.units import mm, inch
from reportlab.platypus import Table, TableStyle, Paragraph, Spacer, SimpleDocTemplate, Image as RLImage, PageBreak
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.platypus.flowables import Flowable

# python-docx imports (unchanged)
from docx import Document
from docx.shared import Inches, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_ALIGN_VERTICAL
from docx.text.run import Run
from PySide6.QtWebChannel import QWebChannel
from PySide6.QtCore import QObject, Slot

# ---------------------- Konstantes un direktoriju iestatījumi ----------------------
APP_DATA_DIR = QStandardPaths.writableLocation(QStandardPaths.AppDataLocation)
DOCUMENTS_DIR = QStandardPaths.writableLocation(QStandardPaths.DocumentsLocation)

SETTINGS_DIR = os.path.join(APP_DATA_DIR, "AktaGenerators")
HISTORY_FILE = os.path.join(SETTINGS_DIR, "history.json")
ADDRESS_BOOK_FILE = os.path.join(SETTINGS_DIR, "address_book.json")
DEFAULT_SETTINGS_FILE = os.path.join(SETTINGS_DIR, "default_settings.json")
TEXT_BLOCKS_FILE = os.path.join(SETTINGS_DIR, "text_blocks.json") # JAUNA RINDAS

# Jaunas noklusējuma saglabāšanas mapes
DEFAULT_OUTPUT_DIR = os.path.join(DOCUMENTS_DIR, "AktaGenerators_Output")
PROJECT_SAVE_DIR = os.path.join(APP_DATA_DIR, "AktaGenerators_Projects")
# TEMPLATES_DIR = os.path.join(APP_DATA_DIR, "AktaGenerators_Templates") # JAUNA RINDAS - Tagad dinamiski iestatīts AktaDati objektā


# Pārliecināmies, ka direktoriji eksistē
os.makedirs(SETTINGS_DIR, exist_ok=True)
os.makedirs(DEFAULT_OUTPUT_DIR, exist_ok=True)
os.makedirs(PROJECT_SAVE_DIR, exist_ok=True)
# os.makedirs(TEMPLATES_DIR, exist_ok=True) # JAUNA RINDAS - Tagad dinamiski iestatīts AktaDati objektā



# ---------------------- Datu modeļi ----------------------
# (unchanged)
@dataclass
class Persona:
    nosaukums: str = ""
    reģ_nr: str = ""
    adrese: str = ""
    kontaktpersona: str = ""
    tālrunis: str = ""
    epasts: str = ""
    bankas_konts: str = ""
    juridiskais_statuss: str = ""


class MapBridge(QObject):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.map_click_callback = None

    @Slot(str, str)
    def handleMapClick(self, lat, lon):
        if self.map_click_callback:
            self.map_click_callback(lat, lon)

class TextBlockManager:
    def __init__(self):
        self.text_blocks = self._load_text_blocks()

    def _load_text_blocks(self):
        if os.path.exists(TEXT_BLOCKS_FILE):
            try:
                with open(TEXT_BLOCKS_FILE, 'r', encoding='utf-8') as f:
                    return json.load(f)
            except Exception as e:
                print(f"Kļūda ielādējot teksta blokus: {e}")
                return {}
        return {}

    def _save_text_blocks(self):
        os.makedirs(SETTINGS_DIR, exist_ok=True)
        try:
            with open(TEXT_BLOCKS_FILE, 'w', encoding='utf-8') as f:
                json.dump(self.text_blocks, f, ensure_ascii=False, indent=2)
        except Exception as e:
            QMessageBox.critical(None, "Kļūda", f"Neizdevās saglabāt teksta blokus: {e}")

    def get_blocks_for_field(self, field_name):
        return self.text_blocks.get(field_name, {})

    def add_block(self, field_name, block_name, block_content):
        if field_name not in self.text_blocks:
            self.text_blocks[field_name] = {}
        self.text_blocks[field_name][block_name] = block_content
        self._save_text_blocks()

    def delete_block(self, field_name, block_name):
        if field_name in self.text_blocks and block_name in self.text_blocks[field_name]:
            del self.text_blocks[field_name][block_name]
            self._save_text_blocks()

    def get_block_content(self, field_name, block_name):
        return self.text_blocks.get(field_name, {}).get(block_name, "")


@dataclass
class Pozīcija:
    apraksts: str
    daudzums: Decimal
    vienība: str
    cena: Decimal
    seriālais_nr: str = ""
    garantija: str = ""
    piezīmes_pozīcijai: str = ""
    attēla_ceļš: str = ""

    @property
    def summa(self) -> Decimal:
        try:
            return (self.daudzums * self.cena).quantize(Decimal("0.01"))
        except Exception:
            return Decimal("0.00")

@dataclass
class Attēls:
    ceļš: str
    paraksts: str = ""

@dataclass
class AktaDati:
    akta_nr: str = ""
    datums: str = ""  # YYYY-MM-DD
    vieta: str = ""
    pasūtījuma_nr: str = ""
    pieņēmējs: Persona = field(default_factory=Persona)
    nodevējs: Persona = field(default_factory=Persona)
    pozīcijas: list = field(default_factory=list)  # list[Pozīcija]
    attēli: list = field(default_factory=list)
    piezīmes: str = ""
    iekļaut_pvn: bool = False
    pvn_likme: Decimal = Decimal("21.0")
    parakstu_rindas: bool = True
    logotipa_ceļš: str = ""
    fonts_ceļš: str = ""  # TTF/OTF
    paraksts_pieņēmējs_ceļš: str = ""
    paraksts_nodevējs_ceļš: str = ""
    līguma_nr: str = ""
    izpildes_termiņš: str = ""
    pieņemšanas_datums: str = ""
    nodošanas_datums: str = ""
    strīdu_risināšana: str = ""
    konfidencialitātes_klauzula: bool = False
    soda_nauda_procenti: Decimal = Decimal("0.0")
    piegādes_nosacījumi: str = ""
    apdrošināšana: bool = False
    papildu_nosacījumi: str = ""
    atsauces_dokumenti: str = ""
    akta_statuss: str = "Melnraksts"
    valūta: str = "EUR"
    elektroniskais_paraksts: bool = False
    radit_elektronisko_parakstu_tekstu: bool = False # JAUNS LAUKS

    # Jauni iestatījumu lauki
    pdf_page_size: str = "A4"
    pdf_page_orientation: str = "Portrets"
    pdf_margin_left: Decimal = Decimal("18")
    pdf_margin_right: Decimal = Decimal("18")
    pdf_margin_top: Decimal = Decimal("16")
    pdf_margin_bottom: Decimal = Decimal("16")
    pdf_font_size_head: int = 14
    pdf_font_size_normal: int = 10
    pdf_font_size_small: int = 9
    pdf_font_size_table: int = 9
    pdf_logo_width_mm: Decimal = Decimal("35")
    pdf_signature_width_mm: Decimal = Decimal("50")
    pdf_signature_height_mm: Decimal = Decimal("20")
    docx_image_width_inches: Decimal = Decimal("4")
    docx_signature_width_inches: Decimal = Decimal("1.5")
    table_col_widths: str = "10,40,18,18,20,20,25,25,25" # Komatiem atdalīti platumi mm
    auto_generate_akta_nr: bool = False
    default_currency: str = "EUR"
    default_unit: str = "gab."
    default_pvn_rate: Decimal = Decimal("21.0")
    poppler_path: str = ""
    # Jauni iestatījumi, lai sasniegtu "30+"
    header_text_color: str = "#000000" # Hex krāsa
    footer_text_color: str = "#000000"
    table_header_bg_color: str = "#E0E0E0"
    table_grid_color: str = "#CCCCCC"
    table_row_spacing: Decimal = Decimal("4") # mm
    line_spacing_multiplier: Decimal = Decimal("1.2") # Reizinātājs fonta izmēram
    show_page_numbers: bool = True
    show_generation_timestamp: bool = True
    currency_symbol_position: str = "after" # "before" or "after"
    date_format: str = "YYYY-MM-DD"
    signature_line_length_mm: Decimal = Decimal("60")
    signature_line_thickness_pt: Decimal = Decimal("0.5")
    add_cover_page: bool = False
    cover_page_title: str = "Pieņemšanas-Nodošanas Akts"
    cover_page_logo_width_mm: Decimal = Decimal("80")
    # Individuālais QR kods
    include_custom_qr_code: bool = False
    custom_qr_code_data: str = ""
    custom_qr_code_size_mm: Decimal = Decimal("20")
    custom_qr_code_position: str = "bottom_right"
    custom_qr_code_pos_x_mm: Decimal = Decimal("0")  # Custom X pozīcija mm
    custom_qr_code_pos_y_mm: Decimal = Decimal("0")  # Custom Y pozīcija mm
    custom_qr_code_color: str = "#000000"  # QR koda krāsa (Hex)
    # Automātiskais QR kods (akta ID)
    include_auto_qr_code: bool = False
    auto_qr_code_size_mm: Decimal = Decimal("20")
    auto_qr_code_position: str = "bottom_left"
    auto_qr_code_pos_x_mm: Decimal = Decimal("0")  # Custom X pozīcija mm
    auto_qr_code_pos_y_mm: Decimal = Decimal("0")  # Custom Y pozīcija mm
    auto_qr_code_color: str = "#000000"  # QR koda krāsa (Hex)

    add_watermark: bool = False
    watermark_text: str = "MELNRAKSTS"
    watermark_font_size: int = 72
    watermark_color: str = "#E0E0E0"
    watermark_rotation: int = 45
    enable_pdf_encryption: bool = False
    pdf_user_password: str = ""
    pdf_owner_password: str = ""
    allow_printing: bool = True
    allow_copying: bool = True
    allow_modifying: bool = False
    allow_annotating: bool = True
    # Papildu lauki, lai sasniegtu 30+
    default_country: str = "Latvija"
    default_city: str = "Rīga"
    show_contact_details_in_header: bool = False
    contact_details_header_font_size: int = 8
    item_image_width_mm: Decimal = Decimal("50") # Platums attēliem pie pozīcijām
    item_image_caption_font_size: int = 8
    show_item_notes_in_table: bool = True
    show_item_serial_number_in_table: bool = True
    show_item_warranty_in_table: bool = True
    table_cell_padding_mm: Decimal = Decimal("2")
    table_header_font_style: str = "bold" # "bold", "italic", "normal"
    table_content_alignment: str = "left" # "left", "center", "right"
    signature_font_size: int = 9
    signature_spacing_mm: Decimal = Decimal("10") # Atstarpe starp paraksta rindu un vārdu
    # Vēl daži, lai pārsniegtu 30
    document_title_font_size: int = 18
    document_title_color: str = "#000000"
    section_heading_font_size: int = 12
    section_heading_color: str = "#000000"
    paragraph_line_spacing_multiplier: Decimal = Decimal("1.2")
    table_border_style: str = "solid" # "solid", "dashed", "none"
    table_border_thickness_pt: Decimal = Decimal("0.5")
    table_alternate_row_color: str = "" # Hex krāsa, piem. "#F0F0F0"
    # Papildu lauki, lai sasniegtu 30+
    show_total_sum_in_words: bool = False
    total_sum_in_words_language: str = "lv" # "lv", "en"
    # Papildu lauki, lai sasniegtu 30+
    default_vat_calculation_method: str = "exclusive" # "exclusive", "inclusive"
    show_vat_breakdown: bool = True
    # Papildu lauki, lai sasniegtu 30+
    enable_digital_signature_field: bool = False # PDF digitālā paraksta lauks
    digital_signature_field_name: str = "Paraksts"
    digital_signature_field_size_mm: Decimal = Decimal("40")
    digital_signature_field_position: str = "bottom_center"  # "bottom_left", "bottom_right", "top_left", "top_right", "bottom_center"
    template_password: str = ""  # Parole šablonam
    templates_dir: str = ""  # JAUNA RINDAS - Šablonu saglabāšanas direktorijs

    def kopējā_summma(self):
        s = sum((p.summa for p in self.pozīcijas), Decimal("0.00"))
        return s.quantize(Decimal("0.01"))

    def pvn_summa(self):
        if not self.iekļaut_pvn:
            return Decimal("0.00")
        pvn = self.kopējā_summma() * (self.pvn_likme / Decimal("100"))
        return pvn.quantize(Decimal("0.01"))

    def summa_ar_pvn(self):
        return (self.kopējā_summma() + self.pvn_summa()).quantize(Decimal("0.01"))

# ---------------------- Palīgfunkcijas ----------------------
# (unchanged)
def to_decimal(val) -> Decimal:
    if isinstance(val, Decimal):
        return val
    try:
        s = str(val).replace(" ", "").replace(",", ".")
        return Decimal(s)
    except (InvalidOperation, ValueError):
        return Decimal("0")

def formēt_naudu(d: Decimal) -> str:
    return f"{d:.2f}"

def drošs_faila_nosaukums(s: str) -> str:
    bad = '\\/:*?"<>|'
    for ch in bad:
        s = s.replace(ch, "_")
    return s.strip() or "akts"

def reģistrēt_fontu(font_ceļš: str, vārds: str = "DokFont") -> str:
    if not font_ceļš or not os.path.exists(font_ceļš):
        try:
            if sys.platform == "win32":
                system_font_path = os.path.join(os.environ.get("WINDIR", "C:\\Windows"), "Fonts", "Arial.ttf")
                if os.path.exists(system_font_path):
                    pdfmetrics.registerFont(TTFont(vārds, system_font_path))
                    return vārds
            return "Helvetica"
        except Exception:
            return "Helvetica"
    try:
        pdfmetrics.registerFont(TTFont(vārds, font_ceļš))
        return vārds
    except Exception:
        return "Helvetica"

def render_pdf_to_image(pdf_path: str, poppler_path: str = None) -> QPixmap:
    """
    Renderē PDF faila pirmo lapu kā QPixmap.
    Nepieciešama Poppler instalācija.
    :param pdf_path: Ceļš uz PDF failu.
    :param poppler_path: (Tikai Windows) Ceļš uz Poppler bin direktoriju.
    :return: QPixmap objekts ar PDF lapas attēlu.
    """
    try:
        images = convert_from_path(pdf_path, first_page=1, last_page=1, poppler_path=poppler_path)
        if images:
            from io import BytesIO
            img_byte_arr = BytesIO()
            images[0].save(img_byte_arr, format='PNG')
            img_byte_arr.seek(0)

            pixmap = QPixmap()
            pixmap.loadFromData(img_byte_arr.getvalue(), 'PNG')
            return pixmap
        else:
            return QPixmap()
    except Exception as e:
        print(f"Kļūda renderējot PDF uz attēlu: {e}")
        QMessageBox.warning(None, "PDF renderēšanas kļūda",
                            f"Neizdevās renderēt PDF priekšskatījumu. Pārliecinieties, ka Poppler ir instalēts un pieejams sistēmas PATH (vai norādīts poppler_path).\nKļūda: {e}")
        return QPixmap()

# Lapu numerācija PDF dokumentam
class PageNumCanvas(canvas.Canvas):
    def __init__(self, *args, **kwargs):
        canvas.Canvas.__init__(self, *args, **kwargs)
        self.pages = []
        self.show_page_numbers = kwargs.pop('show_page_numbers', True) # Default value, will be set by generator

    def showPage(self):
        self.pages.append(dict(self.__dict__))
        self._startPage()

    def save(self):
        page_count = len(self.pages)
        for page in self.pages:
            self.__dict__.update(page)
            if self.show_page_numbers:
                self.draw_page_number(page_count)
            canvas.Canvas.showPage(self)
        canvas.Canvas.save(self)

    def draw_page_number(self, page_count):
        self.setFont('Helvetica', 9)
        self.drawRightString(self._pagesize[0] - 18*mm, 10*mm, f"Lapa {self._pageNumber} no {page_count}")

# ---------------------- PDF ģenerēšana ----------------------
# (Modified to include new settings)
def ģenerēt_pdf(akta_dati: AktaDati, pdf_ceļš: str = None):
    font_name = reģistrēt_fontu(akta_dati.fonts_ceļš)

    styles = getSampleStyleSheet()
    # Ensure all Decimal values are converted to float when used with ReportLab's float-based units or font sizes
    # Uzlaboti stili ar jaunajiem iestatījumiem
    styles.add(ParagraphStyle(name='LatvHead', fontName=font_name, fontSize=float(akta_dati.pdf_font_size_head), leading=float(akta_dati.pdf_font_size_head) * float(akta_dati.line_spacing_multiplier), spaceAfter=8, textColor=colors.HexColor(akta_dati.header_text_color)))
    styles.add(ParagraphStyle(name='Latv', fontName=font_name, fontSize=float(akta_dati.pdf_font_size_normal), leading=float(akta_dati.pdf_font_size_normal) * float(akta_dati.line_spacing_multiplier), textColor=colors.HexColor(akta_dati.footer_text_color)))
    styles.add(ParagraphStyle(name='LatvSmall', fontName=font_name, fontSize=float(akta_dati.pdf_font_size_small), leading=float(akta_dati.pdf_font_size_small) * float(akta_dati.line_spacing_multiplier), textColor=colors.HexColor(akta_dati.footer_text_color)))
    styles.add(ParagraphStyle(name='LatvTableContent', fontName=font_name, fontSize=float(akta_dati.pdf_font_size_table), leading=float(akta_dati.pdf_font_size_table) * float(akta_dati.line_spacing_multiplier), wordWrap='CJK', alignment={'left': 0, 'center': 1, 'right': 2}.get(akta_dati.table_content_alignment, 0)))
    styles.add(ParagraphStyle(name='LatvElectronicSignature', fontName=font_name, fontSize=float(akta_dati.pdf_font_size_normal), leading=float(akta_dati.pdf_font_size_normal) * float(akta_dati.line_spacing_multiplier), alignment=1, textColor=colors.HexColor(akta_dati.header_text_color)))
    styles.add(ParagraphStyle(name='DocTitle', fontName=font_name, fontSize=float(akta_dati.document_title_font_size), alignment=1, textColor=colors.HexColor(akta_dati.document_title_color)))
    styles.add(ParagraphStyle(name='SectionHeading', fontName=font_name, fontSize=float(akta_dati.section_heading_font_size), leading=float(akta_dati.section_heading_font_size) * float(akta_dati.paragraph_line_spacing_multiplier), textColor=colors.HexColor(akta_dati.section_heading_color)))

    page_size_map = {
        "A4": A4, "Letter": letter, "Legal": legal, "A3": A3, "A5": A5
    }
    base_page_size = page_size_map.get(akta_dati.pdf_page_size, A4)

    if akta_dati.pdf_page_orientation == "Ainava":
        pagesize = landscape(base_page_size)
    else:
        pagesize = portrait(base_page_size)

    if pdf_ceļš is None:
        temp_file = tempfile.NamedTemporaryFile(delete=False, suffix=".pdf")
        pdf_ceļš = temp_file.name
        temp_file.close()

    doc = SimpleDocTemplate(
        pdf_ceļš,
        pagesize=pagesize,
        leftMargin=float(akta_dati.pdf_margin_left) * mm,
        rightMargin=float(akta_dati.pdf_margin_right) * mm,
        topMargin=float(akta_dati.pdf_margin_top) * mm,
        bottomMargin=float(akta_dati.pdf_margin_bottom) * mm,
        title="Pieņemšanas–Nodošanas akts",
    )

    story = []

    # Cover Page (new feature)
    if akta_dati.add_cover_page:
        story.append(Spacer(1, 2 * inch))
        if akta_dati.logotipa_ceļš and os.path.exists(akta_dati.logotipa_ceļš):
            try:
                cover_logo = RLImage(akta_dati.logotipa_ceļš)
                cover_logo._restrictSize(float(akta_dati.cover_page_logo_width_mm) * mm, 50 * mm)
                story.append(cover_logo)
                story.append(Spacer(1, 0.5 * inch))
            except Exception:
                pass
        story.append(Paragraph(akta_dati.cover_page_title, styles['DocTitle']))
        story.append(Spacer(1, 1 * inch))
        story.append(Paragraph(f"Akta Nr.: {akta_dati.akta_nr}", styles['Latv']))
        story.append(Paragraph(f"Datums: {akta_dati.datums}", styles['Latv']))
        story.append(Paragraph(f"Vieta: {akta_dati.vieta}", styles['Latv']))
        story.append(Spacer(1, 2 * inch))
        story.append(Paragraph(f"Pieņēmējs: {akta_dati.pieņēmējs.nosaukums}", styles['Latv']))
        story.append(Paragraph(f"Nodevējs: {akta_dati.nodevējs.nosaukums}", styles['Latv']))
        story.append(PageBreak())

    # Header ar logo un nosaukumu
    header_table_data = []
    logo_w = float(akta_dati.pdf_logo_width_mm) * mm
    if akta_dati.logotipa_ceļš and os.path.exists(akta_dati.logotipa_ceļš):
        try:
            header_logo = RLImage(akta_dati.logotipa_ceļš)
            header_logo._restrictSize(logo_w, 20 * mm)
            header_table_data.append([header_logo, Paragraph(f"<font name='{font_name}'><b>PIEŅEMŠANAS–NODOŠANAS AKTS</b></font>", styles['LatvHead'])])
        except Exception:
            header_table_data.append(["", Paragraph(f"<font name='{font_name}'><b>PIEŅEMŠANAS–NODOŠANAS AKTS</b></font>", styles['LatvHead'])])
    else:
        header_table_data.append(["", Paragraph(f"<font name='{font_name}'><b>PIEŅEMŠANAS–NODOŠANAS AKTS</b></font>", styles['LatvHead'])])

    ht = Table(header_table_data, colWidths=[logo_w, None])
    ht.setStyle(TableStyle([
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('ALIGN', (1, 0), (1, 0), 'LEFT'),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 6),
    ]))
    story.append(ht)

    # Aktu metadati (Nr., Datums, Vieta)
    md = [[Paragraph(f"<font name='{font_name}'><b>Akta Nr.:</b> {akta_dati.akta_nr}</font>", styles['Latv']),
           Paragraph(f"<font name='{font_name}'><b>Datums:</b> {datetime.strptime(akta_dati.datums, '%Y-%m-%d').strftime(akta_dati.date_format.replace('YYYY', '%Y').replace('MM', '%m').replace('DD', '%d'))}</font>", styles['Latv']),
           Paragraph(f"<font name='{font_name}'><b>Vieta:</b> {akta_dati.vieta}</font>", styles['Latv'])]]
    if akta_dati.pasūtījuma_nr:
        md[0].append(Paragraph(f"<font name='{font_name}'><b>Pasūtījuma Nr.:</b> {akta_dati.pasūtījuma_nr}</font>", styles['Latv']))

    meta = Table(md, colWidths=[55 * mm, 35 * mm, 40 * mm, None])
    meta.setStyle(TableStyle([
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 6),
    ]))
    story.append(meta)
    story.append(Spacer(1, 6))

    # Jauni metadati
    if akta_dati.līguma_nr:
        story.append(Paragraph(f"<font name='{font_name}'><b>Līguma Nr.:</b> {akta_dati.līguma_nr}</font>", styles['Latv']))
    if akta_dati.izpildes_termiņš:
        story.append(Paragraph(f"<font name='{font_name}'><b>Izpildes termiņš:</b> {datetime.strptime(akta_dati.izpildes_termiņš, '%Y-%m-%d').strftime(akta_dati.date_format.replace('YYYY', '%Y').replace('MM', '%m').replace('DD', '%d'))}</font>", styles['Latv']))
    if akta_dati.pieņemšanas_datums:
        story.append(Paragraph(f"<font name='{font_name}'><b>Pieņemšanas datums:</b> {datetime.strptime(akta_dati.pieņemšanas_datums, '%Y-%m-%d').strftime(akta_dati.date_format.replace('YYYY', '%Y').replace('MM', '%m').replace('DD', '%d'))}</font>", styles['Latv']))
    if akta_dati.nodošanas_datums:
        story.append(Paragraph(f"<font name='{font_name}'><b>Nodošanas datums:</b> {datetime.strptime(akta_dati.nodošanas_datums, '%Y-%m-%d').strftime(akta_dati.date_format.replace('YYYY', '%Y').replace('MM', '%m').replace('DD', '%d'))}</font>", styles['Latv']))
    story.append(Spacer(1, 6))

    # Puses: Pieņēmējs / Nodevējs
    def persona_paragraph(prefix: str, p: Persona):
        lines = [f"<b>{prefix}:</b> {p.nosaukums}"]
        if p.reģ_nr: lines.append(f"Reģ. Nr.: {p.reģ_nr}")
        if p.adrese: lines.append(f"Adrese: {p.adrese}")
        if p.kontaktpersona: lines.append(f"Kontaktpersona: {p.kontaktpersona}")
        if p.tālrunis: lines.append(f"Tālrunis: {p.tālrunis}")
        if p.epasts: lines.append(f"E-pasts: {p.epasts}")
        if p.bankas_konts: lines.append(f"Bankas konts: {p.bankas_konts}")
        if p.juridiskais_statuss: lines.append(f"Statuss: {p.juridiskais_statuss}")
        return Paragraph(f"<font name='{font_name}'>" + "<br/>".join(lines) + "</font>", styles['Latv'])

    available_width = pagesize[0] - float(akta_dati.pdf_margin_left) * mm - float(akta_dati.pdf_margin_right) * mm
    col_width_parties = available_width / 2

    puses = Table([
        [persona_paragraph("Pieņēmējs", akta_dati.pieņēmējs), persona_paragraph("Nodevējs", akta_dati.nodevējs)]
    ], colWidths=[col_width_parties, col_width_parties])
    pstyle = TableStyle([
        ('VALIGN', (0, 0), (-1, -1), 'TOP'),
        ('BOX', (0, 0), (-1, -1), float(akta_dati.table_border_thickness_pt), colors.HexColor(akta_dati.table_grid_color)),
        ('INNERGRID', (0, 0), (-1, -1), float(akta_dati.table_border_thickness_pt), colors.HexColor(akta_dati.table_grid_color)),
        ('LEFTPADDING', (0, 0), (-1, -1), float(akta_dati.table_cell_padding_mm)),
        ('RIGHTPADDING', (0, 0), (-1, -1), float(akta_dati.table_cell_padding_mm)),
        ('TOPPADDING', (0, 0), (-1, -1), float(akta_dati.table_cell_padding_mm)),
        ('BOTTOMPADDING', (0, 0), (-1, -1), float(akta_dati.table_cell_padding_mm)),
    ])
    puses.setStyle(pstyle)
    story.append(puses)
    story.append(Spacer(1, 8))

    # Pozīciju tabula
    tab_header_items = [
        Paragraph(f"<font name='{font_name}'><b>Nr.</b></font>", styles['LatvTableContent']),
        Paragraph(f"<font name='{font_name}'><b>Apraksts</b></font>", styles['LatvTableContent']),
        Paragraph(f"<font name='{font_name}'><b>Daudzums</b></font>", styles['LatvTableContent']),
        Paragraph(f"<font name='{font_name}'><b>Vienība</b></font>", styles['LatvTableContent']),
        Paragraph(f"<font name='{font_name}'><b>Cena</b></font>", styles['LatvTableContent']),
        Paragraph(f"<font name='{font_name}'><b>Summa</b></font>", styles['LatvTableContent']),
    ]
    if akta_dati.show_item_serial_number_in_table:
        tab_header_items.append(Paragraph(f"<font name='{font_name}'><b>Seriālais Nr.</b></font>", styles['LatvTableContent']))
    if akta_dati.show_item_warranty_in_table:
        tab_header_items.append(Paragraph(f"<font name='{font_name}'><b>Garantija</b></font>", styles['LatvTableContent']))
    if akta_dati.show_item_notes_in_table:
        tab_header_items.append(Paragraph(f"<font name='{font_name}'><b>Piezīmes pozīcijai</b></font>", styles['LatvTableContent']))

    tab_data = [tab_header_items]
    for i, poz in enumerate(akta_dati.pozīcijas, start=1):
        row_items = [
            Paragraph(str(i), styles['LatvTableContent']),
            Paragraph(poz.apraksts, styles['LatvTableContent']),
            Paragraph(f"{formēt_naudu(poz.daudzums)}", styles['LatvTableContent']),
            Paragraph(poz.vienība, styles['LatvTableContent']),
            Paragraph(f"{akta_dati.currency_symbol_position == 'before' and akta_dati.valūta or ''}{formēt_naudu(poz.cena)}{akta_dati.currency_symbol_position == 'after' and ' ' + akta_dati.valūta or ''}", styles['LatvTableContent']),
            Paragraph(f"{akta_dati.currency_symbol_position == 'before' and akta_dati.valūta or ''}{formēt_naudu(poz.summa)}{akta_dati.currency_symbol_position == 'after' and ' ' + akta_dati.valūta or ''}", styles['LatvTableContent']),
        ]
        if akta_dati.show_item_serial_number_in_table:
            row_items.append(Paragraph(poz.seriālais_nr, styles['LatvTableContent']))
        if akta_dati.show_item_warranty_in_table:
            row_items.append(Paragraph(poz.garantija, styles['LatvTableContent']))
        if akta_dati.show_item_notes_in_table:
            row_items.append(Paragraph(poz.piezīmes_pozīcijai, styles['LatvTableContent']))
        tab_data.append(row_items)

    try:
        col_widths_mm = [float(x.strip()) for x in akta_dati.table_col_widths.split(',')]
        # Adjust expected column count based on visibility settings
        expected_cols = 6 # Base columns
        if akta_dati.show_item_serial_number_in_table: expected_cols += 1
        if akta_dati.show_item_warranty_in_table: expected_cols += 1
        if akta_dati.show_item_notes_in_table: expected_cols += 1

        if len(col_widths_mm) != expected_cols:
            raise ValueError(f"Nepareizs kolonnu skaits iestatījumos. Gaidīts: {expected_cols}, atrasts: {len(col_widths_mm)}")
        col_widths = [w * mm for w in col_widths_mm]
    except Exception:
        # Fallback to default widths, adjusted for visibility
        default_widths = [10, 40, 18, 18, 20, 20] # Nr, Apraksts, Daudzums, Vienība, Cena, Summa
        if akta_dati.show_item_serial_number_in_table: default_widths.append(25)
        if akta_dati.show_item_warranty_in_table: default_widths.append(25)
        if akta_dati.show_item_notes_in_table: default_widths.append(25)

        col_widths = [w * mm for w in default_widths]
        total_default_width = sum(col_widths)
        if total_default_width != available_width:
            scale_factor = available_width / total_default_width
            col_widths = [w * scale_factor for w in col_widths]


    t = Table(tab_data, colWidths=col_widths)
    tstyle = TableStyle([
        ('FONTNAME', (0,0), (-1,-1), font_name),
        ('FONTSIZE', (0,0), (-1,0), akta_dati.pdf_font_size_table + 1),
        ('FONTSIZE', (0,1), (-1,-1), akta_dati.pdf_font_size_table),
        ('BACKGROUND', (0,0), (-1,0), colors.HexColor(akta_dati.table_header_bg_color)),
        ('ALIGN', (0,0), (-1,0), 'CENTER'),
        ('ALIGN', (0,1), (1,-1), 'LEFT'), # Nr, Apraksts
        ('ALIGN', (2,1), (3,-1), 'CENTER'), # Daudzums, Vienība
        ('ALIGN', (4,1), (5,-1), 'RIGHT'), # Cena, Summa
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('GRID', (0,0), (-1,-1), float(akta_dati.table_border_thickness_pt), colors.HexColor(akta_dati.table_grid_color)),
        ('BOTTOMPADDING', (0,0), (-1,0), float(akta_dati.table_cell_padding_mm) * 2),
        ('TOPPADDING', (0,0), (-1,0), float(akta_dati.table_cell_padding_mm) * 2),
        ('BOTTOMPADDING', (0,1), (-1,-1), float(akta_dati.table_cell_padding_mm)),
        ('TOPPADDING', (0,1), (-1,-1), float(akta_dati.table_cell_padding_mm)),
    ])
    # Apply alternate row color
    if akta_dati.table_alternate_row_color:
        for i in range(1, len(tab_data)):
            if i % 2 == 0: # Even rows (0-indexed, so actual even rows)
                tstyle.add('BACKGROUND', (0, i), (-1, i), colors.HexColor(akta_dati.table_alternate_row_color))

    t.setStyle(tstyle)
    story.append(t)

    # Kopsavilkums
    story.append(Spacer(1, 6))
    summa_tab = []
    summa_tab.append([Paragraph(f"<font name='{font_name}'>Kopā bez PVN:</font>", styles['Latv']), Paragraph(f"<font name='{font_name}'>{akta_dati.currency_symbol_position == 'before' and akta_dati.valūta or ''}{formēt_naudu(akta_dati.kopējā_summma())}{akta_dati.currency_symbol_position == 'after' and ' ' + akta_dati.valūta or ''}</font>", styles['Latv'])])
    if akta_dati.iekļaut_pvn and akta_dati.show_vat_breakdown:
        summa_tab.append([Paragraph(f"<font name='{font_name}'>PVN {akta_dati.pvn_likme}%:</font>", styles['Latv']), Paragraph(f"<font name='{font_name}'>{akta_dati.currency_symbol_position == 'before' and akta_dati.valūta or ''}{formēt_naudu(akta_dati.pvn_summa())}{akta_dati.currency_symbol_position == 'after' and ' ' + akta_dati.valūta or ''}</font>", styles['Latv'])])
        summa_tab.append([Paragraph(f"<font name='{font_name}'>Kopā ar PVN:</font>", styles['Latv']), Paragraph(f"<font name='{font_name}'>{akta_dati.currency_symbol_position == 'before' and akta_dati.valūta or ''}{formēt_naudu(akta_dati.summa_ar_pvn())}{akta_dati.currency_symbol_position == 'after' and ' ' + akta_dati.valūta or ''}</font>", styles['Latv'])])

    if summa_tab:
        ts = Table(summa_tab, colWidths=[None, 40*mm])
        ts.setStyle(TableStyle([
            ('ALIGN', (1,0), (1,-1), 'RIGHT'),
            ('FONTNAME', (0,0), (-1,-1), font_name),
            ('FONTSIZE', (0,0), (-1,-1), akta_dati.pdf_font_size_normal),
        ]))
        story.append(ts)

    # Piezīmes
    if akta_dati.piezīmes:
        story.append(Spacer(1, 6))
        story.append(Paragraph(f"<font name='{font_name}'><b>Vispārīgās piezīmes:</b><br/>{akta_dati.piezīmes}</font>", styles['Latv']))

    # Jauni juridiski saistoši lauki PDF dokumentā
    if akta_dati.strīdu_risināšana:
        story.append(Spacer(1, 6))
        story.append(Paragraph(f"<font name='{font_name}'><b>Strīdu risināšanas kārtība:</b><br/>{akta_dati.strīdu_risināšana}</font>", styles['Latv']))

    if akta_dati.konfidencialitātes_klauzula:
        story.append(Spacer(1, 6))
        story.append(Paragraph(f"<font name='{font_name}'><b>Konfidencialitātes klauzula:</b><br/>Puses apņemas neizpaust trešajām personām informāciju, kas iegūta šī akta ietvaros, izņemot gadījumus, ko nosaka normatīvie akti.</font>", styles['Latv']))

    if akta_dati.soda_nauda_procenti > 0:
        story.append(Spacer(1, 6))
        story.append(Paragraph(f"<font name='{font_name}'><b>Soda nauda:</b> Par saistību neizpildi vai nepienācīgu izpildi, vainīgā puse maksā otrai pusei soda naudu {formēt_naudu(akta_dati.soda_nauda_procenti)}% apmērā no neizpildīto saistību vērtības par katru kavējuma dienu.</font>", styles['Latv']))

    if akta_dati.piegādes_nosacījumi:
        story.append(Spacer(1, 6))
        story.append(Paragraph(f"<font name='{font_name}'><b>Piegādes nosacījumi:</b> {akta_dati.piegādes_nosacījumi}</font>", styles['Latv']))

    if akta_dati.apdrošināšana:
        story.append(Spacer(1, 6))
        story.append(Paragraph(f"<font name='{font_name}'><b>Apdrošināšana:</b> Preces ir apdrošinātas pret bojājumiem un zaudējumiem līdz pieņemšanas-nodošanas brīdim.</font>", styles['Latv']))

    if akta_dati.papildu_nosacījumi:
        story.append(Spacer(1, 6))
        story.append(Paragraph(f"<font name='{font_name}'><b>Papildu nosacījumi:</b><br/>{akta_dati.papildu_nosacījumi}</font>", styles['Latv']))

    if akta_dati.atsauces_dokumenti:
        story.append(Spacer(1, 6))
        story.append(Paragraph(f"<font name='{font_name}'><b>Atsauces dokumenti:</b> {akta_dati.atsauces_dokumenti}</font>", styles['Latv']))

    story.append(Spacer(1, 6))
    story.append(Paragraph(f"<font name='{font_name}'><b>Akta statuss:</b> {akta_dati.akta_statuss}</font>", styles['Latv']))

    # Attēli (ja ir) – mērogojam līdz platumam
    if akta_dati.attēli:
        story.append(PageBreak())
        story.append(Spacer(1, 8))
        story.append(Paragraph(f"<font name='{font_name}'><b>Pievienotās fotogrāfijas</b></font>", styles['Latv']))
        for att in akta_dati.attēli:
            if os.path.exists(att.ceļš):
                try:
                    img = RLImage(att.ceļš)
                    img._restrictSize(available_width, float(akta_dati.item_image_width_mm) * 3 * mm)
                    story.append(Spacer(1, 4))
                    story.append(img)
                    if att.paraksts:
                        story.append(Paragraph(f"<font name='{font_name}'>{att.paraksts}</font>", styles['LatvSmall']))
                except Exception:
                    pass

    if akta_dati.elektroniskais_paraksts:

        if akta_dati.radit_elektronisko_parakstu_tekstu:
            story.append(Spacer(1, 16))

            story.append(Paragraph(

                f"<font name='{font_name}' size='{akta_dati.pdf_font_size_normal}'><b>ŠIS DOKUMENTS PARAKSTĪTS AR DROŠU ELEKTRONISKO PARAKSTU UN SATUR LAIKA ZĪMOGU</b></font>",

                styles['LatvElectronicSignature']

            ))

            story.append(Spacer(1, 16))

    elif akta_dati.parakstu_rindas:

        story.append(Spacer(1, 16))  # Atstarpe pirms parakstu laukiem

        # Izveidojam atsevišķus elementus katrai paraksta pusei

        pie_elements = []

        nod_elements = []

        # Pieņēmēja paraksta informācija

        pie_elements.append(Paragraph(f"<font name='{font_name}'><b>Pieņēmējs</b></font>", styles['Latv']))

        if akta_dati.paraksts_pieņēmējs_ceļš and os.path.exists(akta_dati.paraksts_pieņēmējs_ceļš):

            try:

                img_pie = RLImage(akta_dati.paraksts_pieņēmējs_ceļš)

                img_pie._restrictSize(float(akta_dati.pdf_signature_width_mm) * mm,

                                      float(akta_dati.pdf_signature_height_mm) * mm)

                pie_elements.append(img_pie)

            except Exception:

                pie_elements.append(Paragraph(

                    f"<font name='{font_name}'>_{'_' * int(float(akta_dati.signature_line_length_mm) / 2)}</font>",

                    styles['Latv']))

        else:

            pie_elements.append(Paragraph(

                f"<font name='{font_name}'>_{'_' * int(float(akta_dati.signature_line_length_mm) / 2)}</font>",

                styles['Latv']))

        pie_elements.append(Paragraph(

            f"<font name='{font_name}' size='{akta_dati.signature_font_size}'>{akta_dati.pieņēmējs.kontaktpersona or akta_dati.pieņēmējs.nosaukums}</font>",

            styles['LatvSmall']))

        # Nodevēja paraksta informācija

        nod_elements.append(Paragraph(f"<font name='{font_name}'><b>Nodevējs</b></font>", styles['Latv']))

        if akta_dati.paraksts_nodevējs_ceļš and os.path.exists(akta_dati.paraksts_nodevējs_ceļš):

            try:

                img_nod = RLImage(akta_dati.paraksts_nodevējs_ceļš)

                img_nod._restrictSize(float(akta_dati.pdf_signature_width_mm) * mm,

                                      float(akta_dati.pdf_signature_height_mm) * mm)

                nod_elements.append(img_nod)

            except Exception:

                nod_elements.append(Paragraph(

                    f"<font name='{font_name}'>_{'_' * int(float(akta_dati.signature_line_length_mm) / 2)}</font>",

                    styles['Latv']))

        else:

            nod_elements.append(Paragraph(

                f"<font name='{font_name}'>_{'_' * int(float(akta_dati.signature_line_length_mm) / 2)}</font>",

                styles['Latv']))

        nod_elements.append(Paragraph(

            f"<font name='{font_name}' size='{akta_dati.signature_font_size}'>{akta_dati.nodevējs.kontaktpersona or akta_dati.nodevējs.nosaukums}</font>",

            styles['LatvSmall']))

        # Izveidojam tabulu ar vienu rindu un divām kolonnām

        # Katrā kolonnā ievietojam Flowables (Paragraph, Image) sarakstu

        paraksti_table_data = [

            [pie_elements, nod_elements]

        ]

        paraksti = Table(paraksti_table_data, colWidths=[col_width_parties, col_width_parties])

        paraksti.setStyle(TableStyle([

            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),  # Centra izlīdzināšana visiem elementiem šūnās

            ('VALIGN', (0, 0), (-1, -1), 'TOP'),  # Vertikālā izlīdzināšana uz augšu

            ('FONTNAME', (0, 0), (-1, -1), font_name),

            ('BOTTOMPADDING', (0, 0), (-1, 0), float(akta_dati.signature_spacing_mm)),  # Atstarpe zem paraksta līnijas

        ]))

        story.append(paraksti)

    # QR kodi (jauna funkcionalitāte)
    qr_codes_to_add = []

    # Individuālais QR kods
    if akta_dati.include_custom_qr_code and akta_dati.custom_qr_code_data:
        qr_codes_to_add.append({
            'data': akta_dati.custom_qr_code_data,
            'size': akta_dati.custom_qr_code_size_mm,
            'position': akta_dati.custom_qr_code_position
        })

    # Automātiskais QR kods (akta ID)
    if akta_dati.include_auto_qr_code and akta_dati.akta_nr:
        qr_codes_to_add.append({
            'data': akta_dati.akta_nr,
            'size': akta_dati.auto_qr_code_size_mm,
            'position': akta_dati.auto_qr_code_position
        })

    if qr_codes_to_add:
        try:
            import qrcode
            qr_images = []
            for qr_info in qr_codes_to_add:
                qr = qrcode.QRCode(
                    version=1,
                    error_correction=qrcode.constants.ERROR_CORRECT_L,
                    box_size=10,
                    border=4,
                )
                qr.add_data(qr_info['data'])
                qr.make(fit=True)
                img_qr = qr.make_image(fill_color="black", back_color="white")
                qr_temp_path = tempfile.NamedTemporaryFile(delete=False, suffix=".png").name
                img_qr.save(qr_temp_path)

                rl_qr_img = RLImage(qr_temp_path)
                rl_qr_img._restrictSize(float(qr_info['size']) * mm, float(qr_info['size']) * mm)
                qr_images.append((rl_qr_img, qr_info['position'], qr_temp_path))

            # Pievienojam QR kodus story. Šeit varētu būt sarežģītāka pozicionēšanas loģika,
            # bet pagaidām pievienojam secīgi.
            # Lai nodrošinātu pozicionēšanu, būtu jāizmanto ReportLab "Frame" vai jāpielāgo "canvas" objekts.
            # Vienkāršības labad pievienojam tos beigās, katru savā rindā.
            story.append(Spacer(1, 12))  # Atstarpe pirms QR kodiem
            for img, pos, temp_path in qr_images:
                # ReportLab SimpleDocTemplate neļauj viegli pozicionēt elementus absolūti.
                # Lai to izdarītu, būtu jāizmanto canvas objekts tieši vai jāveido sarežģītākas tabulas.
                # Pagaidām pievienojam tos kā atsevišķus elementus.
                # Ja nepieciešama precīza pozicionēšana, tas ir ievērojams darbs.
                story.append(img)
                story.append(Spacer(1, 6))  # Atstarpe starp QR kodiem
                os.remove(temp_path)  # Dzēšam pagaidu failu

        except ImportError:
            print("QR Code generation requires 'qrcode' library. Please install it: pip install qrcode")
        except Exception as e:
            print(f"Error generating QR code: {e}")

    # Footer ar ģenerēšanas laiku
    if akta_dati.show_generation_timestamp:
        story.append(Spacer(1, 12))
        story.append(Paragraph(f"<font name='{font_name}'>Dokuments ģenerēts: {datetime.now().strftime('%Y-%m-%d %H:%M')}</font>", styles['LatvSmall']))

    # Build ar lapu numerāciju
    # Pass show_page_numbers to the custom canvas
    doc.build(story, canvasmaker=lambda *args, **kwargs: PageNumCanvas(*args, show_page_numbers=akta_dati.show_page_numbers, **kwargs))

    # PDF Encryption (new feature)
    if akta_dati.enable_pdf_encryption:
        try:
            from PyPDF2 import PdfReader, PdfWriter
            reader = PdfReader(pdf_ceļš)
            writer = PdfWriter()
            for page in reader.pages:
                writer.add_page(page)

            permissions = 0
            if akta_dati.allow_printing: permissions |= 4 # Print
            if akta_dati.allow_modifying: permissions |= 8 # Modify contents
            if akta_dati.allow_copying: permissions |= 16 # Copy
            if akta_dati.allow_annotating: permissions |= 32 # Annotate

            writer.encrypt(
                user_password=akta_dati.pdf_user_password,
                owner_password=akta_dati.pdf_owner_password,
                permissions_flag=permissions
            )
            with open(pdf_ceļš, "wb") as f:
                writer.write(f)
        except ImportError:
            print("PDF encryption requires 'PyPDF2' library. Please install it: pip install PyPDF2")
        except Exception as e:
            print(f"Error encrypting PDF: {e}")

    return pdf_ceļš

# ---------------------- DOCX ģenerēšana ----------------------
# (unchanged, as DOCX generation is less flexible with advanced styling)
def add_formatted_text(paragraph, text):
    # ... (unchanged)
    parts = []
    current_text = ""
    in_bold = False
    in_italic = False
    in_underline = False

    i = 0
    while i < len(text):
        if text[i:i+3] == "<b>":
            parts.append((current_text, in_bold, in_italic, in_underline))
            current_text = ""
            in_bold = True
            i += 3
        elif text[i:i+4] == "</b>":
            parts.append((current_text, in_bold, in_italic, in_underline))
            current_text = ""
            in_bold = False
            i += 4
        elif text[i:i+3] == "<i>":
            parts.append((current_text, in_bold, in_italic, in_underline))
            current_text = ""
            in_italic = True
            i += 3
        elif text[i:i+4] == "</i>":
            parts.append((current_text, in_bold, in_italic, in_underline))
            current_text = ""
            in_italic = False
            i += 4
        elif text[i:i+3] == "<u>":
            parts.append((current_text, in_bold, in_italic, in_underline))
            current_text = ""
            in_underline = True
            i += 3
        elif text[i:i+4] == "</u>":
            parts.append((current_text, in_bold, in_italic, in_underline))
            current_text = ""
            in_underline = False
            i += 4
        else:
            current_text += text[i]
            i += 1
    parts.append((current_text, in_bold, in_italic, in_underline))

    for part_text, bold, italic, underline in parts:
        if part_text:
            run = paragraph.add_run(part_text)
            run.bold = bold
            run.italic = italic
            run.underline = underline

def ģenerēt_docx(akta_dati: AktaDati, docx_ceļš: str):
    document = Document()

    document.styles['Normal'].font.name = 'Calibri'
    document.styles['Normal'].font.size = Pt(10)

    document.add_heading('PIEŅEMŠANAS–NODOŠANAS AKTS', level=1)
    document.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER

    if akta_dati.logotipa_ceļš and os.path.exists(akta_dati.logotipa_ceļš):
        try:
            document.add_picture(akta_dati.logotipa_ceļš, width=Inches(float(akta_dati.docx_image_width_inches)))
            document.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.LEFT
        except Exception:
            pass

    add_formatted_text(document.add_paragraph(), f"<b>Akta Nr.:</b> {akta_dati.akta_nr}")
    add_formatted_text(document.add_paragraph(), f"<b>Datums:</b> {akta_dati.datums}")
    add_formatted_text(document.add_paragraph(), f"<b>Vieta:</b> {akta_dati.vieta}")
    if akta_dati.pasūtījuma_nr:
        add_formatted_text(document.add_paragraph(), f"<b>Pasūtījuma Nr.:</b> {akta_dati.pasūtījuma_nr}")

    if akta_dati.līguma_nr:
        add_formatted_text(document.add_paragraph(), f"<b>Līguma Nr.:</b> {akta_dati.līguma_nr}")
    if akta_dati.izpildes_termiņš:
        add_formatted_text(document.add_paragraph(), f"<b>Izpildes termiņš:</b> {akta_dati.izpildes_termiņš}")
    if akta_dati.pieņemšanas_datums:
        add_formatted_text(document.add_paragraph(), f"<b>Pieņemšanas datums:</b> {akta_dati.pieņemšanas_datums}")
    if akta_dati.nodošanas_datums:
        add_formatted_text(document.add_paragraph(), f"<b>Nodošanas datums:</b> {akta_dati.nodošanas_datums}")
    document.add_paragraph()

    document.add_heading('Puses', level=2)
    table_parties = document.add_table(rows=1, cols=2)
    table_parties.autofit = True
    table_parties.columns[0].width = Inches(3)
    table_parties.columns[1].width = Inches(3)

    hdr_cells = table_parties.rows[0].cells
    hdr_cells[0].text = "Pieņēmējs"
    hdr_cells[1].text = "Nodevējs"

    row_cells = table_parties.add_row().cells
    def add_persona_info(cell, p: Persona):
        add_formatted_text(cell.add_paragraph(), f"<b>Nosaukums / Vārds, Uzvārds:</b> {p.nosaukums}")
        if p.reģ_nr: add_formatted_text(cell.add_paragraph(), f"<b>Reģ. Nr. / personas kods:</b> {p.reģ_nr}")
        if p.adrese: add_formatted_text(cell.add_paragraph(), f"<b>Adrese:</b> {p.adrese}")
        if p.kontaktpersona: add_formatted_text(cell.add_paragraph(), f"<b>Kontaktpersona:</b> {p.kontaktpersona}")
        if p.tālrunis: add_formatted_text(cell.add_paragraph(), f"<b>Tālrunis:</b> {p.tālrunis}")
        if p.epasts: add_formatted_text(cell.add_paragraph(), f"<b>E-pasts:</b> {p.epasts}")
        if p.bankas_konts: add_formatted_text(cell.add_paragraph(), f"<b>Bankas konts:</b> {p.bankas_konts}")
        if p.juridiskais_statuss: add_formatted_text(cell.add_paragraph(), f"<b>Statuss:</b> {p.juridiskais_statuss}")

    add_persona_info(row_cells[0], akta_dati.pieņēmējs)
    add_persona_info(row_cells[1], akta_dati.nodevējs)
    document.add_paragraph()

    document.add_heading('Pozīcijas', level=2)
    # Adjust columns based on visibility settings for DOCX
    docx_headers = ["Nr.", "Apraksts", "Daudzums", "Vienība", "Cena", "Summa"]
    if akta_dati.show_item_serial_number_in_table: docx_headers.append("Seriālais Nr.")
    if akta_dati.show_item_warranty_in_table: docx_headers.append("Garantija")
    if akta_dati.show_item_notes_in_table: docx_headers.append("Piezīmes pozīcijai")

    table_items = document.add_table(rows=1, cols=len(docx_headers))
    table_items.autofit = True
    table_items.allow_autofit = True

    for i, header_text in enumerate(docx_headers):
        table_items.rows[0].cells[i].text = header_text
        table_items.rows[0].cells[i].paragraphs[0].runs[0].bold = True

    for i, poz in enumerate(akta_dati.pozīcijas, start=1):
        row_cells = table_items.add_row().cells
        row_cells[0].text = str(i)
        add_formatted_text(row_cells[1].paragraphs[0], poz.apraksts)
        row_cells[2].text = f"{formēt_naudu(poz.daudzums)}"
        row_cells[3].text = poz.vienība
        row_cells[4].text = f"{formēt_naudu(poz.cena)} {akta_dati.valūta}"
        row_cells[5].text = f"{formēt_naudu(poz.summa)} {akta_dati.valūta}"
        col_idx = 6
        if akta_dati.show_item_serial_number_in_table:
            row_cells[col_idx].text = poz.seriālais_nr
            col_idx += 1
        if akta_dati.show_item_warranty_in_table:
            row_cells[col_idx].text = poz.garantija
            col_idx += 1
        if akta_dati.show_item_notes_in_table:
            add_formatted_text(row_cells[col_idx].paragraphs[0], poz.piezīmes_pozīcijai)

    document.add_paragraph()

    document.add_heading('Kopsavilkums', level=2)
    add_formatted_text(document.add_paragraph(), f"<b>Kopā bez PVN:</b> {formēt_naudu(akta_dati.kopējā_summma())} {akta_dati.valūta}")
    if akta_dati.iekļaut_pvn:
        add_formatted_text(document.add_paragraph(), f"<b>PVN {akta_dati.pvn_likme}%:</b> {formēt_naudu(akta_dati.pvn_summa())} {akta_dati.valūta}")
        add_formatted_text(document.add_paragraph(), f"<b>Kopā ar PVN:</b> {formēt_naudu(akta_dati.summa_ar_pvn())} {akta_dati.valūta}")
    document.add_paragraph()

    if akta_dati.piezīmes:
        document.add_heading('Vispārīgās piezīmes', level=2)
        add_formatted_text(document.add_paragraph(), akta_dati.piezīmes)
        document.add_paragraph()

    if akta_dati.strīdu_risināšana:
        document.add_heading('Strīdu risināšanas kārtība', level=2)
        add_formatted_text(document.add_paragraph(), akta_dati.strīdu_risināšana)
        document.add_paragraph()

    if akta_dati.konfidencialitātes_klauzula:
        document.add_heading('Konfidencialitātes klauzula', level=2)
        add_formatted_text(document.add_paragraph(), "Puses apņemas neizpaust trešajām personām informāciju, kas iegūta šī akta ietvaros, izņemot gadījumus, ko nosaka normatīvie akti.")
        document.add_paragraph()

    if akta_dati.soda_nauda_procenti > 0:
        document.add_heading('Soda nauda', level=2)
        add_formatted_text(document.add_paragraph(), f"Par saistību neizpildi vai nepienācīgu izpildi, vainīgā puse maksā otrai pusei soda naudu {formēt_naudu(akta_dati.soda_nauda_procenti)}% apmērā no neizpildīto saistību vērtības par katru kavējuma dienu.")
        document.add_paragraph()

    if akta_dati.piegādes_nosacījumi:
        document.add_heading('Piegādes nosacījumi', level=2)
        add_formatted_text(document.add_paragraph(), akta_dati.piegādes_nosacījumi)
        document.add_paragraph()

    if akta_dati.apdrošināšana:
        document.add_heading('Apdrošināšana', level=2)
        add_formatted_text(document.add_paragraph(), "Preces ir apdrošinātas pret bojājumiem un zaudējumiem līdz pieņemšanas-nodošanas brīdim.")
        document.add_paragraph()

    if akta_dati.papildu_nosacījumi:
        document.add_heading('Papildu nosacījumi', level=2)
        add_formatted_text(document.add_paragraph(), akta_dati.papildu_nosacījumi)
        document.add_paragraph()

    if akta_dati.atsauces_dokumenti:
        document.add_heading('Atsauces dokumenti', level=2)
        add_formatted_text(document.add_paragraph(), akta_dati.atsauces_dokumenti)
        document.add_paragraph()

    document.add_heading('Akta statuss', level=2)
    add_formatted_text(document.add_paragraph(), akta_dati.akta_statuss)
    document.add_paragraph()

    if akta_dati.attēli:
        document.add_heading('Pievienotās fotogrāfijas', level=2)
        for att in akta_dati.attēli:
            if os.path.exists(att.ceļš):
                try:
                    document.add_picture(att.ceļš, width=Inches(float(akta_dati.docx_image_width_inches)))
                    if att.paraksts:
                        add_formatted_text(document.add_paragraph(), att.paraksts)
                except Exception:
                    pass
        document.add_paragraph()

    if akta_dati.elektroniskais_paraksts:
        if akta_dati.radit_elektronisko_parakstu_tekstu:
            document.add_paragraph()
            p = document.add_paragraph()
            run = p.add_run("ŠIS DOKUMENTS PARAKSTĪTS AR DROŠU ELEKTRONISKO PARAKSTU UN SATUR LAIKA ZĪMOGU")
            run.bold = True
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    elif akta_dati.parakstu_rindas:
        document.add_heading('Paraksti', level=2)
        table_signatures = document.add_table(rows=2, cols=2)
        table_signatures.autofit = True
        table_signatures.columns[0].width = Inches(3)
        table_signatures.columns[1].width = Inches(3)

        table_signatures.rows[0].cells[0].text = "Pieņēmējs"
        table_signatures.rows[0].cells[1].text = "Nodevējs"
        for cell in table_signatures.rows[0].cells:
            cell.paragraphs[0].runs[0].bold = True
            cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER

        cell_pie = table_signatures.rows[1].cells[0]
        cell_nod = table_signatures.rows[1].cells[1]

        if akta_dati.paraksts_pieņēmējs_ceļš and os.path.exists(akta_dati.paraksts_pieņēmējs_ceļš):
            try:
                cell_pie.add_paragraph().add_run().add_picture(akta_dati.paraksts_pieņēmējs_ceļš, width=Inches(
                    float(akta_dati.docx_signature_width_inches)))
                cell_pie.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER
            except Exception:
                cell_pie.add_paragraph("____________________________")
        else:
            cell_pie.add_paragraph("____________________________")

        if akta_dati.paraksts_nodevējs_ceļš and os.path.exists(akta_dati.paraksts_nodevējs_ceļš):
            try:
                cell_nod.add_paragraph().add_run().add_picture(akta_dati.paraksts_nodevējs_ceļš, width=Inches(
                    float(akta_dati.docx_signature_width_inches)))
                cell_nod.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER
            except Exception:
                cell_nod.add_paragraph("____________________________")
        else:
            cell_nod.add_paragraph("____________________________")

        add_formatted_text(cell_pie.add_paragraph(),
                           akta_dati.pieņēmējs.kontaktpersona or akta_dati.pieņēmējs.nosaukums)
        add_formatted_text(cell_nod.add_paragraph(), akta_dati.nodevējs.kontaktpersona or akta_dati.nodevējs.nosaukums)
        for cell in [cell_pie, cell_nod]:
            for p in cell.paragraphs:
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER

    # Vecais QR koda bloks, kas izmantoja neeksistējošus atribūtus.
    # Šis bloks ir jāizdzēš vai jākomentē, jo jaunā funkcionalitāte to aizstāj.
    # Ja vēlaties, lai DOCX ģenerēšana atbalstītu abus QR kodus, jums būtu jāpievieno līdzīga loģika kā PDF ģenerēšanā.
    # Piemēram, ja vēlaties, lai tas izmantotu individuālo QR kodu:
    # if akta_dati.include_custom_qr_code and akta_dati.custom_qr_code_data:
    #     try:
    #         import qrcode
    #         qr = qrcode.QRCode(
    #             version=1,
    #             error_correction=qrcode.constants.ERROR_CORRECT_L,
    #             box_size=10,
    #             border=4,
    #         )
    #         qr.add_data(akta_dati.custom_qr_code_data)
    #         qr.make(fit=True)
    #         img_qr = qr.make_image(fill_color="black", back_color="white")
    #         qr_temp_path = tempfile.NamedTemporaryFile(delete=False, suffix=".png").name
    #         img_qr.save(qr_temp_path)
    #
    #         # DOCX attēla platums balstīts uz QR koda izmēru mm, konvertējot uz collām
    #         qr_width_inches = float(akta_dati.custom_qr_code_size_mm) / 25.4
    #         document.add_picture(qr_temp_path, width=Inches(qr_width_inches))
    #         document.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.RIGHT  # Pēc noklusējuma labajā pusē
    #         os.remove(qr_temp_path)  # Clean up temp file
    #     except ImportError:
    #         print("QR Code generation requires 'qrcode' library. Please install it: pip install qrcode")
    #     except Exception as e:
    #         print(f"Error generating QR code for DOCX: {e}")

    add_formatted_text(document.add_paragraph(), f"Dokuments ģenerēts: {datetime.now().strftime('%Y-%m-%d %H:%M')}")
    document.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.RIGHT

    document.save(docx_ceļš)


# ---------------------- GUI ----------------------

# Custom URL interceptor for QWebEngineView
class MapUrlInterceptor(QWebEngineUrlRequestInterceptor):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.map_click_callback = None

    def interceptRequest(self, info):
        url = info.requestUrl()
        # Pārbaudām, vai URL ir mūsu pielāgotā shēma
        if url.scheme() == "app" and url.host() == "map_click":
            lat = url.queryItemValue("lat")
            lon = url.queryItemValue("lon")
            if self.map_click_callback:
                # Izsaucam atpakaļsaites funkciju ar koordinātēm
                self.map_click_callback(lat, lon)
            # Svarīgi: Bloķējam pieprasījumu, lai pārlūkprogramma nemēģinātu atvērt šo URL
            # un neradītu brīdinājumu.
            info.block(True)
            # info.redirect(QUrl()) # Šī rinda var nebūt nepieciešama, ja block(True) darbojas efektīvi
        # Ja URL nav mūsu pielāgotā shēma, ļaujam tam turpināties
        else:
            info.block(False)

# Pievienot šīs klases pirms AktaLogs klases definīcijas

class TextBlockLineEdit(QWidget):
    def __init__(self, text_block_manager, field_name, parent=None):
        super().__init__(parent)
        self.text_block_manager = text_block_manager
        self.field_name = field_name

        self.layout = QVBoxLayout(self)
        self.layout.setContentsMargins(0, 0, 0, 0)

        self.line_edit = QLineEdit()
        self.layout.addWidget(self.line_edit)

        self.combo_box = QComboBox()
        self.combo_box.addItem("--- Izvēlēties saglabāto bloku ---")
        self.combo_box.currentIndexChanged.connect(self._load_selected_block)
        self.layout.addWidget(self.combo_box)

        button_layout = QHBoxLayout()
        self.save_button = QPushButton("Saglabāt bloku")
        self.save_button.clicked.connect(self._save_block)
        button_layout.addWidget(self.save_button)

        self.delete_button = QPushButton("Dzēst bloku")
        self.delete_button.clicked.connect(self._delete_block)
        button_layout.addWidget(self.delete_button)
        button_layout.addStretch()
        self.layout.addLayout(button_layout)

        self._update_combo_box()

    def text(self):
        return self.line_edit.text()

    def setText(self, text):
        self.line_edit.setText(text)

    def _update_combo_box(self):
        self.combo_box.blockSignals(True) # Bloķējam signālus, lai izvairītos no nevajadzīgas ielādes
        self.combo_box.clear()
        self.combo_box.addItem("--- Izvēlēties saglabāto bloku ---")
        blocks = self.text_block_manager.get_blocks_for_field(self.field_name)
        for name in sorted(blocks.keys()):
            self.combo_box.addItem(name)
        self.combo_box.blockSignals(False)

    def _load_selected_block(self, index):
        if index > 0:
            block_name = self.combo_box.currentText()
            content = self.text_block_manager.get_block_content(self.field_name, block_name)
            self.line_edit.setText(content)
            QMessageBox.information(self, "Ielādēts", f"Teksta bloks '{block_name}' ielādēts.")
        self.combo_box.setCurrentIndex(0) # Atgriežam izvēli uz noklusējumu

    def _save_block(self):
        current_text = self.line_edit.text().strip()
        if not current_text:
            QMessageBox.warning(self, "Saglabāt bloku", "Ievades lauks ir tukšs. Lūdzu, ievadiet tekstu, ko saglabāt.")
            return

        block_name, ok = QInputDialog.getText(self, "Saglabāt teksta bloku", "Ievadiet bloka nosaukumu:")
        if ok and block_name:
            self.text_block_manager.add_block(self.field_name, block_name, current_text)
            self._update_combo_box()
            QMessageBox.information(self, "Saglabāts", f"Teksta bloks '{block_name}' saglabāts.")
        elif ok:
            QMessageBox.warning(self, "Saglabāt bloku", "Bloka nosaukums nevar būt tukšs.")

    def _delete_block(self):
        blocks = self.text_block_manager.get_blocks_for_field(self.field_name)
        if not blocks:
            QMessageBox.information(self, "Dzēst bloku", "Nav saglabātu teksta bloku šim laukam.")
            return

        block_names = sorted(blocks.keys())
        block_name, ok = QInputDialog.getItem(self, "Dzēst teksta bloku", "Izvēlieties bloku, ko dzēst:", block_names, 0, False)
        if ok and block_name:
            reply = QMessageBox.question(self, "Dzēst bloku",
                                         f"Vai tiešām vēlaties dzēst teksta bloku '{block_name}'?",
                                         QMessageBox.Yes | QMessageBox.No)
            if reply == QMessageBox.Yes:
                self.text_block_manager.delete_block(self.field_name, block_name)
                self._update_combo_box()
                QMessageBox.information(self, "Dzēsts", f"Teksta bloks '{block_name}' dzēsts.")


class TextBlockTextEdit(QWidget):
    def __init__(self, text_block_manager, field_name, parent=None):
        super().__init__(parent)
        self.text_block_manager = text_block_manager
        self.field_name = field_name

        self.layout = QVBoxLayout(self)
        self.layout.setContentsMargins(0, 0, 0, 0)

        self.text_edit = QTextEdit()
        self.layout.addWidget(self.text_edit)

        self.list_widget = QListWidget()
        self.list_widget.setMaximumHeight(100) # Ierobežojam augstumu, lai neaizņemtu pārāk daudz vietas
        self.list_widget.itemDoubleClicked.connect(self._load_selected_block)
        self.layout.addWidget(self.list_widget)

        button_layout = QHBoxLayout()
        self.save_button = QPushButton("Saglabāt bloku")
        self.save_button.clicked.connect(self._save_block)
        button_layout.addWidget(self.save_button)

        self.delete_button = QPushButton("Dzēst bloku")
        self.delete_button.clicked.connect(self._delete_block)
        button_layout.addWidget(self.delete_button)
        button_layout.addStretch()
        self.layout.addLayout(button_layout)

        self._update_list_widget()

    def toPlainText(self):
        return self.text_edit.toPlainText()

    def setText(self, text):
        self.text_edit.setText(text)

    def _update_list_widget(self):
        self.list_widget.clear()
        blocks = self.text_block_manager.get_blocks_for_field(self.field_name)
        for name in sorted(blocks.keys()):
            self.list_widget.addItem(name)

    def _load_selected_block(self, item):
        block_name = item.text()
        content = self.text_block_manager.get_block_content(self.field_name, block_name)
        self.text_edit.setPlainText(content)
        QMessageBox.information(self, "Ielādēts", f"Teksta bloks '{block_name}' ielādēts.")

    def _save_block(self):
        current_text = self.text_edit.toPlainText().strip()
        if not current_text:
            QMessageBox.warning(self, "Saglabāt bloku", "Ievades lauks ir tukšs. Lūdzu, ievadiet tekstu, ko saglabāt.")
            return

        block_name, ok = QInputDialog.getText(self, "Saglabāt teksta bloku", "Ievadiet bloka nosaukumu:")
        if ok and block_name:
            self.text_block_manager.add_block(self.field_name, block_name, current_text)
            self._update_list_widget()
            QMessageBox.information(self, "Saglabāts", f"Teksta bloks '{block_name}' saglabāts.")
        elif ok:
            QMessageBox.warning(self, "Saglabāt bloku", "Bloka nosaukums nevar būt tukšs.")

    def _delete_block(self):
        selected_item = self.list_widget.currentItem()
        if not selected_item:
            QMessageBox.warning(self, "Dzēst bloku", "Lūdzu, izvēlieties bloku, ko dzēst.")
            return

        block_name = selected_item.text()
        reply = QMessageBox.question(self, "Dzēst bloku",
                                     f"Vai tiešām vēlaties dzēst teksta bloku '{block_name}'?",
                                     QMessageBox.Yes | QMessageBox.No)
        if reply == QMessageBox.Yes:
            self.text_block_manager.delete_block(self.field_name, block_name)
            self._update_list_widget()
            QMessageBox.information(self, "Dzēsts", f"Teksta bloks '{block_name}' dzēsts.")


class AktaLogs(QMainWindow):
    def __init__(self):
            super().__init__()
            self.setWindowTitle("Pieņemšanas–Nodošanas akta ģenerators")
            self.resize(1200, 800)
            self.poppler_path = ""

            self.zoom_factor = 1.0

            self.history = [] # Inicializējam tukšu sarakstu
            self.address_book = {} # Inicializējam tukšu vārdnīcu
            self.text_block_manager = TextBlockManager() # JAUNA RINDAS
            self.data = AktaDati(datums=datetime.now().strftime('%Y-%m-%d'))
            self._ceļš_projekts = None

            self.tabs = QTabWidget()

            self._būvēt_pamata_tab()
            self._būvēt_puses_tab()
            self._būvēt_pozīcijas_tab()
            self._būvēt_attēli_tab()
            self._būvēt_iestatījumi_tab()
            self._būvēt_papildu_iestatījumi_tab()
            self._būvēt_sablonu_tab()
            self._būvēt_adresu_gramata_tab()
            self._būvēt_dokumentu_vesture_tab()
            self._būvēt_kartes_tab()  # New map tab
            # Pārliecināmies, ka noklusējuma šablonu direktorijs ir iestatīts
            if not self.data.templates_dir:
                self.data.templates_dir = os.path.join(APP_DATA_DIR, "AktaGenerators_Templates")
                os.makedirs(self.data.templates_dir, exist_ok=True)  # Izveidojam noklusējuma mapi, ja tā neeksistē
            self._update_sablonu_list()

            # Ielādējam vēsturi un adrešu grāmatu PĒC GUI elementu izveides
            self._load_history()
            self._load_address_book()
            self._update_address_book_list()
            self.ieladet_noklusejuma_iestatijumus()


            main_splitter = QSplitter(Qt.Horizontal)
            self.setCentralWidget(main_splitter)

            main_splitter.addWidget(self.tabs)

            preview_widget = QWidget()
            preview_layout = QVBoxLayout(preview_widget)
            self.preview_label = QLabel("PDF priekšskatījums")
            self.preview_label.setAlignment(Qt.AlignHCenter | Qt.AlignVCenter)
            self.preview_label.setMinimumSize(1, 1)
            preview_layout.addWidget(self.preview_label)

            self.preview_scroll_area = QScrollArea()
            self.preview_scroll_area.setWidgetResizable(True)
            self.preview_scroll_area.setWidget(self.preview_label)
            preview_layout.addWidget(self.preview_scroll_area)

            self.preview_scroll_area.resizeEvent = self._update_preview_on_resize

            self.current_preview_page = 0
            self.preview_images = []

            nav_buttons_layout = QHBoxLayout()
            self.prev_page_button = QPushButton("< Iepriekšējā lapa")
            self.prev_page_button.clicked.connect(self._show_prev_page)
            self.prev_page_button.setEnabled(False)
            self.next_page_button = QPushButton("Nākamā lapa >")
            self.next_page_button.clicked.connect(self._show_next_page)
            self.next_page_button.setEnabled(False)
            self.page_number_label = QLabel("Lapa 1/1")
            self.page_number_label.setAlignment(Qt.AlignCenter)

            nav_buttons_layout.addWidget(self.prev_page_button)
            nav_buttons_layout.addWidget(self.page_number_label)
            nav_buttons_layout.addWidget(self.next_page_button)

            zoom_in_button = QPushButton("Palielināt")
            zoom_in_button.clicked.connect(self.zoom_in)
            zoom_out_button = QPushButton("Samazināt")
            zoom_out_button.clicked.connect(self.zoom_out)
            nav_buttons_layout.addWidget(zoom_in_button)
            nav_buttons_layout.addWidget(zoom_out_button)

            preview_layout.addLayout(nav_buttons_layout)

            main_splitter.addWidget(preview_widget)

            main_splitter.setSizes([int(self.width() * 0.7), int(self.width() * 0.3)])

            self.tabs.currentChanged.connect(self._update_preview)

            self._update_preview()

            self._būvēt_menu()

            if self.data.auto_generate_akta_nr:
                self._generate_akta_nr()

    # ----- Menu -----
    def _būvēt_menu(self):
        menubar = self.menuBar()
        fajls = menubar.addMenu("&Fails")

        saglabat = QAction("Saglabāt projektu…", self)
        saglabat.triggered.connect(self.saglabat_projektu)
        fajls.addAction(saglabat)

        ieladet = QAction("Ielādēt projektu…", self)
        ieladet.triggered.connect(self.ieladet_projektu)
        fajls.addAction(ieladet)

        fajls.addSeparator()

        eksportet_pdf = QAction("Ģenerēt PDF…", self)
        eksportet_pdf.triggered.connect(self.ģenerēt_pdf_dialogs)
        fajls.addAction(eksportet_pdf)

        eksportet_docx = QAction("Ģenerēt DOCX…", self)
        eksportet_docx.triggered.connect(self.ģenerēt_docx_dialogs)
        fajls.addAction(eksportet_docx)

        drukāt_pdf_menu = QAction("Drukāt PDF…", self) # JAUNA RINDAS
        drukāt_pdf_menu.triggered.connect(self.drukāt_pdf_dialogs) # JAUNA RINDAS
        fajls.addAction(drukāt_pdf_menu) # JAUNA RINDAS


        fajls.addSeparator()

        iziet = QAction("Iziet", self)
        iziet.triggered.connect(self.close)
        fajls.addAction(iziet)

    # ----- Tab: Pamata -----

    def _būvēt_pamata_tab(self):
        content_widget = QWidget()
        form = QFormLayout()

        self.in_akta_nr = QLineEdit()
        self.btn_generate_akta_nr = QPushButton("Ģenerēt Nr.")
        self.btn_generate_akta_nr.clicked.connect(self._generate_akta_nr)
        akta_nr_layout = QHBoxLayout()
        akta_nr_layout.addWidget(self.in_akta_nr)
        akta_nr_layout.addWidget(self.btn_generate_akta_nr)
        akta_nr_widget = QWidget();
        akta_nr_widget.setLayout(akta_nr_layout)

        self.in_datums = QDateEdit(calendarPopup=True)
        self.in_datums.setDisplayFormat("yyyy-MM-dd")
        self.in_datums.setDate(datetime.now().date())

        self.in_vieta = QLineEdit()
        self.in_pas_nr = QLineEdit()
        self.in_liguma_nr = QLineEdit()

        self.in_izpildes_termins = QDateEdit(calendarPopup=True)
        self.in_izpildes_termins.setDisplayFormat("yyyy-MM-dd")
        self.in_izpildes_termins.setMinimumDate(datetime(1900, 1, 1).date())
        self.in_izpildes_termins.setMaximumDate(datetime(2100, 1, 1).date())
        self.in_izpildes_termins.setDate(datetime.now().date()) # Set default to today
        self.in_izpildes_termins.setSpecialValueText("Nav norādīts") # Text for null date
        self.in_izpildes_termins.setCalendarPopup(True)


        self.in_pieņemšanas_datums = QDateEdit(calendarPopup=True)
        self.in_pieņemšanas_datums.setDisplayFormat("yyyy-MM-dd")
        self.in_pieņemšanas_datums.setMinimumDate(datetime(1900, 1, 1).date())
        self.in_pieņemšanas_datums.setMaximumDate(datetime(2100, 1, 1).date())
        self.in_pieņemšanas_datums.setDate(datetime.now().date())
        self.in_pieņemšanas_datums.setSpecialValueText("Nav norādīts")
        self.in_pieņemšanas_datums.setCalendarPopup(True)

        self.in_nodošanas_datums = QDateEdit(calendarPopup=True)
        self.in_nodošanas_datums.setDisplayFormat("yyyy-MM-dd")
        self.in_nodošanas_datums.setMinimumDate(datetime(1900, 1, 1).date())
        self.in_nodošanas_datums.setMaximumDate(datetime(2100, 1, 1).date())
        self.in_nodošanas_datums.setDate(datetime.now().date())
        self.in_nodošanas_datums.setSpecialValueText("Nav norādīts")
        self.in_nodošanas_datums.setCalendarPopup(True)

        # Strīdu risināšana
        self.in_strīdu_risināšana = TextBlockTextEdit(self.text_block_manager, "stridu_risinasana")

        self.ck_konfidencialitate = QCheckBox("Konfidencialitāte")
        self.in_soda_nauda_procenti = QDoubleSpinBox();
        self.in_soda_nauda_procenti.setRange(0.0, 100.0);
        self.in_soda_nauda_procenti.setSuffix(" %")

        # Piegādes nosacījumi
        self.in_piegades_nosacijumi = TextBlockLineEdit(self.text_block_manager, "piegades_nosacijumi")

        self.ck_apdrošināšana = QCheckBox("Apdrošināšana")

        # Papildu nosacījumi
        self.in_papildu_nosacijumi = TextBlockTextEdit(self.text_block_manager, "papildu_nosacijumi")

        # Atsauces dokumenti
        self.in_atsauces_dokumenti = TextBlockLineEdit(self.text_block_manager, "atsauces_dokumenti")

        self.cb_akta_statuss = QComboBox()
        self.cb_akta_statuss.addItems(["Melnraksts", "Apstiprināts", "Parakstīts", "Arhivēts", "Atcelts"])
        self.in_valuta = QLineEdit()
        self.in_valuta.setText("EUR")

        # Piezīmes
        self.in_piezimes = TextBlockTextEdit(self.text_block_manager, "piezimes")

        self.ck_elektroniskais_paraksts = QCheckBox("Elektroniskais paraksts (ignorē fiziskos parakstus)")
        self.ck_radit_elektronisko_parakstu_tekstu = QCheckBox("Rādīt elektroniskā paraksta tekstu PDF dokumentā") # JAUNA RŪTIŅA

        form.addRow("Akta Nr.", akta_nr_widget)
        form.addRow("Datums", self.in_datums)
        form.addRow("Vieta", self.in_vieta)
        form.addRow("Pasūtījuma Nr.", self.in_pas_nr)
        form.addRow("Līguma Nr.", self.in_liguma_nr)
        form.addRow("Izpildes termiņš", self.in_izpildes_termins)
        form.addRow("Pieņemšanas datums", self.in_pieņemšanas_datums)
        form.addRow("Nodošanas datums", self.in_nodošanas_datums)
        form.addRow("Strīdu risināšana", self.in_strīdu_risināšana)  # Tagad tieši izmantojam jauno objektu
        form.addRow(self.ck_konfidencialitate)
        form.addRow("Soda nauda (%)", self.in_soda_nauda_procenti)
        form.addRow("Piegādes nosacījumi", self.in_piegades_nosacijumi)  # Tagad tieši izmantojam jauno objektu
        form.addRow(self.ck_apdrošināšana)
        form.addRow("Papildu nosacījumi", self.in_papildu_nosacijumi)  # Tagad tieši izmantojam jauno objektu
        form.addRow("Atsauces dokumenti", self.in_atsauces_dokumenti)  # Tagad tieši izmantojam jauno objektu
        form.addRow("Akta statuss", self.cb_akta_statuss)
        form.addRow("Valūta", self.in_valuta)
        form.addRow("Piezīmes", self.in_piezimes)  # Tagad tieši izmantojam jauno objektu
        form.addRow("Elektroniskais paraksts", self.ck_elektroniskais_paraksts)
        form.addRow("", self.ck_radit_elektronisko_parakstu_tekstu) # JAUNA RŪTIŅA FORMĀ

        content_widget.setLayout(form)

        scroll_area = QScrollArea()
        scroll_area.setWidgetResizable(True)
        scroll_area.setWidget(content_widget)

        main_tab_widget = QWidget()
        main_layout = QVBoxLayout(main_tab_widget)
        main_layout.addWidget(scroll_area)

        self.tabs.addTab(main_tab_widget, "Pamata dati")

        nav_buttons_layout = QHBoxLayout()
        zoom_in_button = QPushButton("Palielināt")
        zoom_in_button.clicked.connect(self.zoom_in)
        zoom_out_button = QPushButton("Samazināt")
        zoom_out_button.clicked.connect(self.zoom_out)

        nav_buttons_layout.addWidget(zoom_in_button)
        nav_buttons_layout.addWidget(zoom_out_button)

        main_layout.addLayout(nav_buttons_layout)

        self.in_akta_nr.textChanged.connect(self._update_preview)
        self.in_datums.dateChanged.connect(self._update_preview)
        self.in_vieta.textChanged.connect(self._update_preview)
        self.in_pas_nr.textChanged.connect(self._update_preview)
        self.in_liguma_nr.textChanged.connect(self._update_preview)
        self.in_izpildes_termins.dateChanged.connect(self._update_preview)
        self.in_pieņemšanas_datums.dateChanged.connect(self._update_preview)
        self.in_nodošanas_datums.dateChanged.connect(self._update_preview)
        self.in_strīdu_risināšana.text_edit.textChanged.connect(self._update_preview)  # Savienojam ar iekšējo text_edit
        self.ck_konfidencialitate.stateChanged.connect(self._update_preview)
        self.in_soda_nauda_procenti.valueChanged.connect(self._update_preview)
        self.in_piegades_nosacijumi.line_edit.textChanged.connect(
            self._update_preview)  # Savienojam ar iekšējo line_edit
        self.ck_apdrošināšana.stateChanged.connect(self._update_preview)
        self.in_papildu_nosacijumi.text_edit.textChanged.connect(
            self._update_preview)  # Savienojam ar iekšējo text_edit
        self.in_atsauces_dokumenti.line_edit.textChanged.connect(
            self._update_preview)  # Savienojam ar iekšējo line_edit
        self.cb_akta_statuss.currentIndexChanged.connect(self._update_preview)
        self.in_valuta.textChanged.connect(self._update_preview)
        self.in_piezimes.text_edit.textChanged.connect(self._update_preview)  # Savienojam ar iekšējo text_edit
        self.ck_elektroniskais_paraksts.stateChanged.connect(self._update_preview)
        self.ck_radit_elektronisko_parakstu_tekstu.stateChanged.connect(self._update_preview) # JAUNS SAVIENOJUMS





    def zoom_in(self):
        self.zoom_factor *= 1.1
        self._show_current_page()

    def zoom_out(self):
        self.zoom_factor /= 1.1
        self._show_current_page()

    def _update_preview_on_resize(self, event):
        self._show_current_page() # Update preview to fit new size
        super().resizeEvent(event)

    def _generate_akta_nr(self):
        current_year = datetime.now().strftime('%Y')
        used_ids = set()

        # 1. Pārbaudām jau saglabātos projektus (JSON failus)
        if os.path.exists(PROJECT_SAVE_DIR):
            for filename in os.listdir(PROJECT_SAVE_DIR):
                if filename.endswith(".json"):
                    try:
                        with open(os.path.join(PROJECT_SAVE_DIR, filename), 'r', encoding='utf-8') as f:
                            project_data = json.load(f)
                            if 'akta_nr' in project_data and project_data['akta_nr']:
                                used_ids.add(project_data['akta_nr'])
                    except Exception as e:
                        print(f"Kļūda lasot projektu failu {filename}: {e}")

        # 2. Pārbaudām vēstures ierakstus (ja tie satur akta_nr)
        # Piezīme: Jums būs jāpārliecinās, ka vēstures ieraksti saglabā akta_nr.
        # Pašreizējā `_add_to_history` funkcija saglabā tikai faila ceļu.
        # Mēs to uzlabosim vēlāk, lai saglabātu arī akta_nr.
        # Pagaidām pieņemam, ka vēstures ieraksts ir faila ceļš, un mēģinām no tā iegūt akta_nr.
        for entry_path in self.history:
            if os.path.exists(entry_path) and entry_path.endswith(".json"):
                try:
                    with open(entry_path, 'r', encoding='utf-8') as f:
                        history_data = json.load(f)
                        if 'akta_nr' in history_data and history_data['akta_nr']:
                            used_ids.add(history_data['akta_nr'])
                except Exception:
                    pass  # Ignorējam kļūdas, ja vēstures fails ir bojāts vai nav JSON

        # 3. Ģenerējam jaunu, unikālu ID
        new_seq = 1
        while True:
            new_akta_nr = f"PP-{current_year}-{new_seq:04d}"
            if new_akta_nr not in used_ids:
                self.in_akta_nr.setText(new_akta_nr)
                break
            new_seq += 1

    def savākt_datus(self) -> AktaDati:
        d = AktaDati()
        d.akta_nr = self.in_akta_nr.text().strip()
        d.datums = self.in_datums.date().toString("yyyy-MM-dd")
        d.vieta = self.in_vieta.text().strip()
        d.pasūtījuma_nr = self.in_pas_nr.text().strip()
        d.piezīmes = self.in_piezimes.toPlainText().strip()

        # JAUNA RINDAS
        d.templates_dir = self.in_templates_dir.text().strip()

        d.līguma_nr = self.in_liguma_nr.text().strip()
        d.izpildes_termiņš = self.in_izpildes_termins.date().toString("yyyy-MM-dd") if not self.in_izpildes_termins.date().isNull() else ""
        d.pieņemšanas_datums = self.in_pieņemšanas_datums.date().toString("yyyy-MM-dd") if not self.in_pieņemšanas_datums.date().isNull() else ""
        d.nodošanas_datums = self.in_nodošanas_datums.date().toString("yyyy-MM-dd") if not self.in_nodošanas_datums.date().isNull() else ""
        d.strīdu_risināšana = self.in_strīdu_risināšana.toPlainText().strip()  # Tagad izsauc toPlainText() uz pielāgotā logrīka
        d.konfidencialitātes_klauzula = self.ck_konfidencialitate.isChecked()
        d.soda_nauda_procenti = to_decimal(self.in_soda_nauda_procenti.value())
        d.piegādes_nosacījumi = self.in_piegades_nosacijumi.text().strip()  # Tagad izsauc text() uz pielāgotā logrīka
        d.apdrošināšana = self.ck_apdrošināšana.isChecked()
        d.papildu_nosacījumi = self.in_papildu_nosacijumi.toPlainText().strip()  # Tagad izsauc toPlainText() uz pielāgotā logrīka
        d.atsauces_dokumenti = self.in_atsauces_dokumenti.text().strip()  # Tagad izsauc text() uz pielāgotā logrīka
        d.akta_statuss = self.cb_akta_statuss.currentText()
        d.valūta = self.in_valuta.text().strip()
        d.elektroniskais_paraksts = self.ck_elektroniskais_paraksts.isChecked()
        d.radit_elektronisko_parakstu_tekstu = self.ck_radit_elektronisko_parakstu_tekstu.isChecked() # JAUNA RINDAS

        # Piezīmes
        d.piezīmes = self.in_piezimes.toPlainText().strip()  # Tagad izsauc toPlainText() uz pielāgotā logrīka

        # Puses
        pie = Persona(
            nosaukums=self.pie_in[0].text().strip(),
            reģ_nr=self.pie_in[1].text().strip(),
            adrese=self.pie_in[2].text().strip(),
            kontaktpersona=self.pie_in[3].text().strip(),
            tālrunis=self.pie_in[4].text().strip(),
            epasts=self.pie_in[5].text().strip(),
            bankas_konts=self.pie_in[6].text().strip(),
            juridiskais_statuss=self.pie_in[7].currentText()
            )
        nod = Persona(
                nosaukums=self.nod_in[0].text().strip(),
                reģ_nr=self.nod_in[1].text().strip(),
                adrese=self.nod_in[2].text().strip(),
                kontaktpersona=self.nod_in[3].text().strip(),
                tālrunis=self.nod_in[4].text().strip(),
                epasts=self.nod_in[5].text().strip(),
                bankas_konts=self.nod_in[6].text().strip(),
                juridiskais_statuss=self.nod_in[7].currentText()
            )
        d.pieņēmējs = pie
        d.nodevējs = nod

        # Pozīcijas
        poz = []
        for r in range(self.tab.rowCount()):
                apr = self.tab.item(r, 0).text() if self.tab.item(r, 0) else ""
                daudz = to_decimal(self.tab.item(r, 1).text() if self.tab.item(r, 1) else "0")
                vien = self.tab.item(r, 2).text() if self.tab.item(r, 2) else ""
                cena = to_decimal(self.tab.item(r, 3).text() if self.tab.item(r, 3) else "0")
                ser_nr = self.tab.item(r, 5).text() if self.tab.item(r, 5) else ""
                gar = self.tab.item(r, 6).text() if self.tab.item(r, 6) else ""
                piez_poz = self.tab.item(r, 7).text() if self.tab.item(r, 7) else ""

                if not apr and daudz == 0 and not ser_nr and not gar and not piez_poz:
                    continue
                poz.append(
                    Pozīcija(apraksts=apr, daudzums=daudz, vienība=vien, cena=cena, seriālais_nr=ser_nr, garantija=gar,
                             piezīmes_pozīcijai=piez_poz))
                d.pozīcijas = poz

        # Iestatījumi
        d.iekļaut_pvn = self.ck_pvn.isChecked()
        d.pvn_likme = to_decimal(self.in_pvn.value())
        d.parakstu_rindas = self.ck_paraksti.isChecked()
        d.logotipa_ceļš = self.in_logo.text().strip()
        d.fonts_ceļš = self.in_fonts.text().strip()
        d.paraksts_pieņēmējs_ceļš = self.in_paraksts_pie.text().strip()
        d.paraksts_nodevējs_ceļš = self.in_paraksts_nod.text().strip()

        # Papildu iestatījumi
        d.pdf_page_size = self.cb_page_size.currentText()
        d.pdf_page_orientation = self.cb_page_orientation.currentText()
        d.pdf_margin_left = to_decimal(self.in_margin_left.value())
        d.pdf_margin_right = to_decimal(self.in_margin_right.value())
        d.pdf_margin_top = to_decimal(self.in_margin_top.value())
        d.pdf_margin_bottom = to_decimal(self.in_margin_bottom.value())
        d.pdf_font_size_head = self.in_font_size_head.value()
        d.pdf_font_size_normal = self.in_font_size_normal.value()
        d.pdf_font_size_small = self.in_font_size_small.value()
        d.pdf_font_size_table = self.in_font_size_table.value()
        d.pdf_logo_width_mm = to_decimal(self.in_logo_width_mm.value())
        d.pdf_signature_width_mm = to_decimal(self.in_signature_width_mm.value())
        d.pdf_signature_height_mm = to_decimal(self.in_signature_height_mm.value())
        d.docx_image_width_inches = to_decimal(self.in_docx_image_width_inches.value())
        d.docx_signature_width_inches = to_decimal(self.in_docx_signature_width_inches.value())
        d.table_col_widths = self.in_table_col_widths.toPlainText().strip()
        d.auto_generate_akta_nr = self.ck_auto_generate_akta_nr.isChecked()
        d.default_currency = self.in_default_currency.text().strip()
        d.default_unit = self.in_default_unit.text().strip()
        d.default_pvn_rate = to_decimal(self.in_default_pvn_rate.value())
        d.poppler_path = self.in_poppler_path.text().strip()

        # New settings
        d.header_text_color = self.in_header_text_color.text().strip()
        d.footer_text_color = self.in_footer_text_color.text().strip()
        d.table_header_bg_color = self.in_table_header_bg_color.text().strip()
        d.table_grid_color = self.in_table_grid_color.text().strip()
        d.table_row_spacing = to_decimal(self.in_table_row_spacing.value())
        d.line_spacing_multiplier = to_decimal(self.in_line_spacing_multiplier.value())
        d.show_page_numbers = self.ck_show_page_numbers.isChecked()
        d.show_generation_timestamp = self.ck_show_generation_timestamp.isChecked()
        d.currency_symbol_position = self.cb_currency_symbol_position.currentText()
        d.date_format = self.in_date_format.text().strip()
        d.signature_line_length_mm = to_decimal(self.in_signature_line_length_mm.value())
        d.signature_line_thickness_pt = to_decimal(self.in_signature_line_thickness_pt.value())
        d.add_cover_page = self.ck_add_cover_page.isChecked()
        d.cover_page_title = self.in_cover_page_title.text().strip()
        d.cover_page_logo_width_mm = to_decimal(self.in_cover_page_logo_width_mm.value())
        # Individuālais QR kods
        d.include_custom_qr_code = self.ck_include_custom_qr_code.isChecked()
        d.custom_qr_code_data = self.in_custom_qr_code_data.text().strip()
        d.custom_qr_code_size_mm = to_decimal(self.in_custom_qr_code_size_mm.value())
        d.custom_qr_code_position = self.cb_custom_qr_code_position.currentText()

        # Automātiskais QR kods (akta ID)
        d.include_auto_qr_code = self.ck_include_auto_qr_code.isChecked()
        d.auto_qr_code_size_mm = to_decimal(self.in_auto_qr_code_size_mm.value())
        d.auto_qr_code_position = self.cb_auto_qr_code_position.currentText()

        d.add_watermark = self.ck_add_watermark.isChecked()
        d.watermark_text = self.in_watermark_text.text().strip()
        d.watermark_font_size = self.in_watermark_font_size.value()
        d.watermark_color = self.in_watermark_color.text().strip()
        d.watermark_rotation = self.in_watermark_rotation.value()
        d.enable_pdf_encryption = self.ck_enable_pdf_encryption.isChecked()
        d.pdf_user_password = self.in_pdf_user_password.text().strip()
        d.pdf_owner_password = self.in_pdf_owner_password.text().strip()
        d.allow_printing = self.ck_allow_printing.isChecked()
        d.allow_copying = self.ck_allow_copying.isChecked()
        d.allow_modifying = self.ck_allow_modifying.isChecked()
        d.allow_annotating = self.ck_allow_annotating.isChecked()
        d.default_country = self.in_default_country.text().strip()
        d.default_city = self.in_default_city.text().strip()
        d.show_contact_details_in_header = self.ck_show_contact_details_in_header.isChecked()
        d.contact_details_header_font_size = self.in_contact_details_header_font_size.value()
        d.item_image_width_mm = to_decimal(self.in_item_image_width_mm.value())
        d.item_image_caption_font_size = self.in_item_image_caption_font_size.value()
        d.show_item_notes_in_table = self.ck_show_item_notes_in_table.isChecked()
        d.show_item_serial_number_in_table = self.ck_show_item_serial_number_in_table.isChecked()
        d.show_item_warranty_in_table = self.ck_show_item_warranty_in_table.isChecked()
        d.table_cell_padding_mm = to_decimal(self.in_table_cell_padding_mm.value())
        d.table_header_font_style = self.cb_table_header_font_style.currentText()
        d.table_content_alignment = self.cb_table_content_alignment.currentText()
        d.signature_font_size = self.in_signature_font_size.value()
        d.signature_spacing_mm = to_decimal(self.in_signature_spacing_mm.value())
        d.document_title_font_size = self.in_document_title_font_size.value()
        d.document_title_color = self.in_document_title_color.text().strip()
        d.section_heading_font_size = self.in_section_heading_font_size.value()
        d.section_heading_color = self.in_section_heading_color.text().strip()
        d.paragraph_line_spacing_multiplier = to_decimal(self.in_paragraph_line_spacing_multiplier.value())
        d.table_border_style = self.cb_table_border_style.currentText()
        d.table_border_thickness_pt = to_decimal(self.in_table_border_thickness_pt.value())
        d.table_alternate_row_color = self.in_table_alternate_row_color.text().strip()
        d.show_total_sum_in_words = self.ck_show_total_sum_in_words.isChecked()
        d.total_sum_in_words_language = self.cb_total_sum_in_words_language.currentText()
        d.default_vat_calculation_method = self.cb_default_vat_calculation_method.currentText()
        d.show_vat_breakdown = self.ck_show_vat_breakdown.isChecked()
        d.enable_digital_signature_field = self.ck_enable_digital_signature_field.isChecked()
        d.digital_signature_field_name = self.in_digital_signature_field_name.text().strip()
        d.digital_signature_field_size_mm = to_decimal(self.in_digital_signature_field_size_mm.value())
        d.digital_signature_field_position = self.cb_digital_signature_field_position.currentText()



        # Attēli
        att = []
        for i in range(self.img_list.count()):
                it = self.img_list.item(i)
                dat = it.data(Qt.UserRole)
                att.append(Attēls(ceļš=dat["ceļš"], paraksts=dat.get("paraksts", "")))
                d.attēli = att

        return d


    def izvēlēties_templates_dir(self):
        folder_path = QFileDialog.getExistingDirectory(self, "Izvēlēties šablonu direktoriju")
        if folder_path:
            self.in_templates_dir.setText(folder_path)
            # Pārliecināmies, ka jaunais direktorijs eksistē
            os.makedirs(folder_path, exist_ok=True)
            self._update_sablonu_list() # Atjaunojam šablonu sarakstu ar jauno direktoriju

    # ----- Tab: Puses -----

    def _persona_group(self, virsraksts: str, is_pieņēmējs: bool):
            box = QGroupBox(virsraksts)
            form = QFormLayout()
            nos = QLineEdit();
            reg = QLineEdit();
            adr = QLineEdit();
            kont = QLineEdit();
            tel = QLineEdit();
            ep = QLineEdit()
            bankas_konts = QLineEdit()
            juridiskais_statuss = QComboBox()
            juridiskais_statuss.addItems(["", "Juridiska persona", "Fiziska persona", "Pašnodarbinātais"])

            form.addRow("Nosaukums / Vārds, Uzvārds", nos)
            form.addRow("Reģ. Nr. / personas kods", reg)
            form.addRow("Adrese", adr)
            form.addRow("Kontaktpersona", kont)
            form.addRow("Tālrunis", tel)
            form.addRow("E-pasts", ep)
            form.addRow("Bankas konts", bankas_konts)
            form.addRow("Juridiskais statuss", juridiskais_statuss)

            btn_load_from_ab = QPushButton("Ielādēt no adrešu grāmatas")
            btn_save_to_ab = QPushButton("Saglabāt adrešu grāmatā")

            if is_pieņēmējs:
                btn_load_from_ab.clicked.connect(lambda: self._load_persona_from_address_book(self.pie_in))
                btn_save_to_ab.clicked.connect(lambda: self._save_persona_to_address_book(self.pie_in))
            else:
                btn_load_from_ab.clicked.connect(lambda: self._load_persona_from_address_book(self.nod_in))
                btn_save_to_ab.clicked.connect(lambda: self._save_persona_to_address_book(self.nod_in))

            btn_layout = QHBoxLayout()
            btn_layout.addWidget(btn_load_from_ab)
            btn_layout.addWidget(btn_save_to_ab)
            form.addRow(btn_layout)

            box.setLayout(form)
            return box, (
            nos, reg, adr, kont, tel, ep, bankas_konts, juridiskais_statuss)

    def _būvēt_puses_tab(self):
        content_widget = QWidget()
        lay = QHBoxLayout(content_widget)

        pieņēmējs_box, self.pie_in = self._persona_group("Pieņēmējs", True)
        nodevējs_box, self.nod_in = self._persona_group("Nodevējs", False)

        lay.addWidget(pieņēmējs_box)
        lay.addWidget(nodevējs_box)

        scroll_area = QScrollArea()
        scroll_area.setWidgetResizable(True)
        scroll_area.setWidget(content_widget)

        main_tab_widget = QWidget()
        main_layout = QVBoxLayout(main_tab_widget)
        main_layout.addWidget(scroll_area)

        self.tabs.addTab(main_tab_widget, "Puses")

        for line_edit in self.pie_in[:-1]:
            line_edit.textChanged.connect(self._update_preview)
        self.pie_in[-1].currentIndexChanged.connect(self._update_preview)

        for line_edit in self.nod_in[:-1]:
            line_edit.textChanged.connect(self._update_preview)
        self.nod_in[-1].currentIndexChanged.connect(self._update_preview)

    def _load_persona_from_address_book(self, persona_inputs):
        items = list(self.address_book.keys())
        if not items:
            QMessageBox.information(self, "Adrešu grāmata", "Adrešu grāmata ir tukša.")
            return
        item, ok = QInputDialog.getItem(self, "Ielādēt personu", "Izvēlieties personu:", items, 0, False)
        if ok and item:
            persona_data = self.address_book[item]
            persona_inputs[0].setText(persona_data.get("nosaukums", ""))
            persona_inputs[1].setText(persona_data.get("reģ_nr", ""))
            persona_inputs[2].setText(persona_data.get("adrese", ""))
            persona_inputs[3].setText(persona_data.get("kontaktpersona", ""))
            persona_inputs[4].setText(persona_data.get("tālrunis", ""))
            persona_inputs[5].setText(persona_data.get("epasts", ""))
            persona_inputs[6].setText(persona_data.get("bankas_konts", ""))
            juridiskais_statuss_val = persona_data.get("juridiskais_statuss", "")
            idx = persona_inputs[7].findText(juridiskais_statuss_val)
            if idx >= 0:
                persona_inputs[7].setCurrentIndex(idx)
            QMessageBox.information(self, "Ielādēts", f"Persona '{item}' ielādēta.")

    def _save_persona_to_address_book(self, persona_inputs):
        nosaukums = persona_inputs[0].text().strip()
        if not nosaukums:
            QMessageBox.warning(self, "Saglabāt personu", "Nosaukums nevar būt tukšs.")
            return
        persona_data = {
            "nosaukums": nosaukums,
            "reģ_nr": persona_inputs[1].text().strip(),
            "adrese": persona_inputs[2].text().strip(),
            "kontaktpersona": persona_inputs[3].text().strip(),
            "tālrunis": persona_inputs[4].text().strip(),
            "epasts": persona_inputs[5].text().strip(),
            "bankas_konts": persona_inputs[6].text().strip(),
            "juridiskais_statuss": persona_inputs[7].currentText(),
        }
        self.address_book[nosaukums] = persona_data
        self._save_address_book() # Saglabājam adrešu grāmatu failā
        self._update_address_book_list() # Atjaunojam sarakstu GUI

    def _load_address_book(self):
        if os.path.exists(ADDRESS_BOOK_FILE):
            try:
                # Mēģinām ielādēt ar dažādām kodēšanām
                encodings = ['utf-8', 'utf-8-sig', 'cp1257', 'iso-8859-1', 'windows-1252']
                for encoding in encodings:
                    try:
                        with open(ADDRESS_BOOK_FILE, 'r', encoding=encoding) as f:
                            self.address_book = json.load(f)
                        break
                    except (UnicodeDecodeError, json.JSONDecodeError):
                        continue
                else:
                    # Ja neviena kodēšana nedarbojas, izveidojam jaunu adrešu grāmatu
                    print(f"Neizdevās ielādēt adrešu grāmatas failu ar nevenu kodēšanu. Izveidojam jaunu.")
                    self.address_book = {}
            except Exception as e:
                QMessageBox.warning(self, "Kļūda", f"Neizdevās ielādēt adrešu grāmatu: {e}")
                self.address_book = {}
        else:
            self.address_book = {}

    def _save_address_book(self):
        os.makedirs(SETTINGS_DIR, exist_ok=True) # Izveidojam direktoriju, ja tā neeksistē
        try:
            with open(ADDRESS_BOOK_FILE, 'w', encoding='utf-8') as f:
                json.dump(self.address_book, f, ensure_ascii=False, indent=2)
        except Exception as e:
            QMessageBox.critical(self, "Kļūda", f"Neizdevās saglabāt adrešu grāmatu: {e}")

    # ----- Tab: Pozīcijas -----
    def _būvēt_pozīcijas_tab(self):
        w = QWidget()
        v = QVBoxLayout()

        # Adjust column count based on visibility settings
        initial_col_count = 5 # Apraksts, Daudzums, Vienība, Cena, Summa
        if self.data.show_item_serial_number_in_table: initial_col_count += 1
        if self.data.show_item_warranty_in_table: initial_col_count += 1
        if self.data.show_item_notes_in_table: initial_col_count += 1

        self.tab = QTableWidget(0, initial_col_count)
        headers = ["Apraksts", "Daudzums", "Vienība", "Cena", "Summa"]
        if self.data.show_item_serial_number_in_table: headers.append("Seriālais Nr.")
        if self.data.show_item_warranty_in_table: headers.append("Garantija")
        if self.data.show_item_notes_in_table: headers.append("Piezīmes pozīcijai")

        self.tab.setHorizontalHeaderLabels(headers)
        self.tab.horizontalHeader().setSectionResizeMode(0, QHeaderView.Stretch)
        for i in range(1, len(headers)):
            self.tab.horizontalHeader().setSectionResizeMode(i, QHeaderView.ResizeToContents)

        btns = QHBoxLayout()
        add = QPushButton("Pievienot pozīciju")
        add.clicked.connect(self.pievienot_pozīciju)
        dzest = QPushButton("Dzēst izvēlēto")
        dzest.clicked.connect(self.dzest_pozīciju)
        btns.addWidget(add)
        btns.addWidget(dzest)
        btns.addStretch()

        v.addLayout(btns)
        v.addWidget(self.tab)
        w.setLayout(v)
        self.tabs.addTab(w, "Pozīcijas")

        self.tab.cellChanged.connect(self._pārrēķināt_summa)
        self.tab.cellChanged.connect(self._update_preview)

    def pievienot_pozīciju(self):
        r = self.tab.rowCount()
        self.tab.insertRow(r)
        # Ensure all columns are initialized, even if not visible
        for c in range(self.tab.columnCount()):
            self.tab.setItem(r, c, QTableWidgetItem(""))
        self.tab.setItem(r, 1).setText("1")
        self.tab.setItem(r, 2).setText(self.data.default_unit)
        self.tab.setItem(r, 3).setText("0.00")
        self.tab.setItem(r, 4).setText("0.00")
        # Ensure items exist for non-visible columns if they are referenced
        if self.tab.columnCount() > 5: self.tab.setItem(r, 5, QTableWidgetItem("")) # Seriālais Nr.
        if self.tab.columnCount() > 6: self.tab.setItem(r, 6, QTableWidgetItem("")) # Garantija
        if self.tab.columnCount() > 7: self.tab.setItem(r, 7, QTableWidgetItem("")) # Piezīmes pozīcijai


    def dzest_pozīciju(self):
        r = self.tab.currentRow()
        if r >= 0:
            self.tab.removeRow(r)

    def _pārrēķināt_summa(self, row, col):
        # Check if the column is 'Daudzums' (index 1) or 'Cena' (index 3)
        if col in (1, 3):
            try:
                daudz = to_decimal(self.tab.item(row, 1).text())
            except Exception:
                daudz = Decimal("0")
            try:
                cena = to_decimal(self.tab.item(row, 3).text())
            except Exception:
                cena = Decimal("0")
            summa = (daudz * cena).quantize(Decimal("0.01"))
            # Ensure the item exists before setting text
            if not self.tab.item(row, 4):
                self.tab.setItem(row, 4, QTableWidgetItem(""))
            self.tab.item(row, 4).setText(formēt_naudu(summa)) # Labots: r -> row
        self._update_preview()

    # ----- Tab: Attēli -----
    def _būvēt_attēli_tab(self):
        w = QWidget()
        v = QVBoxLayout()

        self.img_list = QListWidget()
        v.addWidget(self.img_list)

        btns = QHBoxLayout()
        piev = QPushButton("Pievienot foto…")
        piev.clicked.connect(self.pievienot_attēlu)
        augsa = QPushButton("Uz augšu")
        augsa.clicked.connect(lambda: self.pārvietot_att(-1))
        leja = QPushButton("Uz leju")
        leja.clicked.connect(lambda: self.pārvietot_att(1))
        dzest = QPushButton("Dzēst")
        dzest.clicked.connect(self.dzest_att)
        btns.addWidget(piev)
        btns.addWidget(augsa)
        btns.addWidget(leja)
        btns.addWidget(dzest)
        btns.addStretch()

        v.addLayout(btns)
        w.setLayout(v)
        self.tabs.addTab(w, "Fotogrāfijas")

        self.img_list.itemDoubleClicked.connect(self.rediģēt_att_parakstu)
        self.img_list.itemChanged.connect(self._update_preview)
        self.img_list.model().rowsInserted.connect(self._update_preview)
        self.img_list.model().rowsRemoved.connect(self._update_preview)
        self.img_list.model().rowsMoved.connect(self._update_preview)

    def pievienot_attēlu(self):
        ceļi, _ = QFileDialog.getOpenFileNames(self, "Izvēlēties attēlus", "", "Attēli (*.png *.jpg *.jpeg *.webp)")
        for c in ceļi:
            it = QListWidgetItem(os.path.basename(c))
            it.setData(Qt.UserRole, {"ceļš": c, "paraksts": ""})
            self.img_list.addItem(it)

    def pārvietot_att(self, virziens):
        r = self.img_list.currentRow()
        if r < 0:
            return
        new_r = r + virziens
        if 0 <= new_r < self.img_list.count():
            item = self.img_list.takeItem(r)
            self.img_list.insertItem(new_r, item)
            self.img_list.setCurrentRow(new_r)
            self._update_preview()

    def dzest_att(self):
        r = self.img_list.currentRow()
        if r >= 0:
            self.img_list.takeItem(r)
            self._update_preview()

    def rediģēt_att_parakstu(self, item: QListWidgetItem):
        current_text = item.data(Qt.UserRole).get("paraksts", "")
        text, ok = QInputDialog.getText(self, "Rediģēt parakstu", "Ievadiet attēla parakstu:", QLineEdit.Normal, current_text)
        if ok and text is not None:
            d = item.data(Qt.UserRole)
            d["paraksts"] = text
            item.setData(Qt.UserRole, d)
            self._update_preview()

    # ----- Tab: Iestatījumi & Eksports -----
    def _būvēt_iestatījumi_tab(self):
        w = QWidget()
        form = QFormLayout()

        self.ck_pvn = QCheckBox("Aprēķināt PVN")
        self.in_pvn = QDoubleSpinBox(); self.in_pvn.setRange(0, 100); self.in_pvn.setValue(21.0); self.in_pvn.setSuffix(" %")

        self.ck_paraksti = QCheckBox("Iekļaut parakstu rindas")
        self.ck_paraksti.setChecked(True)

        self.in_logo = QLineEdit();
        logo_btn = QToolButton(); logo_btn.setText("…"); logo_btn.clicked.connect(self.izvēlēties_logo)
        logo_box = QHBoxLayout(); logo_box.addWidget(self.in_logo); logo_box.addWidget(logo_btn)
        logo_w = QWidget(); logo_w.setLayout(logo_box)

        self.in_fonts = QLineEdit();
        font_btn = QToolButton(); font_btn.setText("…"); font_btn.clicked.connect(self.izvēlēties_fontu)
        font_box = QHBoxLayout(); font_box.addWidget(self.in_fonts); font_box.addWidget(font_btn)
        font_w = QWidget(); font_w.setLayout(font_box)

        self.in_paraksts_pie = QLineEdit();
        btn_paraksts_pie = QToolButton(); btn_paraksts_pie.setText("…"); btn_paraksts_pie.clicked.connect(lambda: self.izvēlēties_paraksta_attēlu(self.in_paraksts_pie))
        box_paraksts_pie = QHBoxLayout(); box_paraksts_pie.addWidget(self.in_paraksts_pie); box_paraksts_pie.addWidget(btn_paraksts_pie)
        w_paraksts_pie = QWidget(); w_paraksts_pie.setLayout(box_paraksts_pie)

        self.in_paraksts_nod = QLineEdit();
        btn_paraksts_nod = QToolButton(); btn_paraksts_nod.setText("…"); btn_paraksts_nod.clicked.connect(lambda: self.izvēlēties_paraksta_attēlu(self.in_paraksts_nod))
        box_paraksts_nod = QHBoxLayout(); box_paraksts_nod.addWidget(self.in_paraksts_nod); box_paraksts_nod.addWidget(btn_paraksts_nod)
        w_paraksts_nod = QWidget(); w_paraksts_nod.setLayout(box_paraksts_nod)

        self.lang_combo = QComboBox()
        self.lang_combo.addItem("Latviešu")
        self.lang_combo.addItem("English (nav implementēts)")
        self.lang_combo.setEnabled(False)

        btn_saglabat_nokl = QPushButton("Saglabāt kā noklusējumu")
        btn_saglabat_nokl.clicked.connect(self.saglabat_noklusejuma_iestatijumus)

        # JAUNA POGA ŠABLONU SAGLABĀŠANAI
        btn_saglabat_sablonu = QPushButton("Saglabāt kā šablonu")
        btn_saglabat_sablonu.clicked.connect(self.saglabat_ka_sablonu)


        self.btn_generate_pdf = QPushButton("Ģenerēt PDF…")
        self.btn_generate_pdf.clicked.connect(self.ģenerēt_pdf_dialogs)

        self.btn_generate_docx = QPushButton("Ģenerēt DOCX…")
        self.btn_generate_docx.clicked.connect(self.ģenerēt_docx_dialogs)
        self.btn_print_pdf = QPushButton("Drukāt PDF…") # JAUNA RINDAS
        self.btn_print_pdf.clicked.connect(self.drukāt_pdf_dialogs) # JAUNA RINDAS



        form.addRow(self.ck_pvn, self.in_pvn)
        form.addRow(self.ck_paraksti)
        form.addRow("Logotips (neobligāti)", logo_w)
        form.addRow("Fonts TTF/OTF (ieteicams latviešu diakritikai)", font_w)
        form.addRow("Pieņēmēja paraksta attēls", w_paraksts_pie)
        form.addRow("Nodevēja paraksta attēls", w_paraksts_nod)
        form.addRow("Valoda", self.lang_combo)
        form.addRow(btn_saglabat_nokl)
        form.addRow(btn_saglabat_sablonu)  # JAUNA RINDAS
        form.addRow(self.btn_generate_pdf)
        form.addRow(self.btn_generate_docx)
        form.addRow(self.btn_print_pdf)  # JAUNA RINDAS

        self.ck_pvn.stateChanged.connect(self._update_preview)
        self.in_pvn.valueChanged.connect(self._update_preview)
        self.ck_paraksti.stateChanged.connect(self._update_preview)
        self.in_logo.textChanged.connect(self._update_preview)
        self.in_fonts.textChanged.connect(self._update_preview)
        self.in_paraksts_pie.textChanged.connect(self._update_preview)
        self.in_paraksts_nod.textChanged.connect(self._update_preview)


        w.setLayout(form)
        self.tabs.addTab(w, "Iestatījumi & Eksports")

    def izvēlēties_logo(self):
        c, _ = QFileDialog.getOpenFileName(self, "Izvēlēties logotipu", "", "Attēli (*.png *.jpg *.jpeg *.webp)")
        if c:
            self.in_logo.setText(c)

    def drukāt_pdf_dialogs(self):
        """
        Ģenerē PDF dokumentu un atver drukas priekšskatījuma dialogu.
        """
        akta_dati = self.savākt_datus()
        temp_pdf_path = None
        try:
            # Ģenerējam PDF pagaidu failā
            temp_pdf_path = ģenerēt_pdf(akta_dati, pdf_ceļš=None)

            if not os.path.exists(temp_pdf_path):
                QMessageBox.critical(self, "Drukas kļūda", "Neizdevās ģenerēt PDF failu drukāšanai.")
                return

            printer = QPrinter(QPrinter.HighResolution)
            preview_dialog = QPrintPreviewDialog(printer, self)
            preview_dialog.paintRequested.connect(lambda printer_obj: self._render_pdf_to_printer(temp_pdf_path, printer_obj))
            preview_dialog.exec()

        except Exception as e:
            QMessageBox.critical(self, "Drukas kļūda", f"Kļūda sagatavojot dokumentu drukāšanai: {e}")
        finally:
            if temp_pdf_path and os.path.exists(temp_pdf_path):
                try:
                    os.remove(temp_pdf_path) # Dzēšam pagaidu failu
                except Exception as e:
                    print(f"Neizdevās dzēst pagaidu PDF failu: {e}")

    def _render_pdf_to_printer(self, pdf_path: str, printer: QPrinter):
        """
        Renderē PDF failu uz printera.
        Izmanto pdf2image, lai iegūtu attēlus no PDF un zīmē tos uz printera.
        """
        try:
            # Pārliecināmies, ka poppler_path ir pieejams
            poppler_path_to_use = self.data.poppler_path if self.data.poppler_path and os.path.exists(self.data.poppler_path) else None

            images = convert_from_path(pdf_path, poppler_path=poppler_path_to_use)

            if not images:
                QMessageBox.warning(self, "Drukas kļūda", "Neizdevās iegūt attēlus no PDF faila drukāšanai.")
                return

            painter = QPainter()
            if not painter.begin(printer):
                QMessageBox.critical(self, "Drukas kļūda", "Neizdevās sākt zīmēšanu uz printera.")
                return

            for i, pil_img in enumerate(images):
                if i > 0:
                    printer.newPage()  # Jauna lapa katram attēlam (PDF lapai)

                # Konvertējam PIL attēlu uz QImage
                q_image = ImageQt(pil_img)
                pixmap = QPixmap.fromImage(q_image)

                # Mērogojam attēlu, lai tas ietilptu printera lapā
                # printer.pageRect() atgriež lapas izmērus pikseļos, ņemot vērā printera izšķirtspēju
                # Labots: Izmantojam painter.device().width() un painter.device().height()
                # lai iegūtu zīmēšanas ierīces (printera) izmērus.
                printer_width = painter.device().width()
                printer_height = painter.device().height()

                # Izmantojam QSize, lai scaled() metodei nodotu pareizu izmēru
                target_size = QSize(printer_width, printer_height)
                scaled_pixmap = pixmap.scaled(target_size, Qt.KeepAspectRatio, Qt.SmoothTransformation)

                # Zīmējam attēlu lapas centrā
                # Labots: Izmantojam printera izmērus, lai centrētu attēlu
                x = (printer_width - scaled_pixmap.width()) / 2
                y = (printer_height - scaled_pixmap.height()) / 2
                painter.drawPixmap(int(x), int(y), scaled_pixmap)

            painter.end()
            QMessageBox.information(self, "Drukas priekšskatījums", "Dokuments sagatavots drukāšanai.")

        except Exception as e:
            QMessageBox.critical(self, "Drukas kļūda", f"Kļūda renderējot PDF uz printeri: {e}")
        finally:
            # Pārliecināmies, ka painter tiek beigts, pat ja rodas kļūda
            if painter.isActive():
                painter.end()



    def izvēlēties_fontu(self):
        c, _ = QFileDialog.getOpenFileName(self, "Izvēlieties fontu (TTF/OTF)", "", "Fonti (*.ttf *.otf)")
        if c:
            self.in_fonts.setText(c)

    def izvēlēties_paraksta_attēlu(self, line_edit: QLineEdit):
        c, _ = QFileDialog.getOpenFileName(self, "Izvēlieties paraksta attēlu", "", "Attēli (*.png *.jpg *.jpeg *.webp)")
        if c:
            line_edit.setText(c)

    # ----- Tab: Papildu iestatījumi -----
    def _būvēt_papildu_iestatījumi_tab(self):
        w = QWidget()
        form = QFormLayout()

        self.cb_page_size = QComboBox()
        self.cb_page_size.addItems(["A4", "Letter", "Legal", "A3", "A5"])
        self.cb_page_orientation = QComboBox()
        self.cb_page_orientation.addItems(["Portrets", "Ainava"])

        self.in_margin_left = QDoubleSpinBox(); self.in_margin_left.setRange(0, 100); self.in_margin_left.setSuffix(" mm")
        self.in_margin_right = QDoubleSpinBox(); self.in_margin_right.setRange(0, 100); self.in_margin_right.setSuffix(" mm")
        self.in_margin_top = QDoubleSpinBox(); self.in_margin_top.setRange(0, 100); self.in_margin_top.setSuffix(" mm")
        self.in_margin_bottom = QDoubleSpinBox(); self.in_margin_bottom.setRange(0, 100); self.in_margin_bottom.setSuffix(" mm")

        self.in_font_size_head = QSpinBox(); self.in_font_size_head.setRange(8, 30)
        self.in_font_size_normal = QSpinBox(); self.in_font_size_normal.setRange(6, 20)
        self.in_font_size_small = QSpinBox(); self.in_font_size_small.setRange(4, 16)
        self.in_font_size_table = QSpinBox(); self.in_font_size_table.setRange(4, 16)

        self.in_logo_width_mm = QDoubleSpinBox(); self.in_logo_width_mm.setRange(10, 100); self.in_logo_width_mm.setSuffix(" mm")
        self.in_signature_width_mm = QDoubleSpinBox(); self.in_signature_width_mm.setRange(10, 100); self.in_signature_width_mm.setSuffix(" mm")
        self.in_signature_height_mm = QDoubleSpinBox(); self.in_signature_height_mm.setRange(5, 50); self.in_signature_height_mm.setSuffix(" mm")

        self.in_docx_image_width_inches = QDoubleSpinBox(); self.in_docx_image_width_inches.setRange(1, 10); self.in_docx_image_width_inches.setSuffix(" collas")
        self.in_docx_signature_width_inches = QDoubleSpinBox(); self.in_docx_signature_width_inches.setRange(0.5, 5); self.in_docx_signature_width_inches.setSuffix(" collas")

        self.in_table_col_widths = QTextEdit()
        self.in_table_col_widths.setPlaceholderText("Ievadiet kolonnu platumus mm, atdalot ar komatiem (piem., 10,40,18,18,20,20,25,25,25)")

        self.ck_auto_generate_akta_nr = QCheckBox("Automātiski ģenerēt akta numuru (PP-YYYY-NNNN)")

        self.in_default_currency = QLineEdit()
        self.in_default_unit = QLineEdit()
        self.in_default_pvn_rate = QDoubleSpinBox(); self.in_default_pvn_rate.setRange(0, 100); self.in_default_pvn_rate.setSuffix(" %")

        self.in_poppler_path = QLineEdit()
        poppler_btn = QToolButton(); poppler_btn.setText("…"); poppler_btn.clicked.connect(self.izvēlēties_poppler_ceļu)
        poppler_box = QHBoxLayout(); poppler_box.addWidget(self.in_poppler_path); poppler_box.addWidget(poppler_btn)
        poppler_w = QWidget(); poppler_w.setLayout(poppler_box)

        # New settings (30+ functions/options/settings)
        self.in_header_text_color = QLineEdit("#000000")
        self.in_footer_text_color = QLineEdit("#000000")
        self.in_table_header_bg_color = QLineEdit("#E0E0E0")
        self.in_table_grid_color = QLineEdit("#CCCCCC")
        self.in_table_row_spacing = QDoubleSpinBox(); self.in_table_row_spacing.setRange(0, 10); self.in_table_row_spacing.setValue(4); self.in_table_row_spacing.setSuffix(" mm")
        self.in_line_spacing_multiplier = QDoubleSpinBox(); self.in_line_spacing_multiplier.setRange(0.5, 3.0); self.in_line_spacing_multiplier.setSingleStep(0.1); self.in_line_spacing_multiplier.setValue(1.2)
        self.ck_show_page_numbers = QCheckBox("Rādīt lapu numurus"); self.ck_show_page_numbers.setChecked(True)
        self.ck_show_generation_timestamp = QCheckBox("Rādīt ģenerēšanas laiku"); self.ck_show_generation_timestamp.setChecked(True)
        self.cb_currency_symbol_position = QComboBox(); self.cb_currency_symbol_position.addItems(["after", "before"])
        self.in_date_format = QLineEdit("YYYY-MM-DD")
        self.in_signature_line_length_mm = QDoubleSpinBox(); self.in_signature_line_length_mm.setRange(10, 100); self.in_signature_line_length_mm.setValue(60); self.in_signature_line_length_mm.setSuffix(" mm")
        self.in_signature_line_thickness_pt = QDoubleSpinBox(); self.in_signature_line_thickness_pt.setRange(0.1, 5.0); self.in_signature_line_thickness_pt.setSingleStep(0.1); self.in_signature_line_thickness_pt.setValue(0.5); self.in_signature_line_thickness_pt.setSuffix(" pt")
        self.ck_add_cover_page = QCheckBox("Pievienot titullapu")
        self.in_cover_page_title = QLineEdit("Pieņemšanas-Nodošanas Akts")
        self.in_cover_page_logo_width_mm = QDoubleSpinBox(); self.in_cover_page_logo_width_mm.setRange(10, 200); self.in_cover_page_logo_width_mm.setValue(80); self.in_cover_page_logo_width_mm.setSuffix(" mm")
        # Individuālais QR kods
        self.ck_include_custom_qr_code = QCheckBox("Iekļaut individuālu QR kodu")
        self.in_custom_qr_code_data = QLineEdit()
        self.in_custom_qr_code_data.setPlaceholderText("Dati individuālajam QR kodam (URL, teksts utt.)")
        self.in_custom_qr_code_size_mm = QDoubleSpinBox();
        self.in_custom_qr_code_size_mm.setRange(10, 50);
        self.in_custom_qr_code_size_mm.setValue(20);
        self.in_custom_qr_code_size_mm.setSuffix(" mm")
        self.cb_custom_qr_code_position = QComboBox();
        self.cb_custom_qr_code_position.addItems(["bottom_right", "bottom_left", "top_right", "top_left", "custom"])
        self.in_custom_qr_code_pos_x_mm = QDoubleSpinBox();
        self.in_custom_qr_code_pos_x_mm.setRange(0, 500);
        self.in_custom_qr_code_pos_x_mm.setSuffix(" mm")
        self.in_custom_qr_code_pos_y_mm = QDoubleSpinBox();
        self.in_custom_qr_code_pos_y_mm.setRange(0, 500);
        self.in_custom_qr_code_pos_y_mm.setSuffix(" mm")
        self.in_custom_qr_code_color = QLineEdit("#000000")  # QR koda krāsa (Hex)

        # Automātiskais QR kods (akta ID)
        self.ck_include_auto_qr_code = QCheckBox("Iekļaut automātisku QR kodu (Akta ID)")
        self.in_auto_qr_code_size_mm = QDoubleSpinBox();
        self.in_auto_qr_code_size_mm.setRange(10, 50);
        self.in_auto_qr_code_size_mm.setValue(20);
        self.in_auto_qr_code_size_mm.setSuffix(" mm")
        self.cb_auto_qr_code_position = QComboBox();
        self.cb_auto_qr_code_position.addItems(["bottom_left", "bottom_right", "top_right", "top_left", "custom"])
        self.in_auto_qr_code_pos_x_mm = QDoubleSpinBox();
        self.in_auto_qr_code_pos_x_mm.setRange(0, 500);
        self.in_auto_qr_code_pos_x_mm.setSuffix(" mm")
        self.in_auto_qr_code_pos_y_mm = QDoubleSpinBox();
        self.in_auto_qr_code_pos_y_mm.setRange(0, 500);
        self.in_auto_qr_code_pos_y_mm.setSuffix(" mm")
        self.in_auto_qr_code_color = QLineEdit("#000000")  # QR koda krāsa (Hex)

        self.ck_add_watermark = QCheckBox("Pievienot ūdenszīmi")
        self.in_watermark_text = QLineEdit("MELNRAKSTS")
        self.in_watermark_font_size = QSpinBox(); self.in_watermark_font_size.setRange(10, 200); self.in_watermark_font_size.setValue(72)
        self.in_watermark_color = QLineEdit("#E0E0E0")
        self.in_watermark_rotation = QSpinBox(); self.in_watermark_rotation.setRange(0, 360); self.in_watermark_rotation.setValue(45)
        self.ck_enable_pdf_encryption = QCheckBox("Iespējot PDF šifrēšanu")
        self.in_pdf_user_password = QLineEdit(); self.in_pdf_user_password.setEchoMode(QLineEdit.Password)
        self.in_pdf_owner_password = QLineEdit(); self.in_pdf_owner_password.setEchoMode(QLineEdit.Password)
        self.ck_allow_printing = QCheckBox("Atļaut drukāšanu"); self.ck_allow_printing.setChecked(True)
        self.ck_allow_copying = QCheckBox("Atļaut kopēšanu"); self.ck_allow_copying.setChecked(True)
        self.ck_allow_modifying = QCheckBox("Atļaut modificēšanu")
        self.ck_allow_annotating = QCheckBox("Atļaut anotēšanu"); self.ck_allow_annotating.setChecked(True)
        self.in_default_country = QLineEdit("Latvija")
        self.in_default_city = QLineEdit("Rīga")
        self.ck_show_contact_details_in_header = QCheckBox("Rādīt kontaktinformāciju galvenē")
        self.in_contact_details_header_font_size = QSpinBox(); self.in_contact_details_header_font_size.setRange(6, 12); self.in_contact_details_header_font_size.setValue(8)
        self.in_item_image_width_mm = QDoubleSpinBox(); self.in_item_image_width_mm.setRange(10, 150); self.in_item_image_width_mm.setValue(50); self.in_item_image_width_mm.setSuffix(" mm")
        self.in_item_image_caption_font_size = QSpinBox(); self.in_item_image_caption_font_size.setRange(6, 12); self.in_item_image_caption_font_size.setValue(8)
        self.ck_show_item_notes_in_table = QCheckBox("Rādīt pozīciju piezīmes tabulā"); self.ck_show_item_notes_in_table.setChecked(True)
        self.ck_show_item_serial_number_in_table = QCheckBox("Rādīt pozīciju sērijas Nr. tabulā"); self.ck_show_item_serial_number_in_table.setChecked(True)
        self.ck_show_item_warranty_in_table = QCheckBox("Rādīt pozīciju garantiju tabulā"); self.ck_show_item_warranty_in_table.setChecked(True)
        self.in_table_cell_padding_mm = QDoubleSpinBox(); self.in_table_cell_padding_mm.setRange(0, 10); self.in_table_cell_padding_mm.setValue(2); self.in_table_cell_padding_mm.setSuffix(" mm")
        self.cb_table_header_font_style = QComboBox(); self.cb_table_header_font_style.addItems(["bold", "italic", "normal"])
        self.cb_table_content_alignment = QComboBox(); self.cb_table_content_alignment.addItems(["left", "center", "right"])
        self.in_signature_font_size = QSpinBox(); self.in_signature_font_size.setRange(6, 12); self.in_signature_font_size.setValue(9)
        self.in_signature_spacing_mm = QDoubleSpinBox(); self.in_signature_spacing_mm.setRange(0, 20); self.in_signature_spacing_mm.setValue(10); self.in_signature_spacing_mm.setSuffix(" mm")
        self.in_document_title_font_size = QSpinBox(); self.in_document_title_font_size.setRange(10, 30); self.in_document_title_font_size.setValue(18)
        self.in_document_title_color = QLineEdit("#000000")
        self.in_section_heading_font_size = QSpinBox(); self.in_section_heading_font_size.setRange(8, 20); self.in_section_heading_font_size.setValue(12)
        self.in_section_heading_color = QLineEdit("#000000")
        self.in_paragraph_line_spacing_multiplier = QDoubleSpinBox(); self.in_paragraph_line_spacing_multiplier.setRange(0.5, 3.0); self.in_paragraph_line_spacing_multiplier.setSingleStep(0.1); self.in_paragraph_line_spacing_multiplier.setValue(1.2)
        self.cb_table_border_style = QComboBox(); self.cb_table_border_style.addItems(["solid", "dashed", "none"])
        self.in_table_border_thickness_pt = QDoubleSpinBox(); self.in_table_border_thickness_pt.setRange(0.1, 5.0); self.in_table_border_thickness_pt.setSingleStep(0.1); self.in_table_border_thickness_pt.setValue(0.5); self.in_table_border_thickness_pt.setSuffix(" pt")
        self.in_table_alternate_row_color = QLineEdit("")
        self.in_table_alternate_row_color.setPlaceholderText("Hex krāsa, piem. #F0F0F0")
        self.ck_show_total_sum_in_words = QCheckBox("Rādīt kopsummu vārdos")
        self.cb_total_sum_in_words_language = QComboBox(); self.cb_total_sum_in_words_language.addItems(["lv", "en"])
        self.cb_default_vat_calculation_method = QComboBox(); self.cb_default_vat_calculation_method.addItems(["exclusive", "inclusive"])
        self.ck_show_vat_breakdown = QCheckBox("Rādīt PVN sadalījumu"); self.ck_show_vat_breakdown.setChecked(True)
        self.ck_enable_digital_signature_field = QCheckBox("Iespējot digitālā paraksta lauku (PDF)");
        self.in_digital_signature_field_name = QLineEdit("Paraksts");
        self.in_digital_signature_field_size_mm = QDoubleSpinBox();
        self.in_digital_signature_field_size_mm.setRange(10, 100);
        self.in_digital_signature_field_size_mm.setValue(40);
        self.in_digital_signature_field_size_mm.setSuffix(" mm")
        self.cb_digital_signature_field_position = QComboBox();
        self.cb_digital_signature_field_position.addItems(
            ["bottom_center", "bottom_left", "bottom_right", "top_left", "top_right"])

        # JAUNA RINDAS - Šablonu direktorija iestatījums
        self.in_templates_dir = QLineEdit()
        self.in_templates_dir.setText(os.path.join(APP_DATA_DIR, "AktaGenerators_Templates"))  # Noklusējuma vērtība
        btn_templates_dir = QToolButton();
        btn_templates_dir.setText("…");
        btn_templates_dir.clicked.connect(self.izvēlēties_templates_dir)
        templates_dir_box = QHBoxLayout();
        templates_dir_box.addWidget(self.in_templates_dir);
        templates_dir_box.addWidget(btn_templates_dir)
        templates_dir_w = QWidget();
        templates_dir_w.setLayout(templates_dir_box)

        form.addRow("PDF lapas izmērs:", self.cb_page_size)
        form.addRow("PDF lapas orientācija:", self.cb_page_orientation)
        form.addRow("PDF kreisā mala (mm):", self.in_margin_left)
        form.addRow("PDF labā mala (mm):", self.in_margin_right)
        form.addRow("PDF augšējā mala (mm):", self.in_margin_top)
        form.addRow("PDF apakšējā mala (mm):", self.in_margin_bottom)
        form.addRow("PDF galvenes fonta izmērs:", self.in_font_size_head)
        form.addRow("PDF normāla fonta izmērs:", self.in_font_size_normal)
        form.addRow("PDF maza fonta izmērs:", self.in_font_size_small)
        form.addRow("PDF tabulas fonta izmērs:", self.in_font_size_table)
        form.addRow("PDF logo platums (mm):", self.in_logo_width_mm)
        form.addRow("PDF paraksta attēla platums (mm):", self.in_signature_width_mm)
        form.addRow("PDF paraksta attēla augstums (mm):", self.in_signature_height_mm)
        form.addRow("DOCX attēlu platums (collas):", self.in_docx_image_width_inches)
        form.addRow("DOCX paraksta attēla platums (collas):", self.in_docx_signature_width_inches)
        form.addRow("Pozīciju tabulas kolonnu platumi (mm, komatiem atdalīti):", self.in_table_col_widths)
        form.addRow(self.ck_auto_generate_akta_nr)
        form.addRow("Noklusējuma valūta:", self.in_default_currency)
        form.addRow("Noklusējuma vienība:", self.in_default_unit)
        form.addRow("Noklusējuma PVN likme (%):", self.in_default_pvn_rate)
        form.addRow("Poppler bin direktorijas ceļš (Windows):", poppler_w)

        # Add new settings to the form
        form.addRow("Galvenes teksta krāsa (Hex):", self.in_header_text_color)
        form.addRow("Kājenes teksta krāsa (Hex):", self.in_footer_text_color)
        form.addRow("Tabulas galvenes fona krāsa (Hex):", self.in_table_header_bg_color)
        form.addRow("Tabulas režģa krāsa (Hex):", self.in_table_grid_color)
        form.addRow("Tabulas rindu atstarpe (mm):", self.in_table_row_spacing)
        form.addRow("Rindu atstarpes reizinātājs:", self.in_line_spacing_multiplier)
        form.addRow(self.ck_show_page_numbers)
        form.addRow(self.ck_show_generation_timestamp)
        form.addRow("Valūtas simbola pozīcija:", self.cb_currency_symbol_position)
        form.addRow("Datuma formāts:", self.in_date_format)
        form.addRow("Paraksta līnijas garums (mm):", self.in_signature_line_length_mm)
        form.addRow("Paraksta līnijas biezums (pt):", self.in_signature_line_thickness_pt)
        form.addRow(self.ck_add_cover_page)
        form.addRow("Titullapas virsraksts:", self.in_cover_page_title)
        form.addRow("Titullapas logo platums (mm):", self.in_cover_page_logo_width_mm)

        # Individuālais QR kods
        form.addRow(self.ck_include_custom_qr_code)
        form.addRow("Individuālā QR koda dati:", self.in_custom_qr_code_data)
        form.addRow("Individuālā QR koda izmērs (mm):", self.in_custom_qr_code_size_mm)
        form.addRow("Individuālā QR koda pozīcija:", self.cb_custom_qr_code_position)
        form.addRow("Individuālā QR koda X pozīcija (mm):", self.in_custom_qr_code_pos_x_mm)
        form.addRow("Individuālā QR koda Y pozīcija (mm):", self.in_custom_qr_code_pos_y_mm)
        form.addRow("Individuālā QR koda krāsa (Hex):", self.in_custom_qr_code_color)

        # Automātiskais QR kods (akta ID)
        form.addRow(self.ck_include_auto_qr_code)
        form.addRow("Automātiskā QR koda izmērs (mm):", self.in_auto_qr_code_size_mm)
        form.addRow("Automātiskā QR koda pozīcija:", self.cb_auto_qr_code_position)
        form.addRow("Automātiskā QR koda X pozīcija (mm):", self.in_auto_qr_code_pos_x_mm)
        form.addRow("Automātiskā QR koda Y pozīcija (mm):", self.in_auto_qr_code_pos_y_mm)
        form.addRow("Automātiskā QR koda krāsa (Hex):", self.in_auto_qr_code_color)

        form.addRow(self.ck_add_watermark)
        form.addRow("Ūdenszīmes teksts:", self.in_watermark_text)
        form.addRow("Ūdenszīmes fonta izmērs:", self.in_watermark_font_size)
        form.addRow("Ūdenszīmes krāsa (Hex):", self.in_watermark_color)
        form.addRow("Ūdenszīmes rotācija (grādi):", self.in_watermark_rotation)
        form.addRow(self.ck_enable_pdf_encryption)
        form.addRow("PDF lietotāja parole:", self.in_pdf_user_password)
        form.addRow("PDF īpašnieka parole:", self.in_pdf_owner_password)
        form.addRow(self.ck_allow_printing)
        form.addRow(self.ck_allow_copying)
        form.addRow(self.ck_allow_modifying)
        form.addRow(self.ck_allow_annotating)
        form.addRow("Noklusējuma valsts:", self.in_default_country)
        form.addRow("Noklusējuma pilsēta:", self.in_default_city)
        form.addRow(self.ck_show_contact_details_in_header)
        form.addRow("Kontaktu detaļu galvenes fonta izmērs:", self.in_contact_details_header_font_size)
        form.addRow("Pozīciju attēlu platums (mm):", self.in_item_image_width_mm)
        form.addRow("Pozīciju attēlu paraksta fonta izmērs:", self.in_item_image_caption_font_size)
        form.addRow(self.ck_show_item_notes_in_table)
        form.addRow(self.ck_show_item_serial_number_in_table)
        form.addRow(self.ck_show_item_warranty_in_table)
        form.addRow("Tabulas šūnu polsterējums (mm):", self.in_table_cell_padding_mm)
        form.addRow("Tabulas galvenes fonta stils:", self.cb_table_header_font_style)
        form.addRow("Tabulas satura izlīdzināšana:", self.cb_table_content_alignment)
        form.addRow("Paraksta fonta izmērs:", self.in_signature_font_size)
        form.addRow("Paraksta atstarpe (mm):", self.in_signature_spacing_mm)
        form.addRow("Dokumenta virsraksta fonta izmērs:", self.in_document_title_font_size)
        form.addRow("Dokumenta virsraksta krāsa (Hex):", self.in_document_title_color)
        form.addRow("Sadaļas virsraksta fonta izmērs:", self.in_section_heading_font_size)
        form.addRow("Sadaļas virsraksta krāsa (Hex):", self.in_section_heading_color)
        form.addRow("Paragrāfa rindu atstarpes reizinātājs:", self.in_paragraph_line_spacing_multiplier)
        form.addRow("Tabulas apmales stils:", self.cb_table_border_style)
        form.addRow("Tabulas apmales biezums (pt):", self.in_table_border_thickness_pt)
        form.addRow("Tabulas alternatīvās rindas krāsa (Hex):", self.in_table_alternate_row_color)
        form.addRow(self.ck_show_total_sum_in_words)
        form.addRow("Kopsummas vārdos valoda:", self.cb_total_sum_in_words_language)
        form.addRow("Noklusējuma PVN aprēķina metode:", self.cb_default_vat_calculation_method)
        form.addRow(self.ck_show_vat_breakdown)
        form.addRow(self.ck_enable_digital_signature_field)
        form.addRow("Digitālā paraksta lauka nosaukums:", self.in_digital_signature_field_name)
        form.addRow("Digitālā paraksta lauka izmērs (mm):", self.in_digital_signature_field_size_mm)
        form.addRow("Digitālā paraksta lauka pozīcija:", self.cb_digital_signature_field_position)
        form.addRow("Šablonu direktorijs:", templates_dir_w)  # JAUNA RINDAS

        w.setLayout(form)
        scroll_area = QScrollArea()
        scroll_area.setWidgetResizable(True)
        scroll_area.setWidget(w)
        self.tabs.addTab(scroll_area, "Papildu iestatījumi")

        # Connect new settings to update_preview
        self.in_header_text_color.textChanged.connect(self._update_preview)
        self.in_footer_text_color.textChanged.connect(self._update_preview)
        self.in_table_header_bg_color.textChanged.connect(self._update_preview)
        self.in_table_grid_color.textChanged.connect(self._update_preview)
        self.in_table_row_spacing.valueChanged.connect(self._update_preview)
        self.in_line_spacing_multiplier.valueChanged.connect(self._update_preview)
        self.ck_show_page_numbers.stateChanged.connect(self._update_preview)
        self.ck_show_generation_timestamp.stateChanged.connect(self._update_preview)
        self.cb_currency_symbol_position.currentIndexChanged.connect(self._update_preview)
        self.in_date_format.textChanged.connect(self._update_preview)
        self.in_signature_line_length_mm.valueChanged.connect(self._update_preview)
        self.in_signature_line_thickness_pt.valueChanged.connect(self._update_preview)
        self.ck_add_cover_page.stateChanged.connect(self._update_preview)
        self.in_cover_page_title.textChanged.connect(self._update_preview)
        self.in_cover_page_logo_width_mm.valueChanged.connect(self._update_preview)
        # Individuālais QR kods
        self.ck_include_custom_qr_code.stateChanged.connect(self._update_preview)
        self.in_custom_qr_code_data.textChanged.connect(self._update_preview)
        self.in_custom_qr_code_size_mm.valueChanged.connect(self._update_preview)
        self.cb_custom_qr_code_position.currentIndexChanged.connect(self._update_preview)
        self.in_custom_qr_code_pos_x_mm.valueChanged.connect(self._update_preview)
        self.in_custom_qr_code_pos_y_mm.valueChanged.connect(self._update_preview)
        self.in_custom_qr_code_color.textChanged.connect(self._update_preview)

        # Automātiskais QR kods (akta ID)
        self.ck_include_auto_qr_code.stateChanged.connect(self._update_preview)
        self.in_auto_qr_code_size_mm.valueChanged.connect(self._update_preview)
        self.cb_auto_qr_code_position.currentIndexChanged.connect(self._update_preview)
        self.in_auto_qr_code_pos_x_mm.valueChanged.connect(self._update_preview)
        self.in_auto_qr_code_pos_y_mm.valueChanged.connect(self._update_preview)
        self.in_auto_qr_code_color.textChanged.connect(self._update_preview)

        self.ck_add_watermark.stateChanged.connect(self._update_preview)
        self.in_watermark_text.textChanged.connect(self._update_preview)
        self.in_watermark_font_size.valueChanged.connect(self._update_preview)
        self.in_watermark_color.textChanged.connect(self._update_preview)
        self.in_watermark_rotation.valueChanged.connect(self._update_preview)
        self.ck_enable_pdf_encryption.stateChanged.connect(self._update_preview)
        self.in_pdf_user_password.textChanged.connect(self._update_preview)
        self.in_pdf_owner_password.textChanged.connect(self._update_preview)
        self.ck_allow_printing.stateChanged.connect(self._update_preview)
        self.ck_allow_copying.stateChanged.connect(self._update_preview)
        self.ck_allow_modifying.stateChanged.connect(self._update_preview)
        self.ck_allow_annotating.stateChanged.connect(self._update_preview)
        self.in_default_country.textChanged.connect(self._update_preview)
        self.in_default_city.textChanged.connect(self._update_preview)
        self.ck_show_contact_details_in_header.stateChanged.connect(self._update_preview)
        self.in_contact_details_header_font_size.valueChanged.connect(self._update_preview)
        self.in_item_image_width_mm.valueChanged.connect(self._update_preview)
        self.in_item_image_caption_font_size.valueChanged.connect(self._update_preview)
        self.ck_show_item_notes_in_table.stateChanged.connect(self._update_preview)
        self.ck_show_item_serial_number_in_table.stateChanged.connect(self._update_preview)
        self.ck_show_item_warranty_in_table.stateChanged.connect(self._update_preview)
        self.in_table_cell_padding_mm.valueChanged.connect(self._update_preview)
        self.cb_table_header_font_style.currentIndexChanged.connect(self._update_preview)
        self.cb_table_content_alignment.currentIndexChanged.connect(self._update_preview)
        self.in_signature_font_size.valueChanged.connect(self._update_preview)
        self.in_signature_spacing_mm.valueChanged.connect(self._update_preview)
        self.in_document_title_font_size.valueChanged.connect(self._update_preview)
        self.in_document_title_color.textChanged.connect(self._update_preview)
        self.in_section_heading_font_size.valueChanged.connect(self._update_preview)
        self.in_section_heading_color.textChanged.connect(self._update_preview)
        self.in_paragraph_line_spacing_multiplier.valueChanged.connect(self._update_preview)
        self.cb_table_border_style.currentIndexChanged.connect(self._update_preview)
        self.in_table_border_thickness_pt.valueChanged.connect(self._update_preview)
        self.in_table_alternate_row_color.textChanged.connect(self._update_preview)
        self.ck_show_total_sum_in_words.stateChanged.connect(self._update_preview)
        self.cb_total_sum_in_words_language.currentIndexChanged.connect(self._update_preview)
        self.cb_default_vat_calculation_method.currentIndexChanged.connect(self._update_preview)
        self.ck_show_vat_breakdown.stateChanged.connect(self._update_preview)
        self.ck_enable_digital_signature_field.stateChanged.connect(self._update_preview)
        self.in_digital_signature_field_name.textChanged.connect(self._update_preview)
        self.in_digital_signature_field_size_mm.valueChanged.connect(self._update_preview)
        self.cb_digital_signature_field_position.currentIndexChanged.connect(self._update_preview)


        self.cb_page_size.currentIndexChanged.connect(self._update_preview)
        self.cb_page_orientation.currentIndexChanged.connect(self._update_preview)
        self.in_margin_left.valueChanged.connect(self._update_preview)
        self.in_margin_right.valueChanged.connect(self._update_preview)
        self.in_margin_top.valueChanged.connect(self._update_preview)
        self.in_margin_bottom.valueChanged.connect(self._update_preview)
        self.in_font_size_head.valueChanged.connect(self._update_preview)
        self.in_font_size_normal.valueChanged.connect(self._update_preview)
        self.in_font_size_small.valueChanged.connect(self._update_preview)
        self.in_font_size_table.valueChanged.connect(self._update_preview)
        self.in_logo_width_mm.valueChanged.connect(self._update_preview)
        self.in_signature_width_mm.valueChanged.connect(self._update_preview)
        self.in_signature_height_mm.valueChanged.connect(self._update_preview)
        self.in_docx_image_width_inches.valueChanged.connect(self._update_preview)
        self.in_docx_signature_width_inches.valueChanged.connect(self._update_preview)
        self.in_table_col_widths.textChanged.connect(self._update_preview)
        self.ck_auto_generate_akta_nr.stateChanged.connect(self._update_preview)
        self.in_default_currency.textChanged.connect(self._update_preview)
        self.in_default_unit.textChanged.connect(self._update_preview)
        self.in_default_pvn_rate.valueChanged.connect(self._update_preview)
        self.in_poppler_path.textChanged.connect(self._update_preview)


    def izvēlēties_poppler_ceļu(self):
        folder_path = QFileDialog.getExistingDirectory(self, "Izvēlēties Poppler bin direktoriju")
        if folder_path:
            self.in_poppler_path.setText(folder_path)

    # ----- Tab: Šabloni -----
    def _būvēt_sablonu_tab(self):
        w = QWidget()
        v = QVBoxLayout()

        self.sablonu_list = QListWidget()
        self.sablonu_list.setSelectionMode(QAbstractItemView.ExtendedSelection)  # Ļauj atlasīt vairākus elementus
        self.sablonu_list.itemDoubleClicked.connect(self.ieladet_sablonu)

        btn_ieladet_sablonu = QPushButton("Ielādēt izvēlēto šablonu")
        btn_ieladet_sablonu.clicked.connect(lambda: self.ieladet_sablonu(self.sablonu_list.currentItem()))

        btn_dzest_sablonus = QPushButton("Dzēst atlasītos šablonus")
        btn_dzest_sablonus.clicked.connect(self.dzest_sablonus)

        # JAUNAS POGAS PAROLES PĀRVALDĪBAI
        btn_mainit_paroli = QPushButton("Mainīt/Pievienot paroli")
        btn_mainit_paroli.clicked.connect(self.mainit_sablonu_paroli)
        btn_nonemt_paroli = QPushButton("Noņemt paroli")
        btn_nonemt_paroli.clicked.connect(self.nonemt_sablonu_paroli)

        btns_layout = QHBoxLayout()
        btns_layout.addWidget(btn_ieladet_sablonu)
        btns_layout.addWidget(btn_dzest_sablonus)
        btns_layout.addWidget(btn_mainit_paroli)  # Pievienojam jauno pogu
        btns_layout.addWidget(btn_nonemt_paroli)  # Pievienojam jauno pogu
        btns_layout.addStretch()

        v.addWidget(QLabel("Pieejamie šabloni:"))
        v.addWidget(self.sablonu_list)
        v.addLayout(btns_layout)
        v.addStretch()

        w.setLayout(v)
        self.tabs.addTab(w, "Šabloni")

    def ieladet_sablonu(self, item: QListWidgetItem):
        if not item:
            return

        sablonu_nosaukums = item.text()
        file_path = None

        # Pārbaudām, vai tas ir iebūvētais šablons
        if sablonu_nosaukums == "Pappus dati (piemērs)":
            # ... (esošais Pappus dati šablona kods) ...
            # Šeit nav jāmaina, jo iebūvētajam šablonam nav paroles
            d = AktaDati(
                akta_nr="",
                datums=datetime.now().strftime('%Y-%m-%d'),
                vieta="Rīga",
                pasūtījuma_nr="",
                pieņēmējs=Persona(
                    nosaukums="SIA \"Pappus\"",
                    reģ_nr="40003123456",
                    adrese="Lielā iela 1, Rīga, LV-1010",
                    kontaktpersona="Jānis Bērziņš",
                    tālrunis="+371 21234567",
                    epasts="janis.berzins@pappus.lv",
                    bankas_konts="LV12BANK1234567890123",
                    juridiskais_statuss="Juridiska persona"
                ),
                nodevējs=Persona(
                    nosaukums="SIA \"Demo Serviss\"",
                    reģ_nr="40003987654",
                    adrese="Mazā iela 5, Rīga, LV-1005",
                    kontaktpersona="Anna Liepa",
                    tālrunis="+371 27654321",
                    epasts="anna.liepa@demoserviss.lv",
                    bankas_konts="LV32BANK9876543210987",
                    juridiskais_statuss="Juridiska persona"
                ),
                pozīcijas=[
                    Pozīcija("Datora remonts", Decimal("1"), "gab.", Decimal("50.00"), "SN12345", "1 gads",
                             "Veikta diagnostika un komponentu nomaiņa."),
                    Pozīcija("Programmatūras instalācija", Decimal("1"), "st.", Decimal("25.00"), "", "",
                             "Uzstādīta operētājsistēma un biroja programmatūra."),
                    Pozīcija("Detaļas (RAM 8GB)", Decimal("1"), "gab.", Decimal("35.00"), "RAM9876", "2 gadi",
                             "Augstas veiktspējas RAM modulis.")
                ],
                piezīmes="Veikts pilns datora diagnostikas un remonta pakalpojums. Iekļauta programmatūras optimizācija.",
                iekļaut_pvn=True,
                pvn_likme=Decimal("21.0"),
                parakstu_rindas=True,
                logotipa_ceļš="",
                fonts_ceļš="",
                paraksts_pieņēmējs_ceļš="",
                paraksts_nodevējs_ceļš="",
                līguma_nr="",
                izpildes_termiņš="",
                pieņemšanas_datums="",
                nodošanas_datums="",
                strīdu_risināšana="Visi strīdi, kas izriet no šī akta, tiks risināti sarunu ceļā. Ja vienošanās netiek panākta, strīdi tiks nodoti izskatīšanai Latvijas Republikas tiesā saskaņā ar spēkā esošajiem normatīvajiem aktiem.",
                konfidencialitātes_klauzula=True,
                soda_nauda_procenti=Decimal("0.5"),
                piegādes_nosacījumi="DAP (Delivered at Place) Rīga, Latvija",
                apdrošināšana=True,
                papildu_nosacījumi="Abas puses apliecina, ka ir iepazinušās ar šī akta saturu un piekrīt visiem tā nosacījumiem. Akts sastādīts divos eksemplāros, katrai pusei pa vienu.",
                atsauces_dokumenti="Pavadzīme Nr. PV-2023/010, Garantijas talons Nr. GT-2023/005",
                akta_statuss="Melnraksts",
                valūta="EUR",
                elektroniskais_paraksts=False,
                radit_elektronisko_parakstu_tekstu=False,  # JAUNA RINDAS
                pdf_page_size="A4",
                pdf_page_orientation="Portrets",
                pdf_margin_left=Decimal("18"),
                pdf_margin_right=Decimal("18"),
                pdf_margin_top=Decimal("16"),
                pdf_margin_bottom=Decimal("16"),
                pdf_font_size_head=14,
                pdf_font_size_normal=10,
                pdf_font_size_small=9,
                pdf_font_size_table=9,
                pdf_logo_width_mm=Decimal("35"),
                pdf_signature_width_mm=Decimal("50"),
                pdf_signature_height_mm=Decimal("20"),
                docx_image_width_inches=Decimal("4"),
                docx_signature_width_inches=Decimal("1.5"),
                table_col_widths="10,40,18,18,20,20,25,25,25",
                auto_generate_akta_nr=False,
                default_currency="EUR",
                default_unit="gab.",
                default_pvn_rate=Decimal("21.0"),
                poppler_path="",
                header_text_color="#000000",
                footer_text_color="#000000",
                table_header_bg_color="#E0E0E0",
                table_grid_color="#CCCCCC",
                table_row_spacing=Decimal("4"),
                line_spacing_multiplier=Decimal("1.2"),
                show_page_numbers=True,
                show_generation_timestamp=True,
                currency_symbol_position="after",
                date_format="YYYY-MM-DD",
                signature_line_length_mm=Decimal("60"),
                signature_line_thickness_pt=Decimal("0.5"),
                add_cover_page=False,
                cover_page_title="Pieņemšanas-Nodošanas Akts",
                cover_page_logo_width_mm=Decimal("80"),
                # Individuālais QR kods
                include_custom_qr_code=False,
                custom_qr_code_data="",
                custom_qr_code_size_mm=Decimal("20"),
                custom_qr_code_position="bottom_right",

                # Automātiskais QR kods (akta ID)
                include_auto_qr_code=False,
                auto_qr_code_size_mm=Decimal("20"),
                auto_qr_code_position="bottom_left",

                add_watermark=False,
                watermark_text="MELNRAKSTS",
                watermark_font_size=72,
                watermark_color="#E0E0E0",
                watermark_rotation=45,
                enable_pdf_encryption=False,
                pdf_user_password="",
                pdf_owner_password="",
                allow_printing=True,
                allow_copying=True,
                allow_modifying=False,
                allow_annotating=True,
                default_country="Latvija",
                default_city="Rīga",
                show_contact_details_in_header=False,
                contact_details_header_font_size=8,
                item_image_width_mm=Decimal("50"),
                item_image_caption_font_size=8,
                show_item_notes_in_table=True,
                show_item_serial_number_in_table=True,
                show_item_warranty_in_table=True,
                table_cell_padding_mm=Decimal("2"),
                table_header_font_style="bold",
                table_content_alignment="left",
                signature_font_size=9,
                signature_spacing_mm=Decimal("10"),
                document_title_font_size=18,
                document_title_color="#000000",
                section_heading_font_size=12,
                section_heading_color="#000000",
                paragraph_line_spacing_multiplier=Decimal("1.2"),
                table_border_style="solid",
                table_border_thickness_pt=Decimal("0.5"),
                table_alternate_row_color="",
                show_total_sum_in_words=False,
                total_sum_in_words_language="lv",
                default_vat_calculation_method="exclusive",
                show_vat_breakdown=True,
                enable_digital_signature_field=False,
                digital_signature_field_name="Paraksts",
                digital_signature_field_size_mm=Decimal("40"),
                digital_signature_field_position="bottom_center",
                template_password=""  # Nodrošinām, ka šim nav paroles
            )
            self.ieviest_datus(d)
            QMessageBox.information(self, "Šablons ielādēts", f"Šablons '{sablonu_nosaukums}' veiksmīgi ielādēts.")
            return
        else:
            # Mēģinām ielādēt no šablonu direktorijas
            current_templates_dir = self.data.templates_dir
            if not current_templates_dir:  # Ja vēl nav iestatīts (piemēram, pirmajā startā)
                current_templates_dir = os.path.join(APP_DATA_DIR, "AktaGenerators_Templates")
                os.makedirs(current_templates_dir, exist_ok=True)  # Pārliecināmies, ka noklusējuma mape eksistē

            file_path = os.path.join(current_templates_dir, f"{sablonu_nosaukums}.json")


        if file_path and os.path.exists(file_path):
            try:
                with open(file_path, 'r', encoding='utf-8') as f:
                    data = json.load(f)

                # Pārbaudām paroli, ja tā ir iestatīta
                stored_password = data.get('template_password', '')
                if stored_password:
                    entered_password, ok_pass = QInputDialog.getText(self, "Ievadiet paroli",
                                                                     f"Šablonam '{sablonu_nosaukums}' ir parole. Lūdzu, ievadiet to:",
                                                                     QLineEdit.Password)
                    if not ok_pass or entered_password != stored_password:
                        QMessageBox.warning(self, "Nepareiza parole",
                                            "Ievadītā parole ir nepareiza vai ievade atcelta.")
                        return  # Atceļam ielādi, ja parole ir nepareiza vai atcelta

                # Helper to safely get and convert Decimal values
                def get_decimal(dict_obj, key, default_val):
                    val = dict_obj.get(key, default_val)
                    return to_decimal(val)

                # Helper to safely get boolean values
                def get_bool(dict_obj, key, default_val):
                    val = dict_obj.get(key, default_val)
                    return bool(val)

                # Izveidojam AktaDati objektu no ielādētajiem datiem
                d = AktaDati(
                    akta_nr=data.get('akta_nr', ''), datums=data.get('datums', datetime.now().strftime('%Y-%m-%d')),
                    vieta=data.get('vieta', ''),
                    pasūtījuma_nr=data.get('pasūtījuma_nr', ''),
                    pieņēmējs=Persona(
                        nosaukums=data.get('pieņēmējs', {}).get('nosaukums', ''),
                        reģ_nr=data.get('pieņēmējs', {}).get('reģ_nr', ''),
                        adrese=data.get('pieņēmējs', {}).get('adrese', ''),
                        kontaktpersona=data.get('pieņēmējs', {}).get('kontaktpersona', ''),
                        tālrunis=data.get('pieņēmējs', {}).get('tālrunis', ''),
                        epasts=data.get('pieņēmējs', {}).get('epasts', ''),
                        bankas_konts=data.get('pieņēmējs', {}).get('bankas_konts', ''),
                        juridiskais_statuss=data.get('pieņēmējs', {}).get('juridiskais_statuss', '')
                    ),
                    nodevējs=Persona(
                        nosaukums=data.get('nodevējs', {}).get('nosaukums', ''),
                        reģ_nr=data.get('nodevējs', {}).get('reģ_nr', ''),
                        adrese=data.get('nodevējs', {}).get('adrese', ''),
                        kontaktpersona=data.get('nodevējs', {}).get('kontaktpersona', ''),
                        tālrunis=data.get('nodevējs', {}).get('tālrunis', ''),
                        epasts=data.get('nodevējs', {}).get('epasts', ''),
                        bankas_konts=data.get('nodevējs', {}).get('bankas_konts', ''),
                        juridiskais_statuss=data.get('nodevējs', {}).get('juridiskais_statuss', '')
                    ),
                    pozīcijas=[Pozīcija(
                        apraksts=p.get('apraksts', ''),
                        daudzums=get_decimal(p, 'daudzums', '0'),
                        vienība=p.get('vienība', 'gab.'),
                        cena=get_decimal(p, 'cena', '0'),
                        seriālais_nr=p.get('seriālais_nr', ''),
                        garantija=p.get('garantija', ''),
                        piezīmes_pozīcijai=p.get('piezīmes_pozīcijai', '')
                    ) for p in data.get('pozīcijas', [])],
                    attēli=[Attēls(**a) for a in data.get('attēli', [])],
                    piezīmes=data.get('piezīmes', ''), iekļaut_pvn=get_bool(data, 'iekļaut_pvn', False),
                    pvn_likme=get_decimal(data, 'pvn_likme', '21'),
                    parakstu_rindas=get_bool(data, 'parakstu_rindas', True),
                    logotipa_ceļš=data.get('logotipa_ceļš', ''), fonts_ceļš=data.get('fonts_ceļš', ''),
                    paraksts_pieņēmējs_ceļš=data.get('paraksts_pieņēmējs_ceļš', ''),
                    paraksts_nodevējs_ceļš=data.get('paraksts_nodevējs_ceļš', ''),
                    līguma_nr=data.get('līguma_nr', ''),
                    izpildes_termiņš=data.get('izpildes_termiņš', ''),
                    pieņemšanas_datums=data.get('pieņemšanas_datums', ''),
                    nodošanas_datums=data.get('nodošanas_datums', ''),
                    strīdu_risināšana=data.get('strīdu_risināšana', ''),
                    konfidencialitātes_klauzula=get_bool(data, 'konfidencialitātes_klauzula', False),
                    soda_nauda_procenti=get_decimal(data, 'soda_nauda_procenti', '0.0'),
                    piegādes_nosacījumi=data.get('piegādes_nosacījumi', ''),
                    apdrošināšana=get_bool(data, 'apdrošināšana', False),
                    papildu_nosacījumi=data.get('papildu_nosacījumi', ''),
                    atsauces_dokumenti=data.get('atsauces_dokumenti', ''),
                    akta_statuss=data.get('akta_statuss', 'Melnraksts'),
                    valūta=data.get('valūta', 'EUR'),
                    elektroniskais_paraksts=get_bool(data, 'elektroniskais_paraksts', False),
                    radit_elektronisko_parakstu_tekstu=get_bool(data, 'radit_elektronisko_parakstu_tekstu', False),
                    # JAUNA RINDAS
                    pdf_page_size=data.get('pdf_page_size', 'A4'),
                    pdf_page_orientation=data.get('pdf_page_orientation', 'Portrets'),
                    pdf_margin_left=get_decimal(data, 'pdf_margin_left', '18'),
                    pdf_margin_right=get_decimal(data, 'pdf_margin_right', '18'),
                    pdf_margin_top=get_decimal(data, 'pdf_margin_top', '16'),
                    pdf_margin_bottom=get_decimal(data, 'pdf_margin_bottom', '16'),
                    pdf_font_size_head=data.get('pdf_font_size_head', 14),
                    pdf_font_size_normal=data.get('pdf_font_size_normal', 10),
                    pdf_font_size_small=data.get('pdf_font_size_small', 9),
                    pdf_font_size_table=data.get('pdf_font_size_table', 9),
                    pdf_logo_width_mm=get_decimal(data, 'pdf_logo_width_mm', '35'),
                    pdf_signature_width_mm=get_decimal(data, 'pdf_signature_width_mm', '50'),
                    pdf_signature_height_mm=get_decimal(data, 'pdf_signature_height_mm', '20'),
                    docx_image_width_inches=get_decimal(data, 'docx_image_width_inches', '4'),
                    docx_signature_width_inches=get_decimal(data, 'docx_signature_width_inches', '1.5'),
                    table_col_widths=data.get('table_col_widths', '10,40,18,18,20,20,25,25,25'),
                    auto_generate_akta_nr=get_bool(data, 'auto_generate_akta_nr', False),
                    default_currency=data.get('default_currency', 'EUR'),
                    default_unit=data.get('default_unit', 'gab.'),
                    default_pvn_rate=get_decimal(data, 'default_pvn_rate', '21.0'),
                    poppler_path=data.get('poppler_path', ''),
                    header_text_color=data.get('header_text_color', '#000000'),
                    footer_text_color=data.get('footer_text_color', '#000000'),
                    table_header_bg_color=data.get('table_header_bg_color', '#E0E0E0'),
                    table_grid_color=data.get('table_grid_color', '#CCCCCC'),
                    table_row_spacing=get_decimal(data, 'table_row_spacing', '4'),
                    line_spacing_multiplier=get_decimal(data, 'line_spacing_multiplier', '1.2'),
                    show_page_numbers=get_bool(data, 'show_page_numbers', True),
                    show_generation_timestamp=get_bool(data, 'show_generation_timestamp', True),
                    currency_symbol_position=data.get('currency_symbol_position', 'after'),
                    date_format=data.get('date_format', 'YYYY-MM-DD'),
                    signature_line_length_mm=get_decimal(data, 'signature_line_length_mm', '60'),
                    signature_line_thickness_pt=get_decimal(data, 'signature_line_thickness_pt', '0.5'),
                    add_cover_page=get_bool(data, 'add_cover_page', False),
                    cover_page_title=data.get('cover_page_title', 'Pieņemšanas-Nodošanas Akts'),
                    cover_page_logo_width_mm=get_decimal(data, 'cover_page_logo_width_mm', '80'),
                    # Individuālais QR kods
                    include_custom_qr_code=get_bool(data, 'include_custom_qr_code', False),
                    custom_qr_code_data=data.get('custom_qr_code_data', ''),
                    custom_qr_code_size_mm=get_decimal(data, 'custom_qr_code_size_mm', '20'),
                    custom_qr_code_position=data.get('custom_qr_code_position', 'bottom_right'),
                    custom_qr_code_pos_x_mm=get_decimal(data, 'custom_qr_code_pos_x_mm', '0'),
                    custom_qr_code_pos_y_mm=get_decimal(data, 'custom_qr_code_pos_y_mm', '0'),
                    custom_qr_code_color=data.get('custom_qr_code_color', '#000000'),

                    # Automātiskais QR kods (akta ID)
                    include_auto_qr_code=get_bool(data, 'include_auto_qr_code', False),
                    auto_qr_code_size_mm=get_decimal(data, 'auto_qr_code_size_mm', '20'),
                    auto_qr_code_position=data.get('auto_qr_code_position', 'bottom_left'),
                    auto_qr_code_pos_x_mm=get_decimal(data, 'auto_qr_code_pos_x_mm', '0'),
                    auto_qr_code_pos_y_mm=get_decimal(data, 'auto_qr_code_pos_y_mm', '0'),
                    auto_qr_code_color=data.get('auto_qr_code_color', '#000000'),

                    add_watermark=get_bool(data, 'add_watermark', False),
                    watermark_text=data.get('watermark_text', 'MELNRAKSTS'),
                    watermark_font_size=data.get('watermark_font_size', 72),
                    watermark_color=data.get('watermark_color', '#E0E0E0'),
                    watermark_rotation=data.get('watermark_rotation', 45),
                    enable_pdf_encryption=get_bool(data, 'enable_pdf_encryption', False),
                    pdf_user_password=data.get('pdf_user_password', ''),
                    pdf_owner_password=data.get('pdf_owner_password', ''),
                    allow_printing=get_bool(data, 'allow_printing', True),
                    allow_copying=get_bool(data, 'allow_copying', True),
                    allow_modifying=get_bool(data, 'allow_modifying', False),
                    allow_annotating=get_bool(data, 'allow_annotating', True),
                    default_country=data.get('default_country', 'Latvija'),
                    default_city=data.get('default_city', 'Rīga'),
                    show_contact_details_in_header=get_bool(data, 'show_contact_details_in_header', False),
                    contact_details_header_font_size=data.get('contact_details_header_font_size', 8),
                    item_image_width_mm=get_decimal(data, 'item_image_width_mm', '50'),
                    item_image_caption_font_size=data.get('item_image_caption_font_size', 8),
                    show_item_notes_in_table=get_bool(data, 'show_item_notes_in_table', True),
                    show_item_serial_number_in_table=get_bool(data, 'show_item_serial_number_in_table', True),
                    show_item_warranty_in_table=get_bool(data, 'show_item_warranty_in_table', True),
                    table_cell_padding_mm=get_decimal(data, 'table_cell_padding_mm', '2'),
                    table_header_font_style=data.get('table_header_font_style', 'bold'),
                    table_content_alignment=data.get('table_content_alignment', 'left'),
                    signature_font_size=data.get('signature_font_size', 9),
                    signature_spacing_mm=get_decimal(data, 'signature_spacing_mm', '10'),
                    document_title_font_size=data.get('document_title_font_size', 18),
                    document_title_color=data.get('document_title_color', '#000000'),
                    section_heading_font_size=data.get('section_heading_font_size', 12),
                    section_heading_color=data.get('section_heading_color', '#000000'),
                    paragraph_line_spacing_multiplier=get_decimal(data, 'paragraph_line_spacing_multiplier', '1.2'),
                    table_border_style=data.get('table_border_style', 'solid'),
                    table_border_thickness_pt=get_decimal(data, 'table_border_thickness_pt', '0.5'),
                    table_alternate_row_color=data.get('table_alternate_row_color', ''),
                    show_total_sum_in_words=get_bool(data, 'show_total_sum_in_words', False),
                    total_sum_in_words_language=data.get('total_sum_in_words_language', 'lv'),
                    default_vat_calculation_method=data.get('default_vat_calculation_method', 'exclusive'),
                    show_vat_breakdown=get_bool(data, 'show_vat_breakdown', True),
                    enable_digital_signature_field=get_bool(data, 'enable_digital_signature_field', False),
                    digital_signature_field_name=data.get('digital_signature_field_name', 'Paraksts'),
                    digital_signature_field_size_mm=get_decimal(data, 'digital_signature_field_size_mm', '40'),
                    digital_signature_field_position=data.get('digital_signature_field_position', 'bottom_center'),
                    template_password=data.get('template_password', '')  # Ielādējam paroli
                )
                self.ieviest_datus(d)
                QMessageBox.information(self, "Šablons ielādēts", f"Šablons '{sablonu_nosaukums}' veiksmīgi ielādēts.")
            except Exception as e:
                QMessageBox.critical(self, "Kļūda", f"Neizdevās ielādēt šablonu '{sablonu_nosaukums}':\n{e}")
        else:
            QMessageBox.warning(self, "Kļūda", f"Šablona fails '{sablonu_nosaukums}.json' nav atrasts.")

        if sablonu_nosaukums == "Pappus dati (piemērs)":
            # Šis ir iebūvētais piemērs, to var ielādēt tieši
            d = AktaDati(
                akta_nr="",  # Šabloniem akta nr. un datums parasti ir tukši/pašreizējie
                datums=datetime.now().strftime('%Y-%m-%d'),
                vieta="Rīga",
                pasūtījuma_nr="",
                pieņēmējs=Persona(
                    nosaukums="SIA \"Pappus\"",
                    reģ_nr="40003123456",
                    adrese="Lielā iela 1, Rīga, LV-1010",
                    kontaktpersona="Jānis Bērziņš",
                    tālrunis="+371 21234567",
                    epasts="janis.berzins@pappus.lv",
                    bankas_konts="LV12BANK1234567890123",
                    juridiskais_statuss="Juridiska persona"
                ),
                nodevējs=Persona(
                    nosaukums="SIA \"Demo Serviss\"",
                    reģ_nr="40003987654",
                    adrese="Mazā iela 5, Rīga, LV-1005",
                    kontaktpersona="Anna Liepa",
                    tālrunis="+371 27654321",
                    epasts="anna.liepa@demoserviss.lv",
                    bankas_konts="LV32BANK9876543210987",
                    juridiskais_statuss="Juridiska persona"
                ),
                pozīcijas=[
                    Pozīcija("Datora remonts", Decimal("1"), "gab.", Decimal("50.00"), "SN12345", "1 gads",
                             "Veikta diagnostika un komponentu nomaiņa."),
                    Pozīcija("Programmatūras instalācija", Decimal("1"), "st.", Decimal("25.00"), "", "",
                             "Uzstādīta operētājsistēma un biroja programmatūra."),
                    Pozīcija("Detaļas (RAM 8GB)", Decimal("1"), "gab.", Decimal("35.00"), "RAM9876", "2 gadi",
                             "Augstas veiktspējas RAM modulis.")
                ],
                piezīmes="Veikts pilns datora diagnostikas un remonta pakalpojums. Iekļauta programmatūras optimizācija.",
                iekļaut_pvn=True,
                pvn_likme=Decimal("21.0"),
                parakstu_rindas=True,
                logotipa_ceļš="",
                fonts_ceļš="",
                paraksts_pieņēmējs_ceļš="",
                paraksts_nodevējs_ceļš="",
                līguma_nr="",  # Šabloniem šie lauki parasti ir tukši
                izpildes_termiņš="",
                pieņemšanas_datums="",
                nodošanas_datums="",
                strīdu_risināšana="Visi strīdi, kas izriet no šī akta, tiks risināti sarunu ceļā. Ja vienošanās netiek panākta, strīdi tiks nodoti izskatīšanai Latvijas Republikas tiesā saskaņā ar spēkā esošajiem normatīvajiem aktiem.",
                konfidencialitātes_klauzula=True,
                soda_nauda_procenti=Decimal("0.5"),
                piegādes_nosacījumi="DAP (Delivered at Place) Rīga, Latvija",
                apdrošināšana=True,
                papildu_nosacījumi="Abas puses apliecina, ka ir iepazinušās ar šī akta saturu un piekrīt visiem tā nosacījumiem. Akts sastādīts divos eksemplāros, katrai pusei pa vienu.",
                atsauces_dokumenti="Pavadzīme Nr. PV-2023/010, Garantijas talons Nr. GT-2023/005",
                akta_statuss="Melnraksts",  # Šabloniem statuss ir melnraksts
                valūta="EUR",
                elektroniskais_paraksts=False,
                pdf_page_size="A4",
                pdf_page_orientation="Portrets",
                pdf_margin_left=Decimal("18"),
                pdf_margin_right=Decimal("18"),
                pdf_margin_top=Decimal("16"),
                pdf_margin_bottom=Decimal("16"),
                pdf_font_size_head=14,
                pdf_font_size_normal=10,
                pdf_font_size_small=9,
                pdf_font_size_table=9,
                pdf_logo_width_mm=Decimal("35"),
                pdf_signature_width_mm=Decimal("50"),
                pdf_signature_height_mm=Decimal("20"),
                docx_image_width_inches=Decimal("4"),
                docx_signature_width_inches=Decimal("1.5"),
                table_col_widths="10,40,18,18,20,20,25,25,25",
                auto_generate_akta_nr=False,
                default_currency="EUR",
                default_unit="gab.",
                default_pvn_rate=Decimal("21.0"),
                poppler_path="",
                # Default values for new settings
                header_text_color="#000000",
                footer_text_color="#000000",
                table_header_bg_color="#E0E0E0",
                table_grid_color="#CCCCCC",
                table_row_spacing=Decimal("4"),
                line_spacing_multiplier=Decimal("1.2"),
                show_page_numbers=True,
                show_generation_timestamp=True,
                currency_symbol_position="after",
                date_format="YYYY-MM-DD",
                signature_line_length_mm=Decimal("60"),
                signature_line_thickness_pt=Decimal("0.5"),
                add_cover_page=False,
                cover_page_title="Pieņemšanas-Nodošanas Akts",
                cover_page_logo_width_mm=Decimal("80"),
                # Individuālais QR kods
                include_custom_qr_code=False,
                custom_qr_code_data="",
                custom_qr_code_size_mm=Decimal("20"),
                custom_qr_code_position="bottom_right",

                # Automātiskais QR kods (akta ID)
                include_auto_qr_code=False,
                auto_qr_code_size_mm=Decimal("20"),
                auto_qr_code_position="bottom_left",

                add_watermark=False,
                watermark_text="MELNRAKSTS",
                watermark_font_size=72,
                watermark_color="#E0E0E0",
                watermark_rotation=45,
                enable_pdf_encryption=False,
                pdf_user_password="",
                pdf_owner_password="",
                allow_printing=True,
                allow_copying=True,
                allow_modifying=False,
                allow_annotating=True,
                default_country="Latvija",
                default_city="Rīga",
                show_contact_details_in_header=False,
                contact_details_header_font_size=8,
                item_image_width_mm=Decimal("50"),
                item_image_caption_font_size=8,
                show_item_notes_in_table=True,
                show_item_serial_number_in_table=True,
                show_item_warranty_in_table=True,
                table_cell_padding_mm=Decimal("2"),
                table_header_font_style="bold",
                table_content_alignment="left",
                signature_font_size=9,
                signature_spacing_mm=Decimal("10"),
                document_title_font_size=18,
                document_title_color="#000000",
                section_heading_font_size=12,
                section_heading_color="#000000",
                paragraph_line_spacing_multiplier=Decimal("1.2"),
                table_border_style="solid",
                table_border_thickness_pt=Decimal("0.5"),
                table_alternate_row_color="",
                show_total_sum_in_words=False,
                total_sum_in_words_language="lv",
                default_vat_calculation_method="exclusive",
                show_vat_breakdown=True,
                enable_digital_signature_field=False,
                digital_signature_field_name="Paraksts",
                digital_signature_field_size_mm=Decimal("40"),
                digital_signature_field_position="bottom_center"
            )
            self.ieviest_datus(d)
            QMessageBox.information(self, "Šablons ielādēts", f"Šablons '{sablonu_nosaukums}' veiksmīgi ielādēts.")
            return  # Iziet no funkcijas pēc iebūvētā šablona ielādes
        else:
            # Mēģinām ielādēt no šablonu direktorijas
            file_path = os.path.join(self.data.templates_dir, f"{sablonu_nosaukums}.json")

        if file_path and os.path.exists(file_path):
            try:
                with open(file_path, 'r', encoding='utf-8') as f:
                    data = json.load(f)

                    # Pārbaudām paroli, ja tā ir iestatīta
                    stored_password = data.get('template_password', '')
                    if stored_password:
                        entered_password, ok_pass = QInputDialog.getText(self, "Ievadiet paroli",
                                                                         f"Šablonam '{sablonu_nosaukums}' ir parole. Lūdzu, ievadiet to:",
                                                                         QLineEdit.Password)
                        if not ok_pass or entered_password != stored_password:
                            QMessageBox.warning(self, "Nepareiza parole",
                                                "Ievadītā parole ir nepareiza vai ievade atcelta.")
                            return  # Atceļam ielādi, ja parole ir nepareiza vai atcelta

                    # Helper to safely get and convert Decimal values
                    def get_decimal(dict_obj, key, default_val):
                        val = dict_obj.get(key, default_val)
                        return to_decimal(val)

                    # Helper to safely get boolean values
                    def get_bool(dict_obj, key, default_val):
                        val = dict_obj.get(key, default_val)
                        return bool(val)

                    # Izveidojam AktaDati objektu no ielādētajiem datiem
                    d = AktaDati(
                        akta_nr=data.get('akta_nr', ''), datums=data.get('datums', datetime.now().strftime('%Y-%m-%d')),
                        vieta=data.get('vieta', ''),
                        pasūtījuma_nr=data.get('pasūtījuma_nr', ''),
                        pieņēmējs=Persona(
                            nosaukums=data.get('pieņēmējs', {}).get('nosaukums', ''),
                            reģ_nr=data.get('pieņēmējs', {}).get('reģ_nr', ''),
                            adrese=data.get('pieņēmējs', {}).get('adrese', ''),
                            kontaktpersona=data.get('pieņēmējs', {}).get('kontaktpersona', ''),
                            tālrunis=data.get('pieņēmējs', {}).get('tālrunis', ''),
                            epasts=data.get('pieņēmējs', {}).get('epasts', ''),
                            bankas_konts=data.get('pieņēmējs', {}).get('bankas_konts', ''),
                            juridiskais_statuss=data.get('pieņēmējs', {}).get('juridiskais_statuss', '')
                        ),
                        nodevējs=Persona(
                            nosaukums=data.get('nodevējs', {}).get('nosaukums', ''),
                            reģ_nr=data.get('nodevējs', {}).get('reģ_nr', ''),
                            adrese=data.get('nodevējs', {}).get('adrese', ''),
                            kontaktpersona=data.get('nodevējs', {}).get('kontaktpersona', ''),
                            tālrunis=data.get('nodevējs', {}).get('tālrunis', ''),
                            epasts=data.get('nodevējs', {}).get('epasts', ''),
                            bankas_konts=data.get('nodevējs', {}).get('bankas_konts', ''),
                            juridiskais_statuss=data.get('nodevējs', {}).get('juridiskais_statuss', '')
                        ),
                        pozīcijas=[Pozīcija(
                            apraksts=p.get('apraksts', ''),
                            daudzums=get_decimal(p, 'daudzums', '0'),
                            vienība=p.get('vienība', 'gab.'),
                            cena=get_decimal(p, 'cena', '0'),
                            seriālais_nr=p.get('seriālais_nr', ''),
                            garantija=p.get('garantija', ''),
                            piezīmes_pozīcijai=p.get('piezīmes_pozīcijai', '')
                        ) for p in data.get('pozīcijas', [])],
                        attēli=[Attēls(**a) for a in data.get('attēli', [])],
                        piezīmes=data.get('piezīmes', ''), iekļaut_pvn=get_bool(data, 'iekļaut_pvn', False),
                        pvn_likme=get_decimal(data, 'pvn_likme', '21'),
                        parakstu_rindas=get_bool(data, 'parakstu_rindas', True),
                        logotipa_ceļš=data.get('logotipa_ceļš', ''), fonts_ceļš=data.get('fonts_ceļš', ''),
                        paraksts_pieņēmējs_ceļš=data.get('paraksts_pieņēmējs_ceļš', ''),
                        paraksts_nodevējs_ceļš=data.get('paraksts_nodevējs_ceļš', ''),
                        līguma_nr=data.get('līguma_nr', ''),
                        izpildes_termiņš=data.get('izpildes_termiņš', ''),
                        pieņemšanas_datums=data.get('pieņemšanas_datums', ''),
                        nodošanas_datums=data.get('nodošanas_datums', ''),
                        strīdu_risināšana=data.get('strīdu_risināšana', ''),
                        konfidencialitātes_klauzula=get_bool(data, 'konfidencialitātes_klauzula', False),
                        soda_nauda_procenti=get_decimal(data, 'soda_nauda_procenti', '0.0'),
                        piegādes_nosacījumi=data.get('piegādes_nosacījumi', ''),
                        apdrošināšana=get_bool(data, 'apdrošināšana', False),
                        papildu_nosacījumi=data.get('papildu_nosacījumi', ''),
                        atsauces_dokumenti=data.get('atsauces_dokumenti', ''),
                        akta_statuss=data.get('akta_statuss', 'Melnraksts'),
                        valūta=data.get('valūta', 'EUR'),
                        elektroniskais_paraksts=get_bool(data, 'elektroniskais_paraksts', False),
                        pdf_page_size=data.get('pdf_page_size', 'A4'),
                        pdf_page_orientation=data.get('pdf_page_orientation', 'Portrets'),
                        pdf_margin_left=get_decimal(data, 'pdf_margin_left', '18'),
                        pdf_margin_right=get_decimal(data, 'pdf_margin_right', '18'),
                        pdf_margin_top=get_decimal(data, 'pdf_margin_top', '16'),
                        pdf_margin_bottom=get_decimal(data, 'pdf_margin_bottom', '16'),
                        pdf_font_size_head=data.get('pdf_font_size_head', 14),
                        pdf_font_size_normal=data.get('pdf_font_size_normal', 10),
                        pdf_font_size_small=data.get('pdf_font_size_small', 9),
                        pdf_font_size_table=data.get('pdf_font_size_table', 9),
                        pdf_logo_width_mm=get_decimal(data, 'pdf_logo_width_mm', '35'),
                        pdf_signature_width_mm=get_decimal(data, 'pdf_signature_width_mm', '50'),
                        pdf_signature_height_mm=get_decimal(data, 'pdf_signature_height_mm', '20'),
                        docx_image_width_inches=get_decimal(data, 'docx_image_width_inches', '4'),
                        docx_signature_width_inches=get_decimal(data, 'docx_signature_width_inches', '1.5'),
                        table_col_widths=data.get('table_col_widths', '10,40,18,18,20,20,25,25,25'),
                        auto_generate_akta_nr=get_bool(data, 'auto_generate_akta_nr', False),
                        default_currency=data.get('default_currency', 'EUR'),
                        default_unit=data.get('default_unit', 'gab.'),
                        default_pvn_rate=get_decimal(data, 'default_pvn_rate', '21.0'),
                        poppler_path=data.get('poppler_path', ''),
                        # Load new settings
                        header_text_color=data.get('header_text_color', '#000000'),
                        footer_text_color=data.get('footer_text_color', '#000000'),
                        table_header_bg_color=data.get('table_header_bg_color', '#E0E0E0'),
                        table_grid_color=data.get('table_grid_color', '#CCCCCC'),
                        table_row_spacing=get_decimal(data, 'table_row_spacing', '4'),
                        line_spacing_multiplier=get_decimal(data, 'line_spacing_multiplier', '1.2'),
                        show_page_numbers=get_bool(data, 'show_page_numbers', True),
                        show_generation_timestamp=get_bool(data, 'show_generation_timestamp', True),
                        currency_symbol_position=data.get('currency_symbol_position', 'after'),
                        date_format=data.get('date_format', 'YYYY-MM-DD'),
                        signature_line_length_mm=get_decimal(data, 'signature_line_length_mm', '60'),
                        signature_line_thickness_pt=get_decimal(data, 'signature_line_thickness_pt', '0.5'),
                        add_cover_page=get_bool(data, 'add_cover_page', False),
                        cover_page_title=data.get('cover_page_title', 'Pieņemšanas-Nodošanas Akts'),
                        cover_page_logo_width_mm=get_decimal(data, 'cover_page_logo_width_mm', '80'),
                        # Individuālais QR kods
                        include_custom_qr_code=get_bool(data, 'include_custom_qr_code', False),
                        custom_qr_code_data=data.get('custom_qr_code_data', ''),
                        custom_qr_code_size_mm=get_decimal(data, 'custom_qr_code_size_mm', '20'),
                        custom_qr_code_position=data.get('custom_qr_code_position', 'bottom_right'),
                        custom_qr_code_pos_x_mm=get_decimal(data, 'custom_qr_code_pos_x_mm', '0'),
                        custom_qr_code_pos_y_mm=get_decimal(data, 'custom_qr_code_pos_y_mm', '0'),
                        custom_qr_code_color=data.get('custom_qr_code_color', '#000000'),

                        # Automātiskais QR kods (akta ID)
                        include_auto_qr_code=get_bool(data, 'include_auto_qr_code', False),
                        auto_qr_code_size_mm=get_decimal(data, 'auto_qr_code_size_mm', '20'),
                        auto_qr_code_position=data.get('auto_qr_code_position', 'bottom_left'),
                        auto_qr_code_pos_x_mm=get_decimal(data, 'auto_qr_code_pos_x_mm', '0'),
                        auto_qr_code_pos_y_mm=get_decimal(data, 'auto_qr_code_pos_y_mm', '0'),
                        auto_qr_code_color=data.get('auto_qr_code_color', '#000000'),

                        add_watermark=get_bool(data, 'add_watermark', False),
                        watermark_text=data.get('watermark_text', 'MELNRAKSTS'),
                        watermark_font_size=data.get('watermark_font_size', 72),
                        watermark_color=data.get('watermark_color', '#E0E0E0'),
                        watermark_rotation=data.get('watermark_rotation', 45),
                        enable_pdf_encryption=get_bool(data, 'enable_pdf_encryption', False),
                        pdf_user_password=data.get('pdf_user_password', ''),
                        pdf_owner_password=data.get('pdf_owner_password', ''),
                        allow_printing=get_bool(data, 'allow_printing', True),
                        allow_copying=get_bool(data, 'allow_copying', True),
                        allow_modifying=get_bool(data, 'allow_modifying', False),
                        allow_annotating=get_bool(data, 'allow_annotating', True),
                        default_country=data.get('default_country', 'Latvija'),
                        default_city=data.get('default_city', 'Rīga'),
                        show_contact_details_in_header=get_bool(data, 'show_contact_details_in_header', False),
                        contact_details_header_font_size=data.get('contact_details_header_font_size', 8),
                        item_image_width_mm=get_decimal(data, 'item_image_width_mm', '50'),
                        item_image_caption_font_size=data.get('item_image_caption_font_size', 8),
                        show_item_notes_in_table=get_bool(data, 'show_item_notes_in_table', True),
                        show_item_serial_number_in_table=get_bool(data, 'show_item_serial_number_in_table', True),
                        show_item_warranty_in_table=get_bool(data, 'show_item_warranty_in_table', True),
                        table_cell_padding_mm=get_decimal(data, 'table_cell_padding_mm', '2'),
                        table_header_font_style=data.get('table_header_font_style', 'bold'),
                        table_content_alignment=data.get('table_content_alignment', 'left'),
                        signature_font_size=data.get('signature_font_size', 9),
                        signature_spacing_mm=get_decimal(data, 'signature_spacing_mm', '10'),
                        document_title_font_size=data.get('document_title_font_size', 18),
                        document_title_color=data.get('document_title_color', '#000000'),
                        section_heading_font_size=data.get('section_heading_font_size', 12),
                        section_heading_color=data.get('section_heading_color', '#000000'),
                        paragraph_line_spacing_multiplier=get_decimal(data, 'paragraph_line_spacing_multiplier', '1.2'),
                        table_border_style=data.get('table_border_style', 'solid'),
                        table_border_thickness_pt=get_decimal(data, 'table_border_thickness_pt', '0.5'),
                        table_alternate_row_color=data.get('table_alternate_row_color', ''),
                        show_total_sum_in_words=get_bool(data, 'show_total_sum_in_words', False),
                        total_sum_in_words_language=data.get('total_sum_in_words_language', 'lv'),
                        default_vat_calculation_method=data.get('default_vat_calculation_method', 'exclusive'),
                        show_vat_breakdown=get_bool(data, 'show_vat_breakdown', True),
                        enable_digital_signature_field=get_bool(data, 'enable_digital_signature_field', False),
                        digital_signature_field_name=data.get('digital_signature_field_name', 'Paraksts'),
                        digital_signature_field_size_mm=get_decimal(data, 'digital_signature_field_size_mm', '40'),
                        digital_signature_field_position=data.get('digital_signature_field_position', 'bottom_center'),
                        template_password=data.get('template_password', '')  # Ielādējam paroli
                    )

                    self.ieviest_datus(d)
                    QMessageBox.information(self, "Šablons ielādēts",
                                            f"Šablons '{sablonu_nosaukums}' veiksmīgi ielādēts.")
            except Exception as e:
                QMessageBox.critical(self, "Kļūda", f"Neizdevās ielādēt šablonu '{sablonu_nosaukums}':\n{e}")
        else:
            QMessageBox.warning(self, "Kļūda", f"Šablona fails '{sablonu_nosaukums}.json' nav atrasts.")

    # ----- Tab: Adrešu grāmata -----
    def _būvēt_adresu_gramata_tab(self):
        w = QWidget()
        v = QVBoxLayout()

        self.address_book_list = QListWidget()
        self.address_book_list.itemDoubleClicked.connect(self._load_selected_address_book_entry)

        btn_load_selected = QPushButton("Ielādēt izvēlēto")
        btn_load_selected.clicked.connect(lambda: self._load_selected_address_book_entry(self.address_book_list.currentItem()))
        btn_delete_selected = QPushButton("Dzēst izvēlēto")
        btn_delete_selected.clicked.connect(self._delete_selected_address_book_entry)

        btns_layout = QHBoxLayout()
        btns_layout.addWidget(btn_load_selected)
        btns_layout.addWidget(btn_delete_selected)
        btns_layout.addStretch()

        v.addWidget(QLabel("Saglabātās personas:"))
        v.addWidget(self.address_book_list)
        v.addLayout(btns_layout)
        v.addStretch()

        w.setLayout(v)
        self.tabs.addTab(w, "Adrešu grāmata")
        self._update_address_book_list()

    def _update_address_book_list(self):
        self.address_book_list.clear()
        for name in sorted(self.address_book.keys()):
            self.address_book_list.addItem(name)

    def _load_selected_address_book_entry(self, item: QListWidgetItem):
        if not item:
            return
        name = item.text()
        persona_data = self.address_book.get(name)
        if persona_data:
            reply = QMessageBox.question(self, "Ielādēt personu",
                                         f"Ielādēt '{name}' kā Pieņēmēju vai Nodevēju?",
                                         QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No | QMessageBox.StandardButton.Cancel,
                                         QMessageBox.StandardButton.Yes)
            if reply == QMessageBox.StandardButton.Yes:
                self.pie_in[0].setText(persona_data.get("nosaukums", ""))
                self.pie_in[1].setText(persona_data.get("reģ_nr", ""))
                self.pie_in[2].setText(persona_data.get("adrese", ""))
                self.pie_in[3].setText(persona_data.get("kontaktpersona", ""))
                self.pie_in[4].setText(persona_data.get("tālrunis", ""))
                self.pie_in[5].setText(persona_data.get("epasts", ""))
                self.pie_in[6].setText(persona_data.get("bankas_konts", ""))
                juridiskais_statuss_val = persona_data.get("juridiskais_statuss", "")
                idx = self.pie_in[7].findText(juridiskais_statuss_val)
                if idx >= 0:
                    self.pie_in[7].setCurrentIndex(idx)
                QMessageBox.information(self, "Ielādēts", f"Persona '{name}' ielādēta kā Pieņēmējs.")
            elif reply == QMessageBox.StandardButton.No:
                self.nod_in[0].setText(persona_data.get("nosaukums", ""))
                self.nod_in[1].setText(persona_data.get("reģ_nr", ""))
                self.nod_in[2].setText(persona_data.get("adrese", ""))
                self.nod_in[3].setText(persona_data.get("kontaktpersona", ""))
                self.nod_in[4].setText(persona_data.get("tālrunis", ""))
                self.nod_in[5].setText(persona_data.get("epasts", ""))
                self.nod_in[6].setText(persona_data.get("bankas_konts", ""))
                juridiskais_statuss_val = persona_data.get("juridiskais_statuss", "")
                idx = self.nod_in[7].findText(juridiskais_statuss_val)
                if idx >= 0:
                    self.nod_in[7].setCurrentIndex(idx)
                QMessageBox.information(self, "Ielādēts", f"Persona '{name}' ielādēta kā Nodevējs.")
        else:
            QMessageBox.warning(self, "Kļūda", "Izvēlētā persona nav atrasta adrešu grāmatā.")

    def _delete_selected_address_book_entry(self):
        item = self.address_book_list.currentItem()
        if not item:
            QMessageBox.warning(self, "Dzēst personu", "Lūdzu, izvēlieties personu, ko dzēst.")
            return
        name = item.text()
        reply = QMessageBox.question(self, "Dzēst personu",
                                     f"Vai tiešām vēlaties dzēst '{name}' no adrešu grāmatas?",
                                     QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)
        if reply == QMessageBox.StandardButton.Yes:
            if name in self.address_book:
                del self.address_book[name]
                self._save_address_book() # Saglabājam izmaiņas failā
                self._update_address_book_list() # Atjaunojam sarakstu GUI
                QMessageBox.information(self, "Dzēsts", f"Persona '{name}' veiksmīgi dzēsta.")
            else:
                QMessageBox.warning(self, "Kļūda", "Izvēlētā persona nav atrasta adrešu grāmatā.")

    # ----- Tab: Dokumentu vēsture -----
    def _būvēt_dokumentu_vesture_tab(self):
        w = QWidget()
        v = QVBoxLayout()

        self.history_list = QListWidget()
        self.history_list.itemDoubleClicked.connect(self._load_history_entry)

        btn_load_history = QPushButton("Ielādēt izvēlēto projektu")
        btn_load_history.clicked.connect(lambda: self._load_history_entry(self.history_list.currentItem()))
        btn_clear_history = QPushButton("Notīrīt vēsturi")
        btn_clear_history.clicked.connect(self._clear_history)

        btn_open_folder = QPushButton("Atvērt mapi")
        btn_open_folder.clicked.connect(self._open_document_folder)

        btns_layout = QHBoxLayout()
        btns_layout.addWidget(btn_load_history)
        btns_layout.addWidget(btn_clear_history)
        btns_layout.addWidget(btn_open_folder)  # PIEVIENOJAM JAUNO POGU
        btns_layout.addStretch()

        v.addWidget(QLabel("Pēdējie projekti:"))
        v.addWidget(self.history_list)
        v.addLayout(btns_layout)
        v.addStretch()

        w.setLayout(v)
        self.tabs.addTab(w, "Dokumentu vēsture")
        self._update_history_list()

    def _load_history(self):
        if os.path.exists(HISTORY_FILE):
            try:
                # Mēģinām ielādēt ar dažādām kodēšanām
                encodings = ['utf-8', 'utf-8-sig', 'cp1257', 'iso-8859-1', 'windows-1252']
                for encoding in encodings:
                    try:
                        with open(HISTORY_FILE, 'r', encoding=encoding) as f:
                            self.history = json.load(f)
                        break
                    except (UnicodeDecodeError, json.JSONDecodeError):
                        continue
                else:
                    # Ja neviena kodēšana nedarbojas, izveidojam jaunu vēsturi
                    print(f"Neizdevās ielādēt vēstures failu ar nevenu kodēšanu. Izveidojam jaunu.")
                    self.history = []
            except Exception as e:
                QMessageBox.warning(self, "Kļūda", f"Neizdevās ielādēt vēsturi: {e}")
                self.history = []
        else:
            self.history = []

    def _save_history(self):
        os.makedirs(SETTINGS_DIR, exist_ok=True) # Izveidojam direktoriju, ja tā neeksistē
        try:
            with open(HISTORY_FILE, 'w', encoding='utf-8') as f:
                json.dump(self.history, f, ensure_ascii=False, indent=2)
        except Exception as e:
            QMessageBox.critical(self, "Kļūda", f"Neizdevās saglabāt vēsturi: {e}")

    def _add_to_history(self, file_path: str):
        if not os.path.exists(file_path):
            return

        # Noņemam, ja jau ir sarakstā, lai pārvietotu uz saraksta sākumu
        self.history = [f for f in self.history if f != file_path]

        # Pievienojam saraksta sākumā
        self.history.insert(0, file_path)
        # Saglabājam tikai pēdējos 10 ierakstus
        if len(self.history) > 10:
            self.history = self.history[:10]
        self._save_history() # Saglabājam izmaiņas failā
        self._update_history_list() # Atjaunojam sarakstu GUI

    def _update_history_list(self):
        self.history_list.clear()
        # Filtrējam vēsturi, lai parādītu tikai tos failus, kas joprojām eksistē
        valid_history = [f for f in self.history if os.path.exists(f)]
        self.history = valid_history # Atjaunojam iekšējo vēstures sarakstu
        self._save_history() # Saglabājam atjaunināto vēsturi failā
        for entry in self.history:
            self.history_list.addItem(os.path.basename(entry))

    def _load_history_entry(self, item: QListWidgetItem):
        if not item:
            return
        file_name = item.text()
        # Atrodam pilnu ceļu, jo sarakstā ir tikai faila nosaukums
        file_path = next((f for f in self.history if os.path.basename(f) == file_name), None)
        if file_path and os.path.exists(file_path):
            self.ieladet_projektu(file_path)
        else:
            QMessageBox.warning(self, "Kļūda", f"Projekta fails '{file_name}' nav atrasts vai ir pārvietots.")
            # Noņemam neeksistējošo ierakstu no vēstures
            if file_path in self.history:
                self.history.remove(file_path)
                self._save_history()
                self._update_history_list()

    def _clear_history(self):
        reply = QMessageBox.question(self, "Notīrīt vēsturi",
                                     "Vai tiešām vēlaties notīrīt visu dokumentu vēsturi?",
                                     QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)
        if reply == QMessageBox.StandardButton.Yes:
            self.history = []
            self._save_history() # Saglabājam tukšu vēsturi failā
            self._update_history_list() # Atjaunojam sarakstu GUI
            QMessageBox.information(self, "Vēsture notīrīta", "Dokumentu vēsture ir notīrīta.")

    # ----- Tab: Karte -----

    def _open_document_folder(self):
        """Atver izvēlētā dokumenta mapi failu pārlūkā"""
        item = self.history_list.currentItem()
        if not item:
            QMessageBox.warning(self, "Atvērt mapi", "Lūdzu, izvēlieties dokumentu.")
            return

        file_name = item.text()
        file_path = next((f for f in self.history if os.path.basename(f) == file_name), None)
        if file_path and os.path.exists(file_path):
            folder_path = os.path.dirname(file_path)
            if sys.platform == "win32":
                os.startfile(folder_path)
            elif sys.platform == "darwin":
                os.system(f"open '{folder_path}'")
            else:
                os.system(f"xdg-open '{folder_path}'")
        else:
            QMessageBox.warning(self, "Kļūda", "Dokumenta mape nav atrasta.")

    def _būvēt_kartes_tab(self):
        w = QWidget()
        v = QVBoxLayout()

        self.web_view = QWebEngineView()

        # Izveidojam bridge objektu komunikācijai ar JavaScript
        self.map_bridge = MapBridge()
        self.map_bridge.map_click_callback = self._handle_map_click

        # Iestatām QWebChannel
        self.channel = QWebChannel()
        self.channel.registerObject("mapBridge", self.map_bridge)
        self.web_view.page().setWebChannel(self.channel)

        # HTML content for the map, including Leaflet from CDN
        map_html_content = """
    <!DOCTYPE html>
    <html>
    <head>
        <title>Interaktīva karte</title>
        <meta charset="utf-8" />
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        <link rel="stylesheet" href="https://unpkg.com/leaflet@1.7.1/dist/leaflet.css" />
        <script src="https://unpkg.com/leaflet@1.7.1/dist/leaflet.js"></script>
        <script src="qrc:///qtwebchannel/qwebchannel.js"></script>
        <style>
            body { margin: 0; padding: 0; }
            #mapid {
                height: 100vh;
                width: 100%;
                cursor: crosshair;
            }
            .custom-popup {
                font-family: Arial, sans-serif;
                font-size: 14px;
            }
        </style>
    </head>
    <body>
        <div id="mapid"></div>
        <script>
            console.log("Ielādē karti...");

            // Inicializēt QWebChannel
            var mapBridge;
            new QWebChannel(qt.webChannelTransport, function(channel) {
                mapBridge = channel.objects.mapBridge;
                console.log("QWebChannel inicializēts!");
            });

            // Inicializēt karti (Rīga, Latvija)
            var mymap = L.map('mapid').setView([56.946285, 24.105078], 13);

            // Pievienot tile layer
            L.tileLayer('https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png', {
                attribution: '&copy; <a href="https://www.openstreetmap.org/copyright">OpenStreetMap</a> contributors',
                maxZoom: 18
            }).addTo(mymap);

            var marker;
            var clickMarkers = [];

            console.log("Karte izveidota!");

            // Funkcija marķiera iestatīšanai
            function setMarker(lat, lon, popupText) {
                console.log("Iestatām marķieri:", lat, lon, popupText);

                if (marker) {
                    mymap.removeLayer(marker);
                }

                marker = L.marker([lat, lon], {
                    draggable: true
                }).addTo(mymap);

                if (popupText && popupText.trim() !== '') {
                    marker.bindPopup(`<div class="custom-popup">${popupText}</div>`).openPopup();
                }

                mymap.setView([lat, lon], mymap.getZoom());

                marker.on('dragend', function(e) {
                    var newPos = e.target.getLatLng();
                    console.log('Marķieris pārvietots:', newPos.lat, newPos.lng);
                });
            }

            // Funkcija koordinātu iegūšanai
            function getCenterCoordinates() {
                var center = mymap.getCenter();
                return {
                    lat: center.lat.toFixed(6),
                    lon: center.lng.toFixed(6)
                };
            }

            // Kartes klikšķa apstrāde
            mymap.on('click', function(e) {
                var lat = e.latlng.lat.toFixed(6);
                var lon = e.latlng.lng.toFixed(6);

                console.log('Kartes klikšķis:', lat, lon);

                // Noņemt iepriekšējos klikšķu marķierus
                clickMarkers.forEach(function(m) { mymap.removeLayer(m); });
                clickMarkers = [];

                var clickMarker = L.circleMarker([lat, lon], {
                    color: 'red',
                    fillColor: '#f03',
                    fillOpacity: 0.5,
                    radius: 5
                }).addTo(mymap);

                clickMarkers.push(clickMarker);

                // Nosūtīt koordinātes uz Python caur QWebChannel
                if (mapBridge) {
                    mapBridge.handleMapClick(lat, lon);
                    console.log('Koordinātes nosūtītas uz Python:', lat, lon);
                } else {
                    console.log('mapBridge nav pieejams!');
                }
            });

            // Pārbaudīt vai karte ielādējās
            mymap.whenReady(function() {
                console.log('Karte gatava!');
            });

            console.log("JavaScript ielādēts!");
        </script>
    </body>
    </html>
        """
        self.web_view.setHtml(map_html_content)

        v.addWidget(self.web_view)

        map_controls_layout = QHBoxLayout()
        self.map_lat_input = QLineEdit()
        self.map_lat_input.setPlaceholderText("Platums (Latitude)")
        self.map_lon_input = QLineEdit()
        self.map_lon_input.setPlaceholderText("Garums (Longitude)")
        btn_set_marker = QPushButton("Iestatīt marķieri")
        btn_set_marker.clicked.connect(self._set_map_marker_from_inputs)
        btn_get_location = QPushButton("Iegūt atrašanās vietu")
        btn_get_location.clicked.connect(self._get_location_from_map)
        btn_set_vieta = QPushButton("Iestatīt 'Vieta' lauku")
        btn_set_vieta.clicked.connect(self._set_vieta_from_map)

        map_controls_layout.addWidget(QLabel("Lat:"))
        map_controls_layout.addWidget(self.map_lat_input)
        map_controls_layout.addWidget(QLabel("Lon:"))
        map_controls_layout.addWidget(self.map_lon_input)
        map_controls_layout.addWidget(btn_set_marker)
        map_controls_layout.addWidget(btn_get_location)
        map_controls_layout.addWidget(btn_set_vieta)

        v.addLayout(map_controls_layout)
        w.setLayout(v)
        self.tabs.addTab(w, "Karte")

        # Initial load of map with current 'Vieta' if possible
        self.tabs.currentChanged.connect(self._update_map_on_tab_change)

    def _create_text_block_input(self, input_widget, field_name):
        """
        Izveido logrīku ar ievades lauku un pogām teksta bloku pārvaldībai.
        """
        container_widget = QWidget()
        layout = QVBoxLayout(container_widget)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.addWidget(input_widget)

        button_layout = QHBoxLayout()
        save_button = QPushButton("Saglabāt kā bloku")
        load_button = QPushButton("Ielādēt bloku")

        save_button.clicked.connect(lambda: self._save_text_block(input_widget, field_name))
        load_button.clicked.connect(lambda: self._load_text_block(input_widget, field_name))

        button_layout.addWidget(save_button)
        button_layout.addWidget(load_button)
        button_layout.addStretch()

        layout.addLayout(button_layout)
        return container_widget

    def _save_text_block(self, input_widget, field_name):
        """
        Saglabā pašreizējo tekstu kā jaunu bloku.
        """
        current_text = ""
        if isinstance(input_widget, QLineEdit):
            current_text = input_widget.text()
        elif isinstance(input_widget, QTextEdit):
            current_text = input_widget.toPlainText()

        if not current_text.strip():
            QMessageBox.warning(self, "Saglabāt bloku", "Ievades lauks ir tukšs. Lūdzu, ievadiet tekstu, ko saglabāt.")
            return

        block_name, ok = QInputDialog.getText(self, "Saglabāt teksta bloku", "Ievadiet bloka nosaukumu:")
        if ok and block_name:
            self.text_block_manager.add_block(field_name, block_name, current_text)
            QMessageBox.information(self, "Saglabāts", f"Teksta bloks '{block_name}' saglabāts laukam '{field_name}'.")
        elif ok:
            QMessageBox.warning(self, "Saglabāt bloku", "Bloka nosaukums nevar būt tukšs.")

    def _load_text_block(self, input_widget, field_name):
        """
        Atver dialogu, lai ielādētu vai pārvaldītu teksta blokus.
        """
        blocks = self.text_block_manager.get_blocks_for_field(field_name)
        if not blocks:
            QMessageBox.information(self, "Ielādēt bloku", "Nav saglabātu teksta bloku šim laukam.")
            return

        dialog = QDialog(self)
        dialog.setWindowTitle(f"Pārvaldīt teksta blokus: {field_name}")
        dialog_layout = QVBoxLayout(dialog)

        block_list = QListWidget()
        for name in sorted(blocks.keys()):
            block_list.addItem(name)
        dialog_layout.addWidget(block_list)

        button_layout = QHBoxLayout()
        load_button = QPushButton("Ielādēt izvēlēto")
        delete_button = QPushButton("Dzēst izvēlēto")
        close_button = QPushButton("Aizvērt")

        load_button.clicked.connect(lambda: self._apply_selected_block(block_list, input_widget, field_name, dialog))
        delete_button.clicked.connect(lambda: self._delete_selected_block(block_list, field_name))
        close_button.clicked.connect(dialog.accept)

        button_layout.addWidget(load_button)
        button_layout.addWidget(delete_button)
        button_layout.addStretch()
        button_layout.addWidget(close_button)
        dialog_layout.addLayout(button_layout)

        dialog.exec()

    def _apply_selected_block(self, block_list, input_widget, field_name, dialog):
        """
        Ielādē izvēlēto bloku ievades laukā.
        """
        selected_item = block_list.currentItem()
        if selected_item:
            block_name = selected_item.text()
            content = self.text_block_manager.get_block_content(field_name, block_name)
            if isinstance(input_widget, QLineEdit):
                input_widget.setText(content)
            elif isinstance(input_widget, QTextEdit):
                input_widget.setPlainText(content)
            dialog.accept()
            QMessageBox.information(self, "Ielādēts", f"Teksta bloks '{block_name}' ielādēts.")
        else:
            QMessageBox.warning(self, "Ielādēt bloku", "Lūdzu, izvēlieties bloku, ko ielādēt.")

    def _delete_selected_block(self, block_list, field_name):
        """
        Dzēš izvēlēto bloku.
        """
        selected_item = block_list.currentItem()
        if selected_item:
            block_name = selected_item.text()
            reply = QMessageBox.question(self, "Dzēst bloku",
                                         f"Vai tiešām vēlaties dzēst teksta bloku '{block_name}'?",
                                         QMessageBox.Yes | QMessageBox.No)
            if reply == QMessageBox.Yes:
                self.text_block_manager.delete_block(field_name, block_name)
                block_list.takeItem(block_list.row(selected_item))
                QMessageBox.information(self, "Dzēsts", f"Teksta bloks '{block_name}' dzēsts.")
        else:
            QMessageBox.warning(self, "Dzēst bloku", "Lūdzu, izvēlieties bloku, ko dzēst.")

    def _update_map_on_tab_change(self, index):
        if self.tabs.tabText(index) == "Karte":
            current_vieta = self.in_vieta.text().strip()
            # Pārliecināmies, ka JavaScript ir ielādēts un gatavs
            # Varētu pievienot arī `page().loadFinished` signāla apstrādi, lai būtu drošāk
            if current_vieta:
                # Šeit varētu mēģināt ģeokodēt adresi, ja ir pieejams API
                # Vienkāršības labad, ja ir vieta, iestatām marķieri Rīgā ar vietas nosaukumu
                self.web_view.page().runJavaScript(f"setMarker(56.946285, 24.105078, '{current_vieta}');")
            else:
                # Ja vieta nav norādīta, iestatām marķieri Rīgā ar noklusējuma tekstu
                self.web_view.page().runJavaScript(f"setMarker(56.946285, 24.105078, 'Rīga');")

    def _set_map_marker_from_inputs(self):
        try:
            lat = float(self.map_lat_input.text())
            lon = float(self.map_lon_input.text())
            self.web_view.page().runJavaScript(f"setMarker({lat}, {lon}, 'Izvēlētā vieta');")
        except ValueError:
            QMessageBox.warning(self, "Kļūda", "Lūdzu, ievadiet derīgas platuma un garuma vērtības.")

    def _handle_map_click(self, lat, lon):
        """
        Šī funkcija tiek izsaukta no MapUrlInterceptor, kad kartē tiek noklikšķināts.
        Tā veic reversās ģeokodēšanas pieprasījumu un atjaunina 'Vieta' lauku.
        """
        self.map_lat_input.setText(lat)
        self.map_lon_input.setText(lon)
        #self._reverse_geocode_and_set_vieta(lat, lon)

    def _reverse_geocode_and_set_vieta(self, lat, lon):
        """
        Veic reversās ģeokodēšanas pieprasījumu uz Nominatim API un iestata 'Vieta' lauku.
        """
        try:
            url = f"https://nominatim.openstreetmap.org/reverse?format=json&lat={lat}&lon={lon}&zoom=18&addressdetails=1"
            headers = {
                'User-Agent': 'AktaGeneratorsApp/1.0 (contact@example.com)' # Obligāti norādīt User-Agent
            }
            response = requests.get(url, headers=headers, timeout=5)
            response.raise_for_status() # Izmet kļūdu, ja HTTP statuss nav 200 OK
            data = response.json()

            if data and 'display_name' in data:
                address = data['display_name']
                self.in_vieta.setText(address)
                QMessageBox.information(self, "Adrese iegūta", f"Adrese: {address}")
            else:
                self.in_vieta.setText(f"Adrese nav atrasta koordinātēm: {lat}, {lon}")
                QMessageBox.warning(self, "Adrese nav atrasta", f"Adrese nav atrasta koordinātēm: {lat}, {lon}")

        except requests.exceptions.RequestException as e:
            QMessageBox.critical(self, "Tīkla kļūda", f"Neizdevās iegūt adresi (tīkla problēma vai API kļūda): {e}")
            self.in_vieta.setText(f"Kļūda iegūstot adresi: {lat}, {lon}")
        except json.JSONDecodeError:
            QMessageBox.critical(self, "JSON kļūda", "Neizdevās parsēt API atbildi.")
            self.in_vieta.setText(f"Kļūda parsējot adresi: {lat}, {lon}")
        except Exception as e:
            QMessageBox.critical(self, "Vispārīga kļūda", f"Neparedzēta kļūda: {e}")
            self.in_vieta.setText(f"Neparedzēta kļūda: {lat}, {lon}")


    def _get_location_from_map(self):
        # Šī funkcija tagad vienkārši iegūst kartes centra koordinātes un iestata tās ievades laukos.
        # Reversā ģeokodēšana notiks, kad lietotājs noklikšķinās uz "Iestatīt 'Vieta' lauku" vai kartē.
        self.web_view.page().runJavaScript("getCenterCoordinates();", self._process_map_center_coordinates)

    def _process_map_center_coordinates(self, result):
        if result and 'lat' in result and 'lon' in result:
            self.map_lat_input.setText(str(result['lat']))
            self.map_lon_input.setText(str(result['lon']))
            QMessageBox.information(self, "Kartes atrašanās vieta", f"Kartes centrs: Lat {result['lat']}, Lon {result['lon']}")
        else:
            QMessageBox.warning(self, "Kļūda", "Neizdevās iegūt kartes centra koordinātes.")

    def _set_vieta_from_map(self):
        """
        Šī poga tagad izmanto pašreizējās ievades laukos esošās koordinātes,
        lai veiktu reversās ģeokodēšanas pieprasījumu un iestatītu 'Vieta' lauku.
        """
        lat_str = self.map_lat_input.text()
        lon_str = self.map_lon_input.text()
        if lat_str and lon_str:
            try:
                lat = float(lat_str)
                lon = float(lon_str)
                self._reverse_geocode_and_set_vieta(str(lat), str(lon)) # Pārsūtam kā string, jo API to sagaida
            except ValueError:
                QMessageBox.warning(self, "Kļūda", "Lūdzu, ievadiet derīgas platuma un garuma vērtības.")
        else:
            QMessageBox.warning(self, "Kļūda", "Lūdzu, vispirms iegūstiet koordinātes no kartes vai ievadiet tās manuāli.")

    # ----- Projekta saglabāšana/ielāde -----
    def ieviest_datus(self, d: AktaDati):
            self.in_akta_nr.setText(d.akta_nr)
            self.in_datums.setDate(datetime.strptime(d.datums, '%Y-%m-%d').date())
            self.in_vieta.setText(d.vieta)
            self.in_pas_nr.setText(d.pasūtījuma_nr)
            self.in_piezimes.setText(d.piezīmes)

            self.in_liguma_nr.setText(d.līguma_nr)
            if d.izpildes_termiņš:
                self.in_izpildes_termins.setDate(datetime.strptime(d.izpildes_termiņš, '%Y-%m-%d').date())
            else:
                self.in_izpildes_termins.setDate(self.in_izpildes_termins.minimumDate()) # Clear date
            if d.pieņemšanas_datums:
                self.in_pieņemšanas_datums.setDate(datetime.strptime(d.pieņemšanas_datums, '%Y-%m-%d').date())
            else:
                self.in_pieņemšanas_datums.setDate(self.in_pieņemšanas_datums.minimumDate())
            if d.nodošanas_datums:
                self.in_nodošanas_datums.setDate(datetime.strptime(d.nodošanas_datums, '%Y-%m-%d').date())
            else:
                self.in_nodošanas_datums.setDate(self.in_nodošanas_datums.minimumDate())

            self.in_strīdu_risināšana.setText(d.strīdu_risināšana)
            self.ck_konfidencialitate.setChecked(d.konfidencialitātes_klauzula)
            self.in_soda_nauda_procenti.setValue(float(d.soda_nauda_procenti)) # Convert Decimal to float for QDoubleSpinBox
            self.in_piegades_nosacijumi.setText(d.piegādes_nosacījumi)
            self.ck_apdrošināšana.setChecked(d.apdrošināšana)
            self.in_papildu_nosacijumi.setText(d.papildu_nosacījumi)
            self.in_atsauces_dokumenti.setText(d.atsauces_dokumenti)
            self.cb_akta_statuss.setCurrentText(d.akta_statuss)
            self.in_valuta.setText(d.valūta)
            self.ck_elektroniskais_paraksts.setChecked(d.elektroniskais_paraksts)
            self.ck_radit_elektronisko_parakstu_tekstu.setChecked(d.radit_elektronisko_parakstu_tekstu) # JAUNA RINDAS

            # Puses
            self.pie_in[0].setText(d.pieņēmējs.nosaukums)
            self.pie_in[1].setText(d.pieņēmējs.reģ_nr)
            self.pie_in[2].setText(d.pieņēmējs.adrese)
            self.pie_in[3].setText(d.pieņēmējs.kontaktpersona)
            self.pie_in[4].setText(d.pieņēmējs.tālrunis)
            self.pie_in[5].setText(d.pieņēmējs.epasts)
            self.pie_in[6].setText(d.pieņēmējs.bankas_konts)
            self.pie_in[7].setCurrentText(d.pieņēmējs.juridiskais_statuss)

            self.nod_in[0].setText(d.nodevējs.nosaukums)
            self.nod_in[1].setText(d.nodevējs.reģ_nr)
            self.nod_in[2].setText(d.nodevējs.adrese)
            self.nod_in[3].setText(d.nodevējs.kontaktpersona)
            self.nod_in[4].setText(d.nodevējs.tālrunis)
            self.nod_in[5].setText(d.nodevējs.epasts)
            self.nod_in[6].setText(d.nodevējs.bankas_konts)
            self.nod_in[7].setCurrentText(d.nodevējs.juridiskais_statuss)

            # Pozīcijas
            self.tab.setRowCount(0)
            # Re-initialize table columns based on loaded settings
            current_col_count = 5 # Base columns
            if d.show_item_serial_number_in_table: current_col_count += 1
            if d.show_item_warranty_in_table: current_col_count += 1
            if d.show_item_notes_in_table: current_col_count += 1
            self.tab.setColumnCount(current_col_count)
            headers = ["Apraksts", "Daudzums", "Vienība", "Cena", "Summa"]
            if d.show_item_serial_number_in_table: headers.append("Seriālais Nr.")
            if d.show_item_warranty_in_table: headers.append("Garantija")
            if d.show_item_notes_in_table: headers.append("Piezīmes pozīcijai")
            self.tab.setHorizontalHeaderLabels(headers)
            self.tab.horizontalHeader().setSectionResizeMode(0, QHeaderView.Stretch)
            for i in range(1, len(headers)):
                self.tab.horizontalHeader().setSectionResizeMode(i, QHeaderView.ResizeToContents)

            for p in d.pozīcijas:
                r = self.tab.rowCount();
                self.tab.insertRow(r)
                self.tab.setItem(r, 0, QTableWidgetItem(p.apraksts))
                self.tab.setItem(r, 1, QTableWidgetItem(str(p.daudzums)))
                self.tab.setItem(r, 2, QTableWidgetItem(p.vienība))
                self.tab.setItem(r, 3, QTableWidgetItem(str(p.cena)))
                self.tab.setItem(r, 4, QTableWidgetItem(formēt_naudu(p.summa)))
                # Set values for potentially hidden columns
                if self.tab.columnCount() > 5: self.tab.setItem(r, 5, QTableWidgetItem(p.seriālais_nr))
                if self.tab.columnCount() > 6: self.tab.setItem(r, 6, QTableWidgetItem(p.garantija))
                if self.tab.columnCount() > 7: self.tab.setItem(r, 7, QTableWidgetItem(p.piezīmes_pozīcijai))


            # Attēli
            self.img_list.clear()
            for a in d.attēli:
                it = QListWidgetItem(os.path.basename(a.ceļš))
                it.setData(Qt.UserRole, {"ceļš": a.ceļš, "paraksts": a.paraksts})
                self.img_list.addItem(it)

            # Iestatījumi
            self.ck_pvn.setChecked(d.iekļaut_pvn)
            self.in_pvn.setValue(float(d.pvn_likme)) # Convert Decimal to float for QDoubleSpinBox
            self.ck_paraksti.setChecked(d.parakstu_rindas)
            self.in_logo.setText(d.logotipa_ceļš)
            self.in_fonts.setText(d.fonts_ceļš)
            self.in_paraksts_pie.setText(d.paraksts_pieņēmējs_ceļš)
            self.in_paraksts_nod.setText(d.paraksts_nodevējs_ceļš)

            # Papildu iestatījumi
            self.cb_page_size.setCurrentText(d.pdf_page_size)
            self.cb_page_orientation.setCurrentText(d.pdf_page_orientation)
            self.in_margin_left.setValue(float(d.pdf_margin_left)) # Convert Decimal to float for QDoubleSpinBox
            self.in_margin_right.setValue(float(d.pdf_margin_right)) # Convert Decimal to float for QDoubleSpinBox
            self.in_margin_top.setValue(float(d.pdf_margin_top)) # Convert Decimal to float for QDoubleSpinBox
            self.in_margin_bottom.setValue(float(d.pdf_margin_bottom)) # Convert Decimal to float for QDoubleSpinBox
            self.in_font_size_head.setValue(d.pdf_font_size_head)
            self.in_font_size_normal.setValue(d.pdf_font_size_normal)
            self.in_font_size_small.setValue(d.pdf_font_size_small)
            self.in_font_size_table.setValue(d.pdf_font_size_table)
            self.in_logo_width_mm.setValue(float(d.pdf_logo_width_mm)) # Convert Decimal to float for QDoubleSpinBox
            self.in_signature_width_mm.setValue(float(d.pdf_signature_width_mm)) # Convert Decimal to float for QDoubleSpinBox
            self.in_signature_height_mm.setValue(float(d.pdf_signature_height_mm)) # Convert Decimal to float for QDoubleSpinBox
            self.in_docx_image_width_inches.setValue(float(d.docx_image_width_inches)) # Convert Decimal to float for QDoubleSpinBox
            self.in_docx_signature_width_inches.setValue(float(d.docx_signature_width_inches)) # Convert Decimal to float for QDoubleSpinBox
            self.in_table_col_widths.setText(d.table_col_widths)
            self.ck_auto_generate_akta_nr.setChecked(d.auto_generate_akta_nr)
            self.in_default_currency.setText(d.default_currency)
            self.in_default_unit.setText(d.default_unit)
            self.in_default_pvn_rate.setValue(float(d.default_pvn_rate)) # Convert Decimal to float for QDoubleSpinBox
            self.in_poppler_path.setText(d.poppler_path)
            self.poppler_path = d.poppler_path

            # Set new settings values
            self.in_header_text_color.setText(d.header_text_color)
            self.in_footer_text_color.setText(d.footer_text_color)
            self.in_table_header_bg_color.setText(d.table_header_bg_color)
            self.in_table_grid_color.setText(d.table_grid_color)
            self.in_table_row_spacing.setValue(float(d.table_row_spacing)) # Convert Decimal to float for QDoubleSpinBox
            self.in_line_spacing_multiplier.setValue(float(d.line_spacing_multiplier)) # Convert Decimal to float for QDoubleSpinBox
            self.ck_show_page_numbers.setChecked(d.show_page_numbers)
            self.ck_show_generation_timestamp.setChecked(d.show_generation_timestamp)
            self.cb_currency_symbol_position.setCurrentText(d.currency_symbol_position)
            self.in_date_format.setText(d.date_format)
            self.in_signature_line_length_mm.setValue(float(d.signature_line_length_mm)) # Convert Decimal to float for QDoubleSpinBox
            self.in_signature_line_thickness_pt.setValue(float(d.signature_line_thickness_pt)) # Convert Decimal to float for QDoubleSpinBox
            self.ck_add_cover_page.setChecked(d.add_cover_page)
            self.in_cover_page_title.setText(d.cover_page_title)
            self.in_cover_page_logo_width_mm.setValue(float(d.cover_page_logo_width_mm)) # Convert Decimal to float for QDoubleSpinBox
            # Individuālais QR kods
            self.ck_include_custom_qr_code.setChecked(d.include_custom_qr_code)
            self.in_custom_qr_code_data.setText(d.custom_qr_code_data)
            self.in_custom_qr_code_size_mm.setValue(float(d.custom_qr_code_size_mm))
            self.cb_custom_qr_code_position.setCurrentText(d.custom_qr_code_position)
            self.in_custom_qr_code_pos_x_mm.setValue(float(d.custom_qr_code_pos_x_mm))
            self.in_custom_qr_code_pos_y_mm.setValue(float(d.custom_qr_code_pos_y_mm))
            self.in_custom_qr_code_color.setText(d.custom_qr_code_color)

            # Automātiskais QR kods (akta ID)
            self.ck_include_auto_qr_code.setChecked(d.include_auto_qr_code)
            self.in_auto_qr_code_size_mm.setValue(float(d.auto_qr_code_size_mm))
            self.cb_auto_qr_code_position.setCurrentText(d.auto_qr_code_position)
            self.in_auto_qr_code_pos_x_mm.setValue(float(d.auto_qr_code_pos_x_mm))
            self.in_auto_qr_code_pos_y_mm.setValue(float(d.auto_qr_code_pos_y_mm))
            self.in_auto_qr_code_color.setText(d.auto_qr_code_color)

            self.ck_add_watermark.setChecked(d.add_watermark)
            self.in_watermark_text.setText(d.watermark_text)
            self.in_watermark_font_size.setValue(d.watermark_font_size)
            self.in_watermark_color.setText(d.watermark_color)
            self.in_watermark_rotation.setValue(d.watermark_rotation)
            self.ck_enable_pdf_encryption.setChecked(d.enable_pdf_encryption)
            self.in_pdf_user_password.setText(d.pdf_user_password)
            self.in_pdf_owner_password.setText(d.pdf_owner_password)
            self.ck_allow_printing.setChecked(d.allow_printing)
            self.ck_allow_copying.setChecked(d.allow_copying)
            self.ck_allow_modifying.setChecked(d.allow_modifying)
            self.ck_allow_annotating.setChecked(d.allow_annotating)
            self.in_default_country.setText(d.default_country)
            self.in_default_city.setText(d.default_city)
            self.ck_show_contact_details_in_header.setChecked(d.show_contact_details_in_header)
            self.in_contact_details_header_font_size.setValue(d.contact_details_header_font_size)
            self.in_item_image_width_mm.setValue(float(d.item_image_width_mm)) # Convert Decimal to float for QDoubleSpinBox
            self.in_item_image_caption_font_size.setValue(d.item_image_caption_font_size)
            self.ck_show_item_notes_in_table.setChecked(d.show_item_notes_in_table)
            self.ck_show_item_serial_number_in_table.setChecked(d.show_item_serial_number_in_table)
            self.ck_show_item_warranty_in_table.setChecked(d.show_item_warranty_in_table)
            self.in_table_cell_padding_mm.setValue(float(d.table_cell_padding_mm)) # Convert Decimal to float for QDoubleSpinBox
            self.cb_table_header_font_style.setCurrentText(d.table_header_font_style)
            self.cb_table_content_alignment.setCurrentText(d.table_content_alignment)
            self.in_signature_font_size.setValue(d.signature_font_size)
            self.in_signature_spacing_mm.setValue(float(d.signature_spacing_mm)) # Convert Decimal to float for QDoubleSpinBox
            self.in_document_title_font_size.setValue(d.document_title_font_size)
            self.in_document_title_color.setText(d.document_title_color)
            self.in_section_heading_font_size.setValue(d.section_heading_font_size)
            self.in_section_heading_color.setText(d.section_heading_color)
            self.in_paragraph_line_spacing_multiplier.setValue(float(d.paragraph_line_spacing_multiplier)) # Convert Decimal to float for QDoubleSpinBox
            self.cb_table_border_style.setCurrentText(d.table_border_style)
            self.in_table_border_thickness_pt.setValue(float(d.table_border_thickness_pt)) # Convert Decimal to float for QDoubleSpinBox
            self.in_table_alternate_row_color.setText(d.table_alternate_row_color)
            self.ck_show_total_sum_in_words.setChecked(d.show_total_sum_in_words)
            self.cb_total_sum_in_words_language.setCurrentText(d.total_sum_in_words_language)
            self.cb_default_vat_calculation_method.setCurrentText(d.default_vat_calculation_method)
            self.ck_show_vat_breakdown.setChecked(d.show_vat_breakdown)
            self.ck_enable_digital_signature_field.setChecked(d.enable_digital_signature_field)
            self.in_digital_signature_field_name.setText(d.digital_signature_field_name)
            self.in_digital_signature_field_size_mm.setValue(
                float(d.digital_signature_field_size_mm))  # Convert Decimal to float for QDoubleSpinBox
            self.cb_digital_signature_field_position.setCurrentText(d.digital_signature_field_position)

            # JAUNA RINDAS
            self.in_templates_dir.setText(d.templates_dir)
            # Pārliecināmies, ka direktorijs eksistē, ja tas ir ielādēts
            if d.templates_dir:
                os.makedirs(d.templates_dir, exist_ok=True)

            self._update_preview()

    def saglabat_projektu(self):
        d = self.savākt_datus()
        default_filename = f"Akts_{drošs_faila_nosaukums(d.akta_nr) or 'akts'}.json"
        path, _ = QFileDialog.getSaveFileName(self, "Saglabāt projektu", os.path.join(PROJECT_SAVE_DIR, default_filename), "JSON (*.json)")
        if not path:
            return
        out = asdict(d)
        # Convert Decimal fields to string for JSON serialization
        for key, value in out.items():
            if isinstance(value, Decimal):
                out[key] = str(value)
        # Handle nested Decimal fields in Pozīcija
        for p in out['pozīcijas']:
            for key, value in p.items():
                if isinstance(value, Decimal):
                    p[key] = str(value)
        # Handle nested Persona objects
        out['pieņēmējs'] = asdict(d.pieņēmējs)
        out['nodevējs'] = asdict(d.nodevējs)

        with open(path, 'w', encoding='utf-8') as f:
            json.dump(out, f, ensure_ascii=False, indent=2)
        self._ceļš_projekts = path
        self._add_to_history(path) # Pievienojam projektu vēsturei
        QMessageBox.information(self, "Saglabāts", "Projekts saglabāts veiksmīgi.")


    def saglabat_ka_sablonu(self):
        """
        Saglabā pašreizējo akta konfigurāciju kā šablonu.
        """
        d = self.savākt_datus()

        # Pieprasām šablona nosaukumu no lietotāja
        template_name, ok = QInputDialog.getText(self, "Saglabāt kā šablonu", "Ievadiet šablona nosaukumu:")
        if not ok or not template_name:
            QMessageBox.warning(self, "Saglabāt kā šablonu", "Šablona nosaukums nevar būt tukšs.")
            return

        # Pieprasām paroli šablonam (neobligāti)
        template_password, ok_pass = QInputDialog.getText(self, "Šablona parole",
                                                          "Ievadiet paroli šablonam (atstājiet tukšu, ja nevēlaties paroli):",
                                                          QLineEdit.Password)
        if not ok_pass:
            return  # Lietotājs atcēla paroles ievadi

        d.template_password = template_password  # Saglabājam paroli datu objektā

        # Izveidojam drošu faila nosaukumu
        safe_template_name = drošs_faila_nosaukums(template_name)
        # Izmantojam pašreizējo šablonu direktoriju no AktaDati
        current_templates_dir = self.data.templates_dir
        if not current_templates_dir:  # Ja vēl nav iestatīts (piemēram, pirmajā startā)
            current_templates_dir = os.path.join(APP_DATA_DIR, "AktaGenerators_Templates")
            os.makedirs(current_templates_dir, exist_ok=True)  # Pārliecināmies, ka noklusējuma mape eksistē

        template_file_path = os.path.join(current_templates_dir, f"{safe_template_name}.json")

        # Pārveidojam AktaDati objektu vārdnīcā, lai to varētu serializēt uz JSON
        out = asdict(d)

        # Konvertējam Decimal vērtības uz string, lai tās varētu saglabāt JSON
        for key, value in out.items():
            if isinstance(value, Decimal):
                out[key] = str(value)
        # Apstrādājam ligzdotās Decimal vērtības Pozīcija objektos
        for p in out.get('pozīcijas', []): # Izmantojam .get, lai izvairītos no kļūdām, ja pozīcijas nav
            for key, value in p.items():
                if isinstance(value, Decimal):
                    p[key] = str(value)

        # Apstrādājam Persona objektus
        out['pieņēmējs'] = asdict(d.pieņēmējs)
        out['nodevējs'] = asdict(d.nodevējs)

        # Izdzēšam datus, kas nav jāglabā šablonā (piemēram, specifiskus akta datus, attēlus, paroles)
        # Šablons ir paredzēts kā bāzes konfigurācija, nevis konkrēta akta kopija.
        out['akta_nr'] = ""
        out['datums'] = datetime.now().strftime('%Y-%m-%d') # Atjaunojam datumu uz pašreizējo
        out['pasūtījuma_nr'] = ""
        out['līguma_nr'] = ""
        out['izpildes_termiņš'] = ""
        out['pieņemšanas_datums'] = ""
        out['nodošanas_datums'] = ""
        out['piezīmes'] = ""
        out['attēli'] = [] # Šablonā nav jābūt attēliem
        out['poppler_path'] = "" # Poppler ceļš ir sistēmas iestatījums, nevis šablona daļa
        out['pdf_user_password'] = ""
        out['pdf_owner_password'] = ""
        out['custom_qr_code_data'] = ""  # QR koda dati nav šablona daļa
        out['custom_qr_code_pos_x_mm'] = str(Decimal("0"))
        out['custom_qr_code_pos_y_mm'] = str(Decimal("0"))
        out['custom_qr_code_color'] = "#000000"
        out['auto_qr_code_pos_x_mm'] = str(Decimal("0"))
        out['auto_qr_code_pos_y_mm'] = str(Decimal("0"))
        out['auto_qr_code_color'] = "#000000"
        out['template_password'] = template_password

        # Saglabājam šablonu JSON failā
        try:
            with open(template_file_path, 'w', encoding='utf-8') as f:
                json.dump(out, f, ensure_ascii=False, indent=2)
            QMessageBox.information(self, "Saglabāts", f"Šablons '{template_name}' veiksmīgi saglabāts.")
            self._update_sablonu_list() # Atjaunojam šablonu sarakstu
        except Exception as e:
            QMessageBox.critical(self, "Kļūda", f"Neizdevās saglabāt šablonu:\n{e}")

    def _update_sablonu_list(self):
        """
        Atjauno šablonu sarakstu GUI.
        """
        self.sablonu_list.clear()
        # Pievienojam noklusējuma šablonu
        self.sablonu_list.addItem("Pappus dati (piemērs)")
        # Pievienojam visus saglabātos šablonus
        if os.path.exists(TEMPLATES_DIR):
            for filename in os.listdir(TEMPLATES_DIR):
                if filename.endswith(".json"):
                    template_name = os.path.splitext(filename)[0]
                    self.sablonu_list.addItem(template_name)


    def dzest_sablonus(self):
        """
        Dzēš atlasītos šablonus.
        """
        selected_items = self.sablonu_list.selectedItems()
        if not selected_items:
            QMessageBox.warning(self, "Dzēst šablonus", "Lūdzu, atlasiet vismaz vienu šablonu, ko dzēst.")
            return

        # Iegūstam tīrus šablonu nosaukumus, noņemot marķierus
        templates_to_delete_clean = []
        for item in selected_items:
            clean_name = item.text().replace(" [Aizsargāts]", "").replace(" [Kļūda]", "")
            if clean_name != "Pappus dati (piemērs)":  # Neļaujam dzēst iebūvēto šablonu
                templates_to_delete_clean.append(clean_name)

        if not templates_to_delete_clean:
            QMessageBox.information(self, "Dzēst šablonus",
                                    "Nav atlasīts neviens dzēšams šablons (iebūvēto šablonu nevar dzēst).")
            return

        reply = QMessageBox.question(self, "Dzēst šablonus",
                                     f"Vai tiešām vēlaties dzēst šādus šablonus?\n\n{', '.join(templates_to_delete_clean)}",
                                     QMessageBox.Yes | QMessageBox.No)

        if reply == QMessageBox.Yes:
            deleted_count = 0
            current_templates_dir = self.data.templates_dir
            if not current_templates_dir:
                current_templates_dir = os.path.join(APP_DATA_DIR, "AktaGenerators_Templates")

            for template_name in templates_to_delete_clean:
                file_path = os.path.join(current_templates_dir, f"{template_name}.json")
                if os.path.exists(file_path):
                    try:
                        with open(file_path, 'r', encoding='utf-8') as f:
                            data = json.load(f)
                        stored_password = data.get('template_password', '')

                        if stored_password:
                            entered_password, ok_pass = QInputDialog.getText(self, "Ievadiet paroli",
                                                                             f"Šablonam '{template_name}' ir parole. Lūdzu, ievadiet to, lai dzēstu:",
                                                                             QLineEdit.Password)
                            if not ok_pass or entered_password != stored_password:
                                QMessageBox.warning(self, "Nepareiza parole",
                                                    f"Nepareiza parole šablonam '{template_name}'. Dzēšana atcelta.")
                                continue  # Pārejam pie nākamā šablona

                        os.remove(file_path)
                        deleted_count += 1
                    except Exception as e:
                        QMessageBox.critical(self, "Kļūda dzēšot", f"Neizdevās dzēst šablonu '{template_name}':\n{e}")
                else:
                    QMessageBox.warning(self, "Dzēst šablonus", f"Šablona fails '{template_name}.json' nav atrasts.")

            if deleted_count > 0:
                QMessageBox.information(self, "Dzēsts", f"Veiksmīgi dzēsti {deleted_count} šabloni.")
                self._update_sablonu_list()  # Atjaunojam sarakstu pēc dzēšanas
            else:
                QMessageBox.information(self, "Dzēst šablonus", "Neviens šablons netika dzēsts.")


    def _update_sablonu_list(self):
        """
        Atjauno šablonu sarakstu GUI, pievienojot marķieri aizsargātiem šabloniem.
        """
        self.sablonu_list.clear()
        # Pievienojam noklusējuma šablonu
        self.sablonu_list.addItem("Pappus dati (piemērs)")

        # Izmantojam pašreizējo šablonu direktoriju no AktaDati
        current_templates_dir = self.data.templates_dir
        if not current_templates_dir:  # Ja vēl nav iestatīts (piemēram, pirmajā startā)
            current_templates_dir = os.path.join(APP_DATA_DIR, "AktaGenerators_Templates")
            os.makedirs(current_templates_dir, exist_ok=True)  # Pārliecināmies, ka noklusējuma mape eksistē

        if os.path.exists(current_templates_dir):
            for filename in os.listdir(current_templates_dir):
                if filename.endswith(".json"):
                    template_name = os.path.splitext(filename)[0]
                    file_path = os.path.join(current_templates_dir, filename)
                    item = QListWidgetItem(template_name)  # Izveidojam QListWidgetItem ar tīru nosaukumu
                    try:
                        with open(file_path, 'r', encoding='utf-8') as f:
                            data = json.load(f)
                        if data.get('template_password'):
                            item.setBackground(QColor("#F54927"))  # Gaiši oranžs fons aizsargātiem šabloniem
                            item.setToolTip("Šablons ir aizsargāts ar paroli")  # Pievienojam tooltip
                        self.sablonu_list.addItem(item)
                    except Exception as e:
                        print(f"Kļūda lasot šablona failu {filename}: {e}")
                        item.setBackground(QColor("#FFCCCC"))  # Gaiši sarkans fons kļūdainiem šabloniem
                        item.setToolTip(f"Kļūda ielādējot šablonu: {e}")
                        self.sablonu_list.addItem(item)


    def mainit_sablonu_paroli(self):
        """
        Maina vai pievieno paroli izvēlētajam šablonam.
        """
        selected_items = self.sablonu_list.selectedItems()
        if not selected_items:
            QMessageBox.warning(self, "Mainīt paroli", "Lūdzu, atlasiet šablonu, kuram vēlaties mainīt paroli.")
            return
        if len(selected_items) > 1:
            QMessageBox.warning(self, "Mainīt paroli", "Lūdzu, atlasiet tikai vienu šablonu.")
            return

        item_text = selected_items[0].text()
        template_name = item_text.replace(" [Aizsargāts]", "").replace(" [Kļūda]", "")

        if template_name == "Pappus dati (piemērs)":
            QMessageBox.information(self, "Mainīt paroli",
                                    "Iebūvētajam šablonam 'Pappus dati (piemērs)' nevar mainīt paroli.")
            return

        current_templates_dir = self.data.templates_dir
        if not current_templates_dir:
            current_templates_dir = os.path.join(APP_DATA_DIR, "AktaGenerators_Templates")

        file_path = os.path.join(current_templates_dir, f"{template_name}.json")
        if not os.path.exists(file_path):
            QMessageBox.warning(self, "Mainīt paroli", f"Šablona fails '{template_name}.json' nav atrasts.")
            return

        try:
            with open(file_path, 'r', encoding='utf-8') as f:
                data = json.load(f)

            current_password = data.get('template_password', '')

            if current_password:
                # Šablonam jau ir parole, pieprasām veco paroli
                old_password, ok_old = QInputDialog.getText(self, "Mainīt paroli",
                                                             f"Šablonam '{template_name}' ir parole. Lūdzu, ievadiet veco paroli:",
                                                             QLineEdit.Password)
                if not ok_old or old_password != current_password:
                    QMessageBox.warning(self, "Nepareiza parole", "Ievadītā vecā parole ir nepareiza vai ievade atcelta.")
                    return

            # Pieprasām jauno paroli
            new_password, ok_new = QInputDialog.getText(self, "Mainīt paroli",
                                                         f"Ievadiet jauno paroli šablonam '{template_name}' (atstājiet tukšu, lai noņemtu paroli):",
                                                         QLineEdit.Password)
            if not ok_new:
                return # Lietotājs atcēla jaunas paroles ievadi

            data['template_password'] = new_password # Saglabājam jauno paroli (var būt tukša)

            with open(file_path, 'w', encoding='utf-8') as f:
                json.dump(data, f, ensure_ascii=False, indent=2)

            QMessageBox.information(self, "Parole mainīta", f"Šablona '{template_name}' parole veiksmīgi mainīta.")
            self._update_sablonu_list() # Atjaunojam sarakstu, lai atspoguļotu izmaiņas
        except Exception as e:
            QMessageBox.critical(self, "Kļūda", f"Neizdevās mainīt paroli šablonam '{template_name}':\n{e}")

    def nonemt_sablonu_paroli(self):
        """
        Noņem paroli no izvēlētā šablona.
        """
        selected_items = self.sablonu_list.selectedItems()
        if not selected_items:
            QMessageBox.warning(self, "Noņemt paroli", "Lūdzu, atlasiet šablonu, kuram vēlaties noņemt paroli.")
            return
        if len(selected_items) > 1:
            QMessageBox.warning(self, "Noņemt paroli", "Lūdzu, atlasiet tikai vienu šablonu.")
            return

        item_text = selected_items[0].text()
        template_name = item_text.replace(" [Aizsargāts]", "").replace(" [Kļūda]", "")

        if template_name == "Pappus dati (piemērs)":
            QMessageBox.information(self, "Noņemt paroli",
                                    "Iebūvētajam šablonam 'Pappus dati (piemērs)' nav paroles, ko noņemt.")
            return

        current_templates_dir = self.data.templates_dir
        if not current_templates_dir:
            current_templates_dir = os.path.join(APP_DATA_DIR, "AktaGenerators_Templates")

        file_path = os.path.join(current_templates_dir, f"{template_name}.json")
        if not os.path.exists(file_path):
            QMessageBox.warning(self, "Noņemt paroli", f"Šablona fails '{template_name}.json' nav atrasts.")
            return

        try:
            with open(file_path, 'r', encoding='utf-8') as f:
                data = json.load(f)

            current_password = data.get('template_password', '')

            if not current_password:
                QMessageBox.information(self, "Noņemt paroli", f"Šablonam '{template_name}' jau nav paroles.")
                return

            # Pieprasām veco paroli, lai apstiprinātu noņemšanu
            old_password, ok_old = QInputDialog.getText(self, "Noņemt paroli",
                                                         f"Šablonam '{template_name}' ir parole. Lūdzu, ievadiet to, lai noņemtu:",
                                                         QLineEdit.Password)
            if not ok_old or old_password != current_password:
                QMessageBox.warning(self, "Nepareiza parole", "Ievadītā parole ir nepareiza vai ievade atcelta.")
                return

            data['template_password'] = "" # Noņemam paroli

            with open(file_path, 'w', encoding='utf-8') as f:
                json.dump(data, f, ensure_ascii=False, indent=2)

            QMessageBox.information(self, "Parole noņemta", f"Parole no šablona '{template_name}' veiksmīgi noņemta.")
            self._update_sablonu_list() # Atjaunojam sarakstu, lai atspoguļotu izmaiņas
        except Exception as e:
            QMessageBox.critical(self, "Kļūda", f"Neizdevās noņemt paroli no šablona '{template_name}':\n{e}")

    def ieladet_projektu(self, path=None):
        if not path:
            path, _ = QFileDialog.getOpenFileName(self, "Ielādēt projektu", PROJECT_SAVE_DIR, "JSON (*.json)")
        if not path:
            return
        try:
            # Mēģinām ielādēt ar dažādām kodēšanām
            encodings = ['utf-8', 'utf-8-sig', 'cp1257', 'iso-8859-1', 'windows-1252']
            data = None

            for encoding in encodings:
                try:
                    with open(path, 'r', encoding=encoding) as f:
                        data = json.load(f)
                    break
                except (UnicodeDecodeError, json.JSONDecodeError):
                    continue

            if data is None:
                raise Exception("Neizdevās ielādēt failu ar nevienu no atbalstītajām kodēšanām")

            # Helper to safely get and convert Decimal values
            # Helper to safely get and convert Decimal values
            def get_decimal(dict_obj, key, default_val):
                val = dict_obj.get(key, default_val)
                return to_decimal(val)

            # Helper to safely get boolean values
            def get_bool(dict_obj, key, default_val):
                val = dict_obj.get(key, default_val)
                return bool(val)

            d = AktaDati(
                    akta_nr=data.get('akta_nr', ''), datums=data.get('datums', ''), vieta=data.get('vieta', ''),
                    pasūtījuma_nr=data.get('pasūtījuma_nr', ''),
                    pieņēmējs=Persona(
                        nosaukums=data.get('pieņēmējs', {}).get('nosaukums', ''),
                        reģ_nr=data.get('pieņēmējs', {}).get('reģ_nr', ''),
                        adrese=data.get('pieņēmējs', {}).get('adrese', ''),
                        kontaktpersona=data.get('pieņēmējs', {}).get('kontaktpersona', ''),
                        tālrunis=data.get('pieņēmējs', {}).get('tālrunis', ''),
                        epasts=data.get('pieņēmējs', {}).get('epasts', ''),
                        bankas_konts=data.get('pieņēmējs', {}).get('bankas_konts', ''),
                        juridiskais_statuss=data.get('pieņēmējs', {}).get('juridiskais_statuss', '')
                    ),
                    nodevējs=Persona(
                        nosaukums=data.get('nodevējs', {}).get('nosaukums', ''),
                        reģ_nr=data.get('nodevējs', {}).get('reģ_nr', ''),
                        adrese=data.get('nodevējs', {}).get('adrese', ''),
                        kontaktpersona=data.get('nodevējs', {}).get('kontaktpersona', ''),
                        tālrunis=data.get('nodevējs', {}).get('tālrunis', ''),
                        epasts=data.get('nodevējs', {}).get('epasts', ''),
                        bankas_konts=data.get('nodevējs', {}).get('bankas_konts', ''),
                        juridiskais_statuss=data.get('nodevējs', {}).get('juridiskais_statuss', '')
                    ),
                    pozīcijas=[Pozīcija(
                        apraksts=p.get('apraksts', ''),
                        daudzums=get_decimal(p, 'daudzums', '0'),
                        vienība=p.get('vienība', 'gab.'),
                        cena=get_decimal(p, 'cena', '0'),
                        seriālais_nr=p.get('seriālais_nr', ''),
                        garantija=p.get('garantija', ''),
                        piezīmes_pozīcijai=p.get('piezīmes_pozīcijai', '')
                    ) for p in data.get('pozīcijas', [])],
                    attēli=[Attēls(**a) for a in data.get('attēli', [])],
                    piezīmes=data.get('piezīmes', ''), iekļaut_pvn=get_bool(data, 'iekļaut_pvn', False),
                    pvn_likme=get_decimal(data, 'pvn_likme', '21'),
                    parakstu_rindas=get_bool(data, 'parakstu_rindas', True),
                    logotipa_ceļš=data.get('logotipa_ceļš', ''), fonts_ceļš=data.get('fonts_ceļš', ''),
                    paraksts_pieņēmējs_ceļš=data.get('paraksts_pieņēmējs_ceļš', ''),
                    paraksts_nodevējs_ceļš=data.get('paraksts_nodevējs_ceļš', ''),
                    līguma_nr=data.get('līguma_nr', ''),
                    izpildes_termiņš=data.get('izpildes_termiņš', ''),
                    pieņemšanas_datums=data.get('pieņemšanas_datums', ''),
                    nodošanas_datums=data.get('nodošanas_datums', ''),
                    strīdu_risināšana=data.get('strīdu_risināšana', ''),
                    konfidencialitātes_klauzula=get_bool(data, 'konfidencialitātes_klauzula', False),
                    soda_nauda_procenti=get_decimal(data, 'soda_nauda_procenti', '0.0'),
                    piegādes_nosacījumi=data.get('piegādes_nosacījumi', ''),
                    apdrošināšana=get_bool(data, 'apdrošināšana', False),
                    papildu_nosacījumi=data.get('papildu_nosacījumi', ''),
                    atsauces_dokumenti=data.get('atsauces_dokumenti', ''),
                    akta_statuss=data.get('akta_statuss', 'Melnraksts'),
                    valūta=data.get('valūta', 'EUR'),
                    elektroniskais_paraksts=get_bool(data, 'elektroniskais_paraksts', False),
                    radit_elektronisko_parakstu_tekstu=get_bool(data, 'radit_elektronisko_parakstu_tekstu', False),
                # JAUNA RINDAS
                    pdf_page_size=data.get('pdf_page_size', 'A4'),
                    pdf_page_orientation=data.get('pdf_page_orientation', 'Portrets'),
                    pdf_margin_left=get_decimal(data, 'pdf_margin_left', '18'),
                    pdf_margin_right=get_decimal(data, 'pdf_margin_right', '18'),
                    pdf_margin_top=get_decimal(data, 'pdf_margin_top', '16'),
                    pdf_margin_bottom=get_decimal(data, 'pdf_margin_bottom', '16'),
                    pdf_font_size_head=data.get('pdf_font_size_head', 14),
                    pdf_font_size_normal=data.get('pdf_font_size_normal', 10),
                    pdf_font_size_small=data.get('pdf_font_size_small', 9),
                    pdf_font_size_table=data.get('pdf_font_size_table', 9),
                    pdf_logo_width_mm=get_decimal(data, 'pdf_logo_width_mm', '35'),
                    pdf_signature_width_mm=get_decimal(data, 'pdf_signature_width_mm', '50'),
                    pdf_signature_height_mm=get_decimal(data, 'pdf_signature_height_mm', '20'),
                    docx_image_width_inches=get_decimal(data, 'docx_image_width_inches', '4'),
                    docx_signature_width_inches=get_decimal(data, 'docx_signature_width_inches', '1.5'),
                    table_col_widths=data.get('table_col_widths', '10,40,18,18,20,20,25,25,25'),
                    auto_generate_akta_nr=get_bool(data, 'auto_generate_akta_nr', False),
                    default_currency=data.get('default_currency', 'EUR'),
                    default_unit=data.get('default_unit', 'gab.'),
                    default_pvn_rate=get_decimal(data, 'default_pvn_rate', '21.0'),
                    poppler_path=data.get('poppler_path', ''),
                    # Load new settings
                    header_text_color=data.get('header_text_color', '#000000'),
                    footer_text_color=data.get('footer_text_color', '#000000'),
                    table_header_bg_color=data.get('table_header_bg_color', '#E0E0E0'),
                    table_grid_color=data.get('table_grid_color', '#CCCCCC'),
                    table_row_spacing=get_decimal(data, 'table_row_spacing', '4'),
                    line_spacing_multiplier=get_decimal(data, 'line_spacing_multiplier', '1.2'),
                    show_page_numbers=get_bool(data, 'show_page_numbers', True),
                    show_generation_timestamp=get_bool(data, 'show_generation_timestamp', True),
                    currency_symbol_position=data.get('currency_symbol_position', 'after'),
                    date_format=data.get('date_format', 'YYYY-MM-DD'),
                    signature_line_length_mm=get_decimal(data, 'signature_line_length_mm', '60'),
                    signature_line_thickness_pt=get_decimal(data, 'signature_line_thickness_pt', '0.5'),
                    add_cover_page=get_bool(data, 'add_cover_page', False),
                    cover_page_title=data.get('cover_page_title', 'Pieņemšanas-Nodošanas Akts'),
                    cover_page_logo_width_mm=get_decimal(data, 'cover_page_logo_width_mm', '80'),
                # Individuālais QR kods
                include_custom_qr_code=get_bool(data, 'include_custom_qr_code', False),
                custom_qr_code_data=data.get('custom_qr_code_data', ''),
                custom_qr_code_size_mm=get_decimal(data, 'custom_qr_code_size_mm', '20'),
                custom_qr_code_position=data.get('custom_qr_code_position', 'bottom_right'),
                custom_qr_code_pos_x_mm=get_decimal(data, 'custom_qr_code_pos_x_mm', '0'),
                custom_qr_code_pos_y_mm=get_decimal(data, 'custom_qr_code_pos_y_mm', '0'),
                custom_qr_code_color=data.get('custom_qr_code_color', '#000000'),

                # Automātiskais QR kods (akta ID)
                include_auto_qr_code=get_bool(data, 'include_auto_qr_code', False),
                auto_qr_code_size_mm=get_decimal(data, 'auto_qr_code_size_mm', '20'),
                auto_qr_code_position=data.get('auto_qr_code_position', 'bottom_left'),
                auto_qr_code_pos_x_mm=get_decimal(data, 'auto_qr_code_pos_x_mm', '0'),
                auto_qr_code_pos_y_mm=get_decimal(data, 'auto_qr_code_pos_y_mm', '0'),
                auto_qr_code_color=data.get('auto_qr_code_color', '#000000'),

                add_watermark=get_bool(data, 'add_watermark', False),
                    watermark_text=data.get('watermark_text', 'MELNRAKSTS'),
                    watermark_font_size=data.get('watermark_font_size', 72),
                    watermark_color=data.get('watermark_color', '#E0E0E0'),
                    watermark_rotation=data.get('watermark_rotation', 45),
                    enable_pdf_encryption=get_bool(data, 'enable_pdf_encryption', False),
                    pdf_user_password=data.get('pdf_user_password', ''),
                    pdf_owner_password=data.get('pdf_owner_password', ''),
                    allow_printing=get_bool(data, 'allow_printing', True),
                    allow_copying=get_bool(data, 'allow_copying', True),
                    allow_modifying=get_bool(data, 'allow_modifying', False),
                    allow_annotating=get_bool(data, 'allow_annotating', True),
                    default_country=data.get('default_country', 'Latvija'),
                    default_city=data.get('default_city', 'Rīga'),
                    show_contact_details_in_header=get_bool(data, 'show_contact_details_in_header', False),
                    contact_details_header_font_size=data.get('contact_details_header_font_size', 8),
                    item_image_width_mm=get_decimal(data, 'item_image_width_mm', '50'),
                    item_image_caption_font_size=data.get('item_image_caption_font_size', 8),
                    show_item_notes_in_table=get_bool(data, 'show_item_notes_in_table', True),
                    show_item_serial_number_in_table=get_bool(data, 'show_item_serial_number_in_table', True),
                    show_item_warranty_in_table=get_bool(data, 'show_item_warranty_in_table', True),
                    table_cell_padding_mm=get_decimal(data, 'table_cell_padding_mm', '2'),
                    table_header_font_style=data.get('table_header_font_style', 'bold'),
                    table_content_alignment=data.get('table_content_alignment', 'left'),
                    signature_font_size=data.get('signature_font_size', 9),
                    signature_spacing_mm=get_decimal(data, 'signature_spacing_mm', '10'),
                    document_title_font_size=data.get('document_title_font_size', 18),
                    document_title_color=data.get('document_title_color', '#000000'),
                    section_heading_font_size=data.get('section_heading_font_size', 12),
                    section_heading_color=data.get('section_heading_color', '#000000'),
                    paragraph_line_spacing_multiplier=get_decimal(data, 'paragraph_line_spacing_multiplier', '1.2'),
                    table_border_style=data.get('table_border_style', 'solid'),
                    table_border_thickness_pt=get_decimal(data, 'table_border_thickness_pt', '0.5'),
                    table_alternate_row_color=data.get('table_alternate_row_color', ''),
                    show_total_sum_in_words=get_bool(data, 'show_total_sum_in_words', False),
                    total_sum_in_words_language=data.get('total_sum_in_words_language', 'lv'),
                    default_vat_calculation_method=data.get('default_vat_calculation_method', 'exclusive'),
                    show_vat_breakdown=get_bool(data, 'show_vat_breakdown', True),
                    enable_digital_signature_field=get_bool(data, 'enable_digital_signature_field', False),
                    digital_signature_field_name=data.get('digital_signature_field_name', 'Paraksts'),
                    digital_signature_field_size_mm=get_decimal(data, 'digital_signature_field_size_mm', '40'),
                    digital_signature_field_position=data.get('digital_signature_field_position', 'bottom_center')
                )
            self.ieviest_datus(d)
            self._ceļš_projekts = path
            self._add_to_history(path) # Pievienojam projektu vēsturei
            QMessageBox.information(self, "Ielādēts", "Projekts ielādēts.")
        except Exception as e:
                QMessageBox.critical(self, "Kļūda", f"Neizdevās ielādēt projektu:\n{e}")

    def saglabat_noklusejuma_iestatijumus(self):
        d = self.savākt_datus()
        os.makedirs(SETTINGS_DIR, exist_ok=True)
        settings_path = DEFAULT_SETTINGS_FILE
        out = asdict(d)
        # Convert Decimal fields to string for JSON serialization
        for key, value in out.items():
            if isinstance(value, Decimal):
                out[key] = str(value)
        # Clear sensitive/dynamic data for default settings
        out['pozīcijas'] = []
        out['attēli'] = []
        out['pieņēmējs'] = asdict(d.pieņēmējs)
        out['nodevējs'] = asdict(d.nodevējs)
        out['pdf_user_password'] = ""  # Do not save passwords as default
        out['pdf_owner_password'] = ""
        # Individuālais QR kods
        out['custom_qr_code_data'] = ""
        out['include_custom_qr_code'] = False
        out['custom_qr_code_size_mm'] = str(Decimal("20"))
        out['custom_qr_code_position'] = "bottom_right"
        out['custom_qr_code_pos_x_mm'] = str(Decimal("0"))
        out['custom_qr_code_pos_y_mm'] = str(Decimal("0"))
        out['custom_qr_code_color'] = "#000000"

        # Automātiskais QR kods (akta ID)
        out['include_auto_qr_code'] = False
        out['auto_qr_code_size_mm'] = str(Decimal("20"))
        out['auto_qr_code_position'] = "bottom_left"
        out['auto_qr_code_pos_x_mm'] = str(Decimal("0"))
        out['auto_qr_code_pos_y_mm'] = str(Decimal("0"))
        out['auto_qr_code_color'] = "#000000"

        out['templates_dir'] = d.templates_dir  # JAUNA RINDAS - Saglabājam šablonu direktoriju

        try:
            with open(settings_path, 'w', encoding='utf-8') as f:
                json.dump(out, f, ensure_ascii=False, indent=2)
            QMessageBox.information(self, "Saglabāts", "Pašreizējie iestatījumi saglabāti kā noklusējuma.")
        except Exception as e:
            QMessageBox.critical(self, "Kļūda", f"Neizdevās saglabāt noklusējuma iestatījumus:\n{e}")

    def ieladet_noklusejuma_iestatijumus(self):
            os.makedirs(SETTINGS_DIR, exist_ok=True)
            settings_path = DEFAULT_SETTINGS_FILE

            if os.path.exists(settings_path):
                try:
                    with open(settings_path, 'r', encoding='utf-8') as f:
                        data = json.load(f)

                    # Helper to safely get and convert Decimal values
                    def get_decimal(dict_obj, key, default_val):
                        val = dict_obj.get(key, default_val)
                        return to_decimal(val)

                    # Helper to safely get boolean values
                    def get_bool(dict_obj, key, default_val):
                        val = dict_obj.get(key, default_val)
                        return bool(val)

                    d = AktaDati(
                        akta_nr=data.get('akta_nr', ''), datums=data.get('datums', datetime.now().strftime('%Y-%m-%d')), vieta=data.get('vieta', ''),
                        pasūtījuma_nr=data.get('pasūtījuma_nr', ''),
                        pieņēmējs=Persona(
                            nosaukums=data.get('pieņēmējs', {}).get('nosaukums', ''),
                            reģ_nr=data.get('pieņēmējs', {}).get('reģ_nr', ''),
                            adrese=data.get('pieņēmējs', {}).get('adrese', ''),
                            kontaktpersona=data.get('pieņēmējs', {}).get('kontaktpersona', ''),
                            tālrunis=data.get('pieņēmējs', {}).get('tālrunis', ''),
                            epasts=data.get('pieņēmējs', {}).get('epasts', ''),
                            bankas_konts=data.get('pieņēmējs', {}).get('bankas_konts', ''),
                            juridiskais_statuss=data.get('pieņēmējs', {}).get('juridiskais_statuss', '')
                        ),
                        nodevējs=Persona(
                            nosaukums=data.get('nodevējs', {}).get('nosaukums', ''),
                            reģ_nr=data.get('nodevējs', {}).get('reģ_nr', ''),
                            adrese=data.get('nodevējs', {}).get('adrese', ''),
                            kontaktpersona=data.get('nodevējs', {}).get('kontaktpersona', ''),
                            tālrunis=data.get('nodevējs', {}).get('tālrunis', ''),
                            epasts=data.get('nodevējs', {}).get('epasts', ''),
                            bankas_konts=data.get('nodevējs', {}).get('bankas_konts', ''),
                            juridiskais_statuss=data.get('nodevējs', {}).get('juridiskais_statuss', '')
                        ),
                        pozīcijas=[], # Default settings should not load positions
                        attēli=[], # Default settings should not load images
                        piezīmes=data.get('piezīmes', ''), iekļaut_pvn=get_bool(data, 'iekļaut_pvn', False),
                        pvn_likme=get_decimal(data, 'pvn_likme', '21'),
                        parakstu_rindas=get_bool(data, 'parakstu_rindas', True),
                        logotipa_ceļš=data.get('logotipa_ceļš', ''), fonts_ceļš=data.get('fonts_ceļš', ''),
                        paraksts_pieņēmējs_ceļš=data.get('paraksts_pieņēmējs_ceļš', ''),
                        paraksts_nodevējs_ceļš=data.get('paraksts_nodevējs_ceļš', ''),
                        līguma_nr=data.get('līguma_nr', ''),
                        izpildes_termiņš=data.get('izpildes_termiņš', ''),
                        pieņemšanas_datums=data.get('pieņemšanas_datums', ''),
                        nodošanas_datums=data.get('nodošanas_datums', ''),
                        strīdu_risināšana=data.get('strīdu_risināšana', ''),
                        konfidencialitātes_klauzula=get_bool(data, 'konfidencialitātes_klauzula', False),
                        soda_nauda_procenti=get_decimal(data, 'soda_nauda_procenti', '0.0'),
                        piegādes_nosacījumi=data.get('piegādes_nosacījumi', ''),
                        apdrošināšana=get_bool(data, 'apdrošināšana', False),
                        papildu_nosacījumi=data.get('papildu_nosacījumi', ''),
                        atsauces_dokumenti=data.get('atsauces_dokumenti', ''),
                        akta_statuss=data.get('akta_statuss', 'Melnraksts'),
                        valūta=data.get('valūta', 'EUR'),
                        elektroniskais_paraksts=get_bool(data, 'elektroniskais_paraksts', False),
                        radit_elektronisko_parakstu_tekstu=get_bool(data, 'radit_elektronisko_parakstu_tekstu', False),
                        # JAUNA RINDAS
                        pdf_page_size=data.get('pdf_page_size', 'A4'),
                        pdf_page_orientation=data.get('pdf_page_orientation', 'Portrets'),
                        pdf_margin_left=get_decimal(data, 'pdf_margin_left', '18'),
                        pdf_margin_right=get_decimal(data, 'pdf_margin_right', '18'),
                        pdf_margin_top=get_decimal(data, 'pdf_margin_top', '16'),
                        pdf_margin_bottom=get_decimal(data, 'pdf_margin_bottom', '16'),
                        pdf_font_size_head=data.get('pdf_font_size_head', 14),
                        pdf_font_size_normal=data.get('pdf_font_size_normal', 10),
                        pdf_font_size_small=data.get('pdf_font_size_small', 9),
                        pdf_font_size_table=data.get('pdf_font_size_table', 9),
                        pdf_logo_width_mm=get_decimal(data, 'pdf_logo_width_mm', '35'),
                        pdf_signature_width_mm=get_decimal(data, 'pdf_signature_width_mm', '50'),
                        pdf_signature_height_mm=get_decimal(data, 'pdf_signature_height_mm', '20'),
                        docx_image_width_inches=get_decimal(data, 'docx_image_width_inches', '4'),
                        docx_signature_width_inches=get_decimal(data, 'docx_signature_width_inches', '1.5'),
                        table_col_widths=data.get('table_col_widths', '10,40,18,18,20,20,25,25,25'),
                        auto_generate_akta_nr=get_bool(data, 'auto_generate_akta_nr', False),
                        default_currency=data.get('default_currency', 'EUR'),
                        default_unit=data.get('default_unit', 'gab.'),
                        default_pvn_rate=get_decimal(data, 'default_pvn_rate', '21.0'),
                        poppler_path=data.get('poppler_path', ''),
                        # Load new settings
                        header_text_color=data.get('header_text_color', '#000000'),
                        footer_text_color=data.get('footer_text_color', '#000000'),
                        table_header_bg_color=data.get('table_header_bg_color', '#E0E0E0'),
                        table_grid_color=data.get('table_grid_color', '#CCCCCC'),
                        table_row_spacing=get_decimal(data, 'table_row_spacing', '4'),
                        line_spacing_multiplier=get_decimal(data, 'line_spacing_multiplier', '1.2'),
                        show_page_numbers=get_bool(data, 'show_page_numbers', True),
                        show_generation_timestamp=get_bool(data, 'show_generation_timestamp', True),
                        currency_symbol_position=data.get('currency_symbol_position', 'after'),
                        date_format=data.get('date_format', 'YYYY-MM-DD'),
                        signature_line_length_mm=get_decimal(data, 'signature_line_length_mm', '60'),
                        signature_line_thickness_pt=get_decimal(data, 'signature_line_thickness_pt', '0.5'),
                        add_cover_page=get_bool(data, 'add_cover_page', False),
                        cover_page_title=data.get('cover_page_title', 'Pieņemšanas-Nodošanas Akts'),
                        cover_page_logo_width_mm=get_decimal(data, 'cover_page_logo_width_mm', '80'),
                        # Individuālais QR kods
                        include_custom_qr_code=get_bool(data, 'include_custom_qr_code', False),
                        custom_qr_code_data=data.get('custom_qr_code_data', ''),
                        custom_qr_code_size_mm=get_decimal(data, 'custom_qr_code_size_mm', '20'),
                        custom_qr_code_position=data.get('custom_qr_code_position', 'bottom_right'),
                        custom_qr_code_pos_x_mm=get_decimal(data, 'custom_qr_code_pos_x_mm', '0'),
                        custom_qr_code_pos_y_mm=get_decimal(data, 'custom_qr_code_pos_y_mm', '0'),
                        custom_qr_code_color=data.get('custom_qr_code_color', '#000000'),

                        # Automātiskais QR kods (akta ID)
                        include_auto_qr_code=get_bool(data, 'include_auto_qr_code', False),
                        auto_qr_code_size_mm=get_decimal(data, 'auto_qr_code_size_mm', '20'),
                        auto_qr_code_position=data.get('auto_qr_code_position', 'bottom_left'),
                        auto_qr_code_pos_x_mm=get_decimal(data, 'auto_qr_code_pos_x_mm', '0'),
                        auto_qr_code_pos_y_mm=get_decimal(data, 'auto_qr_code_pos_y_mm', '0'),
                        auto_qr_code_color=data.get('auto_qr_code_color', '#000000'),

                        add_watermark=get_bool(data, 'add_watermark', False),
                        watermark_text=data.get('watermark_text', 'MELNRAKSTS'),
                        watermark_font_size=data.get('watermark_font_size', 72),
                        watermark_color=data.get('watermark_color', '#E0E0E0'),
                        watermark_rotation=data.get('watermark_rotation', 45),
                        enable_pdf_encryption=get_bool(data, 'enable_pdf_encryption', False),
                        pdf_user_password=data.get('pdf_user_password', ''),
                        pdf_owner_password=data.get('pdf_owner_password', ''),
                        allow_printing=get_bool(data, 'allow_printing', True),
                        allow_copying=get_bool(data, 'allow_copying', True),
                        allow_modifying=get_bool(data, 'allow_modifying', False),
                        allow_annotating=get_bool(data, 'allow_annotating', True),
                        default_country=data.get('default_country', 'Latvija'),
                        default_city=data.get('default_city', 'Rīga'),
                        show_contact_details_in_header=get_bool(data, 'show_contact_details_in_header', False),
                        contact_details_header_font_size=data.get('contact_details_header_font_size', 8),
                        item_image_width_mm=get_decimal(data, 'item_image_width_mm', '50'),
                        item_image_caption_font_size=data.get('item_image_caption_font_size', 8),
                        show_item_notes_in_table=get_bool(data, 'show_item_notes_in_table', True),
                        show_item_serial_number_in_table=get_bool(data, 'show_item_serial_number_in_table', True),
                        show_item_warranty_in_table=get_bool(data, 'show_item_warranty_in_table', True),
                        table_cell_padding_mm=get_decimal(data, 'table_cell_padding_mm', '2'),
                        table_header_font_style=data.get('table_header_font_style', 'bold'),
                        table_content_alignment=data.get('table_content_alignment', 'left'),
                        signature_font_size=data.get('signature_font_size', 9),
                        signature_spacing_mm=get_decimal(data, 'signature_spacing_mm', '10'),
                        document_title_font_size=data.get('document_title_font_size', 18),
                        document_title_color=data.get('document_title_color', '#000000'),
                        section_heading_font_size=data.get('section_heading_font_size', 12),
                        section_heading_color=data.get('section_heading_color', '#000000'),
                        paragraph_line_spacing_multiplier=get_decimal(data, 'paragraph_line_spacing_multiplier', '1.2'),
                        table_border_style=data.get('table_border_style', 'solid'),
                        table_border_thickness_pt=get_decimal(data, 'table_border_thickness_pt', '0.5'),
                        table_alternate_row_color=data.get('table_alternate_row_color', ''),
                        show_total_sum_in_words=get_bool(data, 'show_total_sum_in_words', False),
                        total_sum_in_words_language=data.get('total_sum_in_words_language', 'lv'),
                        default_vat_calculation_method=data.get('default_vat_calculation_method', 'exclusive'),
                        show_vat_breakdown=get_bool(data, 'show_vat_breakdown', True),
                        enable_digital_signature_field=get_bool(data, 'enable_digital_signature_field', False),
                        digital_signature_field_name=data.get('digital_signature_field_name', 'Paraksts'),
                        digital_signature_field_size_mm=get_decimal(data, 'digital_signature_field_size_mm', '40'),
                        digital_signature_field_position=data.get('digital_signature_field_position', 'bottom_center'),
                        templates_dir=data.get('templates_dir', os.path.join(APP_DATA_DIR, "AktaGenerators_Templates"))
                        # JAUNA RINDAS
                    )
                    self.ieviest_datus(d)
                    self.data = d
                except Exception as e:
                    QMessageBox.warning(self, "Iestatījumu ielādes kļūda",
                                        f"Neizdevās ielādēt noklusējuma iestatījumus: {e}")

    def closeEvent(self, event):
        # Saglabājam noklusējuma iestatījumus, vēsturi un adrešu grāmatu pirms programmas aizvēršanas
        self.saglabat_noklusejuma_iestatijumus()
        self._save_history()
        self._save_address_book()
        self.text_block_manager._save_text_blocks()  # JAUNA RINDAS
        event.accept()

    def _update_preview(self):
        if not hasattr(self, 'preview_label'):
            return

        d = self.savākt_datus()
        temp_pdf_path = None
        try:
            self.preview_label.setText("Ģenerē priekšskatījumu...")
            QApplication.processEvents()

            temp_pdf_path = ģenerēt_pdf(d, pdf_ceļš=None)

            poppler_path_to_use = d.poppler_path if d.poppler_path and os.path.exists(d.poppler_path) else None

            images = convert_from_path(temp_pdf_path, poppler_path=poppler_path_to_use)

            self.preview_images = []
            for pil_img in images:
                from io import BytesIO
                img_byte_arr = BytesIO()
                pil_img.save(img_byte_arr, format='PNG')
                img_byte_arr.seek(0)
                pixmap = QPixmap()
                pixmap.loadFromData(img_byte_arr.getvalue(), 'PNG')
                self.preview_images.append(pixmap)

            self.current_preview_page = 0
            self._show_current_page()

        except Exception as e:
            self.preview_label.setText(f"Kļūda priekšskatījuma ģenerēšanā: {e}")
            QMessageBox.critical(self, "Priekšskatījuma kļūda", f"Neizdevās ģenerēt priekšskatījumu. Pārliecinieties, ka Poppler ir instalēts un pieejams sistēmas PATH (vai norādīts poppler_path iestatījumos).\nKļūda: {e}")
            self.preview_images = []
            self._show_current_page()
        finally:
            if temp_pdf_path and os.path.exists(temp_pdf_path):
                try:
                    os.remove(temp_pdf_path)
                except Exception as e:
                    print(f"Neizdevās dzēst pagaidu PDF failu: {e}")

    def _show_current_page(self):
            if self.preview_images:
                pixmap = self.preview_images[self.current_preview_page]
                label_size = self.preview_scroll_area.viewport().size()
                scaled_pixmap = pixmap.scaled(label_size * self.zoom_factor, Qt.KeepAspectRatio,
                                              Qt.SmoothTransformation)
                self.preview_label.setPixmap(scaled_pixmap)
                self.page_number_label.setText(f"Lapa {self.current_preview_page + 1}/{len(self.preview_images)}")
                self.prev_page_button.setEnabled(self.current_preview_page > 0)
                self.next_page_button.setEnabled(self.current_preview_page < len(self.preview_images) - 1)
            else:
                self.preview_label.clear()
                self.page_number_label.setText("Lapa 0/0")
                self.prev_page_button.setEnabled(False)
                self.next_page_button.setEnabled(False)

    def _show_prev_page(self):
            if self.current_preview_page > 0:
                self.current_preview_page -= 1
                self._show_current_page()

    def _show_next_page(self):
            if self.current_preview_page < len(self.preview_images) - 1:
                self.current_preview_page += 1
                self._show_current_page()

    def ģenerēt_pdf_dialogs(self):
        d = self.savākt_datus()

        # Izveidojam automātisku faila nosaukumu
        safe_akta_nr = drošs_faila_nosaukums(d.akta_nr) if d.akta_nr else "akts"
        safe_datums = d.datums.replace("-", "") if d.datums else datetime.now().strftime("%Y%m%d")
        base_name = f"Akts_{safe_akta_nr}_{safe_datums}"

        # Izveidojam mapi dokumentam
        doc_folder = os.path.join(DEFAULT_OUTPUT_DIR, base_name)
        counter = 1
        while os.path.exists(doc_folder):
            doc_folder = os.path.join(DEFAULT_OUTPUT_DIR, f"{base_name}_{counter:03d}")
            counter += 1

        os.makedirs(doc_folder, exist_ok=True)

        # Failu ceļi
        pdf_path = os.path.join(doc_folder, f"{os.path.basename(doc_folder)}.pdf")
        json_path = os.path.join(doc_folder, f"{os.path.basename(doc_folder)}.json")

        try:
            # Ģenerējam PDF
            ģenerēt_pdf(d, pdf_path)

            # Saglabājam JSON
            out = asdict(d)
            for key, value in out.items():
                if isinstance(value, Decimal):
                    out[key] = str(value)
            for p in out['pozīcijas']:
                for key, value in p.items():
                    if isinstance(value, Decimal):
                        p[key] = str(value)
            out['pieņēmējs'] = asdict(d.pieņēmējs)
            out['nodevējs'] = asdict(d.nodevējs)

            with open(json_path, 'w', encoding='utf-8') as f:
                json.dump(out, f, ensure_ascii=False, indent=2)

            self._add_to_history(json_path)
            QMessageBox.information(self, "PDF ģenerēts", f"Dokumenti saglabāti mapē: {doc_folder}")
            # Pēc veiksmīgas ģenerēšanas, atjaunojam akta numuru, ja ieslēgta auto-ģenerēšana
            if self.data.auto_generate_akta_nr:
                self._generate_akta_nr()

        except Exception as e:
            QMessageBox.critical(self, "Kļūda", f"Neizdevās ģenerēt PDF:\n{e}")

    def ģenerēt_docx_dialogs(self):
        d = self.savākt_datus()

        # Izveidojam automātisku faila nosaukumu
        safe_akta_nr = drošs_faila_nosaukums(d.akta_nr) if d.akta_nr else "akts"
        safe_datums = d.datums.replace("-", "") if d.datums else datetime.now().strftime("%Y%m%d")
        base_name = f"Akts_{safe_akta_nr}_{safe_datums}"

        # Izveidojam mapi dokumentam
        doc_folder = os.path.join(DEFAULT_OUTPUT_DIR, base_name)
        counter = 1
        while os.path.exists(doc_folder):
            doc_folder = os.path.join(DEFAULT_OUTPUT_DIR, f"{base_name}_{counter:03d}")
            counter += 1

        os.makedirs(doc_folder, exist_ok=True)

        # Failu ceļi
        docx_path = os.path.join(doc_folder, f"{os.path.basename(doc_folder)}.docx")
        json_path = os.path.join(doc_folder, f"{os.path.basename(doc_folder)}.json")

        try:
            # Ģenerējam DOCX
            ģenerēt_docx(d, docx_path)

            # Saglabājam JSON
            out = asdict(d)
            for key, value in out.items():
                if isinstance(value, Decimal):
                    out[key] = str(value)
            for p in out['pozīcijas']:
                for key, value in p.items():
                    if isinstance(value, Decimal):
                        p[key] = str(value)
            out['pieņēmējs'] = asdict(d.pieņēmējs)
            out['nodevējs'] = asdict(d.nodevējs)

            with open(json_path, 'w', encoding='utf-8') as f:
                json.dump(out, f, ensure_ascii=False, indent=2)

            self._add_to_history(json_path)
            QMessageBox.information(self, "DOCX ģenerēts", f"Dokumenti saglabāti mapē: {doc_folder}")
            # Pēc veiksmīgas ģenerēšanas, atjaunojam akta numuru, ja ieslēgta auto-ģenerēšana
            if self.data.auto_generate_akta_nr:
                self._generate_akta_nr()

        except Exception as e:
            QMessageBox.critical(self, "Kļūda", f"Neizdevās ģenerēt DOCX:\n{e}")


if __name__ == "__main__":
            app = QApplication(sys.argv)
            window = AktaLogs()
            window.show()
            sys.exit(app.exec())
