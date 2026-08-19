import sys
import ctypes
from typing import Optional
import os
import hashlib
import secrets
from PySide6.QtWidgets import QMenu, QDialog, QFormLayout, QDialogButtonBox
from PySide6.QtGui import QAction
import csv

# ============================
# Noliktava (jauns): produktu katalogs + atlikumi (lokāli, APP_DATA_DIR)
# ============================
from dataclasses import dataclass as _dataclass
from dataclasses import asdict as _asdict
from dataclasses import field as _field

@_dataclass
class NoliktavasPrece:
    sku: str = ""
    nosaukums: str = ""
    vieniba: str = "gab."
    cena: str = ""          # kā teksts (atļauj komatus), konvertējam ar to_decimal()
    pvn_likme: str = ""     # pēc izvēles
    atlikums: float = 0.0
    piezimes: str = ""
    foto_path: str = ""
    kategorija: str = ""
    apakskategorija: str = ""
    svitrkods: str = ""
    serialais_numurs: str = ""
    partijas_numurs: str = ""
    pavadzimes_numurs: str = ""
    iepirkuma_valuta: str = "EUR"
    noliktavas_nosaukums: str = "Pamatnoliktava"
    atrasanas_vieta: str = ""
    piegadatajs: str = ""
    razotajs: str = ""
    iepirkuma_datums: str = ""
    deriguma_termiņš: str = ""
    minimalais_atlikums: float = 0.0
    statuss: str = "Aktīva"
    supplier_code: str = ""
    supplier_email: str = ""
    supplier_phone: str = ""
    supplier_api_url: str = ""
    supplier_product_url: str = ""
    supplier_lead_time_days: int = 0
    preferred_supplier: bool = False
    hs_kods: str = ""
    izcelsmes_valsts: str = ""
    neto_svars: str = ""
    bruto_svars: str = ""
    dokumentu_mape: str = ""
    inventory_id: str = ""
    dokumenti: list = _field(default_factory=list)
    last_sync_at: str = ""

InventoryItem = NoliktavasPrece

class NoliktavaDB:
    def __init__(self, path: str):
        self.path = path
        self.movements_path = os.path.splitext(path)[0] + '_kustibas.json'
        self.warehouses_path = os.path.splitext(path)[0] + '_noliktavas.json'
        self.items: list[NoliktavasPrece] = []
        self.movements: list[dict] = []
        self.saved_warehouses: list[str] = []
        self.load()
        self._load_movements()
        self._load_saved_warehouses()
        self._ensure_inventory_ids(persist=True)

    def load(self):
        try:
            if os.path.exists(self.path):
                with open(self.path, 'r', encoding='utf-8') as f:
                    raw = json.load(f) or []
                self.items = []
                for r in raw:
                    if not isinstance(r, dict):
                        continue
                    try:
                        self.items.append(NoliktavasPrece(**(r or {})))
                    except Exception:
                        self.items.append(NoliktavasPrece(
                            sku=str(r.get('sku','')),
                            nosaukums=str(r.get('nosaukums','')),
                            vieniba=str(r.get('vieniba','gab.')) or 'gab.',
                            cena=str(r.get('cena','')),
                            pvn_likme=str(r.get('pvn_likme','')),
                            atlikums=float(r.get('atlikums',0.0) or 0.0),
                            piezimes=str(r.get('piezimes','')),
                            foto_path=str(r.get('foto_path','')),
                            kategorija=str(r.get('kategorija','')),
                            svitrkods=str(r.get('svitrkods','')),
                            serialais_numurs=str(r.get('serialais_numurs','')),
                            partijas_numurs=str(r.get('partijas_numurs','')),
                            pavadzimes_numurs=str(r.get('pavadzimes_numurs','')),
                            iepirkuma_valuta=str(r.get('iepirkuma_valuta','EUR') or 'EUR'),
                            noliktavas_nosaukums=str(r.get('noliktavas_nosaukums','Pamatnoliktava') or 'Pamatnoliktava'),
                            atrasanas_vieta=str(r.get('atrasanas_vieta','')),
                            piegadatajs=str(r.get('piegadatajs','')),
                            razotajs=str(r.get('razotajs','')),
                            iepirkuma_datums=str(r.get('iepirkuma_datums','')),
                            deriguma_termiņš=str(r.get('deriguma_termiņš','')),
                            minimalais_atlikums=float(r.get('minimalais_atlikums', 0.0) or 0.0),
                            statuss=str(r.get('statuss','Aktīva') or 'Aktīva'),
                            supplier_code=str(r.get('supplier_code','')),
                            supplier_email=str(r.get('supplier_email','')),
                            supplier_phone=str(r.get('supplier_phone','')),
                            supplier_api_url=str(r.get('supplier_api_url','')),
                            supplier_product_url=str(r.get('supplier_product_url','')),
                            supplier_lead_time_days=int(r.get('supplier_lead_time_days', 0) or 0),
                            preferred_supplier=bool(r.get('preferred_supplier', False)),
                            hs_kods=str(r.get('hs_kods','')),
                            izcelsmes_valsts=str(r.get('izcelsmes_valsts','')),
                            neto_svars=str(r.get('neto_svars','')),
                            bruto_svars=str(r.get('bruto_svars','')),
                            dokumentu_mape=str(r.get('dokumentu_mape','')),
                            inventory_id=str(r.get('inventory_id','')),
                            dokumenti=list(r.get('dokumenti', []) or []),
                            last_sync_at=str(r.get('last_sync_at','')),
                        ))
        except Exception:
            self.items = []

    def _load_movements(self):
        try:
            if os.path.exists(self.movements_path):
                with open(self.movements_path, 'r', encoding='utf-8') as f:
                    self.movements = json.load(f) or []
            else:
                self.movements = []
        except Exception:
            self.movements = []

    def _save_movements(self):
        try:
            os.makedirs(os.path.dirname(self.movements_path), exist_ok=True)
            with open(self.movements_path, 'w', encoding='utf-8') as f:
                json.dump(self.movements or [], f, ensure_ascii=False, indent=2)
        except Exception:
            pass


    def _load_saved_warehouses(self):
        try:
            if os.path.exists(self.warehouses_path):
                with open(self.warehouses_path, 'r', encoding='utf-8') as f:
                    raw = json.load(f) or []
                self.saved_warehouses = [str(x).strip() for x in raw if str(x).strip()]
            else:
                self.saved_warehouses = []
        except Exception:
            self.saved_warehouses = []

    def _save_saved_warehouses(self):
        try:
            os.makedirs(os.path.dirname(self.warehouses_path), exist_ok=True)
            merged = self.warehouse_names(include_item_scan=True)
            with open(self.warehouses_path, 'w', encoding='utf-8') as f:
                json.dump(merged, f, ensure_ascii=False, indent=2)
        except Exception:
            pass

    def add_warehouse(self, name: str):
        return self.add_warehouse_name(name)

    def delete_warehouse(self, name: str):
        return self.delete_warehouse_name(name)

    def add_warehouse_name(self, name: str):
        name = (name or '').strip()
        if not name:
            return False
        if name not in self.saved_warehouses:
            self.saved_warehouses.append(name)
            self.saved_warehouses.sort(key=lambda x: x.lower())
            self._save_saved_warehouses()
        return True

    def delete_warehouse_name(self, name: str):
        name = (name or '').strip()
        if not name:
            return False
        self.saved_warehouses = [x for x in (self.saved_warehouses or []) if (x or '').strip().lower() != name.lower()]
        self._save_saved_warehouses()
        return True

    def save(self):
        try:
            os.makedirs(os.path.dirname(self.path), exist_ok=True)
            with open(self.path, 'w', encoding='utf-8') as f:
                json.dump([_asdict(i) for i in (self.items or [])], f, ensure_ascii=False, indent=2)
            self._save_saved_warehouses()
        except Exception:
            pass

    def _num(self, val, default=0.0):
        try:
            return float(str(val or default).replace(',', '.'))
        except Exception:
            return float(default)

    def _new_inventory_id(self) -> str:
        return f"INV-{secrets.token_hex(8)}"

    def _ensure_inventory_ids(self, persist: bool = False):
        changed = False
        seen = set()
        for it in (self.items or []):
            inv_id = (getattr(it, 'inventory_id', '') or '').strip()
            if not inv_id or inv_id in seen:
                setattr(it, 'inventory_id', self._new_inventory_id())
                changed = True
                inv_id = getattr(it, 'inventory_id', '')
            seen.add(inv_id)
        if changed and persist:
            self.save()
        return changed

    def get_by_inventory_id(self, inventory_id: str):
        inventory_id = (inventory_id or '').strip()
        if not inventory_id:
            return None
        for it in (self.items or []):
            if (getattr(it, 'inventory_id', '') or '').strip() == inventory_id:
                return it
        return None

    def delete_by_inventory_ids(self, inventory_ids):
        ids = {(x or '').strip() for x in (inventory_ids or []) if (x or '').strip()}
        if not ids:
            return 0
        kept = []
        removed = 0
        for it in (self.items or []):
            if (getattr(it, 'inventory_id', '') or '').strip() in ids:
                self.record_movement(it, 'Dzēšana', -self._num(getattr(it, 'atlikums', 0.0)), reference='Prece dzēsta no noliktavas')
                removed += 1
            else:
                kept.append(it)
        self.items = kept
        if removed:
            self.save()
        return removed

    def duplicate_by_inventory_ids(self, inventory_ids):
        ids = [(x or '').strip() for x in (inventory_ids or []) if (x or '').strip()]
        if not ids:
            return 0
        created = 0
        clones = []
        for inv_id in ids:
            src = self.get_by_inventory_id(inv_id)
            if src is None:
                continue
            payload = _asdict(src)
            payload['inventory_id'] = self._new_inventory_id()
            payload['dokumenti'] = list(payload.get('dokumenti', []) or [])
            payload['last_sync_at'] = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
            clones.append(NoliktavasPrece(**payload))
            created += 1
        if clones:
            self.items.extend(clones)
            self.save()
        return created

    def warehouse_names(self, include_item_scan: bool = True) -> list[str]:
        out = []
        for name in (self.saved_warehouses or []):
            name = (name or '').strip()
            if name and name not in out:
                out.append(name)
        if include_item_scan:
            for it in (self.items or []):
                name = (getattr(it, 'noliktavas_nosaukums', '') or '').strip()
                if name and name not in out:
                    out.append(name)
        return sorted(out, key=lambda x: x.lower())

    def upsert(self, item: NoliktavasPrece):
        inv_id = (getattr(item, 'inventory_id', '') or '').strip()
        if not inv_id:
            item.inventory_id = self._new_inventory_id()
            inv_id = item.inventory_id
        for i, it in enumerate(self.items):
            if (getattr(it, 'inventory_id', '') or '').strip() == inv_id:
                old_qty = self._num(getattr(self.items[i], 'atlikums', 0.0))
                new_qty = self._num(getattr(item, 'atlikums', 0.0))
                self.items[i] = item
                self.save()
                if abs(new_qty - old_qty) > 1e-9:
                    self.record_movement(item, 'Korekcija', new_qty - old_qty, reference='Kartītes saglabāšana')
                return
        self.items.append(item)
        self.save()
        self.record_movement(item, 'Jauna prece', self._num(getattr(item, 'atlikums', 0.0)), reference='Kartītes izveide')

    def delete_by_row(self, row: int):
        try:
            it = self.items[row]
            self.record_movement(it, 'Dzēšana', -self._num(getattr(it, 'atlikums', 0.0)), reference='Prece dzēsta no noliktavas')
            self.items.pop(row)
            self.save()
        except Exception:
            pass

    def delete_by_sku(self, sku: str):
        sku = (sku or '').strip()
        if not sku:
            return False
        for idx, it in enumerate(list(self.items or [])):
            if (getattr(it, 'sku', '') or '').strip() == sku:
                self.record_movement(it, 'Dzēšana', -self._num(getattr(it, 'atlikums', 0.0)), reference='Prece dzēsta no noliktavas')
                self.items.pop(idx)
                self.save()
                return True
        return False

    def find(self, q: str, include_zero: bool = True) -> list[NoliktavasPrece]:
        q = (q or '').strip().lower()
        items = list(self.items or [])
        if not include_zero:
            items = [it for it in items if self._num(getattr(it, 'atlikums', 0.0)) > 0]
        if not q:
            return items
        out = []
        for it in items:
            hay = ' '.join([
                str(getattr(it, 'sku', '') or ''),
                str(getattr(it, 'svitrkods', '') or ''),
                str(getattr(it, 'nosaukums', '') or ''),
                str(getattr(it, 'kategorija', '') or ''),
                str(getattr(it, 'apakskategorija', '') or ''),
                str(getattr(it, 'serialais_numurs', '') or ''),
                str(getattr(it, 'partijas_numurs', '') or ''),
                str(getattr(it, 'noliktavas_nosaukums', '') or ''),
                str(getattr(it, 'atrasanas_vieta', '') or ''),
                str(getattr(it, 'pavadzimes_numurs', '') or ''),
                str(getattr(it, 'piegadatajs', '') or ''),
                str(getattr(it, 'razotajs', '') or ''),
                str(getattr(it, 'piezimes', '') or ''),
            ]).lower()
            if q in hay:
                out.append(it)
        return out

    def get_by_sku(self, sku: str):
        sku = (sku or '').strip()
        if not sku:
            return None
        for it in (self.items or []):
            if (it.sku or '').strip() == sku:
                return it
        return None

    def record_movement(self, item: NoliktavasPrece | None, movement_type: str, qty: float, reference: str = '', warehouse: str = '', note: str = ''):
        try:
            it = item or NoliktavasPrece()
            row = {
                'datetime': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
                'movement_type': movement_type,
                'sku': getattr(it, 'sku', '') or '',
                'nosaukums': getattr(it, 'nosaukums', '') or '',
                'qty': float(qty or 0.0),
                'vieniba': getattr(it, 'vieniba', '') or 'gab.',
                'warehouse': warehouse or (getattr(it, 'noliktavas_nosaukums', '') or ''),
                'reference': reference or '',
                'note': note or '',
            }
            self.movements.append(row)
            self.movements = self.movements[-5000:]
            self._save_movements()
        except Exception:
            pass

    def receive_item(self, sku: str, qty: float, reference: str = '', note: str = '') -> tuple[bool, str]:
        sku = (sku or '').strip()
        qty = self._num(qty)
        if not sku or qty <= 0:
            return False, 'Nederīgs SKU vai daudzums.'
        it = self.get_by_sku(sku)
        if it is None:
            return False, 'Prece pēc SKU nav atrasta.'
        it.atlikums = self._num(getattr(it, 'atlikums', 0.0)) + qty
        self.save()
        self.record_movement(it, 'Saņemšana', qty, reference=reference, note=note)
        return True, f'Atlikums palielināts līdz {self._num(it.atlikums):g}'

    def adjust_stock(self, sku: str, new_qty: float, reference: str = '', note: str = '') -> tuple[bool, str]:
        sku = (sku or '').strip()
        new_qty = self._num(new_qty)
        if not sku:
            return False, 'Nederīgs SKU.'
        it = self.get_by_sku(sku)
        if it is None:
            return False, 'Prece pēc SKU nav atrasta.'
        old_qty = self._num(getattr(it, 'atlikums', 0.0))
        diff = new_qty - old_qty
        it.atlikums = new_qty
        self.save()
        self.record_movement(it, 'Inventarizācijas korekcija', diff, reference=reference, note=note)
        return True, f'Atlikums koriģēts līdz {new_qty:g}'

    def issue_item(self, sku: str, qty: float, auto_remove_depleted: bool = True, reference: str = '', note: str = '') -> tuple[bool, str]:
        sku = (sku or '').strip()
        qty = self._num(qty)
        if not sku or qty <= 0:
            return False, 'Nederīgs SKU vai daudzums.'
        for idx, it in enumerate(self.items or []):
            if (it.sku or '').strip() == sku:
                current = self._num(getattr(it, 'atlikums', 0.0))
                new_stock = current - qty
                if auto_remove_depleted and new_stock <= 0:
                    self.record_movement(it, 'Izrakstīšana', -qty, reference=reference, note=note)
                    self.items.pop(idx)
                    self.save()
                    return True, 'Prece izrakstīta un izņemta no aktīvās noliktavas, jo atlikums sasniedza 0.'
                it.atlikums = new_stock if new_stock > 0 else 0.0
                self.items[idx] = it
                self.save()
                self.record_movement(it, 'Izrakstīšana', -qty, reference=reference, note=note)
                return True, f'Atlikums atjaunots: {self._num(it.atlikums):g}'
        return False, 'Prece pēc SKU nav atrasta.'

    def summary(self, include_zero: bool = False) -> dict:
        items = self.find('', include_zero=include_zero)
        total_items = len(items)
        total_qty = 0.0
        low_stock = 0
        total_value = 0.0
        warehouses = set()
        for it in items:
            qty = self._num(getattr(it, 'atlikums', 0.0))
            total_qty += qty
            min_qty = self._num(getattr(it, 'minimalais_atlikums', 0.0))
            if qty <= max(min_qty, 3.0):
                low_stock += 1
            price = self._num(getattr(it, 'cena', '0'))
            total_value += qty * price
            wh = (getattr(it, 'noliktavas_nosaukums', '') or '').strip()
            if wh:
                warehouses.add(wh)
        return {
            'total_items': total_items,
            'total_qty': total_qty,
            'low_stock': low_stock,
            'total_value': total_value,
            'warehouse_count': len(warehouses),
        }


# ============================
# FIX: path-like safety (dict -> str) to allow repeated PDF generation
# ============================
# Dažkārt (pēc pirmās ģenerēšanas) UI stāvoklī ceļi var nonākt kā dict (piem. Qt.UserRole).
# Tas izraisa TypeError "expected str, bytes or os.PathLike object, not dict" pie os.path / open() / ReportLab.
# Šis ir minimāls "drop-in" ielāps: padara os.fspath tolerantāku pret dict/tuple/list,
# lai visur, kur Python/stdlib iekšēji izsauc os.fspath(), mēs vienmēr saņemtu string ceļu.

_ORIG_OS_FSPATH = os.fspath

def _safe_fspath(p):
    try:
        if p is None:
            return ""
        # QFileDialog dažreiz atgriež (path, filter)
        if isinstance(p, (list, tuple)) and p:
            return _safe_fspath(p[0])
        # UI/state var iedot dict ar ceļu
        if isinstance(p, dict):
            for k in ("path", "ceļš", "celsh", "file", "filepath", "filename", "value", "pdf", "json"):
                v = p.get(k)
                if v:
                    return _safe_fspath(v)
            return ""
        return _ORIG_OS_FSPATH(p)
    except TypeError:
        # Pēdējais glābiņš: pārvēršam par str (labāk FileNotFoundError nekā TypeError)
        try:
            return str(p)
        except Exception:
            return ""

# Monkey-patch: ietekmē os.path.*, open(), reportlab utt.
os.fspath = _safe_fspath


def resource_path(relative_path: str) -> str:
    """Atgriež pareizu ceļu uz resursu gan Python režīmā, gan PyInstaller EXE režīmā."""
    base_path = getattr(sys, "_MEIPASS", os.path.abspath("."))
    return os.path.join(base_path, relative_path)

def set_windows_app_id(app_id: str) -> None:
    """Uzliek Windows AppUserModelID, lai Taskbar/Alt+Tab izmantotu pareizo ikonu."""
    try:
        ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(app_id)
    except Exception:
        pass

import os.path
import io
import re

import json
import base64
import tempfile
from datetime import datetime, timedelta
from decimal import Decimal, InvalidOperation
from PySide6.QtGui import QColor, QPageSize
from dataclasses import dataclass, asdict, field
import shutil
import requests
import subprocess # Jauns imports

# Pārliecināties, ka šīs ir importētas no PySide6.QtWidgets
from PySide6.QtWidgets import (
    QApplication, QWidget, QMainWindow, QLabel, QLineEdit, QTextEdit, QPushButton,
    QFileDialog, QSpinBox, QDoubleSpinBox, QTableWidget, QTableWidgetItem, QHeaderView,
    QToolButton, QTabWidget, QFormLayout, QVBoxLayout, QHBoxLayout, QMessageBox, QCheckBox,
    QListWidget, QListWidgetItem, QGroupBox, QComboBox, QInputDialog, QSplitter, QScrollArea, QDateEdit, QAbstractItemView, QMenu, QGridLayout, QFrame, QSizePolicy, QTextBrowser
)

from PIL import Image
from PIL.ImageQt import ImageQt # JAUNS IMPORTS
from PySide6.QtGui import QPainter # JAUNS IMPORTS

from pdf2image import convert_from_path
from PySide6.QtGui import QPixmap

from PySide6.QtCore import Qt, QSize, QSettings, QStandardPaths, QUrl, QPoint, QTimer, QThread, Signal, QDate, QEvent, QDate
from PySide6.QtGui import QAction, QIcon, QDesktopServices
from PySide6.QtPrintSupport import QPrinter, QPrintDialog, QPrintPreviewDialog # JAUNS IMPORTS



# Import for WebEngine
from PySide6.QtWebEngineWidgets import QWebEngineView
from PySide6.QtWebEngineCore import QWebEngineUrlRequestInterceptor # For intercepting URL changes

# ReportLab imports (unchanged)
from reportlab.lib.pagesizes import A4, landscape, portrait, letter, legal, A3, A5
from reportlab.graphics import renderPDF
from reportlab.graphics.shapes import Drawing
from reportlab.graphics.barcode import qr as rl_qr
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
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import RGBColor
from PySide6.QtWebChannel import QWebChannel
from PySide6.QtCore import QObject, Slot
from PySide6.QtCore import Qt, QSize, QSettings, QStandardPaths, QUrl, QPoint, QTimer, QThread, Signal
import copy
from dataclasses import asdict
import platform

# ---------------------- Konstantes un direktoriju iestatījumi ----------------------
APP_DATA_DIR = QStandardPaths.writableLocation(QStandardPaths.AppDataLocation)
DOCUMENTS_DIR = QStandardPaths.writableLocation(QStandardPaths.DocumentsLocation)

# ==============================
# Iestatījumi (vienkārši JSON)
# ==============================

def _settings_path() -> str:
    return os.path.join(APP_DATA_DIR, "settings.json")

def load_settings() -> dict:
    try:
        p = _settings_path()
        if os.path.exists(p):
            with open(p, "r", encoding="utf-8") as f:
                return json.load(f) or {}
    except Exception:
        pass
    return {}

def save_settings(data: dict):
    try:
        os.makedirs(APP_DATA_DIR, exist_ok=True)
        with open(_settings_path(), "w", encoding="utf-8") as f:
            json.dump(data or {}, f, ensure_ascii=False, indent=2)
    except Exception:
        pass



SETTINGS_DIR = os.path.join(APP_DATA_DIR, "AktaGenerators")
HISTORY_FILE = os.path.join(SETTINGS_DIR, "history.json")
ADDRESS_BOOK_FILE = os.path.join(SETTINGS_DIR, "address_book.json")
DEFAULT_SETTINGS_FILE = os.path.join(SETTINGS_DIR, "default_settings.json")
AKTA_NR_COUNTER_FILE = os.path.join(SETTINGS_DIR, "akta_nr_counter.json")  # JAUNS: stabilai secīgai akta numuru ģenerācijai
TEXT_BLOCKS_FILE = os.path.join(SETTINGS_DIR, "text_blocks.json") # JAUNA RINDAS

# Jaunas noklusējuma saglabāšanas mapes
DEFAULT_OUTPUT_DIR = os.path.join(DOCUMENTS_DIR, "AktaGenerators_Output")
PROJECT_SAVE_DIR = os.path.join(APP_DATA_DIR, "AktaGenerators_Projects")
# TEMPLATES_DIR = os.path.join(APP_DATA_DIR, "AktaGenerators_Templates") # JAUNA RINDAS - Tagad dinamiski iestatīts AktaDati objektā

# ==============================
# Audit logs + Undo/Redo (GLOBAL)
# ==============================


# ============================
# Produkta izvēle no Noliktavas (jauns, profesionāls "meklētājs" ar daudzumiem)
# ============================
class ProductPickerDialog(QDialog):
    def __init__(self, parent, noliktava_db: NoliktavaDB):
        super().__init__(parent)
        self.setWindowTitle("Pievienot preces no noliktavas")
        self.resize(860, 560)
        self.db = noliktava_db
        self._result = []

        v = QVBoxLayout(self)
        top = QHBoxLayout()
        self.in_search = QLineEdit()
        self.in_search.setPlaceholderText("Meklēt pēc SKU / nosaukuma / piezīmēm…")
        self.in_search.textChanged.connect(self._refresh)
        top.addWidget(self.in_search, 1)
        btn_clear = QPushButton("Notīrīt")
        btn_clear.clicked.connect(lambda: self.in_search.setText(""))
        top.addWidget(btn_clear, 0)
        v.addLayout(top)

        self.tbl = QTableWidget()
        self.tbl.setColumnCount(7)
        self.tbl.setHorizontalHeaderLabels(["✓", "SKU", "Nosaukums", "Vienība", "Cena", "Atlikums", "Daudzums"])
        self.tbl.horizontalHeader().setStretchLastSection(True)
        try:
            self.tbl.verticalHeader().setVisible(False)
            self.tbl.setSelectionBehavior(QTableWidget.SelectRows)
        except Exception:
            pass
        v.addWidget(self.tbl, 1)

        bottom = QHBoxLayout()
        bottom.addWidget(QLabel("Atzīmē preces, norādi daudzumus, tad spied “Pievienot izvēlētos”."), 1)
        btn_add = QPushButton("Pievienot izvēlētos")
        btn_add.clicked.connect(self._accept_selected)
        bottom.addWidget(btn_add, 0)
        btn_cancel = QPushButton("Atcelt")
        btn_cancel.clicked.connect(self.reject)
        bottom.addWidget(btn_cancel, 0)
        v.addLayout(bottom)

        self._refresh()

    def _refresh(self):
        items = self.db.find(self.in_search.text() if self.db else "")
        self.tbl.setRowCount(len(items))
        for r, it in enumerate(items):
            chk = QCheckBox()
            chk.setChecked(False)
            self.tbl.setCellWidget(r, 0, chk)
            self.tbl.setItem(r, 1, QTableWidgetItem(it.sku or ""))
            self.tbl.setItem(r, 2, QTableWidgetItem(it.nosaukums or ""))
            self.tbl.setItem(r, 3, QTableWidgetItem(it.vieniba or ""))
            self.tbl.setItem(r, 4, QTableWidgetItem(str(it.cena or "")))
            self.tbl.setItem(r, 5, QTableWidgetItem(str(it.atlikums if it.atlikums is not None else "")))
            qty = QDoubleSpinBox()
            qty.setDecimals(3); qty.setMinimum(0.0); qty.setMaximum(1e9); qty.setValue(1.0)
            self.tbl.setCellWidget(r, 6, qty)
        try:
            self.tbl.resizeColumnsToContents()
            self.tbl.setColumnWidth(0, 42)
            self.tbl.setColumnWidth(1, 140)
            self.tbl.setColumnWidth(3, 90)
            self.tbl.setColumnWidth(4, 90)
            self.tbl.setColumnWidth(5, 90)
            self.tbl.setColumnWidth(6, 110)
        except Exception:
            pass
        self._items_view = items

    def _accept_selected(self):
        out = []
        for r in range(self.tbl.rowCount()):
            chk = self.tbl.cellWidget(r, 0)
            qtyw = self.tbl.cellWidget(r, 6)
            if isinstance(chk, QCheckBox) and chk.isChecked():
                q = float(qtyw.value()) if isinstance(qtyw, QDoubleSpinBox) else 0.0
                if q > 0:
                    out.append((self._items_view[r], q))
        if not out:
            QMessageBox.information(self, "Nav izvēlēts", "Lūdzu atzīmē vismaz vienu preci un norādi daudzumu.")
            return
        self._result = out
        self.accept()

    def get_selection(self):
        return list(self._result or [])


class InventoryTransferOptionsDialog(QDialog):
    def __init__(self, parent, item: NoliktavasPrece, qty: float = 1.0, field_options=None, position_headers=None):
        super().__init__(parent)
        self.setWindowTitle("Pārnese uz Pozīcijām")
        self.resize(920, 700)
        self._item = item
        self._field_options = list(field_options or [])
        self._position_headers = list(position_headers or [])

        v = QVBoxLayout(self)
        info = QLabel(f"Prece: <b>{item.nosaukums or item.sku or ''}</b> &nbsp;&nbsp; SKU: {item.sku or '-'}")
        info.setWordWrap(True)
        v.addWidget(info)

        qty_row = QHBoxLayout()
        qty_row.addWidget(QLabel("Daudzums pozīcijā:"))
        self.sp_qty = QDoubleSpinBox()
        self.sp_qty.setDecimals(3)
        self.sp_qty.setRange(0.001, 1e9)
        self.sp_qty.setValue(float(qty or 1.0))
        qty_row.addWidget(self.sp_qty)
        qty_row.addStretch(1)
        v.addLayout(qty_row)

        v.addWidget(QLabel("Izvēlies, kurās Pozīciju tabulas kolonnās ievietot Noliktavas laukus:"))
        self.tbl = QTableWidget(len(self._field_options), 4)
        self.tbl.setHorizontalHeaderLabels(["Iekļaut", "Noliktavas lauks", "Vērtība", "Pozīciju kolonna"])
        self.tbl.verticalHeader().setVisible(False)
        self.tbl.setAlternatingRowColors(True)
        targets = [
            ("ignore", "Nepārsūtīt"),
            ("apraksts", "Apraksts"),
            ("notes", "Piezīmes pozīcijai"),
            ("serial", "Seriālais Nr."),
            ("warranty", "Garantija"),
            ("foto", "Foto"),
        ]
        for col_idx, header in self._position_headers:
            header = (header or '').strip()
            if not header:
                continue
            targets.append((f'col:{col_idx}', f'Pozīciju kolonna: {header}'))
        default_target = {
            'nosaukums': 'apraksts', 'sku': 'notes', 'svitrkods': 'notes', 'kategorija': 'notes', 'apakskategorija': 'notes',
            'noliktavas_nosaukums': 'notes', 'atrasanas_vieta': 'notes', 'piegadatajs': 'notes', 'razotajs': 'notes',
            'partijas_numurs': 'notes', 'serialais_numurs': 'serial', 'pavadzimes_numurs': 'notes', 'iepirkuma_datums': 'notes',
            'deriguma_termiņš': 'warranty', 'statuss': 'notes', 'cena': 'ignore', 'pvn_likme': 'notes', 'vieniba': 'ignore',
            'foto_path': 'foto', 'piezimes': 'notes', 'supplier_code': 'notes', 'supplier_email': 'notes', 'supplier_phone': 'notes',
            'hs_kods': 'notes', 'izcelsmes_valsts': 'notes'
        }
        for r, (key, label) in enumerate(self._field_options):
            chk = QCheckBox()
            chk.setChecked(key in ('nosaukums', 'serialais_numurs', 'foto_path', 'piezimes', 'sku'))
            self.tbl.setCellWidget(r, 0, chk)
            self.tbl.setItem(r, 1, QTableWidgetItem(label))
            val = getattr(item, key, '') if hasattr(item, key) else ''
            self.tbl.setItem(r, 2, QTableWidgetItem(str(val or '')))
            cb = QComboBox()
            for target_key, target_label in targets:
                cb.addItem(target_label, target_key)
            ix = cb.findData(default_target.get(key, 'notes'))
            cb.setCurrentIndex(ix if ix >= 0 else 0)
            self.tbl.setCellWidget(r, 3, cb)

        self.tbl.horizontalHeader().setStretchLastSection(True)
        self.tbl.resizeColumnsToContents()
        v.addWidget(self.tbl, 1)

        v.addWidget(QLabel("Dokumenti, kurus automātiski pievienot kā pielikumus šai precei:"))
        self.list_docs = QListWidget()
        docs = list(getattr(item, 'dokumenti', []) or [])
        for doc in docs:
            if not isinstance(doc, dict):
                continue
            path = str(doc.get('path') or doc.get('ceļš') or '').strip()
            if not path:
                continue
            doc_type = str(doc.get('type') or doc.get('tips') or 'Dokuments').strip() or 'Dokuments'
            title = str(doc.get('name') or doc.get('nosaukums') or os.path.basename(path)).strip() or os.path.basename(path)
            itemw = QListWidgetItem(f"[{doc_type}] {title}")
            itemw.setData(Qt.UserRole, {'path': path, 'type': doc_type, 'name': title})
            itemw.setFlags(itemw.flags() | Qt.ItemIsUserCheckable)
            itemw.setCheckState(Qt.Checked)
            self.list_docs.addItem(itemw)
        v.addWidget(self.list_docs, 1)

        docs_btns = QHBoxLayout()
        btn_all = QPushButton("Visi dokumenti")
        btn_none = QPushButton("Nevienu")
        btn_inv = QPushButton("Invertēt")
        for btn, state in ((btn_all, Qt.Checked), (btn_none, Qt.Unchecked)):
            btn.clicked.connect(lambda _=False, st=state: [self.list_docs.item(i).setCheckState(st) for i in range(self.list_docs.count())])
        btn_inv.clicked.connect(lambda: [self.list_docs.item(i).setCheckState(Qt.Unchecked if self.list_docs.item(i).checkState() == Qt.Checked else Qt.Checked) for i in range(self.list_docs.count())])
        docs_btns.addWidget(btn_all)
        docs_btns.addWidget(btn_none)
        docs_btns.addWidget(btn_inv)
        docs_btns.addStretch(1)
        v.addLayout(docs_btns)

        bb = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        bb.accepted.connect(self.accept)
        bb.rejected.connect(self.reject)
        v.addWidget(bb)

    def get_result(self):
        mapping = {}
        selected_fields = []
        for r, (key, label) in enumerate(self._field_options):
            chk = self.tbl.cellWidget(r, 0)
            cb = self.tbl.cellWidget(r, 3)
            if isinstance(chk, QCheckBox) and chk.isChecked():
                selected_fields.append(key)
                if isinstance(cb, QComboBox):
                    mapping[key] = cb.currentData() or 'ignore'
        docs = []
        for i in range(self.list_docs.count()):
            it = self.list_docs.item(i)
            if it.checkState() == Qt.Checked:
                docs.append(dict(it.data(Qt.UserRole) or {}))
        return {
            'qty': float(self.sp_qty.value()),
            'selected_fields': selected_fields,
            'column_mapping': mapping,
            'documents': docs,
        }


class AuditLogger:
    """Vienkāršs, ātrs audit žurnāls (JSON Lines).
    Katrs ieraksts ir viena JSON rinda, lai failu var lasīt/filtrēt arī ārpus programmas.
    """

    def __init__(self, log_path: str):
        self.log_path = _coerce_path(log_path)
        os.makedirs(os.path.dirname(self.log_path), exist_ok=True)
        self._in_memory = []  # pēdējie N ieraksti UI

    def write(self, event: str, details: dict | None = None, user: str = ""):
        try:
            ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            row = {
                "ts": ts,
                "user": user or "",
                "event": event or "",
                "details": details or {},
            }
            with open(self.log_path, "a", encoding="utf-8") as f:
                f.write(json.dumps(row, ensure_ascii=False) + "\n")
            self._in_memory.append(row)
            # limit memory
            if len(self._in_memory) > 500:
                self._in_memory = self._in_memory[-500:]
        except Exception:
            # audit nedrīkst nogāzt app
            pass

    def tail(self, n: int = 200):
        try:
            if self._in_memory:
                return self._in_memory[-n:]
            # Ja nav atmiņā, mēģinām nolasīt faila beigas (vienkārši)
            if not self.log_path or not os.path.exists(self.log_path):
                return []
            rows = []
            with open(self.log_path, "r", encoding="utf-8") as f:
                for line in f:
                    line = line.strip()
                    if not line:
                        continue
                    try:
                        rows.append(json.loads(line))
                    except Exception:
                        continue
            return rows[-n:]
        except Exception:
            return []


class UndoRedoManager:
    """Vienkāršs Undo/Redo ar 'snapshots' (adresēs grāmata + projekts).
    Nav QUndoStack, bet ir stabils, viegli uzturams un pietiekams 99% gadījumu.
    """
    def __init__(self, max_steps: int = 50):
        self.max_steps = max_steps
        self._undo = []
        self._redo = []

    def clear_redo(self):
        self._redo = []

    def push_undo(self, state: dict):
        self._undo.append(state)
        if len(self._undo) > self.max_steps:
            self._undo = self._undo[-self.max_steps:]
        self.clear_redo()

    def can_undo(self) -> bool:
        return len(self._undo) > 0

    def can_redo(self) -> bool:
        return len(self._redo) > 0

    def pop_undo(self) -> dict | None:
        if not self._undo:
            return None
        return self._undo.pop()

    def push_redo(self, state: dict):
        self._redo.append(state)
        if len(self._redo) > self.max_steps:
            self._redo = self._redo[-self.max_steps:]

    def pop_redo(self) -> dict | None:
        if not self._redo:
            return None
        return self._redo.pop()




# Pārliecināmies, ka direktoriji eksistē
os.makedirs(SETTINGS_DIR, exist_ok=True)
os.makedirs(DEFAULT_OUTPUT_DIR, exist_ok=True)
os.makedirs(PROJECT_SAVE_DIR, exist_ok=True)
# os.makedirs(TEMPLATES_DIR, exist_ok=True) # JAUNA RINDAS - Tagad dinamiski iestatīts AktaDati objektā




# ---------------------- UI tēma (modernāks izskats) ----------------------
def apply_modern_theme(app: QApplication, dark: bool = True):
    """Iestata modernu Fusion stilu + noapaļotus elementus (bez ārējām atkarībām)."""
    try:
        app.setStyle("Fusion")
    except Exception:
        pass

    # High-DPI (īpaši Windows)
    try:
        # High-DPI (PySide6 jaunākās versijās daļa atribūtu ir deprecated; atstājam droši)
        if hasattr(Qt, 'AA_EnableHighDpiScaling'):
            try:
                QApplication.setAttribute(Qt.AA_EnableHighDpiScaling, True)
            except Exception:
                pass
        if hasattr(Qt, 'AA_UseHighDpiPixmaps'):
            try:
                QApplication.setAttribute(Qt.AA_UseHighDpiPixmaps, True)
            except Exception:
                pass
    except Exception:
        pass

    # Palete (tumšā pēc noklusējuma)
    pal = app.palette()
    if dark:
        from PySide6.QtGui import QPalette, QColor
        pal.setColor(QPalette.Window, QColor(20, 22, 28))
        pal.setColor(QPalette.WindowText, QColor(230, 230, 230))
        pal.setColor(QPalette.Base, QColor(28, 30, 38))
        pal.setColor(QPalette.AlternateBase, QColor(34, 36, 46))
        pal.setColor(QPalette.ToolTipBase, QColor(255, 255, 255))
        pal.setColor(QPalette.ToolTipText, QColor(20, 22, 28))
        pal.setColor(QPalette.Text, QColor(230, 230, 230))
        pal.setColor(QPalette.Button, QColor(34, 36, 46))
        pal.setColor(QPalette.ButtonText, QColor(230, 230, 230))
        pal.setColor(QPalette.BrightText, QColor(255, 0, 0))
        pal.setColor(QPalette.Link, QColor(92, 170, 255))
        pal.setColor(QPalette.Highlight, QColor(92, 170, 255))
        pal.setColor(QPalette.HighlightedText, QColor(10, 10, 10))
        app.setPalette(pal)

    # Viegls “modern” QSS
    app.setStyleSheet("""
        QMainWindow { background: transparent; }
        QWidget { font-size: 10.5pt; }
        QLineEdit, QTextEdit, QPlainTextEdit, QSpinBox, QDoubleSpinBox, QDateEdit, QComboBox {
            padding: 6px 8px;
            border: 1px solid rgba(255,255,255,0.12);
            border-radius: 10px;
        }
        QLineEdit:focus, QTextEdit:focus, QPlainTextEdit:focus, QSpinBox:focus, QDoubleSpinBox:focus, QDateEdit:focus, QComboBox:focus {
            border: 1px solid rgba(92,170,255,0.9);
        }
        QPushButton, QToolButton {
            padding: 7px 10px;
            border-radius: 10px;
            border: 1px solid rgba(255,255,255,0.14);
        }
        QPushButton:hover, QToolButton:hover { border: 1px solid rgba(92,170,255,0.7); }
        QPushButton:pressed, QToolButton:pressed { padding-top: 8px; padding-bottom: 6px; }
        QTabBar::tab {
            padding: 8px 12px;
            margin: 2px;
            border-radius: 10px;
            background: #0f1520;
            color: #d7deea;
            border: 1px solid rgba(255,255,255,0.08);
        }
        QTabBar::tab:selected {
            background: #cfd4db;
            color: #111827;
            border: 1px solid rgba(255,255,255,0.18);
            font-weight: 700;
        }
        QTabBar::tab:hover:!selected {
            background: #1a2231;
        }
        QTabWidget::pane { border: 0px; top: 2px; }
        QGroupBox {
            border: 1px solid rgba(255,255,255,0.10);
            border-radius: 14px;
            margin-top: 10px;
        }
        QGroupBox::title { subcontrol-origin: margin; left: 10px; padding: 0 6px; }
        QHeaderView::section {
            padding: 6px 8px;
            border: 0px;
            border-right: 1px solid rgba(255,255,255,0.10);
        }
        QTableWidget { border-radius: 14px; border: 1px solid rgba(255,255,255,0.10); }
        QScrollArea { border: 0px; }
        QMessageBox { font-size: 10.5pt; }
    """)

# ---------------------- Datu modeļi ----------------------
# (unchanged)

@dataclass
class Persona:
    nosaukums: str = ""
    reģ_nr: str = ""
    adrese: str = ""
    kontaktpersona: str = ""
    amats: str = ""
    pilnvaras_pamats: str = ""
    tālrunis: str = ""
    epasts: str = ""
    web_lapa: str = ""
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
class AtsaucesDokuments:
    """Reāls atsauces dokuments, ko pievieno PDF beigās kā atvasinājumu."""
    ceļš: str
    nosaukums: str = ""
    scale_pct: int = 100


@dataclass
class AktaDati:
    akta_nr: str = ""
    datums: str = ""  # YYYY-MM-DD
    vieta: str = ""
    pasūtījuma_nr: str = ""
    # --- Dokumenta tips / nosaukums (jauns) ---
    # doc_tips: "akta" | "pavadzime" | "rekins" (izmanto UI + PDF virsrakstam)
    doc_tips: str = "akta"
    dokumenta_nosaukums: str = "Pieņemšanas–Nodošanas akts"
    # Rēķinam / pavadzīmei (pēc izvēles)
    apmaksas_termins: str = ""   # YYYY-MM-DD
    piegades_datums: str = ""    # YYYY-MM-DD
    party_mode: str = "puses"
    rekvizitu_virsraksts: str = "Rekvizīti"
    pieņēmēja_loma: str = "Pieņēmējs"
    nodevēja_loma: str = "Iekārtas/u /pakalpojuma/u nodevējs"

    pieņēmējs: Persona = field(default_factory=Persona)
    rekviziti: Persona = field(default_factory=Persona)
    nodevējs: Persona = field(default_factory=Persona)
    pozīcijas: list = field(default_factory=list)  # list[Pozīcija]
    attēli: list = field(default_factory=list)
    piezīmes: str = ""
    iekļaut_pvn: bool = False
    pvn_likme: Decimal = Decimal("21.0")
    parakstu_rindas: bool = True
    paraksta_rezims: str = "physical"
    paraksta_nav_teksts: str = "Dokuments ir spēkā bez paraksta, ja tas sagatavots un apstiprināts atbilstoši pušu noteiktajai kārtībai."
    paraksta_vards_rekviziti: str = ""
    paraksta_vards_pienemejs: str = ""
    paraksta_vards_nodevejs: str = ""
    # Papildu parakstu rindas zem galvenajām parakstu rindām.
    # Katrs ieraksts: {"nosaukums": "Saskaņoja", "parakstitajs": "Vārds Uzvārds"}
    papildu_parakstu_rindas: list = field(default_factory=list)
    logotipa_ceļš: str = ""
    fonts_ceļš: str = ""  # TTF/OTF
    paraksts_pieņēmējs_ceļš: str = ""
    paraksts_nodevējs_ceļš: str = ""
    līguma_nr: str = ""
    izpildes_termiņš: str = ""
    pieņemšanas_datums: str = ""
    nodošanas_datums: str = ""
    ieklaut_izpildes_terminu: bool = True
    ieklaut_pienemsanas_datumu: bool = True
    ieklaut_nodosanas_datumu: bool = True
    strīdu_risināšana: str = ""
    konfidencialitātes_klauzula: bool = False
    soda_nauda_procenti: Decimal = Decimal("0.0")
    piegādes_nosacījumi: str = ""
    apdrošināšana: bool = False
    apdrošināšana_teksts: str = ""
    papildu_nosacījumi: str = ""
    atsauces_dokumenti: str = ""
    atsauces_dokumenti_faili: list = field(default_factory=list)  # list[AtsaucesDokuments]
    akta_statuss: str = "Melnraksts"
    valūta: str = "EUR"
    elektroniskais_paraksts: bool = False
    radit_elektronisko_parakstu_tekstu: bool = False # JAUNS LAUKS
    qr_kods_enabled: bool = True  # QR kods apakšā pa kreisi
    qr_kods_ieklaut_pozicijas: bool = True
    qr_only_first_page: bool = False
    qr_verification_url_enabled: bool = False
    qr_verification_base_url: str = ""
    qr_kods_izmers_mm: Decimal = Decimal("25.0")

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
    auto_generate_akta_nr: bool = True
    default_execution_days: int = 5  # Noklusējuma dienu skaits izpildes termiņam (datums + N)
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
    show_item_photo_in_table: bool = True
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
    docx_template_path: str = ""  # (ja vēlies) DOCX šablons ar {{placeholders}}
    custom_columns: list = field(
    default_factory=list)  # Saraksts ar pielāgotajām kolonnām: [{'name': 'Kolonna1', 'data': []}]

    # --- JAUNS: Pozīciju tabulas kolonnu pārvaldība (UI + PDF) ---
    # poz_columns_config struktūra:
    # {
    #   "nr": {"title": "Nr.", "visible": True},
    #   "apraksts": {"title": "Apraksts", "visible": True},
    #   "daudzums": {"title": "Daudzums", "visible": True},
    #   "vieniba": {"title": "Vienība", "visible": True},
    #   "cena": {"title": "Cena", "visible": True},
    #   "summa": {"title": "Summa", "visible": True},
    #   "serial": {"title": "Seriālais Nr.", "visible": True},
    #   "warranty": {"title": "Garantija", "visible": True},
    #   "notes": {"title": "Piezīmes pozīcijai", "visible": True},
    #   "foto": {"title": "Foto", "visible": True}
    # }
    poz_columns_config: dict = field(default_factory=dict)

    # Vai rādīt cenu apkopojumu (kopsavilkumu) zem pozīciju tabulas (PDF)
    show_price_summary: bool = True
    # Saglabā lietotāja kolonnu secību Pozīciju tabulai (UI -> PDF)
    poz_columns_visual_order: list = field(default_factory=list)
    # Saglabā kolonnu platumus/izkārtojumu (QHeaderView.saveState) Pozīciju tabulai (UI)
    poz_header_state_b64: str = ""
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


def _default_poz_columns_config() -> dict:
    """Noklusētā pozīciju kolonnu konfigurācija (UI + PDF)."""
    return {
        "nr": {"title": "Nr.", "visible": True},
        "apraksts": {"title": "Apraksts", "visible": True},
        "daudzums": {"title": "Daudzums", "visible": True},
        "vieniba": {"title": "Vienība", "visible": True},
        "cena": {"title": "Cena", "visible": True},
        "summa": {"title": "Summa", "visible": True},
        "serial": {"title": "Seriālais Nr.", "visible": True},
        "warranty": {"title": "Garantija", "visible": True},
        "notes": {"title": "Piezīmes pozīcijai", "visible": True},
        "foto": {"title": "Foto", "visible": True},
    }


def _merge_poz_columns_config(cfg: Optional[dict]) -> dict:
    """Droši apvieno lietotāja konfigurāciju ar noklusējuma vērtībām."""
    base = _default_poz_columns_config()
    out = {}
    cfg = cfg if isinstance(cfg, dict) else {}
    for k, v in base.items():
        cur = cfg.get(k, {}) if isinstance(cfg.get(k, {}), dict) else {}
        out[k] = {
            "title": str(cur.get("title", v.get("title", ""))),
            "visible": bool(cur.get("visible", v.get("visible", True))),
        }
    # Saglabājam arī nezināmus atslēgu ierakstus (ja nākotnē pievieno jaunas kolonnas)
    for k, cur in cfg.items():
        if k in out:
            continue
        if isinstance(cur, dict):
            out[k] = {"title": str(cur.get("title", k)), "visible": bool(cur.get("visible", True))}
    return out


DOCUMENT_TYPE_PRESETS = {
    "akta": {"title": "Pieņemšanas–Nodošanas akts", "party_mode": "puses"},
    "pavadzime": {"title": "Pavadzīme", "party_mode": "puses"},
    "rekins": {"title": "Rēķins", "party_mode": "puses"},
    "norakstisanas_akts": {"title": "Norakstīšanas akts", "party_mode": "rekviziti"},
    "inventarizacijas_akts": {"title": "Inventarizācijas akts", "party_mode": "rekviziti"},
    "nodosanas_akts": {"title": "Nodošanas akts", "party_mode": "puses"},
    "atgriesanas_akts": {"title": "Atgriešanas akts", "party_mode": "puses"},
    "defektacijas_akts": {"title": "Defektācijas akts", "party_mode": "rekviziti"},
    "cits": {"title": "Cits dokuments", "party_mode": "puses"},
}

PARTY_MODE_LABELS = {
    "puses": "Puses",
    "rekviziti": "Rekvizīti",
    "abi": "Puses + Rekvizīti",
    "neviens": "Bez pusēm / rekvizītiem",
}

def _doc_type_title(doc_tips: str, fallback: str = "Dokuments") -> str:
    tip = (doc_tips or "").strip().lower()
    preset = DOCUMENT_TYPE_PRESETS.get(tip, {})
    return str(preset.get("title") or fallback or "Dokuments")


def _effective_cover_page_title(akta_dati) -> str:
    doc_title = (getattr(akta_dati, "dokumenta_nosaukums", "") or "").strip() or _doc_type_title(getattr(akta_dati, "doc_tips", "akta"), "Dokuments")
    cover_title = (getattr(akta_dati, "cover_page_title", "") or "").strip()
    auto_titles = {
        "Pieņemšanas-Nodošanas Akts",
        "Pieņemšanas–Nodošanas akts",
        "Pieņemšanas-Nodošanas akts",
    }
    try:
        auto_titles.update(str(v.get("title", "")).strip() for v in DOCUMENT_TYPE_PRESETS.values())
    except Exception:
        pass
    if (not cover_title) or (cover_title in auto_titles):
        return doc_title
    return cover_title


def _doc_number_label(akta_dati: AktaDati) -> str:
    tip = (getattr(akta_dati, "doc_tips", "") or "akta").strip().lower()
    return {
        "akta": "Akta Nr.",
        "pavadzime": "Pavadzīmes Nr.",
        "rekins": "Rēķina Nr.",
        "norakstisanas_akts": "Akta Nr.",
        "inventarizacijas_akts": "Akta Nr.",
        "nodosanas_akts": "Akta Nr.",
        "atgriesanas_akts": "Akta Nr.",
        "defektacijas_akts": "Akta Nr.",
        "cits": "Dokumenta Nr.",
    }.get(tip, "Dokumenta Nr.")


def _safe_fmt_date(date_str: str, date_format: str = "YYYY-MM-DD") -> str:
    s = (date_str or "").strip()
    if not s:
        return ""
    try:
        return datetime.strptime(s, '%Y-%m-%d').strftime(date_format.replace('YYYY', '%Y').replace('MM', '%m').replace('DD', '%d'))
    except Exception:
        return s


def _persona_has_content(p: Optional[Persona]) -> bool:
    if not p:
        return False
    try:
        return any(bool((getattr(p, k, "") or "").strip()) for k in [
            "nosaukums", "reģ_nr", "adrese", "kontaktpersona", "amats",
            "pilnvaras_pamats", "tālrunis", "epasts", "web_lapa",
            "bankas_konts", "juridiskais_statuss"
        ])
    except Exception:
        return False


def _party_role_label(akta_dati: AktaDati, role_key: str, default: str) -> str:
    try:
        if role_key == 'pieņēmējs':
            return (getattr(akta_dati, 'pieņēmēja_loma', '') or default).strip() or default
        if role_key == 'nodevējs':
            return (getattr(akta_dati, 'nodevēja_loma', '') or default).strip() or default
    except Exception:
        pass
    return default


def _iter_visible_parties(akta_dati: AktaDati):
    out = []
    party_mode = (getattr(akta_dati, 'party_mode', '') or 'puses').strip().lower()
    if party_mode not in ('puses', 'abi'):
        return out
    pie = getattr(akta_dati, 'pieņēmējs', None)
    nod = getattr(akta_dati, 'nodevējs', None)
    if _persona_has_content(pie):
        out.append((_party_role_label(akta_dati, 'pieņēmējs', 'Pieņēmējs'), pie, getattr(akta_dati, 'paraksts_pieņēmējs_ceļš', '') or ''))
    if _persona_has_content(nod):
        out.append((_party_role_label(akta_dati, 'nodevējs', 'Iekārtas/u /pakalpojuma/u nodevējs'), nod, getattr(akta_dati, 'paraksts_nodevējs_ceļš', '') or ''))
    return out


def _iter_visible_rekviziti(akta_dati: AktaDati):
    out = []
    party_mode = (getattr(akta_dati, 'party_mode', '') or 'puses').strip().lower()
    if party_mode not in ('rekviziti', 'abi'):
        return out
    rek = getattr(akta_dati, 'rekviziti', None)
    if _persona_has_content(rek):
        label = (getattr(akta_dati, 'rekvizitu_virsraksts', '') or 'Rekvizīti').strip() or 'Rekvizīti'
        out.append((label, rek))
    return out


def _resolve_signature_name(akta_dati: AktaDati, role_key: str, party: Persona):
    override_map = {
        'rekviziti': getattr(akta_dati, 'paraksta_vards_rekviziti', '') or '',
        'pieņēmējs': getattr(akta_dati, 'paraksta_vards_pienemejs', '') or '',
        'nodevējs': getattr(akta_dati, 'paraksta_vards_nodevejs', '') or '',
    }
    custom_name = str(override_map.get(role_key, '') or '').strip()
    if custom_name:
        return custom_name
    return (getattr(party, 'kontaktpersona', '') or getattr(party, 'nosaukums', '') or '').strip()


def _iter_signature_parties(akta_dati: AktaDati):
    out = []
    for lbl, party, sig_path in _iter_visible_parties(akta_dati):
        out.append((lbl, _resolve_signature_name(akta_dati, 'pieņēmējs' if 'pieņēm' in lbl.lower() else 'nodevējs', party), sig_path))
    for lbl, party in _iter_visible_rekviziti(akta_dati):
        out.append((lbl, _resolve_signature_name(akta_dati, 'rekviziti', party), ''))
    return out


def _iter_extra_signature_rows(akta_dati: AktaDati):
    """Atgriež lietotāja definētās papildu parakstu rindas zem galvenajām parakstu rindām."""
    out = []
    raw = getattr(akta_dati, 'papildu_parakstu_rindas', []) or []
    if isinstance(raw, str):
        raw = raw.splitlines()
    for item in raw:
        label = ''
        signer = ''
        if isinstance(item, dict):
            label = str(item.get('nosaukums') or item.get('label') or item.get('loma') or '').strip()
            signer = str(item.get('parakstitajs') or item.get('parakstītājs') or item.get('vards') or item.get('name') or '').strip()
        else:
            text = str(item or '').strip()
            if '|' in text:
                label, signer = [x.strip() for x in text.split('|', 1)]
            elif ';' in text:
                label, signer = [x.strip() for x in text.split(';', 1)]
            else:
                label = text
        if label or signer:
            out.append({'nosaukums': label or 'Paraksts', 'parakstitajs': signer})
    return out


def _extra_signature_rows_to_text(rows) -> str:
    lines = []
    for item in rows or []:
        if isinstance(item, dict):
            label = str(item.get('nosaukums') or '').strip()
            signer = str(item.get('parakstitajs') or item.get('parakstītājs') or '').strip()
        else:
            label = str(item or '').strip()
            signer = ''
        if label or signer:
            lines.append(f"{label} | {signer}" if signer else label)
    return "\n".join(lines)


def _parse_extra_signature_rows_text(text: str) -> list:
    rows = []
    for line in str(text or '').splitlines():
        line = line.strip()
        if not line:
            continue
        if '|' in line:
            label, signer = [x.strip() for x in line.split('|', 1)]
        elif ';' in line:
            label, signer = [x.strip() for x in line.split(';', 1)]
        else:
            label, signer = line, ''
        if label or signer:
            rows.append({'nosaukums': label or 'Paraksts', 'parakstitajs': signer})
    return rows


def _signature_header_needed(signature_parties):
    try:
        parts = list(signature_parties or [])
    except Exception:
        parts = []
    if len(parts) <= 1:
        return False
    labels = []
    for item in parts:
        try:
            lbl = str(item[0] or '').strip()
        except Exception:
            lbl = ''
        if lbl:
            labels.append(lbl.lower())
    if not labels:
        return False
    generic = {'rekvizīti', 'rekviziti', 'parakstītājs', 'parakstitajs', 'paraksts'}
    return any(lbl not in generic for lbl in labels) or len(set(labels)) > 1


def _render_persona_lines(prefix: str, p: Persona, include_prefix: bool = True):
    title = (prefix or '').strip()
    name = (getattr(p, 'nosaukums', '') or '').strip()
    lines = []
    if include_prefix and title and name:
        lines.append(f"<b>{title}:</b> {name}")
    elif title and not include_prefix:
        if name:
            lines.append(f"<b>{title}</b>")
            lines.append(name)
        else:
            lines.append(f"<b>{title}</b>")
    elif name:
        lines.append(name)
    if getattr(p, 'reģ_nr', ''): lines.append(f"Reģ. Nr.: {p.reģ_nr}")
    if getattr(p, 'adrese', ''): lines.append(f"Adrese: {p.adrese}")
    if getattr(p, 'kontaktpersona', ''): lines.append(f"Kontaktpersona: {p.kontaktpersona}")
    if getattr(p, 'amats', ''): lines.append(f"Amats: {p.amats}")
    if getattr(p, 'pilnvaras_pamats', '') and str(p.pilnvaras_pamats).strip() not in ('', 'Pilnvaras pamats'):
        lines.append(f"Pilnvaras pamats: {p.pilnvaras_pamats}")
    if getattr(p, 'tālrunis', ''): lines.append(f"Tālrunis: {p.tālrunis}")
    if getattr(p, 'epasts', ''):
        em = p.epasts.strip()
        lines.append(f'E-pasts: <a href="mailto:{em}">{em}</a>')
    if getattr(p, 'web_lapa', ''):
        url = p.web_lapa.strip()
        if url and not re.match(r"^[a-zA-Z]+://", url):
            url = "https://" + url
        disp = p.web_lapa.strip()
        lines.append(f'Web lapa: <a href="{url}">{disp}</a>')
    if getattr(p, 'bankas_konts', ''): lines.append(f"Bankas konts: {p.bankas_konts}")
    if getattr(p, 'juridiskais_statuss', ''): lines.append(f"Statuss: {p.juridiskais_statuss}")
    return lines

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

def _coerce_path(p):
    """Normalizē ievadi uz faila ceļu (str) vai atgriež None.
    Vajadzīgs, jo dažās vietās signāli/ieraksti var nodot dict/tuple u.c.
    Pieņem arī dict ar tipiskām atslēgām: path, ceļš, file, filepath, json, pdf.
    """
    if p is None:
        return None
    # Qt signāli dažreiz padod bool (checked). To ignorējam.
    if isinstance(p, bool):
        return None
    # Ja padots tuple/list (piem. (path, filter))
    if isinstance(p, (list, tuple)) and p:
        return _coerce_path(p[0])
    # Ja padots dict ar ceļu
    if isinstance(p, dict):
        for k in ("path", "ceļš", "celsh", "file", "filepath", "json", "pdf", "value"):
            v = p.get(k)
            if isinstance(v, (str, bytes, os.PathLike)) and v:
                return os.fspath(v)
        return None
    # Parasts path-like
    if isinstance(p, (str, bytes, os.PathLike)):
        try:
            s = os.fspath(p)
        except Exception:
            return None
        return s if s else None
    # Cits tips (piem. int fd) šeit nav vajadzīgs
    return None




def _path_exists(p) -> bool:
    """Droša os.path.exists versija, kas pieņem arī dict/tuple u.c."""
    pp = _coerce_path(p)
    if not pp:
        return False
    try:
        return os.path.exists(pp)
    except Exception:
        return False

def _atomic_write_bytes(target_path: str, data: bytes):
    """Droši pārraksta failu (Windows lock-safe): raksta uz pagaidu failu un tad os.replace()."""
    try:
        target_path = os.path.abspath(target_path)
        os.makedirs(os.path.dirname(target_path), exist_ok=True)
    except Exception:
        pass

    tmp_fd = None
    tmp_path = None
    try:
        # Pagaidu fails tajā pašā mapē (lai os.replace būtu atomisks arī Windows)
        dir_name = os.path.dirname(os.path.abspath(target_path)) or "."
        fd, tmp_path = tempfile.mkstemp(prefix=".tmp_", suffix=".pdf", dir=dir_name)
        tmp_fd = fd
        with os.fdopen(fd, "wb") as f:
            f.write(data)
        tmp_fd = None
        os.replace(tmp_path, target_path)
        tmp_path = None
    finally:
        try:
            if tmp_fd is not None:
                os.close(tmp_fd)
        except Exception:
            pass
        try:
            if tmp_path and os.path.exists(tmp_path):
                os.remove(tmp_path)
        except Exception:
            pass

def _atomic_write_pdfwriter(target_path: str, writer_obj):
    """Droši saglabā PyPDF2 PdfWriter uz failu, neatsitoties pret file-lock."""
    buf = io.BytesIO()
    writer_obj.write(buf)
    _atomic_write_bytes(target_path, buf.getvalue())


def _prepare_unencrypted_pdf_for_render(pdf_path: str, password: str = "") -> tuple[str, Optional[str]]:
    """Ja PDF ir šifrēts, izveido atšifrētu pagaidu kopiju renderēšanai (poppler/pdf2image/Qt).
    Atgriež (render_path, temp_path_to_cleanup). Ja nav šifrēts, temp_path_to_cleanup būs None.

    Kāpēc tas vajadzīgs:
    - pdf2image/poppler bieži uzkārtina procesu, ja PDF ir šifrēts.
    - Priekšskatījums programmā nedrīkst prasīt paroli.
    """
    try:
        from PyPDF2 import PdfReader, PdfWriter
    except Exception:
        # Ja nav PyPDF2, vienkārši mēģinām renderēt oriģinālu (labāk nekā crash)
        return pdf_path, None

    try:
        with open(pdf_path, "rb") as rf:
            reader = PdfReader(rf)
            if not getattr(reader, "is_encrypted", False):
                return pdf_path, None

            # Mēģinām atšifrēt (PyPDF2: decrypt() atgriež 0/1/2 vai True/False atkarībā no versijas)
            try:
                ok = reader.decrypt(password or "")
            except Exception:
                ok = 0

            if not ok:
                # Nevarējām atšifrēt — neatgriežam šifrētu uz poppler (lai neuzkar),
                # bet metīsim izņēmumu, ko UI var parādīt kā kļūdu.
                raise RuntimeError("PDF ir šifrēts un paroli neizdevās pielietot priekšskatījumam/drukai.")

            writer = PdfWriter()
            for p in reader.pages:
                writer.add_page(p)

        # Saglabājam atšifrētu pagaidu PDF tajā pašā mapē (drošāk Windows)
        dir_name = os.path.dirname(os.path.abspath(pdf_path)) or "."
        fd, tmp_path = tempfile.mkstemp(prefix=".tmp_decrypted_", suffix=".pdf", dir=dir_name)
        os.close(fd)
        _atomic_write_pdfwriter(tmp_path, writer)
        return tmp_path, tmp_path

    except Exception:
        # Ja kaut kas noiet greizi, labāk lai caller redz kļūdu (nevis hang).
        raise


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

def render_pdf_to_image(pdf_path: str, poppler_path: str = None, password: str = "") -> QPixmap:
    """
    Renderē PDF faila pirmo lapu kā QPixmap.
    Nepieciešama Poppler instalācija.
    :param pdf_path: Ceļš uz PDF failu.
    :param poppler_path: (Tikai Windows) Ceļš uz Poppler bin direktoriju.
    :param password: Ja PDF ir šifrēts, lietotāja parole atšifrēšanai priekšskatījumam.
    :return: QPixmap objekts ar PDF lapas attēlu.
    """
    try:
        render_path, tmp_cleanup = _prepare_unencrypted_pdf_for_render(pdf_path, password=password)
        try:
            images = convert_from_path(render_path, first_page=1, last_page=1, poppler_path=poppler_path)
        finally:
            try:
                if tmp_cleanup and os.path.exists(tmp_cleanup):
                    os.remove(tmp_cleanup)
            except Exception:
                pass

        if images:
            from io import BytesIO
            img_byte_arr = BytesIO()
            images[0].save(img_byte_arr, format="PNG")
            img_byte_arr.seek(0)

            pixmap = QPixmap()
            pixmap.loadFromData(img_byte_arr.getvalue(), "PNG")
            return pixmap

        return QPixmap()

    except Exception as e:
        print(f"Kļūda renderējot PDF uz attēlu: {e}")
        QMessageBox.warning(
            None,
            "PDF renderēšanas kļūda",
            "Neizdevās renderēt PDF priekšskatījumu. "
            "Ja PDF ir šifrēts, pārliecinieties, ka parole ir pareiza. "
            f"\nKļūda: {e}"
        )
        return QPixmap()

# Lapu numerācija PDF dokumentam


# Lapu numerācija + dekorācijas PDF dokumentam
class DecoratedCanvas(canvas.Canvas):
    """Canvas, kas pievieno lapu numerāciju, ūdenszīmi un QR kodus (ja ieslēgts)."""

    def __init__(self, *args, **kwargs):
        self.pages = []
        self.akta_dati: AktaDati = kwargs.pop("akta_dati", None)
        self.show_page_numbers = kwargs.pop("show_page_numbers", True)
        super().__init__(*args, **kwargs)

        # Sagatavojam QR kodu attēlus vienreiz (ja vajag)
        self._qr_images = []
        try:
            if self.akta_dati:
                qr_items = []
                if self.akta_dati.include_custom_qr_code and self.akta_dati.custom_qr_code_data:
                    qr_items.append(("custom", self.akta_dati.custom_qr_code_data,
                                     float(self.akta_dati.custom_qr_code_size_mm),
                                     self.akta_dati.custom_qr_code_position,
                                     float(self.akta_dati.custom_qr_code_pos_x_mm),
                                     float(self.akta_dati.custom_qr_code_pos_y_mm)))
                if self.akta_dati.include_auto_qr_code and self.akta_dati.akta_nr:
                    qr_items.append(("auto", self.akta_dati.akta_nr,
                                     float(self.akta_dati.auto_qr_code_size_mm),
                                     self.akta_dati.auto_qr_code_position,
                                     float(self.akta_dati.auto_qr_code_pos_x_mm),
                                     float(self.akta_dati.auto_qr_code_pos_y_mm)))

                if qr_items:
                    import qrcode
                    from reportlab.lib.utils import ImageReader
                    for _, data, size_mm, pos, x_mm, y_mm in qr_items:
                        qr = qrcode.QRCode(version=1, error_correction=qrcode.constants.ERROR_CORRECT_L,
                                           box_size=10, border=2)
                        qr.add_data(data)
                        qr.make(fit=True)
                        img = qr.make_image(fill_color="black", back_color="white")
                        # ImageReader strādā ar PIL Image
                        self._qr_images.append({
                            "reader": ImageReader(img),
                            "size_pt": size_mm * mm,
                            "pos": pos,
                            "x_pt": x_mm * mm,
                            "y_pt": y_mm * mm,
                        })
        except Exception as e:
            # Nekrītam ārā, ja qrcode nav uzinstalēts vai rodas kļūda
            print(f"QR sagatavošanas kļūda: {e}")
            self._qr_images = []

    def showPage(self):
        self.pages.append(dict(self.__dict__))
        self._startPage()

    def save(self):
        page_count = len(self.pages)
        for page in self.pages:
            self.__dict__.update(page)

            # Dekorācijas uz katras lapas
            self._draw_staple_mark()
            self._draw_watermark()
            self._draw_qr_codes()

            if self.show_page_numbers:
                self._draw_page_number(page_count)

            canvas.Canvas.showPage(self)
        canvas.Canvas.save(self)


    def _draw_staple_mark(self):
        """Uzzīmē īsu vertikālu līniju augšējā kreisajā stūrī (skavotāja atzīme) uz katras lapas."""
        try:
            w, h = self._pagesize
            self.saveState()
            # Diskrēta, drukai draudzīga krāsa
            self.setStrokeColor(colors.HexColor("#94A3B8"))  # slate-400
            self.setLineWidth(0.7)
            # Pozīcija: nedaudz no kreisās malas, pie augšas
            x = 6 * mm
            y1 = h - 12 * mm
            y2 = h - 32 * mm
            self.line(x, y1, x, y2)
            self.restoreState()
        except Exception:
            try:
                self.restoreState()
            except Exception:
                pass

    def _draw_page_number(self, page_count: int):
        # Augšējais labais stūris
        self.setFont("Helvetica", 9)
        w, h = self._pagesize
        self.drawRightString(w - 18 * mm, h - 10 * mm, f"Lapa {self._pageNumber} no {page_count}")

    def _draw_watermark(self):
        if not self.akta_dati or not getattr(self.akta_dati, "add_watermark", False):
            return
        text = getattr(self.akta_dati, "watermark_text", "") or ""
        if not text.strip():
            return

        try:
            from reportlab.lib.colors import HexColor
            self.saveState()
            self.setFillColor(HexColor(getattr(self.akta_dati, "watermark_color", "#E0E0E0")))
            self.setFont("Helvetica-Bold", float(getattr(self.akta_dati, "watermark_font_size", 72)))
            w, h = self._pagesize
            self.translate(w / 2, h / 2)
            self.rotate(float(getattr(self.akta_dati, "watermark_rotation", 45)))
            self.drawCentredString(0, 0, text)
            self.restoreState()
        except Exception as e:
            print(f"Ūdenszīmes kļūda: {e}")

    def _draw_qr_codes(self):
        if not self._qr_images:
            return

        try:
            w, h = self._pagesize
            left = float(getattr(self.akta_dati, "pdf_margin_left", 18)) * mm if self.akta_dati else 18 * mm
            right = float(getattr(self.akta_dati, "pdf_margin_right", 18)) * mm if self.akta_dati else 18 * mm
            top = float(getattr(self.akta_dati, "pdf_margin_top", 16)) * mm if self.akta_dati else 16 * mm
            bottom = float(getattr(self.akta_dati, "pdf_margin_bottom", 16)) * mm if self.akta_dati else 16 * mm

            for q in self._qr_images:
                size = q["size_pt"]
                pos = q["pos"]
                if pos == "bottom_left":
                    x, y = left, bottom
                elif pos == "bottom_right":
                    x, y = w - right - size, bottom
                elif pos == "top_left":
                    x, y = left, h - top - size
                elif pos == "top_right":
                    x, y = w - right - size, h - top - size
                elif pos == "custom":
                    x, y = q["x_pt"], q["y_pt"]
                else:
                    # fallback
                    x, y = w - right - size, bottom

                self.drawImage(q["reader"], x, y, width=size, height=size, mask='auto')
        except Exception as e:
            print(f"QR zīmēšanas kļūda: {e}")


# ---------------------- Atsauces dokumentu (pielikumu) apstrāde ----------------------
def _find_soffice_exe() -> Optional[str]:
    """Atrod LibreOffice/soffice izpildāmo failu.
    Atgriež pilnu ceļu vai None, ja nav atrasts.
    """
    import shutil

    # 1) mēģinam no PATH
    p = shutil.which("soffice") or shutil.which("libreoffice")
    if p:
        return p

    # 2) tipiskie ceļi Windows
    if sys.platform == "win32":
        candidates = [
            os.path.join(os.environ.get("ProgramFiles", ""), "LibreOffice", "program", "soffice.exe"),
            os.path.join(os.environ.get("ProgramFiles(x86)", ""), "LibreOffice", "program", "soffice.exe"),
            r"C:\Program Files\LibreOffice\program\soffice.exe",
            r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
        ]
        for c in candidates:
            if c and os.path.exists(c):
                return c

    return None



def _normalize_reference_doc_payload(ref, fallback_name: str = "") -> dict:
    """Normalizē atsauces dokumenta ierakstu uz dict formu.

    Atbalsta vecos ierakstus kā str/path, dataclass/object ar laukiem `ceļš`/`nosaukums`,
    kā arī dict ar sinonīmiem `path`, `file`, `name`, `title`, `scale`, `scale_pct`.
    """
    payload = {}
    if isinstance(ref, dict):
        payload = dict(ref)
    elif isinstance(ref, (str, os.PathLike)):
        payload = {"ceļš": os.fspath(ref)}
    elif ref is not None:
        payload = {
            "ceļš": getattr(ref, "ceļš", "") or getattr(ref, "path", "") or getattr(ref, "file", ""),
            "nosaukums": getattr(ref, "nosaukums", "") or getattr(ref, "name", "") or getattr(ref, "title", ""),
            "scale_pct": getattr(ref, "scale_pct", None) or getattr(ref, "scale", None),
        }

    path = str(payload.get("ceļš") or payload.get("path") or payload.get("file") or "").strip()
    name = str(payload.get("nosaukums") or payload.get("name") or payload.get("title") or fallback_name or "").strip()
    try:
        scale_pct = int(payload.get("scale_pct", payload.get("scale", 100)) or 100)
    except Exception:
        scale_pct = 100
    scale_pct = max(10, min(500, scale_pct))
    if not name and path:
        name = os.path.basename(path)
    return {
        "ceļš": path,
        "nosaukums": name,
        "scale_pct": scale_pct,
    }


def _reference_doc_display_text(ref) -> str:
    payload = _normalize_reference_doc_payload(ref)
    name = payload.get("nosaukums") or os.path.basename(payload.get("ceļš") or "") or "Dokuments"
    scale_pct = int(payload.get("scale_pct", 100) or 100)
    return f"{name} ({scale_pct}%)" if scale_pct != 100 else name


def _reference_doc_path(ref) -> str:
    return _normalize_reference_doc_payload(ref).get("ceļš", "")




def _scale_pdf_page_to_fit(page, scale_pct: int = 100):
    """Samēro PDF lapas saturu procentos, atstājot to pašā lapas izmērā un centrējot.

    Kāpēc šis ir vajadzīgs:
    - daļā PyPDF2/pypdf versiju `merge_transformed_page` nav vai strādā nestabili;
    - daži PDF izmanto tikai `add_transformation`, citi – `merge_page` kombinācijā ar jaunu tukšu lapu.

    Funkcija cenšas izmantot vairākas saderīgas pieejas. Ja mērogs ir 100%, atgriež oriģinālo lapu.
    """
    try:
        scale = max(10, min(500, int(scale_pct or 100))) / 100.0
    except Exception:
        scale = 1.0

    if abs(scale - 1.0) < 0.0001:
        return page

    try:
        from copy import deepcopy
        from PyPDF2 import Transformation
        try:
            from PyPDF2._page import PageObject
        except Exception:
            try:
                from PyPDF2.pdf import PageObject
            except Exception:
                PageObject = None

        width = float(page.mediabox.width)
        height = float(page.mediabox.height)
        tx = (width - (width * scale)) / 2.0
        ty = (height - (height * scale)) / 2.0
        transform = Transformation().scale(scale, scale).translate(tx, ty)

        src = deepcopy(page)

        # 1) Modernākais ceļš: transformējam kopiju un iemergējam tukšā lapā.
        try:
            if hasattr(src, 'add_transformation'):
                src.add_transformation(transform)
                if PageObject is not None:
                    blank = PageObject.create_blank_page(width=width, height=height)
                    blank.merge_page(src)
                    try:
                        blank.compress_content_streams()
                    except Exception:
                        pass
                    return blank
        except Exception:
            pass

        # 2) Vecāks PyPDF2 variants.
        try:
            if PageObject is not None:
                blank = PageObject.create_blank_page(width=width, height=height)
                if hasattr(blank, 'merge_transformed_page'):
                    blank.merge_transformed_page(src, transform)
                    try:
                        blank.compress_content_streams()
                    except Exception:
                        pass
                    return blank
                if hasattr(blank, 'mergeTransformedPage'):
                    blank.mergeTransformedPage(src, transform)
                    try:
                        blank.compress_content_streams()
                    except Exception:
                        pass
                    return blank
        except Exception:
            pass

        # 3) Pēdējais fallback – transformējam pašu kopiju un atgriežam to.
        try:
            if hasattr(src, 'scale_by') and hasattr(src, 'add_transformation'):
                src.add_transformation(transform)
                return src
        except Exception:
            pass

        return page
    except Exception:
        return page

def _convert_attachment_to_pdf(input_path: str, out_dir: str) -> Optional[str]:
    """Konvertē pielikumu uz PDF, lai to varētu pievienot akta beigās.

    Atbalsts:
      - PDF: atgriež oriģinālo ceļu
      - DOC/DOCX: LibreOffice (soffice) -> PDF (writer_pdf_Export)
      - XLS/XLSX/ODS: LibreOffice -> PDF (calc_pdf_Export)
      - PPT/PPTX: LibreOffice -> PDF (impress_pdf_Export)

    Atgriež PDF ceļu vai None, ja neizdevās.
    """
    input_path = _coerce_path(input_path)
    if not input_path:
        return None

    input_path = os.path.abspath(input_path)
    if not os.path.exists(input_path):
        return None

    ext = os.path.splitext(input_path)[1].lower()
    if ext == ".pdf":
        return input_path

    supported = {".docx", ".doc", ".xlsx", ".xls", ".ods", ".odt", ".pptx", ".ppt"}
    if ext not in supported:
        return None

    soffice = _find_soffice_exe()
    if not soffice:
        # Nav LibreOffice -> nevar padarīt aplūkojamu PDF beigās
        return None

    # izvēlamies pareizo LO filtru (stabilāk priekš XLSX)
    if ext in {".xls", ".xlsx", ".ods"}:
        convert_to = "pdf:calc_pdf_Export"
    elif ext in {".ppt", ".pptx"}:
        convert_to = "pdf:impress_pdf_Export"
    else:
        convert_to = "pdf:writer_pdf_Export"

    try:
        os.makedirs(out_dir, exist_ok=True)

        cmd = [
            soffice,
            "--headless",
            "--nologo",
            "--nolockcheck",
            "--nodefault",
            "--norestore",
            "--convert-to", convert_to,
            "--outdir", out_dir,
            input_path,
        ]

        res = subprocess.run(cmd, capture_output=True, text=True, timeout=30, **({"creationflags": subprocess.CREATE_NO_WINDOW} if os.name == "nt" and hasattr(subprocess, "CREATE_NO_WINDOW") else {}))

        if res.returncode != 0:
            # noder debuggam (konsolē), bet UI netraucē
            print(f"Konvertēšanas kļūda ({input_path}): {res.stderr or res.stdout}")
            return None

        # meklējam izveidoto PDF tieši šajā mapē
        pdfs = [os.path.join(out_dir, f) for f in os.listdir(out_dir) if f.lower().endswith('.pdf')]
        if not pdfs:
            return None

        # Ja ir vairāki, ņemam jaunāko
        pdfs.sort(key=lambda fp: os.path.getmtime(fp), reverse=True)
        return pdfs[0]

    except subprocess.TimeoutExpired:
        # LibreOffice iestrēga (piem., liels DOCX/XLSX vai dialogi).
        return None
    except FileNotFoundError:
        # izpildāmais fails nav atrasts
        return None
    except Exception as e:
        print(f"Konvertēšanas izņēmums ({input_path}): {e}")
        return None


def _make_annex_title_pdf(title: str, out_path: str, pagesize=A4, font_name: str = "Helvetica"):
    """Izveido vienas lapas PDF informācijas lapu (ja pielikumu nevar konvertēt).
    Teksts ir mazs, augšējā kreisajā stūrī, ar pareizām garumzīmēm (izmanto font_name).
    """
    # nodrošinām, ka mape eksistē (citādi Windows met [Errno 2])
    try:
        os.makedirs(os.path.dirname(os.path.abspath(out_path)), exist_ok=True)
    except Exception:
        pass

    from reportlab.pdfbase.pdfmetrics import stringWidth
    from reportlab.lib.utils import simpleSplit

    c = canvas.Canvas(out_path, pagesize=pagesize)
    w, h = pagesize

    left = 18 * mm
    top_y = h - 12 * mm
    max_w = w - 2 * left

    # Bold variants only if Helvetica, citādi izmantojam to pašu fontu (drošāk diakritikām)
    bold_font = "Helvetica-Bold" if (font_name or "") == "Helvetica" else (font_name or "Helvetica")
    normal_font = font_name or "Helvetica"

    def fit_one_line(txt: str, fnt: str, size: float) -> str:
        if stringWidth(txt, fnt, size) <= max_w:
            return txt
        ell = "…"
        # brutāla, bet stabila saīsināšana
        for cut in range(len(txt), 0, -1):
            cand = txt[:cut].rstrip() + ell
            if stringWidth(cand, fnt, size) <= max_w:
                return cand
        return ell

    title_fit = fit_one_line(title, bold_font, 10)

    c.setFillColor(colors.HexColor("#0F172A"))
    c.setFont(bold_font, 10)
    c.drawString(left, top_y, title_fit)

    c.setStrokeColor(colors.HexColor("#CBD5E1"))
    c.setLineWidth(0.8)
    c.line(left, top_y - 4, w - left, top_y - 4)

    info = "Pielikums pie pieņemšanas–nodošanas akta. Dokuments pievienots automātiski."
    c.setFont(normal_font, 9)
    c.setFillColor(colors.HexColor("#334155"))
    lines = simpleSplit(info, normal_font, 9, max_w)
    y = top_y - 14
    for ln in lines[:5]:
        c.drawString(left, y, ln)
        y -= 11

    c.showPage()
    c.save()


def _overlay_text_on_pdf_page(page, text: str, font_name: str, font_size: float, x_pt: float, y_pt: float, bold: bool = False):
    """Uzvelk tekstu uz dotās PDF lapas (PyPDF2 PageObject), izmantojot ReportLab overlay.

    Svarīgi:
    - Diakritikām jābūt redzamām -> izmanto font_name (TTF)
    - Teksts NETIEK izvilkts ārpus lapas -> ja par garu, saīsina ar "…"
    """
    try:
        from PyPDF2 import PdfReader
        from reportlab.pdfgen import canvas as rl_canvas
        from reportlab.pdfbase.pdfmetrics import stringWidth
    except Exception:
        return page

    try:
        w = float(page.mediabox.width)
        h = float(page.mediabox.height)
        left = float(x_pt)
        max_w = w - left - (18 * mm)  # labā mala

        # izvēlamies fontu (bold tikai Helvetica gadījumā)
        fnt = font_name if font_name else "Helvetica"
        if bold and fnt == "Helvetica":
            fnt = "Helvetica-Bold"

        def fit_one_line(txt: str) -> str:
            if stringWidth(txt, fnt, float(font_size)) <= max_w:
                return txt
            ell = "…"
            for cut in range(len(txt), 0, -1):
                cand = txt[:cut].rstrip() + ell
                if stringWidth(cand, fnt, float(font_size)) <= max_w:
                    return cand
            return ell

        safe_text = fit_one_line(text)

        buf = io.BytesIO()
        c = rl_canvas.Canvas(buf, pagesize=(w, h))
        c.setFont(fnt, float(font_size))
        c.setFillColor(colors.HexColor("#0F172A"))
        # drošības dēļ y ierobežojam
        y = min(max(float(y_pt), 0), h - 2)
        c.drawString(left, y, safe_text)
        c.save()
        buf.seek(0)
        overlay_page = PdfReader(buf).pages[0]
        page.merge_page(overlay_page)
    except Exception:
        pass
    return page


def _apply_global_page_numbers_to_pdf(pdf_path: str, akta_dati: AktaDati, font_name: str):
    """Pievieno lapu numerāciju visam PDF (arī pielikumiem), lai kopējais skaits ir pareizs."""
    if not getattr(akta_dati, "show_page_numbers", True):
        return

    try:
        from PyPDF2 import PdfReader, PdfWriter
        from reportlab.pdfgen import canvas as rl_canvas
    except Exception as e:
        print(f"Nevar uzlikt lapu numerāciju (trūkst PyPDF2/ReportLab): {e}")
        return

    try:
        reader = PdfReader(pdf_path)
        total = len(reader.pages)
        writer = PdfWriter()

        for idx, page in enumerate(reader.pages, start=1):
            w = float(page.mediabox.width)
            h = float(page.mediabox.height)

            buf = io.BytesIO()
            c = rl_canvas.Canvas(buf, pagesize=(w, h))
            c.setFillColor(colors.HexColor("#0F172A"))
            c.setFont(font_name if font_name else "Helvetica", 9)
            # Augšējais labais stūris (tāpat kā DecoratedCanvas)
            c.drawRightString(w - 18 * mm, h - 10 * mm, f"Lapa {idx} no {total}")
            c.save()
            buf.seek(0)

            overlay_page = PdfReader(buf).pages[0]
            page.merge_page(overlay_page)
            writer.add_page(page)

        _atomic_write_pdfwriter(pdf_path, writer)
    except Exception as e:
        print(f"Lapu numerācijas kļūda: {e}")



def _apply_stapler_mark_to_pdf(pdf_path: str, akta_dati: AktaDati, color_hex: str = "#94A3B8"):
    """Pievieno 'skavotāja līniju' katrai PDF lapai (arī pielikumiem), pēcapstrādē ar PyPDF2."""
    try:
        from PyPDF2 import PdfReader, PdfWriter
        from reportlab.pdfgen import canvas as rl_canvas
    except Exception as e:
        print(f"Nevar uzlikt skavotāja līniju (trūkst PyPDF2/ReportLab): {e}")
        return

    try:
        reader = PdfReader(pdf_path)
        writer = PdfWriter()

        for page_index, page in enumerate(reader.pages):
            if getattr(akta_dati, 'qr_kods_tikai_pirma_lapa', False) and page_index != 0:
                writer.add_page(page)
                continue

            w = float(page.mediabox.width)
            h = float(page.mediabox.height)

            buf = io.BytesIO()
            c = rl_canvas.Canvas(buf, pagesize=(w, h))
            try:
                c.setStrokeColor(colors.HexColor(color_hex))
            except Exception:
                c.setStrokeColor(colors.HexColor("#94A3B8"))
            c.setLineWidth(0.7)

            x = 6 * mm
            y1 = h - 12 * mm
            y2 = h - 32 * mm
            c.line(x, y1, x, y2)

            c.save()
            buf.seek(0)

            overlay = PdfReader(buf).pages[0]
            page.merge_page(overlay)
            writer.add_page(page)

        _atomic_write_pdfwriter(pdf_path, writer)
    except Exception as e:
        print(f"Skavotāja līnijas kļūda: {e}")




def _build_qr_payload(akta_dati: AktaDati) -> str:
    """QR saturs.
    - URL režīms: QR satur verifikācijas URL ar parametriem (a,d,pc,h).
    - Pretējā gadījumā: kompakts JSON ar akta datiem un (saīsinātām) pozīcijām.
    """
    try:
        if getattr(akta_dati, "qr_kods_url_mode", False) and (getattr(akta_dati, "qr_kods_url", "") or "").strip():
            base = akta_dati.qr_kods_url.strip()
            poz_count = len(getattr(akta_dati, "pozīcijas", []) or [])
            full = json.dumps(
                {"akta_nr": akta_dati.akta_nr, "datums": akta_dati.datums, "poz_count": poz_count},
                ensure_ascii=False,
                separators=(",", ":"),
            )
            h = hashlib.sha256(full.encode("utf-8")).hexdigest()[:16]
            from urllib.parse import urlencode
            qs = urlencode({"a": akta_dati.akta_nr, "d": akta_dati.datums, "pc": poz_count, "h": h})
            return base + ("&" if "?" in base else "?") + qs

        payload = {"v": 1, "akta_nr": akta_dati.akta_nr, "datums": akta_dati.datums, "vieta": getattr(akta_dati, "vieta", "")}

        if getattr(akta_dati, "qr_kods_ieklaut_pozicijas", True):
            items = []
            for i, p in enumerate(getattr(akta_dati, "pozīcijas", []) or []):
                if i >= 20:
                    break
                try:
                    name = (getattr(p, "nosaukums", "") or "").strip()
                    if len(name) > 40:
                        name = name[:40] + "…"
                    items.append({"n": name, "q": str(getattr(p, "daudzums", "")), "u": (getattr(p, "vienība", "") or "").strip(),
                                  "sn": (getattr(p, "sērijas_nr", "") or "").strip()[:30]})
                except Exception:
                    continue
            payload["poz"] = items
            payload["poz_count"] = len(getattr(akta_dati, "pozīcijas", []) or [])

        s = json.dumps(payload, ensure_ascii=False, separators=(",", ":"))
        if len(s) > 1200 and "poz" in payload:
            payload["poz"] = payload["poz"][:8]
            s = json.dumps(payload, ensure_ascii=False, separators=(",", ":"))

        if len(s) > 1200:
            poz = getattr(akta_dati, "pozīcijas", []) or []
            full = json.dumps(
                {"akta_nr": akta_dati.akta_nr, "datums": akta_dati.datums,
                 "pozīcijas": [asdict(p) if hasattr(p, "__dataclass_fields__") else str(p) for p in poz]},
                ensure_ascii=False,
                separators=(",", ":"),
            )
            h = hashlib.sha256(full.encode("utf-8")).hexdigest()[:16]
            s = json.dumps({"v": 1, "akta_nr": akta_dati.akta_nr, "datums": akta_dati.datums, "poz_count": len(poz), "poz_hash": h},
                           ensure_ascii=False, separators=(",", ":"))
        return s
    except Exception:
        return ""


def _apply_qr_to_pdf(pdf_path: str, akta_dati: AktaDati):
    """Pievieno QR kodu PDF apakšējā kreisajā stūrī.
    - QR tikai pirmajā lapā, ja qr_kods_tikai_pirma_lapa=True
    - URL režīmā uz QR uzliek klikšķināmu linku
    """
    try:
        if not getattr(akta_dati, "qr_kods_enabled", True):
            return
        from PyPDF2 import PdfReader, PdfWriter
        from reportlab.pdfgen import canvas as rl_canvas
    except Exception:
        return

    payload = _build_qr_payload(akta_dati)
    if not payload:
        return

    try:
        reader = PdfReader(pdf_path)
        writer = PdfWriter()

        size_mm = float(getattr(akta_dati, "qr_kods_izmers_mm", Decimal("18.0")))
        first_only = bool(getattr(akta_dati, "qr_kods_tikai_pirma_lapa", True))

        for page_index, page in enumerate(reader.pages):
            if first_only and page_index != 0:
                writer.add_page(page)
                continue

            w = float(page.mediabox.width)
            h = float(page.mediabox.height)

            try:
                ml = float(getattr(akta_dati, "pdf_margin_left", Decimal("15.0")))
                mb = float(getattr(akta_dati, "pdf_margin_bottom", Decimal("15.0")))
            except Exception:
                ml, mb = 15.0, 15.0

            max_size = max(10.0, min(size_mm, ml - 2.0, mb - 2.0))
            qr_size = max_size * mm

            x = 2 * mm
            y = 2 * mm

            buf = io.BytesIO()
            c = rl_canvas.Canvas(buf, pagesize=(w, h))

            try:
                if getattr(akta_dati, "qr_kods_url_mode", False) and (getattr(akta_dati, "qr_kods_url", "") or "").strip():
                    c.linkURL(payload, (x, y, x + qr_size, y + qr_size), relative=0)
            except Exception:
                pass

            try:
                widget = rl_qr.QrCodeWidget(payload)
                bounds = widget.getBounds()
                bw = bounds[2] - bounds[0]
                bh = bounds[3] - bounds[1]
                scale = qr_size / max(bw, bh)
                d = Drawing(qr_size, qr_size, transform=[scale, 0, 0, scale, 0, 0])
                d.add(widget)
                renderPDF.draw(d, c, x, y)
            except Exception:
                pass

            c.save()
            buf.seek(0)
            overlay = PdfReader(buf).pages[0]
            page.merge_page(overlay)
            writer.add_page(page)

        _atomic_write_pdfwriter(pdf_path, writer)
    except Exception:
        pass


def _append_reference_docs_to_pdf(main_pdf_path: str, akta_dati: AktaDati, pagesize=A4, font_name: str = "Helvetica") -> str:
    """Pievieno atsauces dokumentus PDF beigās kā Atvasinājums 1, 2, 3...

    Prasības:
    - "Atvasinājums X: ..." ir mazs teksts augšējā kreisajā stūrī TAJĀ PAŠĀ lapā, kur sākas pielikums
    - Teksts nedrīkst pārsniegt lapas robežas (saīsinām ar "…")
    - Ja DOCX/XLSX nevar konvertēt uz PDF, pievienojam informācijas lapu ar korektām garumzīmēm
    """
    refs = getattr(akta_dati, "atsauces_dokumenti_faili", []) or []
    if not refs:
        return main_pdf_path

    try:
        from PyPDF2 import PdfReader, PdfWriter
    except Exception as e:
        print(f"PyPDF2 nav pieejams pielikumiem: {e}")
        return main_pdf_path

    tmp_root = tempfile.mkdtemp(prefix="akta_refs_")
    try:
        writer = PdfWriter()

        # Pamata PDF
        reader_main = PdfReader(main_pdf_path)
        for p in reader_main.pages:
            writer.add_page(p)

        annex_no = 0
        for ref in refs:
            try:
                payload = _normalize_reference_doc_payload(ref)
                ref_path = payload.get("ceļš", "")
                ref_name = payload.get("nosaukums", "") or os.path.basename(ref_path)
                ref_scale_pct = int(payload.get('scale_pct', 100) or 100)

                if not ref_path or not os.path.exists(ref_path):
                    continue

                annex_no += 1

                # Konvertējam katru pielikumu savā apakšmapē, lai nekad nesajauktu PDF nosaukumus
                tmp_dir = os.path.join(tmp_root, f"conv_{annex_no}")
                os.makedirs(tmp_dir, exist_ok=True)
                converted = _convert_attachment_to_pdf(ref_path, tmp_dir)

                if not converted:
                    info_pdf = os.path.join(tmp_dir, f"Atvasinajums_{annex_no}_info.pdf")
                    _make_annex_title_pdf(
                        f"Atvasinājums {annex_no}: {ref_name} (neizdevās konvertēt uz PDF)",
                        info_pdf,
                        pagesize=pagesize,
                        font_name=font_name if font_name else "Helvetica"
                    )
                    r = PdfReader(info_pdf)
                    for p in r.pages:
                        writer.add_page(p)
                    continue

                r = PdfReader(converted)
                if not r.pages:
                    continue

                # Uz pirmās pielikuma lapas uzliekam label (mazs, top-left, saīsināts ja vajag)
                first = _scale_pdf_page_to_fit(r.pages[0], ref_scale_pct)
                label = f"Atvasinājums {annex_no}: {ref_name}"
                _overlay_text_on_pdf_page(
                    first,
                    label,
                    font_name=font_name if font_name else "Helvetica",
                    font_size=9,
                    x_pt=18 * mm,
                    y_pt=float(first.mediabox.height) - 12 * mm,
                    bold=False
                )
                writer.add_page(first)

                for p in r.pages[1:]:
                    writer.add_page(_scale_pdf_page_to_fit(p, ref_scale_pct))

            except Exception as e:
                print(f"Pielikuma pievienošanas kļūda: {e}")

        _atomic_write_pdfwriter(main_pdf_path, writer)

    finally:
        try:
            shutil.rmtree(tmp_root, ignore_errors=True)
        except Exception:
            pass

    return main_pdf_path

# ---------------------- PDF ģenerēšana ----------------------

def ģenerēt_pdf(akta_dati: AktaDati, pdf_ceļš: str = None, include_reference_docs: bool = True, encrypt_pdf: bool = True):
    # --- FIX v46: normalize pdf_ceļš if dict leaked from state ---
    if isinstance(pdf_ceļš, dict):
        pdf_ceļš = pdf_ceļš.get('path') or ''
    font_name = reģistrēt_fontu(akta_dati.fonts_ceļš)
    doc_title = (getattr(akta_dati, "dokumenta_nosaukums", "") or "").strip() or _doc_type_title(getattr(akta_dati, "doc_tips", "akta"), "Dokuments")

    styles = getSampleStyleSheet()
    # Ensure all Decimal values are converted to float when used with ReportLab's float-based units or font sizes
    # Uzlaboti stili ar jaunajiem iestatījumiem
    styles.add(ParagraphStyle(name='LatvHead', fontName=font_name, fontSize=float(akta_dati.pdf_font_size_head), leading=float(akta_dati.pdf_font_size_head) * float(akta_dati.line_spacing_multiplier), spaceAfter=8, textColor=colors.HexColor(akta_dati.header_text_color)))
    styles.add(ParagraphStyle(name='Latv', fontName=font_name, fontSize=float(akta_dati.pdf_font_size_normal), leading=float(akta_dati.pdf_font_size_normal) * float(akta_dati.line_spacing_multiplier), textColor=colors.HexColor(akta_dati.footer_text_color)))
    styles.add(ParagraphStyle(name='LatvSmall', fontName=font_name, fontSize=float(akta_dati.pdf_font_size_small), leading=float(akta_dati.pdf_font_size_small) * float(akta_dati.line_spacing_multiplier), textColor=colors.HexColor(akta_dati.footer_text_color)))
    styles.add(ParagraphStyle(name='LatvTableContent', fontName=font_name, fontSize=float(akta_dati.pdf_font_size_table), leading=float(akta_dati.pdf_font_size_table) * float(akta_dati.line_spacing_multiplier), wordWrap='LTR', splitLongWords=0, alignment={'left': 0, 'center': 1, 'right': 2}.get(akta_dati.table_content_alignment, 0)))
    styles.add(ParagraphStyle(name='LatvElectronicSignature', fontName=font_name, fontSize=float(akta_dati.pdf_font_size_normal), leading=float(akta_dati.pdf_font_size_normal) * float(akta_dati.line_spacing_multiplier), alignment=1, textColor=colors.HexColor(akta_dati.header_text_color)))
    styles.add(ParagraphStyle(name='DocTitle', fontName=font_name, fontSize=float(akta_dati.document_title_font_size), alignment=1, textColor=colors.HexColor(akta_dati.document_title_color)))
    styles.add(ParagraphStyle(name='SectionHeading', fontName=font_name, fontSize=float(akta_dati.section_heading_font_size), leading=float(akta_dati.section_heading_font_size) * float(akta_dati.paragraph_line_spacing_multiplier), textColor=colors.HexColor(akta_dati.section_heading_color)))

    # ---------------------- Premium juridiskais dizains (uzlabojumi) ----------------------
    # Bold fonts: ja tiek lietots Helvetica, izmantojam Helvetica-Bold, citādi atstājam fontu (ja nav bold varianta).
    bold_font_name = "Helvetica-Bold" if font_name == "Helvetica" else font_name

    # Papildu stili skaidrai hierarhijai (juridisks + moderns)
    styles.add(ParagraphStyle(
        name='FieldLabel',
        fontName=bold_font_name,
        fontSize=float(akta_dati.pdf_font_size_normal),
        leading=float(akta_dati.pdf_font_size_normal) * float(akta_dati.line_spacing_multiplier),
        textColor=colors.HexColor("#0F172A")
    ))
    styles.add(ParagraphStyle(
        name='FieldValue',
        fontName=font_name,
        fontSize=float(akta_dati.pdf_font_size_normal),
        leading=float(akta_dati.pdf_font_size_normal) * float(akta_dati.line_spacing_multiplier),
        textColor=colors.HexColor("#0F172A")
    ))
    styles.add(ParagraphStyle(
        name='SectionBar',
        fontName=bold_font_name,
        fontSize=float(max(11, int(akta_dati.section_heading_font_size))),
        leading=float(max(11, int(akta_dati.section_heading_font_size))) * float(akta_dati.paragraph_line_spacing_multiplier),
        textColor=colors.HexColor("#0F172A"),
        backColor=colors.HexColor("#F1F5F9"),
        borderPadding=6,
        spaceBefore=10,
        spaceAfter=6,
    ))
    styles.add(ParagraphStyle(
        name='LegalSectionTitle',
        fontName=bold_font_name,
        fontSize=11,
        leading=14,
        textColor=colors.HexColor("#0F172A"),
        spaceBefore=10,
        spaceAfter=4,
    ))
    styles.add(ParagraphStyle(
        name='LegalBody',
        fontName=font_name,
        fontSize=float(akta_dati.pdf_font_size_normal),
        leading=float(akta_dati.pdf_font_size_normal) * float(akta_dati.line_spacing_multiplier),
        textColor=colors.HexColor("#0F172A"),
        spaceAfter=6,
    ))
    styles.add(ParagraphStyle(
        name='TableHeader',
        fontName=bold_font_name,
        fontSize=float(akta_dati.pdf_font_size_table),
        leading=float(akta_dati.pdf_font_size_table) * float(akta_dati.line_spacing_multiplier),
        textColor=colors.HexColor("#0F172A"),
        alignment=1,  # CENTER
    ))
    # Ja nav iestatīts alternējošās rindas tonis, piešķiram klusu "enterprise" noklusējumu
    if not getattr(akta_dati, "table_alternate_row_color", ""):
        akta_dati.table_alternate_row_color = "#F8FAFC"

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
        title=doc_title,
    )

    story = []
    available_width = pagesize[0] - float(akta_dati.pdf_margin_left) * mm - float(akta_dati.pdf_margin_right) * mm

    # Cover Page (new feature)
    if akta_dati.add_cover_page:
        story.append(Spacer(1, 2 * inch))
        if akta_dati.logotipa_ceļš and _path_exists(akta_dati.logotipa_ceļš):
            try:
                cover_logo = RLImage(_coerce_path(akta_dati.logotipa_ceļš))
                cover_logo._restrictSize(float(akta_dati.cover_page_logo_width_mm) * mm, 50 * mm)
                story.append(cover_logo)
                story.append(Spacer(1, 0.5 * inch))
            except Exception:
                pass
        story.append(Paragraph(_effective_cover_page_title(akta_dati), styles['DocTitle']))
        story.append(Spacer(1, 1 * inch))
        tip = (getattr(akta_dati, "doc_tips", "") or "akta").strip()
        nr_label = _doc_number_label(akta_dati)
        story.append(Paragraph(f"{nr_label}: {akta_dati.akta_nr}", styles['Latv']))
        story.append(Paragraph(f"Datums: {_safe_fmt_date(akta_dati.datums, akta_dati.date_format)}", styles['Latv']))
        if tip == "rekins" and (getattr(akta_dati, "apmaksas_termins", "") or "").strip():
            story.append(Paragraph(f"Apmaksas termiņš: {akta_dati.apmaksas_termins}", styles['Latv']))
        if tip == "pavadzime" and (getattr(akta_dati, "piegades_datums", "") or "").strip():
            story.append(Paragraph(f"Piegādes datums: {akta_dati.piegades_datums}", styles['Latv']))
        story.append(Paragraph(f"Vieta: {akta_dati.vieta}", styles['Latv']))
        story.append(Spacer(1, 2 * inch))
        for role_label, party, _sig in _iter_visible_parties(akta_dati):
            story.append(Paragraph(f"{role_label}: {party.nosaukums}", styles['Latv']))
            if getattr(party, 'tālrunis', ''):
                story.append(Paragraph(f"{role_label} tālrunis: {party.tālrunis}", styles['Latv']))
        for rek_label, rek in _iter_visible_rekviziti(akta_dati):
            story.append(Paragraph(f"{rek_label}: {getattr(rek, 'nosaukums', '')}", styles['Latv']))
            if getattr(rek, 'tālrunis', ''):
                story.append(Paragraph(f"{rek_label} tālrunis: {rek.tālrunis}", styles['Latv']))
        story.append(PageBreak())

    # Header ar logo un nosaukumu
    header_table_data = []
    logo_w = float(akta_dati.pdf_logo_width_mm) * mm
    if (not getattr(akta_dati, 'cover_page_enabled', False)) and akta_dati.logotipa_ceļš and _path_exists(akta_dati.logotipa_ceļš):
        try:
            header_logo = RLImage(_coerce_path(akta_dati.logotipa_ceļš))
            header_logo._restrictSize(logo_w, 20 * mm)
            header_table_data.append([header_logo, Paragraph(f"<font name='{bold_font_name}'>{doc_title.upper()}</font>", styles['LatvHead'])])
        except Exception:
            header_table_data.append(["", Paragraph(f"<font name='{bold_font_name}'>{doc_title.upper()}</font>", styles['LatvHead'])])
    else:
        header_table_data.append(["", Paragraph(f"<font name='{bold_font_name}'>{doc_title.upper()}</font>", styles['LatvHead'])])

    ht = Table(header_table_data, colWidths=[logo_w, None])
    ht.setStyle(TableStyle([
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('ALIGN', (1, 0), (1, 0), 'LEFT'),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 6),
    ]))
    story.append(ht)

    # Aktu metadati (Nr., Datums, Vieta)
    nr_label = _doc_number_label(akta_dati)
    md = [[Paragraph(f"<font name='{font_name}'><b>{nr_label}:</b> {akta_dati.akta_nr}</font>", styles['Latv']),
           Paragraph(f"<font name='{font_name}'><b>Datums:</b> {_safe_fmt_date(akta_dati.datums, akta_dati.date_format)}</font>", styles['Latv']),
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
    if getattr(akta_dati, 'ieklaut_izpildes_terminu', True) and akta_dati.izpildes_termiņš:
        story.append(Paragraph(f"<font name='{font_name}'><b>Izpildes termiņš:</b> {_safe_fmt_date(akta_dati.izpildes_termiņš, akta_dati.date_format)}</font>", styles['Latv']))
    if getattr(akta_dati, 'ieklaut_pienemsanas_datumu', True) and akta_dati.pieņemšanas_datums:
        story.append(Paragraph(f"<font name='{font_name}'><b>Pieņemšanas datums:</b> {_safe_fmt_date(akta_dati.pieņemšanas_datums, akta_dati.date_format)}</font>", styles['Latv']))
    if getattr(akta_dati, 'ieklaut_nodosanas_datumu', True) and akta_dati.nodošanas_datums:
        story.append(Paragraph(f"<font name='{font_name}'><b>Nodošanas datums:</b> {_safe_fmt_date(akta_dati.nodošanas_datums, akta_dati.date_format)}</font>", styles['Latv']))
    story.append(Spacer(1, 6))

    # Puses: dinamiski rāda tikai aizpildītās sadaļas
    def persona_paragraph(prefix: str, p: Persona):
        lines = [f"<b>{prefix}:</b> {p.nosaukums}"]
        if p.reģ_nr: lines.append(f"Reģ. Nr.: {p.reģ_nr}")
        if p.adrese: lines.append(f"Adrese: {p.adrese}")
        if p.kontaktpersona: lines.append(f"Kontaktpersona: {p.kontaktpersona}")
        if getattr(p, 'amats', ''): lines.append(f"Amats: {p.amats}")
        if getattr(p, 'pilnvaras_pamats', ''): lines.append(f"Pilnvaras pamats: {p.pilnvaras_pamats}")
        if p.tālrunis: lines.append(f"Tālrunis: {p.tālrunis}")
        if p.epasts:
            em = p.epasts.strip()
            lines.append(f"E-pasts: <a href=\"mailto:{em}\">{em}</a>")
        if getattr(p, "web_lapa", ""):
            url = p.web_lapa.strip()
            if url and not re.match(r"^[a-zA-Z]+://", url):
                url = "https://" + url
            disp = p.web_lapa.strip()
            lines.append(f"Web lapa: <a href=\"{url}\">{disp}</a>")
        if p.bankas_konts: lines.append(f"Bankas konts: {p.bankas_konts}")
        if p.juridiskais_statuss: lines.append(f"Statuss: {p.juridiskais_statuss}")
        try:
            extra_map = getattr(akta_dati, 'custom_fields', {}) or {}
            sec_key = 'rek'
            role_key = (prefix or '').strip().lower()
            if 'pieņēm' in role_key or 'pienem' in role_key:
                sec_key = 'pie'
            elif 'nodev' in role_key:
                sec_key = 'nod'
            extra_fields = extra_map.get(sec_key, {}) if isinstance(extra_map, dict) else {}
            if isinstance(extra_fields, dict):
                for extra_name, extra_val in extra_fields.items():
                    extra_name = str(extra_name or '').strip()
                    extra_val = str(extra_val or '').strip()
                    if extra_name and extra_val:
                        lines.append(f"{extra_name}: {extra_val}")
        except Exception:
            pass
        return Paragraph(f"<font name='{font_name}'>" + "<br/>".join(lines) + "</font>", styles['Latv'])

    visible_parties = _iter_visible_parties(akta_dati)
    if visible_parties:
        story.append(Paragraph("PUSES", styles['SectionBar']))
        available_width = pagesize[0] - float(akta_dati.pdf_margin_left) * mm - float(akta_dati.pdf_margin_right) * mm
        col_width_parties = available_width / max(1, len(visible_parties))

        puses = Table([[persona_paragraph(lbl, party) for lbl, party, _sig in visible_parties]],
                      colWidths=[col_width_parties] * len(visible_parties))
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

    visible_rekviziti = _iter_visible_rekviziti(akta_dati)
    available_width = pagesize[0] - float(akta_dati.pdf_margin_left) * mm - float(akta_dati.pdf_margin_right) * mm
    if visible_rekviziti:
        for rek_label, rek in visible_rekviziti:
            story.append(Paragraph(str(rek_label or 'REKVIZĪTI').upper(), styles['SectionBar']))
            rek_tbl = Table([[persona_paragraph(rek_label, rek)]], colWidths=[available_width])
            rek_tbl.setStyle(TableStyle([
                ('VALIGN', (0, 0), (-1, -1), 'TOP'),
                ('BOX', (0, 0), (-1, -1), float(akta_dati.table_border_thickness_pt), colors.HexColor(akta_dati.table_grid_color)),
                ('INNERGRID', (0, 0), (-1, -1), float(akta_dati.table_border_thickness_pt), colors.HexColor(akta_dati.table_grid_color)),
                ('LEFTPADDING', (0, 0), (-1, -1), float(akta_dati.table_cell_padding_mm)),
                ('RIGHTPADDING', (0, 0), (-1, -1), float(akta_dati.table_cell_padding_mm)),
                ('TOPPADDING', (0, 0), (-1, -1), float(akta_dati.table_cell_padding_mm)),
                ('BOTTOMPADDING', (0, 0), (-1, -1), float(akta_dati.table_cell_padding_mm)),
            ]))
            story.append(rek_tbl)
            story.append(Spacer(1, 8))

    # Pozīciju tabula
    story.append(Paragraph("POZĪCIJAS", styles['SectionBar']))
    # JAUNS: kolonnas var būt paslēptas/ pārdēvētas (poz_columns_config + custom_columns.visible)
    poz_cfg = _merge_poz_columns_config(getattr(akta_dati, "poz_columns_config", None))
    def _col_visible(key: str) -> bool:
        try:
            return bool(poz_cfg.get(key, {}).get("visible", True))
        except Exception:
            return True
    def _col_title(key: str, default: str) -> str:
        try:
            t = str(poz_cfg.get(key, {}).get("title", default))
            return t or default
        except Exception:
            return default

    # Kolonnu secība PDF tabulai (ņem vērā GUI pārvietošanu / drag&drop)
    # Nr. kolonna PDF tabulā ir atsevišķa un (ja ieslēgta) vienmēr ir pirmā.
    base_kind = {
        "apraksts": "text",
        "daudzums": "num",
        "vieniba": "text",
        "cena": "money",
        "summa": "money",
        "serial": "text",
        "warranty": "text",
        "notes": "text",
        "foto": "foto",
    }

    # Pielāgotās kolonnas (no datiem)
    custom_cols = []
    try:
        custom_cols = list(getattr(akta_dati, "custom_columns", []) or [])
    except Exception:
        custom_cols = []

    # Iespējamās atslēgas secībai (bez "nr")
    base_keys = ["apraksts", "daudzums", "vieniba", "cena", "summa", "serial", "warranty", "notes"]
    custom_keys = [f"custom:{i}" for i in range(len(custom_cols))]
    possible = set(base_keys + custom_keys + ["foto"])

    # Vēlamā secība no GUI (ja ir). Ja nav – izmantojam noklusējumu.
    desired = []
    try:
        desired = list(getattr(akta_dati, "poz_columns_visual_order", []) or [])
    except Exception:
        desired = []
    desired = [k for k in desired if k in possible]

    if not desired:
        desired = list(base_keys) + list(custom_keys) + ["foto"]
    else:
        # pieliekam trūkstošās kolonnas klāt (lai nezaudējam jaunpievienotas kolonnas)
        for k in base_keys:
            if k not in desired:
                desired.append(k)
        for k in custom_keys:
            if k not in desired:
                desired.append(k)
        if "foto" not in desired:
            desired.append("foto")

    # Foto kolonna var atrasties jebkurā vietā (tāpat kā GUI). 
    # Neuzspiežam tai pozīciju — izmantojam lietotāja (GUI) secību.

    # Uztaisām PDF kolonnu plānu, ņemot vērā redzamību
    col_plan = []  # list[tuple(key, kind)]
    if _col_visible("nr"):
        col_plan.append(("nr", "nr"))

    for key in desired:
        if key.startswith("custom:"):
            try:
                ci = int(key.split(":", 1)[1])
                if 0 <= ci < len(custom_cols):
                    cdef = custom_cols[ci]
                    if isinstance(cdef, dict) and bool(cdef.get("visible", True)):
                        col_plan.append((key, "text"))
            except Exception:
                pass
            continue

        if key == "foto":
            if _col_visible("foto"):
                col_plan.append(("foto", "foto"))
            continue

        # bāzes kolonnas
        if key in base_keys and _col_visible(key):
            col_plan.append((key, base_kind.get(key, "text")))

    # Galvene
    tab_header_items = []
    for key, _kind in col_plan:
        if key == "nr":
            tab_header_items.append(Paragraph(_col_title("nr", "Nr."), styles['TableHeader']))
        elif key.startswith("custom:"):
            try:
                ci = int(key.split(":", 1)[1])
                nm = str(custom_cols[ci].get("name", "")) if isinstance(custom_cols[ci], dict) else ""
            except Exception:
                nm = ""
            tab_header_items.append(Paragraph(nm or "", styles['TableHeader']))
        else:
            default_map = {
                "apraksts": "Apraksts",
                "daudzums": "Daudzums",
                "vieniba": "Vienība",
                "cena": "Cena",
                "summa": "Summa",
                "serial": "Seriālais Nr.",
                "warranty": "Garantija",
                "notes": "Piezīmes pozīcijai",
                "foto": "Foto",
            }
            tab_header_items.append(Paragraph(_col_title(key, default_map.get(key, key)), styles['TableHeader']))

    tab_data = [tab_header_items]

    # Rindas
    for i, poz in enumerate(akta_dati.pozīcijas, start=1):
        row_items = []
        for key, kind in col_plan:
            if key == "nr":
                row_items.append(Paragraph(str(i), styles['LatvTableContent']))
            elif key == "apraksts":
                row_items.append(Paragraph(poz.apraksts, styles['LatvTableContent']))
            elif key == "daudzums":
                row_items.append(Paragraph(f"{formēt_naudu(poz.daudzums)}", styles['LatvTableContent']))
            elif key == "vieniba":
                row_items.append(Paragraph(poz.vienība, styles['LatvTableContent']))
            elif key == "cena":
                row_items.append(Paragraph(f"{akta_dati.currency_symbol_position == 'before' and akta_dati.valūta or ''}{formēt_naudu(poz.cena)}{akta_dati.currency_symbol_position == 'after' and ' ' + akta_dati.valūta or ''}", styles['LatvTableContent']))
            elif key == "summa":
                row_items.append(Paragraph(f"{akta_dati.currency_symbol_position == 'before' and akta_dati.valūta or ''}{formēt_naudu(poz.summa)}{akta_dati.currency_symbol_position == 'after' and ' ' + akta_dati.valūta or ''}", styles['LatvTableContent']))
            elif key == "serial":
                row_items.append(Paragraph(poz.seriālais_nr, styles['LatvTableContent']))
            elif key == "warranty":
                row_items.append(Paragraph(poz.garantija, styles['LatvTableContent']))
            elif key == "notes":
                row_items.append(Paragraph(poz.piezīmes_pozīcijai, styles['LatvTableContent']))
            elif key.startswith("custom:"):
                try:
                    ci = int(key.split(":", 1)[1])
                    data_list = custom_cols[ci].get("data", []) if isinstance(custom_cols[ci], dict) else []
                    v = str(data_list[i-1]) if isinstance(data_list, list) and (i-1) < len(data_list) else ""
                except Exception:
                    v = ""
                row_items.append(Paragraph(v, styles['LatvTableContent']))
            elif key == "foto":
                if getattr(poz, 'attēla_ceļš', '') and os.path.exists(poz.attēla_ceļš):
                    try:
                        img_thumb = RLImage(poz.attēla_ceļš)
                        img_thumb._restrictSize(18 * mm, 14 * mm)
                        row_items.append(img_thumb)
                    except Exception:
                        row_items.append(Paragraph("", styles['LatvTableContent']))
                else:
                    row_items.append(Paragraph("", styles['LatvTableContent']))
            else:
                row_items.append(Paragraph("", styles['LatvTableContent']))
        tab_data.append(row_items)

    # Kolonnu platumi
    expected_cols = len(col_plan)
    col_widths = None
    try:
        col_widths_mm = [float(x.strip()) for x in (akta_dati.table_col_widths or "").split(',') if x.strip()]
        if len(col_widths_mm) == expected_cols:
            col_widths = [w * mm for w in col_widths_mm]
    except Exception:
        col_widths = None

    if not col_widths:
        # Noklusējuma platumi (mm) atkarībā no kolonnu tipa
        default_map_mm = {
            "nr": 10,
            "text": 35,
            "num": 18,
            "money": 20,
            "foto": 18,
        }
        dw = []
        for _k, kind in col_plan:
            dw.append(default_map_mm.get(kind, 25))
        col_widths = [w * mm for w in dw]
        total_default_width = sum(col_widths)
        if total_default_width > 0 and total_default_width != available_width:
            scale_factor = available_width / total_default_width
            col_widths = [w * scale_factor for w in col_widths]

    t = Table(tab_data, colWidths=col_widths)

    tstyle = TableStyle([
        ('FONTNAME', (0,0), (-1,-1), font_name),
        ('FONTNAME', (0,0), (-1,0), bold_font_name),
        ('FONTSIZE', (0,0), (-1,0), akta_dati.pdf_font_size_table),
        ('FONTSIZE', (0,1), (-1,-1), akta_dati.pdf_font_size_table),
        ('BACKGROUND', (0,0), (-1,0), colors.HexColor(akta_dati.table_header_bg_color or "#E5E7EB")),
        ('TEXTCOLOR', (0,0), (-1,0), colors.HexColor("#0F172A")),
        ('ALIGN', (0,0), (-1,0), 'CENTER'),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('GRID', (0,0), (-1,-1), 0.4, colors.HexColor(akta_dati.table_grid_color or "#CBD5E1")),
        ('LINEBELOW', (0,0), (-1,0), 1.0, colors.HexColor("#94A3B8")),
        ('BOTTOMPADDING', (0,0), (-1,0), float(akta_dati.table_cell_padding_mm) * 2),
        ('TOPPADDING', (0,0), (-1,0), float(akta_dati.table_cell_padding_mm) * 2),
        ('BOTTOMPADDING', (0,1), (-1,-1), float(akta_dati.table_cell_padding_mm)),
        ('TOPPADDING', (0,1), (-1,-1), float(akta_dati.table_cell_padding_mm)),
    ])

    # Kolonnu izlīdzinājumi pēc tipa
    for col_i, (_k, kind) in enumerate(col_plan):
        if kind in ("money",):
            tstyle.add('ALIGN', (col_i, 1), (col_i, -1), 'RIGHT')
        elif kind in ("num", "nr"):
            tstyle.add('ALIGN', (col_i, 1), (col_i, -1), 'CENTER')
        elif kind == "foto":
            tstyle.add('ALIGN', (col_i, 1), (col_i, -1), 'CENTER')
        else:
            tstyle.add('ALIGN', (col_i, 1), (col_i, -1), 'LEFT')

    # Apply alternate row color
    if akta_dati.table_alternate_row_color:
        for i in range(1, len(tab_data)):
            if i % 2 == 0: # Even rows (0-indexed, so actual even rows)
                tstyle.add('BACKGROUND', (0, i), (-1, i), colors.HexColor(akta_dati.table_alternate_row_color))

    t.setStyle(tstyle)
    story.append(t)

    # Kopsavilkums (var izslēgt)
    if getattr(akta_dati, "show_price_summary", True):
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
                ('LINEBELOW', (0,0), (-1,-1), 0.6, colors.HexColor('#CBD5E1')),
            ]))
            story.append(ts)

    # Piezīmes
    if akta_dati.piezīmes:
        story.append(Spacer(1, 6))
        story.append(Paragraph(f"<font name='{font_name}'><b>Vispārīgās piezīmes:</b><br/>{akta_dati.piezīmes}</font>", styles['Latv']))

    # Jauni juridiski saistoši lauki PDF dokumentā
    if akta_dati.strīdu_risināšana:
        story.append(Paragraph("Strīdu risināšanas kārtība", styles['LegalSectionTitle']))
        story.append(Paragraph(f"<font name='{font_name}'>{akta_dati.strīdu_risināšana}</font>", styles['LegalBody']))

    if akta_dati.konfidencialitātes_klauzula:
        story.append(Paragraph("Konfidencialitātes klauzula", styles['LegalSectionTitle']))
        story.append(Paragraph(
            f"<font name='{font_name}'>Puses apņemas neizpaust trešajām personām informāciju, kas iegūta šī akta ietvaros, "
            f"izņemot gadījumus, ko nosaka normatīvie akti.</font>",
            styles['LegalBody']
        ))

    if akta_dati.soda_nauda_procenti > 0:
        story.append(Spacer(1, 6))
        story.append(Paragraph(f"<font name='{font_name}'><b>Soda nauda:</b> Par saistību neizpildi vai nepienācīgu izpildi, vainīgā puse maksā otrai pusei soda naudu {formēt_naudu(akta_dati.soda_nauda_procenti)}% apmērā no neizpildīto saistību vērtības par katru kavējuma dienu.</font>", styles['Latv']))

    if akta_dati.piegādes_nosacījumi:
        story.append(Spacer(1, 6))
        story.append(Paragraph(f"<font name='{font_name}'><b>Piegādes nosacījumi:</b> {akta_dati.piegādes_nosacījumi}</font>", styles['Latv']))

    if akta_dati.apdrošināšana:
        story.append(Spacer(1, 6))
        apd_text = (getattr(akta_dati, "apdrošināšana_teksts", "") or "").strip()
        if not apd_text:
            apd_text = "Preces ir apdrošinātas pret bojājumiem un zaudējumiem līdz pieņemšanas-nodošanas brīdim."
        story.append(Paragraph(f"<font name='{font_name}'><b>Apdrošināšana:</b> {apd_text}</font>", styles['Latv']))

    if akta_dati.papildu_nosacījumi:
        story.append(Spacer(1, 6))
        story.append(Paragraph(f"<font name='{font_name}'><b>Papildu nosacījumi:</b><br/>{akta_dati.papildu_nosacījumi}</font>", styles['Latv']))

    if akta_dati.atsauces_dokumenti:
        story.append(Spacer(1, 6))
        story.append(Paragraph(f"<font name='{font_name}'><b>Atsauces dokumenti:</b> {akta_dati.atsauces_dokumenti}</font>", styles['Latv']))

    story.append(Spacer(1, 6))
    # Akta statuss (izcelts "juridisks premium" bloks)
    status_tbl = Table([[Paragraph(f"<font name='{bold_font_name}'>Akta statuss:</font>", styles['FieldLabel']),
                         Paragraph(f"<font name='{font_name}'>{akta_dati.akta_statuss}</font>", styles['FieldValue'])]],
                       colWidths=[35*mm, None])
    # JAUNS: Akta statusa fona krāsa PDF (atkarīga no izvēlētā statusa)
    _status_bg_map = {
        "Melnraksts": "#F1F5F9",   # pelēcīgs
        "Apstiprināts": "#DBEAFE", # zils
        "Parakstīts": "#DCFCE7",   # zaļš
        "Arhivēts": "#EDE9FE",     # violets
        "Atcelts": "#FEE2E2",      # sarkans
    }
    status_bg_color = _status_bg_map.get((akta_dati.akta_statuss or "").strip(), "#F8FAFC")

    status_tbl.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,-1), colors.HexColor(status_bg_color)),
        ('BOX', (0,0), (-1,-1), 1.0, colors.HexColor("#334155")),
        ('INNERGRID', (0,0), (-1,-1), 0.0, colors.white),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('LEFTPADDING', (0,0), (-1,-1), 8),
        ('RIGHTPADDING', (0,0), (-1,-1), 8),
        ('TOPPADDING', (0,0), (-1,-1), 6),
        ('BOTTOMPADDING', (0,0), (-1,-1), 6),
    ]))
    story.append(Spacer(1, 8))
    story.append(status_tbl)


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

    signature_mode = (getattr(akta_dati, 'paraksta_rezims', 'physical') or 'physical').strip().lower()
    if akta_dati.elektroniskais_paraksts:
        signature_mode = 'electronic'

    if signature_mode == 'electronic' and akta_dati.radit_elektronisko_parakstu_tekstu:
        story.append(Spacer(1, 16))
        story.append(Paragraph(
            f"<font name='{font_name}' size='{akta_dati.pdf_font_size_normal}'><b>ŠIS DOKUMENTS PARAKSTĪTS AR DROŠU ELEKTRONISKO PARAKSTU UN SATUR LAIKA ZĪMOGU</b></font>",
            styles['LatvElectronicSignature']
        ))

    elif signature_mode == 'no_signature_required':
        no_sig_text = (getattr(akta_dati, 'paraksta_nav_teksts', '') or '').strip()
        if no_sig_text:
            story.append(Spacer(1, 12))
            story.append(Paragraph(f"<font name='{font_name}'><b>{no_sig_text}</b></font>", styles['Latv']))

    elif akta_dati.parakstu_rindas and signature_mode == 'physical':
        visible_parties = _iter_signature_parties(akta_dati)
        if visible_parties:
            story.append(Spacer(1, 12))
            story.append(Paragraph("PARAKSTI", styles['SectionBar']))
            sig_row = []
            for _lbl, signer_name, _sig in visible_parties:
                nm = (signer_name or '').strip()
                sig_row.append(Paragraph(f"<font name='{font_name}'>____________________________<br/>{nm}</font>", styles['Latv']))
            table_rows = [sig_row]
            if _signature_header_needed(visible_parties):
                hdr_row = [Paragraph(f"<font name='{bold_font_name}'>{lbl}</font>", styles['TableHeader']) for lbl, _name, _sig in visible_parties]
                table_rows = [hdr_row, sig_row]
            sig_tbl = Table(table_rows, colWidths=[available_width / max(1, len(visible_parties))] * len(visible_parties))
            sig_tbl.setStyle(TableStyle([
                ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
                ('ALIGN', (0,0), (-1,-1), 'CENTER'),
                ('BOX', (0,0), (-1,-1), 0.6, colors.HexColor('#CBD5E1')),
                ('INNERGRID', (0,0), (-1,-1), 0.6, colors.HexColor('#CBD5E1')),
                ('TOPPADDING', (0,0), (-1,-1), 8),
                ('BOTTOMPADDING', (0,0), (-1,-1), 8),
            ]))
            story.append(sig_tbl)

        extra_signature_rows = _iter_extra_signature_rows(akta_dati)
        if extra_signature_rows:
            story.append(Spacer(1, 8))
            extra_rows = []
            for row in extra_signature_rows:
                lbl = (row.get('nosaukums') or 'Paraksts').strip()
                nm = (row.get('parakstitajs') or '').strip()
                extra_rows.append([Paragraph(f"<font name='{bold_font_name}'>{lbl}</font>", styles['TableHeader'])])
                extra_rows.append([Paragraph(f"<font name='{font_name}'>____________________________<br/>{nm}</font>", styles['Latv'])])
            extra_tbl = Table(extra_rows, colWidths=[available_width])
            extra_tbl.setStyle(TableStyle([
                ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
                ('ALIGN', (0,0), (-1,-1), 'CENTER'),
                ('BOX', (0,0), (-1,-1), 0.6, colors.HexColor('#CBD5E1')),
                ('INNERGRID', (0,0), (-1,-1), 0.6, colors.HexColor('#CBD5E1')),
                ('TOPPADDING', (0,0), (-1,-1), 8),
                ('BOTTOMPADDING', (0,0), (-1,-1), 8),
            ]))
            story.append(extra_tbl)

    doc.build(story)

    try:
        if include_reference_docs:
            _append_reference_docs_to_pdf(pdf_ceļš, akta_dati, pagesize=pagesize, font_name=font_name)
    except Exception as e:
        print(f"Atsauces pielikumu pievienošanas kļūda: {e}")

    try:
        if getattr(akta_dati, 'include_custom_qr_code', False) or getattr(akta_dati, 'include_auto_qr_code', False):
            _apply_qr_to_pdf(pdf_ceļš, akta_dati)
    except Exception as e:
        print(f"QR koda kļūda: {e}")

    try:
        _apply_global_page_numbers_to_pdf(pdf_ceļš, akta_dati, font_name)
    except Exception as e:
        print(f"Lapu numuru kļūda: {e}")

    try:
        if getattr(akta_dati, 'show_stapler_mark', False):
            _apply_stapler_mark_to_pdf(pdf_ceļš, akta_dati)
    except Exception as e:
        print(f"Skavotāja atzīmes kļūda: {e}")

    return pdf_ceļš


def ģenerēt_docx(akta_dati: AktaDati, docx_ceļš: str = None) -> str:
    doc_title = (getattr(akta_dati, 'dokumenta_nosaukums', '') or '').strip() or _doc_type_title(getattr(akta_dati, 'doc_tips', 'akta'))
    nr_label = _doc_number_label(akta_dati)

    if docx_ceļš is None:
        temp_file = tempfile.NamedTemporaryFile(delete=False, suffix='.docx')
        docx_ceļš = temp_file.name
        temp_file.close()

    document = Document()

    try:
        section = document.sections[0]
        section.top_margin = Mm(15)
        section.bottom_margin = Mm(15)
        section.left_margin = Mm(18)
        section.right_margin = Mm(18)
    except Exception:
        pass

    try:
        document.core_properties.title = doc_title
    except Exception:
        pass

    p = document.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run(doc_title.upper())
    r.bold = True

    meta = document.add_paragraph()
    add_formatted_text(meta, f"{nr_label}: {akta_dati.akta_nr}\nDatums: {_safe_fmt_date(akta_dati.datums, akta_dati.date_format)}\nVieta: {akta_dati.vieta}")

    tip = (getattr(akta_dati, 'doc_tips', '') or 'akta').strip().lower()
    if tip == 'rekins' and (getattr(akta_dati, 'apmaksas_termins', '') or '').strip():
        add_formatted_text(document.add_paragraph(), f"Apmaksas termiņš: {akta_dati.apmaksas_termins}")
    if tip == 'pavadzime' and (getattr(akta_dati, 'piegades_datums', '') or '').strip():
        add_formatted_text(document.add_paragraph(), f"Piegādes datums: {akta_dati.piegades_datums}")

    visible_parties = _iter_visible_parties(akta_dati)
    if visible_parties:
        hdr = document.add_paragraph()
        hdr.add_run('PUSES').bold = True
        table = document.add_table(rows=1, cols=len(visible_parties))
        table.autofit = True
        try:
            table.style = 'Table Grid'
        except Exception:
            pass
        for i, (lbl, party, _sig) in enumerate(visible_parties):
            cell = table.rows[0].cells[i]
            txt = [lbl + ': ' + (party.nosaukums or '')]
            if getattr(party, 'reģ_nr', ''): txt.append(f"Reģ. Nr.: {party.reģ_nr}")
            if getattr(party, 'adrese', ''): txt.append(f"Adrese: {party.adrese}")
            if getattr(party, 'kontaktpersona', ''): txt.append(f"Kontaktpersona: {party.kontaktpersona}")
            if getattr(party, 'amats', ''): txt.append(f"Amats: {party.amats}")
            if getattr(party, 'pilnvaras_pamats', ''): txt.append(f"Pilnvaras pamats: {party.pilnvaras_pamats}")
            if getattr(party, 'tālrunis', ''): txt.append(f"Tālrunis: {party.tālrunis}")
            if getattr(party, 'epasts', ''): txt.append(f"E-pasts: {party.epasts}")
            if getattr(party, 'bankas_konts', ''): txt.append(f"Bankas konts: {party.bankas_konts}")
            add_formatted_text(cell.paragraphs[0], "\n".join(txt))

    hdr = document.add_paragraph()
    hdr.add_run('POZĪCIJAS').bold = True
    poz_cfg = _merge_poz_columns_config(getattr(akta_dati, 'poz_columns_config', None))
    custom_cols = list(getattr(akta_dati, 'custom_columns', []) or [])
    col_defs = []
    base_defs = [
        ('nr', 'Nr.'), ('apraksts', 'Apraksts'), ('daudzums', 'Daudzums'), ('vieniba', 'Vienība'),
        ('cena', 'Cena'), ('summa', 'Summa'), ('serial', 'Seriālais Nr.'), ('warranty', 'Garantija'), ('notes', 'Piezīmes pozīcijai')
    ]
    for key, default_title in base_defs:
        c = poz_cfg.get(key, {}) if isinstance(poz_cfg, dict) else {}
        if bool(c.get('visible', True)):
            col_defs.append((key, str(c.get('title', default_title)) or default_title))
    for i, cdef in enumerate(custom_cols):
        if isinstance(cdef, dict) and bool(cdef.get('visible', True)):
            col_defs.append((f'custom:{i}', str(cdef.get('name', '')) or f'Papildu kolonna {i+1}'))

    table = document.add_table(rows=1, cols=max(1, len(col_defs)))
    table.autofit = True
    try:
        table.style = 'Table Grid'
    except Exception:
        pass
    for i, (_key, title) in enumerate(col_defs):
        cell = table.rows[0].cells[i]
        cell.text = title
        if cell.paragraphs and cell.paragraphs[0].runs:
            cell.paragraphs[0].runs[0].bold = True

    for idx, poz in enumerate(akta_dati.pozīcijas, start=1):
        row = table.add_row().cells
        for ci, (key, _title) in enumerate(col_defs):
            val = ''
            if key == 'nr': val = str(idx)
            elif key == 'apraksts': val = poz.apraksts
            elif key == 'daudzums': val = formēt_naudu(poz.daudzums)
            elif key == 'vieniba': val = poz.vienība
            elif key == 'cena': val = f"{akta_dati.valūta} {formēt_naudu(poz.cena)}" if akta_dati.currency_symbol_position == 'before' else f"{formēt_naudu(poz.cena)} {akta_dati.valūta}"
            elif key == 'summa': val = f"{akta_dati.valūta} {formēt_naudu(poz.summa)}" if akta_dati.currency_symbol_position == 'before' else f"{formēt_naudu(poz.summa)} {akta_dati.valūta}"
            elif key == 'serial': val = poz.seriālais_nr
            elif key == 'warranty': val = poz.garantija
            elif key == 'notes': val = poz.piezīmes_pozīcijai
            elif key.startswith('custom:'):
                try:
                    ci2 = int(key.split(':', 1)[1])
                    data_list = custom_cols[ci2].get('data', []) if isinstance(custom_cols[ci2], dict) else []
                    val = str(data_list[idx-1]) if isinstance(data_list, list) and idx-1 < len(data_list) else ''
                except Exception:
                    val = ''
            row[ci].text = val

    if getattr(akta_dati, 'show_price_summary', True):
        document.add_paragraph()
        add_formatted_text(document.add_paragraph(), f"Kopā bez PVN: {formēt_naudu(akta_dati.kopējā_summma())} {akta_dati.valūta}")
        if akta_dati.iekļaut_pvn and getattr(akta_dati, 'show_vat_breakdown', True):
            add_formatted_text(document.add_paragraph(), f"PVN {akta_dati.pvn_likme}%: {formēt_naudu(akta_dati.pvn_summa())} {akta_dati.valūta}")
            add_formatted_text(document.add_paragraph(), f"Kopā ar PVN: {formēt_naudu(akta_dati.summa_ar_pvn())} {akta_dati.valūta}")

    if akta_dati.piezīmes:
        add_formatted_text(document.add_paragraph(), f"Vispārīgās piezīmes: {akta_dati.piezīmes}")
    if akta_dati.papildu_nosacījumi:
        add_formatted_text(document.add_paragraph(), f"Papildu nosacījumi: {akta_dati.papildu_nosacījumi}")
    if akta_dati.atsauces_dokumenti:
        add_formatted_text(document.add_paragraph(), f"Atsauces dokumenti: {akta_dati.atsauces_dokumenti}")

    signature_mode = (getattr(akta_dati, 'paraksta_rezims', 'physical') or 'physical').strip().lower()
    if akta_dati.elektroniskais_paraksts:
        signature_mode = 'electronic'
    signature_parties = _iter_signature_parties(akta_dati)
    if signature_mode == 'electronic' and akta_dati.radit_elektronisko_parakstu_tekstu:
        p = document.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r = p.add_run('ŠIS DOKUMENTS PARAKSTĪTS AR DROŠU ELEKTRONISKO PARAKSTU UN SATUR LAIKA ZĪMOGU')
        r.bold = True
    elif signature_mode == 'no_signature_required':
        no_sig_text = (getattr(akta_dati, 'paraksta_nav_teksts', '') or '').strip()
        if no_sig_text:
            p = document.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            r = p.add_run(no_sig_text)
            r.bold = True
    elif akta_dati.parakstu_rindas and signature_mode == 'physical' and signature_parties:
        hdr = document.add_paragraph()
        hdr.add_run('PARAKSTI').bold = True
        header_needed = _signature_header_needed(signature_parties)
        table_signatures = document.add_table(rows=(2 if header_needed else 1), cols=len(signature_parties))
        table_signatures.autofit = True
        try:
            table_signatures.style = 'Table Grid'
        except Exception:
            pass
        for i, (lbl, signer_name, _sig) in enumerate(signature_parties):
            row_idx = 0
            if header_needed:
                table_signatures.rows[0].cells[i].text = lbl
                row_idx = 1
            cell = table_signatures.rows[row_idx].cells[i]
            add_formatted_text(cell.paragraphs[0], '____________________________')
            add_formatted_text(cell.add_paragraph(), (signer_name or '').strip())
            for p in cell.paragraphs:
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER

        extra_signature_rows = _iter_extra_signature_rows(akta_dati)
        if extra_signature_rows:
            table_extra = document.add_table(rows=len(extra_signature_rows) * 2, cols=1)
            try:
                table_extra.style = 'Table Grid'
            except Exception:
                pass
            rr = 0
            for row in extra_signature_rows:
                lbl = (row.get('nosaukums') or 'Paraksts').strip()
                nm = (row.get('parakstitajs') or '').strip()
                table_extra.rows[rr].cells[0].text = lbl
                table_extra.rows[rr].cells[0].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
                rr += 1
                cell = table_extra.rows[rr].cells[0]
                add_formatted_text(cell.paragraphs[0], '____________________________')
                add_formatted_text(cell.add_paragraph(), nm)
                for p in cell.paragraphs:
                    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                rr += 1

    add_formatted_text(document.add_paragraph(), f"Dokuments ģenerēts: {datetime.now().strftime('%Y-%m-%d %H:%M')}")
    document.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.RIGHT
    document.save(docx_ceļš)
    return docx_ceļš


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



class DraggablePreviewLabel(QLabel):
    """QLabel, kas ļauj ar peles vilcienu pārvietot PDF priekšskatījumu (pan) pat pie zoom."""
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self._scroll_area = None
        self._dragging = False
        self._last_pos = None
        self.setCursor(Qt.OpenHandCursor)

    def set_scroll_area(self, sa: QScrollArea):
        self._scroll_area = sa

    def mousePressEvent(self, event):
        if event.button() == Qt.LeftButton and self._scroll_area is not None:
            self._dragging = True
            self._last_pos = event.globalPosition().toPoint()
            self.setCursor(Qt.ClosedHandCursor)
            event.accept()
            return
        super().mousePressEvent(event)

    def mouseMoveEvent(self, event):
        if self._dragging and self._scroll_area is not None and self._last_pos is not None:
            p = event.globalPosition().toPoint()
            dx = p.x() - self._last_pos.x()
            dy = p.y() - self._last_pos.y()
            self._last_pos = p

            h = self._scroll_area.horizontalScrollBar()
            v = self._scroll_area.verticalScrollBar()
            h.setValue(h.value() - dx)
            v.setValue(v.value() - dy)
            event.accept()
            return
        super().mouseMoveEvent(event)

    def mouseReleaseEvent(self, event):
        if event.button() == Qt.LeftButton:
            self._dragging = False
            self._last_pos = None
            self.setCursor(Qt.OpenHandCursor)
            event.accept()
            return
        super().mouseReleaseEvent(event)



class PannableScrollArea(QScrollArea):
    """QScrollArea ar 'hand-drag' panning PDF priekšskatījumam (strādā arī, ja klikšķis ir tukšajā zonā)."""
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self._dragging = False
        self._last_pos = None
        self.setCursor(Qt.OpenHandCursor)
        # Centrs, ja saturs mazāks par viewport
        try:
            self.setAlignment(Qt.AlignCenter)
        except Exception:
            pass

    def mousePressEvent(self, event):
        if event.button() == Qt.LeftButton:
            self._dragging = True
            self._last_pos = event.globalPosition().toPoint()
            self.setCursor(Qt.ClosedHandCursor)
            event.accept()
            return
        super().mousePressEvent(event)

    def mouseMoveEvent(self, event):
        if self._dragging and self._last_pos is not None:
            p = event.globalPosition().toPoint()
            dx = p.x() - self._last_pos.x()
            dy = p.y() - self._last_pos.y()
            self._last_pos = p

            h = self.horizontalScrollBar()
            v = self.verticalScrollBar()
            h.setValue(h.value() - dx)
            v.setValue(v.value() - dy)
            event.accept()
            return
        super().mouseMoveEvent(event)

    def mouseReleaseEvent(self, event):
        if event.button() == Qt.LeftButton:
            self._dragging = False
            self._last_pos = None
            self.setCursor(Qt.OpenHandCursor)
            event.accept()
            return
        super().mouseReleaseEvent(event)

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


    # Proxy helper, so this widget can be used like a QTextEdit in the rest of the code
    def setPlaceholderText(self, text: str):
        """Set placeholder text on the internal QTextEdit (if supported by Qt version)."""
        if hasattr(self.text_edit, "setPlaceholderText"):
            self.text_edit.setPlaceholderText(text)
        else:
            # Older Qt versions: QTextEdit might not support placeholder text
            # (keep silent to avoid crashing)
            pass

    def setPlainText(self, text: str):
        self.text_edit.setPlainText(text)

    def toPlainText(self) -> str:
        return self.text_edit.toPlainText()

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


# ---------------------- Priekšskatījuma ģenerēšana fonā (lai DOCX/XLSX nekad neuzkar UI) ----------------------
class _PreviewBuildWorker(QObject):
    """Ģenerē priekšskatījuma PDF un pārvērš to PNG baitos fonā.

    Svarīgi: QPixmap nedrīkst veidot fonā, tāpēc worker atgriež PNG baitu sarakstu.
    """

    finished = Signal(str, list, int)  # data_hash, png_bytes_list, old_page
    failed = Signal(str, str)  # data_hash, error_message

    def __init__(self, d: 'AktaDati', data_hash: str, old_page: int):
        super().__init__()
        # --- Settings (persist across restarts) ---
        self._qt_settings = QSettings("AktaGenerators", "AktaGeneratorsApp")
        self._settings = load_settings()

        self._d = d
        self._hash = data_hash
        self._old_page = old_page

    def run(self):
        temp_pdf_path = None
        try:
            # Priekšskatījumā iekļaujam arī atsauces dokumentus (DOCX/XLSX u.c.), bet fonā, lai UI neuzkar.
            temp_pdf_path = ģenerēt_pdf(self._d, pdf_ceļš=None, include_reference_docs=True, encrypt_pdf=False)

            poppler_path_to_use = self._d.poppler_path if getattr(self._d, 'poppler_path', None) and os.path.exists(self._d.poppler_path) else None
            images = convert_from_path(temp_pdf_path, poppler_path=poppler_path_to_use)

            out_bytes = []
            from io import BytesIO
            for pil_img in images:
                bio = BytesIO()
                pil_img.save(bio, format='PNG')
                out_bytes.append(bio.getvalue())

            self.finished.emit(self._hash, out_bytes, self._old_page)
        except Exception as e:
            self.failed.emit(self._hash, str(e))
        finally:
            if temp_pdf_path and os.path.exists(temp_pdf_path):
                try:
                    os.remove(temp_pdf_path)
                except Exception:
                    pass


class AktaLogs(QMainWindow):
    def _parse_num_lv_en(self, value, default=0.0):
        try:
            s = str(value or '').strip()
            if not s:
                return float(default)
            s = s.replace(' ', '').replace(' ', '')
            if ',' in s and '.' in s:
                if s.rfind(',') > s.rfind('.'):
                    s = s.replace('.', '').replace(',', '.')
                else:
                    s = s.replace(',', '')
            else:
                if ',' in s:
                    s = s.replace(',', '.')
            filtered = ''.join(ch for ch in s if ch in '0123456789.-')
            if filtered in ('', '-', '.', '-.'):
                return float(default)
            return float(filtered)
        except Exception:
            return float(default)

    def __init__(self):
        super().__init__()


        # Ikona logam + Taskbar (Windows)
        try:
            self.setWindowIcon(QIcon(resource_path("Akta_Generators_Icon.ico")))
        except Exception:
            pass

        # --- JAUNS: Audit + Undo/Redo ---
        self._current_user = os.getenv("USERNAME") or os.getenv("USER") or ""
        self._audit_logger = AuditLogger(os.path.join(APP_DATA_DIR, "audit_log.jsonl"))
        self._undo_mgr = UndoRedoManager(max_steps=80)

        # Ātri undo/redo (Ctrl+Z / Ctrl+Y)
        self._act_undo = QAction("Undo", self)
        self._act_undo.setShortcut("Ctrl+Z")
        self._act_undo.triggered.connect(self.undo_action)

        self._act_redo = QAction("Redo", self)
        self._act_redo.setShortcut("Ctrl+Y")
        self._act_redo.triggered.connect(self.redo_action)

        self.addAction(self._act_undo)
        self.addAction(self._act_redo)

        # Status bar: Undo/Redo indikatori
        try:
            sb = self.statusBar()
            self._undo_status_label = QLabel("Undo: 0")
            self._redo_status_label = QLabel("Redo: 0")
            self._undo_status_label.setMinimumWidth(90)
            self._redo_status_label.setMinimumWidth(90)
            sb.addPermanentWidget(self._undo_status_label)
            sb.addPermanentWidget(self._redo_status_label)
        except Exception:
            self._undo_status_label = None
            self._redo_status_label = None

        # Izmaiņu izsekošana laukiem (audit + undo checkpoints)
        self._track_enabled = True
        self._last_widget_values = {}
        try:
            app = QApplication.instance()
            if app is not None:
                app.installEventFilter(self)
        except Exception:
            pass

        self._update_undo_redo_indicators()
        self.setWindowTitle("Pieņemšanas–Nodošanas akta ģenerators")
        self.resize(1200, 800)
        self.poppler_path = ""
        self.zoom_factor = 1.0
        self.history = [] # Inicializējam tukšu sarakstu
        self.address_book = {} # Inicializējam tukšu vārdnīcu
        self.text_block_manager = TextBlockManager() # JAUNA RINDAS
        self.data = AktaDati(
            datums=datetime.now().strftime('%Y-%m-%d'),
            pieņemšanas_datums=datetime.now().strftime('%Y-%m-%d'),
            nodošanas_datums=datetime.now().strftime('%Y-%m-%d'),
            izpildes_termiņš=(datetime.now() + timedelta(days=5)).strftime('%Y-%m-%d'),
        )

        # Noliktava (jauns)
        self._noliktava = NoliktavaDB(os.path.join(APP_DATA_DIR, "noliktava.json"))

        self._ceļš_projekts = None
        # Jauni atribūti ātruma uzlabošanai
        self.preview_timer = QTimer(self)
        self.preview_timer.setSingleShot(True)
        self._map_address_target = None  # QLineEdit, kurā ieliekam adresi no kartes
        self.preview_timer.timeout.connect(self._do_update_preview)
        self.preview_cache = {}  # Kešatmiņa: {'data_hash': {'images': [...], 'page_count': int}}
        self.last_data_hash = None  # Pēdējais datu hash
        self._autosave_path = os.path.join(APP_DATA_DIR, 'autosave_last_session.json')
        self._last_autosave_hash = None
        self.autosave_timer = QTimer(self)
        self.autosave_timer.setInterval(30000)
        self.autosave_timer.timeout.connect(self._autosave_snapshot)
        self.autosave_timer.start()

        # Priekšskatījuma ģenerēšana fonā (DOCX/XLSX konvertācija u.c.)
        self._preview_thread: Optional[QThread] = None
        self._preview_worker: Optional[_PreviewBuildWorker] = None
        self._requested_preview_hash: Optional[str] = None
        self._pending_preview_request = None  # (AktaDati, data_hash, old_page)

        self.tabs = QTabWidget()
        try:
            self.tabs.setUsesScrollButtons(True)
            self.tabs.tabBar().setExpanding(False)
            self.tabs.tabBar().setElideMode(Qt.ElideNone)
            self.tabs.setDocumentMode(False)
        except Exception:
            pass

        self._būvēt_pamata_tab()
        self._būvēt_puses_tab()
        self._būvēt_pozīcijas_tab()
        self._būvēt_noliktava_tab()  # Jauns: noliktavas sistēma
        self._būvēt_attēli_tab()
        self._būvēt_iestatījumi_tab()
        self._būvēt_papildu_iestatījumi_tab()
        self._būvēt_sablonu_tab()
        self._būvēt_adresu_gramata_tab()
        self._būvēt_audit_tab()
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



        # --- FIRST RUN: automātiski ielādē iebūvēto šablonu "Testa dati (piemērs)" tikai pirmajā palaišanas reizē ---
        QTimer.singleShot(0, self._auto_load_test_template_first_run)
        main_splitter = QSplitter(Qt.Horizontal)
        self.setCentralWidget(main_splitter)

        main_splitter.addWidget(self.tabs)

        preview_widget = QWidget()
        preview_layout = QVBoxLayout(preview_widget)
        self.preview_label = QLabel("PDF priekšskatījums")
        self.preview_label.setAlignment(Qt.AlignCenter)
        self.preview_label.setMinimumSize(1, 1)
        self.preview_label.setAttribute(Qt.WA_TransparentForMouseEvents, True)

        # Pannable scroll area, lai var vilkt (hand-drag) arī tukšajā zonā
        self.preview_scroll_area = PannableScrollArea()
        self.preview_scroll_area.setWidgetResizable(False)
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
        self._būvēt_toolbar()

        if self.data.auto_generate_akta_nr:
            self._generate_akta_nr()

    # ----- Menu -----

    def _auto_load_test_template_first_run(self):
        """Pirmajā palaišanas reizē automātiski ielādē iebūvēto šablonu 'Testa dati (piemērs)'.

        Mehānisms izmanto settings.json karodziņu 'first_run_test_template_loaded', lai ielāde notiktu tikai 1x.
        Ja šablonu saraksts vēl nav gatavs, karodziņš netiek uzstādīts un ielāde tiks mēģināta nākamajā startā.
        """
        try:
            st = load_settings() or {}
            if st.get("first_run_test_template_loaded", False):
                return

            target_name = "Testa dati (piemērs)"
            target_item = None

            # Šabloni ir listē (sablonu_list), ko programma jau aizpilda ar _update_sablonu_list()
            if hasattr(self, "sablonu_list") and self.sablonu_list is not None:
                for i in range(self.sablonu_list.count()):
                    it = self.sablonu_list.item(i)
                    if it and it.text().strip() == target_name:
                        target_item = it
                        break

            if target_item is None:
                return

            # Ielādējam tāpat kā lietotājs manuāli (dubultklikšķis uz šablona)
            self.ieladet_sablonu(target_item)

            # Atzīmējam, ka auto-load ir izpildīts
            st["first_run_test_template_loaded"] = True
            save_settings(st)

            # Atsvaidzinām priekšskatījumu (ja preview jau ir uzbūvēts)
            try:
                self._update_preview()
            except Exception:
                pass

        except Exception as e:
            # Nekad nekritinām aplikāciju auto-load dēļ
            print(f"Auto-load testa šablonam neizdevās: {e}")

    def _būvēt_menu(self):
        menubar = self.menuBar()
        fajls = menubar.addMenu("&Fails")

        # Parakstīšana (eParaksts)
        act_sign_now = QAction("Ģenerēt + Parakstīt (eParaksts)…", self)
        act_sign_now.triggered.connect(self.generate_and_sign_current)

        act_sign_last = QAction("Parakstīt pēdējo PDF…", self)
        act_sign_last.triggered.connect(lambda: self.sign_file_with_eparaksts(None))

        fajls.addSeparator()
        fajls.addAction(act_sign_now)
        fajls.addAction(act_sign_last)

        iestatijumi = menubar.addMenu("&Iestatījumi")
        act_set_ep = QAction("eParaksts…", self)
        act_set_ep.triggered.connect(self._open_settings_eparaksts)
        iestatijumi.addAction(act_set_ep)

        rediget = menubar.addMenu("&Rediģēt")
        rediget.addAction(self._act_undo)
        rediget.addAction(self._act_redo)


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

        # Papildu izvēlnes
        riki = menubar.addMenu("&Rīki")

        atjaunot = QAction("Atjaunot priekšskatījumu", self)
        atjaunot.setShortcut("F5")
        atjaunot.triggered.connect(self._update_preview)
        riki.addAction(atjaunot)

        atvert_iestat = QAction("Atvērt iestatījumu mapi", self)
        atvert_iestat.triggered.connect(self._open_settings_folder)
        riki.addAction(atvert_iestat)

        riki.addSeparator()
        notirit_kesu = QAction("Notīrīt priekšskatījuma kešu", self)
        notirit_kesu.triggered.connect(self._clear_preview_cache)
        riki.addAction(notirit_kesu)

        palidziba = menubar.addMenu("&Palīdzība")
        par = QAction("Par programmu", self)
        par.triggered.connect(self._show_about_dialog)
        palidziba.addAction(par)

        isceli = QAction("Īsceļi", self)
        isceli.triggered.connect(self._show_shortcuts_dialog)
        palidziba.addAction(isceli)

        palidziba.addSeparator()
        atvert_map = QAction("Atvērt programmas mapi", self)
        atvert_map.triggered.connect(self._open_app_folder)
        palidziba.addAction(atvert_map)


    # ----- Palīgfunkcijas izvēlnēm -----
    def _open_settings_folder(self):
        try:
            os.makedirs(SETTINGS_DIR, exist_ok=True)
            os.startfile(SETTINGS_DIR)  # type: ignore[attr-defined]
        except Exception:
            try:
                import subprocess
                if sys.platform.startswith("darwin"):
                    subprocess.Popen(["open", SETTINGS_DIR])
                else:
                    subprocess.Popen(["xdg-open", SETTINGS_DIR])
            except Exception:
                QMessageBox.information(self, "Iestatījumi", f"Iestatījumu mape:\n{SETTINGS_DIR}")

    def _open_app_folder(self):
        try:
            app_dir = os.path.dirname(os.path.abspath(__file__))
            os.startfile(app_dir)  # type: ignore[attr-defined]
        except Exception:
            try:
                import subprocess
                app_dir = os.path.dirname(os.path.abspath(__file__))
                if sys.platform.startswith("darwin"):
                    subprocess.Popen(["open", app_dir])
                else:
                    subprocess.Popen(["xdg-open", app_dir])
            except Exception:
                QMessageBox.information(self, "Mape", os.path.dirname(os.path.abspath(__file__)))

    def _clear_preview_cache(self):
        try:
            if hasattr(self, "preview_cache") and isinstance(self.preview_cache, dict):
                self.preview_cache.clear()
            if hasattr(self, "last_data_hash"):
                self.last_data_hash = None
            self.statusBar().showMessage("Priekšskatījuma kešs notīrīts", 3000)
        except Exception:
            pass
        self._update_preview()

    def _show_about_dialog(self):
        about_txt = (
            "Akta ģenerators\n"
            "\n"
            f"Iestatījumu mape: {SETTINGS_DIR}\n"
            "\n"
            "Šī programma palīdz veidot pieņemšanas–nodošanas aktus un eksportēt PDF/DOCX."
        )
        QMessageBox.information(self, "Par programmu", about_txt)

    def _show_shortcuts_dialog(self):
        shortcuts_txt = (
            "Īsceļi\n"
            "\n"
            "F5      — atjaunot priekšskatījumu\n"
            "Ctrl+S  — saglabāt projektu\n"
            "Ctrl+O  — ielādēt projektu\n"
            "Ctrl+P  — ģenerēt PDF\n"
        )
        QMessageBox.information(self, "Īsceļi", shortcuts_txt)

    def _būvēt_toolbar(self):
        tb = self.addToolBar("Galvenais")
        tb.setMovable(False)

        act_save = QAction("Saglabāt", self)
        act_save.setShortcut("Ctrl+S")
        act_save.triggered.connect(self.saglabat_projektu)
        tb.addAction(act_save)

        act_load = QAction("Ielādēt", self)
        act_load.setShortcut("Ctrl+O")
        act_load.triggered.connect(self.ieladet_projektu)
        tb.addAction(act_load)

        tb.addSeparator()

        act_pdf = QAction("PDF", self)
        act_pdf.setShortcut("Ctrl+P")
        act_pdf.triggered.connect(self.ģenerēt_pdf_dialogs)
        tb.addAction(act_pdf)

        act_docx = QAction("DOCX", self)
        act_docx.triggered.connect(self.ģenerēt_docx_dialogs)
        tb.addAction(act_docx)

        tb.addSeparator()

        act_print = QAction("Drukāt", self)
        act_print.triggered.connect(self.drukāt_pdf_dialogs)
        tb.addAction(act_print)

        self.statusBar().showMessage("Gatavs")


    # ----- Tab: Pamata -----

    def _būvēt_pamata_tab(self):
        content_widget = QWidget()
        form = QFormLayout()
        # --- Dokumenta tips / nosaukums (jauns) ---
        self.cmb_doc_tips = QComboBox()
        for _key, _cfg in DOCUMENT_TYPE_PRESETS.items():
            self.cmb_doc_tips.addItem(_cfg.get("title", _key), _key)

        self.cb_party_mode = QComboBox()
        self.cb_party_mode.addItem(PARTY_MODE_LABELS["puses"], "puses")
        self.cb_party_mode.addItem(PARTY_MODE_LABELS["rekviziti"], "rekviziti")
        self.cb_party_mode.addItem(PARTY_MODE_LABELS["abi"], "abi")
        self.cb_party_mode.addItem(PARTY_MODE_LABELS["neviens"], "neviens")

        self.in_doc_nosaukums = QLineEdit()
        self.in_doc_nosaukums.setPlaceholderText("Dokumenta nosaukums (var mainīt)")

        # Datumu lauki ar "ielādēt sistēmas datumu" pogu
        def _date_row_widget(date_edit: QDateEdit, attr_name: str = ""):
            wrap = QWidget()
            h = QHBoxLayout(wrap)
            h.setContentsMargins(0, 0, 0, 0)
            h.setSpacing(6)
            date_edit.setCalendarPopup(True)
            date_edit.setDisplayFormat("yyyy-MM-dd")
            btn_today = QPushButton("Ielādēt sistēmas datumu")
            btn_today.setToolTip("Ielādēt sistēmas datumu")
            btn_today.setMinimumWidth(180)
            def _set_today():
                try:
                    date_edit.setDate(QDate.currentDate())
                    date_edit.dateChanged.emit(date_edit.date())
                except Exception:
                    try:
                        date_edit.setDate(QDate.currentDate())
                    except Exception:
                        pass
                try:
                    self._update_preview()
                except Exception:
                    pass
            btn_today.clicked.connect(_set_today)
            if attr_name:
                try:
                    setattr(self, attr_name, btn_today)
                except Exception:
                    pass
            h.addWidget(date_edit, 1)
            h.addWidget(btn_today, 0)
            return wrap

        self.in_apmaksas_termins = QDateEdit()
        self.in_piegades_datums = QDateEdit()
        self.w_apmaksas_termins = _date_row_widget(self.in_apmaksas_termins, "btn_apmaksas_today")
        self.w_piegades_datums = _date_row_widget(self.in_piegades_datums, "btn_piegades_today")

        # Noklusējumi
        try:
            self.in_doc_nosaukums.setText(self.cmb_doc_tips.currentText())
            self.in_apmaksas_termins.setDate(QDate.currentDate())
            self.in_piegades_datums.setDate(QDate.currentDate())
        except Exception:
            pass

        # Rādām/slēpjam laukus atkarībā no dokumenta tipa + atjaunojam nosaukumu
        def _sync_cover_title_with_doc_title(force: bool = False):
            try:
                if not hasattr(self, 'in_cover_page_title'):
                    return
                current_cover_title = (self.in_cover_page_title.text() or "").strip()
                current_doc_title = (self.in_doc_nosaukums.text() or "").strip() or (self.cmb_doc_tips.currentText() or "Dokuments")
                auto_titles = {"", "Pieņemšanas-Nodošanas Akts", "Pieņemšanas–Nodošanas akts", "Pieņemšanas-Nodošanas akts", "Pavadzīme", "Rēķins"}
                if force or current_cover_title in auto_titles:
                    self.in_cover_page_title.setText(current_doc_title)
            except Exception:
                pass

        def _apply_doc_tips_ui():
            try:
                tip = self.cmb_doc_tips.currentData() or "akta"
                cur_title = (self.in_doc_nosaukums.text() or "").strip()
                preset = DOCUMENT_TYPE_PRESETS.get(tip, {})
                default_titles = [str(v.get("title", "")) for v in DOCUMENT_TYPE_PRESETS.values()]
                if (not cur_title) or (cur_title in default_titles):
                    self.in_doc_nosaukums.setText(self.cmb_doc_tips.currentText())

                default_party_mode = preset.get("party_mode", "puses")
                self.cb_party_mode.setCurrentIndex(max(0, self.cb_party_mode.findData(default_party_mode)))
                self.w_apmaksas_termins.setVisible(tip == "rekins")
                self.w_piegades_datums.setVisible(tip == "pavadzime")
                self._apply_party_mode_ui()
                self._update_party_role_titles()
                _sync_cover_title_with_doc_title(force=True)
            except Exception:
                pass

        self.cmb_doc_tips.currentIndexChanged.connect(_apply_doc_tips_ui)
        self.cb_party_mode.currentIndexChanged.connect(self._apply_party_mode_ui)
        self.in_doc_nosaukums.textChanged.connect(lambda _t: _sync_cover_title_with_doc_title())
        _apply_doc_tips_ui()

        form.addRow("Dokumenta tips", self.cmb_doc_tips)
        form.addRow("Sadaļu režīms", self.cb_party_mode)
        form.addRow("Dokumenta nosaukums", self.in_doc_nosaukums)
        form.addRow("Apmaksas termiņš (rēķinam)", self.w_apmaksas_termins)
        form.addRow("Piegādes datums (pavadzīmei)", self.w_piegades_datums)

        self.in_akta_nr = QLineEdit()
        self.btn_generate_akta_nr = QPushButton("Ģenerēt Nr.")
        self.btn_generate_akta_nr.clicked.connect(self._generate_akta_nr)
        # Mazs, neuzkrītošs "restart" taustiņš akta numura skaitītājam
        self.btn_reset_akta_nr = QPushButton("↺")
        self.btn_reset_akta_nr.setToolTip("Pārstartēt akta numura skaitītāju (atgriezt uz 1)")
        self.btn_reset_akta_nr.setFixedSize(34, 34)
        self.btn_reset_akta_nr.setFlat(True)
        self.btn_reset_akta_nr.clicked.connect(self._reset_akta_nr_counter)
        akta_nr_layout = QHBoxLayout()
        akta_nr_layout.addWidget(self.in_akta_nr)
        akta_nr_layout.addWidget(self.btn_generate_akta_nr)
        akta_nr_layout.addWidget(self.btn_reset_akta_nr)
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
        self.in_izpildes_termins.setDate((datetime.now() + timedelta(days=int(getattr(self.data, 'default_execution_days', 5)))).date())  # default: today + N days
        self.in_izpildes_termins.setSpecialValueText("Nav norādīts") # Text for null date
        self.in_izpildes_termins.setCalendarPopup(True)
        self.ck_ieklaut_izpildes_terminu = QCheckBox("Iekļaut")
        self.ck_ieklaut_izpildes_terminu.setChecked(True)

        self.in_pieņemšanas_datums = QDateEdit(calendarPopup=True)
        self.in_pieņemšanas_datums.setDisplayFormat("yyyy-MM-dd")
        self.in_pieņemšanas_datums.setMinimumDate(datetime(1900, 1, 1).date())
        self.in_pieņemšanas_datums.setMaximumDate(datetime(2100, 1, 1).date())
        self.in_pieņemšanas_datums.setDate(datetime.now().date())
        self.in_pieņemšanas_datums.setSpecialValueText("Nav norādīts")
        self.in_pieņemšanas_datums.setCalendarPopup(True)
        self.ck_ieklaut_pienemsanas_datumu = QCheckBox("Iekļaut")
        self.ck_ieklaut_pienemsanas_datumu.setChecked(True)

        self.in_nodošanas_datums = QDateEdit(calendarPopup=True)
        self.in_nodošanas_datums.setDisplayFormat("yyyy-MM-dd")
        self.in_nodošanas_datums.setMinimumDate(datetime(1900, 1, 1).date())
        self.in_nodošanas_datums.setMaximumDate(datetime(2100, 1, 1).date())
        self.in_nodošanas_datums.setDate(datetime.now().date())
        self.in_nodošanas_datums.setSpecialValueText("Nav norādīts")
        self.in_nodošanas_datums.setCalendarPopup(True)
        self.ck_ieklaut_nodosanas_datumu = QCheckBox("Iekļaut")
        self.ck_ieklaut_nodosanas_datumu.setChecked(True)

        def _make_optional_date_widget(date_edit, include_checkbox):
            wrap = QWidget()
            lay = QHBoxLayout(wrap)
            lay.setContentsMargins(0, 0, 0, 0)
            lay.setSpacing(8)
            lay.addWidget(self._wrap_date_with_system_button(date_edit), 1)
            lay.addWidget(include_checkbox, 0)
            def _sync_optional_date_state(checked):
                try:
                    date_edit.setEnabled(bool(checked))
                except Exception:
                    pass
            include_checkbox.toggled.connect(_sync_optional_date_state)
            _sync_optional_date_state(include_checkbox.isChecked())
            return wrap

        self.w_izpildes_termins_optional = _make_optional_date_widget(self.in_izpildes_termins, self.ck_ieklaut_izpildes_terminu)
        self.w_pienemsanas_datums_optional = _make_optional_date_widget(self.in_pieņemšanas_datums, self.ck_ieklaut_pienemsanas_datumu)
        self.w_nodosanas_datums_optional = _make_optional_date_widget(self.in_nodošanas_datums, self.ck_ieklaut_nodosanas_datumu)

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
        self.list_atsauces_faili = QListWidget()
        self.list_atsauces_faili.setMinimumHeight(90)
        self.list_atsauces_faili.setSelectionMode(QAbstractItemView.SingleSelection)
        self.btn_add_atsauce_failu = QPushButton("Pievienot atsauces failu")
        self.btn_remove_atsauce_failu = QPushButton("Noņemt izvēlēto")
        self.btn_scale_atsauce_failu = QPushButton("Mērogs %")
        self.btn_add_atsauce_failu.clicked.connect(self._add_reference_doc)
        self.btn_remove_atsauce_failu.clicked.connect(self._remove_reference_doc)
        self.btn_scale_atsauce_failu.clicked.connect(self._set_reference_doc_scale)
        try:
            _refs_model = self.list_atsauces_faili.model()
            _refs_model.rowsInserted.connect(self._invalidate_preview_cache_and_refresh)
            _refs_model.rowsRemoved.connect(self._invalidate_preview_cache_and_refresh)
            _refs_model.dataChanged.connect(self._invalidate_preview_cache_and_refresh)
            _refs_model.layoutChanged.connect(self._invalidate_preview_cache_and_refresh)
        except Exception:
            pass

        self.cb_akta_statuss = QComboBox()
        self.cb_akta_statuss.addItems(["Melnraksts", "Apstiprināts", "Parakstīts", "Arhivēts", "Atcelts"])
        self.in_valuta = QComboBox()
        self.in_valuta.setEditable(True)
        self.in_valuta.addItems(["EUR €","USD $","GBP £","NOK kr","SEK kr","DKK kr","PLN zł","CHF CHF","CZK Kč","HUF Ft","RON lei","BGN лв","JPY ¥","CNY ¥","AUD $","CAD $","NZD $","TRY ₺","UAH ₴","RUB ₽"]) 
        self.in_valuta.setCurrentText("EUR €")

        # Piezīmes
        self.in_piezimes = TextBlockTextEdit(self.text_block_manager, "piezimes")

        self.ck_elektroniskais_paraksts = QCheckBox("Elektroniskais paraksts (ignorē fiziskos parakstus)")
        self.ck_radit_elektronisko_parakstu_tekstu = QCheckBox("Rādīt elektroniskā paraksta tekstu PDF dokumentā") # JAUNA RŪTIŅA

        form.addRow("Akta Nr.", akta_nr_widget)
        form.addRow("Datums", self._wrap_date_with_system_button(self.in_datums))
        form.addRow("Vieta", self.in_vieta)
        form.addRow("Pasūtījuma Nr.", self.in_pas_nr)
        form.addRow("Līguma Nr.", self.in_liguma_nr)
        form.addRow("Izpildes termiņš", self.w_izpildes_termins_optional)
        form.addRow("Pieņemšanas datums", self.w_pienemsanas_datums_optional)
        form.addRow("Nodošanas datums", self.w_nodosanas_datums_optional)
        form.addRow("Strīdu risināšana", self.in_strīdu_risināšana)  # Tagad tieši izmantojam jauno objektu
        form.addRow(self.ck_konfidencialitate)
        form.addRow("Soda nauda (%)", self.in_soda_nauda_procenti)
        form.addRow("Piegādes nosacījumi", self.in_piegades_nosacijumi)  # Tagad tieši izmantojam jauno objektu
        form.addRow(self.ck_apdrošināšana)

        # Apdrošināšanas teksts (kā teksta bloks – līdzīgi kā citur)
        self.in_apdrosinasana_teksts = TextBlockTextEdit(self.text_block_manager, "apdrosinasana_teksts")
        self.in_apdrosinasana_teksts.setPlaceholderText("Ierakstiet apdrošināšanas tekstu…")
        form.addRow("Apdrošināšanas teksts", self.in_apdrosinasana_teksts)

        # Rādīt/Slēpt teksta lauku atkarībā no atzīmes
        def _toggle_apdrosinasana_text():
            show = self.ck_apdrošināšana.isChecked()
            self.in_apdrosinasana_teksts.setVisible(show)
            # QLabel no QFormLayout nav tieši pieejams; paslēpjam arī etiķeti
            try:
                lbl = form.labelForField(self.in_apdrosinasana_teksts)
                if lbl:
                    lbl.setVisible(show)
            except Exception:
                pass
        self.ck_apdrošināšana.stateChanged.connect(lambda *_: (_toggle_apdrosinasana_text(), self._update_preview()))
        _toggle_apdrosinasana_text()
        form.addRow("Papildu nosacījumi", self.in_papildu_nosacijumi)  # Tagad tieši izmantojam jauno objektu
        form.addRow("Atsauces dokumenti", self.in_atsauces_dokumenti)  # Tagad tieši izmantojam jauno objektu
        # Atsauces faili (reāli pielikumi PDF beigās)
        w_refs = QWidget()
        v_refs = QVBoxLayout(w_refs)
        btn_row = QHBoxLayout()
        btn_row.addWidget(self.btn_add_atsauce_failu)
        btn_row.addWidget(self.btn_remove_atsauce_failu)
        btn_row.addWidget(self.btn_scale_atsauce_failu)
        btn_row.addStretch(1)
        v_refs.addLayout(btn_row)
        v_refs.addWidget(self.list_atsauces_faili)
        form.addRow("Atsauces faili", w_refs)
        form.addRow("Akta statuss", self.cb_akta_statuss)
        
        # --- JAUNS: PDF šifrēšana (parole) Pamata datos ---
        # Priekšskatījums programmā joprojām tiek ģenerēts NEšifrēts (skat. encrypt_pdf=False preview worker),
        # bet eksportā/saglabāšanā šī parole tiks piemērota PDF failam.
        self.ck_pdf_encrypt_basic = QCheckBox("Šifrēt PDF ar paroli")
        try:
            self.ck_pdf_encrypt_basic.setChecked(bool(getattr(self.data, "enable_pdf_encryption", False)))
        except Exception:
            self.ck_pdf_encrypt_basic.setChecked(False)

        self.in_pdf_password_basic = QLineEdit()
        self.in_pdf_password_basic.setEchoMode(QLineEdit.Password)
        self.in_pdf_password_basic.setPlaceholderText("Parole (ja atstāj tukšu – atvērsies bez paroles)")
        try:
            self.in_pdf_password_basic.setText(str(getattr(self.data, "pdf_user_password", "") or ""))
        except Exception:
            pass

        # Neliekam spiest preview uzreiz – bet, ja lietotājs maina, lai eksportā viss būtu saglabāts.
        try:
            self.ck_pdf_encrypt_basic.stateChanged.connect(self._update_preview)
            self.in_pdf_password_basic.textChanged.connect(self._update_preview)
        except Exception:
            pass

        form.addRow(self.ck_pdf_encrypt_basic)
        form.addRow("PDF parole", self.in_pdf_password_basic)

        form.addRow("Valūta", self.in_valuta)
        form.addRow("Piezīmes", self.in_piezimes)  # Tagad tieši izmantojam jauno objektu
        form.addRow("Elektroniskais paraksts", self.ck_elektroniskais_paraksts)
        form.addRow("", self.ck_radit_elektronisko_parakstu_tekstu) # JAUNA RŪTIŅA FORMĀ
        # --- JAUNS: QR iestatījumu UI (droši, ja nav izveidots) ---
        if not hasattr(self, "ck_qr_kods"):
            self.ck_qr_kods = QCheckBox("QR kods apakšā pa kreisi (akta dati)")
            self.ck_qr_kods.setChecked(True)
        if not hasattr(self, "ck_qr_first_page"):
            self.ck_qr_first_page = QCheckBox("QR tikai pirmajā lapā")
            self.ck_qr_first_page.setChecked(True)
        if not hasattr(self, "ck_qr_url_mode"):
            self.ck_qr_url_mode = QCheckBox("QR kā verifikācijas saite (URL)")
            self.ck_qr_url_mode.setChecked(False)
        if not hasattr(self, "le_qr_url"):
            self.le_qr_url = QLineEdit()
            self.le_qr_url.setPlaceholderText("QR URL, piem.: https://kulinics.id.lv/verify")
        form.addRow("", self.ck_qr_kods)
        form.addRow("", self.ck_qr_first_page)
        form.addRow("", self.ck_qr_url_mode)
        form.addRow("QR URL", self.le_qr_url)
        # --- JAUNS: QR UI elementi (izveidojam, ja nav) ---
        if not hasattr(self, "ck_qr_kods"):
            self.ck_qr_kods = QCheckBox("QR kods apakšā pa kreisi (akta dati)")
            self.ck_qr_kods.setChecked(True)
        if not hasattr(self, "ck_qr_only_first"):
            self.ck_qr_only_first = QCheckBox("QR tikai pirmajā lapā")
            self.ck_qr_only_first.setChecked(True)
        if not hasattr(self, "ck_qr_url_mode"):
            self.ck_qr_url_mode = QCheckBox("QR kā verifikācijas saite (URL)")
            self.ck_qr_url_mode.setChecked(False)
        if not hasattr(self, "le_qr_url"):
            self.le_qr_url = QLineEdit()
            self.le_qr_url.setPlaceholderText("QR URL, piem.: https://kulinics.id.lv/verify")
        form.addRow("", self.ck_qr_only_first)
        # --- FIX: QR URL checkbox alias (dažādi nosaukumi) ---
        if not hasattr(self, "ck_qr_use_url") and hasattr(self, "ck_qr_url_mode"):
            self.ck_qr_use_url = self.ck_qr_url_mode
        elif not hasattr(self, "ck_qr_url_mode") and hasattr(self, "ck_qr_use_url"):
            self.ck_qr_url_mode = self.ck_qr_use_url
        elif not hasattr(self, "ck_qr_use_url") and not hasattr(self, "ck_qr_url_mode"):
            self.ck_qr_use_url = QCheckBox("QR kā verifikācijas saite (URL)")
            self.ck_qr_use_url.setChecked(False)
            self.ck_qr_url_mode = self.ck_qr_use_url
        form.addRow("", self.ck_qr_use_url)
        # --- FIX: QR URL lineEdit alias (dažādi nosaukumi) ---
        if not hasattr(self, "le_qr_base_url") and hasattr(self, "le_qr_url"):
            self.le_qr_base_url = self.le_qr_url
        elif not hasattr(self, "le_qr_url") and hasattr(self, "le_qr_base_url"):
            self.le_qr_url = self.le_qr_base_url
        elif not hasattr(self, "le_qr_base_url") and not hasattr(self, "le_qr_url"):
            self.le_qr_base_url = QLineEdit()
            self.le_qr_base_url.setPlaceholderText("QR URL, piem.: https://kulinics.id.lv/verify")
            self.le_qr_url = self.le_qr_base_url
        form.addRow("QR URL", self.le_qr_base_url)

        content_widget.setLayout(form)

        scroll_area = QScrollArea()
        scroll_area.setWidgetResizable(True)
        try:
            scroll_area.setAlignment(Qt.AlignTop | Qt.AlignLeft)
        except Exception:
            pass
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
        self.in_valuta.currentTextChanged.connect(self._update_preview)
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
        """Ģenerē jaunu unikālu akta numuru pēc shēmas PP-YYYY-NNNN.

        Uzlabojums:
        - uztur stabilu secīgo skaitītāju failā (AKTA_NR_COUNTER_FILE), lai nākamais numurs vienmēr būtu
          konsekvents pat tad, ja projekts vēl nav saglabāts vai history nav atjaunots.
        - papildus pārbauda jau izmantotos numurus projektos un vēsturē, lai nerastos dublikāti.
        """
        try:
            prefix = "PP"
            current_year = datetime.now().strftime('%Y')

            # ---------------- 1) Nolasa jau izmantotos numurus ----------------
            used_ids = set()

            # 1) Projekti (JSON) mapē
            if os.path.isdir(PROJECT_SAVE_DIR):
                for filename in os.listdir(PROJECT_SAVE_DIR):
                    if not filename.lower().endswith('.json'):
                        continue
                    fp = os.path.join(PROJECT_SAVE_DIR, filename)
                    try:
                        with open(fp, 'r', encoding='utf-8') as f:
                            project_data = json.load(f)
                        akta_nr = project_data.get('akta_nr') if isinstance(project_data, dict) else None
                        if akta_nr:
                            used_ids.add(str(akta_nr))
                    except Exception:
                        pass  # klusām ignorējam bojātus failus

            # 2) Vēsture (history.json)
            if os.path.exists(HISTORY_FILE):
                try:
                    with open(HISTORY_FILE, 'r', encoding='utf-8') as f:
                        history = json.load(f)
                except Exception:
                    history = None

                if isinstance(history, dict):
                    items = history.get('items') or history.get('history') or []
                else:
                    items = history or []

                if isinstance(items, list):
                    for item in items:
                        try:
                            if isinstance(item, dict):
                                ak = item.get('akta_nr')
                                if ak:
                                    used_ids.add(str(ak))

                                for k in ('json_path', 'project_path', 'path'):
                                    p = item.get(k)
                                    if isinstance(p, str) and p.lower().endswith('.json') and os.path.exists(p):
                                        try:
                                            with open(p, 'r', encoding='utf-8') as f:
                                                pd = json.load(f)
                                            ak2 = pd.get('akta_nr') if isinstance(pd, dict) else None
                                            if ak2:
                                                used_ids.add(str(ak2))
                                        except Exception:
                                            pass
                            elif isinstance(item, str) and item.lower().endswith('.json') and os.path.exists(item):
                                try:
                                    with open(item, 'r', encoding='utf-8') as f:
                                        pd = json.load(f)
                                    ak2 = pd.get('akta_nr') if isinstance(pd, dict) else None
                                    if ak2:
                                        used_ids.add(str(ak2))
                                except Exception:
                                    pass
                        except Exception:
                            pass

            # ---------------- 2) Aprēķina MAX izmantoto secību šim gadam ----------------
            seq_re = re.compile(rf'^{re.escape(prefix)}-(\d{{4}})-(\d{{4}})$')
            max_used_seq_this_year = 0
            for ak in used_ids:
                mm = seq_re.match(str(ak).strip())
                if not mm:
                    continue
                year, seq_s = mm.group(1), mm.group(2)
                if year != current_year:
                    continue
                try:
                    max_used_seq_this_year = max(max_used_seq_this_year, int(seq_s))
                except Exception:
                    pass

            # ---------------- 3) Nolasa / atjauno skaitītāju failā ----------------
            counters = {}
            if os.path.exists(AKTA_NR_COUNTER_FILE):
                try:
                    with open(AKTA_NR_COUNTER_FILE, 'r', encoding='utf-8') as f:
                        counters = json.load(f) if f else {}
                except Exception:
                    counters = {}

            last_seq = 0
            try:
                last_seq = int(counters.get(current_year, 0))
            except Exception:
                last_seq = 0

            # nodrošinām, ka skaitītājs nekad nav mazāks par reāli izmantoto max
            base_seq = max(last_seq, max_used_seq_this_year)

            # ---------------- 4) Atrod nākamo brīvo secību (bez dublikātiem) ----------------
            new_seq = base_seq + 1
            while True:
                new_akta_nr = f"{prefix}-{current_year}-{new_seq:04d}"
                if new_akta_nr not in used_ids:
                    # Uzreiz saglabājam skaitītāju, lai nākamais ģenerējums būtu konsekvents
                    counters[current_year] = new_seq
                    try:
                        os.makedirs(os.path.dirname(AKTA_NR_COUNTER_FILE), exist_ok=True)
                        with open(AKTA_NR_COUNTER_FILE, 'w', encoding='utf-8') as f:
                            json.dump(counters, f, ensure_ascii=False, indent=2)
                    except Exception:
                        pass

                    if hasattr(self, 'in_akta_nr') and self.in_akta_nr:
                        self.in_akta_nr.setText(new_akta_nr)
                    break
                new_seq += 1

        except Exception as e:
            print(f"Akta nr. ģenerēšanas kļūda: {e}")

    def _reset_akta_nr_counter(self):
        """Pārstartē akta numura skaitītāju, lai nākamais numurs sākas no 1.

        Piezīme: šī darbība apzināti neanalizē jau izmantotos numurus projektos/vēsturē.
        Ja vēlies pilnīgu unikālitāti, izmanto 'Ģenerēt Nr.' pēc reset.
        """
        try:
            prefix = "PP"
            current_year = datetime.now().strftime('%Y')

            counters = {}
            if os.path.exists(AKTA_NR_COUNTER_FILE):
                try:
                    with open(AKTA_NR_COUNTER_FILE, 'r', encoding='utf-8') as f:
                        counters = json.load(f) if f else {}
                except Exception:
                    counters = {}

            # Atgriež uz 1 un saglabā
            counters[current_year] = 1
            try:
                os.makedirs(os.path.dirname(AKTA_NR_COUNTER_FILE), exist_ok=True)
                with open(AKTA_NR_COUNTER_FILE, 'w', encoding='utf-8') as f:
                    json.dump(counters, f, ensure_ascii=False, indent=2)
            except Exception:
                pass

            # Uzstāda lauku uz 0001
            new_akta_nr = f"{prefix}-{current_year}-0001"
            if hasattr(self, 'in_akta_nr') and self.in_akta_nr:
                self.in_akta_nr.setText(new_akta_nr)

            try:
                if hasattr(self, 'statusBar') and callable(self.statusBar):
                    self.statusBar().showMessage("Akta numura skaitītājs pārstartēts uz 1", 3000)
            except Exception:
                pass

        except Exception as e:
            print(f"Akta nr. skaitītāja reset kļūda: {e}")


    def savākt_datus(self) -> AktaDati:
        d = AktaDati()
        d.akta_nr = self.in_akta_nr.text().strip()
        d.datums = self.in_datums.date().toString("yyyy-MM-dd")
        d.vieta = self.in_vieta.text().strip()
        d.pasūtījuma_nr = self.in_pas_nr.text().strip()
        # Dokumenta tips / nosaukums (jauns)
        try:
            d.doc_tips = self.cmb_doc_tips.currentData() or "akta"
            d.dokumenta_nosaukums = (self.in_doc_nosaukums.text() or "").strip() or (self.cmb_doc_tips.currentText() if hasattr(self, 'cmb_doc_tips') else "")
            d.party_mode = self.cb_party_mode.currentData() or 'puses'
            d.apmaksas_termins = self.in_apmaksas_termins.date().toString("yyyy-MM-dd") if hasattr(self, 'in_apmaksas_termins') and not self.in_apmaksas_termins.date().isNull() else ""
            d.piegades_datums = self.in_piegades_datums.date().toString("yyyy-MM-dd") if hasattr(self, 'in_piegades_datums') and not self.in_piegades_datums.date().isNull() else ""
        except Exception:
            pass
        d.piezīmes = self.in_piezimes.toPlainText().strip()

        # JAUNA RINDAS
        d.templates_dir = self.in_templates_dir.text().strip()

        d.līguma_nr = self.in_liguma_nr.text().strip()
        d.ieklaut_izpildes_terminu = self.ck_ieklaut_izpildes_terminu.isChecked()
        d.ieklaut_pienemsanas_datumu = self.ck_ieklaut_pienemsanas_datumu.isChecked()
        d.ieklaut_nodosanas_datumu = self.ck_ieklaut_nodosanas_datumu.isChecked()
        d.izpildes_termiņš = self.in_izpildes_termins.date().toString("yyyy-MM-dd") if d.ieklaut_izpildes_terminu and not self.in_izpildes_termins.date().isNull() else ""
        d.pieņemšanas_datums = self.in_pieņemšanas_datums.date().toString("yyyy-MM-dd") if d.ieklaut_pienemsanas_datumu and not self.in_pieņemšanas_datums.date().isNull() else ""
        d.nodošanas_datums = self.in_nodošanas_datums.date().toString("yyyy-MM-dd") if d.ieklaut_nodosanas_datumu and not self.in_nodošanas_datums.date().isNull() else ""
        d.strīdu_risināšana = self.in_strīdu_risināšana.toPlainText().strip()  # Tagad izsauc toPlainText() uz pielāgotā logrīka
        d.konfidencialitātes_klauzula = self.ck_konfidencialitate.isChecked()
        d.soda_nauda_procenti = to_decimal(self.in_soda_nauda_procenti.value())
        d.piegādes_nosacījumi = self.in_piegades_nosacijumi.text().strip()  # Tagad izsauc text() uz pielāgotā logrīka
        d.apdrošināšana = self.ck_apdrošināšana.isChecked()
        d.apdrošināšana_teksts = self.in_apdrosinasana_teksts.toPlainText().strip() if hasattr(self, 'in_apdrosinasana_teksts') else ""
        d.papildu_nosacījumi = self.in_papildu_nosacijumi.toPlainText().strip()  # Tagad izsauc toPlainText() uz pielāgotā logrīka
        d.atsauces_dokumenti = self.in_atsauces_dokumenti.text().strip()  # Tagad izsauc text() uz pielāgotā logrīka
        d.atsauces_dokumenti_faili = []
        for i in range(self.list_atsauces_faili.count()):
            it = self.list_atsauces_faili.item(i)
            payload = _normalize_reference_doc_payload(it.data(Qt.UserRole), fallback_name=it.text())
            if payload.get('ceļš'):
                d.atsauces_dokumenti_faili.append(payload)

        d.akta_statuss = self.cb_akta_statuss.currentText()
        d.valūta = (self.in_valuta.currentText().strip().split(' ')[0] if self.in_valuta.currentText().strip() else "")
        d.elektroniskais_paraksts = self.ck_elektroniskais_paraksts.isChecked()
        d.radit_elektronisko_parakstu_tekstu = self.ck_radit_elektronisko_parakstu_tekstu.isChecked() # JAUNA RINDAS
        d.qr_kods_enabled = self.ck_qr_kods.isChecked() if hasattr(self, 'ck_qr_kods') else True
        d.qr_kods_tikai_pirma_lapa = self.ck_qr_only_first.isChecked() if hasattr(self, 'ck_qr_only_first') else True
        d.qr_kods_url_mode = (self.ck_qr_url_mode.isChecked() if hasattr(self, 'ck_qr_url_mode') else (self.ck_qr_use_url.isChecked() if hasattr(self, 'ck_qr_use_url') else False))
        d.qr_kods_url = (self.le_qr_url.text().strip() if hasattr(self, 'le_qr_url') else (self.le_qr_base_url.text().strip() if hasattr(self, 'le_qr_base_url') else ""))
        d.qr_only_first_page = self.ck_qr_only_first.isChecked() if hasattr(self, 'ck_qr_only_first') else False
        d.qr_verification_url_enabled = self.ck_qr_use_url.isChecked() if hasattr(self, 'ck_qr_use_url') else False
        d.qr_verification_base_url = (self.le_qr_base_url.text().strip() if hasattr(self, 'le_qr_base_url') else '')
        try:
            if hasattr(self, '_settings') and self._settings is not None:
                self._settings["qr_base_url"] = d.qr_verification_base_url
                save_settings(self._settings)
        except Exception:
            pass

        # Piezīmes
        d.piezīmes = self.in_piezimes.toPlainText().strip()  # Tagad izsauc toPlainText() uz pielāgotā logrīka

        # Puses / Rekvizīti
        d.party_mode = (self.cb_party_mode.currentData() if hasattr(self, 'cb_party_mode') else 'puses') or 'puses'
        d.rekvizitu_virsraksts = ((self.in_rekvizitu_virsraksts.text() if hasattr(self, 'in_rekvizitu_virsraksts') else 'Rekvizīti') or 'Rekvizīti').strip()
        d.pieņēmēja_loma = ((self.cb_pie_loma.currentText() if hasattr(self, 'cb_pie_loma') else 'Pieņēmējs') or 'Pieņēmējs').strip() or 'Pieņēmējs'
        d.nodevēja_loma = ((self.cb_nod_loma.currentText() if hasattr(self, 'cb_nod_loma') else 'Iekārtas/u /pakalpojuma/u nodevējs') or 'Iekārtas/u /pakalpojuma/u nodevējs').strip() or 'Iekārtas/u /pakalpojuma/u nodevējs'
        d.pieņēmējs = self._persona_from_inputs(self.pie_in) if hasattr(self, 'pie_in') else Persona()
        d.nodevējs = self._persona_from_inputs(self.nod_in) if hasattr(self, 'nod_in') else Persona()
        d.rekviziti = self._persona_from_inputs(self.rek_in) if hasattr(self, 'rek_in') else Persona()

        # Pozīcijas
        poz = []
        idx = self._poz_col_indices()
        for r in range(self.tab.rowCount()):
            apr = self.tab.item(r, idx["apraksts"]).text() if self.tab.item(r, idx["apraksts"]) else ""
            daudz = to_decimal(self.tab.item(r, idx["daudzums"]).text() if self.tab.item(r, idx["daudzums"]) else "0")
            vien = self.tab.item(r, idx["vieniba"]).text() if self.tab.item(r, idx["vieniba"]) else ""
            cena = to_decimal(self.tab.item(r, idx["cena"]).text() if self.tab.item(r, idx["cena"]) else "0")

            foto_path = ""
            if "foto" in idx:
                foto_path = self.tab.item(r, idx["foto"]).text() if self.tab.item(r, idx["foto"]) else ""

            ser_nr = self.tab.item(r, idx.get("serial", -1)).text() if idx.get("serial") is not None and self.tab.item(r, idx.get("serial")) else ""
            gar = self.tab.item(r, idx.get("warranty", -1)).text() if idx.get("warranty") is not None and self.tab.item(r, idx.get("warranty")) else ""
            piez_poz = self.tab.item(r, idx.get("notes", -1)).text() if idx.get("notes") is not None and self.tab.item(r, idx.get("notes")) else ""

            if not apr and daudz == 0 and not ser_nr and not gar and not piez_poz and not foto_path:
                continue

            poz.append(Pozīcija(
                apraksts=apr,
                daudzums=daudz,
                vienība=vien,
                cena=cena,
                seriālais_nr=ser_nr,
                garantija=gar,
                piezīmes_pozīcijai=piez_poz,
                attēla_ceļš=foto_path
            ))

        d.pozīcijas = poz

        # Pielāgotās kolonnas
        d.custom_columns = self.data.custom_columns.copy()
        # Atjaunināt pielāgoto kolonnu datus no tabulas
        for col_idx, col in enumerate(d.custom_columns):
            col_data = []
            for r in range(self.tab.rowCount()):
                item = self.tab.item(r, idx["custom_start"] + col_idx)  # Pielāgotās kolonnas sākas pēc standarta kolonnu bloka
                col_data.append(item.text() if item else "")
            col['data'] = col_data

        # JAUNS: pozīciju kolonnu konfigurācija + kopsavilkums
        try:
            d.poz_columns_config = self._poz_cfg().copy() if isinstance(self._poz_cfg(), dict) else {}
        except Exception:
            d.poz_columns_config = {}
        try:
            d.show_price_summary = bool(self.ck_show_price_summary.isChecked()) if hasattr(self, "ck_show_price_summary") else bool(getattr(self.data, "show_price_summary", True))
        except Exception:
            d.show_price_summary = bool(getattr(self.data, "show_price_summary", True))

        # JAUNS: saglabā Pozīciju tabulas kolonnu secību (GUI -> PDF)
        try:
            d.poz_columns_visual_order = self._poz_get_visual_order_keys()
        except Exception:
            try:
                d.poz_columns_visual_order = list(getattr(self.data, 'poz_columns_visual_order', []) or [])
            except Exception:
                d.poz_columns_visual_order = []

        # JAUNS: saglabā Pozīciju tabulas kolonnu platumus/izkārtojumu (GUI)
        try:
            hdr = self.tab.horizontalHeader()
            st = hdr.saveState()
            try:
                st_bytes = bytes(st)
            except Exception:
                st_bytes = st.data() if hasattr(st, "data") else b""
            d.poz_header_state_b64 = base64.b64encode(st_bytes).decode("ascii") if st_bytes else ""
        except Exception:
            try:
                d.poz_header_state_b64 = str(getattr(self.data, "poz_header_state_b64", "") or "")
            except Exception:
                d.poz_header_state_b64 = ""

        # Iestatījumi
        d.iekļaut_pvn = self.ck_pvn.isChecked()
        d.pvn_likme = to_decimal(self.in_pvn.value())
        d.parakstu_rindas = self.ck_paraksti.isChecked()
        d.paraksta_rezims = str(self.cb_paraksta_rezims.currentData() or 'physical')
        d.paraksta_nav_teksts = self.in_paraksta_nav_teksts.text().strip()
        d.paraksta_vards_rekviziti = self.in_paraksta_vards_rekviziti.text().strip()
        d.paraksta_vards_pienemejs = self.in_paraksta_vards_pienemejs.text().strip()
        d.paraksta_vards_nodevejs = self.in_paraksta_vards_nodevejs.text().strip()
        d.papildu_parakstu_rindas = _parse_extra_signature_rows_text(self.in_papildu_parakstu_rindas.toPlainText() if hasattr(self, 'in_papildu_parakstu_rindas') else '')
        d.logotipa_ceļš = self.in_logo.text().strip()
        d.fonts_ceļš = self.in_fonts.text().strip()
        d.docx_template_path = self.in_docx_template.text().strip()

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
        d.default_execution_days = int(self.in_default_execution_days.value())
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
        d.cover_page_title = (self.in_cover_page_title.text() or '').strip() or ((self.in_doc_nosaukums.text() or '').strip() or (self.cmb_doc_tips.currentText() if hasattr(self, 'cmb_doc_tips') else 'Dokuments'))
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
        # --- JAUNS: Pamata datu PDF šifrēšanas lauki (ja eksistē) ---
        try:
            if hasattr(self, "ck_pdf_encrypt_basic") and hasattr(self, "in_pdf_password_basic"):
                d.enable_pdf_encryption = self.ck_pdf_encrypt_basic.isChecked()
                d.pdf_user_password = self.in_pdf_password_basic.text().strip()
                # Owner parole var palikt no "Iestatījumi & Eksports" (ja lietotājs to lieto),
                # bet ja nav, ģenerēsies automātiski šifrēšanas brīdī.
        except Exception:
            pass

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
        if hasattr(self, "photos_table") and self.photos_table is not None:
            for r in range(self.photos_table.rowCount()):
                it = self.photos_table.item(r, 3)
                dat = it.data(Qt.UserRole) if it is not None else None
                if not dat:
                    continue
                att.append(Attēls(ceļš=dat["ceļš"], paraksts=dat.get("paraksts", "")))
        else:
            # Back-compat (ja kādā vecā stāvoklī vēl ir img_list)
            if getattr(self, "img_list", None) is not None:
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

    # ----- Tab: Puses / Rekvizīti -----

    def _apply_party_mode_ui(self):
        try:
            mode = (self.cb_party_mode.currentData() or 'puses') if hasattr(self, 'cb_party_mode') else 'puses'
            if hasattr(self, 'tab_puses_widget') and self.tab_puses_widget is not None:
                self.tab_puses_widget.setVisible(mode in ('puses', 'abi'))
            if hasattr(self, 'tab_rekviziti_widget') and self.tab_rekviziti_widget is not None:
                self.tab_rekviziti_widget.setVisible(mode in ('rekviziti', 'abi'))
            if hasattr(self, 'grp_rekviziti') and self.grp_rekviziti is not None:
                self.grp_rekviziti.setTitle((self.in_rekvizitu_virsraksts.text() or 'Rekvizīti').strip())
        except Exception:
            pass
        try:
            self._update_preview()
        except Exception:
            pass

    def _update_party_role_titles(self, _text=''):
        try:
            if hasattr(self, 'grp_pieņēmējs') and self.grp_pieņēmējs is not None:
                title1 = ((self.cb_pie_loma.currentText() if hasattr(self, 'cb_pie_loma') else 'Pieņēmējs') or 'Pieņēmējs').strip() or 'Pieņēmējs'
                self.grp_pieņēmējs.setTitle(title1)
            if hasattr(self, 'grp_nodevējs') and self.grp_nodevējs is not None:
                title2 = ((self.cb_nod_loma.currentText() if hasattr(self, 'cb_nod_loma') else 'Iekārtas/u /pakalpojuma/u nodevējs') or 'Iekārtas/u /pakalpojuma/u nodevējs').strip() or 'Iekārtas/u /pakalpojuma/u nodevējs'
                self.grp_nodevējs.setTitle(title2)
        except Exception:
            pass

    def _on_rekviziti_title_changed(self, _text=''):
        try:
            if hasattr(self, 'grp_rekviziti') and self.grp_rekviziti is not None:
                self.grp_rekviziti.setTitle((self.in_rekvizitu_virsraksts.text() or 'Rekvizīti').strip())
        except Exception:
            pass
        try:
            self._update_preview()
        except Exception:
            pass

    def _persona_to_inputs(self, persona_inputs, persona_data: dict):
        if not persona_inputs:
            return
        persona_inputs[0].setText(persona_data.get("nosaukums", ""))
        persona_inputs[1].setText(persona_data.get("reģ_nr", ""))
        persona_inputs[2].setText(persona_data.get("adrese", ""))
        persona_inputs[3].setText(persona_data.get("kontaktpersona", ""))
        persona_inputs[4].setText(persona_data.get("amats", ""))
        pilnvaras_val = persona_data.get("pilnvaras_pamats", "")
        if hasattr(persona_inputs[5], 'findText'):
            idx = persona_inputs[5].findText(pilnvaras_val)
            if idx >= 0:
                persona_inputs[5].setCurrentIndex(idx)
        persona_inputs[6].setText(persona_data.get("tālrunis", ""))
        persona_inputs[7].setText(persona_data.get("epasts", ""))
        persona_inputs[8].setText(persona_data.get("web_lapa", ""))
        persona_inputs[9].setText(persona_data.get("bankas_konts", ""))
        juridiskais_statuss_val = persona_data.get("juridiskais_statuss", "")
        if hasattr(persona_inputs[10], 'findText'):
            idx = persona_inputs[10].findText(juridiskais_statuss_val)
            if idx >= 0:
                persona_inputs[10].setCurrentIndex(idx)

    def _persona_from_inputs(self, persona_inputs) -> Persona:
        return Persona(
            nosaukums=persona_inputs[0].text().strip(),
            reģ_nr=persona_inputs[1].text().strip(),
            adrese=persona_inputs[2].text().strip(),
            kontaktpersona=persona_inputs[3].text().strip(),
            amats=persona_inputs[4].text().strip(),
            pilnvaras_pamats=persona_inputs[5].currentText(),
            tālrunis=persona_inputs[6].text().strip(),
            epasts=persona_inputs[7].text().strip(),
            web_lapa=persona_inputs[8].text().strip(),
            bankas_konts=persona_inputs[9].text().strip(),
            juridiskais_statuss=persona_inputs[10].currentText()
        )

    def _persona_group(self, virsraksts: str, is_pieņēmējs: bool, target_key: str = None):
            box = QGroupBox(virsraksts)
            form = QFormLayout()
            nos = QLineEdit();
            reg = QLineEdit();
            adr = QLineEdit();
            kont = QLineEdit();
            tel = QLineEdit();
            ep = QLineEdit()
            web = QLineEdit(); web.setPlaceholderText("https://...")
            bankas_konts = QLineEdit()
            juridiskais_statuss = QComboBox()
            juridiskais_statuss.addItems(["", "Juridiska persona", "Fiziska persona", "Pašnodarbinātais"])

            form.addRow("Nosaukums / Vārds, Uzvārds", nos)
            form.addRow("Reģ. Nr. / personas kods", reg)
            adr_row = QWidget()
            adr_row_layout = QHBoxLayout()
            adr_row_layout.setContentsMargins(0, 0, 0, 0)
            adr_row_layout.setSpacing(6)
            btn_map = QToolButton()
            btn_map.setText("📍")
            btn_map.setToolTip("Atlasīt adresi kartē")
            btn_map.clicked.connect(lambda: self._begin_address_pick(adr))
            adr_row_layout.addWidget(adr, 1)
            adr_row_layout.addWidget(btn_map)
            adr_row.setLayout(adr_row_layout)
            form.addRow("Adrese", adr_row)
            form.addRow("Kontaktpersona", kont)

            amats = QLineEdit()
            pilnvaras_pamats = QComboBox()
            pilnvaras_pamats.addItems(["Pilnvaras pamats", "Līgums", "Uzņēmuma īpašumtiesības", "Cits"])

            form.addRow("Amats", amats)
            form.addRow("Pilnvaras pamats", pilnvaras_pamats)

            form.addRow("Tālrunis", tel)
            form.addRow("E-pasts", ep)
            form.addRow("Web lapa", web)
            form.addRow("Bankas konts", bankas_konts)
            form.addRow("Juridiskais statuss", juridiskais_statuss)

            btn_load_from_ab = QPushButton("Ielādēt no adrešu grāmatas")
            btn_save_to_ab = QPushButton("Saglabāt adrešu grāmatā")
            btn_pick_from_ur = QPushButton("Izvēlēties uzņēmumu (UR)")
            btn_pick_from_ur.setToolTip("Meklēt un izvēlēties uzņēmumu no data.gov.lv (Uzņēmumu reģistrs) un aizpildīt laukus automātiski")

            effective_target_key = (target_key or ("pie" if is_pieņēmējs else "nod")).strip().lower()

            def _target_inputs():
                return getattr(self, f"{effective_target_key}_in", None) or (pie_in if False else None)

            btn_load_from_ab.clicked.connect(lambda: self._load_persona_from_address_book(_target_inputs()))
            btn_save_to_ab.clicked.connect(lambda: self._save_persona_to_address_book(_target_inputs()))
            btn_pick_from_ur.clicked.connect(lambda: self._open_ur_company_picker(_target_inputs()))

            btn_layout = QHBoxLayout()
            btn_layout.addWidget(btn_load_from_ab)
            btn_layout.addWidget(btn_save_to_ab)
            btn_layout.addWidget(btn_pick_from_ur)
            form.addRow(btn_layout)

            box.setLayout(form)
            return box, (nos, reg, adr, kont, amats, pilnvaras_pamats, tel, ep, web, bankas_konts, juridiskais_statuss)


    # ---------------------- data.gov.lv (Uzņēmumu reģistrs) integrācija ----------------------
    _UR_REGISTER_RESOURCE_ID = "25e80bf3-f107-4ab4-89ef-251b5b9374e9"  # Uzņēmumu reģistrs (register.csv) DataStore
    _UR_OFFICERS_RESOURCE_ID = "e665114a-73c2-4375-9470-55874b4cfa6b"  # Amatpersonas (officers.csv) DataStore
    _UR_API_BASE = "https://data.gov.lv/dati/api/action"

    def _ur_api_get(self, endpoint: str, params: dict, timeout: int = 15):
        """Drošs GET uz CKAN API (atgriež dict vai paceļ Exception)."""
        import requests
        from urllib.parse import urlencode

        url = f"{self._UR_API_BASE}/{endpoint}"
        r = requests.get(url, params=params, timeout=timeout, headers={'User-Agent': 'Pienemsanas-Nodosanas-Akts/1.0'})
        r.raise_for_status()
        js = r.json()
        if not isinstance(js, dict) or not js.get('success'):
            raise RuntimeError(f"UR API atbilde nav veiksmīga: {js}")
        return js.get('result')

    def _ur_company_search(self, query: str, limit: int = 25):
        """Meklē uzņēmumus UR datu kopā pēc nosaukuma vai reģ.nr. (atgriež sarakstu ar dict)."""
        q = (query or '').strip()
        if not q:
            return []
        # Kešatmiņa (ļoti vienkārša)
        try:
            if not hasattr(self, '_ur_cache_search'):
                self._ur_cache_search = {}
            cache_key = f"{q.lower()}|{limit}"
            if cache_key in self._ur_cache_search:
                return self._ur_cache_search[cache_key]
        except Exception:
            pass

        # Datastore SQL – precīzāk nekā 'q='
        # NB: lietojam parametru aizvietošanu tikai ar manuālu escaping (CKAN datastore_search_sql neatbalsta parametrus),
        # tāpēc rūpīgi izfiltrējam vienkāršam LIKE vaicājumam.
        safe = q.replace("'", "''")
        # Ja ievadīts ciparu sākums, prioritetizējam regcode
        where = f"(lower(name) like lower('%{safe}%'))"
        if q.replace(' ', '').isdigit():
            where = f"(cast(regcode as text) like '{safe}%') OR {where}"

        sql = (
            "SELECT regcode, name, address, type_text, regtype_text "
            f"FROM \"{self._UR_REGISTER_RESOURCE_ID}\" "
            f"WHERE {where} "
            "ORDER BY regcode DESC "
            f"LIMIT {int(max(1, min(limit, 50)))}"
        )

        res = self._ur_api_get('datastore_search_sql', {'sql': sql})
        records = (res or {}).get('records') or []
        try:
            self._ur_cache_search[cache_key] = records
            # ierobežojam cache izmēru
            if len(self._ur_cache_search) > 200:
                # izmetam pirmo atslēgu
                k0 = next(iter(self._ur_cache_search))
                self._ur_cache_search.pop(k0, None)
        except Exception:
            pass
        return records

    def _ur_company_details(self, regcode: str) -> dict:
        """Atgriež pilnu ierakstu no UR reģistra pēc reģ.nr. (vai {})."""
        rc = (str(regcode or '').strip())
        if not rc:
            return {}
        try:
            if not hasattr(self, '_ur_cache_details'):
                self._ur_cache_details = {}
            if rc in self._ur_cache_details:
                return self._ur_cache_details.get(rc) or {}
        except Exception:
            pass

        safe = rc.replace("'", "''")
        sql = (
            "SELECT regcode, sepa, name, address, type_text, regtype_text, registered, terminated, closed, reregistration_term "
            f"FROM \"{self._UR_REGISTER_RESOURCE_ID}\" "
            f"WHERE cast(regcode as text) = '{safe}' "
            "LIMIT 1"
        )
        res = self._ur_api_get('datastore_search_sql', {'sql': sql})
        recs = (res or {}).get('records') or []
        rec = recs[0] if recs else {}
        try:
            self._ur_cache_details[rc] = rec
            if len(self._ur_cache_details) > 500:
                k0 = next(iter(self._ur_cache_details))
                self._ur_cache_details.pop(k0, None)
        except Exception:
            pass
        return rec or {}

    def _ur_officers_for_regcode(self, regcode: str, limit: int = 400):
        """Atrod amatpersonas dotajam reģ.nr. (atgriež sarakstu ar dict).

        Primāri izmanto CKAN DataStore SQL. Ja resursam nav DataStore skata vai SQL vaicājums neizdodas,
        mēģina 'datastore_search' ar q=regcode un pēc tam filtrē rezultātus Python pusē.
        """
        rc = (str(regcode or '').strip())
        if not rc:
            return []
        try:
            if not hasattr(self, '_ur_cache_officers'):
                self._ur_cache_officers = {}
            if rc in self._ur_cache_officers:
                return self._ur_cache_officers[rc]
        except Exception:
            pass

        safe = rc.replace("'", "''")
        records = []
        # 1) Mēģinām ar SQL (ātrāk un precīzāk)
        try:
            sql = (
                f"SELECT * FROM \"{self._UR_OFFICERS_RESOURCE_ID}\" "
                f"WHERE (cast(at_legal_entity_registration_number as text) = '{safe}' "
                f"   OR cast(legal_entity_registration_number as text) = '{safe}') "
                "ORDER BY registered_on DESC "
                f"LIMIT {int(max(1, min(limit, 1000)))}"
            )
            res = self._ur_api_get('datastore_search_sql', {'sql': sql})
            records = (res or {}).get('records') or []
        except Exception:
            records = []

        # 2) Fallback: datastore_search ar q=regcode (der, ja SQL nav pieejams / resursam nav skata)
        if not records:
            try:
                # Ņemam vairāk, jo q= meklēšana var atgriezt arī citus ierakstus
                fetch_limit = int(max(50, min(limit, 1000)))
                offset = 0
                out = []
                while offset < 3000 and len(out) < fetch_limit:
                    res = self._ur_api_get('datastore_search', {
                        'resource_id': self._UR_OFFICERS_RESOURCE_ID,
                        'q': rc,
                        'limit': fetch_limit,
                        'offset': offset
                    })
                    chunk = (res or {}).get('records') or []
                    if not chunk:
                        break
                    out.extend(chunk)
                    offset += fetch_limit
                    # ja nav vairāk, beidzam
                    if len(chunk) < fetch_limit:
                        break

                def get_any(it, keys):
                    for k in keys:
                        if k in it and it.get(k) not in (None, ''):
                            return it.get(k)
                    return None

                # filtrējam pēc jebkuras zināmās reģ.nr. kolonnas
                filtered = []
                for it in out:
                    v1 = get_any(it, ['at_legal_entity_registration_number', 'legal_entity_registration_number',
                                     'registration_number', 'regcode', 'legal_entity_regcode'])
                    if v1 is None:
                        # dažkārt ir iekšā kā skaitlis citā formātā
                        continue
                    if str(v1).strip() == rc:
                        filtered.append(it)

                # ja filtrs nav atradis, tomēr atgriežam visu "out" (labāk nekā tukšs, pick_best nosvērs)
                records = filtered or out
            except Exception:
                records = []

        try:
            self._ur_cache_officers[rc] = records
            if len(self._ur_cache_officers) > 500:
                k0 = next(iter(self._ur_cache_officers))
                self._ur_cache_officers.pop(k0, None)
        except Exception:
            pass
        return records

    def _pick_best_contact_from_officers(self, officers: list) -> tuple:
        """Atgriež (name, position, pk_prefix, rights_note) pēc labākā minējuma.

        Šī funkcija ir tolerantāka pret kolonnu nosaukumu variācijām (piem., name/full_name/person_name, position/role u.c.).
        pk_prefix ir, piem., '123456-' (ja pieejams maskētais personas kods), rights_note – īsa piezīme par pārstāvības tiesībām.
        """
        if not officers:
            return ("", "", "", "")

        def get_first(it: dict, keys: list):
            for k in keys:
                if k in it and it.get(k) not in (None, ""):
                    return it.get(k)
            return None

        def sget(it: dict, keys: list) -> str:
            v = get_first(it, keys)
            return str(v or '').strip()

        def norm(s: str) -> str:
            s = (str(s or '').strip().lower()
                 .replace('ē','e').replace('ā','a').replace('ī','i').replace('ū','u')
                 .replace('č','c').replace('š','s').replace('ģ','g').replace('ķ','k')
                 .replace('ļ','l').replace('ņ','n').replace('ž','z'))
            # samazinām atstarpju/komatu ietekmi
            s = re.sub(r"\s+", " ", s)
            return s

        def pk_prefix_from_mask(mask: str) -> str:
            m = str(mask or '').strip()
            if not m:
                return ""
            # bieži ir formāts 123456-*****
            if '-' in m:
                left = m.split('-', 1)[0]
                left = ''.join(ch for ch in left if ch.isdigit())
                if len(left) >= 6:
                    return left[:6] + "-"
            digits = ''.join(ch for ch in m if ch.isdigit())
            if len(digits) >= 6:
                return digits[:6] + "-"
            return ""

        def rights_note(it: dict) -> str:
            rt = sget(it, ['rights_of_representation_type', 'representation_rights_type', 'rights_type'])
            atleast = get_first(it, ['representation_with_at_least', 'representation_minimum', 'with_at_least'])
            parts = []
            if rt:
                parts.append(rt)
            try:
                if atleast not in (None, "", 0, "0"):
                    parts.append(f"kopā ar vismaz {int(atleast)}")
            except Exception:
                pass
            return "; ".join(parts)

        # atslēgas dažādos resursa izlaidumos var atšķirties
        NAME_KEYS = ['name', 'full_name', 'person_name', 'person', 'officer_name']
        POS_KEYS  = ['position', 'role', 'position_text', 'position_lv', 'role_text', 'office']
        GB_KEYS   = ['governing_body', 'body', 'governing_body_text']
        PK_KEYS   = ['latvian_identity_number_masked', 'identity_number_masked', 'personal_code_masked', 'person_code_masked']

        def score(it: dict) -> int:
            pos_raw = sget(it, POS_KEYS)
            gb_raw = sget(it, GB_KEYS)
            pos = norm(pos_raw)
            gb = norm(gb_raw)
            sc = 0

            # Valdes priekšsēdētājs / priekšsēdētājs
            chair_words = ['valdes priekssedet', 'priekssedetajs', 'chairman', 'chair', 'priekšsēdētājs', 'prieksedetajs']
            if any(w in pos for w in chair_words):
                sc += 200

            # Valdes loceklis
            if 'valdes locekl' in pos or 'board member' in pos:
                sc += 140

            # Ja governing_body norāda uz valdi
            if 'valde' in gb or 'board' in gb:
                sc += 40

            # Papildpunkti, ja ir pārstāvības tiesības
            if sget(it, ['rights_of_representation_type', 'representation_rights_type', 'rights_type']):
                sc += 10

            # Papildpunkti, ja ir vārds/uzvārds
            if sget(it, NAME_KEYS):
                sc += 10

            # Nedaudz preferējam ierakstus ar maskētu PK (lai varam ielikt 1. daļu)
            if sget(it, PK_KEYS):
                sc += 3

            return sc

        best = None
        best_sc = -1
        for it in officers:
            try:
                sc = score(it)
            except Exception:
                sc = 0
            if sc > best_sc:
                best_sc = sc
                best = it
        if not best:
            best = officers[0]

        nm = sget(best, NAME_KEYS)
        pos = sget(best, POS_KEYS)
        pkp = pk_prefix_from_mask(sget(best, PK_KEYS))
        rn = rights_note(best)
        return (nm, pos, pkp, rn)

    def _open_ur_company_picker(self, persona_inputs):
        """Atver meklēšanas dialogu un aizpilda personu laukus no UR datiem."""
        if not persona_inputs:
            return
        try:
            from PySide6.QtCore import QTimer, Qt
            from PySide6.QtWidgets import (
                QDialog, QVBoxLayout, QHBoxLayout, QLabel, QLineEdit, QListWidget, QListWidgetItem,
                QPushButton, QDialogButtonBox, QMessageBox, QApplication
            )
        except Exception:
            QMessageBox.warning(self, "Kļūda", "Neizdevās ielādēt Qt komponentes uzņēmumu izvēlei.")
            return

        dlg = QDialog(self)
        dlg.setWindowTitle("Izvēlēties uzņēmumu (data.gov.lv / UR)")
        dlg.resize(760, 520)
        v = QVBoxLayout(dlg)

        lbl = QLabel("Ievadi uzņēmuma nosaukumu vai reģistrācijas numuru (meklē UR atvērtajos datos):")
        v.addWidget(lbl)

        row = QHBoxLayout()
        le = QLineEdit()
        le.setPlaceholderText("piem.: SIA PAPPUS vai 4000...")
        row.addWidget(le, 1)
        btn_search = QPushButton("Meklēt")
        row.addWidget(btn_search)
        v.addLayout(row)

        info = QLabel(" ")
        info.setWordWrap(True)
        v.addWidget(info)

        lst = QListWidget()
        v.addWidget(lst, 1)

        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        buttons.button(QDialogButtonBox.Ok).setText("Aizpildīt")
        buttons.button(QDialogButtonBox.Cancel).setText("Atcelt")
        buttons.button(QDialogButtonBox.Ok).setEnabled(False)
        v.addWidget(buttons)

        state = {'results': []}

        def render_results(records):
            lst.clear()
            state['results'] = records or []
            for rec in state['results']:
                name = str(rec.get('name') or '').strip()
                regcode = str(rec.get('regcode') or '').strip()
                addr = str(rec.get('address') or '').strip()
                type_text = str(rec.get('type_text') or '').strip()
                line = f"{name}  ({regcode})"
                if type_text:
                    line += f" — {type_text}"
                if addr:
                    line += f" — {addr}"
                it = QListWidgetItem(line)
                it.setData(Qt.UserRole, rec)
                lst.addItem(it)
            info.setText(f"Atrasti ieraksti: {len(state['results'])}")
            buttons.button(QDialogButtonBox.Ok).setEnabled(lst.currentItem() is not None)

        def do_search():
            q = le.text().strip()
            if not q:
                info.setText("Ievadi meklējamo frāzi.")
                render_results([])
                return
            info.setText("Meklē... (data.gov.lv)")
            try:
                recs = self._ur_company_search(q, limit=30)
                render_results(recs)
                if not recs:
                    info.setText("Nekas netika atrasts. Pamēģini citu frāzi vai ievadi reģ.nr.")
            except Exception as e:
                render_results([])
                info.setText("Neizdevās saņemt datus no data.gov.lv (pārbaudi internetu vai ugunsmūri).")
                QMessageBox.warning(self, "UR meklēšana", f"Neizdevās meklēt uzņēmumus:\n{e}")

        # Debounce meklēšanai, lai nepārslogotu API
        t = QTimer(dlg)
        t.setSingleShot(True)
        t.setInterval(350)
        le.textChanged.connect(lambda: t.start())
        t.timeout.connect(do_search)
        btn_search.clicked.connect(do_search)

        lst.currentItemChanged.connect(lambda cur, prev: buttons.button(QDialogButtonBox.Ok).setEnabled(cur is not None))
        lst.itemDoubleClicked.connect(lambda _: buttons.accepted.emit())

        def apply_selected():
            item = lst.currentItem()
            if not item:
                return
            rec = item.data(Qt.UserRole) or {}
            name = str(rec.get('name') or '').strip()
            regcode = str(rec.get('regcode') or '').strip()
            addr = str(rec.get('address') or '').strip()

            # Aizpildām bāzes laukus
            try:
                persona_inputs[0].setText(name)
                persona_inputs[1].setText(regcode)
                persona_inputs[2].setText(addr)
                # juridiska persona pēc noklusējuma
                try:
                    idx = persona_inputs[10].findText("Juridiska persona")
                    if idx >= 0:
                        persona_inputs[10].setCurrentIndex(idx)
                except Exception:
                    pass
            except Exception:
                pass

            # Papildu detaļas (piem., SEPA/IBAN) – mēģinām paņemt pilno ierakstu pēc reģ.nr.
            try:
                det = self._ur_company_details(regcode) or {}
                det_name = str(det.get('name') or '').strip()
                det_addr = str(det.get('address') or '').strip()
                sepa = str(det.get('sepa') or '').strip()

                # Pārrakstām ar detalizētākiem datiem, ja tie ir pieejami
                if det_name:
                    persona_inputs[0].setText(det_name)
                if det_addr:
                    persona_inputs[2].setText(det_addr)

                # Bankas konts – pārrakstām (ja UR nav, notīrām, lai nepaliek vecais)
                persona_inputs[9].setText(sepa or "")
            except Exception:
                # klusais fallback – bāzes dati jau aizpildīti
                pass


            # Pilnvaras pamats – juridiskai personai bieži der "Uzņēmuma īpašumtiesības"
            try:
                if hasattr(persona_inputs[5], 'currentText') and persona_inputs[5].currentText() in ("", "Pilnvaras pamats"):
                    idxp = persona_inputs[5].findText("Uzņēmuma īpašumtiesības")
                    if idxp >= 0:
                        persona_inputs[5].setCurrentIndex(idxp)
            except Exception:
                pass

            # Mēģinām paņemt valdes priekšsēdētāju / valdes locekli kā kontaktpersonu + amatu
            try:
                officers = self._ur_officers_for_regcode(regcode)
                nm, pos, pkp, rn = self._pick_best_contact_from_officers(officers)

                # Kontaktpersona – pārrakstām (ja nav atrasts, notīrām, lai nepaliek vecais)
                if nm:
                    if pkp:
                        persona_inputs[3].setText(f"{nm} (pk: {pkp})")
                    else:
                        persona_inputs[3].setText(nm)
                else:
                    persona_inputs[3].setText("")

                # Amats – pārrakstām (ja nav atrasts, notīrām)
                if pos:
                    if rn:
                        persona_inputs[4].setText(f"{pos} | {rn}")
                    else:
                        persona_inputs[4].setText(pos)
                else:
                    persona_inputs[4].setText("")
            except Exception:
                # klusais fallback – bāzes dati jau aizpildīti
                pass



            dlg.accept()

        buttons.accepted.connect(apply_selected)
        buttons.rejected.connect(dlg.reject)

        dlg.exec()
        try:
            self._update_party_role_titles()
            self._update_preview()
        except Exception:
            pass

    def _būvēt_puses_tab(self):
        self.cb_pie_loma = QComboBox()
        self.cb_pie_loma.setEditable(True)
        self.cb_pie_loma.addItems([
            "Pieņēmējs", "Pieņēmējs / Pārdevējs", "Saņēmējs", "Pircējs", "Pasūtītājs", "Komisija", "Atbildīgā persona"
        ])
        self.cb_nod_loma = QComboBox()
        self.cb_nod_loma.setEditable(True)
        self.cb_nod_loma.addItems([
            "Iekārtas/u /pakalpojuma/u nodevējs", "Iekārtas/u /pakalpojuma/u nodevējs / Pircējs", "Nodevējs", "Piegādātājs", "Izpildītājs", "Nosūtītājs"
        ])

        self.grp_pieņēmējs, self.pie_in = self._persona_group("Pieņēmējs", True)
        self.grp_nodevējs, self.nod_in = self._persona_group("Iekārtas/u /pakalpojuma/u nodevējs", False)

        def _build_party_page(title_widget, group_widget):
            page = QWidget()
            page_layout = QVBoxLayout(page)
            page_layout.setContentsMargins(8, 8, 8, 8)
            page_layout.setSpacing(8)
            page_layout.addWidget(title_widget)
            page_layout.addWidget(group_widget)
            page_layout.addStretch(1)
            area = QScrollArea()
            area.setWidgetResizable(True)
            try:
                area.setAlignment(Qt.AlignTop | Qt.AlignLeft)
            except Exception:
                pass
            area.setWidget(page)
            return area

        pie_title = QWidget()
        pie_title_l = QVBoxLayout(pie_title)
        pie_title_l.setContentsMargins(0, 0, 0, 0)
        pie_title_l.setSpacing(6)
        pie_title_l.addWidget(QLabel("Pieņēmēja virsraksts"))
        pie_title_l.addWidget(self.cb_pie_loma)

        nod_title = QWidget()
        nod_title_l = QVBoxLayout(nod_title)
        nod_title_l.setContentsMargins(0, 0, 0, 0)
        nod_title_l.setSpacing(6)
        nod_title_l.addWidget(QLabel("Nodevēja virsraksts"))
        nod_title_l.addWidget(self.cb_nod_loma)

        self.party_tabs = QTabWidget()
        self.party_tabs.setUsesScrollButtons(True)
        self.party_tabs.addTab(_build_party_page(pie_title, self.grp_pieņēmējs), "Pieņēmējs")
        self.party_tabs.addTab(_build_party_page(nod_title, self.grp_nodevējs), "Nodevējs")

        self.tab_puses_widget = QWidget()
        main_layout = QVBoxLayout(self.tab_puses_widget)
        main_layout.setContentsMargins(6, 6, 6, 6)
        main_layout.addWidget(self.party_tabs)

        self.tabs.addTab(self.tab_puses_widget, "Puses")
        self.cb_pie_loma.currentTextChanged.connect(self._update_party_role_titles)
        self.cb_pie_loma.currentTextChanged.connect(self._update_preview)
        self.cb_nod_loma.currentTextChanged.connect(self._update_party_role_titles)
        self.cb_nod_loma.currentTextChanged.connect(self._update_preview)

        for w in self.pie_in:
            if isinstance(w, QLineEdit):
                w.textChanged.connect(self._update_preview)
            elif isinstance(w, QComboBox):
                w.currentIndexChanged.connect(self._update_preview)

        for w in self.nod_in:
            if isinstance(w, QLineEdit):
                w.textChanged.connect(self._update_preview)
            elif isinstance(w, QComboBox):
                w.currentIndexChanged.connect(self._update_preview)

        rek_tab = QWidget()
        rek_layout = QVBoxLayout(rek_tab)
        rek_layout.setContentsMargins(8, 8, 8, 8)
        rek_layout.setSpacing(8)
        self.in_rekvizitu_virsraksts = QLineEdit("Rekvizīti")
        self.in_rekvizitu_virsraksts.setPlaceholderText("Sadaļas virsraksts, piem. Rekvizīti")
        self.in_rekvizitu_virsraksts.textChanged.connect(self._on_rekviziti_title_changed)
        rek_layout.addWidget(self.in_rekvizitu_virsraksts)
        self.grp_rekviziti, self.rek_in = self._persona_group("Rekvizīti", True, target_key="rek")
        rek_layout.addWidget(self.grp_rekviziti)
        rek_layout.addStretch(1)

        scroll_rek = QScrollArea()
        scroll_rek.setWidgetResizable(True)
        try:
            scroll_rek.setAlignment(Qt.AlignTop | Qt.AlignLeft)
        except Exception:
            pass
        scroll_rek.setWidget(rek_tab)

        self.tab_rekviziti_widget = QWidget()
        tab_rek_layout = QVBoxLayout(self.tab_rekviziti_widget)
        tab_rek_layout.setContentsMargins(6, 6, 6, 6)
        tab_rek_layout.addWidget(scroll_rek)
        self.tabs.addTab(self.tab_rekviziti_widget, "Rekvizīti")

        for w in self.rek_in:
            if isinstance(w, QLineEdit):
                w.textChanged.connect(self._update_preview)
            elif isinstance(w, QComboBox):
                w.currentIndexChanged.connect(self._update_preview)

        self._apply_party_mode_ui()

    def _load_persona_from_address_book(self, persona_inputs):
        if not persona_inputs:
            return
        self._undo_mgr.push_undo(self._snapshot_state('AB_LOAD_TO_PUSES'))
        self._audit('AB_LOAD_TO_PUSES', {})
        items = list(self.address_book.keys())
        if not items:
            QMessageBox.information(self, "Adrešu grāmata", "Adrešu grāmata ir tukša.")
            return
        item, ok = QInputDialog.getItem(self, "Ielādēt personu", "Izvēlieties personu:", items, 0, False)
        if ok and item:
            persona_data = self.address_book[item]
            # --- JAUNS: ja ir parole, prasa to pirms ielādes uz Puses tab ---
            if not self._ab_require_password(item, persona_data, "ielādēt"):
                return
            self._persona_to_inputs(persona_inputs, persona_data)
            try:
                self._update_party_role_titles()
                self._update_preview()
            except Exception:
                pass
            QMessageBox.information(self, "Ielādēts", f"Persona '{item}' ielādēta.")
    # --- JAUNS: Adrešu grāmatas ieraksta nosaukuma ģenerēšana/lietotāja izvēle ---
    def _ab_generate_default_entry_name(self, persona_data: dict) -> str:
        """Noklusējuma adrešu grāmatas ieraksta nosaukums: 'Uzņēmums — Kontaktpersona'."""
        try:
            comp = (persona_data.get("nosaukums") or "").strip()
            kp = (persona_data.get("kontaktpersona") or "").strip()
            if comp and kp:
                return f"{comp} — {kp}"
            return comp or kp or "Persona"
        except Exception:
            return "Persona"

    def _ab_make_unique_key(self, base: str) -> str:
        """Nodrošina unikālu atslēgu adrešu grāmatā (ja jau eksistē, pievieno (2), (3)...)."""
        base = (base or "").strip() or "Persona"
        if base not in getattr(self, "address_book", {}):
            return base
        i = 2
        while True:
            k = f"{base} ({i})"
            if k not in self.address_book:
                return k
            i += 1

    def _ab_get_entry_name_for_save(self, persona_data: dict) -> str:
        """Atgriež ieraksta nosaukumu atkarībā no režīma (auto/lietotājs)."""
        default_name = self._ab_generate_default_entry_name(persona_data)
        try:
            use_auto = True
            if hasattr(self, "chk_ab_auto_name") and self.chk_ab_auto_name is not None:
                use_auto = bool(self.chk_ab_auto_name.isChecked())
        except Exception:
            use_auto = True

        if use_auto:
            return default_name

        # Lietotājs pats ievada nosaukumu
        name, ok = QInputDialog.getText(
            self,
            "Saglabāt adrešu grāmatā",
            "Ievadi nosaukumu (lai atšķirtu vairākas puses vienam uzņēmumam):",
            QLineEdit.Normal,
            default_name
        )
        if not ok:
            return ""
        return (name or "").strip()


    def _save_persona_to_address_book(self, persona_inputs):
        if not persona_inputs:
            return
        self._undo_mgr.push_undo(self._snapshot_state('AB_SAVE'))
        self._audit('AB_SAVE', {})
        nosaukums = persona_inputs[0].text().strip()
        if not nosaukums:
            QMessageBox.warning(self, "Saglabāt personu", "Nosaukums nevar būt tukšs.")
            return
        persona_data = {
            "nosaukums": nosaukums,
            "reģ_nr": persona_inputs[1].text().strip(),
            "adrese": persona_inputs[2].text().strip(),
            "kontaktpersona": persona_inputs[3].text().strip(),
            "amats": persona_inputs[4].text().strip(),
            "pilnvaras_pamats": persona_inputs[5].currentText(),
            "tālrunis": persona_inputs[6].text().strip(),
            "epasts": persona_inputs[7].text().strip(),
            "web_lapa": persona_inputs[8].text().strip(),
            "bankas_konts": persona_inputs[9].text().strip(),
            "juridiskais_statuss": persona_inputs[10].currentText(),
        }
        entry_name = self._ab_get_entry_name_for_save(persona_data)
        if not entry_name:
            return
        entry_name = self._ab_make_unique_key(entry_name)
        self.address_book[entry_name] = persona_data
        self._save_address_book() # Saglabājam adrešu grāmatu failā
        self._update_address_book_list() # Atjaunojam sarakstu GUI
        try:
            self._update_preview()
        except Exception:
            pass

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

        # --- JAUNS: vienmēr būvējam pilnu kolonnu komplektu (kolonnas var paslēpt/atkal parādīt) ---
        # Bāzes kolonnas (UI secība ir fiksēta, lai aprēķini un saglabāšana vienmēr strādātu)
        base_headers = [
            ("apraksts", "Apraksts"),
            ("daudzums", "Daudzums"),
            ("vieniba", "Vienība"),
            ("cena", "Cena"),
            ("summa", "Summa"),
            ("serial", "Seriālais Nr."),
            ("warranty", "Garantija"),
            ("notes", "Piezīmes pozīcijai"),
        ]

        # Pārliecināmies, ka konfigurācija eksistē un ir pilna
        try:
            self.data.poz_columns_config = _merge_poz_columns_config(getattr(self.data, "poz_columns_config", None))
        except Exception:
            self.data.poz_columns_config = _merge_poz_columns_config({})

        # Nodrošinām, ka custom column dict satur arī "visible" (atpakaļsavietojami)
        try:
            for col in getattr(self.data, "custom_columns", []) or []:
                if isinstance(col, dict) and "visible" not in col:
                    col["visible"] = True
        except Exception:
            pass

        headers = []
        for key, default_title in base_headers:
            title = default_title
            try:
                title = str(self.data.poz_columns_config.get(key, {}).get("title", default_title))
            except Exception:
                title = default_title
            headers.append(title)

        # Pielāgotās kolonnas (vienmēr pirms Foto)
        for col in getattr(self.data, "custom_columns", []) or []:
            try:
                headers.append(str(col.get('name', '')))
            except Exception:
                headers.append("")

        # Foto kolonna vienmēr pēdējā (UI), bet var paslēpt
        foto_title = "Foto"
        try:
            foto_title = str(self.data.poz_columns_config.get("foto", {}).get("title", "Foto"))
        except Exception:
            pass
        headers.append(foto_title)

        self.tab = QTableWidget(0, len(headers))
        self.tab.setHorizontalHeaderLabels(headers)

        # --- JAUNS: Undo/Redo pozīciju tabulai ---
        self._ensure_positions_undo_hook()

        # --- JAUNS: kopsavilkuma rādīšana (zem tabulas PDF) ---
        self.ck_show_price_summary = QCheckBox("Rādīt cenu apkopojumu zem tabulas (PDF)")
        try:
            self.ck_show_price_summary.setChecked(bool(getattr(self.data, "show_price_summary", True)))
        except Exception:
            self.ck_show_price_summary.setChecked(True)
        self.ck_show_price_summary.stateChanged.connect(self._update_preview)
        v.addWidget(self.ck_show_price_summary)

        # Padarīt kolonnas pielāgojamas velkot
        for i in range(len(headers)):
            self.tab.horizontalHeader().setSectionResizeMode(i, QHeaderView.Interactive)

        # --- JAUNS: kolonnu konteksta izvēlne (paslēpt/parādīt/pārdēvēt) ---
        try:
            hh = self.tab.horizontalHeader()
            hh.setContextMenuPolicy(Qt.CustomContextMenu)
            hh.customContextMenuRequested.connect(self._poz_header_context_menu)
            # Ļaujam vilkt kolonnas ar peli + saglabājam secību
            try:
                hh.setSectionsMovable(True)
            except Exception:
                pass
            try:
                hh.sectionMoved.connect(lambda *_: self._poz_store_visual_order())
            except Exception:
                pass
            # Ja ir saglabāta secība, pielietojam to uzreiz
            try:
                self._poz_apply_stored_visual_order()
            except Exception:
                pass

        except Exception:
            pass

        # Pielietojam kolonnu redzamību
        self._poz_apply_column_visibility()

        # Teksta pārlikšana šūnās (nevārdi netiek pārrauti; pāriet nākamajā rindā)
        self.tab.setWordWrap(True)
        self.tab.setTextElideMode(Qt.ElideNone)

        # Auto pielāgot rindas/kolonnas pēc teksta (de-bounce, lai nebremzē rakstot)
        self._poz_autoresize_timer = QTimer(self)
        self._poz_autoresize_timer.setSingleShot(True)

        def _apply_poz_autoresize():
            try:
                self.tab.resizeRowsToContents()
                self.tab.resizeColumnsToContents()
            except Exception:
                pass

        self._poz_autoresize_timer.timeout.connect(_apply_poz_autoresize)

        def _schedule_poz_autoresize(*_):
            try:
                self._poz_autoresize_timer.start(120)
            except Exception:
                pass

        self.tab.cellChanged.connect(_schedule_poz_autoresize)

        btns = QHBoxLayout()
        add = QPushButton("Pievienot no noliktavas")
        add.setToolTip("Atver noliktavas meklētāju un pievieno izvēlētās preces pozīciju tabulai")
        add.clicked.connect(self.pievienot_pozīcijas_no_noliktavas)
        add_empty = QPushButton("Tukša pozīcija")
        add_empty.setToolTip("Pievieno tukšu rindu manuālai ievadei")
        add_empty.clicked.connect(self.pievienot_tukšu_pozīciju)
        dzest = QPushButton("Dzēst izvēlēto")
        dzest.clicked.connect(self.dzest_pozīciju)
        cfg_btn = QPushButton("Kolonnu iestatījumi")
        cfg_btn.setToolTip("Ar peles labo pogu uz kolonnas virsraksta var paslēpt/parādīt un pārdēvēt kolonnas")
        cfg_btn.clicked.connect(self._poz_open_settings_for_selected_column)
        add_col = QPushButton("Pievienot kolonnu")
        add_col.clicked.connect(self.pievienot_kolonnu)
        del_col = QPushButton("Dzēst kolonnu")
        del_col.clicked.connect(self.dzest_kolonnu)
        btns.addWidget(add)
        btns.addWidget(add_empty)
        btns.addWidget(dzest)
        btns.addWidget(cfg_btn)
        btns.addWidget(add_col)
        btns.addWidget(del_col)
        btns.addStretch()

        v.addLayout(btns)
        v.addWidget(self.tab)
        w.setLayout(v)
        self.tabs.addTab(w, "Pozīcijas")

        self.tab.cellChanged.connect(self._pārrēķināt_summa)
        self.tab.cellChanged.connect(self._update_preview)


    def _poz_col_indices(self) -> dict:
        """Atgriež kolonnu indeksus 'Pozīcijas' tabulai atbilstoši iestatījumiem.
        Foto kolonna vienmēr ir PĒDĒJĀ (arī pēc pielāgotajām kolonnām).
        """
        # JAUNS: fiksēts bāzes kolonnu bloks, lai jebkuru kolonnu var paslēpt/atkal parādīt,
        # nepārkārtojot indeksus (paslēpšana notiek ar setColumnHidden).
        idx = {
            "apraksts": 0,
            "daudzums": 1,
            "vieniba": 2,
            "cena": 3,
            "summa": 4,
            "serial": 5,
            "warranty": 6,
            "notes": 7,
        }
        idx["custom_start"] = 8
        idx["foto"] = idx["custom_start"] + len(getattr(self.data, "custom_columns", []) or [])
        idx["col_count"] = idx["foto"] + 1
        return idx




    # --- JAUNS: Pozīciju kolonnu kārtība (viegla pārvietošana starp jebkurām kolonnām) ---
    def _poz_get_visual_order_keys(self) -> list:
        """Atgriež kolonnu atslēgas vizuālajā secībā (kā redz lietotājs)."""
        if not hasattr(self, "tab") or not self.tab:
            return []
        hh = self.tab.horizontalHeader()
        order = []
        try:
            for visual in range(self.tab.columnCount()):
                logical = hh.logicalIndex(visual)
                key = self._poz_key_from_column(int(logical))
                if key:
                    order.append(key)
        except Exception:
            return []
        return order

    def _poz_store_visual_order(self):
        """Saglabā pašreizējo vizuālo kolonnu secību (lai tā saglabājas pēc restart)."""
        try:
            self.data.poz_columns_visual_order = self._poz_get_visual_order_keys()
        except Exception:
            pass

    def _poz_apply_stored_visual_order(self):
        """Pielieto saglabāto vizuālo kolonnu secību, pārvietojot header sekcijas."""
        if not hasattr(self, "tab") or not self.tab:
            return
        try:
            desired = list(getattr(self.data, "poz_columns_visual_order", []) or [])
        except Exception:
            desired = []
        if not desired:
            return

        hh = self.tab.horizontalHeader()
        # Pašreizējais key -> logical index
        key_to_logical = {}
        try:
            for logical in range(self.tab.columnCount()):
                k = self._poz_key_from_column(int(logical))
                if k:
                    key_to_logical[k] = int(logical)
        except Exception:
            return

        # Filtrējam tikai eksistējošās
        desired = [k for k in desired if k in key_to_logical]
        if not desired:
            return

        # Pieliekam klāt jaunas (piem., tikko pievienota custom kolonna)
        current = self._poz_get_visual_order_keys()
        for k in current:
            if k not in desired:
                # Foto vienmēr pēdējā, ja eksistē
                if k == "foto":
                    continue
                desired.append(k)
        if "foto" in current and "foto" not in desired:
            desired.append("foto")
        # Ja foto ir sarakstā, pārbīdam uz beigām
        if "foto" in desired:
            desired = [k for k in desired if k != "foto"] + ["foto"]

        # Pārvietojam vizuālajā secībā pa vienam
        try:
            for target_visual, key in enumerate(desired):
                logical = key_to_logical.get(key)
                if logical is None:
                    continue
                cur_visual = hh.visualIndex(logical)
                if cur_visual != target_visual and cur_visual >= 0:
                    hh.moveSection(cur_visual, target_visual)
        except Exception:
            pass

        # Pārrakstam saglabāto secību, lai tā atbilst realitātei
        self._poz_store_visual_order()

    # --- JAUNS: Pozīciju kolonnu redzamība/pārdēvēšana ---
    def _poz_cfg(self) -> dict:
        try:
            self.data.poz_columns_config = _merge_poz_columns_config(getattr(self.data, "poz_columns_config", None))
        except Exception:
            self.data.poz_columns_config = _merge_poz_columns_config({})
        return self.data.poz_columns_config

    def _poz_is_visible(self, key: str) -> bool:
        cfg = self._poz_cfg()
        try:
            return bool(cfg.get(key, {}).get("visible", True))
        except Exception:
            return True

    def _poz_set_visible(self, key: str, visible: bool):
        cfg = self._poz_cfg()
        if key not in cfg:
            cfg[key] = {"title": key, "visible": bool(visible)}
        else:
            try:
                cfg[key]["visible"] = bool(visible)
            except Exception:
                cfg[key] = {"title": str(cfg.get(key, {}).get("title", key)), "visible": bool(visible)}

    def _poz_set_title(self, key: str, title: str):
        cfg = self._poz_cfg()
        if key not in cfg:
            cfg[key] = {"title": str(title), "visible": True}
        else:
            try:
                cfg[key]["title"] = str(title)
            except Exception:
                cfg[key] = {"title": str(title), "visible": bool(cfg.get(key, {}).get("visible", True))}

    def _poz_apply_column_visibility(self):
        """Pielieto kolonnu redzamību tabulā + sinhronizē ar vecajiem iestatījumu laukiem."""
        if not hasattr(self, "tab") or not self.tab:
            return
        idx = self._poz_col_indices()
        cfg = self._poz_cfg()

        # Bāzes kolonnas
        base_keys = ["apraksts", "daudzums", "vieniba", "cena", "summa", "serial", "warranty", "notes"]
        for k in base_keys:
            col = idx.get(k)
            if col is None or col >= self.tab.columnCount():
                continue
            self.tab.setColumnHidden(col, not bool(cfg.get(k, {}).get("visible", True)))

        # Pielāgotās kolonnas (visible lauks katrai custom kolonnai)
        custom_start = idx.get("custom_start", 8)
        for i, col in enumerate(getattr(self.data, "custom_columns", []) or []):
            ui_col = custom_start + i
            if ui_col >= self.tab.columnCount():
                continue
            vis = True
            try:
                vis = bool(col.get("visible", True)) if isinstance(col, dict) else True
            except Exception:
                vis = True
            self.tab.setColumnHidden(ui_col, not vis)

        # Foto
        foto_col = idx.get("foto")
        if foto_col is not None and foto_col < self.tab.columnCount():
            self.tab.setColumnHidden(foto_col, not bool(cfg.get("foto", {}).get("visible", True)))

        # Atpakaļsavietojamība: vecie iestatījumi, ko izmanto PDF/iestatījumu tabs
        try:
            self.data.show_item_serial_number_in_table = bool(cfg.get("serial", {}).get("visible", True))
            self.data.show_item_warranty_in_table = bool(cfg.get("warranty", {}).get("visible", True))
            self.data.show_item_notes_in_table = bool(cfg.get("notes", {}).get("visible", True))
            self.data.show_item_photo_in_table = bool(cfg.get("foto", {}).get("visible", True))
        except Exception:
            pass

        # JAUNS: ja Foto kolonna ir ieslēgta, nodrošinām, ka pogas tiek atjaunotas visām rindām
        try:
            if bool(cfg.get("foto", {}).get("visible", True)):
                for r in range(self.tab.rowCount()):
                    self._ensure_photo_cell(r)
        except Exception:
            pass

        # Kopsavilkums zem tabulas
        try:
            if hasattr(self, "ck_show_price_summary") and self.ck_show_price_summary:
                self.data.show_price_summary = bool(self.ck_show_price_summary.isChecked())
        except Exception:
            pass

    def _poz_key_from_column(self, col: int) -> Optional[str]:
        idx = self._poz_col_indices()
        for k in ["apraksts", "daudzums", "vieniba", "cena", "summa", "serial", "warranty", "notes"]:
            if idx.get(k) == col:
                return k
        # custom
        cs = idx.get("custom_start", 8)
        foto = idx.get("foto")
        if foto is not None and col == foto:
            return "foto"
        if cs is not None and col >= cs and (foto is None or col < foto):
            return f"custom:{col - cs}"
        return None

    # --- JAUNS: droša kolonnu iestatījumu atvēršana pēc atlasītās kolonnas ---
    def _poz_open_settings_for_selected_column(self):
        """Atver kolonnu iestatījumu izvēlni tieši atlasītajai kolonnai.

        (Tas salabo situāciju, kad darbības attiecās uz nepareizu kolonnu, jo tika izmantotas
        nepareizas koordinātes.)
        """
        if not hasattr(self, "tab") or not self.tab:
            return
        hh = self.tab.horizontalHeader()
        col = self.tab.currentColumn()
        if col < 0:
            # ja nekas nav atlasīts, mēģinām atrast pirmo redzamo
            for c in range(self.tab.columnCount()):
                if not self.tab.isColumnHidden(c):
                    col = c
                    break
        if col < 0:
            return

        try:
            x = hh.sectionPosition(col) + max(6, int(hh.sectionSize(col) * 0.5))
            global_pos = hh.mapToGlobal(QPoint(x, hh.height()))
            self._poz_header_context_menu(QPoint(x, int(hh.height() / 2)), forced_col=col, global_pos=global_pos)
        except Exception:
            self._poz_header_context_menu(QPoint(0, 0), forced_col=col)

    def _poz_swap_custom_columns(self, i: int, j: int):
        """Samaina divas pielāgotās kolonnas (UI + datu sarakstā), saglabājot Foto vienmēr pēdējā."""
        try:
            customs = getattr(self.data, "custom_columns", []) or []
            if not (0 <= i < len(customs) and 0 <= j < len(customs)):
                return
            if i == j:
                return

            idx = self._poz_col_indices()
            cs = idx.get("custom_start", 8)
            ui_i = cs + i
            ui_j = cs + j
            if ui_i >= self.tab.columnCount() or ui_j >= self.tab.columnCount():
                return

            # swap definīcijas
            customs[i], customs[j] = customs[j], customs[i]
            self.data.custom_columns = customs

            # swap header tekstu
            hi = self.tab.horizontalHeaderItem(ui_i)
            hj = self.tab.horizontalHeaderItem(ui_j)
            ti = hi.text() if hi else ""
            tj = hj.text() if hj else ""
            self.tab.setHorizontalHeaderItem(ui_i, QTableWidgetItem(tj))
            self.tab.setHorizontalHeaderItem(ui_j, QTableWidgetItem(ti))

            # swap slēpšanas stāvokli
            hid_i = self.tab.isColumnHidden(ui_i)
            hid_j = self.tab.isColumnHidden(ui_j)

            # swap platumu
            try:
                w_i = self.tab.columnWidth(ui_i)
                w_j = self.tab.columnWidth(ui_j)
                self.tab.setColumnWidth(ui_i, w_j)
                self.tab.setColumnWidth(ui_j, w_i)
            except Exception:
                pass

            # swap šūnu saturu
            for r in range(self.tab.rowCount()):
                it_i = self.tab.takeItem(r, ui_i)
                it_j = self.tab.takeItem(r, ui_j)
                self.tab.setItem(r, ui_i, it_j)
                self.tab.setItem(r, ui_j, it_i)

                wgi = self.tab.cellWidget(r, ui_i)
                wgj = self.tab.cellWidget(r, ui_j)
                if wgi is not None or wgj is not None:
                    self.tab.removeCellWidget(r, ui_i)
                    self.tab.removeCellWidget(r, ui_j)
                    if wgj is not None:
                        self.tab.setCellWidget(r, ui_i, wgj)
                    if wgi is not None:
                        self.tab.setCellWidget(r, ui_j, wgi)

            self.tab.setColumnHidden(ui_i, hid_j)
            self.tab.setColumnHidden(ui_j, hid_i)

        except Exception:
            return

    def _poz_header_context_menu(self, pos: QPoint, forced_col: Optional[int] = None, global_pos: Optional[QPoint] = None):
        """Konteksta izvēlne kolonnu virsrakstiem: paslēpt/parādīt/pārdēvēt.

        `pos` ir header lokālajās koordinātēs. Ja izvēlne tiek atvērta no pogas,
        izmantojam `forced_col`, lai vienmēr strādātu ar atlasīto kolonnu.
        """
        if not hasattr(self, "tab") or not self.tab:
            return

        hh = self.tab.horizontalHeader()
        col = int(forced_col) if forced_col is not None else hh.logicalIndexAt(pos)
        if col < 0:
            col = -1

        menu = QMenu(self)

        # Paslēpt izvēlēto
        if col >= 0:
            key = self._poz_key_from_column(col)

            # JAUNS: kolonnu pārvietošana (strādā jebkurai kolonnai, arī starp default kolonnām)
            try:
                hh2 = self.tab.horizontalHeader()
                cur_visual = hh2.visualIndex(col)
            except Exception:
                hh2 = None
                cur_visual = -1

            if hh2 is not None and cur_visual >= 0:
                act_left = menu.addAction("Pārvietot pa kreisi")
                act_right = menu.addAction("Pārvietot pa labi")
                act_left.setEnabled(cur_visual > 0)
                act_right.setEnabled(cur_visual < self.tab.columnCount() - 1)

                def _mv_left():
                    try:
                        v = hh2.visualIndex(col)
                        if v > 0:
                            hh2.moveSection(v, v - 1)
                            self._poz_store_visual_order()
                            self._update_preview()
                    except Exception:
                        pass

                def _mv_right():
                    try:
                        v = hh2.visualIndex(col)
                        if 0 <= v < self.tab.columnCount() - 1:
                            hh2.moveSection(v, v + 1)
                            self._poz_store_visual_order()
                            self._update_preview()
                    except Exception:
                        pass

                act_left.triggered.connect(_mv_left)
                act_right.triggered.connect(_mv_right)

                # Pārvietot uz konkrētu pozīciju (jebkur)
                sub_mv = menu.addMenu("Pārvietot uz…")
                try:
                    titles = []
                    for visual in range(self.tab.columnCount()):
                        logical = hh2.logicalIndex(visual)
                        it = self.tab.horizontalHeaderItem(int(logical))
                        t = it.text() if it else str(self._poz_key_from_column(int(logical)) or "")
                        titles.append(t)

                    for target_visual, t in enumerate(titles):
                        act = sub_mv.addAction(f"{target_visual+1}. {t}")
                        def _make_move(tv=target_visual):
                            def _do():
                                try:
                                    v = hh2.visualIndex(col)
                                    if v >= 0 and tv >= 0:
                                        hh2.moveSection(v, tv)
                                        self._poz_store_visual_order()
                                        self._update_preview()
                                except Exception:
                                    pass
                            return _do
                        act.triggered.connect(_make_move())
                except Exception:
                    pass

                menu.addSeparator()
            act_hide = menu.addAction("Paslēpt šo kolonnu")
            def _hide_selected():
                if not key:
                    return
                if key.startswith("custom:"):
                    try:
                        i = int(key.split(":", 1)[1])
                        if 0 <= i < len(getattr(self.data, "custom_columns", []) or []):
                            self.data.custom_columns[i]["visible"] = False
                    except Exception:
                        pass
                else:
                    self._poz_set_visible(key, False)
                self._poz_apply_column_visibility()
                self._update_preview()
            act_hide.triggered.connect(_hide_selected)

            # Pārdēvēt
            act_rename = menu.addAction("Pārdēvēt kolonnu…")
            def _rename_selected():
                if not key:
                    return
                cur_title = ""
                try:
                    it = self.tab.horizontalHeaderItem(col)
                    cur_title = it.text() if it else ""
                except Exception:
                    cur_title = ""
                new_title, ok = QInputDialog.getText(self, "Pārdēvēt kolonnu", "Jaunais nosaukums:", text=cur_title)
                if not (ok and new_title.strip()):
                    return
                new_title = new_title.strip()

                if key.startswith("custom:"):
                    try:
                        i = int(key.split(":", 1)[1])
                        if 0 <= i < len(getattr(self.data, "custom_columns", []) or []):
                            self.data.custom_columns[i]["name"] = new_title
                    except Exception:
                        pass
                else:
                    self._poz_set_title(key, new_title)

                try:
                    self.tab.setHorizontalHeaderItem(col, QTableWidgetItem(new_title))
                except Exception:
                    pass
                self._update_preview()
            act_rename.triggered.connect(_rename_selected)

        # Parādīt paslēptās
        sub = menu.addMenu("Parādīt kolonnas")
        idx = self._poz_col_indices()
        # bāzes + foto
        for k in ["apraksts", "daudzums", "vieniba", "cena", "summa", "serial", "warranty", "notes", "foto"]:
            c = idx.get(k)
            if c is None or c >= self.tab.columnCount():
                continue
            hidden = self.tab.isColumnHidden(c)
            title = ""
            try:
                it = self.tab.horizontalHeaderItem(c)
                title = it.text() if it else k
            except Exception:
                title = k
            act = sub.addAction(title)
            act.setCheckable(True)
            act.setChecked(not hidden)
            def _toggle_factory(key=k, col_index=c):
                def _t(checked: bool):
                    self._poz_set_visible(key, bool(checked))
                    self._poz_apply_column_visibility()
                    self._update_preview()
                return _t
            act.toggled.connect(_toggle_factory())

        # custom kolonnas
        customs = getattr(self.data, "custom_columns", []) or []
        if customs:
            sub2 = menu.addMenu("Parādīt pielāgotās")
            cs = idx.get("custom_start", 8)
            for i, coldef in enumerate(customs):
                ui_col = cs + i
                if ui_col >= self.tab.columnCount():
                    continue
                title = ""
                try:
                    title = str(coldef.get("name", f"Kolonna {i+1}"))
                except Exception:
                    title = f"Kolonna {i+1}"
                vis = True
                try:
                    vis = bool(coldef.get("visible", True))
                except Exception:
                    vis = True
                act = sub2.addAction(title)
                act.setCheckable(True)
                act.setChecked(vis)
                def _toggle_custom_factory(i_=i):
                    def _t(checked: bool):
                        try:
                            if 0 <= i_ < len(self.data.custom_columns):
                                self.data.custom_columns[i_]["visible"] = bool(checked)
                        except Exception:
                            pass
                        self._poz_apply_column_visibility()
                        self._update_preview()
                    return _t
                act.toggled.connect(_toggle_custom_factory())

        # JAUNS: ja tiek padots precīzs globālais punkts (piem., no pogas), izmantojam to
        try:
            if isinstance(global_pos, QPoint):
                menu.exec(global_pos)
            else:
                menu.exec(hh.mapToGlobal(pos))
        except Exception:
            try:
                menu.exec(hh.mapToGlobal(pos))
            except Exception:
                pass


    
    def _row_from_photo_button(self, btn: QToolButton) -> int:
        """Droši nosaka rindu pēc foto pogas pozīcijas tabulas viewportā."""
        try:
            from PySide6.QtCore import QPoint
            vp = self.tab.viewport()
            p = btn.mapTo(vp, QPoint(btn.width() // 2, btn.height() // 2))
            idx = self.tab.indexAt(p)
            return idx.row()
        except Exception:
            return -1

    def _ensure_photo_cell(self, row: int):
        """Ieliek Foto izvēles/dzēšanas pogu konkrētai rindai (ja foto kolonna ir ieslēgta)."""
        # Foto kolonna vienmēr eksistē, bet var būt paslēpta
        try:
            if not self._poz_is_visible("foto"):
                return
        except Exception:
            pass
        col = self._poz_col_indices().get("foto")
        if col is None:
            return

        # Item glabā faila ceļu
        it = self.tab.item(row, col)
        if not it:
            it = QTableWidgetItem("")
            it.setFlags(it.flags() & ~Qt.ItemIsEditable)
            self.tab.setItem(row, col, it)

        # Ja jau ir ielikts widgets – nepārrakstām, tikai atsvaidzinām tekstu/ikonas
        w = self.tab.cellWidget(row, col)
        if isinstance(w, QToolButton):
            self._refresh_photo_button_ui(row, w)
            return

        btn = QToolButton()
        btn.setAutoRaise(True)
        btn.setToolTip("Pievienot vai dzēst foto šai pozīcijai")

        # Galvenais klikšķis = izvēlēties/mainīt
        btn.clicked.connect(lambda _=False, b=btn: self._choose_photo_for_button(b))

        # Menu ar "Izvēlēties/Mainīt" un "Dzēst"
        menu = QMenu(btn)
        act_choose = menu.addAction("Izvēlēties / Mainīt")
        act_choose.triggered.connect(lambda _=False, b=btn: self._choose_photo_for_button(b))
        act_clear = menu.addAction("Dzēst foto")
        act_clear.triggered.connect(lambda _=False, b=btn: self._clear_photo_for_button(b))
        btn.setMenu(menu)
        btn.setPopupMode(QToolButton.MenuButtonPopup)

        self.tab.setCellWidget(row, col, btn)
        self._refresh_photo_button_ui(row, btn)

    def _refresh_photo_button_ui(self, row: int, btn: QToolButton):
        """Atjauno foto pogas UI (tekstu/ikonu) balstoties uz šūnas ceļu."""
        col = self._poz_col_indices().get("foto")
        if col is None:
            return
        path = ""
        it = self.tab.item(row, col)
        if it:
            path = (it.text() or "").strip()

        if path and os.path.exists(path):
            btn.setText("Mainīt")
            try:
                ico = QIcon(path)
                if not ico.isNull():
                    btn.setIcon(ico)
                    btn.setIconSize(QSize(18, 18))
            except Exception:
                btn.setIcon(QIcon())
        else:
            btn.setText("Izvēlēties")
            btn.setIcon(QIcon())

        # Menu "Dzēst" – atslēdzam, ja nav ko dzēst
        if btn.menu():
            acts = btn.menu().actions()
            # pieņemam, ka otrā ir "Dzēst foto"
            for a in acts:
                if "Dzēst" in a.text():
                    a.setEnabled(bool(path))

    def _choose_photo_for_button(self, btn: QToolButton):
        row = self._row_from_photo_button(btn)
        if row < 0:
            return
        self._choose_photo_for_row(row)

    def _clear_photo_for_button(self, btn: QToolButton):
        row = self._row_from_photo_button(btn)
        if row < 0:
            return
        self._clear_photo_for_row(row)

    def _clear_photo_for_row(self, row: int):
        col = self._poz_col_indices().get("foto")
        if col is None:
            return

        it = self.tab.item(row, col)
        if not it:
            it = QTableWidgetItem("")
            it.setFlags(it.flags() & ~Qt.ItemIsEditable)
            self.tab.setItem(row, col, it)

        it.setText("")

        w = self.tab.cellWidget(row, col)
        if isinstance(w, QToolButton):
            self._refresh_photo_button_ui(row, w)

        self._update_preview()

    def _choose_photo_from_sender(self):
        btn = self.sender()
        if isinstance(btn, QToolButton):
            self._choose_photo_for_button(btn)

    def _choose_photo_for_row(self, row: int):
        col = self._poz_col_indices().get("foto")
        if col is None:
            return

        path, _ = QFileDialog.getOpenFileName(
            self, "Izvēlēties foto pozīcijai",
            "", "Attēli (*.png *.jpg *.jpeg *.webp *.bmp)"
        )
        if not path:
            return

        it = self.tab.item(row, col)
        if not it:
            it = QTableWidgetItem("")
            it.setFlags(it.flags() & ~Qt.ItemIsEditable)
            self.tab.setItem(row, col, it)
        it.setText(path)

        w = self.tab.cellWidget(row, col)
        if isinstance(w, QToolButton):
            self._refresh_photo_button_ui(row, w)

        self._update_preview()

    def pievienot_kolonnu(self):
        col_name, ok = QInputDialog.getText(self, "Pievienot kolonnu", "Ievadiet kolonnas nosaukumu:")
        if ok and col_name.strip():
            col_name = col_name.strip()

            # Pārbaudīt, vai kolonna jau eksistē
            headers = [self.tab.horizontalHeaderItem(i).text() for i in range(self.tab.columnCount())]
            if col_name in headers:
                QMessageBox.warning(self, "Kļūda", f"Kolonna '{col_name}' jau eksistē.")
                return

            # Foto vienmēr ir pēdējā kolonna -> insertējam jauno kolonnu PIRMS Foto
            idx = self._poz_col_indices()
            foto_col = idx.get("foto", self.tab.columnCount())
            self.tab.insertColumn(foto_col)
            self.tab.setHorizontalHeaderItem(foto_col, QTableWidgetItem(col_name))
            self.tab.horizontalHeader().setSectionResizeMode(foto_col, QHeaderView.Interactive)

            # Pievienojam datiem (atpakaļsavietojami: ja nav 'visible', pievienojam)
            new_col = {'name': col_name, 'data': [''] * self.tab.rowCount(), 'visible': True}
            self.data.custom_columns.append(new_col)

            # Atjaunojam Foto virsrakstu (jo tas ir pabīdīts pa labi)
            try:
                foto_idx_new = self._poz_col_indices().get("foto")
                if foto_idx_new is not None and foto_idx_new < self.tab.columnCount():
                    foto_title = str(self.data.poz_columns_config.get("foto", {}).get("title", "Foto"))
                    self.tab.setHorizontalHeaderItem(foto_idx_new, QTableWidgetItem(foto_title))
            except Exception:
                pass

            self._poz_apply_column_visibility()
            self._update_preview()

    def dzest_kolonnu(self):
        # Dzēšam tikai pielāgotās kolonnas (custom_columns), jo bāzes kolonnas var tikai paslēpt.
        customs = [c for c in (getattr(self.data, "custom_columns", []) or []) if isinstance(c, dict)]
        if not customs:
            QMessageBox.information(self, "Dzēst kolonnu", "Nav pielāgotu kolonnu, ko dzēst.")
            return

        names = [str(c.get("name", "")) for c in customs]
        col_name, ok = QInputDialog.getItem(self, "Dzēst kolonnu", "Izvēlieties kolonnu, ko dzēst:", names, 0, False)
        if not (ok and col_name):
            return

        # Noņemam no UI pēc indeksa (custom_start + custom_index)
        try:
            idx = self._poz_col_indices()
            custom_idx = names.index(col_name)
            ui_col = idx.get("custom_start", 8) + custom_idx
            if 0 <= ui_col < self.tab.columnCount():
                self.tab.removeColumn(ui_col)
        except Exception:
            pass

        # Noņemam no datiem
        for i, col in enumerate(self.data.custom_columns):
            if isinstance(col, dict) and str(col.get('name', '')) == str(col_name):
                del self.data.custom_columns[i]
                break

        # Atjaunojam Foto virsrakstu + redzamības iestatījumus
        try:
            foto_idx_new = self._poz_col_indices().get("foto")
            if foto_idx_new is not None and foto_idx_new < self.tab.columnCount():
                foto_title = str(self.data.poz_columns_config.get("foto", {}).get("title", "Foto"))
                self.tab.setHorizontalHeaderItem(foto_idx_new, QTableWidgetItem(foto_title))
        except Exception:
            pass

        self._poz_apply_column_visibility()
        self._update_preview()

    def pievienot_tukšu_pozīciju(self):
        r = self.tab.rowCount()
        self.tab.insertRow(r)

        # Inicializējam visas kolonnas
        for c in range(self.tab.columnCount()):
            self.tab.setItem(r, c, QTableWidgetItem(""))

        idx = self._poz_col_indices()

        # Noklusējumi
        if self.tab.item(r, idx["daudzums"]):
            self.tab.item(r, idx["daudzums"]).setText("1")
        if self.tab.item(r, idx["vieniba"]):
            self.tab.item(r, idx["vieniba"]).setText(self.data.default_unit)
        if self.tab.item(r, idx["cena"]):
            self.tab.item(r, idx["cena"]).setText("0.00")
        if self.tab.item(r, idx["summa"]):
            self.tab.item(r, idx["summa"]).setText("0.00")

        # Foto poga (ja ieslēgta)
        self._ensure_photo_cell(r)
    def pievienot_pozīcijas_no_noliktavas(self):
        try:
            db = getattr(self, "_noliktava", None)
            if db is None:
                QMessageBox.warning(self, "Noliktava nav pieejama", "Noliktavas datubāze nav inicializēta.")
                return
            dlg = ProductPickerDialog(self, db)
            if dlg.exec() != QDialog.Accepted:
                return
            for it, qty in dlg.get_selection():
                self._append_position_from_item(it, qty)
            self._update_preview()
        except Exception as e:
            QMessageBox.warning(self, "Kļūda", f"Neizdevās pievienot preces no noliktavas.\n\n{e}")

    def _append_position_from_item(self, it: NoliktavasPrece, qty: float):
        r = self.tab.rowCount()
        self.tab.insertRow(r)
        for c in range(self.tab.columnCount()):
            self.tab.setItem(r, c, QTableWidgetItem(""))

        idx = self._poz_col_indices()

        desc = (it.nosaukums or "").strip()
        sku = (it.sku or "").strip()
        if sku and desc:
            desc = f"[{sku}] {desc}"
        elif sku:
            desc = sku

        if self.tab.item(r, idx["apraksts"]):
            self.tab.item(r, idx["apraksts"]).setText(desc)
        if self.tab.item(r, idx["daudzums"]):
            try:
                self.tab.item(r, idx["daudzums"]).setText(("{:.3f}".format(float(qty))).rstrip("0").rstrip("."))
            except Exception:
                self.tab.item(r, idx["daudzums"]).setText(str(qty))
        if self.tab.item(r, idx["vieniba"]):
            self.tab.item(r, idx["vieniba"]).setText((it.vieniba or self.data.default_unit).strip() or self.data.default_unit)
        if self.tab.item(r, idx["cena"]):
            self.tab.item(r, idx["cena"]).setText(str(it.cena or ""))

        try:
            if "pvn" in idx and self.tab.item(r, idx["pvn"]):
                pvn = (it.pvn_likme or "").strip()
                if pvn:
                    self.tab.item(r, idx["pvn"]).setText(pvn)
        except Exception:
            pass

        try:
            foto_col = idx.get("foto")
            if foto_col is not None and self.tab.item(r, foto_col):
                self.tab.item(r, foto_col).setText(getattr(it, "foto_path", "") or "")
                try:
                    self._ensure_photo_cell(r)
                except Exception:
                    pass
        except Exception:
            pass

        try:
            self._aprēķināt_pozīciju_summa(r)
        except Exception:
            pass

    def pievienot_pozīciju(self):
        self.pievienot_pozīcijas_no_noliktavas()

    def dzest_pozīciju(self):
        r = self.tab.currentRow()
        if r >= 0:
            self.tab.removeRow(r)

    
    def _pārrēķināt_summa(self, row, col):
        """Pārrēķina 'Summa' pēc daudzuma un cenas, izmantojot dinamiskos kolonnu indeksus."""
        try:
            idx = self._poz_col_indices()
            col_daudz = idx.get("daudzums", 1)
            col_cena = idx.get("cena", 3)
            col_summa = idx.get("summa", 4)

            if col not in (col_daudz, col_cena):
                return

            daudz = to_decimal(self.tab.item(row, col_daudz).text()) if self.tab.item(row, col_daudz) else Decimal("0")
            cena = to_decimal(self.tab.item(row, col_cena).text()) if self.tab.item(row, col_cena) else Decimal("0")
            summa = (daudz * cena).quantize(Decimal("0.01"))

            if not self.tab.item(row, col_summa):
                self.tab.setItem(row, col_summa, QTableWidgetItem(""))
            self.tab.item(row, col_summa).setText(formēt_naudu(summa))
        finally:
            self._update_preview()

    # ----- Tab: Attēli -----
    
    # --- Tab: Attēli -----

    # ============================
    # Noliktava TAB (jauns)
    # ============================
    def _būvēt_noliktava_tab(self):
        w = QWidget()
        outer = QVBoxLayout(w)
        outer.setContentsMargins(0, 0, 0, 0)
        outer.setSpacing(0)

        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QFrame.NoFrame)
        try:
            scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarAsNeeded)
            scroll.setVerticalScrollBarPolicy(Qt.ScrollBarAsNeeded)
            scroll.setAlignment(Qt.AlignTop | Qt.AlignLeft)
        except Exception:
            pass

        content = QWidget()
        root = QVBoxLayout(content)
        root.setContentsMargins(8, 8, 8, 8)
        root.setSpacing(8)

        def _style_info_label(lbl):
            lbl.setStyleSheet("padding:8px 12px; border:1px solid #2d3750; border-radius:10px; background:#141a26; font-weight:600;")
            try:
                lbl.setWordWrap(True)
                lbl.setMinimumHeight(40)
                lbl.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
            except Exception:
                pass

        stats_grid = QGridLayout()
        stats_grid.setContentsMargins(0, 0, 0, 0)
        stats_grid.setHorizontalSpacing(8)
        stats_grid.setVerticalSpacing(8)
        self.lbl_nol_total_items = QLabel("Preču kartītes: 0")
        self.lbl_nol_total_qty = QLabel("Kopējais atlikums: 0")
        self.lbl_nol_total_value = QLabel("Krājumu vērtība: 0.00 EUR")
        self.lbl_nol_low_stock = QLabel("Zema atlikuma preces: 0")
        self.lbl_nol_warehouses = QLabel("Noliktavas: 0")
        stats_labels = [self.lbl_nol_total_items, self.lbl_nol_total_qty, self.lbl_nol_total_value, self.lbl_nol_low_stock, self.lbl_nol_warehouses]
        for lbl in stats_labels:
            _style_info_label(lbl)
        stats_grid.addWidget(self.lbl_nol_total_items, 0, 0)
        stats_grid.addWidget(self.lbl_nol_total_qty, 0, 1)
        stats_grid.addWidget(self.lbl_nol_total_value, 0, 2)
        stats_grid.addWidget(self.lbl_nol_low_stock, 1, 0)
        stats_grid.addWidget(self.lbl_nol_warehouses, 1, 1)
        stats_grid.setColumnStretch(0, 1)
        stats_grid.setColumnStretch(1, 1)
        stats_grid.setColumnStretch(2, 1)
        root.addLayout(stats_grid)

        filter_box = QGroupBox("Meklēšana un filtri")
        filter_grid = QGridLayout(filter_box)
        filter_grid.setContentsMargins(10, 12, 10, 10)
        filter_grid.setHorizontalSpacing(8)
        filter_grid.setVerticalSpacing(8)
        self.in_noliktava_search = QLineEdit()
        self.in_noliktava_search.setPlaceholderText("Meklēt pēc SKU, nosaukuma, svītrkoda, partijas, piegādātāja vai noliktavas")
        self.cb_noliktava_filter = QComboBox()
        self.cb_noliktava_filter.addItems(["Aktīvās preces", "Zems atlikums", "Visas preces"])
        self.cb_noliktava_warehouse = QComboBox()
        self.cb_noliktava_warehouse.addItem("Visas noliktavas")
        self.chk_auto_remove_depleted = QCheckBox("Pie 0 atlikuma izņemt no aktīvās noliktavas")
        self.chk_auto_remove_depleted.setChecked(True)
        try:
            self.in_noliktava_search.setMinimumHeight(36)
            self.cb_noliktava_filter.setMinimumHeight(36)
            self.cb_noliktava_warehouse.setMinimumHeight(36)
        except Exception:
            pass
        filter_grid.addWidget(QLabel("Meklēt:"), 0, 0)
        filter_grid.addWidget(self.in_noliktava_search, 0, 1, 1, 3)
        filter_grid.addWidget(QLabel("Skats:"), 1, 0)
        filter_grid.addWidget(self.cb_noliktava_filter, 1, 1)
        filter_grid.addWidget(QLabel("Noliktava:"), 1, 2)
        filter_grid.addWidget(self.cb_noliktava_warehouse, 1, 3)
        filter_grid.addWidget(self.chk_auto_remove_depleted, 2, 0, 1, 4)
        filter_grid.setColumnStretch(1, 1)
        filter_grid.setColumnStretch(3, 1)
        root.addWidget(filter_box)

        actions_box = QGroupBox("Darbības")
        actions_grid = QGridLayout(actions_box)
        actions_grid.setContentsMargins(10, 12, 10, 10)
        actions_grid.setHorizontalSpacing(8)
        actions_grid.setVerticalSpacing(8)
        btn_refresh = QPushButton("Atsvaidzināt")
        btn_add = QPushButton("Saglabāt kartīti")
        btn_edit = QPushButton("Rediģēt / dokumenti")
        btn_clear = QPushButton("Jauna kartīte")
        btn_delete = QPushButton("Dzēst kartīti(s)")
        btn_duplicate = QPushButton("Dublēt kartīti(s)")
        btn_to_poz = QPushButton("Uz pozīcijām")
        btn_issue = QPushButton("Izrakstīt")
        btn_receive = QPushButton("Saņemt")
        btn_adjust = QPushButton("Korekcija")
        btn_import = QPushButton("Imports CSV / Excel")
        btn_export = QPushButton("Eksports CSV")
        btn_export_xlsx = QPushButton("Eksports Excel")
        btn_export_pdf = QPushButton("Eksports PDF")
        btn_export_moves = QPushButton("Kustību žurnāls")
        button_tips = {
            btn_refresh: "Pārlādēt noliktavas sarakstu un kopsavilkumu",
            btn_add: "Saglabāt vai atjaunināt izvēlēto preces kartīti",
            btn_edit: "Ielādēt atlasīto produktu rediģēšanai un dokumentu pievienošanai",
            btn_clear: "Notīrīt formu un veidot jaunu preces kartīti",
            btn_delete: "Dzēst atlasītās preces kartītes no noliktavas",
            btn_duplicate: "Izveidot atlasīto noliktavas kartīšu kopijas",
            btn_to_poz: "Pārnest atlasīto preci uz Pozīciju tabu ar kolonnu izvēli",
            btn_issue: "Samazināt atlikumu / izrakstīt no noliktavas",
            btn_receive: "Palielināt atlikumu noliktavā",
            btn_adjust: "Veikt inventarizācijas korekciju",
            btn_import: "Importēt noliktavas datus no CSV vai Excel",
            btn_export: "Eksportēt noliktavu CSV formātā",
            btn_export_xlsx: "Eksportēt noliktavu Excel formātā",
            btn_export_pdf: "Eksportēt noliktavu PDF formātā",
            btn_export_moves: "Eksportēt kustību žurnālu CSV formātā",
        }
        action_buttons = [
            btn_refresh, btn_add, btn_edit, btn_clear,
            btn_delete, btn_duplicate, btn_to_poz, btn_issue, btn_receive,
            btn_adjust, btn_import, btn_export, btn_export_xlsx,
            btn_export_pdf, btn_export_moves,
        ]
        for b, tip in button_tips.items():
            b.setToolTip(tip)
            try:
                b.setMinimumHeight(38)
                b.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
            except Exception:
                pass
        for idx, b in enumerate(action_buttons):
            actions_grid.addWidget(b, idx // 4, idx % 4)
        for col in range(4):
            actions_grid.setColumnStretch(col, 1)
        root.addWidget(actions_box)

        table_box = QGroupBox("Preču saraksts")
        table_layout = QVBoxLayout(table_box)
        table_layout.setContentsMargins(10, 12, 10, 10)
        self.tbl_noliktava = QTableWidget(0, 17)
        self.tbl_noliktava.setHorizontalHeaderLabels([
            "SKU", "Svītrkods", "Nosaukums", "Kategorija", "Noliktava", "Atrašanās vieta", "Vienība",
            "Iepirkuma valūta", "Cena", "PVN %", "Atlikums", "Min. atlik.", "Pavadzīme Nr.",
            "Piegādātājs", "Statuss", "Piezīmes", "Foto"
        ])
        self.tbl_noliktava.horizontalHeader().setStretchLastSection(False)
        self.tbl_noliktava.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self.tbl_noliktava.setSelectionMode(QAbstractItemView.SelectionMode.ExtendedSelection)
        self.tbl_noliktava.setAlternatingRowColors(True)
        self.tbl_noliktava.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        try:
            self.tbl_noliktava.verticalHeader().setVisible(False)
            self.tbl_noliktava.setMinimumHeight(170)
            self.tbl_noliktava.setMaximumHeight(220)
            hdr = self.tbl_noliktava.horizontalHeader()
            hdr.setSectionResizeMode(0, QHeaderView.ResizeToContents)
            hdr.setSectionResizeMode(1, QHeaderView.ResizeToContents)
            hdr.setSectionResizeMode(2, QHeaderView.Stretch)
            for i in range(3, 15):
                hdr.setSectionResizeMode(i, QHeaderView.ResizeToContents)
            hdr.setSectionResizeMode(15, QHeaderView.Stretch)
            hdr.setSectionResizeMode(16, QHeaderView.ResizeToContents)
        except Exception:
            pass
        try:
            self.tbl_noliktava.setContextMenuPolicy(Qt.CustomContextMenu)
            def _open_inventory_context_menu(pos):
                menu = QMenu(self.tbl_noliktava)
                act_edit = menu.addAction('Rediģēt kartīti')
                act_docs = menu.addAction('Pievienot dokumentus')
                act_open_doc = menu.addAction('Atvērt izvēlēto dokumentu')
                chosen = menu.exec(self.tbl_noliktava.viewport().mapToGlobal(pos))
                if chosen == act_edit:
                    _edit_selected_item(False)
                elif chosen == act_docs:
                    _edit_selected_item(True)
                elif chosen == act_open_doc:
                    self._inventory_open_selected_product_doc()
            self.tbl_noliktava.customContextMenuRequested.connect(_open_inventory_context_menu)
        except Exception:
            pass
        table_layout.addWidget(self.tbl_noliktava)
        root.addWidget(table_box)

        editor_box = QGroupBox("Preces kartīte")
        editor_root = QVBoxLayout(editor_box)
        editor_root.setContentsMargins(10, 12, 10, 10)
        editor_root.setSpacing(8)

        warehouse_box = QGroupBox("Noliktavu saraksts")
        warehouse_grid = QGridLayout(warehouse_box)
        warehouse_grid.setContentsMargins(8, 8, 8, 8)
        warehouse_grid.setHorizontalSpacing(8)
        warehouse_grid.setVerticalSpacing(8)
        self.cb_saved_warehouses = QComboBox()
        self.cb_saved_warehouses.setEditable(True)
        self.cb_saved_warehouses.setInsertPolicy(QComboBox.NoInsert)
        btn_add_wh = QPushButton("Saglabāt noliktavu")
        btn_del_wh = QPushButton("Dzēst noliktavu")
        btn_apply_wh = QPushButton("Ielikt kartītē")
        for b in (btn_add_wh, btn_del_wh, btn_apply_wh):
            try:
                b.setMinimumHeight(36)
                b.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
            except Exception:
                pass
        warehouse_grid.addWidget(QLabel("Noliktavas nosaukums:"), 0, 0)
        warehouse_grid.addWidget(self.cb_saved_warehouses, 0, 1, 1, 3)
        warehouse_grid.addWidget(btn_add_wh, 1, 1)
        warehouse_grid.addWidget(btn_del_wh, 1, 2)
        warehouse_grid.addWidget(btn_apply_wh, 1, 3)
        warehouse_grid.setColumnStretch(1, 1)
        editor_root.addWidget(warehouse_box)

        self.in_prod_sku = QLineEdit()
        self.in_prod_barcode = QLineEdit()
        self.in_prod_name = QLineEdit()
        self.in_prod_category = QLineEdit()
        self.in_prod_subcategory = QLineEdit()
        self.in_prod_unit = QLineEdit(); self.in_prod_unit.setText("gab.")
        self.in_prod_price = QLineEdit()
        self.in_prod_currency = QComboBox(); self.in_prod_currency.addItems(["EUR", "USD", "GBP", "SEK", "NOK"])
        self.in_prod_vat = QLineEdit(); self.in_prod_vat.setPlaceholderText("21")
        self.in_prod_stock = QDoubleSpinBox(); self.in_prod_stock.setRange(-1e9, 1e9); self.in_prod_stock.setDecimals(3)
        self.in_prod_min_stock = QDoubleSpinBox(); self.in_prod_min_stock.setRange(0, 1e9); self.in_prod_min_stock.setDecimals(3)
        self.in_prod_warehouse = QComboBox(); self.in_prod_warehouse.setEditable(True); self.in_prod_warehouse.setInsertPolicy(QComboBox.NoInsert)
        self.in_prod_location = QLineEdit()
        self.in_prod_supplier = QLineEdit()
        self.in_prod_manufacturer = QLineEdit()
        self.in_prod_delivery_doc = QLineEdit()
        self.in_prod_batch = QLineEdit()
        self.in_prod_serial = QLineEdit()
        self.in_prod_purchase_date = QDateEdit(calendarPopup=True); self.in_prod_purchase_date.setDate(QDate.currentDate())
        self.in_prod_expiry_date = QDateEdit(calendarPopup=True); self.in_prod_expiry_date.setDate(QDate.currentDate())
        self.cb_prod_status = QComboBox(); self.cb_prod_status.addItems(["Aktīva", "Rezervēta", "Norakstāma", "Arhivēta"])
        self.in_prod_notes = QTextEdit(); self.in_prod_notes.setFixedHeight(120)
        self.in_prod_photo = QLineEdit(); self.in_prod_photo.setPlaceholderText("Foto ceļš")
        self.in_supplier_code = QLineEdit()
        self.in_supplier_email = QLineEdit()
        self.in_supplier_phone = QLineEdit()
        self.in_supplier_api_url = QLineEdit()
        self.in_supplier_product_url = QLineEdit()
        self.in_supplier_lead_days = QSpinBox(); self.in_supplier_lead_days.setRange(0, 3650)
        self.ck_preferred_supplier = QCheckBox("Primārais piegādātājs")
        self.in_prod_hs_code = QLineEdit()
        self.in_prod_origin_country = QLineEdit()
        self.in_prod_net_weight = QLineEdit()
        self.in_prod_gross_weight = QLineEdit()
        self.in_prod_docs_folder = QLineEdit(); self.in_prod_docs_folder.setPlaceholderText("Dokumentu mape")
        self.list_prod_docs = QListWidget()
        self.list_prod_docs.setAlternatingRowColors(True)
        self.list_prod_docs.itemDoubleClicked.connect(lambda *_: self._inventory_open_selected_product_doc())

        photo_row = QWidget()
        photo_row_l = QHBoxLayout(photo_row)
        photo_row_l.setContentsMargins(0, 0, 0, 0)
        photo_row_l.setSpacing(6)
        btn_pick_prod_photo = QPushButton("Izvēlēties foto")
        btn_clear_prod_photo = QPushButton("Dzēst foto")
        btn_pick_prod_photo.clicked.connect(lambda: self._choose_inventory_photo())
        btn_clear_prod_photo.clicked.connect(lambda: self.in_prod_photo.setText(""))
        photo_row_l.addWidget(self.in_prod_photo, 1)
        photo_row_l.addWidget(btn_pick_prod_photo)
        photo_row_l.addWidget(btn_clear_prod_photo)

        card_tabs = QTabWidget()
        try:
            card_tabs.setUsesScrollButtons(True)
        except Exception:
            pass

        page_main = QWidget()
        page_main_l = QVBoxLayout(page_main)
        page_main_l.setContentsMargins(6, 6, 6, 6)
        main_form = QFormLayout()
        main_form.setLabelAlignment(Qt.AlignLeft)
        main_form.setFormAlignment(Qt.AlignTop)
        main_form.setFieldGrowthPolicy(QFormLayout.AllNonFixedFieldsGrow)
        main_form.addRow("SKU", self.in_prod_sku)
        main_form.addRow("Svītrkods", self.in_prod_barcode)
        main_form.addRow("Nosaukums", self.in_prod_name)
        main_form.addRow("Kategorija", self.in_prod_category)
        main_form.addRow("Apakškategorija", self.in_prod_subcategory)
        main_form.addRow("Vienība", self.in_prod_unit)
        main_form.addRow("Cena", self.in_prod_price)
        main_form.addRow("Iepirkuma valūta", self.in_prod_currency)
        main_form.addRow("PVN %", self.in_prod_vat)
        main_form.addRow("Atlikums", self.in_prod_stock)
        page_main_l.addLayout(main_form)
        page_main_l.addStretch(1)

        page_log = QWidget()
        page_log_l = QVBoxLayout(page_log)
        page_log_l.setContentsMargins(6, 6, 6, 6)
        log_form = QFormLayout()
        log_form.setLabelAlignment(Qt.AlignLeft)
        log_form.setFormAlignment(Qt.AlignTop)
        log_form.setFieldGrowthPolicy(QFormLayout.AllNonFixedFieldsGrow)
        log_form.addRow("Minimālais atlikums", self.in_prod_min_stock)
        log_form.addRow("Noliktavas nosaukums", self.in_prod_warehouse)
        log_form.addRow("Atrašanās vieta", self.in_prod_location)
        log_form.addRow("Piegādātājs", self.in_prod_supplier)
        log_form.addRow("Ražotājs", self.in_prod_manufacturer)
        log_form.addRow("Pavadzīmes Nr.", self.in_prod_delivery_doc)
        log_form.addRow("Partijas Nr.", self.in_prod_batch)
        log_form.addRow("Sērijas Nr.", self.in_prod_serial)
        log_form.addRow("Iepirkuma datums", self._wrap_date_with_system_button(self.in_prod_purchase_date))
        log_form.addRow("Derīguma termiņš", self._wrap_date_with_system_button(self.in_prod_expiry_date))
        log_form.addRow("Statuss", self.cb_prod_status)
        page_log_l.addLayout(log_form)
        page_log_l.addStretch(1)

        page_notes = QWidget()
        page_notes_l = QVBoxLayout(page_notes)
        page_notes_l.setContentsMargins(6, 6, 6, 6)
        page_notes_l.setSpacing(8)
        page_notes_l.addWidget(QLabel("Piezīmes"))
        page_notes_l.addWidget(self.in_prod_notes)
        page_notes_l.addWidget(QLabel("Foto"))
        page_notes_l.addWidget(photo_row)

        page_supplier = QWidget()
        page_supplier_l = QVBoxLayout(page_supplier)
        page_supplier_l.setContentsMargins(6, 6, 6, 6)
        supplier_form = QFormLayout()
        supplier_form.setLabelAlignment(Qt.AlignLeft)
        supplier_form.setFormAlignment(Qt.AlignTop)
        supplier_form.setFieldGrowthPolicy(QFormLayout.AllNonFixedFieldsGrow)
        supplier_form.addRow("Piegādātāja kods", self.in_supplier_code)
        supplier_form.addRow("Piegādātāja e-pasts", self.in_supplier_email)
        supplier_form.addRow("Piegādātāja tālrunis", self.in_supplier_phone)
        supplier_form.addRow("API / katalogs URL", self.in_supplier_api_url)
        supplier_form.addRow("Produkta URL", self.in_supplier_product_url)
        supplier_form.addRow("Piegādes laiks (dienas)", self.in_supplier_lead_days)
        supplier_form.addRow("", self.ck_preferred_supplier)
        supplier_form.addRow("HS kods", self.in_prod_hs_code)
        supplier_form.addRow("Izcelsmes valsts", self.in_prod_origin_country)
        supplier_form.addRow("Neto svars", self.in_prod_net_weight)
        supplier_form.addRow("Bruto svars", self.in_prod_gross_weight)
        page_supplier_l.addLayout(supplier_form)
        page_supplier_l.addStretch(1)

        page_docs = QWidget()
        page_docs_l = QVBoxLayout(page_docs)
        page_docs_l.setContentsMargins(6, 6, 6, 6)
        docs_form = QFormLayout()
        docs_form.setLabelAlignment(Qt.AlignLeft)
        docs_form.addRow("Dokumentu mape", self.in_prod_docs_folder)
        page_docs_l.addLayout(docs_form)
        docs_btn_row = QHBoxLayout()
        btn_add_prod_doc = QPushButton("Pievienot dokumentus")
        btn_remove_prod_doc = QPushButton("Dzēst izvēlēto")
        btn_open_prod_doc = QPushButton("Atvērt")
        btn_add_prod_doc.clicked.connect(self._inventory_add_product_docs)
        btn_remove_prod_doc.clicked.connect(self._inventory_remove_selected_product_doc)
        btn_open_prod_doc.clicked.connect(self._inventory_open_selected_product_doc)
        docs_btn_row.addWidget(btn_add_prod_doc)
        docs_btn_row.addWidget(btn_remove_prod_doc)
        docs_btn_row.addWidget(btn_open_prod_doc)
        docs_btn_row.addStretch(1)
        page_docs_l.addLayout(docs_btn_row)
        page_docs_l.addWidget(self.list_prod_docs, 1)

        card_tabs.addTab(page_main, "Pamatdati")
        card_tabs.addTab(page_log, "Loģistika")
        card_tabs.addTab(page_supplier, "Piegādātāji / muita")
        card_tabs.addTab(page_notes, "Piezīmes un foto")
        card_tabs.addTab(page_docs, "Dokumenti")
        editor_root.addWidget(card_tabs)
        root.addWidget(editor_box)
        root.addStretch(1)

        def _safe_text(item):
            return item.text() if item else ""

        def _refresh_warehouse_lists():
            try:
                names = self._noliktava.warehouse_names()
                current_filter = self.cb_noliktava_warehouse.currentText()
                current_editor = self.in_prod_warehouse.currentText()
                current_saved = self.cb_saved_warehouses.currentText()
                for combo, first in ((self.cb_noliktava_warehouse, 'Visas noliktavas'), (self.in_prod_warehouse, None), (self.cb_saved_warehouses, None)):
                    combo.blockSignals(True)
                    combo.clear()
                    if first:
                        combo.addItem(first)
                    for name in names:
                        combo.addItem(name)
                    combo.blockSignals(False)
                ix = self.cb_noliktava_warehouse.findText(current_filter)
                self.cb_noliktava_warehouse.setCurrentIndex(ix if ix >= 0 else 0)
                if current_editor:
                    self.in_prod_warehouse.setCurrentText(current_editor)
                elif self.in_prod_warehouse.count() and not self.in_prod_warehouse.currentText():
                    self.in_prod_warehouse.setCurrentText(self.in_prod_warehouse.itemText(0))
                if current_saved:
                    self.cb_saved_warehouses.setCurrentText(current_saved)
            except Exception:
                pass

        def _clear_editor():
            try:
                self.in_prod_sku.clear(); self.in_prod_barcode.clear(); self.in_prod_name.clear(); self.in_prod_category.clear(); self.in_prod_subcategory.clear()
                self.in_prod_unit.setText('gab.'); self.in_prod_price.clear(); self.in_prod_currency.setCurrentText('EUR'); self.in_prod_vat.clear()
                self.in_prod_stock.setValue(0.0); self.in_prod_min_stock.setValue(0.0); self.in_prod_warehouse.setCurrentText('Pamatnoliktava'); self.in_prod_location.clear()
                self.in_prod_supplier.clear(); self.in_prod_manufacturer.clear(); self.in_prod_delivery_doc.clear(); self.in_prod_batch.clear(); self.in_prod_serial.clear()
                self.in_prod_purchase_date.setDate(QDate.currentDate()); self.in_prod_expiry_date.setDate(QDate.currentDate()); self.cb_prod_status.setCurrentText('Aktīva')
                self.in_prod_notes.setPlainText(''); self.in_prod_photo.clear(); self.in_supplier_code.clear(); self.in_supplier_email.clear(); self.in_supplier_phone.clear()
                self.in_supplier_api_url.clear(); self.in_supplier_product_url.clear(); self.in_supplier_lead_days.setValue(0); self.ck_preferred_supplier.setChecked(False)
                self._current_inventory_edit_key = ''
                self.in_prod_hs_code.clear(); self.in_prod_origin_country.clear(); self.in_prod_net_weight.clear(); self.in_prod_gross_weight.clear(); self.in_prod_docs_folder.clear(); self.list_prod_docs.clear()
            except Exception:
                pass

        def _fill_editor_from_selected():
            try:
                r = self.tbl_noliktava.currentRow()
                if r < 0:
                    return
                row_item = self.tbl_noliktava.item(r, 0)
                row_key = row_item.data(Qt.UserRole) if row_item is not None else ''
                sku = _safe_text(row_item)
                it = self._noliktava.get_by_inventory_id(row_key) if row_key else None
                if it is None:
                    it = self._noliktava.get_by_sku(sku)
                if it is None:
                    return
                self._current_inventory_edit_key = getattr(it, 'inventory_id', '') or ''
                self.in_prod_sku.setText(it.sku)
                self.in_prod_barcode.setText(getattr(it, 'svitrkods', '') or '')
                self.in_prod_name.setText(it.nosaukums)
                self.in_prod_category.setText(getattr(it, 'kategorija', '') or '')
                self.in_prod_subcategory.setText(getattr(it, 'apakskategorija', '') or '')
                self.in_prod_unit.setText(it.vieniba or 'gab.')
                self.in_prod_currency.setCurrentText(getattr(it, 'iepirkuma_valuta', 'EUR') or 'EUR')
                self.in_prod_price.setText(it.cena)
                self.in_prod_vat.setText(it.pvn_likme)
                self.in_prod_stock.setValue(float(getattr(it, 'atlikums', 0.0) or 0.0))
                self.in_prod_min_stock.setValue(float(getattr(it, 'minimalais_atlikums', 0.0) or 0.0))
                self.in_prod_warehouse.setCurrentText(getattr(it, 'noliktavas_nosaukums', '') or '')
                self.in_prod_location.setText(getattr(it, 'atrasanas_vieta', '') or '')
                self.in_prod_supplier.setText(getattr(it, 'piegadatajs', '') or '')
                self.in_prod_manufacturer.setText(getattr(it, 'razotajs', '') or '')
                self.in_prod_delivery_doc.setText(getattr(it, 'pavadzimes_numurs', '') or '')
                self.in_prod_batch.setText(getattr(it, 'partijas_numurs', '') or '')
                self.in_prod_serial.setText(getattr(it, 'serialais_numurs', '') or '')
                try:
                    if getattr(it, 'iepirkuma_datums', ''):
                        self.in_prod_purchase_date.setDate(QDate.fromString(getattr(it, 'iepirkuma_datums', ''), 'yyyy-MM-dd'))
                except Exception:
                    pass
                try:
                    if getattr(it, 'deriguma_termiņš', ''):
                        self.in_prod_expiry_date.setDate(QDate.fromString(getattr(it, 'deriguma_termiņš', ''), 'yyyy-MM-dd'))
                except Exception:
                    pass
                self.cb_prod_status.setCurrentText(getattr(it, 'statuss', 'Aktīva') or 'Aktīva')
                self.in_prod_notes.setPlainText(getattr(it, 'piezimes', '') or '')
                self.in_prod_photo.setText(getattr(it, 'foto_path', '') or '')
                self.in_supplier_code.setText(getattr(it, 'supplier_code', '') or '')
                self.in_supplier_email.setText(getattr(it, 'supplier_email', '') or '')
                self.in_supplier_phone.setText(getattr(it, 'supplier_phone', '') or '')
                self.in_supplier_api_url.setText(getattr(it, 'supplier_api_url', '') or '')
                self.in_supplier_product_url.setText(getattr(it, 'supplier_product_url', '') or '')
                self.in_supplier_lead_days.setValue(int(getattr(it, 'supplier_lead_time_days', 0) or 0))
                self.ck_preferred_supplier.setChecked(bool(getattr(it, 'preferred_supplier', False)))
                self.in_prod_hs_code.setText(getattr(it, 'hs_kods', '') or '')
                self.in_prod_origin_country.setText(getattr(it, 'izcelsmes_valsts', '') or '')
                self.in_prod_net_weight.setText(getattr(it, 'neto_svars', '') or '')
                self.in_prod_gross_weight.setText(getattr(it, 'bruto_svars', '') or '')
                self.in_prod_docs_folder.setText(getattr(it, 'dokumentu_mape', '') or '')
                self._inventory_set_doc_list(getattr(it, 'dokumenti', []) or [])
            except Exception:
                pass

        self.tbl_noliktava.itemSelectionChanged.connect(_fill_editor_from_selected)

        def _edit_selected_item(open_docs_tab: bool = True):
            try:
                r = self.tbl_noliktava.currentRow()
                if r < 0:
                    QMessageBox.information(self, 'Noliktava', 'Atlasiet preces kartīti, kuru rediģēt.')
                    return
                _fill_editor_from_selected()
                try:
                    if open_docs_tab:
                        card_tabs.setCurrentIndex(4)
                    else:
                        card_tabs.setCurrentIndex(0)
                except Exception:
                    pass
                try:
                    editor_box.setFocus()
                except Exception:
                    pass
            except Exception as e:
                QMessageBox.critical(self, 'Kļūda', f'Neizdevās atvērt kartīti rediģēšanai: {e}')

        self.tbl_noliktava.cellDoubleClicked.connect(lambda *_: _edit_selected_item(True))

        def _apply_inventory_row_style(row: int, stock_value: float, min_value: float):
            try:
                if stock_value <= 0:
                    bg = QColor('#3a1820')
                elif stock_value <= max(min_value, 0.0):
                    bg = QColor('#3a2f12')
                else:
                    bg = QColor('#142033')
                for c in range(self.tbl_noliktava.columnCount()):
                    item = self.tbl_noliktava.item(row, c)
                    if item is not None:
                        item.setBackground(bg)
            except Exception:
                pass

        def _refresh_summary():
            try:
                include_zero = self.cb_noliktava_filter.currentText() == 'Visas preces'
                s = self._noliktava.summary(include_zero=include_zero)
                self.lbl_nol_total_items.setText(f"Preču kartītes: {int(s['total_items'])}")
                self.lbl_nol_total_qty.setText(f"Kopējais atlikums: {s['total_qty']:.3f}".rstrip('0').rstrip('.'))
                self.lbl_nol_total_value.setText(f"Krājumu vērtība: {s['total_value']:.2f} EUR")
                self.lbl_nol_low_stock.setText(f"Zema atlikuma preces: {int(s['low_stock'])}")
                self.lbl_nol_warehouses.setText(f"Noliktavas: {int(s.get('warehouse_count', 0))}")
            except Exception:
                pass

        def _refresh_table():
            q = (self.in_noliktava_search.text() or '').strip()
            mode = self.cb_noliktava_filter.currentText()
            include_zero = mode == 'Visas preces'
            items = self._noliktava.find(q, include_zero=include_zero)
            wh = self.cb_noliktava_warehouse.currentText()
            if wh and wh != 'Visas noliktavas':
                items = [it for it in items if (getattr(it, 'noliktavas_nosaukums', '') or '') == wh]
            if mode == 'Zems atlikums':
                items = [it for it in items if float(getattr(it, 'atlikums', 0.0) or 0.0) <= max(float(getattr(it, 'minimalais_atlikums', 0.0) or 0.0), 3.0)]
            self.tbl_noliktava.setRowCount(0)
            for it in items:
                r = self.tbl_noliktava.rowCount()
                self.tbl_noliktava.insertRow(r)
                vals = [
                    getattr(it, 'sku', ''), getattr(it, 'svitrkods', ''), getattr(it, 'nosaukums', ''), getattr(it, 'kategorija', ''),
                    getattr(it, 'noliktavas_nosaukums', ''), getattr(it, 'atrasanas_vieta', ''), getattr(it, 'vieniba', ''), getattr(it, 'iepirkuma_valuta', ''),
                    getattr(it, 'cena', ''), getattr(it, 'pvn_likme', ''), str(getattr(it, 'atlikums', '')), str(getattr(it, 'minimalais_atlikums', '')),
                    getattr(it, 'pavadzimes_numurs', ''), getattr(it, 'piegadatajs', ''), getattr(it, 'statuss', ''), getattr(it, 'piezimes', ''), getattr(it, 'foto_path', '')
                ]
                row_key = getattr(it, 'inventory_id', '') or ''
                for c, val in enumerate(vals):
                    cell = QTableWidgetItem(str(val or ''))
                    if c == 0:
                        cell.setData(Qt.UserRole, row_key)
                    self.tbl_noliktava.setItem(r, c, cell)
                try:
                    _apply_inventory_row_style(r, float(getattr(it, 'atlikums', 0.0) or 0.0), float(getattr(it, 'minimalais_atlikums', 0.0) or 0.0))
                except Exception:
                    pass
            _refresh_summary()
            _refresh_warehouse_lists()

        def _save_item():
            try:
                sku = (self.in_prod_sku.text() or '').strip()
                if not sku:
                    QMessageBox.warning(self, 'Noliktava', 'Lūdzu ievadiet SKU.')
                    return
                item = NoliktavasPrece(
                    sku=sku,
                    nosaukums=(self.in_prod_name.text() or '').strip(),
                    vieniba=(self.in_prod_unit.text() or 'gab.').strip() or 'gab.',
                    cena=(self.in_prod_price.text() or '').strip(),
                    pvn_likme=(self.in_prod_vat.text() or '').strip(),
                    serialais_numurs=(self.in_prod_serial.text() or '').strip(),
                    piezimes=self.in_prod_notes.toPlainText().strip(),
                    foto_path=(self.in_prod_photo.text() or '').strip(),
                    svitrkods=(self.in_prod_barcode.text() or '').strip(),
                    kategorija=(self.in_prod_category.text() or '').strip(),
                    apakskategorija=(self.in_prod_subcategory.text() or '').strip(),
                    iepirkuma_valuta=(self.in_prod_currency.currentText() or 'EUR').strip() or 'EUR',
                    noliktavas_nosaukums=(self.in_prod_warehouse.currentText() or '').strip(),
                    atrasanas_vieta=(self.in_prod_location.text() or '').strip(),
                    piegadatajs=(self.in_prod_supplier.text() or '').strip(),
                    razotajs=(self.in_prod_manufacturer.text() or '').strip(),
                    partijas_numurs=(self.in_prod_batch.text() or '').strip(),
                    iepirkuma_datums=self.in_prod_purchase_date.date().toString('yyyy-MM-dd'),
                    deriguma_termiņš=self.in_prod_expiry_date.date().toString('yyyy-MM-dd'),
                    minimalais_atlikums=float(self.in_prod_min_stock.value()),
                    statuss=(self.cb_prod_status.currentText() or 'Aktīva').strip(),
                    pavadzimes_numurs=(self.in_prod_delivery_doc.text() or '').strip(),
                    atlikums=float(self.in_prod_stock.value()),
                    supplier_code=(self.in_supplier_code.text() or '').strip(),
                    supplier_email=(self.in_supplier_email.text() or '').strip(),
                    supplier_phone=(self.in_supplier_phone.text() or '').strip(),
                    supplier_api_url=(self.in_supplier_api_url.text() or '').strip(),
                    supplier_product_url=(self.in_supplier_product_url.text() or '').strip(),
                    supplier_lead_time_days=int(self.in_supplier_lead_days.value()),
                    preferred_supplier=bool(self.ck_preferred_supplier.isChecked()),
                    hs_kods=(self.in_prod_hs_code.text() or '').strip(),
                    izcelsmes_valsts=(self.in_prod_origin_country.text() or '').strip(),
                    neto_svars=(self.in_prod_net_weight.text() or '').strip(),
                    bruto_svars=(self.in_prod_gross_weight.text() or '').strip(),
                    dokumentu_mape=(self.in_prod_docs_folder.text() or '').strip(),
                    inventory_id=(getattr(self, '_current_inventory_edit_key', '') or '').strip(),
                    dokumenti=self._inventory_collect_doc_list(),
                    last_sync_at=datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
                )
                self._noliktava.upsert(item)
                _refresh_table()
                QMessageBox.information(self, 'Noliktava', 'Kartīte saglabāta.')
            except Exception as e:
                QMessageBox.critical(self, 'Kļūda', f'Neizdevās saglabāt kartīti: {e}')

        def _delete_selected_item():
            try:
                rows = sorted({idx.row() for idx in self.tbl_noliktava.selectionModel().selectedRows()})
                if not rows:
                    r = self.tbl_noliktava.currentRow()
                    if r >= 0:
                        rows = [r]
                if not rows:
                    return
                keys = []
                labels = []
                for r in rows:
                    row_item = self.tbl_noliktava.item(r, 0)
                    if row_item is None:
                        continue
                    keys.append(row_item.data(Qt.UserRole) or '')
                    labels.append(_safe_text(row_item) or f'Rinda {r+1}')
                if not keys:
                    return
                question = f'Dzēst atlasītās kartītes ({len(keys)})?' if len(keys) > 1 else f'Dzēst kartīti ar SKU {labels[0]}?'
                if QMessageBox.question(self, 'Dzēst', question) != QMessageBox.StandardButton.Yes:
                    return
                self._noliktava.delete_by_inventory_ids(keys)
                _refresh_table(); _clear_editor()
            except Exception as e:
                QMessageBox.critical(self, 'Kļūda', str(e))

        def _duplicate_selected_items():
            try:
                rows = sorted({idx.row() for idx in self.tbl_noliktava.selectionModel().selectedRows()})
                if not rows:
                    r = self.tbl_noliktava.currentRow()
                    if r >= 0:
                        rows = [r]
                if not rows:
                    QMessageBox.information(self, 'Noliktava', 'Atlasiet vismaz vienu preces kartīti.')
                    return
                keys = []
                for r in rows:
                    row_item = self.tbl_noliktava.item(r, 0)
                    if row_item is not None:
                        keys.append(row_item.data(Qt.UserRole) or '')
                created = self._noliktava.duplicate_by_inventory_ids(keys)
                _refresh_table()
                QMessageBox.information(self, 'Noliktava', f'Dublētas kartītes: {created}.')
            except Exception as e:
                QMessageBox.critical(self, 'Kļūda', f'Neizdevās dublēt kartītes: {e}')

        def _export_selected_to_positions():
            try:
                r = self.tbl_noliktava.currentRow()
                if r < 0:
                    QMessageBox.information(self, 'Noliktava', 'Atlasiet preces kartīti.')
                    return
                row_item = self.tbl_noliktava.item(r, 0)
                row_key = row_item.data(Qt.UserRole) if row_item is not None else ''
                sku = _safe_text(row_item)
                item = self._noliktava.get_by_inventory_id(row_key) if row_key else None
                if item is None:
                    item = self._noliktava.get_by_sku(sku)
                if item is None:
                    return
                field_options = [
                    ('nosaukums', 'Nosaukums / apraksts'),
                    ('sku', 'SKU'),
                    ('svitrkods', 'Svītrkods'),
                    ('kategorija', 'Kategorija'),
                    ('apakskategorija', 'Apakškategorija'),
                    ('noliktavas_nosaukums', 'Noliktava'),
                    ('atrasanas_vieta', 'Atrašanās vieta'),
                    ('piegadatajs', 'Piegādātājs'),
                    ('razotajs', 'Ražotājs'),
                    ('partijas_numurs', 'Partijas Nr.'),
                    ('serialais_numurs', 'Seriālais Nr.'),
                    ('pavadzimes_numurs', 'Pavadzīmes Nr.'),
                    ('iepirkuma_datums', 'Iepirkuma datums'),
                    ('deriguma_termiņš', 'Derīguma termiņš'),
                    ('statuss', 'Statuss'),
                    ('cena', 'Cena'),
                    ('pvn_likme', 'PVN %'),
                    ('vieniba', 'Vienība'),
                    ('foto_path', 'Foto'),
                    ('piezimes', 'Piezīmes'),
                ]
                opts = self._choose_inventory_transfer_options(item, max(1.0, float(getattr(item, 'atlikums', 1.0) or 1.0)), field_options)
                if not opts:
                    return
                self._pievienot_poziciju_no_noliktavas(item, qty=float(opts.get('qty', 1.0) or 1.0), selected_fields=opts.get('selected_fields') or ['nosaukums'], column_mapping=opts.get('column_mapping') or {}, attach_docs=opts.get('documents') or [])
                self._update_preview()
                QMessageBox.information(self, 'Pozīcijas', 'Prece pievienota pozīcijām.')
            except Exception as e:
                QMessageBox.critical(self, 'Kļūda', f'Neizdevās pievienot pozīcijām: {e}')

        def _issue_selected_item():
            try:
                r = self.tbl_noliktava.currentRow()
                if r < 0:
                    QMessageBox.information(self, 'Noliktava', 'Atlasiet preces kartīti.')
                    return
                row_item = self.tbl_noliktava.item(r, 0)
                sku = (row_item.data(Qt.UserRole) if row_item is not None else '') or _safe_text(row_item)
                qty, ok = QInputDialog.getDouble(self, 'Izrakstīt', 'Daudzums:', 1.0, 0.001, 1e9, 3)
                if not ok:
                    return
                ref, _ = QInputDialog.getText(self, 'Atsauce', 'Dokumenta / pavadzīmes Nr. (neobligāti):')
                ok2, msg = self._noliktava.issue_item(sku, qty, auto_remove_depleted=self.chk_auto_remove_depleted.isChecked(), reference=ref)
                if ok2:
                    _refresh_table(); _clear_editor(); QMessageBox.information(self, 'Noliktava', msg)
                else:
                    QMessageBox.warning(self, 'Noliktava', msg)
            except Exception as e:
                QMessageBox.critical(self, 'Kļūda', str(e))

        def _receive_selected_item():
            try:
                r = self.tbl_noliktava.currentRow()
                if r < 0:
                    QMessageBox.information(self, 'Noliktava', 'Atlasiet preces kartīti.')
                    return
                row_item = self.tbl_noliktava.item(r, 0)
                sku = (row_item.data(Qt.UserRole) if row_item is not None else '') or _safe_text(row_item)
                qty, ok = QInputDialog.getDouble(self, 'Saņemt noliktavā', 'Daudzums:', 1.0, 0.001, 1e9, 3)
                if not ok:
                    return
                ref, _ = QInputDialog.getText(self, 'Atsauce', 'Pavadzīmes Nr. / atsauce (neobligāti):')
                ok2, msg = self._noliktava.receive_item(sku, qty, reference=ref)
                if ok2:
                    _refresh_table(); QMessageBox.information(self, 'Noliktava', msg)
                else:
                    QMessageBox.warning(self, 'Noliktava', msg)
            except Exception as e:
                QMessageBox.critical(self, 'Kļūda', str(e))

        def _adjust_selected_item():
            try:
                r = self.tbl_noliktava.currentRow()
                if r < 0:
                    QMessageBox.information(self, 'Noliktava', 'Atlasiet preces kartīti.')
                    return
                sku = _safe_text(self.tbl_noliktava.item(r, 0))
                current = 0.0
                try:
                    current = float(_safe_text(self.tbl_noliktava.item(r, 10)) or 0.0)
                except Exception:
                    current = 0.0
                qty, ok = QInputDialog.getDouble(self, 'Inventarizācijas korekcija', 'Jaunais atlikums:', current, -1e9, 1e9, 3)
                if not ok:
                    return
                ref, _ = QInputDialog.getText(self, 'Atsauce', 'Piezīme / inventarizācijas akts (neobligāti):')
                ok2, msg = self._noliktava.adjust_stock(sku, qty, reference=ref)
                if ok2:
                    _refresh_table(); QMessageBox.information(self, 'Noliktava', msg)
                else:
                    QMessageBox.warning(self, 'Noliktava', msg)
            except Exception as e:
                QMessageBox.critical(self, 'Kļūda', str(e))

        def _export_csv_inventory():
            try:
                path, _ = QFileDialog.getSaveFileName(self, 'Eksportēt noliktavu CSV', '', 'CSV (*.csv)')
                if not path:
                    return
                import csv
                with open(path, 'w', newline='', encoding='utf-8-sig') as f:
                    wr = csv.writer(f, delimiter=';')
                    wr.writerow(['SKU','Svītrkods','Nosaukums','Kategorija','Apakškategorija','Vienība','Cena','PVN %','Atlikums','Minimālais atlikums','Iepirkuma valūta','Noliktava','Atrašanās vieta','Piegādātājs','Ražotājs','Pavadzīmes Nr.','Partijas Nr.','Sērijas Nr.','Iepirkuma datums','Derīguma termiņš','Statuss','Piezīmes','Foto'])
                    for it in self._noliktava.items:
                        wr.writerow([
                            getattr(it,'sku',''), getattr(it,'svitrkods',''), getattr(it,'nosaukums',''), getattr(it,'kategorija',''), getattr(it,'apakskategorija',''),
                            getattr(it,'vieniba',''), getattr(it,'cena',''), getattr(it,'pvn_likme',''), getattr(it,'atlikums',''), getattr(it,'minimalais_atlikums',''),
                            getattr(it,'iepirkuma_valuta',''), getattr(it,'noliktavas_nosaukums',''), getattr(it,'atrasanas_vieta',''), getattr(it,'piegadatajs',''),
                            getattr(it,'razotajs',''), getattr(it,'pavadzimes_numurs',''), getattr(it,'partijas_numurs',''), getattr(it,'serialais_numurs',''),
                            getattr(it,'iepirkuma_datums',''), getattr(it,'deriguma_termiņš',''), getattr(it,'statuss',''), getattr(it,'piezimes',''), getattr(it,'foto_path','')
                        ])
                QMessageBox.information(self, 'Eksports', 'CSV eksports pabeigts.')
            except Exception as e:
                QMessageBox.critical(self, 'Kļūda', str(e))

        def _export_xlsx_inventory():
            try:
                path, _ = QFileDialog.getSaveFileName(self, 'Eksportēt noliktavu Excel', '', 'Excel (*.xlsx)')
                if not path:
                    return
                if not path.lower().endswith('.xlsx'):
                    path += '.xlsx'
                from openpyxl import Workbook
                wb = Workbook()
                ws = wb.active
                ws.title = 'Noliktava'
                headers = ['SKU','Svītrkods','Nosaukums','Kategorija','Apakškategorija','Vienība','Cena','PVN %','Atlikums','Minimālais atlikums','Iepirkuma valūta','Noliktava','Atrašanās vieta','Piegādātājs','Ražotājs','Pavadzīmes Nr.','Partijas Nr.','Sērijas Nr.','Iepirkuma datums','Derīguma termiņš','Statuss','Piezīmes','Foto']
                ws.append(headers)
                for it in self._noliktava.items:
                    ws.append([
                        getattr(it,'sku',''), getattr(it,'svitrkods',''), getattr(it,'nosaukums',''), getattr(it,'kategorija',''), getattr(it,'apakskategorija',''),
                        getattr(it,'vieniba',''), getattr(it,'cena',''), getattr(it,'pvn_likme',''), getattr(it,'atlikums',''), getattr(it,'minimalais_atlikums',''),
                        getattr(it,'iepirkuma_valuta',''), getattr(it,'noliktavas_nosaukums',''), getattr(it,'atrasanas_vieta',''), getattr(it,'piegadatajs',''),
                        getattr(it,'razotajs',''), getattr(it,'pavadzimes_numurs',''), getattr(it,'partijas_numurs',''), getattr(it,'serialais_numurs',''),
                        getattr(it,'iepirkuma_datums',''), getattr(it,'deriguma_termiņš',''), getattr(it,'statuss',''), getattr(it,'piezimes',''), getattr(it,'foto_path','')
                    ])
                wb.save(path)
                QMessageBox.information(self, 'Eksports', 'Excel eksports pabeigts.')
            except Exception as e:
                QMessageBox.critical(self, 'Kļūda', str(e))

        def _export_pdf_inventory():
            try:
                path, _ = QFileDialog.getSaveFileName(self, 'Eksportēt noliktavu PDF', '', 'PDF (*.pdf)')
                if not path:
                    return
                if not path.lower().endswith('.pdf'):
                    path += '.pdf'
                c = canvas.Canvas(path, pagesize=A4)
                width, height = A4
                y = height - 30
                c.setFont('Helvetica-Bold', 14)
                c.drawString(30, y, 'Noliktavas atlikumu saraksts')
                y -= 24
                c.setFont('Helvetica', 8)
                headers = ['SKU','Nosaukums','Noliktava','Vienība','Cena','Atlikums','Piegādātājs','Statuss']
                colx = [30, 90, 240, 330, 370, 420, 470, 540]
                for x, h in zip(colx, headers):
                    c.drawString(x, y, h)
                y -= 14
                for it in self._noliktava.items:
                    if y < 40:
                        c.showPage(); y = height - 30; c.setFont('Helvetica', 8)
                    vals = [getattr(it,'sku',''), getattr(it,'nosaukums',''), getattr(it,'noliktavas_nosaukums',''), getattr(it,'vieniba',''), getattr(it,'cena',''), str(getattr(it,'atlikums','')), getattr(it,'piegadatajs',''), getattr(it,'statuss','')]
                    for x, v in zip(colx, vals):
                        c.drawString(x, y, str(v)[:28])
                    y -= 12
                c.save()
                QMessageBox.information(self, 'Eksports', 'PDF eksports pabeigts.')
            except Exception as e:
                QMessageBox.critical(self, 'Kļūda', str(e))

        def _import_inventory_file():
            try:
                path, _ = QFileDialog.getOpenFileName(self, 'Importēt noliktavas datus', '', 'CSV/Excel (*.csv *.xlsx *.xls)')
                if not path:
                    return
                rows = []
                if path.lower().endswith('.csv'):
                    import csv
                    with open(path, 'r', encoding='utf-8-sig', newline='') as f:
                        sample = f.read(2048)
                        f.seek(0)
                        delimiter = ';' if sample.count(';') >= sample.count(',') else ','
                        rd = csv.DictReader(f, delimiter=delimiter)
                        rows = list(rd)
                else:
                    from openpyxl import load_workbook
                    wb = load_workbook(path, read_only=True, data_only=True)
                    ws = wb.active
                    header = [str(c.value or '').strip() for c in next(ws.iter_rows(min_row=1, max_row=1))]
                    for row in ws.iter_rows(min_row=2, values_only=True):
                        rows.append({header[i]: row[i] for i in range(min(len(header), len(row)))})

                imported = 0
                forced_imported = 0
                skipped = []

                def g(row, *names):
                    for n in names:
                        if n in row and row.get(n) not in (None, ''):
                            return str(row.get(n)).strip()
                    lower = {str(k).strip().lower(): v for k, v in row.items()}
                    for n in names:
                        v = lower.get(str(n).strip().lower())
                        if v not in (None, ''):
                            return str(v).strip()
                    return ''

                existing_signatures = set()
                for existing_item in list(getattr(self._noliktava, 'items', []) or []):
                    existing_signatures.add(_inventory_import_signature(existing_item))
                seen_signatures = set(existing_signatures)

                for row_idx, row in enumerate(rows, start=2):
                    sku = g(row, 'SKU', 'sku')
                    if not sku:
                        continue
                    item = InventoryItem(
                        sku=sku,
                        nosaukums=g(row, 'Nosaukums', 'nosaukums', 'Name'),
                        vieniba=g(row, 'Vienība', 'vieniba', 'Unit') or 'gab.',
                        cena=g(row, 'Cena', 'cena', 'Price'),
                        pvn_likme=g(row, 'PVN %', 'pvn', 'PVN', 'VAT'),
                        serialais_numurs=g(row, 'Sērijas Nr.', 'Seriālais Nr.', 'serialais_numurs', 'serial'),
                        piezimes=g(row, 'Piezīmes', 'piezimes', 'Notes'),
                        foto_path=g(row, 'Foto', 'foto', 'foto_path'),
                        svitrkods=g(row, 'Svītrkods', 'svitrkods', 'Barcode'),
                        kategorija=g(row, 'Kategorija', 'kategorija', 'Category'),
                        apakskategorija=g(row, 'Apakškategorija', 'apakskategorija', 'Subcategory'),
                        iepirkuma_valuta=g(row, 'Iepirkuma valūta', 'iepirkuma_valuta', 'Valūta', 'Currency') or 'EUR',
                        noliktavas_nosaukums=g(row, 'Noliktava', 'Noliktavas nosaukums', 'noliktava', 'warehouse'),
                        atrasanas_vieta=g(row, 'Atrašanās vieta', 'atrasanas_vieta', 'location'),
                        piegadatajs=g(row, 'Piegādātājs', 'piegadatajs', 'Supplier'),
                        razotajs=g(row, 'Ražotājs', 'razotajs', 'Manufacturer'),
                        partijas_numurs=g(row, 'Partijas Nr.', 'partijas_numurs', 'Batch'),
                        iepirkuma_datums=g(row, 'Iepirkuma datums', 'iepirkuma_datums', 'Purchase date'),
                        deriguma_termiņš=g(row, 'Derīguma termiņš', 'deriguma_termiņš', 'Expiry date'),
                        minimalais_atlikums=self._parse_num_lv_en(g(row, 'Minimālais atlikums', 'Min. atlik.', 'minimalais_atlikums', 'Minimum stock')),
                        statuss=g(row, 'Statuss', 'statuss', 'Status') or 'Aktīva',
                        pavadzimes_numurs=g(row, 'Pavadzīmes Nr.', 'pavadzimes_numurs', 'Reference'),
                        atlikums=self._parse_num_lv_en(g(row, 'Atlikums', 'atlikums', 'Stock')),
                        inventory_id=self._noliktava._new_inventory_id(),
                    )
                    sig = _inventory_import_signature(item)
                    if sig in seen_signatures:
                        skipped.append({
                            'row_no': row_idx,
                            'item': item,
                            'reason': _inventory_duplicate_reason(item),
                        })
                        continue
                    self._noliktava.upsert(item)
                    seen_signatures.add(sig)
                    imported += 1

                if skipped:
                    dlg = InventoryImportDuplicatesDialog(self, skipped)
                    if dlg.exec() == QDialog.Accepted:
                        for payload in dlg.selected_rows():
                            self._noliktava.upsert(payload.get('item'))
                            forced_imported += 1

                _refresh_table()
                msg = f'Importētas {imported} kartītes.'
                if skipped:
                    msg += f' Identiskas rindas sākotnēji netika iekļautas: {len(skipped)}.'
                    if forced_imported:
                        msg += f' Pēc manuālas apstiprināšanas papildus apstrādātas: {forced_imported}.'
                QMessageBox.information(self, 'Imports', msg)
            except Exception as e:
                QMessageBox.critical(self, 'Kļūda', f'Neizdevās importēt datus: {e}')

        def _export_movements_csv():
            try:
                path, _ = QFileDialog.getSaveFileName(self, 'Eksportēt kustību žurnālu', '', 'CSV (*.csv)')
                if not path:
                    return
                import csv
                with open(path, 'w', newline='', encoding='utf-8-sig') as f:
                    fieldnames = ['timestamp','type','sku','name','qty','vieniba','warehouse','reference','note']
                    wr = csv.DictWriter(f, fieldnames=fieldnames, delimiter=';', extrasaction='ignore')
                    wr.writeheader()
                    for row in getattr(self._noliktava, 'movements', []) or []:
                        src = dict(row or {})
                        normalized = {
                            'timestamp': src.get('timestamp') or src.get('datetime') or '',
                            'type': src.get('type') or src.get('movement_type') or '',
                            'sku': src.get('sku') or '',
                            'name': src.get('name') or src.get('nosaukums') or '',
                            'qty': src.get('qty') or '',
                            'vieniba': src.get('vieniba') or '',
                            'warehouse': src.get('warehouse') or src.get('noliktava') or '',
                            'reference': src.get('reference') or '',
                            'note': src.get('note') or src.get('piezimes') or '',
                        }
                        wr.writerow(normalized)
                QMessageBox.information(self, 'Eksports', 'Kustību žurnāls eksportēts.')
            except Exception as e:
                QMessageBox.critical(self, 'Kļūda', str(e))

        def _save_warehouse_name():
            name = (self.cb_saved_warehouses.currentText() or '').strip()
            if not name:
                return
            try:
                self._noliktava.add_warehouse(name)
            except Exception:
                pass
            _refresh_warehouse_lists()
            self.cb_saved_warehouses.setCurrentText(name)

        def _delete_warehouse_name():
            name = (self.cb_saved_warehouses.currentText() or '').strip()
            if not name:
                return
            if QMessageBox.question(self, 'Noliktava', f'Dzēst noliktavu "{name}" no saraksta?') != QMessageBox.StandardButton.Yes:
                return
            try:
                self._noliktava.remove_warehouse(name)
            except Exception:
                pass
            _refresh_warehouse_lists()

        def _apply_selected_warehouse_name():
            name = (self.cb_saved_warehouses.currentText() or '').strip()
            if name:
                self.in_prod_warehouse.setCurrentText(name)

        self.in_noliktava_search.textChanged.connect(_refresh_table)
        self.cb_noliktava_filter.currentIndexChanged.connect(_refresh_table)
        self.cb_noliktava_warehouse.currentIndexChanged.connect(_refresh_table)
        btn_refresh.clicked.connect(_refresh_table)
        btn_add.clicked.connect(_save_item)
        btn_edit.clicked.connect(lambda: _edit_selected_item(True))
        btn_clear.clicked.connect(_clear_editor)
        btn_delete.clicked.connect(_delete_selected_item)
        btn_duplicate.clicked.connect(_duplicate_selected_items)
        btn_to_poz.clicked.connect(_export_selected_to_positions)
        btn_issue.clicked.connect(_issue_selected_item)
        btn_receive.clicked.connect(_receive_selected_item)
        btn_adjust.clicked.connect(_adjust_selected_item)
        btn_import.clicked.connect(_import_inventory_file)
        btn_export.clicked.connect(_export_csv_inventory)
        btn_export_xlsx.clicked.connect(_export_xlsx_inventory)
        btn_export_pdf.clicked.connect(_export_pdf_inventory)
        btn_export_moves.clicked.connect(_export_movements_csv)
        btn_add_wh.clicked.connect(_save_warehouse_name)
        btn_del_wh.clicked.connect(_delete_warehouse_name)
        btn_apply_wh.clicked.connect(_apply_selected_warehouse_name)

        self._nol_refresh = _refresh_table
        self._nol_clear_editor = _clear_editor
        self.tab_noliktava_widget = w
        scroll.setWidget(content)
        outer.addWidget(scroll)
        self.tabs.addTab(w, 'Noliktava')
        _refresh_table()

    def _choose_inventory_photo(self):
        try:
            path, _ = QFileDialog.getOpenFileName(self, 'Izvēlēties noliktavas preces foto', '', 'Attēli (*.png *.jpg *.jpeg *.webp *.bmp)')
            if path:
                self.in_prod_photo.setText(path)
        except Exception as e:
            QMessageBox.warning(self, 'Noliktava', f'Neizdevās izvēlēties foto: {e}')

    def _export_noliktava_xlsx(self):
        path, _ = QFileDialog.getSaveFileName(self, 'Eksportēt noliktavu Excel', '', 'Excel (*.xlsx)')
        if not path:
            return
        try:
            from openpyxl import Workbook
            wb = Workbook()
            ws = wb.active
            ws.title = 'Noliktava'
            headers = ['SKU', 'Svītrkods', 'Nosaukums', 'Kategorija', 'Apakškategorija', 'Noliktava', 'Atrašanās vieta', 'Vienība', 'Iepirkuma valūta', 'Cena', 'PVN %', 'Atlikums', 'Minimālais atlikums', 'Pavadzīmes Nr.', 'Piegādātājs', 'Ražotājs', 'Partijas Nr.', 'Sērijas Nr.', 'Iepirkuma datums', 'Derīguma termiņš', 'Statuss', 'Piezīmes', 'Foto']
            ws.append(headers)
            for it in (self._noliktava.items or []):
                ws.append([
                    it.sku, getattr(it, 'svitrkods', '') or '', it.nosaukums, getattr(it, 'kategorija', '') or '', getattr(it, 'apakskategorija', '') or '',
                    getattr(it, 'noliktavas_nosaukums', '') or '', getattr(it, 'atrasanas_vieta', '') or '', it.vieniba, getattr(it, 'iepirkuma_valuta', 'EUR') or 'EUR',
                    it.cena, it.pvn_likme, it.atlikums, getattr(it, 'minimalais_atlikums', 0.0) or 0.0, getattr(it, 'pavadzimes_numurs', '') or '',
                    getattr(it, 'piegadatajs', '') or '', getattr(it, 'razotajs', '') or '', getattr(it, 'partijas_numurs', '') or '', getattr(it, 'serialais_numurs', '') or '',
                    getattr(it, 'iepirkuma_datums', '') or '', getattr(it, 'deriguma_termiņš', '') or '', getattr(it, 'statuss', '') or '', getattr(it, 'piezimes', '') or '', getattr(it, 'foto_path', '') or ''
                ])
            for col in ws.columns:
                max_len = 0
                col_letter = col[0].column_letter
                for cell in col:
                    max_len = max(max_len, len(str(cell.value or '')))
                ws.column_dimensions[col_letter].width = min(max_len + 2, 40)
            wb.save(path)
            QMessageBox.information(self, 'Noliktava', 'Excel eksports pabeigts.')
        except Exception as e:
            QMessageBox.warning(self, 'Noliktava', f'Neizdevās eksportēt Excel: {e}')

    def _export_noliktava_pdf(self):
        path, _ = QFileDialog.getSaveFileName(self, 'Eksportēt noliktavu PDF', '', 'PDF (*.pdf)')
        if not path:
            return
        try:
            doc = SimpleDocTemplate(path, pagesize=landscape(A4), leftMargin=10*mm, rightMargin=10*mm, topMargin=10*mm, bottomMargin=10*mm)
            styles = getSampleStyleSheet()
            story = [Paragraph('Noliktavas atlikums', styles['Title']), Spacer(1, 4*mm)]
            s = self._noliktava.summary(include_zero=True)
            story.append(Paragraph(f"Kartītes: {int(s['total_items'])} | Kopējais atlikums: {s['total_qty']:.3f} | Vērtība: {s['total_value']:.2f} EUR | Noliktavas: {int(s.get('warehouse_count', 0))}", styles['Normal']))
            story.append(Spacer(1, 4*mm))
            data = [["SKU", "Nosaukums", "Noliktava", "Vieta", "Val.", "Cena", "Atlikums", "Min.", "Pavadzīme Nr.", "Piegādātājs", "Statuss"]]
            for it in (self._noliktava.items or []):
                data.append([
                    it.sku, it.nosaukums, getattr(it, 'noliktavas_nosaukums', '') or '', getattr(it, 'atrasanas_vieta', '') or '', getattr(it, 'iepirkuma_valuta', 'EUR') or 'EUR',
                    it.cena, str(it.atlikums), str(getattr(it, 'minimalais_atlikums', 0.0) or 0.0), getattr(it, 'pavadzimes_numurs', '') or '', getattr(it, 'piegadatajs', '') or '', getattr(it, 'statuss', '') or ''
                ])
            tbl = Table(data, repeatRows=1)
            tbl.setStyle(TableStyle([
                ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#d9e2f3')),
                ('GRID', (0,0), (-1,-1), 0.4, colors.grey),
                ('FONTNAME', (0,0), (-1,0), 'Helvetica-Bold'),
                ('VALIGN', (0,0), (-1,-1), 'TOP'),
                ('FONTSIZE', (0,0), (-1,-1), 7),
                ('ROWBACKGROUNDS', (0,1), (-1,-1), [colors.whitesmoke, colors.HexColor('#f3f6fa')]),
            ]))
            story.append(tbl)
            doc.build(story)
            QMessageBox.information(self, 'Noliktava', 'PDF eksports pabeigts.')
        except Exception as e:
            QMessageBox.warning(self, 'Noliktava', f'Neizdevās eksportēt PDF: {e}')

    def _pievienot_poziciju_no_noliktavas(self, it: NoliktavasPrece, qty: float = 1.0, selected_fields=None):
        """Pievieno pozīciju tabulai no noliktavas preces ar lietotāja izvēlētu lauku pārnesi."""
        try:
            if not hasattr(self, "tab") or self.tab is None:
                QMessageBox.warning(self, "Noliktava", "Nav atrasta pozīciju tabula.")
                return

            row = self.tab.rowCount()
            self.tab.insertRow(row)

            headers = []
            for c in range(self.tab.columnCount()):
                hi = self.tab.horizontalHeaderItem(c)
                headers.append(hi.text().strip().lower() if hi else "")

            def col_idx_contains(keywords):
                for c, h in enumerate(headers):
                    for k in keywords:
                        if k in h:
                            return c
                return None

            c_desc = col_idx_contains(["apraksts"])
            c_qty = col_idx_contains(["daudzums"])
            c_unit = col_idx_contains(["vienīb", "vieniba"])
            c_price = col_idx_contains(["cena"])
            c_sum = col_idx_contains(["summa"])
            c_photo = col_idx_contains(["foto"])
            c_serial = col_idx_contains(["seriā", "serial"])
            c_notes = col_idx_contains(["piezīm", "piezim"])
            c_warranty = col_idx_contains(["garant"])

            desc_lines = []
            field_title_map = {
                'sku': 'SKU', 'svitrkods': 'Svītrkods', 'nosaukums': 'Nosaukums', 'kategorija': 'Kategorija', 'apakskategorija': 'Apakškategorija',
                'pavadzimes_numurs': 'Pavadzīmes Nr.', 'piegadatajs': 'Piegādātājs', 'razotajs': 'Ražotājs', 'partijas_numurs': 'Partijas Nr.',
                'serialais_numurs': 'Sērijas Nr.', 'noliktavas_nosaukums': 'Noliktava', 'atrasanas_vieta': 'Atrašanās vieta', 'iepirkuma_datums': 'Iepirkuma datums',
                'deriguma_termiņš': 'Derīguma termiņš', 'iepirkuma_valuta': 'Valūta', 'piezimes': 'Piezīmes', 'foto_path': 'Foto'
            }
            selected_fields = selected_fields or [('nosaukums', 'Nosaukums', None)]
            selected_keys = [x[0] if isinstance(x, tuple) else x for x in selected_fields]

            for key in selected_keys:
                val = getattr(it, key, '') if hasattr(it, key) else ''
                if val in (None, ''):
                    continue
                if key == 'nosaukums':
                    desc_lines.insert(0, str(val))
                else:
                    desc_lines.append(f"{field_title_map.get(key, key)}: {val}")

            if not desc_lines:
                desc_lines = [it.nosaukums or it.sku or '']
            desc = '\\n'.join([x for x in desc_lines if x])

            if c_desc is not None:
                self.tab.setItem(row, c_desc, QTableWidgetItem(desc))
            if c_qty is not None:
                self.tab.setItem(row, c_qty, QTableWidgetItem(str(qty)))
            if c_unit is not None:
                self.tab.setItem(row, c_unit, QTableWidgetItem(it.vieniba or "gab."))
            if c_price is not None:
                self.tab.setItem(row, c_price, QTableWidgetItem(it.cena or ""))
            if c_photo is not None and ('foto_path' in selected_keys or getattr(it, 'foto_path', '')):
                self.tab.setItem(row, c_photo, QTableWidgetItem(getattr(it, "foto_path", "") or ""))
                try:
                    self._ensure_photo_cell(row)
                except Exception:
                    pass
            if c_serial is not None and getattr(it, 'serialais_numurs', ''):
                self.tab.setItem(row, c_serial, QTableWidgetItem(getattr(it, 'serialais_numurs', '') or ''))
            if c_notes is not None:
                notes = []
                for key in selected_keys:
                    if key in ('nosaukums', 'serialais_numurs', 'foto_path'):
                        continue
                    val = getattr(it, key, '') if hasattr(it, key) else ''
                    if val not in (None, ''):
                        notes.append(f"{field_title_map.get(key, key)}: {val}")
                if notes:
                    self.tab.setItem(row, c_notes, QTableWidgetItem(' | '.join(notes)))
            if c_warranty is not None and getattr(it, 'deriguma_termiņš', ''):
                self.tab.setItem(row, c_warranty, QTableWidgetItem(getattr(it, 'deriguma_termiņš', '') or ''))

            try:
                price = float(str(it.cena or "0").replace(",", "."))
            except Exception:
                price = 0.0
            summa = qty * price
            if c_sum is not None:
                self.tab.setItem(row, c_sum, QTableWidgetItem(f"{summa:.2f}"))

            for meth in ["_atjaunot_kopsummas", "_refresh_totals", "atjaunot_kopsummas"]:
                if hasattr(self, meth):
                    try:
                        getattr(self, meth)()
                        break
                    except Exception:
                        pass
        except Exception as e:
            QMessageBox.warning(self, "Noliktava", f"Neizdevās pievienot pozīciju: {e}")


    def _būvēt_attēli_tab(self):
        """Fotogrāfiju pievienošana tabulas veidā (ar nosaukumu/ aprakstu rediģēšanu)."""
        w = QWidget()
        v = QVBoxLayout()

        # Tabula: [Priekšskats | Nosaukums/Apraksts | Fails | ... (iekšējie dati)]
        self.photos_table = QTableWidget(0, 4)
        self.photos_table.setObjectName("photos_table")
        self.photos_table.setHorizontalHeaderLabels(["Foto", "Nosaukums / Apraksts", "Fails", ""])
        self.photos_table.verticalHeader().setVisible(False)
        self.photos_table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.photos_table.setSelectionMode(QAbstractItemView.SingleSelection)
        self.photos_table.setEditTriggers(QAbstractItemView.DoubleClicked | QAbstractItemView.SelectedClicked | QAbstractItemView.EditKeyPressed)
        self.photos_table.setAlternatingRowColors(True)
        self.photos_table.setShowGrid(True)
        self.photos_table.setColumnHidden(3, True)  # iekšējais (UserRole) datu items

        hh = self.photos_table.horizontalHeader()
        hh.setSectionResizeMode(0, QHeaderView.ResizeToContents)
        hh.setSectionResizeMode(1, QHeaderView.Stretch)
        hh.setSectionResizeMode(2, QHeaderView.Stretch)

        self.photos_table.setContextMenuPolicy(Qt.CustomContextMenu)
        self.photos_table.customContextMenuRequested.connect(self._photos_context_menu)

        v.addWidget(self.photos_table)

        btns = QHBoxLayout()

        piev = QPushButton("Pievienot foto…")
        piev.clicked.connect(self.pievienot_attēlu)

        # Mazas pārvietošanas pogas (strādā uz atlasīto rindu)
        augsa = QToolButton(); augsa.setText("▲"); augsa.setToolTip("Pārvietot uz augšu")
        augsa.clicked.connect(lambda: self.pārvietot_att(-1))
        leja = QToolButton(); leja.setText("▼"); leja.setToolTip("Pārvietot uz leju")
        leja.clicked.connect(lambda: self.pārvietot_att(1))

        dzest = QToolButton(); dzest.setText("🗑"); dzest.setToolTip("Dzēst atlasīto")
        dzest.clicked.connect(self.dzest_att)

        btns.addWidget(piev)
        btns.addSpacing(8)
        btns.addWidget(augsa)
        btns.addWidget(leja)
        btns.addWidget(dzest)
        btns.addStretch()

        v.addLayout(btns)
        w.setLayout(v)
        self.tabs.addTab(w, "Fotogrāfijas")

        # Kad maina nosaukumu/ aprakstu → atjauno preview
        self.photos_table.itemChanged.connect(self._photos_item_changed)
        self.photos_table.cellDoubleClicked.connect(self._photos_cell_double_clicked)

        # Back-compat: dažās vietās kods izmanto self.img_list (vēsturiskais QListView/QListWidget).
        # Lai neko “nenolauztu”, mēs norādām to uz jauno tabulu, kurai arī ir model() ar tiem pašiem signāliem.
        self.img_list = self.photos_table
        try:
            m = self.img_list.model()
            m.rowsInserted.connect(self._update_preview)
            m.rowsRemoved.connect(self._update_preview)
            m.rowsMoved.connect(self._update_preview)
            # drošībai – ja tiek pārrakstīti dati bez rindu izmaiņām
            m.dataChanged.connect(self._update_preview)
            m.layoutChanged.connect(self._update_preview)
        except Exception:
            # Ja kādā vidē model() nav pieejams/atšķiras, nelaužam startu – preview tāpat atjaunosies no citiem trigeriem.
            pass

    
    def _photos_make_thumb_widget(self, path: str) -> QWidget:
        lbl = QLabel()
        lbl.setAlignment(Qt.AlignCenter)
        lbl.setMinimumWidth(90)
        lbl.setMinimumHeight(70)
        lbl.setToolTip(path)
        try:
            pm = QPixmap(path)
            if not pm.isNull():
                pm = pm.scaled(120, 90, Qt.KeepAspectRatio, Qt.SmoothTransformation)
                lbl.setPixmap(pm)
            else:
                lbl.setText("—")
        except Exception:
            lbl.setText("—")
        return lbl

    def _photos_row_payload(self, row: int) -> Optional[dict]:
        """Iekšējie dati par rindu (ceļš+paraksts)."""
        it = self.photos_table.item(row, 3)
        if it is None:
            return None
        return it.data(Qt.UserRole) or None

    def _photos_set_row_payload(self, row: int, payload: dict):
        it = self.photos_table.item(row, 3)
        if it is None:
            it = QTableWidgetItem()
            self.photos_table.setItem(row, 3, it)
        it.setData(Qt.UserRole, payload)

    def _photos_add_row(self, path: str, caption: str = ""):
        row = self.photos_table.rowCount()
        self.photos_table.insertRow(row)

        # Thumbnail
        self.photos_table.setCellWidget(row, 0, self._photos_make_thumb_widget(path))

        # Caption (editable)
        cap_item = QTableWidgetItem(caption or "")
        cap_item.setToolTip("Dubultklikšķis, lai rediģētu nosaukumu/ aprakstu")
        self.photos_table.setItem(row, 1, cap_item)

        # File (read-only)
        file_item = QTableWidgetItem(os.path.basename(path))
        file_item.setFlags(file_item.flags() & ~Qt.ItemIsEditable)
        file_item.setToolTip(path)
        self.photos_table.setItem(row, 2, file_item)

        # Hidden payload
        hidden = QTableWidgetItem()
        self.photos_table.setItem(row, 3, hidden)
        self._photos_set_row_payload(row, {"ceļš": path, "paraksts": caption or ""})

        self.photos_table.setCurrentCell(row, 1)

    def _photos_context_menu(self, pos: QPoint):
        menu = QMenu(self)
        row = self.photos_table.rowAt(pos.y())
        if row < 0:
            act_add = menu.addAction("Pievienot foto…")
            act_add.triggered.connect(self.pievienot_attēlu)
            menu.exec(self.photos_table.viewport().mapToGlobal(pos))
            return

        payload = self._photos_row_payload(row) or {}
        path = payload.get("ceļš", "")

        act_open = menu.addAction("Atvērt failu")
        act_open.triggered.connect(lambda: self._open_file_in_os(path) if path else None)

        act_reveal = menu.addAction("Atvērt mapē")
        act_reveal.triggered.connect(lambda: self._reveal_file_in_os(path) if path else None)

        menu.addSeparator()
        act_up = menu.addAction("Pārvietot uz augšu")
        act_up.triggered.connect(lambda: self._photos_move_selected(-1))
        act_down = menu.addAction("Pārvietot uz leju")
        act_down.triggered.connect(lambda: self._photos_move_selected(1))

        menu.addSeparator()
        act_del = menu.addAction("Dzēst")
        act_del.triggered.connect(self.dzest_att)

        menu.exec(self.photos_table.viewport().mapToGlobal(pos))

    def _photos_item_changed(self, item: QTableWidgetItem):
        # ja mainās caption kolonna → atjauno payload + preview
        if item is None:
            return
        if item.column() != 1:
            return
        row = item.row()
        payload = self._photos_row_payload(row) or {}
        payload["paraksts"] = item.text()
        self._photos_set_row_payload(row, payload)
        self._update_preview()

    def _photos_cell_double_clicked(self, row: int, col: int):
        # dubultklikšķis uz foto -> atvērt failu
        if col == 0:
            payload = self._photos_row_payload(row) or {}
            path = payload.get("ceļš", "")
            if path:
                self._open_file_in_os(path)

    def _photos_move_selected(self, direction: int):
        row = self.photos_table.currentRow()
        if row < 0:
            return
        new_row = row + direction
        if new_row < 0 or new_row >= self.photos_table.rowCount():
            return

        payload = self._photos_row_payload(row) or {}
        caption = (self.photos_table.item(row, 1).text() if self.photos_table.item(row, 1) else payload.get("paraksts", ""))
        path = payload.get("ceļš", "")

        # remove
        self.photos_table.blockSignals(True)
        self.photos_table.removeRow(row)
        self.photos_table.insertRow(new_row)

        # rebuild at new position
        self.photos_table.setCellWidget(new_row, 0, self._photos_make_thumb_widget(path))
        cap_item = QTableWidgetItem(caption or "")
        self.photos_table.setItem(new_row, 1, cap_item)
        file_item = QTableWidgetItem(os.path.basename(path))
        file_item.setFlags(file_item.flags() & ~Qt.ItemIsEditable)
        file_item.setToolTip(path)
        self.photos_table.setItem(new_row, 2, file_item)
        hidden = QTableWidgetItem()
        self.photos_table.setItem(new_row, 3, hidden)
        self._photos_set_row_payload(new_row, {"ceļš": path, "paraksts": caption or ""})
        self.photos_table.blockSignals(False)

        self.photos_table.setCurrentCell(new_row, 1)
        self._update_preview()

    def pievienot_attēlu(self):
        ceļi, _ = QFileDialog.getOpenFileNames(
            self,
            "Izvēlēties attēlus",
            "",
            "Attēli (*.png *.jpg *.jpeg *.webp)"
        )
        for c in ceļi:
            self._photos_add_row(c, "")

        if ceļi:
            self._update_preview()

    def pārvietot_att(self, virziens):
        self._photos_move_selected(virziens)

    def dzest_att(self):
        r = self.photos_table.currentRow()
        if r >= 0:
            self.photos_table.removeRow(r)
            self._update_preview()

    # Back-compat: vecais API (vairs neizmantojam, bet atstājam, lai nekas nelūzt)
    def rediģēt_att_parakstu(self, *args, **kwargs):
        # Tagad paraksts/nosaukums rediģējas tieši tabulā (2. kolonna).
        return
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
        self.cb_paraksta_rezims = QComboBox()
        self.cb_paraksta_rezims.addItem("Fiziska paraksta rindas", "physical")
        self.cb_paraksta_rezims.addItem("Elektroniski parakstīts dokuments", "electronic")
        self.cb_paraksta_rezims.addItem("Dokuments ir spēkā bez paraksta", "no_signature_required")
        self.in_paraksta_nav_teksts = QLineEdit("Dokuments ir spēkā bez paraksta, ja tas sagatavots un apstiprināts atbilstoši pušu noteiktajai kārtībai.")
        self.in_paraksta_vards_rekviziti = QLineEdit(); self.in_paraksta_vards_rekviziti.setPlaceholderText("Tukšs = izmantot kontaktpersonu no Rekvizītiem")
        self.in_paraksta_vards_pienemejs = QLineEdit(); self.in_paraksta_vards_pienemejs.setPlaceholderText("Tukšs = izmantot Pieņēmēja kontaktpersonu")
        self.in_paraksta_vards_nodevejs = QLineEdit(); self.in_paraksta_vards_nodevejs.setPlaceholderText("Tukšs = izmantot Nodevēja kontaktpersonu")
        self.in_papildu_parakstu_rindas = QTextEdit()
        self.in_papildu_parakstu_rindas.setPlaceholderText("Viena papildu paraksta rinda katrā rindā. Formāts: Paraksta nosaukums | Parakstītāja vārds\nPiemērs: Saskaņoja | Jānis Bērziņš")
        self.in_papildu_parakstu_rindas.setMaximumHeight(90)

        self.in_logo = QLineEdit();
        logo_btn = QToolButton(); logo_btn.setText("…"); logo_btn.clicked.connect(self.izvēlēties_logo)
        logo_box = QHBoxLayout(); logo_box.addWidget(self.in_logo); logo_box.addWidget(logo_btn)
        logo_w = QWidget(); logo_w.setLayout(logo_box)

        self.in_fonts = QLineEdit();
        font_btn = QToolButton(); font_btn.setText("…"); font_btn.clicked.connect(self.izvēlēties_fontu)
        font_box = QHBoxLayout(); font_box.addWidget(self.in_fonts); font_box.addWidget(font_btn)
        font_w = QWidget(); font_w.setLayout(font_box)

        # DOCX šablons (nav obligāts)
        self.in_docx_template = QLineEdit()
        docx_tpl_btn = QToolButton(); docx_tpl_btn.setText("…")
        docx_tpl_btn.clicked.connect(self.izvēlēties_docx_sablonu)
        docx_tpl_box = QHBoxLayout(); docx_tpl_box.addWidget(self.in_docx_template); docx_tpl_box.addWidget(docx_tpl_btn)
        docx_tpl_w = QWidget(); docx_tpl_w.setLayout(docx_tpl_box)

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

        self.btn_export_zip = QPushButton("Saglabāt ZIP…")
        self.btn_export_zip.clicked.connect(self.ģenerēt_zip_dialogs)
        self.btn_print_pdf = QPushButton("Drukāt PDF…") # JAUNA RINDAS
        self.btn_print_pdf.clicked.connect(self.drukāt_pdf_dialogs) # JAUNA RINDAS



        form.addRow(self.ck_pvn, self.in_pvn)
        form.addRow(self.ck_paraksti)
        form.addRow("Paraksta veids", self.cb_paraksta_rezims)
        form.addRow("Teksts režīmam bez paraksta", self.in_paraksta_nav_teksts)
        form.addRow("Parakstītājs no Rekvizītiem", self.in_paraksta_vards_rekviziti)
        form.addRow("Parakstītājs — Pieņēmējs", self.in_paraksta_vards_pienemejs)
        form.addRow("Parakstītājs — Nodevējs", self.in_paraksta_vards_nodevejs)
        form.addRow("Papildu parakstu rindas", self.in_papildu_parakstu_rindas)
        form.addRow("Logotips (neobligāti)", logo_w)
        form.addRow("Fonts TTF/OTF (ieteicams latviešu diakritikai)", font_w)
        form.addRow("DOCX šablons (opc.)", docx_tpl_w)
        form.addRow("Pieņēmēja paraksta attēls", w_paraksts_pie)
        form.addRow("Nodevēja paraksta attēls", w_paraksts_nod)
        form.addRow("Valoda", self.lang_combo)
        form.addRow(btn_saglabat_nokl)
        form.addRow(btn_saglabat_sablonu)  # JAUNA RINDAS
        form.addRow(self.btn_generate_pdf)
        form.addRow(self.btn_generate_docx)
        form.addRow(self.btn_export_zip)
        form.addRow(self.btn_print_pdf)  # JAUNA RINDAS

        self.ck_pvn.stateChanged.connect(self._update_preview)
        self.in_pvn.valueChanged.connect(self._update_preview)
        self.ck_paraksti.stateChanged.connect(self._update_preview)
        self.cb_paraksta_rezims.currentIndexChanged.connect(self._on_paraksta_rezims_changed)
        self.ck_elektroniskais_paraksts.stateChanged.connect(self._sync_signature_mode_from_checkbox)
        self.in_paraksta_nav_teksts.textChanged.connect(self._update_preview)
        self.in_paraksta_vards_rekviziti.textChanged.connect(self._update_preview)
        self.in_paraksta_vards_pienemejs.textChanged.connect(self._update_preview)
        self.in_paraksta_vards_nodevejs.textChanged.connect(self._update_preview)
        self.in_papildu_parakstu_rindas.textChanged.connect(self._update_preview)
        self.in_logo.textChanged.connect(self._update_preview)
        self.in_fonts.textChanged.connect(self._update_preview)
        self.in_paraksts_pie.textChanged.connect(self._update_preview)
        self.in_paraksts_nod.textChanged.connect(self._update_preview)

        self._on_paraksta_rezims_changed()

        w.setLayout(form)
        self.tabs.addTab(w, "Iestatījumi & Eksports")

    
    def _wrap_date_with_system_button(self, date_edit: QDateEdit) -> QWidget:
        """Ietin QDateEdit ar pogu, kas ielādē sistēmas (Windows) šodienas datumu."""
        w = QWidget()
        h = QHBoxLayout()
        h.setContentsMargins(0, 0, 0, 0)
        h.addWidget(date_edit, 1)
        btn = QToolButton()
        btn.setText("Ielādēt sistēmas datumu")
        btn.clicked.connect(lambda: date_edit.setDate(datetime.now().date()))
        h.addWidget(btn)
        w.setLayout(h)
        return w

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
            temp_pdf_path = ģenerēt_pdf(akta_dati, pdf_ceļš=None, encrypt_pdf=False)

            if not os.path.exists(temp_pdf_path):
                QMessageBox.critical(self, "Drukas kļūda", "Neizdevās ģenerēt PDF failu drukāšanai.")
                return

            printer = QPrinter(QPrinter.HighResolution)

            # Lietojam PROGRAMMAS lapas izmēru (nevis printera/sistēmas noklusējumu, kas bieži ir Letter)
            try:
                page_name = (akta_dati.pdf_page_size or "A4").upper().strip()
            except Exception:
                page_name = "A4"

            page_map = {
                "A4": QPageSize.A4,
                "A3": QPageSize.A3,
                "A5": QPageSize.A5,
                "LETTER": QPageSize.Letter,
                "LEGAL": QPageSize.Legal,
            }
            qt_page = page_map.get(page_name, QPageSize.A4)
            printer.setPageSize(QPageSize(qt_page))
            # Orientācija pēc iestatījuma
            try:
                if getattr(akta_dati, "pdf_orientation", "Portrait").lower().startswith("land"):
                    printer.setOrientation(QPrinter.Landscape)
                else:
                    printer.setOrientation(QPrinter.Portrait)
            except Exception:
                pass

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
            # Renderēšana tiek izsaukta vairākas reizes (preview pārzīmēšana). Lai neuzkārtu programmu,
            # kešojam lapu attēlus atmiņā vienam pdf_path.
            if not hasattr(self, "_print_images_cache"):
                self._print_images_cache = {}

            cache_key = os.path.abspath(pdf_path)
            images = self._print_images_cache.get(cache_key)

            if images is None:
                # Pārliecināmies, ka poppler_path ir pieejams
                poppler_path_to_use = self.data.poppler_path if self.data.poppler_path and os.path.exists(self.data.poppler_path) else None
                # Mazliet zemāka DPI vērtība = ātrāk un stabilāk drukas priekšskatījumā
                render_path, tmp_cleanup = _prepare_unencrypted_pdf_for_render(pdf_path, password=getattr(self.data, "pdf_user_password", "") or "")
                try:
                    images = convert_from_path(render_path, poppler_path=poppler_path_to_use, dpi=150)
                finally:
                    try:
                        if tmp_cleanup and os.path.exists(tmp_cleanup):
                            os.remove(tmp_cleanup)
                    except Exception:
                        pass

                self._print_images_cache[cache_key] = images

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

    def izvēlēties_docx_sablonu(self):
        c, _ = QFileDialog.getOpenFileName(self, "Izvēlēties DOCX šablonu", "", "Word dokuments (*.docx)")
        if c:
            self.in_docx_template.setText(c)
            self._update_preview()

    def _on_paraksta_rezims_changed(self, *args, **kwargs):
        try:
            mode = self.cb_paraksta_rezims.currentData() if hasattr(self, 'cb_paraksta_rezims') else 'physical'
            is_physical = (mode == 'physical')
            is_no_sig = (mode == 'no_signature_required')
            for w in (getattr(self, 'in_paraksts_pie', None), getattr(self, 'in_paraksts_nod', None), getattr(self, 'ck_paraksti', None)):
                if w is not None:
                    w.setEnabled(is_physical)
            for w in (getattr(self, 'in_paraksta_vards_rekviziti', None), getattr(self, 'in_paraksta_vards_pienemejs', None), getattr(self, 'in_paraksta_vards_nodevejs', None)):
                if w is not None:
                    w.setEnabled(is_physical)
            if hasattr(self, 'in_paraksta_nav_teksts') and self.in_paraksta_nav_teksts is not None:
                self.in_paraksta_nav_teksts.setEnabled(is_no_sig)
            if hasattr(self, 'ck_elektroniskais_paraksts') and self.ck_elektroniskais_paraksts is not None:
                self.ck_elektroniskais_paraksts.blockSignals(True)
                self.ck_elektroniskais_paraksts.setChecked(mode == 'electronic')
                self.ck_elektroniskais_paraksts.blockSignals(False)
        except Exception:
            pass
        self._update_preview()

    def _sync_signature_mode_from_checkbox(self, *args, **kwargs):
        try:
            if not hasattr(self, 'cb_paraksta_rezims') or self.cb_paraksta_rezims is None:
                self._update_preview()
                return
            target_mode = 'electronic' if bool(self.ck_elektroniskais_paraksts.isChecked()) else 'physical'
            current_mode = str(self.cb_paraksta_rezims.currentData() or 'physical')
            if current_mode != target_mode:
                idx = self.cb_paraksta_rezims.findData(target_mode)
                if idx >= 0:
                    self.cb_paraksta_rezims.blockSignals(True)
                    self.cb_paraksta_rezims.setCurrentIndex(idx)
                    self.cb_paraksta_rezims.blockSignals(False)
            self._on_paraksta_rezims_changed()
        except Exception:
            self._update_preview()

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

        self.in_default_execution_days = QSpinBox()
        self.in_default_execution_days.setRange(0, 365)
        self.in_default_execution_days.setValue(5)
        self.in_default_execution_days.setSuffix(" d.")
        self.in_default_execution_days.setToolTip("Izpildes termiņš = šodiena + N dienas (noklusējums jaunam aktam)")

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
        # --- JAUNS: sinhronizācija ar Pamata datu šifrēšanas lauku (ja tāds ir) ---
        try:
            if hasattr(self, "ck_pdf_encrypt_basic") and hasattr(self, "in_pdf_password_basic"):
                # sākotnējā sinhronizācija
                self.ck_enable_pdf_encryption.setChecked(self.ck_pdf_encrypt_basic.isChecked())
                if self.in_pdf_password_basic.text().strip():
                    self.in_pdf_user_password.setText(self.in_pdf_password_basic.text().strip())

                # divvirzienu sinhronizācija
                self.ck_enable_pdf_encryption.stateChanged.connect(lambda *_: self.ck_pdf_encrypt_basic.setChecked(self.ck_enable_pdf_encryption.isChecked()))
                self.in_pdf_user_password.textChanged.connect(lambda *_: self.in_pdf_password_basic.setText(self.in_pdf_user_password.text()))
                self.ck_pdf_encrypt_basic.stateChanged.connect(lambda *_: self.ck_enable_pdf_encryption.setChecked(self.ck_pdf_encrypt_basic.isChecked()))
                self.in_pdf_password_basic.textChanged.connect(lambda *_: self.in_pdf_user_password.setText(self.in_pdf_password_basic.text()))
        except Exception:
            pass

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
        form.addRow("Izpildes termiņš + dienas (noklusējums):", self.in_default_execution_days)
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
        if sablonu_nosaukums == "Testa dati (piemērs)":
            # ... (esošais Testa dati šablona kods) ...
            # Šeit nav jāmaina, jo iebūvētajam šablonam nav paroles
            d = AktaDati(
                akta_nr="",
                datums=datetime.now().strftime('%Y-%m-%d'),
                vieta="Rīga",
                pasūtījuma_nr="",
                pieņēmējs=Persona(
                    nosaukums="SIA \"Tests\"",
                    reģ_nr="00000000000",
                    adrese="Testa iela 1, Rīga, LV-0000",
                    kontaktpersona="Jānis Tests",
                    tālrunis="+371 200000000",
                    epasts="janis.tests@tests.lv",
                    bankas_konts="LV00BANK1234567890123",
                    juridiskais_statuss="Juridiska persona"
                ),
                nodevējs=Persona(
                    nosaukums="SIA \"Demo Serviss\"",
                    reģ_nr="00000000000",
                    adrese="Testa iela 1, Rīga, LV-0000",
                    kontaktpersona="Jānis Tests",
                    tālrunis="+371 200000000",
                    epasts="anna.tests@demoserviss.lv",
                    bankas_konts="LV00BANK1234567890123",
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
                paraksta_rezims="physical",
                paraksta_nav_teksts="Dokuments ir spēkā bez paraksta, ja tas sagatavots un apstiprināts atbilstoši pušu noteiktajai kārtībai.",
                paraksta_vards_rekviziti="",
                paraksta_vards_pienemejs="",
                paraksta_vards_nodevejs="",
                papildu_parakstu_rindas=[],
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
                pieņēmēja_loma="Pieņēmējs",
                nodevēja_loma="Iekārtas/u /pakalpojuma/u nodevējs",
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
            self._undo_mgr.push_undo(self._snapshot_state('TEMPLATE_LOAD'))
            self.ieviest_datus(d)
            QMessageBox.information(self, "Šablons ielādēts", f"Šablons '{sablonu_nosaukums}' veiksmīgi ielādēts.")
            return
        else:
            # Mēģinām ielādēt no šablonu direktorijas (atbalsta arī vecos ceļus)
            file_path = self._resolve_template_path(sablonu_nosaukums)


        file_path = _coerce_path(file_path)
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
                    doc_tips=data.get('doc_tips', 'akta'),
                    dokumenta_nosaukums=data.get('dokumenta_nosaukums', _doc_type_title(data.get('doc_tips', 'akta'))),
                    apmaksas_termins=data.get('apmaksas_termins', ''),
                    piegades_datums=data.get('piegades_datums', ''),
                    party_mode=data.get('party_mode', DOCUMENT_TYPE_PRESETS.get(data.get('doc_tips', 'akta'), {}).get('party_mode', 'puses')),
                    rekvizitu_virsraksts=data.get('rekvizitu_virsraksts', 'Rekvizīti'),
                    pieņēmēja_loma=data.get('pieņēmēja_loma', 'Pieņēmējs'),
                    nodevēja_loma=data.get('nodevēja_loma', 'Iekārtas/u /pakalpojuma/u nodevējs'),
                    pieņēmējs=Persona(
                        nosaukums=data.get('pieņēmējs', {}).get('nosaukums', ''),
                        reģ_nr=data.get('pieņēmējs', {}).get('reģ_nr', ''),
                        adrese=data.get('pieņēmējs', {}).get('adrese', ''),
                        kontaktpersona=data.get('pieņēmējs', {}).get('kontaktpersona', ''),
                        amats=data.get('pieņēmējs', {}).get('amats', ''),
                        pilnvaras_pamats=data.get('pieņēmējs', {}).get('pilnvaras_pamats', ''),
                        tālrunis=data.get('pieņēmējs', {}).get('tālrunis', ''),
                        epasts=data.get('pieņēmējs', {}).get('epasts', ''),
                        web_lapa=data.get('pieņēmējs', {}).get('web_lapa', ''),
                        bankas_konts=data.get('pieņēmējs', {}).get('bankas_konts', ''),
                        juridiskais_statuss=data.get('pieņēmējs', {}).get('juridiskais_statuss', '')
                    ),
                    nodevējs=Persona(
                        nosaukums=data.get('nodevējs', {}).get('nosaukums', ''),
                        reģ_nr=data.get('nodevējs', {}).get('reģ_nr', ''),
                        adrese=data.get('nodevējs', {}).get('adrese', ''),
                        kontaktpersona=data.get('nodevējs', {}).get('kontaktpersona', ''),
                        amats=data.get('nodevējs', {}).get('amats', ''),
                        pilnvaras_pamats=data.get('nodevējs', {}).get('pilnvaras_pamats', ''),
                        tālrunis=data.get('nodevējs', {}).get('tālrunis', ''),
                        epasts=data.get('nodevējs', {}).get('epasts', ''),
                        web_lapa=data.get('nodevējs', {}).get('web_lapa', ''),
                        bankas_konts=data.get('nodevējs', {}).get('bankas_konts', ''),
                        juridiskais_statuss=data.get('nodevējs', {}).get('juridiskais_statuss', '')
                    ),
                    rekviziti=Persona(
                        nosaukums=data.get('rekviziti', {}).get('nosaukums', ''),
                        reģ_nr=data.get('rekviziti', {}).get('reģ_nr', ''),
                        adrese=data.get('rekviziti', {}).get('adrese', ''),
                        kontaktpersona=data.get('rekviziti', {}).get('kontaktpersona', ''),
                        amats=data.get('rekviziti', {}).get('amats', ''),
                        pilnvaras_pamats=data.get('rekviziti', {}).get('pilnvaras_pamats', ''),
                        tālrunis=data.get('rekviziti', {}).get('tālrunis', ''),
                        epasts=data.get('rekviziti', {}).get('epasts', ''),
                        web_lapa=data.get('rekviziti', {}).get('web_lapa', ''),
                        bankas_konts=data.get('rekviziti', {}).get('bankas_konts', ''),
                        juridiskais_statuss=data.get('rekviziti', {}).get('juridiskais_statuss', '')
                    ),
                    pozīcijas=[Pozīcija(
                        apraksts=p.get('apraksts', ''),
                        daudzums=get_decimal(p, 'daudzums', '0'),
                        vienība=p.get('vienība', 'gab.'),
                        cena=get_decimal(p, 'cena', '0'),
                        seriālais_nr=p.get('seriālais_nr', ''),
                        garantija=p.get('garantija', ''),
                        piezīmes_pozīcijai=p.get('piezīmes_pozīcijai', ''),
                        attēla_ceļš=p.get('attēla_ceļš', '')
                    ) for p in data.get('pozīcijas', [])],
                    custom_columns=data.get('custom_columns', []) if isinstance(data.get('custom_columns', []), list) else [],
                    poz_columns_config=data.get('poz_columns_config', {}) if isinstance(data.get('poz_columns_config', {}), dict) else {},
                    show_price_summary=get_bool(data, 'show_price_summary', True),
                    poz_columns_visual_order=data.get('poz_columns_visual_order', []) if isinstance(data.get('poz_columns_visual_order', []), list) else [],
                    attēli=[Attēls(**a) for a in data.get('attēli', [])],
                    piezīmes=data.get('piezīmes', ''), iekļaut_pvn=get_bool(data, 'iekļaut_pvn', False),
                    pvn_likme=get_decimal(data, 'pvn_likme', '21'),
                    parakstu_rindas=get_bool(data, 'parakstu_rindas', True),
                    paraksta_rezims=data.get('paraksta_rezims', 'electronic' if get_bool(data, 'elektroniskais_paraksts', False) else 'physical'),
                    paraksta_nav_teksts=data.get('paraksta_nav_teksts', 'Dokuments ir spēkā bez paraksta, ja tas sagatavots un apstiprināts atbilstoši pušu noteiktajai kārtībai.'),
                    paraksta_vards_rekviziti=data.get('paraksta_vards_rekviziti', ''),
                    paraksta_vards_pienemejs=data.get('paraksta_vards_pienemejs', ''),
                    paraksta_vards_nodevejs=data.get('paraksta_vards_nodevejs', ''),
                    papildu_parakstu_rindas=data.get('papildu_parakstu_rindas', []) if isinstance(data.get('papildu_parakstu_rindas', []), list) else [],
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
                        apdrošināšana_teksts=data.get('apdrošināšana_teksts', ''),
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
                self._undo_mgr.push_undo(self._snapshot_state('TEMPLATE_LOAD'))
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
                paraksta_rezims="physical",
                paraksta_nav_teksts="Dokuments ir spēkā bez paraksta, ja tas sagatavots un apstiprināts atbilstoši pušu noteiktajai kārtībai.",
                paraksta_vards_rekviziti="",
                paraksta_vards_pienemejs="",
                paraksta_vards_nodevejs="",
                papildu_parakstu_rindas=[],
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
                pieņēmēja_loma="Pieņēmējs",
                nodevēja_loma="Iekārtas/u /pakalpojuma/u nodevējs",
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
            self._undo_mgr.push_undo(self._snapshot_state('TEMPLATE_LOAD'))
            self.ieviest_datus(d)
            QMessageBox.information(self, "Šablons ielādēts", f"Šablons '{sablonu_nosaukums}' veiksmīgi ielādēts.")
            return  # Iziet no funkcijas pēc iebūvētā šablona ielādes
        else:
            # Mēģinām ielādēt no šablonu direktorijas
            file_path = os.path.join(self.data.templates_dir, f"{sablonu_nosaukums}.json")

        file_path = _coerce_path(file_path)
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
                            web_lapa=data.get('nodevējs', {}).get('web_lapa', ''),
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
                            piezīmes_pozīcijai=p.get('piezīmes_pozīcijai', ''),
                            attēla_ceļš=p.get('attēla_ceļš', '')
                        ) for p in data.get('pozīcijas', [])],
                        custom_columns=data.get('custom_columns', []) if isinstance(data.get('custom_columns', []), list) else [],
                        poz_columns_config=data.get('poz_columns_config', {}) if isinstance(data.get('poz_columns_config', {}), dict) else {},
                        show_price_summary=get_bool(data, 'show_price_summary', True),
                    poz_columns_visual_order=data.get('poz_columns_visual_order', []) if isinstance(data.get('poz_columns_visual_order', []), list) else [],
                        attēli=[Attēls(**a) for a in data.get('attēli', [])],
                        piezīmes=data.get('piezīmes', ''), iekļaut_pvn=get_bool(data, 'iekļaut_pvn', False),
                        pvn_likme=get_decimal(data, 'pvn_likme', '21'),
                        parakstu_rindas=get_bool(data, 'parakstu_rindas', True),
                        paraksta_rezims=data.get('paraksta_rezims', 'electronic' if get_bool(data, 'elektroniskais_paraksts', False) else 'physical'),
                        paraksta_nav_teksts=data.get('paraksta_nav_teksts', 'Dokuments ir spēkā bez paraksta, ja tas sagatavots un apstiprināts atbilstoši pušu noteiktajai kārtībai.'),
                        paraksta_vards_rekviziti=data.get('paraksta_vards_rekviziti', ''),
                        paraksta_vards_pienemejs=data.get('paraksta_vards_pienemejs', ''),
                        paraksta_vards_nodevejs=data.get('paraksta_vards_nodevejs', ''),
                    papildu_parakstu_rindas=data.get('papildu_parakstu_rindas', []) if isinstance(data.get('papildu_parakstu_rindas', []), list) else [],
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
                        apdrošināšana_teksts=data.get('apdrošināšana_teksts', ''),
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

                    self._undo_mgr.push_undo(self._snapshot_state('TEMPLATE_LOAD'))
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

        

        
        # --- JAUNS: labā klikšķa konteksta izvēlne adrešu grāmatai ---
        self.address_book_list.setContextMenuPolicy(Qt.CustomContextMenu)
        self.address_book_list.customContextMenuRequested.connect(self._show_address_book_context_menu)
# --- JAUNS: nosaukuma režīms saglabāšanai ---
        self.chk_ab_auto_name = QCheckBox("Nosaukumu ģenerē sistēma automātiski (Uzņēmums + Kontaktpersona)")
        self.chk_ab_auto_name.setChecked(True)
        btn_load_selected = QPushButton("Ielādēt izvēlēto")
        btn_load_selected.clicked.connect(lambda: self._load_selected_address_book_entry(self.address_book_list.currentItem()))
        btn_delete_selected = QPushButton("Dzēst izvēlēto")
        btn_delete_selected.clicked.connect(self._delete_selected_address_book_entry)

        btns_layout = QHBoxLayout()
        btns_layout.addWidget(btn_load_selected)
        btns_layout.addWidget(btn_delete_selected)
        btns_layout.addStretch()

        v.addWidget(QLabel("Saglabātās personas:"))
        v.addWidget(self.chk_ab_auto_name)
        v.addWidget(self.address_book_list)
        v.addLayout(btns_layout)
        v.addStretch()

        w.setLayout(v)
        self.tabs.addTab(w, "Adrešu grāmata")
        self._update_address_book_list()

    # ======================
    # Adrešu grāmata: labais klikšķis + parole + rediģēšana
    # ======================



    def _būvēt_audit_tab(self):
        """Audit tab: rāda pēdējos ierakstus un ļauj eksportēt."""
        tab = QWidget()
        v = QVBoxLayout(tab)

        top = QHBoxLayout()
        self._audit_filter = QLineEdit()
        self._audit_filter.setPlaceholderText("Filtrs (meklē event / user / detaļās)…")
        btn_refresh = QPushButton("Atjaunot")
        btn_refresh.clicked.connect(lambda: self._refresh_audit_table(limit=400))
        btn_export = QPushButton("Eksportēt…")
        btn_export.clicked.connect(self._export_audit_log)

        top.addWidget(QLabel("Audit logs:"))
        top.addWidget(self._audit_filter, 1)
        top.addWidget(btn_refresh)
        top.addWidget(btn_export)
        v.addLayout(top)

        self._audit_table = QTableWidget(0, 4)
        self._audit_table.setHorizontalHeaderLabels(["Laiks", "Lietotājs", "Notikums", "Detaļas"])
        self._audit_table.horizontalHeader().setStretchLastSection(True)
        self._audit_table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self._audit_table.setSelectionBehavior(QAbstractItemView.SelectRows)
        v.addWidget(self._audit_table, 1)

        # Undo/Redo pogas (ērti redzamas)
        h = QHBoxLayout()
        btn_undo = QPushButton("Undo (Ctrl+Z)")
        btn_redo = QPushButton("Redo (Ctrl+Y)")
        btn_undo.clicked.connect(self.undo_action)
        btn_redo.clicked.connect(self.redo_action)
        h.addWidget(btn_undo)
        h.addWidget(btn_redo)
        h.addStretch(1)
        v.addLayout(h)

        self._audit_filter.textChanged.connect(lambda: self._refresh_audit_table(limit=400))
        self.tabs.addTab(tab, "Audit")

        self._refresh_audit_table(limit=200)

    def _color_for_audit_event(self, event_name: str):
        try:
            e = (event_name or "").upper()
            if e.startswith("AB_"):
                return QColor(219, 234, 254)  # light blue
            if e.startswith("PROJECT_"):
                return QColor(220, 252, 231)  # light green
            if e.startswith("GENERATE_"):
                return QColor(255, 237, 213)  # light orange
            if e.startswith("POZ_"):
                return QColor(243, 232, 255)  # light purple
            if e.startswith("FIELD_") or e.startswith("COMBO_") or e.startswith("CHECK_"):
                return QColor(241, 245, 249)  # light slate
            if e in ("UNDO", "REDO"):
                return QColor(226, 232, 240)  # gray
            return None
        except Exception:
            return None

    def _refresh_audit_table(self, limit: int = 200):
        if not hasattr(self, "_audit_table") or self._audit_table is None:
            return
        flt = (self._audit_filter.text() or "").strip().lower() if hasattr(self, "_audit_filter") else ""
        rows = self._audit_logger.tail(limit)
        # filtrējam
        if flt:
            def ok(r):
                try:
                    s = (r.get("ts","") + " " + r.get("user","") + " " + r.get("event","") + " " + json.dumps(r.get("details",{}), ensure_ascii=False)).lower()
                    return flt in s
                except Exception:
                    return True
            rows = [r for r in rows if ok(r)]

        self._audit_table.setRowCount(0)
        for r in rows[::-1]:  # newest first
            row = self._audit_table.rowCount()
            self._audit_table.insertRow(row)
            self._audit_table.setItem(row, 0, QTableWidgetItem(str(r.get("ts",""))))
            self._audit_table.setItem(row, 1, QTableWidgetItem(str(r.get("user",""))))
            self._audit_table.setItem(row, 2, QTableWidgetItem(str(r.get("event",""))))
            it0 = QTableWidgetItem(str(r.get("ts","")))
            it1 = QTableWidgetItem(str(r.get("user","")))
            it2 = QTableWidgetItem(str(r.get("event","")))
            it3 = QTableWidgetItem(json.dumps(r.get("details", {}), ensure_ascii=False))
            self._audit_table.setItem(row, 0, it0)
            self._audit_table.setItem(row, 1, it1)
            self._audit_table.setItem(row, 2, it2)
            self._audit_table.setItem(row, 3, it3)
            bg = self._color_for_audit_event(str(r.get("event","")))
            if bg is not None:
                for it in (it0, it1, it2, it3):
                    it.setBackground(bg)


    def _update_undo_redo_indicators(self):
        try:
            if not hasattr(self, "_undo_mgr"):
                return
            u = len(getattr(self._undo_mgr, "_undo", []))
            r = len(getattr(self._undo_mgr, "_redo", []))
            if getattr(self, "_undo_status_label", None) is not None:
                self._undo_status_label.setText(f"Undo: {u}")
                self._undo_status_label.setStyleSheet("padding:2px; color: " + ("#16a34a" if u>0 else "#94a3b8"))
            if getattr(self, "_redo_status_label", None) is not None:
                self._redo_status_label.setText(f"Redo: {r}")
                self._redo_status_label.setStyleSheet("padding:2px; color: " + ("#2563eb" if r>0 else "#94a3b8"))
        except Exception:
            pass

    def eventFilter(self, obj, event):
        # Globāla izmaiņu izsekošana: FIELD_EDIT / COMBO_CHANGE / CHECK_TOGGLE
        try:
            if not getattr(self, "_track_enabled", False):
                return super().eventFilter(obj, event)

            et = event.type()
            # LineEdit: audit uz FocusOut (kad beidz rediģēt)
            if isinstance(obj, QLineEdit) and et == QEvent.FocusOut:
                key = id(obj)
                cur = obj.text()
                prev = self._last_widget_values.get(key, None)
                if prev is None:
                    self._last_widget_values[key] = cur
                    return super().eventFilter(obj, event)
                if cur != prev:
                    self._undo_mgr.push_undo(self._snapshot_state("FIELD_EDIT"))
                    self._audit("FIELD_EDIT", {"field": obj.objectName() or obj.placeholderText() or "QLineEdit", "value": (cur or "")[:200]})
                    self._last_widget_values[key] = cur
                    self._update_undo_redo_indicators()
                    self._update_preview()
                return super().eventFilter(obj, event)

            # ComboBox: audit uz FocusOut
            if isinstance(obj, QComboBox) and et == QEvent.FocusOut:
                key = id(obj)
                cur = obj.currentText()
                prev = self._last_widget_values.get(key, None)
                if prev is None:
                    self._last_widget_values[key] = cur
                    return super().eventFilter(obj, event)
                if cur != prev:
                    self._undo_mgr.push_undo(self._snapshot_state("COMBO_CHANGE"))
                    self._audit("COMBO_CHANGE", {"field": obj.objectName() or "QComboBox", "value": (cur or "")[:200]})
                    self._last_widget_values[key] = cur
                    self._update_undo_redo_indicators()
                    self._update_preview()
                return super().eventFilter(obj, event)

            # CheckBox: audit uz MouseButtonRelease (toggling)
            if isinstance(obj, QCheckBox) and et == QEvent.MouseButtonRelease:
                key = id(obj)
                cur = bool(obj.isChecked())
                prev = self._last_widget_values.get(key, None)
                if prev is None:
                    self._last_widget_values[key] = cur
                    return super().eventFilter(obj, event)
                if cur != prev:
                    self._undo_mgr.push_undo(self._snapshot_state("CHECK_TOGGLE"))
                    self._audit("CHECK_TOGGLE", {"field": obj.text() or obj.objectName() or "QCheckBox", "value": cur})
                    self._last_widget_values[key] = cur
                    self._update_undo_redo_indicators()
                    self._update_preview()
                return super().eventFilter(obj, event)

        except Exception:
            pass
        return super().eventFilter(obj, event)

    def _ensure_positions_undo_hook(self):
        # Pieslēdz undo/redo checkpoint pozīciju tabulas izmaiņām (debounce)
        try:
            if not hasattr(self, "tab") or self.tab is None:
                return
            if hasattr(self, "_pos_change_timer") and self._pos_change_timer is not None:
                return
            self._pos_change_timer = QTimer(self)
            self._pos_change_timer.setSingleShot(True)
            self._pos_change_timer.timeout.connect(self._commit_positions_change)
            try:
                self.tab.cellChanged.connect(self._on_positions_cell_changed)
            except Exception:
                pass
        except Exception:
            pass

    def _on_positions_cell_changed(self, row: int, col: int):
        try:
            self._last_pos_change = (row, col)
            if hasattr(self, "_pos_change_timer") and self._pos_change_timer is not None:
                self._pos_change_timer.start(700)
        except Exception:
            pass

    def _commit_positions_change(self):
        try:
            self._undo_mgr.push_undo(self._snapshot_state("POZ_CHANGE"))
            rc = getattr(self, "_last_pos_change", None)
            details = {"row": rc[0], "col": rc[1]} if rc else {}
            self._audit("POZ_CHANGE", details)
            self._update_undo_redo_indicators()
        except Exception:
            pass


    def _export_audit_log(self):
        try:
            default = os.path.join(PROJECT_SAVE_DIR, "audit_export.jsonl")
            fn, _ = QFileDialog.getSaveFileName(self, "Eksportēt audit log", default, "JSONL (*.jsonl);;Teksts (*.txt)")
            if not fn:
                return
            srcp = self._audit_logger.log_path
            if srcp and os.path.exists(srcp):
                shutil.copy2(srcp, fn)
                QMessageBox.information(self, "OK", "Audit logs eksportēts.")
            else:
                QMessageBox.warning(self, "Nav", "Audit logs fails vēl nav izveidots.")
        except Exception as e:
            QMessageBox.warning(self, "Kļūda", f"Neizdevās eksportēt: {e}")


    def _ab_hash_password(self, password: str, salt: str) -> str:
        # Hash parolei (PBKDF2-HMAC-SHA256).
        try:
            dk = hashlib.pbkdf2_hmac("sha256", password.encode("utf-8"), salt.encode("utf-8"), 120_000)
            return dk.hex()
        except Exception as e:
            print(f"Paroles hash kļūda: {e}")
            return ""

    def _ab_verify_password(self, persona_data: dict, password: str) -> bool:
        # Pārbauda paroli pret saglabāto hash.
        try:
            salt = (persona_data.get("__pw_salt") or "").strip()
            h = (persona_data.get("__pw_hash") or "").strip()
            if not salt or not h:
                return True
            return self._ab_hash_password(password, salt) == h
        except Exception:
            return False

    def _ab_require_password(self, entry_name: str, persona_data: dict, action_label: str = "ielādēt") -> bool:
        # Ja pusei ir parole, prasa to pirms ielādes/rediģēšanas.
        try:
            if not persona_data:
                return False
            if not (persona_data.get("__pw_hash") and persona_data.get("__pw_salt")):
                return True

            pwd, ok = QInputDialog.getText(
                self,
                "Parole nepieciešama",
                f"Lai {action_label} '{entry_name}', ievadi paroli:",
                QLineEdit.Password,
                ""
            )
            if not ok:
                return False
            if self._ab_verify_password(persona_data, pwd):
                return True

            QMessageBox.warning(self, "Nepareiza parole", "Parole nav pareiza.")
            return False
        except Exception:
            return False

    def _show_address_book_context_menu(self, pos):
        # Labā klikšķa izvēlne adrešu grāmatas sarakstam.
        item = self.address_book_list.itemAt(pos)
        menu = QMenu(self)

        act_load = QAction("Ielādēt", self)
        act_edit = QAction("Rediģēt…", self)
        act_set_pw = QAction("Uzlikt / mainīt paroli…", self)
        act_remove_pw = QAction("Noņemt paroli", self)
        act_delete = QAction("Dzēst", self)

        act_rename = QAction("Pārdēvēt…", self)
        act_duplicate = QAction("Dublēt", self)
        act_export_json = QAction("Eksportēt (JSON)…", self)
        act_export_csv = QAction("Eksportēt (CSV)…", self)

        if item is None:
            act_refresh = QAction("Atjaunot sarakstu", self)
            act_refresh.triggered.connect(self._update_address_book_list)
            menu.addAction(act_refresh)
            menu.exec(self.address_book_list.mapToGlobal(pos))
            return

        name = item.text()
        persona_data = self.address_book.get(name, {})

        act_rename.triggered.connect(lambda: self._rename_address_book_entry(name))
        act_duplicate.triggered.connect(lambda: self._duplicate_address_book_entry(name))
        act_export_json.triggered.connect(lambda: self._export_address_book_entry_json(name))
        act_export_csv.triggered.connect(lambda: self._export_address_book_entry_csv(name))

        act_load.triggered.connect(lambda: self._load_selected_address_book_entry(item))
        act_edit.triggered.connect(lambda: self._edit_address_book_entry(name))
        act_set_pw.triggered.connect(lambda: self._set_password_for_address_book_entry(name))
        act_remove_pw.triggered.connect(lambda: self._remove_password_for_address_book_entry(name))
        act_delete.triggered.connect(lambda: self._delete_address_book_entry(name))

        menu.addAction(act_load)
        menu.addAction(act_edit)
        menu.addAction(act_rename)
        menu.addAction(act_duplicate)
        menu.addSeparator()
        menu.addAction(act_export_json)
        menu.addAction(act_export_csv)
        menu.addSeparator()
        menu.addAction(act_set_pw)
        menu.addAction(act_remove_pw)
        menu.addSeparator()
        menu.addAction(act_delete)

        has_pw = bool(persona_data.get("__pw_hash") and persona_data.get("__pw_salt"))
        act_remove_pw.setEnabled(has_pw)

        menu.exec(self.address_book_list.mapToGlobal(pos))

    def _set_password_for_address_book_entry(self, entry_name: str):
        self._undo_mgr.push_undo(self._snapshot_state('AB_SET_PASSWORD'))
        self._audit('AB_SET_PASSWORD', {})
        persona_data = self.address_book.get(entry_name)
        if not persona_data:
            return

        if persona_data.get("__pw_hash") and persona_data.get("__pw_salt"):
            if not self._ab_require_password(entry_name, persona_data, "mainīt paroli"):
                return

        pw1, ok1 = QInputDialog.getText(self, "Uzlikt paroli", f"Ievadi jauno paroli '{entry_name}':", QLineEdit.Password, "")
        if not ok1:
            return
        pw2, ok2 = QInputDialog.getText(self, "Uzlikt paroli", "Atkārto paroli:", QLineEdit.Password, "")
        if not ok2:
            return
        if pw1 != pw2:
            QMessageBox.warning(self, "Kļūda", "Paroles nesakrīt.")
            return
        if not pw1.strip():
            QMessageBox.warning(self, "Kļūda", "Parole nevar būt tukša.")
            return

        salt = secrets.token_hex(8)
        h = self._ab_hash_password(pw1, salt)
        if not h:
            QMessageBox.warning(self, "Kļūda", "Neizdevās uzlikt paroli (hash tukšs).")
            return

        persona_data["__pw_salt"] = salt
        persona_data["__pw_hash"] = h
        self.address_book[entry_name] = persona_data
        self._save_address_book()
        QMessageBox.information(self, "OK", "Parole uzlikta.")

    def _remove_password_for_address_book_entry(self, entry_name: str):
        self._undo_mgr.push_undo(self._snapshot_state('AB_REMOVE_PASSWORD'))
        self._audit('AB_REMOVE_PASSWORD', {})
        persona_data = self.address_book.get(entry_name)
        if not persona_data:
            return
        if persona_data.get("__pw_hash") and persona_data.get("__pw_salt"):
            if not self._ab_require_password(entry_name, persona_data, "noņemt paroli"):
                return
        persona_data.pop("__pw_hash", None)
        persona_data.pop("__pw_salt", None)
        self.address_book[entry_name] = persona_data
        self._save_address_book()
        QMessageBox.information(self, "OK", "Parole noņemta.")

    def _rename_address_book_entry(self, old_name: str):
        self._undo_mgr.push_undo(self._snapshot_state('AB_RENAME'))
        self._audit('AB_RENAME', {})
        if old_name not in self.address_book:
            return
        persona_data = self.address_book.get(old_name, {})
        # Ja ir parole, prasām to pirms pārdēvēšanas
        if not self._ab_require_password(old_name, persona_data, "pārdēvēt"):
            return

        new_name, ok = QInputDialog.getText(
            self,
            "Pārdēvēt ierakstu",
            "Jaunais nosaukums:",
            QLineEdit.Normal,
            old_name
        )
        if not ok:
            return
        new_name = (new_name or "").strip()
        if not new_name:
            QMessageBox.warning(self, "Kļūda", "Nosaukums nevar būt tukšs.")
            return
        if new_name == old_name:
            return

        # Unikāls nosaukums
        if hasattr(self, "_ab_make_unique_key"):
            new_name = self._ab_make_unique_key(new_name)
        else:
            if new_name in self.address_book:
                i = 2
                base = new_name
                while f"{base} ({i})" in self.address_book:
                    i += 1
                new_name = f"{base} ({i})"

        self.address_book[new_name] = persona_data
        self.address_book.pop(old_name, None)
        self._save_address_book()
        self._update_address_book_list()

    def _duplicate_address_book_entry(self, src_name: str):
        self._undo_mgr.push_undo(self._snapshot_state('AB_DUPLICATE'))
        self._audit('AB_DUPLICATE', {})
        if src_name not in self.address_book:
            return
        persona_data = dict(self.address_book.get(src_name, {}))  # shallow copy
        # Ja ir parole, prasām to pirms dublēšanas
        if not self._ab_require_password(src_name, persona_data, "dublēt"):
            return

        base = f"{src_name} (kopija)"
        if hasattr(self, "_ab_make_unique_key"):
            new_name = self._ab_make_unique_key(base)
        else:
            new_name = base
            if new_name in self.address_book:
                i = 2
                while f"{base} ({i})" in self.address_book:
                    i += 1
                new_name = f"{base} ({i})"

        self.address_book[new_name] = persona_data
        self._save_address_book()
        self._update_address_book_list()
        QMessageBox.information(self, "OK", f"Ieraksts dublēts kā: {new_name}")

    def _export_address_book_entry_json(self, entry_name: str):
        if entry_name not in self.address_book:
            return
        persona_data = dict(self.address_book.get(entry_name, {}))
        if not self._ab_require_password(entry_name, persona_data, "eksportēt"):
            return

        # Neeksportējam paroles hash/salt
        persona_data.pop("__pw_hash", None)
        persona_data.pop("__pw_salt", None)

        filename, _ = QFileDialog.getSaveFileName(
            self,
            "Eksportēt JSON",
            f"{entry_name}.json",
            "JSON faili (*.json)"
        )
        if not filename:
            return
        try:
            with open(filename, "w", encoding="utf-8") as f:
                json.dump({"name": entry_name, "data": persona_data}, f, ensure_ascii=False, indent=2)
            QMessageBox.information(self, "OK", "JSON eksports pabeigts.")
        except Exception as e:
            QMessageBox.warning(self, "Kļūda", f"Neizdevās eksportēt JSON: {e}")

    def _export_address_book_entry_csv(self, entry_name: str):
        if entry_name not in self.address_book:
            return
        persona_data = dict(self.address_book.get(entry_name, {}))
        if not self._ab_require_password(entry_name, persona_data, "eksportēt"):
            return

        # Neeksportējam paroles hash/salt
        persona_data.pop("__pw_hash", None)
        persona_data.pop("__pw_salt", None)

        filename, _ = QFileDialog.getSaveFileName(
            self,
            "Eksportēt CSV",
            f"{entry_name}.csv",
            "CSV faili (*.csv)"
        )
        if not filename:
            return
        try:
            keys = list(persona_data.keys())
            with open(filename, "w", newline="", encoding="utf-8") as f:
                w = csv.DictWriter(f, fieldnames=keys, delimiter=";")
                w.writeheader()
                w.writerow(persona_data)
            QMessageBox.information(self, "OK", "CSV eksports pabeigts.")
        except Exception as e:
            QMessageBox.warning(self, "Kļūda", f"Neizdevās eksportēt CSV: {e}")


    def _delete_address_book_entry(self, entry_name: str):
        self._undo_mgr.push_undo(self._snapshot_state('AB_DELETE'))
        self._audit('AB_DELETE', {})
        if entry_name not in self.address_book:
            return
        persona_data = self.address_book.get(entry_name, {})
        # --- JAUNS: dzēšana ir aizsargāta ar paroli (ja tā uzlikta) ---
        if persona_data.get('__pw_hash') and persona_data.get('__pw_salt'):
            if not self._ab_require_password(entry_name, persona_data, 'dzēst'):
                return
        reply = QMessageBox.question(
            self, "Dzēst ierakstu", f"Dzēst '{entry_name}' no adrešu grāmatas?",
            QMessageBox.Yes | QMessageBox.No, QMessageBox.No
        )
        if reply != QMessageBox.Yes:
            return
        self.address_book.pop(entry_name, None)
        self._save_address_book()
        self._update_address_book_list()

    def _edit_address_book_entry(self, entry_name: str):
        self._undo_mgr.push_undo(self._snapshot_state('AB_EDIT'))
        self._audit('AB_EDIT', {})
        # Atver jaunu logu rediģēšanai.
        persona_data = self.address_book.get(entry_name)
        if not persona_data:
            return

        if not self._ab_require_password(entry_name, persona_data, "rediģēt"):
            return

        dlg = QDialog(self)
        dlg.setWindowTitle(f"Rediģēt: {entry_name}")
        layout = QVBoxLayout(dlg)
        form = QFormLayout()

        fields = [
            ("Nosaukums", "nosaukums"),
            ("Reģ. Nr.", "reģ_nr"),
            ("Adrese", "adrese"),
            ("Kontaktpersona", "kontaktpersona"),
            ("Amats", "amats"),
            ("Pilnvaras pamats", "pilnvaras_pamats"),
            ("Tālrunis", "tālrunis"),
            ("E-pasts", "epasts"),
            ("Web lapa", "web_lapa"),
            ("Bankas konts", "bankas_konts"),
            ("Juridiskais statuss", "juridiskais_statuss"),
        ]

        widgets = {}

        for label, key in fields:
            if key == "pilnvaras_pamats":
                cb = QComboBox()
                try:
                    if hasattr(self, "pie_in") and self.pie_in and hasattr(self.pie_in[5], "count"):
                        for i in range(self.pie_in[5].count()):
                            cb.addItem(self.pie_in[5].itemText(i))
                    else:
                        cb.addItems(["Pilnvaras pamats", "Cits"])
                except Exception:
                    cb.addItems(["Pilnvaras pamats", "Cits"])
                val = persona_data.get(key, "")
                idx = cb.findText(val)
                if idx >= 0:
                    cb.setCurrentIndex(idx)
                widgets[key] = cb
                form.addRow(label + ":", cb)
            elif key == "juridiskais_statuss":
                cb = QComboBox()
                try:
                    if hasattr(self, "pie_in") and self.pie_in and hasattr(self.pie_in[10], "count"):
                        for i in range(self.pie_in[10].count()):
                            cb.addItem(self.pie_in[10].itemText(i))
                    else:
                        cb.addItems(["Juridiska persona", "Fiziska persona"])
                except Exception:
                    cb.addItems(["Juridiska persona", "Fiziska persona"])
                val = persona_data.get(key, "")
                idx = cb.findText(val)
                if idx >= 0:
                    cb.setCurrentIndex(idx)
                widgets[key] = cb
                form.addRow(label + ":", cb)
            else:
                le = QLineEdit(str(persona_data.get(key, "") or ""))
                widgets[key] = le
                form.addRow(label + ":", le)

        layout.addLayout(form)

        buttons = QDialogButtonBox(QDialogButtonBox.Save | QDialogButtonBox.Cancel)
        layout.addWidget(buttons)

        def on_save():
            try:
                for _, key in fields:
                    w = widgets.get(key)
                    if w is None:
                        continue
                    if isinstance(w, QComboBox):
                        persona_data[key] = w.currentText()
                    else:
                        persona_data[key] = w.text().strip()
                self.address_book[entry_name] = persona_data
                self._save_address_book()
                self._update_address_book_list()
                dlg.accept()
            except Exception as e:
                QMessageBox.warning(self, "Kļūda", f"Neizdevās saglabāt: {e}")

        buttons.accepted.connect(on_save)
        buttons.rejected.connect(dlg.reject)
        dlg.exec()

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
            # --- JAUNS: ja ir parole, prasa to pirms ielādes ---
            if not self._ab_require_password(name, persona_data, "ielādēt"):
                return
            reply = QMessageBox.question(self, "Ielādēt personu",
                                         f"Ielādēt '{name}' kā Pieņēmēju vai Nodevēju?",
                                         QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No | QMessageBox.StandardButton.Cancel,
                                         QMessageBox.StandardButton.Yes)
            if reply == QMessageBox.StandardButton.Yes:
                self.pie_in[0].setText(persona_data.get("nosaukums", ""))
                self.pie_in[1].setText(persona_data.get("reģ_nr", ""))
                self.pie_in[2].setText(persona_data.get("adrese", ""))
                self.pie_in[3].setText(persona_data.get("kontaktpersona", ""))
                self.pie_in[4].setText(persona_data.get("amats", ""))
                pilnvaras_val = persona_data.get("pilnvaras_pamats", "")
                idx_p = self.pie_in[5].findText(pilnvaras_val)
                if idx_p >= 0:
                    self.pie_in[5].setCurrentIndex(idx_p)
                self.pie_in[6].setText(persona_data.get("tālrunis", ""))
                self.pie_in[7].setText(persona_data.get("epasts", ""))
                self.pie_in[8].setText(persona_data.get("web_lapa", ""))
                self.pie_in[9].setText(persona_data.get("bankas_konts", ""))
                juridiskais_statuss_val = persona_data.get("juridiskais_statuss", "")
                idx = self.pie_in[10].findText(juridiskais_statuss_val)
                if idx >= 0:
                    self.pie_in[10].setCurrentIndex(idx)
                QMessageBox.information(self, "Ielādēts", f"Persona '{name}' ielādēta kā Pieņēmējs.")
            elif reply == QMessageBox.StandardButton.No:
                self.nod_in[0].setText(persona_data.get("nosaukums", ""))
                self.nod_in[1].setText(persona_data.get("reģ_nr", ""))
                self.nod_in[2].setText(persona_data.get("adrese", ""))
                self.nod_in[3].setText(persona_data.get("kontaktpersona", ""))
                self.nod_in[4].setText(persona_data.get("amats", ""))
                pilnvaras_val = persona_data.get("pilnvaras_pamats", "")
                idx_p = self.nod_in[5].findText(pilnvaras_val)
                if idx_p >= 0:
                    self.nod_in[5].setCurrentIndex(idx_p)
                self.nod_in[6].setText(persona_data.get("tālrunis", ""))
                self.nod_in[7].setText(persona_data.get("epasts", ""))
                self.nod_in[8].setText(persona_data.get("web_lapa", ""))
                self.nod_in[9].setText(persona_data.get("bankas_konts", ""))
                juridiskais_statuss_val = persona_data.get("juridiskais_statuss", "")
                idx = self.nod_in[10].findText(juridiskais_statuss_val)
                if idx >= 0:
                    self.nod_in[10].setCurrentIndex(idx)
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
                try:
                    self._update_preview()
                except Exception:
                    pass
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

    
    def _record_generated_document(self, pdf_path: str, json_path: str):
        """Saglabā PDF+JSON kopijas dokumentu vēsturē, lai tās vienmēr būtu pieejamas."""
        try:
            os.makedirs(DOCUMENTS_DIR, exist_ok=True)
            ts = datetime.now().strftime("%Y%m%d_%H%M%S")
            base = f"Akts_{ts}"
            dst_pdf = os.path.join(DOCUMENTS_DIR, base + ".pdf")
            dst_json = os.path.join(DOCUMENTS_DIR, base + ".json")
            shutil.copy2(pdf_path, dst_pdf)
            shutil.copy2(json_path, dst_json)

            # Vēsture tagad ir saraksts ar dict ierakstiem
            if not isinstance(self.history, list):
                self.history = []
            # migrācija: ja bija vecais formāts (stringi)
            new_hist = []
            for it in self.history:
                if isinstance(it, str):
                    new_hist.append({"json": it, "pdf": "", "created": ""})
                elif isinstance(it, dict):
                    new_hist.append(it)
            self.history = new_hist

            self.history.insert(0, {"pdf": dst_pdf, "json": dst_json, "created": datetime.now().isoformat(timespec="seconds")})
            self._save_history()
            self._update_history_list()
        except Exception as e:
            print(f"Vēstures saglabāšanas kļūda: {e}")

    def _add_to_history(self, file_path: str):
        file_path = _coerce_path(file_path)
        if not file_path:
            return
        if not os.path.exists(file_path):
            return

        # Noņemam, ja jau ir sarakstā, lai pārvietotu uz saraksta sākumu
        self.history = [f for f in self.history if f != file_path]

        # Pievienojam saraksta sākumā
        self.history.insert(0, file_path)
        self._save_history() # Saglabājam izmaiņas failā
        self._update_history_list() # Atjaunojam sarakstu GUI

    def _update_history_list(self):
        self.history_list.clear()

        # Migrācija no vecā formāta (saraksts ar failu ceļiem)
        if isinstance(self.history, list) and self.history and isinstance(self.history[0], str):
            self.history = [{"json": p, "pdf": "", "created": ""} for p in self.history]

        valid = []
        for it in (self.history or []):
            if isinstance(it, dict):
                j = it.get("json", "")
                j = _coerce_path(j) or ""
                p = it.get("pdf", "")
                p = _coerce_path(p) or ""
                if j and os.path.exists(j) or (p and os.path.exists(p)):
                    valid.append(it)
            elif isinstance(it, str) and os.path.exists(it):
                valid.append({"json": it, "pdf": "", "created": ""})

        self.history = valid
        self._save_history()

        for it in self.history:
            created = it.get("created", "")
            label = os.path.basename(it.get("pdf") or it.get("json") or "")
            if created:
                label = f"{label}  ({created.replace('T',' ')})"
            item = QListWidgetItem(label)
            item.setData(Qt.UserRole, it)
            self.history_list.addItem(item)

    def _load_history_entry(self, item: QListWidgetItem):
        if not item:
            return
        payload = item.data(Qt.UserRole)
        if isinstance(payload, dict):
            json_path = payload.get("json", "")
            json_path = _coerce_path(json_path) or ""
            if json_path and os.path.exists(json_path):
                self.ieladet_projektu(json_path)
                return
            # vecs ieraksts bez json
            pdf_path = payload.get("pdf", "")
            pdf_path = _coerce_path(pdf_path) or ""
            if pdf_path and os.path.exists(pdf_path):
                QDesktopServices.openUrl(QUrl.fromLocalFile(pdf_path))
                return
        # fallback (ja kāds vecs ieraksts)
        txt = item.text().split("  (")[0]
        # mēģinām atrast pēc nosaukuma
        for it in self.history or []:
            if isinstance(it, dict):
                p = it.get("json") or it.get("pdf")
                if p and os.path.basename(p) == txt and os.path.exists(p):
                    if p.lower().endswith(".json"):
                        self.ieladet_projektu(p)
                    else:
                        QDesktopServices.openUrl(QUrl.fromLocalFile(p))
                    return

        QMessageBox.warning(self, "Kļūda", "Neizdevās atrast vēstures ieraksta failu.")

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
        file_path = next(((_coerce_path(f) or "") for f in self.history if _coerce_path(f) and os.path.basename(_coerce_path(f)) == file_name), None)
        file_path = _coerce_path(file_path)
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

        # Meklēšana kartē (Photon) — ļauj atrast adresi un atlasīt no saraksta
        search_row = QHBoxLayout()
        self.map_search_input = QLineEdit()
        self.map_search_input.setPlaceholderText("Meklēt adresi (piem. Brīvības iela 1, Rīga)")
        btn_search = QPushButton("Meklēt")
        btn_search.clicked.connect(self._search_address_on_map)
        search_row.addWidget(self.map_search_input)
        search_row.addWidget(btn_search)

        self.map_search_results = QListWidget()
        self.map_search_results.setMaximumHeight(140)
        self.map_search_results.itemDoubleClicked.connect(self._apply_map_search_result)

        v.addLayout(search_row)
        v.addWidget(self.map_search_results)

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
        self.tab_kartes = w

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

    

    def _search_address_on_map(self):
        q = (self.map_search_input.text() if hasattr(self, "map_search_input") else "").strip()
        self.map_search_results.clear()
        if not q:
            return
        try:
            # Photon: ātrs un parasti neliek 403 (OpenStreetMap Nominatim bieži bloķē bez User-Agent)
            r = requests.get(
                "https://photon.komoot.io/api/",
                # Photon dažās instancēs neatbalsta `lang` parametru (met 400). Tāpēc to nelietojam.
                params={"q": q, "limit": 10},
                headers={
                    "User-Agent": "AktaGenerators/1.0 (contact: edgars@kulinics.id.lv)",
                    "Accept-Language": "lv,en;q=0.8",
                },
                timeout=10
            )
            r.raise_for_status()
            data = r.json()
            feats = data.get("features", []) or []
            for f in feats:
                props = f.get("properties", {}) or {}
                geom = f.get("geometry", {}) or {}
                coords = geom.get("coordinates", None)
                if not coords or len(coords) < 2:
                    continue
                lon, lat = coords[0], coords[1]
                # Saudzīgi saliekam adresi
                name = props.get("name") or ""
                street = props.get("street") or ""
                housenumber = props.get("housenumber") or ""
                city = props.get("city") or props.get("town") or props.get("village") or ""
                country = props.get("country") or ""
                parts = []
                if name:
                    parts.append(name)
                addr_line = " ".join([p for p in [street, housenumber] if p]).strip()
                if addr_line:
                    parts.append(addr_line)
                if city:
                    parts.append(city)
                if country:
                    parts.append(country)
                label = ", ".join([p for p in parts if p]).strip()
                if not label:
                    label = q
                it = QListWidgetItem(label)
                it.setData(Qt.UserRole, {"lat": lat, "lon": lon, "label": label})
                self.map_search_results.addItem(it)
        except Exception as e:
            QMessageBox.warning(self, "Karte", f"Neizdevās meklēt adresi: {e}")

    def _apply_map_search_result(self, item: QListWidgetItem):
        if not item:
            return
        d = item.data(Qt.UserRole) or {}
        try:
            lat = float(d.get("lat"))
            lon = float(d.get("lon"))
        except Exception:
            return
        label = d.get("label", "")
        self.map_lat_input.setText(f"{lat:.6f}")
        self.map_lon_input.setText(f"{lon:.6f}")
        # Iestatām marķieri kartē
        js = f"setMarker({lat:.6f}, {lon:.6f}, {json.dumps(label, ensure_ascii=False)});"
        self.web_view.page().runJavaScript(js)
    def _set_map_marker_from_inputs(self):
        try:
            lat = float(self.map_lat_input.text())
            lon = float(self.map_lon_input.text())
            self.web_view.page().runJavaScript(f"setMarker({lat}, {lon}, 'Izvēlētā vieta');")
        except ValueError:
            QMessageBox.warning(self, "Kļūda", "Lūdzu, ievadiet derīgas platuma un garuma vērtības.")

    def _begin_address_pick(self, target_lineedit: QLineEdit):
        """
        Aktivizē adreses izvēli kartē konkrētam laukam (Pieņēmējs/Nodevējs).
        Lietotājs var arī rakstīt adresi ar roku – šī ir tikai ērtība.
        """
        self._map_address_target = target_lineedit
        try:
            if hasattr(self, "tab_kartes"):
                self.tabs.setCurrentWidget(self.tab_kartes)
        except Exception:
            pass
        QMessageBox.information(
            self,
            "Adrese no kartes",
            "Noklikšķini kartē uz nepieciešamās vietas – adrese tiks automātiski ielikta laukā."
        )

    def _handle_map_click(self, lat, lon):
        """
        Šī funkcija tiek izsaukta no MapUrlInterceptor, kad kartē tiek noklikšķināts.
        Tā veic reverso ģeokodēšanu un aizpilda adresi:
        - ja lietotājs izvēlējās adresi (Pieņēmējs/Nodevējs), tad aizpilda attiecīgo lauku,
        - citādi aizpilda akta "Vieta" lauku (kā līdz šim).
        """
        self.map_lat_input.setText(lat)
        self.map_lon_input.setText(lon)
        self._reverse_geocode_and_set_vieta(lat, lon)

    def _reverse_geocode_and_set_vieta(self, lat, lon):
        """Reversā ģeokodēšana: mēģina Nominatim; ja 403/429 vai kļūda – izmanto Photon fallback."""
        try:
            lat_s = str(lat).strip()
            lon_s = str(lon).strip()

            headers = {
                # Nominatim prasa identificējamu UA (ar kontaktu)
                "User-Agent": "AktaGenerators/1.0 (kulinics.id.lv; edgars@kulinics.id.lv)",
                "Accept-Language": "lv,en;q=0.8",
                "From": "edgars@kulinics.id.lv",
            }

            # 1) Nominatim
            try:
                resp = requests.get(
                    "https://nominatim.openstreetmap.org/reverse",
                    params={
                        "format": "json",
                        "lat": lat_s,
                        "lon": lon_s,
                        "zoom": 18,
                        "addressdetails": 1
                    },
                    headers=headers,
                    timeout=8
                )
                if resp.status_code not in (403, 429):
                    resp.raise_for_status()
                    data = resp.json() if resp.text else {}

                    if isinstance(data, dict) and data.get("display_name"):
                        address = data["display_name"]

                        target = getattr(self, "_map_address_target", None)
                        if target is not None:
                            target.setText(address)
                            self._map_address_target = None
                        else:
                            self.in_vieta.setText(address)
                        QMessageBox.information(self, "Adrese iegūta", f"Adrese: {address}")
                        return
            except Exception:
                # ignorējam un mēģinam Photon
                pass

            # 2) Photon fallback (bieži strādā, kad Nominatim bloķē)
            presp = requests.get(
                "https://photon.komoot.io/reverse",
                params={"lat": lat_s, "lon": lon_s},
                headers=headers,
                timeout=8
            )
            presp.raise_for_status()
            pdata = presp.json() if presp.text else {}
            feats = pdata.get("features") or []
            if feats:
                props = feats[0].get("properties") or {}
                parts = [
                    props.get("name"),
                    props.get("street"),
                    props.get("housenumber"),
                    props.get("city") or props.get("town") or props.get("village"),
                    props.get("country")
                ]
                address = ", ".join([p for p in parts if p])
                if address:
                    self.in_vieta.setText(address)
                    QMessageBox.information(self, "Adrese iegūta", f"Adrese: {address}")
                    return

            # Fallback uz koordinātēm, ja abi servisi neiedeva adresi
            self.in_vieta.setText(f"{lat_s}, {lon_s}")
            QMessageBox.warning(self, "Brīdinājums", "Neizdevās iegūt adresi. Iestatītas koordinātes.")
        except Exception as e:
            QMessageBox.warning(self, "Tīkla kļūda",
                                f"Neizdevās iegūt adresi (tīkla problēma vai API kļūda): {e}")

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


    
    def _invalidate_preview_cache_and_refresh(self, *args, **kwargs):
        """Piespiedu priekšskatījuma atjaunošana, īpaši atsauces pielikumiem un citiem nestandarta UI datiem."""
        try:
            self.last_data_hash = None
            if hasattr(self, 'preview_cache') and isinstance(self.preview_cache, dict):
                self.preview_cache.clear()
        except Exception:
            pass
        self._update_preview()

    def _add_reference_doc(self):
        """Pievieno atsauces dokumentus (PDF/DOCX/XLSX), kas tiks pievienoti PDF beigās kā atvasinājumi."""
        files, _ = QFileDialog.getOpenFileNames(
            self,
            "Pievienot atsauces dokumentus",
            "",
            "Dokumenti (*.pdf *.docx *.doc *.xlsx *.xls *.odt *.ods *.pptx *.ppt);;Visi faili (*.*)"
        )
        if not files:
            return

        existing = set()
        for i in range(self.list_atsauces_faili.count()):
            it = self.list_atsauces_faili.item(i)
            existing.add(_reference_doc_path(it.data(Qt.UserRole)))

        for p in files:
            p = str(p or '').strip()
            if not p or p in existing:
                continue
            payload = _normalize_reference_doc_payload({"ceļš": p, "nosaukums": os.path.basename(p), "scale_pct": 100})
            it = QListWidgetItem(_reference_doc_display_text(payload))
            it.setData(Qt.UserRole, payload)
            self.list_atsauces_faili.addItem(it)

        self._invalidate_preview_cache_and_refresh()

    def _remove_reference_doc(self):
        it = self.list_atsauces_faili.currentItem()
        if not it:
            return
        row = self.list_atsauces_faili.row(it)
        self.list_atsauces_faili.takeItem(row)
        self._invalidate_preview_cache_and_refresh()

    def _set_reference_doc_scale(self):
        """Maina izvēlētā atsauces dokumenta mērogu procentos priekš pielikuma lapām."""
        it = self.list_atsauces_faili.currentItem()
        if not it:
            QMessageBox.information(self, "Mērogs", "Vispirms izvēlieties atsauces dokumentu sarakstā.")
            return

        payload = _normalize_reference_doc_payload(it.data(Qt.UserRole), fallback_name=it.text())
        current_scale = int(payload.get('scale_pct', 100) or 100)
        value, ok = QInputDialog.getInt(
            self,
            "Dokumenta mērogs",
            "Ievadiet pielikuma mērogu procentos (10–500):",
            current_scale,
            10,
            500,
            5,
        )
        if not ok:
            return

        payload['scale_pct'] = int(value)
        it.setData(Qt.UserRole, payload)
        it.setText(_reference_doc_display_text(payload))
        self._invalidate_preview_cache_and_refresh()


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
            # Sinhronizējam iekšējos datus ar ielādētajiem iestatījumiem
            self.data = d
            # Foto kolonna: lai UI vienmēr būtu iespējams pievienot foto
            try:
                self.data.show_item_photo_in_table = True
            except Exception:
                pass
            try:
                idx_doc = self.cmb_doc_tips.findData(getattr(d, 'doc_tips', 'akta'))
                self.cmb_doc_tips.setCurrentIndex(idx_doc if idx_doc >= 0 else 0)
                self.in_doc_nosaukums.setText((getattr(d, 'dokumenta_nosaukums', '') or _doc_type_title(getattr(d, 'doc_tips', 'akta'))))
                if getattr(d, 'apmaksas_termins', ''):
                    self.in_apmaksas_termins.setDate(datetime.strptime(d.apmaksas_termins, '%Y-%m-%d').date())
                if getattr(d, 'piegades_datums', ''):
                    self.in_piegades_datums.setDate(datetime.strptime(d.piegades_datums, '%Y-%m-%d').date())
            except Exception:
                pass
            self.in_akta_nr.setText(d.akta_nr)
            self.in_datums.setDate(datetime.strptime(d.datums, '%Y-%m-%d').date())
            self.in_vieta.setText(d.vieta)
            self.in_pas_nr.setText(d.pasūtījuma_nr)
            self.in_piezimes.setText(d.piezīmes)

            self.in_liguma_nr.setText(d.līguma_nr)
            self.ck_ieklaut_izpildes_terminu.setChecked(getattr(d, 'ieklaut_izpildes_terminu', bool(d.izpildes_termiņš) if getattr(d, 'izpildes_termiņš', '') else True))
            self.ck_ieklaut_pienemsanas_datumu.setChecked(getattr(d, 'ieklaut_pienemsanas_datumu', bool(d.pieņemšanas_datums) if getattr(d, 'pieņemšanas_datums', '') else True))
            self.ck_ieklaut_nodosanas_datumu.setChecked(getattr(d, 'ieklaut_nodosanas_datumu', bool(d.nodošanas_datums) if getattr(d, 'nodošanas_datums', '') else True))
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
            if hasattr(self, 'in_apdrosinasana_teksts'):
                self.in_apdrosinasana_teksts.setText(getattr(d, 'apdrošināšana_teksts', '') or '')
            self.in_papildu_nosacijumi.setText(d.papildu_nosacījumi)
            self.in_atsauces_dokumenti.setText(d.atsauces_dokumenti)
            self.list_atsauces_faili.clear()
            for ref in getattr(d, 'atsauces_dokumenti_faili', []) or []:
                try:
                    if isinstance(ref, dict):
                        p = ref.get('ceļš', '')
                        n = ref.get('nosaukums', '') or os.path.basename(p)
                    else:
                        p = getattr(ref, 'ceļš', '')
                        n = getattr(ref, 'nosaukums', '') or os.path.basename(p)
                    if p:
                        payload = _normalize_reference_doc_payload(ref, fallback_name=n)
                        it = QListWidgetItem(_reference_doc_display_text(payload))
                        it.setData(Qt.UserRole, payload)
                        self.list_atsauces_faili.addItem(it)
                except Exception:
                    pass
            self.cb_akta_statuss.setCurrentText(d.akta_statuss)
            self.in_valuta.setCurrentText(d.valūta)
            self.ck_elektroniskais_paraksts.setChecked(d.elektroniskais_paraksts)
            self.ck_radit_elektronisko_parakstu_tekstu.setChecked(d.radit_elektronisko_parakstu_tekstu) # JAUNA RINDAS

            # Puses / Rekvizīti
            if hasattr(self, 'pie_in'):
                self._persona_to_inputs(self.pie_in, asdict(d.pieņēmējs))
            if hasattr(self, 'nod_in'):
                self._persona_to_inputs(self.nod_in, asdict(d.nodevējs))
            if hasattr(self, 'rek_in'):
                self._persona_to_inputs(self.rek_in, asdict(getattr(d, 'rekviziti', Persona())))
            if hasattr(self, 'cb_pie_loma'):
                self.cb_pie_loma.setCurrentText(getattr(d, 'pieņēmēja_loma', 'Pieņēmējs'))
            if hasattr(self, 'cb_nod_loma'):
                self.cb_nod_loma.setCurrentText(getattr(d, 'nodevēja_loma', 'Iekārtas/u /pakalpojuma/u nodevējs'))
            self._update_party_role_titles()
            if hasattr(self, 'in_rekvizitu_virsraksts'):
                self.in_rekvizitu_virsraksts.setText(getattr(d, 'rekvizitu_virsraksts', 'Rekvizīti'))
            if hasattr(self, 'cb_party_mode'):
                idx_pm = self.cb_party_mode.findData(getattr(d, 'party_mode', 'puses'))
                if idx_pm >= 0:
                    self.cb_party_mode.setCurrentIndex(idx_pm)
            self._apply_party_mode_ui()


# Pozīcijas
            self.tab.setRowCount(0)

            # JAUNS: vienmēr būvējam pilno kolonnu komplektu (paslēpšana notiek ar setColumnHidden)
            try:
                self.data.poz_columns_config = _merge_poz_columns_config(getattr(d, "poz_columns_config", None))
            except Exception:
                self.data.poz_columns_config = _merge_poz_columns_config({})

            try:
                if hasattr(self, "ck_show_price_summary") and self.ck_show_price_summary:
                    self.ck_show_price_summary.setChecked(bool(getattr(d, "show_price_summary", True)))
            except Exception:
                pass

            base_headers = [
                ("apraksts", "Apraksts"),
                ("daudzums", "Daudzums"),
                ("vieniba", "Vienība"),
                ("cena", "Cena"),
                ("summa", "Summa"),
                ("serial", "Seriālais Nr."),
                ("warranty", "Garantija"),
                ("notes", "Piezīmes pozīcijai"),
            ]

            headers = []
            for key, default_title in base_headers:
                try:
                    headers.append(str(self.data.poz_columns_config.get(key, {}).get("title", default_title)))
                except Exception:
                    headers.append(default_title)

            # Pielāgotās kolonnas
            self.data.custom_columns = getattr(d, "custom_columns", []) or []
            for col in self.data.custom_columns:
                try:
                    if isinstance(col, dict) and "visible" not in col:
                        col["visible"] = True
                except Exception:
                    pass
                try:
                    headers.append(str(col.get("name", "")) if isinstance(col, dict) else "")
                except Exception:
                    headers.append("")

            # Foto vienmēr pēdējā
            try:
                headers.append(str(self.data.poz_columns_config.get("foto", {}).get("title", "Foto")))
            except Exception:
                headers.append("Foto")

            self.tab.setColumnCount(len(headers))
            self.tab.setHorizontalHeaderLabels(headers)

            for i in range(len(headers)):
                self.tab.horizontalHeader().setSectionResizeMode(i, QHeaderView.Interactive)

            # JAUNS: pielietojam saglabāto kolonnu secību (UI) un platumus, ja tādi ir
            try:
                self._poz_apply_stored_visual_order()
            except Exception:
                pass
            try:
                st_b64 = str(getattr(d, 'poz_header_state_b64', '') or '')
                if st_b64:
                    self.tab.horizontalHeader().restoreState(base64.b64decode(st_b64.encode('ascii')))
            except Exception:
                pass

            self._poz_apply_column_visibility()

            for p in d.pozīcijas:
                r = self.tab.rowCount();
                self.tab.insertRow(r)
                self.tab.setItem(r, 0, QTableWidgetItem(p.apraksts))
                self.tab.setItem(r, 1, QTableWidgetItem(str(p.daudzums)))
                self.tab.setItem(r, 2, QTableWidgetItem(p.vienība))
                self.tab.setItem(r, 3, QTableWidgetItem(str(p.cena)))
                self.tab.setItem(r, 4, QTableWidgetItem(formēt_naudu(p.summa)))
                # Set values for potentially hidden columns
                idx = self._poz_col_indices()

                # Pielāgotās kolonnas (ja ir dati)
                try:
                    for col_idx, coldef in enumerate(getattr(self.data, "custom_columns", []) or []):
                        ui_col = idx.get("custom_start", 8) + col_idx
                        if ui_col >= self.tab.columnCount():
                            continue
                        val = ""
                        if isinstance(coldef, dict):
                            data_list = coldef.get("data", [])
                            if isinstance(data_list, list) and r < len(data_list):
                                val = str(data_list[r])
                        self.tab.setItem(r, ui_col, QTableWidgetItem(val))
                except Exception:
                    pass

                # Foto (ja ieslēgts)
                if "foto" in idx and self.tab.columnCount() > idx["foto"]:
                    foto_path = getattr(p, "attēla_ceļš", "") or ""
                    itf = QTableWidgetItem(foto_path)
                    itf.setFlags(itf.flags() & ~Qt.ItemIsEditable)
                    self.tab.setItem(r, idx["foto"], itf)
                    self._ensure_photo_cell(r)
                    w = self.tab.cellWidget(r, idx["foto"])
                    if isinstance(w, QToolButton):
                        if foto_path:
                            w.setText("Mainīt")
                            try:
                                ico = QIcon(foto_path)
                                if not ico.isNull():
                                    w.setIcon(ico)
                                    w.setIconSize(QSize(18, 18))
                            except Exception:
                                pass

                # Seriālais / Garantija / Piezīmes
                if "serial" in idx and self.tab.columnCount() > idx["serial"]:
                    self.tab.setItem(r, idx["serial"], QTableWidgetItem(p.seriālais_nr))
                if "warranty" in idx and self.tab.columnCount() > idx["warranty"]:
                    self.tab.setItem(r, idx["warranty"], QTableWidgetItem(p.garantija))
                if "notes" in idx and self.tab.columnCount() > idx["notes"]:
                    self.tab.setItem(r, idx["notes"], QTableWidgetItem(p.piezīmes_pozīcijai))



            # Attēli
            if hasattr(self, "photos_table") and self.photos_table is not None:
                self.photos_table.blockSignals(True)
                self.photos_table.setRowCount(0)
                for a in d.attēli:
                    self._photos_add_row(a.ceļš, a.paraksts)
                self.photos_table.blockSignals(False)
            else:
                # Back-compat
                if getattr(self, "img_list", None) is not None:
                    self.img_list.clear()
                    for a in d.attēli:
                        it = QListWidgetItem(os.path.basename(a.ceļš))
                        it.setData(Qt.UserRole, {"ceļš": a.ceļš, "paraksts": a.paraksts})
                        self.img_list.addItem(it)

            # Iestatījumi
            self.ck_pvn.setChecked(d.iekļaut_pvn)
            self.in_pvn.setValue(float(d.pvn_likme)) # Convert Decimal to float for QDoubleSpinBox
            self.ck_paraksti.setChecked(d.parakstu_rindas)
            if hasattr(self, 'cb_paraksta_rezims'):
                idx_sig = self.cb_paraksta_rezims.findData(getattr(d, 'paraksta_rezims', 'electronic' if getattr(d, 'elektroniskais_paraksts', False) else 'physical'))
                self.cb_paraksta_rezims.setCurrentIndex(idx_sig if idx_sig >= 0 else 0)
            if hasattr(self, 'in_paraksta_nav_teksts'):
                self.in_paraksta_nav_teksts.setText(getattr(d, 'paraksta_nav_teksts', 'Dokuments ir spēkā bez paraksta, ja tas sagatavots un apstiprināts atbilstoši pušu noteiktajai kārtībai.'))
            if hasattr(self, 'in_paraksta_vards_rekviziti'):
                self.in_paraksta_vards_rekviziti.setText(getattr(d, 'paraksta_vards_rekviziti', '') or '')
            if hasattr(self, 'in_paraksta_vards_pienemejs'):
                self.in_paraksta_vards_pienemejs.setText(getattr(d, 'paraksta_vards_pienemejs', '') or '')
            if hasattr(self, 'in_paraksta_vards_nodevejs'):
                self.in_paraksta_vards_nodevejs.setText(getattr(d, 'paraksta_vards_nodevejs', '') or '')
            if hasattr(self, 'in_papildu_parakstu_rindas'):
                self.in_papildu_parakstu_rindas.setPlainText(_extra_signature_rows_to_text(getattr(d, 'papildu_parakstu_rindas', []) or []))
            self._on_paraksta_rezims_changed()
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
            self.in_default_execution_days.setValue(int(getattr(d, 'default_execution_days', 5)))
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
            self.in_cover_page_title.setText(_effective_cover_page_title(d))
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
            # --- JAUNS: sinhronizē arī Pamata datu šifrēšanas lauku ---
            try:
                if hasattr(self, "ck_pdf_encrypt_basic") and hasattr(self, "in_pdf_password_basic"):
                    self.ck_pdf_encrypt_basic.setChecked(bool(getattr(d, "enable_pdf_encryption", False)))
                    self.in_pdf_password_basic.setText(str(getattr(d, "pdf_user_password", "") or ""))
            except Exception:
                pass

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
        self._undo_mgr.push_undo(self._snapshot_state('PROJECT_SAVE'))
        self._audit('PROJECT_SAVE', {})
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
            for filename in sorted(os.listdir(TEMPLATES_DIR)):
                if not filename.lower().endswith('.json'):
                    continue
                display_name = os.path.splitext(filename)[0]
                self.sablonu_list.addItem(display_name)

    def _resolve_template_path(self, template_display_name: str) -> Optional[str]:
        """
        Atrod šablona .json failu neatkarīgi no tā, kurā mapē tas glabājas.
        Meklē gan pašreizējo templates_dir, gan noklusējuma APP_DATA_DIR/AktaGenerators_Templates.
        """
        name = (template_display_name or "").strip()
        # noņemam marķierus no saraksta (ja tādi ir)
        for marker in ["[Aizsargāts]", "[Kļūda]"]:
            name = name.replace(marker, "").strip()

        # Kandidātu mapes (pirmā – lietotāja iestatītā)
        candidates = []
        if getattr(self.data, "templates_dir", ""):
            candidates.append(self.data.templates_dir)
        # noklusējuma mape
        candidates.append(os.path.join(APP_DATA_DIR, "AktaGenerators_Templates"))
        # drošībai: vecā/alternatīvā (ja kādreiz mainīts)
        candidates.append(os.path.join(os.path.expanduser("~"), "Documents", "AktaGenerators_Templates"))
        # noņemam dublikātus
        uniq = []
        for d in candidates:
            if d and d not in uniq:
                uniq.append(d)

        # mēģinām tiešu nosaukumu
        file_variants = [
            f"{name}.json",
            f"{drošs_faila_nosaukums(name)}.json",
        ]

        for d in uniq:
            try:
                os.makedirs(d, exist_ok=True)
            except Exception:
                pass
            for fn in file_variants:
                p = os.path.join(d, fn)
                if os.path.exists(p):
                    # ja atradām citā mapē nekā iestatīts, sinhronizējam
                    if getattr(self.data, "templates_dir", "") != d:
                        self.data.templates_dir = d
                    return p

        # Ja nav atrasts, mēģinām "fuzzy" – ignorējam atstarpes/_/- un reģistru
        def norm(s: str) -> str:
            return re.sub(r"[\s_\-]+", "", s).lower()

        wanted = norm(name)
        for d in uniq:
            if not d or not os.path.exists(d):
                continue
            for fn in os.listdir(d):
                if not fn.lower().endswith(".json"):
                    continue
                base = os.path.splitext(fn)[0]
                if norm(base) == wanted:
                    p = os.path.join(d, fn)
                    if getattr(self.data, "templates_dir", "") != d:
                        self.data.templates_dir = d
                    return p
        return None


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
            if clean_name != "Testa dati (piemērs)":  # Neļaujam dzēst iebūvēto šablonu
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
        self.sablonu_list.addItem("Testa dati (piemērs)")

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

        if template_name == "Testa dati (piemērs)":
            QMessageBox.information(self, "Mainīt paroli",
                                    "Iebūvētajam šablonam 'Testa dati (piemērs)' nevar mainīt paroli.")
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

        if template_name == "Testa dati (piemērs)":
            QMessageBox.information(self, "Noņemt paroli",
                                    "Iebūvētajam šablonam 'Testa dati (piemērs)' nav paroles, ko noņemt.")
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
        self._undo_mgr.push_undo(self._snapshot_state('PROJECT_LOAD'))
        self._audit('PROJECT_LOAD', {})
        path = _coerce_path(path)
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
                    doc_tips=data.get('doc_tips', 'akta'),
                    dokumenta_nosaukums=data.get('dokumenta_nosaukums', _doc_type_title(data.get('doc_tips', 'akta'))),
                    apmaksas_termins=data.get('apmaksas_termins', ''),
                    piegades_datums=data.get('piegades_datums', ''),
                    party_mode=data.get('party_mode', DOCUMENT_TYPE_PRESETS.get(data.get('doc_tips', 'akta'), {}).get('party_mode', 'puses')),
                    rekvizitu_virsraksts=data.get('rekvizitu_virsraksts', 'Rekvizīti'),
                    pieņēmēja_loma=data.get('pieņēmēja_loma', 'Pieņēmējs'),
                    nodevēja_loma=data.get('nodevēja_loma', 'Iekārtas/u /pakalpojuma/u nodevējs'),
                    pieņēmējs=Persona(
                        nosaukums=data.get('pieņēmējs', {}).get('nosaukums', ''),
                        reģ_nr=data.get('pieņēmējs', {}).get('reģ_nr', ''),
                        adrese=data.get('pieņēmējs', {}).get('adrese', ''),
                        kontaktpersona=data.get('pieņēmējs', {}).get('kontaktpersona', ''),
                        amats=data.get('pieņēmējs', {}).get('amats', ''),
                        pilnvaras_pamats=data.get('pieņēmējs', {}).get('pilnvaras_pamats', ''),
                        tālrunis=data.get('pieņēmējs', {}).get('tālrunis', ''),
                        epasts=data.get('pieņēmējs', {}).get('epasts', ''),
                        web_lapa=data.get('pieņēmējs', {}).get('web_lapa', ''),
                        bankas_konts=data.get('pieņēmējs', {}).get('bankas_konts', ''),
                        juridiskais_statuss=data.get('pieņēmējs', {}).get('juridiskais_statuss', '')
                    ),
                    nodevējs=Persona(
                        nosaukums=data.get('nodevējs', {}).get('nosaukums', ''),
                        reģ_nr=data.get('nodevējs', {}).get('reģ_nr', ''),
                        adrese=data.get('nodevējs', {}).get('adrese', ''),
                        kontaktpersona=data.get('nodevējs', {}).get('kontaktpersona', ''),
                        amats=data.get('nodevējs', {}).get('amats', ''),
                        pilnvaras_pamats=data.get('nodevējs', {}).get('pilnvaras_pamats', ''),
                        tālrunis=data.get('nodevējs', {}).get('tālrunis', ''),
                        epasts=data.get('nodevējs', {}).get('epasts', ''),
                        web_lapa=data.get('nodevējs', {}).get('web_lapa', ''),
                        bankas_konts=data.get('nodevējs', {}).get('bankas_konts', ''),
                        juridiskais_statuss=data.get('nodevējs', {}).get('juridiskais_statuss', '')
                    ),
                    rekviziti=Persona(
                        nosaukums=data.get('rekviziti', {}).get('nosaukums', ''),
                        reģ_nr=data.get('rekviziti', {}).get('reģ_nr', ''),
                        adrese=data.get('rekviziti', {}).get('adrese', ''),
                        kontaktpersona=data.get('rekviziti', {}).get('kontaktpersona', ''),
                        amats=data.get('rekviziti', {}).get('amats', ''),
                        pilnvaras_pamats=data.get('rekviziti', {}).get('pilnvaras_pamats', ''),
                        tālrunis=data.get('rekviziti', {}).get('tālrunis', ''),
                        epasts=data.get('rekviziti', {}).get('epasts', ''),
                        web_lapa=data.get('rekviziti', {}).get('web_lapa', ''),
                        bankas_konts=data.get('rekviziti', {}).get('bankas_konts', ''),
                        juridiskais_statuss=data.get('rekviziti', {}).get('juridiskais_statuss', '')
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
                    paraksta_rezims=data.get('paraksta_rezims', 'electronic' if get_bool(data, 'elektroniskais_paraksts', False) else 'physical'),
                    paraksta_nav_teksts=data.get('paraksta_nav_teksts', 'Dokuments ir spēkā bez paraksta, ja tas sagatavots un apstiprināts atbilstoši pušu noteiktajai kārtībai.'),
                    paraksta_vards_rekviziti=data.get('paraksta_vards_rekviziti', ''),
                    paraksta_vards_pienemejs=data.get('paraksta_vards_pienemejs', ''),
                    paraksta_vards_nodevejs=data.get('paraksta_vards_nodevejs', ''),
                    papildu_parakstu_rindas=data.get('papildu_parakstu_rindas', []) if isinstance(data.get('papildu_parakstu_rindas', []), list) else [],
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
                        apdrošināšana_teksts=data.get('apdrošināšana_teksts', ''),
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
            self._undo_mgr.push_undo(self._snapshot_state('TEMPLATE_LOAD'))
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
        # JAUNS: noklusējumam saglabājam pielāgoto kolonnu definīcijas, bet notīram rindu datus
        try:
            if isinstance(out.get('custom_columns'), list):
                for cc in out['custom_columns']:
                    if isinstance(cc, dict) and 'data' in cc:
                        cc['data'] = []
        except Exception:
            pass
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
                            web_lapa=data.get('nodevējs', {}).get('web_lapa', ''),
                            bankas_konts=data.get('nodevējs', {}).get('bankas_konts', ''),
                            juridiskais_statuss=data.get('nodevējs', {}).get('juridiskais_statuss', '')
                        ),
                        pozīcijas=[], # Default settings should not load positions
                        custom_columns=data.get('custom_columns', []) if isinstance(data.get('custom_columns', []), list) else [],
                        poz_columns_config=data.get('poz_columns_config', {}) if isinstance(data.get('poz_columns_config', {}), dict) else {},
                        show_price_summary=get_bool(data, 'show_price_summary', True),
                        poz_columns_visual_order=data.get('poz_columns_visual_order', []) if isinstance(data.get('poz_columns_visual_order', []), list) else [],
                        poz_header_state_b64=data.get('poz_header_state_b64', '') if isinstance(data.get('poz_header_state_b64', ''), str) else '',
                        attēli=[], # Default settings should not load images
                        piezīmes=data.get('piezīmes', ''), iekļaut_pvn=get_bool(data, 'iekļaut_pvn', False),
                        pvn_likme=get_decimal(data, 'pvn_likme', '21'),
                        parakstu_rindas=get_bool(data, 'parakstu_rindas', True),
                        paraksta_rezims=data.get('paraksta_rezims', 'electronic' if get_bool(data, 'elektroniskais_paraksts', False) else 'physical'),
                        paraksta_nav_teksts=data.get('paraksta_nav_teksts', 'Dokuments ir spēkā bez paraksta, ja tas sagatavots un apstiprināts atbilstoši pušu noteiktajai kārtībai.'),
                        paraksta_vards_rekviziti=data.get('paraksta_vards_rekviziti', ''),
                        paraksta_vards_pienemejs=data.get('paraksta_vards_pienemejs', ''),
                        paraksta_vards_nodevejs=data.get('paraksta_vards_nodevejs', ''),
                    papildu_parakstu_rindas=data.get('papildu_parakstu_rindas', []) if isinstance(data.get('papildu_parakstu_rindas', []), list) else [],
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
                        apdrošināšana_teksts=data.get('apdrošināšana_teksts', ''),
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
                    self._undo_mgr.push_undo(self._snapshot_state('TEMPLATE_LOAD'))
                    self.ieviest_datus(d)
                    self.data = d
                except Exception as e:
                    QMessageBox.warning(self, "Iestatījumu ielādes kļūda",
                                        f"Neizdevās ielādēt noklusējuma iestatījumus: {e}")

    # -----------------------
    # Audit + Undo/Redo API
    # -----------------------
    def _audit(self, event: str, details: dict | None = None):
        try:
            self._audit_logger.write(event, details or {}, user=getattr(self, "_current_user", ""))
            # ja audit tabs ir uzbūvēts, atjaunojam
            if hasattr(self, "_audit_table") and self._audit_table is not None:
                self._refresh_audit_table(limit=200)
        except Exception:
            pass

    def _snapshot_state(self, label: str = "") -> dict:
        """Saglabā stāvokli undo vajadzībām (adrešu grāmata + projekts)."""
        try:
            # adrešu grāmata
            ab = copy.deepcopy(getattr(self, "address_book", {}))

            # projekts (akta dati) – ja savākšana izgāžas, saglabājam tikai AB
            proj = None
            try:
                d = self.savākt_datus()
                proj = asdict(d)
            except Exception:
                proj = None

            # UI: izvēlētais address book entry
            sel = ""
            try:
                it = self.address_book_list.currentItem() if hasattr(self, "address_book_list") else None
                sel = it.text() if it else ""
            except Exception:
                sel = ""

            return {"label": label, "address_book": ab, "project": proj, "ab_selected": sel}
        except Exception:
            return {"label": label, "address_book": copy.deepcopy(getattr(self, "address_book", {}))}

    def _restore_state(self, state: dict):
        """Atjauno stāvokli no undo/redo."""
        if not state:
            return
        # adrešu grāmata
        if "address_book" in state:
            self.address_book = copy.deepcopy(state.get("address_book") or {})
            try:
                self._save_address_book()
            except Exception:
                pass
            try:
                self._update_address_book_list()
            except Exception:
                pass

        # projekts
        proj = state.get("project")
        if proj:
            try:
                # saglabājam/atjaunojam izmantojot esošo loģiku
                d = self._akta_dati_from_dict(proj)
                self._undo_mgr.push_undo(self._snapshot_state('TEMPLATE_LOAD'))
                self.ieviest_datus(d)
            except Exception:
                pass

        # atjaunojam selekciju
        try:
            sel = state.get("ab_selected") or ""
            if sel and hasattr(self, "address_book_list"):
                for i in range(self.address_book_list.count()):
                    if self.address_book_list.item(i).text() == sel:
                        self.address_book_list.setCurrentRow(i)
                        break
        except Exception:
            pass

    def undo_action(self):
        if not self._undo_mgr.can_undo():
            return
        # pašreizējo stāvokli ieliek redo
        try:
            self._undo_mgr.push_redo(self._snapshot_state("redo"))
            st = self._undo_mgr.pop_undo()
            self._restore_state(st)
            self._update_undo_redo_indicators()
            self._audit("UNDO", {"label": st.get("label", "")})
        except Exception:
            pass

    def redo_action(self):
        if not self._undo_mgr.can_redo():
            return
        try:
            self._undo_mgr.push_undo(self._snapshot_state("undo"))
            st = self._undo_mgr.pop_redo()
            self._restore_state(st)
            self._update_undo_redo_indicators()
            self._audit("REDO", {"label": st.get("label", "")})
        except Exception:
            pass

    def _akta_dati_from_dict(self, data: dict) -> AktaDati:
        """Atjauno AktaDati no dict (undo/redo/projekta ielāde)."""
        # Reuse conversion helpers already in file
        def get_decimal(dict_obj, key, default_val):
            val = dict_obj.get(key, default_val)
            return to_decimal(val)

        def get_bool(dict_obj, key, default_val):
            val = dict_obj.get(key, default_val)
            if isinstance(val, bool):
                return val
            if isinstance(val, str):
                return val.lower() in ("true", "1", "yes", "jā", "ja", "y")
            return bool(val)

        d = AktaDati(
            akta_nr=data.get("akta_nr", ""),
            datums=data.get("datums", ""),
            vieta=data.get("vieta", ""),
            pasūtījuma_nr=data.get("pasūtījuma_nr", ""),
            pieņēmējs=Persona(**(data.get("pieņēmējs") or {})),
            nodevējs=Persona(**(data.get("nodevējs") or {})),
            pozīcijas=[],
            attēli=[],
            piezīmes=data.get("piezīmes", ""),
            iekļaut_pvn=get_bool(data, "iekļaut_pvn", True),
            pvn_likme=get_decimal(data, "pvn_likme", "21"),
            piegādes_nosacījumi=data.get("piegādes_nosacījumi", ""),
            papildu_nosacījumi=data.get("papildu_nosacījumi", ""),
            atsauces_dokumenti=data.get("atsauces_dokumenti", ""),
            akta_statuss=data.get("akta_statuss", ""),
            cover_page_enabled=get_bool(data, "cover_page_enabled", True),
            cover_include_logo=get_bool(data, "cover_include_logo", True),
            cover_show_company_name=get_bool(data, "cover_show_company_name", True),
            cover_show_date=get_bool(data, "cover_show_date", True),
            cover_show_place=get_bool(data, "cover_show_place", True),
            cover_show_contacts=get_bool(data, "cover_show_contacts", True),
            cover_show_summary=get_bool(data, "cover_show_summary", True),
            cover_show_signatures=get_bool(data, "cover_show_signatures", True),
            cover_show_attachments=get_bool(data, "cover_show_attachments", True),
            cover_title_text=data.get("cover_title_text", "Pieņemšanas-Nodošanas Akts"),
            cover_subtitle_text=data.get("cover_subtitle_text", ""),
            cover_logo_max_width_mm=get_decimal(data, "cover_logo_max_width_mm", "30"),
            cover_logo_max_height_mm=get_decimal(data, "cover_logo_max_height_mm", "15"),
            digital_signature_enabled=get_bool(data, "digital_signature_enabled", False),
            digital_signature_text=data.get("digital_signature_text", "Dokuments parakstīts ar drošu elektronisko parakstu"),
            digital_signature_show_timestamp=get_bool(data, "digital_signature_show_timestamp", True),
            digital_signature_field_label=data.get("digital_signature_field_label", "Paraksts"),
            digital_signature_field_size_mm=get_decimal(data, "digital_signature_field_size_mm", "40"),
            digital_signature_field_position=data.get("digital_signature_field_position", "bottom_center"),
            qr_kods_enabled=get_bool(data, "qr_kods_enabled", True),
            qr_kods_ieklaut_pozicijas=get_bool(data, "qr_kods_ieklaut_pozicijas", True),
            qr_only_first_page=get_bool(data, "qr_only_first_page", False),
            qr_verification_url_enabled=get_bool(data, "qr_verification_url_enabled", False),
            qr_verification_base_url=data.get("qr_verification_base_url", ""),
            qr_kods_izmers_mm=get_decimal(data, "qr_kods_izmers_mm", "18"),
            qr_kods_tikai_pirma_lapa=get_bool(data, "qr_kods_tikai_pirma_lapa", True),
            qr_kods_url_mode=get_bool(data, "qr_kods_url_mode", False),
            qr_kods_url=data.get("qr_kods_url", ""),
        )

        # pozīcijas
        try:
            for p in (data.get("pozīcijas") or []):
                if isinstance(p, dict):
                    d.pozīcijas.append(Pozīcija(**p))
                else:
                    d.pozīcijas.append(p)
        except Exception:
            pass

        # attēli
        try:
            for a in (data.get("attēli") or []):
                if isinstance(a, dict):
                    d.attēli.append(Attēls(**a))
                else:
                    d.attēli.append(a)
        except Exception:
            pass

        return d



    # -----------------------
    # eParaksts integrācija (praktiska)
    # -----------------------
    def _get_eparaksts_app_path(self) -> str:
        """Atgriež saglabāto eParaksts EXE ceļu.
        Persistējas starp programmas palaišanām (QSettings), ar atpakaļsavietojamību uz settings.json.
        """
        try:
            # 1) QSettings (primārais)
            if not hasattr(self, "_qt_settings") or self._qt_settings is None:
                self._qt_settings = QSettings("AktaGenerators", "AktaGeneratorsApp")
            p = self._qt_settings.value("eparaksts/app_path", "", type=str)
            p = _coerce_path(p) or ""
            if p:
                return p

            # 2) Fallback uz veco settings.json, un migrējam uz QSettings
            if not hasattr(self, "_settings") or self._settings is None:
                self._settings = load_settings()
            p2 = _coerce_path((self._settings or {}).get("eparaksts_app_path", "")) or ""
            if p2:
                try:
                    self._qt_settings.setValue("eparaksts/app_path", p2)
                    self._qt_settings.sync()
                except Exception:
                    pass
            return p2
        except Exception:
            return ""


    def _set_eparaksts_app_path(self, p: str):
        """Saglabā eParaksts EXE ceļu gan QSettings (primāri), gan settings.json (fallback)."""
        try:
            p = _coerce_path(p) or ""
            if not hasattr(self, "_qt_settings") or self._qt_settings is None:
                self._qt_settings = QSettings("AktaGenerators", "AktaGeneratorsApp")
            self._qt_settings.setValue("eparaksts/app_path", p)
            self._qt_settings.sync()

            # atpakaļsavietojamība
            if not hasattr(self, "_settings") or self._settings is None:
                self._settings = load_settings()
            self._settings["eparaksts_app_path"] = p
            save_settings(self._settings)
        except Exception:
            pass


    def _open_settings_eparaksts(self):
        try:
            cur = self._get_eparaksts_app_path()
            msg = ("Norādi eParaksts parakstīšanas programmu (piem., eParakstītājs) EXE failu.\n"
                   "Ja neatstāsi, tiks atvērts PDF ar noklusēto programmu, un parakstīšanu veiksi manuāli.")
            fn, _ = QFileDialog.getOpenFileName(self, "Izvēlēties eParaksts programmu", cur or "", "Programmas (*.exe);;Visi faili (*.*)")
            if not fn:
                return
            self._set_eparaksts_app_path(fn)
            QMessageBox.information(self, "OK", "Saglabāts.")
            self._audit("SETTINGS_EPARAKSTS_APP", {"path": fn})
        except Exception as e:
            QMessageBox.warning(self, "Kļūda", str(e))

    def sign_file_with_eparaksts(self, file_path: str | None = None):
        """Atver failu eParaksts parakstīšanai (ārējā programma)."""
        try:
            p = _coerce_path(file_path or getattr(self, "_last_generated_pdf", ""))
            if not p or not os.path.exists(p):
                QMessageBox.warning(self, "Nav faila", "Nav atrasts PDF, ko parakstīt. Vispirms ģenerē PDF.")
                return

            app = self._get_eparaksts_app_path()

            # Ja lietotājs nav norādījis, mēģinām atrast biežākos ceļus Windows
            if (not app) and platform.system().lower().startswith("win"):
                candidates = [
                    os.path.join(os.environ.get("ProgramFiles", "C:\\Program Files"), "eParakstītājs", "eParakstītājs.exe"),
                    os.path.join(os.environ.get("ProgramFiles(x86)", "C:\\Program Files (x86)"), "eParakstītājs", "eParakstītājs.exe"),
                ]
                for c in candidates:
                    if os.path.exists(c):
                        app = c
                        break

            if app and os.path.exists(app):
                # sākam sekot parakstītajam failam tajā pašā mapē
                self._start_signed_file_watcher(p)
                # Daudzas e-paraksta programmas pieņem failu kā argumentu (ja nepieņem, vismaz atvērsies).
                subprocess.Popen([app, p], shell=False)
                self._audit("EPARAKSTS_OPEN_APP", {"app": app, "file": p})
            else:
                # Fallback: atver ar noklusēto PDF programmu
                try:
                    if platform.system().lower().startswith("win"):
                        os.startfile(p)  # type: ignore
                    else:
                        subprocess.Popen(["xdg-open", p])
                    self._audit("EPARAKSTS_OPEN_FALLBACK", {"file": p})
                except Exception as e:
                    QMessageBox.warning(self, "Kļūda", f"Neizdevās atvērt failu: {e}")
                    return

                QMessageBox.information(
                    self,
                    "Parakstīšana",
                    "PDF ir atvērts. Paraksti to ar eParaksts/eID rīku (ārējā programmā).\n"
                    "Ja vēlies automātiski atvērt eParakstītāju, iestatos norādi tā EXE ceļu: Fails → Iestatījumi → eParaksts."
                )
        except Exception as e:
            QMessageBox.warning(self, "Kļūda", str(e))



    def _open_folder(self, folder: str):
        try:
            folder = _coerce_path(folder)
            if not folder:
                return
            if platform.system().lower().startswith("win"):
                os.startfile(folder)  # type: ignore
            else:
                subprocess.Popen(["xdg-open", folder])
        except Exception:
            pass

    def _start_signed_file_watcher(self, original_pdf: str):
        """Pēc eParaksts atvēršanas sekojam mapē, vai parādās parakstītais fails (.asice/.edoc/.pdf)."""
        try:
            original_pdf = _coerce_path(original_pdf)
            if not original_pdf:
                return
            folder = os.path.dirname(original_pdf)
            stem = os.path.splitext(os.path.basename(original_pdf))[0]
            self._signed_watch_folder = folder
            self._signed_watch_stem = stem
            self._signed_watch_seen = set(os.listdir(folder)) if os.path.isdir(folder) else set()
            if hasattr(self, "_signed_watch_timer") and self._signed_watch_timer is not None:
                self._signed_watch_timer.stop()
            self._signed_watch_timer = QTimer(self)
            self._signed_watch_timer.setInterval(1200)
            self._signed_watch_timer.timeout.connect(self._poll_signed_file)
            self._signed_watch_timer.start()
        except Exception:
            pass

    def _poll_signed_file(self):
        try:
            folder = getattr(self, "_signed_watch_folder", "")
            stem = getattr(self, "_signed_watch_stem", "")
            if not folder or not os.path.isdir(folder):
                return
            now = set(os.listdir(folder))
            new_files = [f for f in (now - getattr(self, "_signed_watch_seen", set()))]
            self._signed_watch_seen = now

            # Meklējam tipiskos parakstītos failus
            candidates = []
            for f in new_files:
                fl = f.lower()
                if fl.endswith((".asice", ".edoc", ".asics", ".bdoc", ".pdf")):
                    # vēlams ar tādu pašu stem
                    if stem and f.startswith(stem):
                        candidates.append(f)
            if not candidates:
                # arī vecā sarakstā var parādīties ar aizturi — pārbaudam visu mapi
                for f in now:
                    fl = f.lower()
                    if stem and f.startswith(stem) and fl.endswith((".asice", ".edoc", ".asics", ".bdoc")):
                        candidates = [f]
                        break

            if candidates:
                signed = os.path.join(folder, candidates[0])
                self._last_signed_file = signed
                try:
                    if hasattr(self, "_signed_watch_timer") and self._signed_watch_timer is not None:
                        self._signed_watch_timer.stop()
                except Exception:
                    pass
                self._audit("EPARAKSTS_SIGNED_DETECTED", {"file": signed})
                QMessageBox.information(
                    self,
                    "Parakstīts",
                    f"Parakstītais fails atrasts:\n{signed}"
                )
        except Exception:
            pass

    def generate_and_sign_current(self):
        """1) Automātiski ģenerē PDF no šī brīža datiem 2) Atver eParaksts parakstīšanai 3) SeKo parakstītajam failam tajā pašā mapē."""
        try:
            d = self.savākt_datus()

            safe_akta_nr = drošs_faila_nosaukums(d.akta_nr) if d.akta_nr else "akts"
            safe_datums = d.datums.replace("-", "") if d.datums else datetime.now().strftime("%Y%m%d")
            default_name = f"{safe_akta_nr}_{safe_datums}.pdf"

            # izvēlamies saglabāšanas vietu
            default_dir = PROJECT_SAVE_DIR if os.path.isdir(PROJECT_SAVE_DIR) else os.getcwd()
            default_path = os.path.join(default_dir, default_name)

            pdf_path, _ = QFileDialog.getSaveFileName(self, "Saglabāt PDF (pirms parakstīšanas)", default_path, "PDF faili (*.pdf)")
            if not pdf_path:
                return

            # Undo checkpoint + audit
            self._undo_mgr.push_undo(self._snapshot_state("GENERATE_AND_SIGN"))
            self._audit("GENERATE_AND_SIGN_START", {"file": pdf_path})

            # ģenerējam PDF tieši uz izvēlēto vietu
            pdf_ceļš = ģenerēt_pdf(d, pdf_path, include_reference_docs=True, encrypt_pdf=True)
            pdf_ceļš = _coerce_path(pdf_ceļš) or _coerce_path(pdf_path)
            self._last_generated_pdf = pdf_ceļš
            self._audit("GENERATE_AND_SIGN_PDF_DONE", {"file": pdf_ceļš})

            # atveram eParaksts un sākam sekot rezultātam
            self.sign_file_with_eparaksts(pdf_ceļš)
            self._start_signed_file_watcher(pdf_ceļš)

            QMessageBox.information(
                self,
                "Parakstīšana",
                "PDF ir saglabāts un atvērts parakstīšanai.\n"
                "Parakstīto failu saglabā tajā pašā mapē.\n"
                "Programma automātiski mēģinās atrast parakstīto failu."
            )
        except Exception as e:
            QMessageBox.warning(self, "Kļūda", str(e))


    def _autosave_snapshot(self):
        try:
            d = self.savākt_datus()
            payload = asdict(d)
            payload['_autosave_saved_at'] = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
            data_str = json.dumps(payload, sort_keys=True, ensure_ascii=False, default=str)
            import hashlib
            data_hash = hashlib.md5(data_str.encode('utf-8')).hexdigest()
            if data_hash == self._last_autosave_hash:
                return
            safe_write_json(self._autosave_path, payload)
            self._last_autosave_hash = data_hash
        except Exception:
            pass

    def closeEvent(self, event):
        # Avāriju/negaidītas aizvēršanās gadījumam saglabājam pēdējo stāvokli
        self._autosave_snapshot()

        # Saglabājam vēsturi, adrešu grāmatu un teksta blokus vienmēr
        self._save_history()
        self._save_address_book()
        self.text_block_manager._save_text_blocks()

        # Noklusējuma iestatījumus saglabājam tikai, ja lietotājs to vēlas
        try:
            ans = QMessageBox.question(
                self,
                "Saglabāt noklusējuma iestatījumus?",
                "Vai saglabāt pašreizējos iestatījumus kā noklusējumu nākamajai palaišanai?",
                QMessageBox.Yes | QMessageBox.No | QMessageBox.Cancel,
                QMessageBox.No
            )
            if ans == QMessageBox.Cancel:
                event.ignore()
                return
            if ans == QMessageBox.Yes:
                self.saglabat_noklusejuma_iestatijumus()
        except Exception:
            # Ja dialoga parādīšana neizdodas, vienkārši neaiztiekam noklusējumu
            pass

        event.accept()

    def _update_preview(self):
        # Atlikt atjaunināšanu par 500 ms, lai izvairītos no biežas PDF ģenerēšanas
        self.preview_timer.start(900)

    def _do_update_preview(self):
        if not hasattr(self, 'preview_label'):
            return

        old_page = getattr(self, 'current_preview_page', 0)
        d = self.savākt_datus()
        # Izveidot hash no datiem, lai pārbaudītu, vai tie mainījušies
        import hashlib
        data_str = json.dumps(asdict(d), sort_keys=True, default=str)
        data_hash = hashlib.md5(data_str.encode('utf-8')).hexdigest()

        # Pārbaudīt kešatmiņu
        if data_hash == self.last_data_hash and self.preview_cache.get(data_hash):
            # Izmantot kešatmiņu
            cached = self.preview_cache[data_hash]
            self.preview_images = cached['images']
            self.current_preview_page = min(self.current_preview_page, len(self.preview_images) - 1)
            self._show_current_page()
            return

        # Smagā daļa (PDF ģenerācija + DOCX/XLSX konvertācija + pdf2image) notiek fonā,
        # lai programma nekad neuzkaras un neaizveras.
        self._requested_preview_hash = data_hash
        self._start_preview_worker(d, data_hash, old_page)


    
    def _is_preview_thread_running(self) -> bool:
        """Droši pārbauda, vai preview QThread vēl skrien (izvairās no 'already deleted' RuntimeError)."""
        if self._preview_thread is None:
            return False
        try:
            return self._preview_thread.isRunning()
        except RuntimeError:
            # C++ objekts jau izdzēsts (deleteLater), bet Python reference palika
            self._preview_thread = None
            self._preview_worker = None
            return False

    def _cleanup_preview_thread_refs(self):
        """Notīra atsauces uz preview thread/worker, kad thread ir beidzies."""
        self._preview_thread = None
        self._preview_worker = None

    def _start_preview_worker(self, d: 'AktaDati', data_hash: str, old_page: int):
        """Startē (vai ieplāno) priekšskatījuma ģenerēšanu fonā."""
        # Ja jau notiek ģenerēšana – nepārtraucam thread vardarbīgi; ieplānojam jaunāko pieprasījumu.
        if self._is_preview_thread_running():
            self._pending_preview_request = (d, data_hash, old_page)
            self.preview_label.setText("Ģenerē priekšskatījumu... (gaida rindā)")
            return

        # Ja vecais thread objekts ir palicis atsaucēs (bet vairs neskrien), droši notīram.
        try:
            if self._preview_thread is not None:
                try:
                    self._preview_thread.quit()
                    self._preview_thread.wait(50)
                except RuntimeError:
                    # jau izdzēsts
                    pass
        except Exception:
            pass

        self.preview_label.setText("Ģenerē priekšskatījumu...")

        thread = QThread(self)
        worker = _PreviewBuildWorker(d, data_hash, old_page)
        worker.moveToThread(thread)

        # Saglabājam atsauces, lai Qt tās negarbāž ārā
        self._preview_thread = thread
        self._preview_worker = worker

        thread.started.connect(worker.run)
        worker.finished.connect(self._on_preview_worker_finished)
        worker.failed.connect(self._on_preview_worker_failed)

        # Dzīves cikls
        worker.finished.connect(thread.quit)
        worker.failed.connect(thread.quit)

        # Kad beidzas, notīram atsauces PIRMS deleteLater (lai nerodas isRunning() uz izdzēsta C++ objekta)
        thread.finished.connect(self._cleanup_preview_thread_refs)
        thread.finished.connect(worker.deleteLater)
        thread.finished.connect(thread.deleteLater)

        thread.start()


    def _on_preview_worker_finished(self, data_hash: str, png_bytes_list: list, old_page: int):
        # Ja pa vidu bija jauns pieprasījums, bet šis nav jaunākais – ignorējam.
        if self._requested_preview_hash is not None and data_hash != self._requested_preview_hash:
            # Ja bija ieplānots jaunāks, palaidīsim to, kad thread beigsies (šeit jau beidzies)
            if self._pending_preview_request:
                d2, h2, op2 = self._pending_preview_request
                self._pending_preview_request = None
                self._requested_preview_hash = h2
                self._start_preview_worker(d2, h2, op2)
            return

        try:
            self.preview_images = []
            for b in png_bytes_list:
                pixmap = QPixmap()
                pixmap.loadFromData(b, 'PNG')
                self.preview_images.append(pixmap)

            # Saglabāt kešatmiņā
            self.preview_cache[data_hash] = {
                'images': self.preview_images.copy(),
                'page_count': len(self.preview_images)
            }
            self.last_data_hash = data_hash

            # Ierobežot kešatmiņas izmēru (saglabāt tikai pēdējos 5)
            if len(self.preview_cache) > 5:
                oldest_key = next(iter(self.preview_cache))
                del self.preview_cache[oldest_key]

            self.current_preview_page = min(max(old_page, 0), len(self.preview_images) - 1) if self.preview_images else 0
            self._show_current_page()

        finally:
            # Palaist pending, ja tāds ir
            if self._pending_preview_request:
                d2, h2, op2 = self._pending_preview_request
                self._pending_preview_request = None
                self._requested_preview_hash = h2
                self._start_preview_worker(d2, h2, op2)


    def _on_preview_worker_failed(self, data_hash: str, error_message: str):
        # Ja tas nav jaunākais pieprasījums, ignorējam.
        if self._requested_preview_hash is not None and data_hash != self._requested_preview_hash:
            if self._pending_preview_request:
                d2, h2, op2 = self._pending_preview_request
                self._pending_preview_request = None
                self._requested_preview_hash = h2
                self._start_preview_worker(d2, h2, op2)
            return

        self.preview_images = []
        self._show_current_page()
        self.preview_label.setText(f"Kļūda priekšskatījumā: {error_message}")
        QMessageBox.critical(
            self,
            "Priekšskatījuma kļūda",
            f"Neizdevās ģenerēt priekšskatījumu (ar atsauces pielikumiem).\n"
            f"Iespējams, trūkst LibreOffice (DOCX/XLSX konvertācijai) vai Poppler (PDF renderēšanai).\n"
            f"Kļūda: {error_message}"
        )

        if self._pending_preview_request:
            d2, h2, op2 = self._pending_preview_request
            self._pending_preview_request = None
            self._requested_preview_hash = h2
            self._start_preview_worker(d2, h2, op2)

    def _show_current_page(self):
        if self.preview_images and 0 <= self.current_preview_page < len(self.preview_images):
            pixmap = self.preview_images[self.current_preview_page]
            label_size = self.preview_scroll_area.viewport().size()
            scaled_pixmap = pixmap.scaled(label_size * self.zoom_factor, Qt.KeepAspectRatio,
                                          Qt.SmoothTransformation)
            self.preview_label.setPixmap(scaled_pixmap)
            # Lai QScrollArea varētu skrollēt (un pan ar peli strādātu), QLabel izmērs jāpielāgo pixmap izmēram
            self.preview_label.resize(scaled_pixmap.size())
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

    
    
    def _savakt_akta_datus(self) -> AktaDati:
        """Savāc datus no UI. Alias vecākiem izsaukumiem (ZIP u.c.)."""
        return self.savākt_datus()

    def _ģenerēt_pdf_failu(self, d: AktaDati, pdf_path: str) -> str:
        """Ģenerē galveno PDF (kā eksportā) un pievieno atsauces pielikumus (Atvasinājumi), ja tādi ir.
        Alias priekš vecākiem izsaukumiem ZIP funkcijā.
        """
        # 1) ģenerējam pamata PDF
        ģenerēt_pdf(d, pdf_path)

        # 2) pieliekam atsauces dokumentus (ja ir)
        try:
            # nododam to pašu fontu, lai latviešu diakritika vienmēr ir korekta
            fn = reģistrēt_fontu(getattr(d, "fonts_ceļš", ""))
            _append_reference_docs_to_pdf(pdf_path, d, font_name=fn)
        except Exception as e:
            print(f"Pielikumu pievienošanas kļūda: {e}")
        return pdf_path

    def ģenerēt_zip_dialogs(self):
        """Saglabā ZIP arhīvu: PDF + visi pielikumi (atsevišķi) + projekta JSON."""
        d0 = self.savākt_datus()
        safe_akta = drošs_faila_nosaukums(d0.akta_nr) if getattr(d0, "akta_nr", "") else "Akts"
        today = datetime.now().strftime("%Y%m%d")
        default_name = f"{safe_akta}_{today}.zip"
        default_path = os.path.join(DEFAULT_OUTPUT_DIR, default_name) if 'DEFAULT_OUTPUT_DIR' in globals() else default_name
        zip_path, _ = QFileDialog.getSaveFileName(self, "Saglabāt ZIP", default_path, "ZIP arhīvs (*.zip)")
        if not zip_path:
            return
        if not zip_path.lower().endswith(".zip"):
            zip_path += ".zip"

        try:
            # Sagatavojam pagaidu failus
            tmp_dir = tempfile.mkdtemp(prefix="akta_zip_")
            pdf_path = os.path.join(tmp_dir, "Akts.pdf")
            json_path = os.path.join(tmp_dir, "projekts.json")

            # Ģenerējam PDF uz pagaidu vietu
            d = self.savākt_datus()
            self._ģenerēt_pdf_failu(d, pdf_path)

            # Saglabājam JSON
            with open(json_path, "w", encoding="utf-8") as f:
                def _json_safe(o):
                    # Pārvērš Decimal un citus JSON-nederīgus tipus par drošu formu
                    from decimal import Decimal
                    if isinstance(o, Decimal):
                        return str(o)
                    if isinstance(o, dict):
                        return {k: _json_safe(v) for k, v in o.items()}
                    if isinstance(o, (list, tuple)):
                        return [_json_safe(v) for v in o]
                    return o

                json.dump(_json_safe(asdict(d)), f, ensure_ascii=False, indent=2)

            import zipfile
            with zipfile.ZipFile(zip_path, "w", compression=zipfile.ZIP_DEFLATED) as z:
                z.write(pdf_path, arcname="Akts.pdf")
                z.write(json_path, arcname="projekts.json")

                # Pielikumi (atsauces dokumenti) kā atsevišķi faili
                used = set()
                for i in range(self.list_atsauces_faili.count()):
                    it = self.list_atsauces_faili.item(i)
                    p = it.data(Qt.UserRole)
                    p = _coerce_path(p)
                    if not p or not os.path.exists(p):
                        continue
                    base = os.path.basename(p)
                    name = base
                    k = 2
                    while name.lower() in used:
                        root, ext = os.path.splitext(base)
                        name = f"{root}_{k}{ext}"
                        k += 1
                    used.add(name.lower())
                    z.write(p, arcname=os.path.join("pielikumi", name))

            # Ierakstām dokumentu vēsturē (kopējam PDF+JSON uz vēstures mapi)
            self._record_generated_document(pdf_path, json_path)

            QMessageBox.information(self, "Gatavs", "ZIP fails saglabāts veiksmīgi.")
        except Exception as e:
            QMessageBox.warning(self, "Kļūda", f"Neizdevās saglabāt ZIP: {e}")

    def ģenerēt_pdf_dialogs(self):
        self._undo_mgr.push_undo(self._snapshot_state('GENERATE_PDF'))
        self._audit('GENERATE_PDF', {})
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
        self._undo_mgr.push_undo(self._snapshot_state('GENERATE_DOCX'))
        self._audit('GENERATE_DOCX', {})
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

    # =======================
    # HOTFIX: PDF path safety
    # =======================

    _original_generate_pdf = ģenerēt_pdf

    def ģenerēt_pdf(
            akta_dati: AktaDati,
            pdf_ceļš: str | None = None,
            include_reference_docs: bool = True,
            encrypt_pdf: bool = True
    ) -> str:
        if isinstance(pdf_ceļš, dict):
            pdf_ceļš = pdf_ceļš.get("path") or pdf_ceļš.get("file") or ""
        elif isinstance(pdf_ceļš, (list, tuple)) and pdf_ceļš:
            pdf_ceļš = pdf_ceļš[0]
        elif pdf_ceļš and not isinstance(pdf_ceļš, (str, os.PathLike)):
            pdf_ceļš = ""

        result = _original_generate_pdf(
            akta_dati,
            os.fspath(pdf_ceļš) if pdf_ceļš else None,
            include_reference_docs=include_reference_docs,
            encrypt_pdf=encrypt_pdf
        )
        return os.fspath(result) if result else ""




def _inventory_collect_doc_list(self):
    docs = []
    try:
        for i in range(self.list_prod_docs.count()):
            it = self.list_prod_docs.item(i)
            payload = it.data(Qt.UserRole)
            if isinstance(payload, dict) and (payload.get('path') or payload.get('ceļš')):
                docs.append(dict(payload))
    except Exception:
        pass
    return docs


def _inventory_set_doc_list(self, docs):
    try:
        self.list_prod_docs.clear()
        for doc in (docs or []):
            if not isinstance(doc, dict):
                continue
            path = str(doc.get('path') or doc.get('ceļš') or '').strip()
            if not path:
                continue
            doc_type = str(doc.get('type') or doc.get('tips') or 'Dokuments').strip() or 'Dokuments'
            title = str(doc.get('name') or doc.get('nosaukums') or os.path.basename(path)).strip() or os.path.basename(path)
            item = QListWidgetItem(f'[{doc_type}] {title}')
            item.setData(Qt.UserRole, {'path': path, 'type': doc_type, 'name': title})
            self.list_prod_docs.addItem(item)
    except Exception:
        pass


def _inventory_add_product_docs(self):
    files, _ = QFileDialog.getOpenFileNames(self, 'Pievienot preces dokumentus', '', 'Dokumenti (*.pdf *.doc *.docx *.xls *.xlsx *.csv *.xml *.jpg *.jpeg *.png *.webp);;Visi faili (*.*)')
    if not files:
        return
    for p in files:
        title = os.path.basename(p)
        doc_type, ok = QInputDialog.getItem(self, 'Dokumenta tips', f'Izvēlies dokumenta tipu failam\n{title}', ['Pavadzīme', 'Iepirkuma dokuments', 'Muitas dokuments', 'Apmaksas dokuments', 'Sertifikāts', 'Cits'], 0, False)
        if not ok:
            doc_type = 'Cits'
        item = QListWidgetItem(f'[{doc_type}] {title}')
        item.setData(Qt.UserRole, {'path': p, 'type': doc_type, 'name': title})
        self.list_prod_docs.addItem(item)
    try:
        if files and not self.in_prod_docs_folder.text().strip():
            self.in_prod_docs_folder.setText(os.path.dirname(files[0]))
    except Exception:
        pass


def _inventory_remove_selected_product_doc(self):
    it = self.list_prod_docs.currentItem()
    if not it:
        return
    self.list_prod_docs.takeItem(self.list_prod_docs.row(it))


def _inventory_open_selected_product_doc(self):
    it = self.list_prod_docs.currentItem()
    if not it:
        return
    payload = it.data(Qt.UserRole) or {}
    path = payload.get('path') or payload.get('ceļš')
    if path and os.path.exists(path):
        try:
            QDesktopServices.openUrl(QUrl.fromLocalFile(path))
        except Exception:
            pass


def _ensure_reference_doc_added(self, doc_payload):
    try:
        if not hasattr(self, 'list_atsauces_faili'):
            return
        path = str((doc_payload or {}).get('path') or (doc_payload or {}).get('ceļš') or '').strip()
        if not path:
            return
        existing = set()
        for i in range(self.list_atsauces_faili.count()):
            it = self.list_atsauces_faili.item(i)
            existing.add(_reference_doc_path(it.data(Qt.UserRole)))
        if path in existing:
            return
        title = str((doc_payload or {}).get('name') or (doc_payload or {}).get('nosaukums') or os.path.basename(path)).strip() or os.path.basename(path)
        payload = _normalize_reference_doc_payload({"ceļš": path, "nosaukums": title, "scale_pct": (doc_payload or {}).get('scale_pct', 100)})
        item = QListWidgetItem(_reference_doc_display_text(payload))
        item.setData(Qt.UserRole, payload)
        self.list_atsauces_faili.addItem(item)
    except Exception:
        pass


def _choose_inventory_export_fields(self, field_options):
    item = getattr(self, '_inventory_current_item_for_export', None)
    opts = self._choose_inventory_transfer_options(item, 1.0, field_options)
    return list(opts.get('selected_fields') or []) if opts else []


def _choose_inventory_transfer_options(self, item, qty, field_options):
    if item is None:
        return None
    position_headers = []
    try:
        if hasattr(self, 'tab') and self.tab is not None:
            for c in range(self.tab.columnCount()):
                hi = self.tab.horizontalHeaderItem(c)
                position_headers.append((c, hi.text().strip() if hi else f'Kolonna {c+1}'))
    except Exception:
        position_headers = []
    dlg = InventoryTransferOptionsDialog(self, item, qty=qty, field_options=field_options, position_headers=position_headers)
    if dlg.exec() != QDialog.Accepted:
        return None
    return dlg.get_result()


def _pievienot_poziciju_no_noliktavas_modern(self, it: NoliktavasPrece, qty: float = 1.0, selected_fields=None, column_mapping=None, attach_docs=None):
    try:
        if not hasattr(self, 'tab') or self.tab is None:
            QMessageBox.warning(self, 'Noliktava', 'Nav atrasta pozīciju tabula.')
            return
        try:
            stock_qty = float(getattr(it, 'atlikums', 0.0) or 0.0)
            req_qty = float(qty or 0.0)
            if req_qty > stock_qty:
                QMessageBox.warning(self, 'Noliktavas brīdinājums', f'Pieprasītais daudzums ({req_qty:.3f}) pārsniedz noliktavas atlikumu ({stock_qty:.3f}). Prece tomēr tiks pievienota pozīcijām.')
        except Exception:
            pass
        idx = self._poz_col_indices()
        row = self.tab.rowCount()
        self.tab.insertRow(row)
        for c in range(self.tab.columnCount()):
            if self.tab.item(row, c) is None:
                self.tab.setItem(row, c, QTableWidgetItem(''))

        selected_fields = list(selected_fields or ['nosaukums'])
        column_mapping = dict(column_mapping or {})
        field_title_map = {
            'sku': 'SKU', 'svitrkods': 'Svītrkods', 'nosaukums': 'Nosaukums', 'kategorija': 'Kategorija', 'apakskategorija': 'Apakškategorija',
            'pavadzimes_numurs': 'Pavadzīmes Nr.', 'piegadatajs': 'Piegādātājs', 'razotajs': 'Ražotājs', 'partijas_numurs': 'Partijas Nr.',
            'serialais_numurs': 'Seriālais Nr.', 'noliktavas_nosaukums': 'Noliktava', 'atrasanas_vieta': 'Atrašanās vieta', 'iepirkuma_datums': 'Iepirkuma datums',
            'deriguma_termiņš': 'Derīguma termiņš', 'iepirkuma_valuta': 'Valūta', 'piezimes': 'Piezīmes', 'foto_path': 'Foto',
            'supplier_code': 'Piegādātāja kods', 'supplier_email': 'Piegādātāja e-pasts', 'supplier_phone': 'Piegādātāja tālrunis', 'hs_kods': 'HS kods', 'izcelsmes_valsts': 'Izcelsmes valsts'
        }
        bucket = {'apraksts': [], 'notes': [], 'serial': [], 'warranty': [], 'foto': []}
        direct_columns = {}
        for key in selected_fields:
            value = getattr(it, key, '') if hasattr(it, key) else ''
            if value in (None, ''):
                continue
            target = column_mapping.get(key, 'notes')
            if target == 'ignore':
                continue
            if isinstance(target, str) and target.startswith('col:'):
                try:
                    col_no = int(str(target).split(':', 1)[1])
                    direct_columns.setdefault(col_no, []).append(str(value))
                except Exception:
                    bucket['notes'].append(f"{field_title_map.get(key, key)}: {value}")
                continue
            if key == 'nosaukums' and target == 'apraksts':
                bucket['apraksts'].insert(0, str(value))
            elif target == 'apraksts':
                bucket['apraksts'].append(f"{field_title_map.get(key, key)}: {value}")
            elif target == 'notes':
                bucket['notes'].append(f"{field_title_map.get(key, key)}: {value}")
            elif target == 'serial':
                bucket['serial'].append(str(value))
            elif target == 'warranty':
                bucket['warranty'].append(str(value))
            elif target == 'foto':
                bucket['foto'].append(str(value))
        apr = '\n'.join([x for x in bucket['apraksts'] if x]).strip() or (it.nosaukums or it.sku or '')
        self.tab.item(row, idx['apraksts']).setText(apr)
        self.tab.item(row, idx['daudzums']).setText((f'{float(qty):.3f}').rstrip('0').rstrip('.'))
        self.tab.item(row, idx['vieniba']).setText((it.vieniba or getattr(self.data, 'default_unit', 'gab.')).strip() or 'gab.')
        self.tab.item(row, idx['cena']).setText(str(it.cena or ''))
        if idx.get('serial') is not None:
            self.tab.item(row, idx['serial']).setText(' | '.join(bucket['serial']))
        if idx.get('warranty') is not None:
            self.tab.item(row, idx['warranty']).setText(' | '.join(bucket['warranty']))
        if idx.get('notes') is not None:
            self.tab.item(row, idx['notes']).setText('\n'.join(bucket['notes']))
        for col_no, values in direct_columns.items():
            if 0 <= int(col_no) < self.tab.columnCount():
                self.tab.item(row, int(col_no)).setText(' | '.join([str(v) for v in values if str(v).strip()]))
        foto_path = (bucket['foto'][0] if bucket['foto'] else getattr(it, 'foto_path', '') or '')
        if idx.get('foto') is not None:
            self.tab.item(row, idx['foto']).setText(foto_path)
            try:
                self._ensure_photo_cell(row)
            except Exception:
                pass
        try:
            self._aprēķināt_pozīciju_summa(row)
        except Exception:
            pass
        for doc in (attach_docs or []):
            self._ensure_reference_doc_added(doc)
        try:
            self._atjaunot_numurus_un_kopsummas()
        except Exception:
            pass
        try:
            self._update_preview()
        except Exception:
            pass
    except Exception as e:
        QMessageBox.warning(self, 'Noliktava', f'Neizdevās pievienot pozīciju: {e}')


def _append_position_from_item_modern(self, it: NoliktavasPrece, qty: float):
    field_options = [
        ('nosaukums', 'Nosaukums / apraksts'), ('sku', 'SKU'), ('svitrkods', 'Svītrkods'), ('kategorija', 'Kategorija'), ('apakskategorija', 'Apakškategorija'),
        ('noliktavas_nosaukums', 'Noliktava'), ('atrasanas_vieta', 'Atrašanās vieta'), ('piegadatajs', 'Piegādātājs'), ('razotajs', 'Ražotājs'), ('partijas_numurs', 'Partijas Nr.'),
        ('serialais_numurs', 'Seriālais Nr.'), ('pavadzimes_numurs', 'Pavadzīmes Nr.'), ('iepirkuma_datums', 'Iepirkuma datums'), ('deriguma_termiņš', 'Derīguma termiņš'),
        ('statuss', 'Statuss'), ('cena', 'Cena'), ('pvn_likme', 'PVN %'), ('vieniba', 'Vienība'), ('foto_path', 'Foto'), ('piezimes', 'Piezīmes'),
        ('supplier_code', 'Piegādātāja kods'), ('supplier_email', 'Piegādātāja e-pasts'), ('supplier_phone', 'Piegādātāja tālrunis'), ('hs_kods', 'HS kods'), ('izcelsmes_valsts', 'Izcelsmes valsts')
    ]
    opts = self._choose_inventory_transfer_options(it, qty, field_options)
    if not opts:
        return
    self._pievienot_poziciju_no_noliktavas(it, qty=float(opts.get('qty', qty) or qty), selected_fields=opts.get('selected_fields') or ['nosaukums'], column_mapping=opts.get('column_mapping') or {}, attach_docs=opts.get('documents') or [])


def pievienot_pozicijas_no_noliktavas_modern(self):
    try:
        db = getattr(self, '_noliktava', None)
        if db is None:
            QMessageBox.warning(self, 'Noliktava nav pieejama', 'Noliktavas datubāze nav inicializēta.')
            return
        dlg = ProductPickerDialog(self, db)
        if dlg.exec() != QDialog.Accepted:
            return
        for it, qty in dlg.get_selection():
            self._append_position_from_item(it, qty)
        self._update_preview()
    except Exception as e:
        QMessageBox.warning(self, 'Kļūda', f'Neizdevās pievienot preces no noliktavas.\n\n{e}')


AktaLogs._inventory_collect_doc_list = _inventory_collect_doc_list
AktaLogs._inventory_set_doc_list = _inventory_set_doc_list
AktaLogs._inventory_add_product_docs = _inventory_add_product_docs
AktaLogs._inventory_remove_selected_product_doc = _inventory_remove_selected_product_doc
AktaLogs._inventory_open_selected_product_doc = _inventory_open_selected_product_doc
AktaLogs._ensure_reference_doc_added = _ensure_reference_doc_added
AktaLogs._choose_inventory_export_fields = _choose_inventory_export_fields
AktaLogs._choose_inventory_transfer_options = _choose_inventory_transfer_options
AktaLogs._pievienot_poziciju_no_noliktavas = _pievienot_poziciju_no_noliktavas_modern
AktaLogs._append_position_from_item = _append_position_from_item_modern
AktaLogs.pievienot_pozīcijas_no_noliktavas = pievienot_pozicijas_no_noliktavas_modern

# =======================
# HOTFIX: modern inventory import duplicate control + stock warnings
# =======================

def _norm_inventory_value(value):
    if value is None:
        return ''
    if isinstance(value, bool):
        return '1' if value else '0'
    if isinstance(value, (int, float)):
        try:
            return ('%.6f' % float(value)).rstrip('0').rstrip('.')
        except Exception:
            return str(value).strip()
    return str(value).strip()


def _inventory_import_signature(item):
    invoice = _norm_inventory_value(getattr(item, 'pavadzimes_numurs', ''))
    base = (
        _norm_inventory_value(getattr(item, 'sku', '')),
        _norm_inventory_value(getattr(item, 'nosaukums', '')),
        _norm_inventory_value(getattr(item, 'vieniba', '')),
        _norm_inventory_value(getattr(item, 'atlikums', 0.0)),
        _norm_inventory_value(getattr(item, 'cena', '')),
        _norm_inventory_value(getattr(item, 'piegadatajs', '')),
        _norm_inventory_value(getattr(item, 'partijas_numurs', '')),
        _norm_inventory_value(getattr(item, 'serialais_numurs', '')),
        _norm_inventory_value(getattr(item, 'noliktavas_nosaukums', '')),
        invoice,
    )
    if invoice:
        return ('invoice_key',) + base
    return ('full_row',) + _inventory_item_signature(item)


def _inventory_duplicate_reason(item):
    invoice = _norm_inventory_value(getattr(item, 'pavadzimes_numurs', ''))
    if invoice:
        return f"Sakrīt produkts un pavadzīmes Nr. {invoice}"
    return 'Identiska rinda jau eksistē importā vai noliktavā'

def _inventory_item_signature(item):
    doc_list = []
    for doc in list(getattr(item, 'dokumenti', []) or []):
        if isinstance(doc, dict):
            doc_list.append((
                _norm_inventory_value(doc.get('type') or doc.get('tips') or ''),
                _norm_inventory_value(doc.get('name') or doc.get('nosaukums') or ''),
                _norm_inventory_value(doc.get('path') or doc.get('ceļš') or ''),
            ))
        else:
            doc_list.append((_norm_inventory_value(doc), '', ''))
    doc_list = tuple(sorted(doc_list))
    return (
        _norm_inventory_value(getattr(item, 'sku', '')),
        _norm_inventory_value(getattr(item, 'nosaukums', '')),
        _norm_inventory_value(getattr(item, 'vieniba', '')),
        _norm_inventory_value(getattr(item, 'cena', '')),
        _norm_inventory_value(getattr(item, 'pvn_likme', '')),
        _norm_inventory_value(getattr(item, 'atlikums', 0.0)),
        _norm_inventory_value(getattr(item, 'minimalais_atlikums', 0.0)),
        _norm_inventory_value(getattr(item, 'svitrkods', '')),
        _norm_inventory_value(getattr(item, 'kategorija', '')),
        _norm_inventory_value(getattr(item, 'apakskategorija', '')),
        _norm_inventory_value(getattr(item, 'noliktavas_nosaukums', '')),
        _norm_inventory_value(getattr(item, 'atrasanas_vieta', '')),
        _norm_inventory_value(getattr(item, 'piegadatajs', '')),
        _norm_inventory_value(getattr(item, 'razotajs', '')),
        _norm_inventory_value(getattr(item, 'partijas_numurs', '')),
        _norm_inventory_value(getattr(item, 'serialais_numurs', '')),
        _norm_inventory_value(getattr(item, 'pavadzimes_numurs', '')),
        _norm_inventory_value(getattr(item, 'iepirkuma_datums', '')),
        _norm_inventory_value(getattr(item, 'deriguma_termiņš', '')),
        _norm_inventory_value(getattr(item, 'statuss', '')),
        _norm_inventory_value(getattr(item, 'iepirkuma_valuta', '')),
        _norm_inventory_value(getattr(item, 'piezimes', '')),
        _norm_inventory_value(getattr(item, 'foto_path', '')),
        doc_list,
    )


class InventoryImportDuplicatesDialog(QDialog):
    def __init__(self, parent, skipped_rows):
        super().__init__(parent)
        self.setWindowTitle('Identisku rindu pārbaude')
        self.resize(980, 520)
        self._rows = list(skipped_rows or [])

        layout = QVBoxLayout(self)
        info = QLabel('Šīs rindas netika importētas, jo programma atrada identiskus datus. Vari atzīmēt tās rindas, kuras tomēr apstrādāt/importēt manuāli.')
        info.setWordWrap(True)
        layout.addWidget(info)

        self.tbl = QTableWidget(len(self._rows), 6)
        self.tbl.setHorizontalHeaderLabels(['Iekļaut', 'Faila rinda', 'SKU', 'Nosaukums', 'Atlikums', 'Iemesls'])
        self.tbl.verticalHeader().setVisible(False)
        self.tbl.setAlternatingRowColors(True)

        for r, payload in enumerate(self._rows):
            chk = QCheckBox()
            chk.setChecked(False)
            self.tbl.setCellWidget(r, 0, chk)
            item = payload.get('item')
            self.tbl.setItem(r, 1, QTableWidgetItem(str(payload.get('row_no') or '')))
            self.tbl.setItem(r, 2, QTableWidgetItem(str(getattr(item, 'sku', '') or '')))
            self.tbl.setItem(r, 3, QTableWidgetItem(str(getattr(item, 'nosaukums', '') or '')))
            self.tbl.setItem(r, 4, QTableWidgetItem(_norm_inventory_value(getattr(item, 'atlikums', 0.0))))
            self.tbl.setItem(r, 5, QTableWidgetItem(str(payload.get('reason') or '')))

        self.tbl.horizontalHeader().setStretchLastSection(True)
        try:
            self.tbl.resizeColumnsToContents()
            self.tbl.setColumnWidth(0, 80)
            self.tbl.setColumnWidth(1, 90)
            self.tbl.setColumnWidth(2, 140)
            self.tbl.setColumnWidth(4, 90)
        except Exception:
            pass
        layout.addWidget(self.tbl, 1)

        row_btns = QHBoxLayout()
        btn_all = QPushButton('Atzīmēt visus')
        btn_none = QPushButton('Noņemt visus')
        btn_all.clicked.connect(lambda: [self.tbl.cellWidget(i, 0).setChecked(True) for i in range(self.tbl.rowCount()) if isinstance(self.tbl.cellWidget(i, 0), QCheckBox)])
        btn_none.clicked.connect(lambda: [self.tbl.cellWidget(i, 0).setChecked(False) for i in range(self.tbl.rowCount()) if isinstance(self.tbl.cellWidget(i, 0), QCheckBox)])
        row_btns.addWidget(btn_all)
        row_btns.addWidget(btn_none)
        row_btns.addStretch(1)
        layout.addLayout(row_btns)

        bb = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        bb.accepted.connect(self.accept)
        bb.rejected.connect(self.reject)
        layout.addWidget(bb)

    def selected_rows(self):
        out = []
        for i, payload in enumerate(self._rows):
            chk = self.tbl.cellWidget(i, 0)
            if isinstance(chk, QCheckBox) and chk.isChecked():
                out.append(payload)
        return out


# =======================
# HOTFIX: PDF path safety
# =======================

# Saglabājam oriģinālo PILNO funkciju un tikai normalizējam ceļu.
_original_generate_pdf = ģenerēt_pdf

def ģenerēt_pdf(
    akta_dati: AktaDati,
    pdf_ceļš: str | None = None,
    include_reference_docs: bool = True,
    encrypt_pdf: bool = True
) -> str:
    if isinstance(pdf_ceļš, dict):
        pdf_ceļš = pdf_ceļš.get("path") or pdf_ceļš.get("file") or ""
    elif isinstance(pdf_ceļš, (list, tuple)) and pdf_ceļš:
        pdf_ceļš = pdf_ceļš[0]
    elif pdf_ceļš and not isinstance(pdf_ceļš, (str, os.PathLike)):
        pdf_ceļš = ""

    result = _original_generate_pdf(
        akta_dati,
        os.fspath(pdf_ceļš) if pdf_ceļš else None,
        include_reference_docs=include_reference_docs,
        encrypt_pdf=encrypt_pdf
    )
    return os.fspath(result) if result else ""




def draw_first_page(canvas, doc):
    try:
        draw_header(canvas, doc, show_logo=False)
    except TypeError:
        draw_header(canvas, doc)

def draw_later_pages(canvas, doc):
    try:
        draw_header(canvas, doc, show_logo=False)
    except TypeError:
        pass



# =======================
# V52 HOTFIX: inventory UX, persistent transfer mapping, supplier/history autocomplete
# =======================
from PySide6.QtWidgets import QCompleter
from PySide6.QtCore import QStringListModel

INVENTORY_HISTORY_FILE = os.path.join(APP_DATA_DIR, "inventory_field_history.json")

def _inventory_field_history_load():
    try:
        if os.path.exists(INVENTORY_HISTORY_FILE):
            with open(INVENTORY_HISTORY_FILE, 'r', encoding='utf-8') as f:
                data = json.load(f) or {}
                if isinstance(data, dict):
                    return data
    except Exception:
        pass
    return {}


def _inventory_field_history_save(data: dict):
    try:
        os.makedirs(os.path.dirname(INVENTORY_HISTORY_FILE), exist_ok=True)
        with open(INVENTORY_HISTORY_FILE, 'w', encoding='utf-8') as f:
            json.dump(data or {}, f, ensure_ascii=False, indent=2)
    except Exception:
        pass


def _inventory_collect_item_snapshot(item):
    if item is None:
        return {}
    docs = []
    for d in list(getattr(item, 'dokumenti', []) or []):
        if isinstance(d, dict):
            docs.append(dict(d))
    return {
        'inventory_id': getattr(item, 'inventory_id', '') or '',
        'sku': getattr(item, 'sku', '') or '',
        'nosaukums': getattr(item, 'nosaukums', '') or '',
        'vieniba': getattr(item, 'vieniba', '') or '',
        'cena': getattr(item, 'cena', '') or '',
        'atlikums': getattr(item, 'atlikums', '') or '',
        'piegadatajs': getattr(item, 'piegadatajs', '') or '',
        'razotajs': getattr(item, 'razotajs', '') or '',
        'pavadzimes_numurs': getattr(item, 'pavadzimes_numurs', '') or '',
        'partijas_numurs': getattr(item, 'partijas_numurs', '') or '',
        'serialais_numurs': getattr(item, 'serialais_numurs', '') or '',
        'statuss': getattr(item, 'statuss', '') or '',
        'kategorija': getattr(item, 'kategorija', '') or '',
        'apakskategorija': getattr(item, 'apakskategorija', '') or '',
        'noliktavas_nosaukums': getattr(item, 'noliktavas_nosaukums', '') or '',
        'atrasanas_vieta': getattr(item, 'atrasanas_vieta', '') or '',
        'iepirkuma_datums': getattr(item, 'iepirkuma_datums', '') or '',
        'deriguma_termiņš': getattr(item, 'deriguma_termiņš', '') or '',
        'pvn_likme': getattr(item, 'pvn_likme', '') or '',
        'iepirkuma_valuta': getattr(item, 'iepirkuma_valuta', '') or '',
        'minimalais_atlikums': getattr(item, 'minimalais_atlikums', '') or '',
        'supplier_code': getattr(item, 'supplier_code', '') or '',
        'supplier_email': getattr(item, 'supplier_email', '') or '',
        'supplier_phone': getattr(item, 'supplier_phone', '') or '',
        'supplier_api_url': getattr(item, 'supplier_api_url', '') or '',
        'supplier_product_url': getattr(item, 'supplier_product_url', '') or '',
        'hs_kods': getattr(item, 'hs_kods', '') or '',
        'izcelsmes_valsts': getattr(item, 'izcelsmes_valsts', '') or '',
        'neto_svars': getattr(item, 'neto_svars', '') or '',
        'bruto_svars': getattr(item, 'bruto_svars', '') or '',
        'piezimes': getattr(item, 'piezimes', '') or '',
        'dokumenti': docs,
    }


def _inventory_details_html(snapshot: dict) -> str:
    rows = []
    order = [
        ('sku', 'SKU'), ('nosaukums', 'Nosaukums'), ('vieniba', 'Vienība'), ('cena', 'Cena'), ('atlikums', 'Atlikums'),
        ('piegadatajs', 'Piegādātājs'), ('razotajs', 'Ražotājs'), ('pavadzimes_numurs', 'Pavadzīmes Nr.'), ('partijas_numurs', 'Partijas Nr.'),
        ('serialais_numurs', 'Sērijas / seriālais Nr.'), ('statuss', 'Statuss'), ('kategorija', 'Kategorija'), ('apakskategorija', 'Apakškategorija'),
        ('noliktavas_nosaukums', 'Noliktava'), ('atrasanas_vieta', 'Atrašanās vieta'), ('iepirkuma_datums', 'Iepirkuma datums'),
        ('deriguma_termiņš', 'Derīguma termiņš'), ('pvn_likme', 'PVN %'), ('iepirkuma_valuta', 'Valūta'), ('minimalais_atlikums', 'Min. atlikums'),
        ('supplier_code', 'Piegādātāja kods'), ('supplier_email', 'Piegādātāja e-pasts'), ('supplier_phone', 'Piegādātāja tālrunis'),
        ('supplier_api_url', 'API / katalogs URL'), ('supplier_product_url', 'Produkta URL'), ('hs_kods', 'HS kods'), ('izcelsmes_valsts', 'Izcelsmes valsts'),
        ('neto_svars', 'Neto svars'), ('bruto_svars', 'Bruto svars'), ('piezimes', 'Piezīmes')
    ]
    for key, label in order:
        val = snapshot.get(key, '')
        if val not in ('', None):
            rows.append(f"<tr><td style='padding:4px 8px; font-weight:600; vertical-align:top'>{label}</td><td style='padding:4px 8px'>{str(val)}</td></tr>")
    docs = snapshot.get('dokumenti') or []
    doc_html = ''
    if docs:
        li = []
        for d in docs:
            name = d.get('name') or d.get('nosaukums') or os.path.basename(str(d.get('path') or d.get('ceļš') or ''))
            typ = d.get('type') or d.get('tips') or 'Dokuments'
            li.append(f"<li>[{typ}] {name}</li>")
        doc_html = "<h4>Dokumenti</h4><ul>" + ''.join(li) + "</ul>"
    return "<html><body><table cellspacing='0' cellpadding='0'>" + ''.join(rows) + "</table>" + doc_html + "</body></html>"


class ProductDetailsDialog(QDialog):
    def __init__(self, parent, item):
        super().__init__(parent)
        self.setWindowTitle("Preces pilna informācija")
        self.resize(760, 620)
        layout = QVBoxLayout(self)
        view = QTextBrowser()
        view.setOpenExternalLinks(True)
        view.setHtml(_inventory_details_html(_inventory_collect_item_snapshot(item)))
        layout.addWidget(view, 1)
        bb = QDialogButtonBox(QDialogButtonBox.Close)
        bb.rejected.connect(self.reject)
        bb.accepted.connect(self.accept)
        try:
            bb.button(QDialogButtonBox.Close).clicked.connect(self.accept)
        except Exception:
            pass
        layout.addWidget(bb)


class ProductPickerDialog(QDialog):
    def __init__(self, parent, noliktava_db: NoliktavaDB):
        super().__init__(parent)
        self.setWindowTitle("Pievienot preces no noliktavas")
        self.resize(1220, 700)
        self.db = noliktava_db
        self._result = []
        self._items_view = []

        root = QVBoxLayout(self)
        top = QHBoxLayout()
        self.in_search = QLineEdit()
        self.in_search.setPlaceholderText("Meklēt pēc SKU / nosaukuma / piegādātāja / pavadzīmes / partijas / piezīmēm…")
        self.in_search.textChanged.connect(self._refresh)
        top.addWidget(self.in_search, 1)
        self.cb_only_positive = QCheckBox("Tikai ar atlikumu > 0")
        self.cb_only_positive.setChecked(True)
        self.cb_only_positive.toggled.connect(self._refresh)
        top.addWidget(self.cb_only_positive)
        btn_clear = QPushButton("Notīrīt")
        btn_clear.clicked.connect(lambda: self.in_search.setText(""))
        top.addWidget(btn_clear)
        root.addLayout(top)

        splitter = QSplitter(Qt.Horizontal)
        self.tbl = QTableWidget()
        self.tbl.setColumnCount(11)
        self.tbl.setHorizontalHeaderLabels(["✓", "SKU", "Nosaukums", "Piegādātājs", "Pavadzīme", "Partija", "Noliktava", "Vienība", "Cena", "Atlikums", "Daudzums"])
        self.tbl.setSelectionBehavior(QTableWidget.SelectRows)
        self.tbl.setSelectionMode(QAbstractItemView.SingleSelection)
        self.tbl.setContextMenuPolicy(Qt.CustomContextMenu)
        self.tbl.customContextMenuRequested.connect(self._open_context_menu)
        self.tbl.itemSelectionChanged.connect(self._update_details_panel)
        self.tbl.itemDoubleClicked.connect(lambda *_: self._show_selected_details())
        splitter.addWidget(self.tbl)

        right = QWidget()
        rv = QVBoxLayout(right)
        rv.setContentsMargins(6, 6, 6, 6)
        self.lbl_summary = QLabel("")
        self.lbl_summary.setWordWrap(True)
        rv.addWidget(self.lbl_summary)
        self.details = QTextBrowser()
        self.details.setOpenExternalLinks(True)
        rv.addWidget(self.details, 1)
        quick = QHBoxLayout()
        btn_details = QPushButton("Pilna informācija")
        btn_details.clicked.connect(self._show_selected_details)
        btn_select_all = QPushButton("Atzīmēt visus")
        btn_select_all.clicked.connect(lambda: self._set_all_checks(True))
        btn_unselect_all = QPushButton("Noņemt visus")
        btn_unselect_all.clicked.connect(lambda: self._set_all_checks(False))
        quick.addWidget(btn_details)
        quick.addWidget(btn_select_all)
        quick.addWidget(btn_unselect_all)
        rv.addLayout(quick)
        splitter.addWidget(right)
        splitter.setStretchFactor(0, 4)
        splitter.setStretchFactor(1, 3)
        root.addWidget(splitter, 1)

        bottom = QHBoxLayout()
        self.lbl_footer = QLabel("")
        self.lbl_footer.setWordWrap(True)
        bottom.addWidget(self.lbl_footer, 1)
        btn_add = QPushButton("Pievienot izvēlētos")
        btn_add.clicked.connect(self._accept_selected)
        bottom.addWidget(btn_add)
        btn_cancel = QPushButton("Atcelt")
        btn_cancel.clicked.connect(self.reject)
        bottom.addWidget(btn_cancel)
        root.addLayout(bottom)

        self._refresh()

    def _set_all_checks(self, state: bool):
        for r in range(self.tbl.rowCount()):
            chk = self.tbl.cellWidget(r, 0)
            if isinstance(chk, QCheckBox):
                chk.setChecked(state)

    def _selected_item(self):
        row = self.tbl.currentRow()
        if row < 0 or row >= len(self._items_view):
            return None
        return self._items_view[row]

    def _show_selected_details(self):
        item = self._selected_item()
        if item is None:
            return
        dlg = ProductDetailsDialog(self, item)
        dlg.exec()

    def _copy_selected_field(self, field_name: str):
        item = self._selected_item()
        if item is None:
            return
        try:
            QApplication.clipboard().setText(str(getattr(item, field_name, '') or ''))
        except Exception:
            pass

    def _open_context_menu(self, pos):
        item = self._selected_item()
        if item is None:
            return
        menu = QMenu(self)
        act_info = menu.addAction("Rādīt visu informāciju")
        act_info.triggered.connect(self._show_selected_details)
        menu.addSeparator()
        for field_name, label in [('sku', 'Kopēt SKU'), ('nosaukums', 'Kopēt nosaukumu'), ('piegadatajs', 'Kopēt piegādātāju'), ('pavadzimes_numurs', 'Kopēt pavadzīmes Nr.'), ('serialais_numurs', 'Kopēt seriālo Nr.')]:
            act = menu.addAction(label)
            act.triggered.connect(lambda _=False, f=field_name: self._copy_selected_field(f))
        menu.addSeparator()
        act_pick = menu.addAction("Atzīmēt šo preci pievienošanai")
        def _toggle_current():
            r = self.tbl.currentRow()
            if r >= 0:
                chk = self.tbl.cellWidget(r, 0)
                if isinstance(chk, QCheckBox):
                    chk.setChecked(True)
        act_pick.triggered.connect(_toggle_current)
        menu.exec(self.tbl.viewport().mapToGlobal(pos))

    def _update_details_panel(self):
        item = self._selected_item()
        if item is None:
            self.details.setHtml("<i>Atlasiet noliktavas ierakstu, lai redzētu pilnu informāciju.</i>")
            return
        self.details.setHtml(_inventory_details_html(_inventory_collect_item_snapshot(item)))

    def _refresh(self):
        q = (self.in_search.text() or '').strip()
        items = self.db.find(q, include_zero=not self.cb_only_positive.isChecked()) if self.db else []
        self.tbl.setRowCount(len(items))
        for r, it in enumerate(items):
            chk = QCheckBox()
            self.tbl.setCellWidget(r, 0, chk)
            values = [
                it.sku or '', it.nosaukums or '', getattr(it, 'piegadatajs', '') or '', getattr(it, 'pavadzimes_numurs', '') or '',
                getattr(it, 'partijas_numurs', '') or '', getattr(it, 'noliktavas_nosaukums', '') or '', it.vieniba or '',
                str(it.cena or ''), str(it.atlikums if it.atlikums is not None else '')
            ]
            for c, val in enumerate(values, start=1):
                twi = QTableWidgetItem(str(val))
                if c == 1:
                    twi.setData(Qt.UserRole, getattr(it, 'inventory_id', '') or '')
                self.tbl.setItem(r, c, twi)
            qty = QDoubleSpinBox()
            qty.setDecimals(3)
            qty.setRange(0.0, 1e9)
            default_qty = 1.0
            try:
                stock = float(getattr(it, 'atlikums', 0.0) or 0.0)
                if stock > 0:
                    default_qty = min(1.0, stock) if stock < 1 else 1.0
            except Exception:
                pass
            qty.setValue(default_qty)
            self.tbl.setCellWidget(r, 10, qty)
        try:
            self.tbl.resizeColumnsToContents()
            self.tbl.setColumnWidth(0, 42)
            self.tbl.setColumnWidth(1, 120)
            self.tbl.setColumnWidth(2, 240)
            self.tbl.setColumnWidth(3, 160)
            self.tbl.setColumnWidth(4, 130)
            self.tbl.setColumnWidth(5, 120)
            self.tbl.setColumnWidth(6, 130)
            self.tbl.setColumnWidth(7, 80)
            self.tbl.setColumnWidth(8, 80)
            self.tbl.setColumnWidth(9, 80)
            self.tbl.setColumnWidth(10, 100)
        except Exception:
            pass
        self._items_view = items
        total = len(getattr(self.db, 'items', []) or []) if self.db else 0
        self.lbl_summary.setText(f"Atlasītajā logā redzami <b>{len(items)}</b> ieraksti. Kopā noliktavā ir <b>{total}</b> preču ieraksti.")
        self.lbl_footer.setText("Padoms: ar labo klikšķi vari redzēt pilnu informāciju un ātrās darbības. Dubultklikšķis atver pilno preces kartītes pārskatu.")
        if items and self.tbl.currentRow() < 0:
            self.tbl.selectRow(0)
        self._update_details_panel()

    def _accept_selected(self):
        out = []
        for r in range(self.tbl.rowCount()):
            chk = self.tbl.cellWidget(r, 0)
            qtyw = self.tbl.cellWidget(r, 10)
            if isinstance(chk, QCheckBox) and chk.isChecked():
                q = float(qtyw.value()) if isinstance(qtyw, QDoubleSpinBox) else 0.0
                if q > 0:
                    out.append((self._items_view[r], q))
        if not out:
            QMessageBox.information(self, "Nav izvēlēts", "Lūdzu atzīmē vismaz vienu preci un norādi daudzumu.")
            return
        self._result = out
        self.accept()

    def get_selection(self):
        return list(self._result or [])


class InventoryTransferOptionsDialog(QDialog):
    def __init__(self, parent, item: NoliktavasPrece, qty: float = 1.0, field_options=None, position_headers=None, saved_preset=None, total_inventory_items: int = 0):
        super().__init__(parent)
        self.setWindowTitle("Pārnese uz Pozīcijām")
        self.resize(980, 760)
        self._item = item
        self._field_options = list(field_options or [])
        self._position_headers = list(position_headers or [])
        self._saved_preset = dict(saved_preset or {})

        v = QVBoxLayout(self)
        info = QLabel(f"Prece: <b>{item.nosaukums or item.sku or ''}</b> &nbsp;&nbsp; SKU: {item.sku or '-'} &nbsp;&nbsp; Pieejams noliktavā: <b>{getattr(item, 'atlikums', 0) or 0}</b>")
        info.setWordWrap(True)
        v.addWidget(info)
        lbl_total = QLabel(f"Kopā noliktavā šobrīd reģistrēti <b>{int(total_inventory_items or 0)}</b> preču ieraksti.")
        lbl_total.setWordWrap(True)
        v.addWidget(lbl_total)

        qty_row = QHBoxLayout()
        qty_row.addWidget(QLabel("Daudzums pozīcijā:"))
        self.sp_qty = QDoubleSpinBox()
        self.sp_qty.setDecimals(3)
        self.sp_qty.setRange(0.001, 1e9)
        self.sp_qty.setValue(float(qty or 1.0))
        qty_row.addWidget(self.sp_qty)
        qty_row.addStretch(1)
        v.addLayout(qty_row)

        v.addWidget(QLabel("Izvēlies, kurās Pozīciju tabulas kolonnās ievietot Noliktavas laukus. Izvēle saglabājas arī pēc programmas aizvēršanas."))
        self.tbl = QTableWidget(len(self._field_options), 4)
        self.tbl.setHorizontalHeaderLabels(["Iekļaut", "Noliktavas lauks", "Vērtība", "Pozīciju kolonna"])
        self.tbl.verticalHeader().setVisible(False)
        self.tbl.setAlternatingRowColors(True)
        targets = [
            ("ignore", "Nepārsūtīt"),
            ("apraksts", "Apraksts"),
            ("notes", "Piezīmes pozīcijai"),
            ("serial", "Seriālais Nr."),
            ("warranty", "Garantija"),
            ("foto", "Foto"),
        ]
        for col_idx, header in self._position_headers:
            header = (header or '').strip()
            if header:
                targets.append((f'col:{col_idx}', f'Pozīciju kolonna: {header}'))
        default_target = {
            'nosaukums': 'apraksts', 'sku': 'notes', 'svitrkods': 'notes', 'kategorija': 'notes', 'apakskategorija': 'notes',
            'noliktavas_nosaukums': 'notes', 'atrasanas_vieta': 'notes', 'piegadatajs': 'notes', 'razotajs': 'notes',
            'partijas_numurs': 'notes', 'serialais_numurs': 'serial', 'pavadzimes_numurs': 'notes', 'iepirkuma_datums': 'notes',
            'deriguma_termiņš': 'warranty', 'statuss': 'notes', 'cena': 'ignore', 'pvn_likme': 'notes', 'vieniba': 'ignore',
            'foto_path': 'foto', 'piezimes': 'notes', 'supplier_code': 'notes', 'supplier_email': 'notes', 'supplier_phone': 'notes',
            'hs_kods': 'notes', 'izcelsmes_valsts': 'notes'
        }
        saved_fields = set(self._saved_preset.get('selected_fields') or [])
        saved_mapping = dict(self._saved_preset.get('column_mapping') or {})
        for r, (key, label) in enumerate(self._field_options):
            chk = QCheckBox()
            default_checked = key in ('nosaukums', 'serialais_numurs', 'foto_path', 'piezimes', 'sku')
            chk.setChecked(key in saved_fields if saved_fields else default_checked)
            self.tbl.setCellWidget(r, 0, chk)
            self.tbl.setItem(r, 1, QTableWidgetItem(label))
            val = getattr(item, key, '') if hasattr(item, key) else ''
            self.tbl.setItem(r, 2, QTableWidgetItem(str(val or '')))
            cb = QComboBox()
            for target_key, target_label in targets:
                cb.addItem(target_label, target_key)
            wanted = saved_mapping.get(key, default_target.get(key, 'notes'))
            ix = cb.findData(wanted)
            cb.setCurrentIndex(ix if ix >= 0 else 0)
            self.tbl.setCellWidget(r, 3, cb)
        self.tbl.horizontalHeader().setStretchLastSection(True)
        self.tbl.resizeColumnsToContents()
        v.addWidget(self.tbl, 1)

        tools = QHBoxLayout()
        btn_defaults = QPushButton("Atjaunot noklusējumu")
        btn_all = QPushButton("Iekļaut visus")
        btn_none = QPushButton("Noņemt visus")
        btn_defaults.clicked.connect(self._reset_defaults)
        btn_all.clicked.connect(lambda: [self.tbl.cellWidget(i, 0).setChecked(True) for i in range(self.tbl.rowCount()) if isinstance(self.tbl.cellWidget(i, 0), QCheckBox)])
        btn_none.clicked.connect(lambda: [self.tbl.cellWidget(i, 0).setChecked(False) for i in range(self.tbl.rowCount()) if isinstance(self.tbl.cellWidget(i, 0), QCheckBox)])
        tools.addWidget(btn_defaults)
        tools.addWidget(btn_all)
        tools.addWidget(btn_none)
        tools.addStretch(1)
        v.addLayout(tools)

        v.addWidget(QLabel("Dokumenti, kurus automātiski pievienot kā pielikumus šai precei:"))
        self.list_docs = QListWidget()
        docs = list(getattr(item, 'dokumenti', []) or [])
        for doc in docs:
            if not isinstance(doc, dict):
                continue
            path = str(doc.get('path') or doc.get('ceļš') or '').strip()
            if not path:
                continue
            doc_type = str(doc.get('type') or doc.get('tips') or 'Dokuments').strip() or 'Dokuments'
            title = str(doc.get('name') or doc.get('nosaukums') or os.path.basename(path)).strip() or os.path.basename(path)
            itemw = QListWidgetItem(f"[{doc_type}] {title}")
            itemw.setData(Qt.UserRole, {'path': path, 'type': doc_type, 'name': title})
            itemw.setFlags(itemw.flags() | Qt.ItemIsUserCheckable)
            itemw.setCheckState(Qt.Checked)
            self.list_docs.addItem(itemw)
        v.addWidget(self.list_docs, 1)

        docs_btns = QHBoxLayout()
        btn_all_docs = QPushButton("Visi dokumenti")
        btn_none_docs = QPushButton("Nevienu")
        btn_inv = QPushButton("Invertēt")
        for btn, state in ((btn_all_docs, Qt.Checked), (btn_none_docs, Qt.Unchecked)):
            btn.clicked.connect(lambda _=False, st=state: [self.list_docs.item(i).setCheckState(st) for i in range(self.list_docs.count())])
        btn_inv.clicked.connect(lambda: [self.list_docs.item(i).setCheckState(Qt.Unchecked if self.list_docs.item(i).checkState() == Qt.Checked else Qt.Checked) for i in range(self.list_docs.count())])
        docs_btns.addWidget(btn_all_docs)
        docs_btns.addWidget(btn_none_docs)
        docs_btns.addWidget(btn_inv)
        docs_btns.addStretch(1)
        v.addLayout(docs_btns)

        bb = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        bb.accepted.connect(self.accept)
        bb.rejected.connect(self.reject)
        v.addWidget(bb)

    def _reset_defaults(self):
        saved_fields = set()
        for r, (key, _label) in enumerate(self._field_options):
            chk = self.tbl.cellWidget(r, 0)
            cb = self.tbl.cellWidget(r, 3)
            if isinstance(chk, QCheckBox):
                chk.setChecked(key in ('nosaukums', 'serialais_numurs', 'foto_path', 'piezimes', 'sku'))
            if isinstance(cb, QComboBox):
                wanted = {'nosaukums': 'apraksts', 'serialais_numurs': 'serial', 'foto_path': 'foto'}.get(key, 'notes')
                ix = cb.findData(wanted)
                cb.setCurrentIndex(ix if ix >= 0 else 0)

    def get_result(self):
        mapping = {}
        selected_fields = []
        for r, (key, label) in enumerate(self._field_options):
            chk = self.tbl.cellWidget(r, 0)
            cb = self.tbl.cellWidget(r, 3)
            if isinstance(chk, QCheckBox) and chk.isChecked():
                selected_fields.append(key)
                if isinstance(cb, QComboBox):
                    mapping[key] = cb.currentData() or 'ignore'
            elif isinstance(cb, QComboBox):
                mapping[key] = cb.currentData() or 'ignore'
        docs = []
        for i in range(self.list_docs.count()):
            it = self.list_docs.item(i)
            if it.checkState() == Qt.Checked:
                docs.append(dict(it.data(Qt.UserRole) or {}))
        return {
            'qty': float(self.sp_qty.value()),
            'selected_fields': selected_fields,
            'column_mapping': mapping,
            'documents': docs,
        }


def _inventory_transfer_preset_get(self):
    try:
        settings = load_settings()
        return dict((settings or {}).get('inventory_transfer_preset_v2') or {})
    except Exception:
        return {}


def _inventory_transfer_preset_set(self, payload):
    try:
        settings = load_settings()
        settings['inventory_transfer_preset_v2'] = dict(payload or {})
        save_settings(settings)
    except Exception:
        pass


def _choose_inventory_transfer_options_v2(self, item, qty, field_options):
    if item is None:
        return None
    position_headers = []
    try:
        if hasattr(self, 'tab') and self.tab is not None:
            for c in range(self.tab.columnCount()):
                hi = self.tab.horizontalHeaderItem(c)
                position_headers.append((c, hi.text().strip() if hi else f'Kolonna {c+1}'))
    except Exception:
        position_headers = []
    saved_preset = self._inventory_transfer_preset_get() if hasattr(self, '_inventory_transfer_preset_get') else {}
    total_items = len(getattr(getattr(self, '_noliktava', None), 'items', []) or [])
    dlg = InventoryTransferOptionsDialog(self, item, qty=qty, field_options=field_options, position_headers=position_headers, saved_preset=saved_preset, total_inventory_items=total_items)
    if dlg.exec() != QDialog.Accepted:
        return None
    result = dlg.get_result()
    try:
        self._inventory_transfer_preset_set({'selected_fields': result.get('selected_fields') or [], 'column_mapping': result.get('column_mapping') or {}})
    except Exception:
        pass
    return result


def _inventory_all_history_candidates(self):
    db = getattr(self, '_noliktava', None)
    data = _inventory_field_history_load()
    result = {}
    field_map = {
        'warehouse': 'noliktavas_nosaukums', 'supplier': 'piegadatajs', 'manufacturer': 'razotajs', 'category': 'kategorija', 'subcategory': 'apakskategorija',
        'location': 'atrasanas_vieta', 'unit': 'vieniba', 'currency': 'iepirkuma_valuta', 'delivery_doc': 'pavadzimes_numurs', 'batch': 'partijas_numurs',
        'serial': 'serialais_numurs', 'supplier_code': 'supplier_code', 'supplier_email': 'supplier_email', 'supplier_phone': 'supplier_phone',
        'supplier_api_url': 'supplier_api_url', 'supplier_product_url': 'supplier_product_url', 'hs_code': 'hs_kods', 'origin_country': 'izcelsmes_valsts',
        'net_weight': 'neto_svars', 'gross_weight': 'bruto_svars', 'docs_folder': 'dokumentu_mape', 'name': 'nosaukums', 'sku': 'sku', 'barcode': 'svitrkods'
    }
    for history_key in field_map:
        result[history_key] = list(dict.fromkeys([x for x in (data.get(history_key) or []) if str(x).strip()]))
    for it in list(getattr(db, 'items', []) or []):
        for history_key, attr in field_map.items():
            val = getattr(it, attr, '') if hasattr(it, attr) else ''
            val = str(val or '').strip()
            if val and val not in result[history_key]:
                result[history_key].append(val)
    # rekvizīti kā papildu uzņēmumu bibliotēka
    try:
        for side in [('nodevejs_nosaukums', 'supplier'), ('pienemejs_nosaukums', 'supplier')]:
            v = str(getattr(getattr(self, 'data', None), side[0], '') or '').strip()
            if v and v not in result[side[1]]:
                result[side[1]].append(v)
    except Exception:
        pass
    return {k: sorted(v, key=lambda x: x.lower()) for k, v in result.items()}


def _inventory_store_current_field_history(self):
    try:
        data = _inventory_field_history_load()
        fields = {
            'warehouse': (self.in_prod_warehouse.currentText() if hasattr(self.in_prod_warehouse, 'currentText') else ''),
            'supplier': self.in_prod_supplier.text(),
            'manufacturer': self.in_prod_manufacturer.text(),
            'category': self.in_prod_category.text(),
            'subcategory': self.in_prod_subcategory.text(),
            'location': self.in_prod_location.text(),
            'unit': self.in_prod_unit.text(),
            'currency': self.in_prod_currency.currentText() if hasattr(self.in_prod_currency, 'currentText') else '',
            'delivery_doc': self.in_prod_delivery_doc.text(),
            'batch': self.in_prod_batch.text(),
            'serial': self.in_prod_serial.text(),
            'supplier_code': self.in_supplier_code.text(),
            'supplier_email': self.in_supplier_email.text(),
            'supplier_phone': self.in_supplier_phone.text(),
            'supplier_api_url': self.in_supplier_api_url.text(),
            'supplier_product_url': self.in_supplier_product_url.text(),
            'hs_code': self.in_prod_hs_code.text(),
            'origin_country': self.in_prod_origin_country.text(),
            'net_weight': self.in_prod_net_weight.text(),
            'gross_weight': self.in_prod_gross_weight.text(),
            'docs_folder': self.in_prod_docs_folder.text(),
            'name': self.in_prod_name.text(),
            'sku': self.in_prod_sku.text(),
            'barcode': self.in_prod_barcode.text(),
        }
        changed = False
        for key, val in fields.items():
            val = str(val or '').strip()
            if not val:
                continue
            arr = list(data.get(key) or [])
            if val not in arr:
                arr.append(val)
                data[key] = arr[-300:]
                changed = True
        if changed:
            _inventory_field_history_save(data)
    except Exception:
        pass


def _apply_completer(widget, values):
    try:
        vals = [str(v).strip() for v in (values or []) if str(v).strip()]
        vals = list(dict.fromkeys(vals))
        comp = QCompleter(vals, widget)
        comp.setCaseSensitivity(Qt.CaseInsensitive)
        comp.setFilterMode(Qt.MatchContains)
        if hasattr(widget, 'setCompleter'):
            widget.setCompleter(comp)
        elif isinstance(widget, QComboBox) and widget.isEditable() and widget.lineEdit() is not None:
            widget.lineEdit().setCompleter(comp)
    except Exception:
        pass


def _inventory_refresh_autocomplete_sources(self):
    try:
        vals = self._inventory_all_history_candidates()
        if hasattr(self, 'cb_saved_warehouses'):
            cur = self.cb_saved_warehouses.currentText()
            self.cb_saved_warehouses.blockSignals(True)
            self.cb_saved_warehouses.clear()
            self.cb_saved_warehouses.addItems(vals.get('warehouse', []))
            self.cb_saved_warehouses.setEditText(cur)
            self.cb_saved_warehouses.blockSignals(False)
        if hasattr(self, 'in_prod_warehouse'):
            cur = self.in_prod_warehouse.currentText()
            self.in_prod_warehouse.blockSignals(True)
            self.in_prod_warehouse.clear()
            self.in_prod_warehouse.addItems(vals.get('warehouse', []))
            self.in_prod_warehouse.setEditText(cur)
            self.in_prod_warehouse.blockSignals(False)
        mapping = {
            'supplier': getattr(self, 'in_prod_supplier', None),
            'manufacturer': getattr(self, 'in_prod_manufacturer', None),
            'category': getattr(self, 'in_prod_category', None),
            'subcategory': getattr(self, 'in_prod_subcategory', None),
            'location': getattr(self, 'in_prod_location', None),
            'unit': getattr(self, 'in_prod_unit', None),
            'delivery_doc': getattr(self, 'in_prod_delivery_doc', None),
            'batch': getattr(self, 'in_prod_batch', None),
            'serial': getattr(self, 'in_prod_serial', None),
            'supplier_code': getattr(self, 'in_supplier_code', None),
            'supplier_email': getattr(self, 'in_supplier_email', None),
            'supplier_phone': getattr(self, 'in_supplier_phone', None),
            'supplier_api_url': getattr(self, 'in_supplier_api_url', None),
            'supplier_product_url': getattr(self, 'in_supplier_product_url', None),
            'hs_code': getattr(self, 'in_prod_hs_code', None),
            'origin_country': getattr(self, 'in_prod_origin_country', None),
            'net_weight': getattr(self, 'in_prod_net_weight', None),
            'gross_weight': getattr(self, 'in_prod_gross_weight', None),
            'docs_folder': getattr(self, 'in_prod_docs_folder', None),
            'name': getattr(self, 'in_prod_name', None),
            'sku': getattr(self, 'in_prod_sku', None),
            'barcode': getattr(self, 'in_prod_barcode', None),
        }
        for key, widget in mapping.items():
            if widget is not None:
                _apply_completer(widget, vals.get(key, []))
        _apply_completer(getattr(self, 'in_prod_currency', None), vals.get('currency', []) + ['EUR', 'USD', 'GBP', 'SEK', 'NOK'])
        _apply_completer(getattr(self, 'cb_saved_warehouses', None), vals.get('warehouse', []))
        _apply_completer(getattr(self, 'in_prod_warehouse', None), vals.get('warehouse', []))
    except Exception:
        pass


def _inventory_bind_history_events(self):
    widgets = [
        getattr(self, 'in_prod_supplier', None), getattr(self, 'in_prod_manufacturer', None), getattr(self, 'in_prod_category', None), getattr(self, 'in_prod_subcategory', None),
        getattr(self, 'in_prod_location', None), getattr(self, 'in_prod_unit', None), getattr(self, 'in_prod_delivery_doc', None), getattr(self, 'in_prod_batch', None),
        getattr(self, 'in_prod_serial', None), getattr(self, 'in_supplier_code', None), getattr(self, 'in_supplier_email', None), getattr(self, 'in_supplier_phone', None),
        getattr(self, 'in_supplier_api_url', None), getattr(self, 'in_supplier_product_url', None), getattr(self, 'in_prod_hs_code', None), getattr(self, 'in_prod_origin_country', None),
        getattr(self, 'in_prod_net_weight', None), getattr(self, 'in_prod_gross_weight', None), getattr(self, 'in_prod_docs_folder', None), getattr(self, 'in_prod_name', None),
        getattr(self, 'in_prod_sku', None), getattr(self, 'in_prod_barcode', None)
    ]
    for w in widgets:
        try:
            w.editingFinished.connect(lambda s=self: s._inventory_store_current_field_history())
        except Exception:
            pass
    for cb in [getattr(self, 'in_prod_warehouse', None), getattr(self, 'cb_saved_warehouses', None), getattr(self, 'in_prod_currency', None)]:
        try:
            if cb is not None and hasattr(cb, 'currentTextChanged'):
                cb.currentTextChanged.connect(lambda _=None, s=self: s._inventory_store_current_field_history())
        except Exception:
            pass


_orig_inventory_build_tab = AktaLogs._būvēt_noliktava_tab

def _būvēt_noliktava_tab_v2(self):
    _orig_inventory_build_tab(self)
    try:
        if hasattr(self, '_noliktava') and self._noliktava is not None:
            setattr(self._noliktava, '_ui_owner', self)
        self._inventory_refresh_autocomplete_sources()
        self._inventory_bind_history_events()
    except Exception:
        pass


AktaLogs._būvēt_noliktava_tab = _būvēt_noliktava_tab_v2
AktaLogs._inventory_transfer_preset_get = _inventory_transfer_preset_get
AktaLogs._inventory_transfer_preset_set = _inventory_transfer_preset_set
AktaLogs._choose_inventory_transfer_options = _choose_inventory_transfer_options_v2
AktaLogs._inventory_all_history_candidates = _inventory_all_history_candidates
AktaLogs._inventory_store_current_field_history = _inventory_store_current_field_history
AktaLogs._inventory_refresh_autocomplete_sources = _inventory_refresh_autocomplete_sources
AktaLogs._inventory_bind_history_events = _inventory_bind_history_events


# refresh autocomplete/history after inventory DB changes
_OLD_NOLIKTAVA_SAVE = NoliktavaDB.save

def _noliktava_save_with_ui_refresh(self):
    res = _OLD_NOLIKTAVA_SAVE(self)
    owner = getattr(self, '_ui_owner', None)
    if owner is not None:
        try:
            owner._inventory_store_current_field_history()
            owner._inventory_refresh_autocomplete_sources()
        except Exception:
            pass
    return res

NoliktavaDB.save = _noliktava_save_with_ui_refresh


# ================================
# v54 enhancements: UI state, document numbers, history password, custom fields
# ================================
UI_STATE_FILE = os.path.join(SETTINGS_DIR, "ui_state.json")
DOC_NR_SETTINGS_FILE = os.path.join(SETTINGS_DIR, "doc_number_settings.json")
HISTORY_ACCESS_FILE = os.path.join(SETTINGS_DIR, "history_access.json")


def _v54_load_json(path, default):
    try:
        if os.path.exists(path):
            with open(path, 'r', encoding='utf-8') as f:
                return json.load(f)
    except Exception:
        pass
    return copy.deepcopy(default)


def _v54_save_json(path, payload):
    try:
        os.makedirs(os.path.dirname(path), exist_ok=True)
        with open(path, 'w', encoding='utf-8') as f:
            json.dump(payload, f, ensure_ascii=False, indent=2)
        return True
    except Exception:
        return False


def _v54_hash_password(password: str, salt: str) -> str:
    return hashlib.sha256((salt + '|' + (password or '')).encode('utf-8')).hexdigest()


def _v54_default_doc_nr_settings():
    return {
        'prefix': 'PP',
        'separator': '-',
        'include_doc_type': False,
        'doc_type_map': {'akta': 'AKT', 'pavadzime': 'PAV', 'rekins': 'REK'},
        'include_year': True,
        'include_month': False,
        'include_day': False,
        'sequence_padding': 4,
        'suffix': '',
        'uppercase': True,
        'year_reset': True,
    }


def _v54_load_doc_nr_settings(self):
    data = _v54_load_json(DOC_NR_SETTINGS_FILE, _v54_default_doc_nr_settings())
    if not isinstance(data, dict):
        data = _v54_default_doc_nr_settings()
    base = _v54_default_doc_nr_settings()
    base.update(data)
    self._doc_nr_settings = base
    return base


def _v54_collect_used_doc_numbers(self):
    used_ids = set()
    try:
        if os.path.isdir(PROJECT_SAVE_DIR):
            for filename in os.listdir(PROJECT_SAVE_DIR):
                if not filename.lower().endswith('.json'):
                    continue
                fp = os.path.join(PROJECT_SAVE_DIR, filename)
                try:
                    with open(fp, 'r', encoding='utf-8') as f:
                        project_data = json.load(f)
                    ak = (project_data or {}).get('akta_nr')
                    if ak:
                        used_ids.add(str(ak).strip())
                except Exception:
                    pass
    except Exception:
        pass

    try:
        hist = _v54_load_json(HISTORY_FILE, [])
        items = hist.get('items') if isinstance(hist, dict) else hist
        if isinstance(items, list):
            for item in items:
                if isinstance(item, dict):
                    ak = item.get('akta_nr') or item.get('doc_no')
                    if ak:
                        used_ids.add(str(ak).strip())
    except Exception:
        pass

    try:
        docs_dir = DOCUMENTS_DIR if 'DOCUMENTS_DIR' in globals() else ''
        if docs_dir and os.path.isdir(docs_dir):
            for filename in os.listdir(docs_dir):
                if not filename.lower().endswith('.json'):
                    continue
                fp = os.path.join(docs_dir, filename)
                try:
                    with open(fp, 'r', encoding='utf-8') as f:
                        doc = json.load(f)
                    ak = (doc or {}).get('akta_nr')
                    if ak:
                        used_ids.add(str(ak).strip())
                except Exception:
                    pass
    except Exception:
        pass
    return used_ids


def _v54_generate_akta_nr(self):
    cfg = _v54_load_doc_nr_settings(self)
    now = datetime.now()
    parts = []
    prefix = str(cfg.get('prefix', '') or '').strip()
    sep = str(cfg.get('separator', '-') or '')
    if prefix:
        parts.append(prefix)
    if cfg.get('include_doc_type'):
        try:
            doc_type = self.cmb_doc_tips.currentData() if hasattr(self, 'cmb_doc_tips') else 'akta'
        except Exception:
            doc_type = 'akta'
        doc_code = (cfg.get('doc_type_map') or {}).get(doc_type, str(doc_type or 'DOC').upper())
        if doc_code:
            parts.append(doc_code)
    if cfg.get('include_year', True):
        parts.append(now.strftime('%Y'))
    if cfg.get('include_month'):
        parts.append(now.strftime('%m'))
    if cfg.get('include_day'):
        parts.append(now.strftime('%d'))

    used_ids = _v54_collect_used_doc_numbers(self)
    pad = max(1, int(cfg.get('sequence_padding', 4) or 4))
    counter_data = _v54_load_json(AKTA_NR_COUNTER_FILE, {})
    scope_parts = []
    if cfg.get('year_reset', True):
        scope_parts.append(now.strftime('%Y'))
    if cfg.get('include_month'):
        scope_parts.append(now.strftime('%m'))
    if cfg.get('include_day'):
        scope_parts.append(now.strftime('%d'))
    scope_key = '|'.join(scope_parts) or 'global'
    n = int(counter_data.get(scope_key, 0) or 0) + 1
    while True:
        seq = str(n).zfill(pad)
        candidate_parts = list(parts) + [seq]
        candidate = sep.join([p for p in candidate_parts if p != ''])
        suffix = str(cfg.get('suffix', '') or '').strip()
        if suffix:
            candidate = f"{candidate}{sep if sep else ''}{suffix}"
        if cfg.get('uppercase', True):
            candidate = candidate.upper()
        if candidate not in used_ids:
            counter_data[scope_key] = n
            _v54_save_json(AKTA_NR_COUNTER_FILE, counter_data)
            self.in_akta_nr.setText(candidate)
            try:
                self._update_preview()
            except Exception:
                pass
            return candidate
        n += 1


def _v54_open_doc_nr_settings(self):
    cfg = _v54_load_doc_nr_settings(self)
    dlg = QDialog(self)
    dlg.setWindowTitle('Numura iestatījumi')
    dlg.resize(560, 420)
    lay = QVBoxLayout(dlg)
    form = QFormLayout()

    prefix = QLineEdit(str(cfg.get('prefix', 'PP')))
    separator = QLineEdit(str(cfg.get('separator', '-')))
    include_doc_type = QCheckBox('Iekļaut dokumenta tipu kodā')
    include_doc_type.setChecked(bool(cfg.get('include_doc_type')))
    include_year = QCheckBox('Iekļaut gadu')
    include_year.setChecked(bool(cfg.get('include_year', True)))
    include_month = QCheckBox('Iekļaut mēnesi')
    include_month.setChecked(bool(cfg.get('include_month')))
    include_day = QCheckBox('Iekļaut dienu')
    include_day.setChecked(bool(cfg.get('include_day')))
    year_reset = QCheckBox('Skaitītāju pārstartēt pa gadiem')
    year_reset.setChecked(bool(cfg.get('year_reset', True)))
    seq_padding = QSpinBox(); seq_padding.setRange(1, 12); seq_padding.setValue(int(cfg.get('sequence_padding', 4) or 4))
    suffix = QLineEdit(str(cfg.get('suffix', '')))
    uppercase = QCheckBox('Automātiski lielie burti')
    uppercase.setChecked(bool(cfg.get('uppercase', True)))
    doc_type_map = QTextEdit()
    dtm = cfg.get('doc_type_map') or {}
    doc_type_map.setPlainText('\n'.join([f"{k}={v}" for k, v in dtm.items()]))
    preview = QLabel('')
    preview.setWordWrap(True)

    form.addRow('Prefikss', prefix)
    form.addRow('Atdalītājs', separator)
    form.addRow(include_doc_type)
    form.addRow(include_year)
    form.addRow(include_month)
    form.addRow(include_day)
    form.addRow('Secības ciparu garums', seq_padding)
    form.addRow('Sufikss', suffix)
    form.addRow(uppercase)
    form.addRow(year_reset)
    form.addRow('Dokumenta tipu kodi (piem. akta=AKT)', doc_type_map)
    form.addRow('Priekšskatījums', preview)
    lay.addLayout(form)

    buttons = QDialogButtonBox(QDialogButtonBox.Save | QDialogButtonBox.Cancel)
    btn_reset = buttons.addButton('Atjaunot noklusējumu', QDialogButtonBox.ResetRole)
    btn_counter_reset = buttons.addButton('Pārstartēt skaitītāju', QDialogButtonBox.ActionRole)
    lay.addWidget(buttons)

    def build_preview():
        local = {
            'prefix': prefix.text().strip(),
            'separator': separator.text(),
            'include_doc_type': include_doc_type.isChecked(),
            'include_year': include_year.isChecked(),
            'include_month': include_month.isChecked(),
            'include_day': include_day.isChecked(),
            'sequence_padding': seq_padding.value(),
            'suffix': suffix.text().strip(),
            'uppercase': uppercase.isChecked(),
            'year_reset': year_reset.isChecked(),
            'doc_type_map': {},
        }
        for line in doc_type_map.toPlainText().splitlines():
            if '=' in line:
                k, v = line.split('=', 1)
                local['doc_type_map'][k.strip()] = v.strip()
        old = getattr(self, '_doc_nr_settings', None)
        self._doc_nr_settings = local
        try:
            preview.setText(_v54_generate_akta_nr(self) or '')
        finally:
            self._doc_nr_settings = local
    for w in [prefix, separator, suffix, doc_type_map]:
        try:
            w.textChanged.connect(lambda *args: build_preview())
        except Exception:
            pass
    for w in [include_doc_type, include_year, include_month, include_day, uppercase, year_reset]:
        w.toggled.connect(lambda *_: build_preview())
    seq_padding.valueChanged.connect(lambda *_: build_preview())
    build_preview()

    def on_reset_defaults():
        d = _v54_default_doc_nr_settings()
        prefix.setText(d['prefix']); separator.setText(d['separator']); include_doc_type.setChecked(d['include_doc_type'])
        include_year.setChecked(d['include_year']); include_month.setChecked(d['include_month']); include_day.setChecked(d['include_day'])
        seq_padding.setValue(d['sequence_padding']); suffix.setText(d['suffix']); uppercase.setChecked(d['uppercase'])
        year_reset.setChecked(d['year_reset'])
        doc_type_map.setPlainText('\n'.join([f"{k}={v}" for k, v in d['doc_type_map'].items()]))

    def on_reset_counter():
        reply = QMessageBox.question(self, 'Skaitītāja pārstartēšana', 'Vai pārstartēt dokumentu numuru skaitītāju?', QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
        if reply == QMessageBox.Yes:
            _v54_save_json(AKTA_NR_COUNTER_FILE, {})
            QMessageBox.information(self, 'Gatavs', 'Skaitītājs pārstartēts.')

    def on_save():
        new_cfg = {
            'prefix': prefix.text().strip(),
            'separator': separator.text(),
            'include_doc_type': include_doc_type.isChecked(),
            'include_year': include_year.isChecked(),
            'include_month': include_month.isChecked(),
            'include_day': include_day.isChecked(),
            'sequence_padding': seq_padding.value(),
            'suffix': suffix.text().strip(),
            'uppercase': uppercase.isChecked(),
            'year_reset': year_reset.isChecked(),
            'doc_type_map': {},
        }
        for line in doc_type_map.toPlainText().splitlines():
            if '=' in line:
                k, v = line.split('=', 1)
                if k.strip():
                    new_cfg['doc_type_map'][k.strip()] = v.strip()
        self._doc_nr_settings = new_cfg
        _v54_save_json(DOC_NR_SETTINGS_FILE, new_cfg)
        dlg.accept()

    btn_reset.clicked.connect(on_reset_defaults)
    btn_counter_reset.clicked.connect(on_reset_counter)
    buttons.accepted.connect(on_save)
    buttons.rejected.connect(dlg.reject)
    dlg.exec()


def _v54_add_doc_nr_button(self):
    try:
        if hasattr(self, 'btn_doc_nr_settings') and self.btn_doc_nr_settings is not None:
            return
        self.btn_doc_nr_settings = QPushButton('Numura iestatījumi')
        self.btn_doc_nr_settings.setToolTip('Mainīt dokumenta numura ģenerēšanas shēmu un unikālo numerāciju')
        self.btn_doc_nr_settings.clicked.connect(lambda: _v54_open_doc_nr_settings(self))
        parent = self.btn_generate_akta_nr.parentWidget()
        layout = parent.layout() if parent is not None else None
        if layout is not None:
            layout.addWidget(self.btn_doc_nr_settings)
    except Exception:
        pass


def _v54_load_ui_state(self):
    st = _v54_load_json(UI_STATE_FILE, {})
    self._ui_state_v54 = st if isinstance(st, dict) else {}
    dates = self._ui_state_v54.get('optional_dates', {}) if isinstance(self._ui_state_v54, dict) else {}
    try:
        mapping = [
            ('izpildes', self.ck_ieklaut_izpildes_terminu, self.in_izpildes_termins),
            ('pieņemšanas', self.ck_ieklaut_pienemsanas_datumu, self.in_pieņemšanas_datums),
            ('nodošanas', self.ck_ieklaut_nodosanas_datumu, self.in_nodošanas_datums),
        ]
        for key, ck, de in mapping:
            info = dates.get(key, {}) if isinstance(dates, dict) else {}
            if isinstance(info, dict):
                if 'checked' in info:
                    ck.setChecked(bool(info.get('checked')))
                if info.get('date'):
                    qd = QDate.fromString(str(info.get('date')), 'yyyy-MM-dd')
                    if qd.isValid():
                        de.setDate(qd)
    except Exception:
        pass


def _v54_save_ui_state(self):
    try:
        st = _v54_load_json(UI_STATE_FILE, {})
        if not isinstance(st, dict):
            st = {}
        st['optional_dates'] = {
            'izpildes': {'checked': bool(self.ck_ieklaut_izpildes_terminu.isChecked()), 'date': self.in_izpildes_termins.date().toString('yyyy-MM-dd')},
            'pieņemšanas': {'checked': bool(self.ck_ieklaut_pienemsanas_datumu.isChecked()), 'date': self.in_pieņemšanas_datums.date().toString('yyyy-MM-dd')},
            'nodošanas': {'checked': bool(self.ck_ieklaut_nodosanas_datumu.isChecked()), 'date': self.in_nodošanas_datums.date().toString('yyyy-MM-dd')},
        }
        st['custom_field_definitions'] = getattr(self, '_custom_field_definitions', {})
        st['custom_field_values'] = self._v54_collect_custom_field_values() if hasattr(self, '_v54_collect_custom_field_values') else {}
        _v54_save_json(UI_STATE_FILE, st)
    except Exception:
        pass


def _v54_bind_optional_date_state(self):
    for obj in [
        self.ck_ieklaut_izpildes_terminu, self.ck_ieklaut_pienemsanas_datumu, self.ck_ieklaut_nodosanas_datumu,
        self.in_izpildes_termins, self.in_pieņemšanas_datums, self.in_nodošanas_datums,
    ]:
        try:
            if hasattr(obj, 'stateChanged'):
                obj.stateChanged.connect(lambda *_: _v54_save_ui_state(self))
            if hasattr(obj, 'dateChanged'):
                obj.dateChanged.connect(lambda *_: _v54_save_ui_state(self))
        except Exception:
            pass


def _v54_custom_fields_default_defs():
    return {'pie': [], 'nod': [], 'rek': []}


def _v54_init_custom_fields(self):
    st = _v54_load_json(UI_STATE_FILE, {})
    defs = st.get('custom_field_definitions', {}) if isinstance(st, dict) else {}
    merged = _v54_custom_fields_default_defs()
    if isinstance(defs, dict):
        for k in merged:
            v = defs.get(k)
            if isinstance(v, list):
                merged[k] = [str(x) for x in v if str(x).strip()]
    self._custom_field_definitions = merged
    self._custom_field_widgets = {'pie': {}, 'nod': {}, 'rek': {}}
    self._custom_field_forms = {}

    def attach(group_box, key, title):
        try:
            parent_layout = group_box.layout()
            ctrl = QHBoxLayout()
            lbl = QLabel('Pielāgotie lauki')
            btn_add = QPushButton('Pievienot lauku')
            btn_manage = QPushButton('Pārvaldīt laukus')
            ctrl.addWidget(lbl)
            ctrl.addStretch(1)
            ctrl.addWidget(btn_add)
            ctrl.addWidget(btn_manage)
            parent_layout.addRow(ctrl)
            inner = QWidget()
            inner_form = QFormLayout(inner)
            inner_form.setContentsMargins(0, 0, 0, 0)
            parent_layout.addRow(inner)
            self._custom_field_forms[key] = inner_form
            btn_add.clicked.connect(lambda *_: _v54_add_custom_field(self, key))
            btn_manage.clicked.connect(lambda *_: _v54_manage_custom_fields(self, key, title))
            _v54_rebuild_custom_fields(self, key)
        except Exception:
            pass

    attach(self.grp_pieņēmējs, 'pie', 'Pieņēmējs')
    attach(self.grp_nodevējs, 'nod', 'Nodevējs')
    attach(self.grp_rekviziti, 'rek', 'Rekvizīti')

    values = st.get('custom_field_values', {}) if isinstance(st, dict) else {}
    if isinstance(values, dict):
        for sec, fields in values.items():
            if sec in self._custom_field_widgets and isinstance(fields, dict):
                for name, value in fields.items():
                    w = self._custom_field_widgets[sec].get(name)
                    if isinstance(w, QLineEdit):
                        w.setText(str(value or ''))


def _v54_rebuild_custom_fields(self, section):
    form = getattr(self, '_custom_field_forms', {}).get(section)
    if form is None:
        return
    while form.rowCount() > 0:
        form.removeRow(0)
    self._custom_field_widgets[section] = {}
    for field_name in self._custom_field_definitions.get(section, []):
        row = QWidget()
        row_l = QHBoxLayout(row)
        row_l.setContentsMargins(0, 0, 0, 0)
        edit = QLineEdit()
        edit.setPlaceholderText(field_name)
        edit.textChanged.connect(lambda *_: _v54_save_ui_state(self))
        btn_del = QToolButton(); btn_del.setText('✕'); btn_del.setToolTip('Dzēst lauku')
        btn_del.clicked.connect(lambda *_args, sec=section, fn=field_name: _v54_delete_custom_field(self, sec, fn))
        row_l.addWidget(edit, 1)
        row_l.addWidget(btn_del)
        form.addRow(field_name, row)
        self._custom_field_widgets[section][field_name] = edit
    _v54_save_ui_state(self)


def _v54_add_custom_field(self, section):
    name, ok = QInputDialog.getText(self, 'Pievienot pielāgoto lauku', 'Lauka nosaukums:')
    name = (name or '').strip()
    if not ok or not name:
        return
    if name in self._custom_field_definitions.get(section, []):
        QMessageBox.information(self, 'Info', 'Šāds lauks jau eksistē.')
        return
    self._custom_field_definitions.setdefault(section, []).append(name)
    _v54_rebuild_custom_fields(self, section)


def _v54_delete_custom_field(self, section, field_name):
    try:
        self._custom_field_definitions[section] = [x for x in self._custom_field_definitions.get(section, []) if x != field_name]
        _v54_rebuild_custom_fields(self, section)
    except Exception:
        pass


def _v54_manage_custom_fields(self, section, title):
    items = list(self._custom_field_definitions.get(section, []))
    if not items:
        QMessageBox.information(self, title, 'Pielāgoto lauku vēl nav.')
        return
    item, ok = QInputDialog.getItem(self, f'{title} lauki', 'Izvēlies lauku pārdēvēšanai vai dzēšanai:', items, 0, False)
    if not ok or not item:
        return
    action, ok2 = QInputDialog.getItem(self, 'Darbība', 'Ko darīt ar izvēlēto lauku?', ['Pārdēvēt', 'Dzēst'], 0, False)
    if not ok2:
        return
    if action == 'Dzēst':
        _v54_delete_custom_field(self, section, item)
        return
    new_name, ok3 = QInputDialog.getText(self, 'Pārdēvēt lauku', 'Jaunais nosaukums:', text=item)
    new_name = (new_name or '').strip()
    if not ok3 or not new_name or new_name == item:
        return
    defs = self._custom_field_definitions.get(section, [])
    defs = [new_name if x == item else x for x in defs]
    self._custom_field_definitions[section] = defs
    vals = self._v54_collect_custom_field_values().get(section, {})
    if item in vals:
        vals[new_name] = vals.pop(item)
    _v54_rebuild_custom_fields(self, section)
    w = self._custom_field_widgets.get(section, {}).get(new_name)
    if isinstance(w, QLineEdit):
        w.setText(vals.get(new_name, ''))


def _v54_collect_custom_field_values(self):
    out = {'pie': {}, 'nod': {}, 'rek': {}}
    for sec, mapping in getattr(self, '_custom_field_widgets', {}).items():
        for name, widget in mapping.items():
            try:
                out[sec][name] = widget.text().strip()
            except Exception:
                out[sec][name] = ''
    return out


def _v54_persona_custom_payload(self, section):
    return {
        'definitions': list(getattr(self, '_custom_field_definitions', {}).get(section, [])),
        'values': dict(self._v54_collect_custom_field_values().get(section, {}))
    }


def _v54_apply_custom_payload(self, section, payload):
    if not isinstance(payload, dict):
        return
    defs = payload.get('definitions')
    if isinstance(defs, list):
        self._custom_field_definitions[section] = [str(x) for x in defs if str(x).strip()]
        _v54_rebuild_custom_fields(self, section)
    vals = payload.get('values')
    if isinstance(vals, dict):
        for name, value in vals.items():
            w = self._custom_field_widgets.get(section, {}).get(name)
            if isinstance(w, QLineEdit):
                w.setText(str(value or ''))
    _v54_save_ui_state(self)


def _v54_patch_address_book_custom_fields():
    old_save = AktaLogs._save_persona_to_address_book
    old_load = AktaLogs._load_persona_from_address_book
    old_load_selected = AktaLogs._load_selected_address_book_entry

    def save_wrapper(self, persona_inputs):
        res = old_save(self, persona_inputs)
        try:
            section = 'pie' if persona_inputs == self.pie_in else 'nod' if persona_inputs == self.nod_in else 'rek'
            entry_name = None
            # find likely saved entry by latest matching name
            nos = persona_inputs[0].text().strip()
            for key, value in reversed(list(self.address_book.items())):
                if isinstance(value, dict) and value.get('nosaukums', '').strip() == nos:
                    entry_name = key
                    break
            if entry_name and entry_name in self.address_book:
                self.address_book[entry_name]['__custom_fields'] = _v54_persona_custom_payload(self, section)
                self._save_address_book()
        except Exception:
            pass
        return res

    def load_wrapper(self, persona_inputs):
        before = set(self.address_book.keys())
        res = old_load(self, persona_inputs)
        try:
            # old dialog already loaded selected item into inputs; infer from exact field match
            section = 'pie' if persona_inputs == self.pie_in else 'nod' if persona_inputs == self.nod_in else 'rek'
            nos = persona_inputs[0].text().strip()
            reg = persona_inputs[1].text().strip()
            for _, data in self.address_book.items():
                if isinstance(data, dict) and data.get('nosaukums', '').strip() == nos and data.get('reģ_nr', '').strip() == reg:
                    _v54_apply_custom_payload(self, section, data.get('__custom_fields') or {})
                    break
        except Exception:
            pass
        return res

    def load_selected_wrapper(self, item):
        res = old_load_selected(self, item)
        try:
            if not item:
                return res
            data = self.address_book.get(item.text()) or {}
            active_index = self.tabs.currentIndex()
            active_text = self.tabs.tabText(active_index)
            if 'Rekviz' in active_text:
                sec = 'rek'
            else:
                sec = 'pie'
            _v54_apply_custom_payload(self, sec, data.get('__custom_fields') or {})
        except Exception:
            pass
        return res

    AktaLogs._save_persona_to_address_book = save_wrapper
    AktaLogs._load_persona_from_address_book = load_wrapper
    AktaLogs._load_selected_address_book_entry = load_selected_wrapper


def _v54_history_settings(self):
    data = _v54_load_json(HISTORY_ACCESS_FILE, {})
    return data if isinstance(data, dict) else {}


def _v54_history_require_access(self):
    cfg = _v54_history_settings(self)
    salt = cfg.get('salt')
    pw_hash = cfg.get('hash')
    if not salt or not pw_hash:
        return True
    if getattr(self, '_history_access_granted', False):
        return True
    pw, ok = QInputDialog.getText(self, 'Dokumentu vēsture', 'Ievadiet paroli, lai atvērtu dokumentu vēsturi:', QLineEdit.Password)
    if not ok:
        return False
    if _v54_hash_password(pw, salt) != pw_hash:
        QMessageBox.warning(self, 'Kļūda', 'Nepareiza parole.')
        return False
    self._history_access_granted = True
    return True


def _v54_set_history_password(self):
    pw1, ok1 = QInputDialog.getText(self, 'Uzlikt paroli', 'Jaunā parole dokumentu vēsturei:', QLineEdit.Password)
    if not ok1:
        return
    pw2, ok2 = QInputDialog.getText(self, 'Uzlikt paroli', 'Atkārtojiet paroli:', QLineEdit.Password)
    if not ok2:
        return
    if not pw1 or pw1 != pw2:
        QMessageBox.warning(self, 'Kļūda', 'Paroles nesakrīt vai ir tukšas.')
        return
    salt = secrets.token_hex(8)
    _v54_save_json(HISTORY_ACCESS_FILE, {'salt': salt, 'hash': _v54_hash_password(pw1, salt)})
    self._history_access_granted = True
    QMessageBox.information(self, 'Saglabāts', 'Parole dokumentu vēsturei saglabāta.')


def _v54_remove_history_password(self):
    if not _v54_history_require_access(self):
        return
    _v54_save_json(HISTORY_ACCESS_FILE, {})
    self._history_access_granted = True
    QMessageBox.information(self, 'Saglabāts', 'Dokumentu vēstures parole noņemta.')


def _v54_patch_history_tab(self):
    try:
        idx = -1
        for i in range(self.tabs.count()):
            if self.tabs.tabText(i) == 'Dokumentu vēsture':
                idx = i
                break
        if idx < 0:
            return
        page = self.tabs.widget(idx)
        layout = page.layout()
        if layout is None:
            return
        row = QHBoxLayout()
        btn_set = QPushButton('Uzlikt paroli')
        btn_set.clicked.connect(lambda *_: _v54_set_history_password(self))
        btn_remove = QPushButton('Noņemt paroli')
        btn_remove.clicked.connect(lambda *_: _v54_remove_history_password(self))
        row.addWidget(btn_set)
        row.addWidget(btn_remove)
        row.addStretch(1)
        layout.insertLayout(0, row)
        self.tabs.currentChanged.connect(lambda i: _v54_on_tab_changed(self, i))
    except Exception:
        pass


def _v54_on_tab_changed(self, index):
    try:
        if self.tabs.tabText(index) == 'Dokumentu vēsture' and not _v54_history_require_access(self):
            self.tabs.setCurrentIndex(0)
    except Exception:
        pass


def _v54_patch_generation_history():
    old_pdf = AktaLogs.ģenerēt_pdf_dialogs
    old_docx = AktaLogs.ģenerēt_docx_dialogs

    def pdf_wrapper(self):
        before = set(os.listdir(DEFAULT_OUTPUT_DIR)) if os.path.isdir(DEFAULT_OUTPUT_DIR) else set()
        res = old_pdf(self)
        try:
            after = set(os.listdir(DEFAULT_OUTPUT_DIR)) if os.path.isdir(DEFAULT_OUTPUT_DIR) else set()
            new_dirs = sorted(list(after - before))
            if new_dirs:
                folder = os.path.join(DEFAULT_OUTPUT_DIR, new_dirs[-1])
                pdf_path = next((os.path.join(folder, f) for f in os.listdir(folder) if f.lower().endswith('.pdf')), None)
                json_path = next((os.path.join(folder, f) for f in os.listdir(folder) if f.lower().endswith('.json')), None)
                if pdf_path and json_path and os.path.exists(pdf_path) and os.path.exists(json_path):
                    self._record_generated_document(pdf_path, json_path)
                    self._update_history_list()
        except Exception:
            pass
        return res

    def docx_wrapper(self):
        before = set(os.listdir(DEFAULT_OUTPUT_DIR)) if os.path.isdir(DEFAULT_OUTPUT_DIR) else set()
        res = old_docx(self)
        try:
            after = set(os.listdir(DEFAULT_OUTPUT_DIR)) if os.path.isdir(DEFAULT_OUTPUT_DIR) else set()
            new_dirs = sorted(list(after - before))
            if new_dirs:
                folder = os.path.join(DEFAULT_OUTPUT_DIR, new_dirs[-1])
                json_path = next((os.path.join(folder, f) for f in os.listdir(folder) if f.lower().endswith('.json')), None)
                docx_path = next((os.path.join(folder, f) for f in os.listdir(folder) if f.lower().endswith('.docx')), None)
                if json_path and docx_path and os.path.exists(json_path):
                    # saglabājam vismaz JSON vēsturē arī DOCX gadījumam
                    entry = {
                        'title': os.path.basename(folder),
                        'json_path': json_path,
                        'docx_path': docx_path,
                        'created_at': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
                        'akta_nr': self.in_akta_nr.text().strip(),
                    }
                    if not isinstance(self.history, list):
                        self.history = []
                    self.history.insert(0, entry)
                    self._save_history()
                    self._update_history_list()
        except Exception:
            pass
        return res

    AktaLogs.ģenerēt_pdf_dialogs = pdf_wrapper
    AktaLogs.ģenerēt_docx_dialogs = docx_wrapper


def _v54_after_init(self):
    _v54_load_doc_nr_settings(self)
    _v54_add_doc_nr_button(self)
    _v54_bind_optional_date_state(self)
    _v54_init_custom_fields(self)
    _v54_patch_history_tab(self)
    _v54_load_ui_state(self)


def _v54_close_wrapper(old_close):
    def wrapper(self, event):
        try:
            _v54_save_ui_state(self)
        except Exception:
            pass
        return old_close(self, event)
    return wrapper


_v54_old_init = AktaLogs.__init__

def _v54_init_wrapper(self, *args, **kwargs):
    _v54_old_init(self, *args, **kwargs)
    _v54_after_init(self)

AktaLogs.__init__ = _v54_init_wrapper
AktaLogs._generate_akta_nr = _v54_generate_akta_nr
AktaLogs._open_doc_nr_settings = _v54_open_doc_nr_settings
AktaLogs._v54_collect_custom_field_values = _v54_collect_custom_field_values
AktaLogs.closeEvent = _v54_close_wrapper(AktaLogs.closeEvent)
_v54_patch_address_book_custom_fields()
_v54_patch_generation_history()



# v61: main block moved to file end so all hotfix patches are applied before app startup


# ==============================
# v58 hotfixes: UI state, custom fields in PDF/preview, stable JSON-only history backup
# ==============================
HISTORY_BACKUP_SETTINGS_KEY = 'history_backup_dir'
DEFAULT_HISTORY_BACKUP_DIR = os.path.join(SETTINGS_DIR, 'document_history_json')


def _v58_get_history_backup_dir():
    try:
        st = load_settings() or {}
        folder = (st.get(HISTORY_BACKUP_SETTINGS_KEY) or '').strip()
        if not folder:
            folder = DEFAULT_HISTORY_BACKUP_DIR
        os.makedirs(folder, exist_ok=True)
        return folder
    except Exception:
        try:
            os.makedirs(DEFAULT_HISTORY_BACKUP_DIR, exist_ok=True)
        except Exception:
            pass
        return DEFAULT_HISTORY_BACKUP_DIR


def _v58_set_history_backup_dir(folder: str):
    folder = (folder or '').strip()
    if not folder:
        return False
    try:
        os.makedirs(folder, exist_ok=True)
        st = load_settings() or {}
        st[HISTORY_BACKUP_SETTINGS_KEY] = folder
        save_settings(st)
        return True
    except Exception:
        return False


def _v58_collect_current_custom_fields(self):
    try:
        return self._v54_collect_custom_field_values() if hasattr(self, '_v54_collect_custom_field_values') else {}
    except Exception:
        return {}


_v58_orig_load_ui_state = _v54_load_ui_state
_v58_orig_save_ui_state = _v54_save_ui_state
_v58_orig_bind_optional_date_state = _v54_bind_optional_date_state
_v58_orig_after_init = _v54_after_init


def _v58_load_ui_state(self):
    try:
        _v58_orig_load_ui_state(self)
    except Exception:
        pass
    try:
        mapping = [
            ('izpildes', self.ck_ieklaut_izpildes_terminu, self.in_izpildes_termins),
            ('pieņemšanas', self.ck_ieklaut_pienemsanas_datumu, self.in_pieņemšanas_datums),
            ('nodošanas', self.ck_ieklaut_nodosanas_datumu, self.in_nodošanas_datums),
        ]
        st = _v54_load_json(UI_STATE_FILE, {})
        dates = st.get('optional_dates', {}) if isinstance(st, dict) else {}
        for key, ck, de in mapping:
            info = dates.get(key, {}) if isinstance(dates, dict) else {}
            if isinstance(info, dict):
                if 'checked' in info:
                    ck.blockSignals(True)
                    ck.setChecked(bool(info.get('checked')))
                    ck.blockSignals(False)
                date_text = str(info.get('date') or '').strip()
                if date_text:
                    qd = QDate.fromString(date_text, 'yyyy-MM-dd')
                    if qd.isValid():
                        de.blockSignals(True)
                        de.setDate(qd)
                        de.blockSignals(False)
    except Exception:
        pass


def _v58_save_ui_state(self):
    try:
        _v58_orig_save_ui_state(self)
    except Exception:
        pass


def _v58_bind_optional_date_state(self):
    try:
        _v58_orig_bind_optional_date_state(self)
    except Exception:
        pass
    for obj in [
        getattr(self, 'ck_ieklaut_izpildes_terminu', None),
        getattr(self, 'ck_ieklaut_pienemsanas_datumu', None),
        getattr(self, 'ck_ieklaut_nodosanas_datumu', None),
        getattr(self, 'in_izpildes_termins', None),
        getattr(self, 'in_pieņemšanas_datums', None),
        getattr(self, 'in_nodošanas_datums', None),
    ]:
        if obj is None:
            continue
        try:
            if hasattr(obj, 'stateChanged'):
                obj.stateChanged.connect(lambda *_args, s=self: _v58_save_ui_state(s))
            if hasattr(obj, 'dateChanged'):
                obj.dateChanged.connect(lambda *_args, s=self: _v58_save_ui_state(s))
        except Exception:
            pass


def _v58_wrap_after_load(old_func):
    def wrapper(self, *args, **kwargs):
        res = old_func(self, *args, **kwargs)
        try:
            for delay in (0, 50, 250):
                QTimer.singleShot(delay, lambda s=self: _v58_load_ui_state(s))
        except Exception:
            pass
        return res
    return wrapper


def _v58_history_items_from_backup_dir(folder: str):
    items = []
    try:
        if not os.path.isdir(folder):
            return []
        for name in os.listdir(folder):
            if not str(name).lower().endswith('.json'):
                continue
            path = os.path.join(folder, name)
            try:
                with open(path, 'r', encoding='utf-8') as f:
                    data = json.load(f) or {}
            except Exception:
                data = {}
            meta = data.get('_history_meta', {}) if isinstance(data, dict) else {}
            akta_nr = str(data.get('akta_nr') or meta.get('akta_nr') or '').strip() if isinstance(data, dict) else ''
            title = str(data.get('dokumenta_nosaukums') or meta.get('title') or os.path.splitext(name)[0]).strip() if isinstance(data, dict) else os.path.splitext(name)[0]
            created = str(meta.get('created_at') or datetime.fromtimestamp(os.path.getmtime(path)).strftime('%Y-%m-%d %H:%M:%S'))
            items.append({'json': path, 'pdf': '', 'created': created, 'akta_nr': akta_nr, 'title': title})
        items.sort(key=lambda x: str(x.get('created') or ''), reverse=True)
    except Exception:
        pass
    return items


def _v58_load_history(self):
    self.history = _v58_history_items_from_backup_dir(_v58_get_history_backup_dir())


def _v58_save_history(self):
    try:
        _v58_set_history_backup_dir(_v58_get_history_backup_dir())
    except Exception:
        pass


def _v58_update_history_list(self):
    self.history = _v58_history_items_from_backup_dir(_v58_get_history_backup_dir())
    try:
        self.history_list.clear()
        for it in self.history:
            title = str(it.get('title') or os.path.basename(it.get('json') or '')).strip()
            akta_nr = str(it.get('akta_nr') or '').strip()
            created = str(it.get('created') or '').strip()
            label = title
            if akta_nr:
                label += f' | {akta_nr}'
            if created:
                label += f' | {created}'
            item = QListWidgetItem(label)
            item.setData(Qt.UserRole, it)
            self.history_list.addItem(item)
        if hasattr(self, '_history_folder_label'):
            self._history_folder_label.setText(f"Mape: {_v58_get_history_backup_dir()}")
    except Exception:
        pass


def _v58_load_history_entry(self, item):
    if not item:
        return
    payload = item.data(Qt.UserRole)
    json_path = payload.get('json', '') if isinstance(payload, dict) else str(payload or '')
    json_path = str(json_path or '').strip()
    if json_path and os.path.exists(json_path):
        self.ieladet_projektu(json_path)
        try:
            for delay in (0, 50, 250):
                QTimer.singleShot(delay, lambda s=self: _v58_load_ui_state(s))
        except Exception:
            pass
        return
    QMessageBox.warning(self, 'Kļūda', 'Izvēlētais JSON fails nav atrasts.')


def _v58_clear_history(self):
    folder = _v58_get_history_backup_dir()
    if QMessageBox.question(self, 'Dzēst vēsturi', f'Dzēst visus JSON failus no dokumentu vēstures mapes?\n\n{folder}', QMessageBox.Yes | QMessageBox.No, QMessageBox.No) != QMessageBox.Yes:
        return
    errs = []
    try:
        for name in os.listdir(folder):
            if str(name).lower().endswith('.json'):
                try:
                    os.remove(os.path.join(folder, name))
                except Exception as e:
                    errs.append(f'{name}: {e}')
    except Exception as e:
        errs.append(str(e))
    self._update_history_list()
    if errs:
        QMessageBox.warning(self, 'Pabeigts ar kļūdām', '\n'.join(errs[:15]))


def _v58_open_document_folder(self):
    folder = _v58_get_history_backup_dir()
    try:
        os.makedirs(folder, exist_ok=True)
    except Exception:
        pass
    try:
        if sys.platform.startswith('win'):
            os.startfile(folder)
        elif sys.platform == 'darwin':
            subprocess.Popen(['open', folder])
        else:
            subprocess.Popen(['xdg-open', folder])
    except Exception as e:
        QMessageBox.warning(self, 'Kļūda', f'Neizdevās atvērt mapi:\n{e}')


def _v58_choose_history_folder(self):
    current = _v58_get_history_backup_dir()
    folder = QFileDialog.getExistingDirectory(self, 'Izvēlieties dokumentu vēstures/backup mapi', current)
    if not folder:
        return
    if _v58_set_history_backup_dir(folder):
        self._load_history()
        self._update_history_list()
        QMessageBox.information(self, 'Saglabāts', 'Dokumentu vēstures/backup mape saglabāta.')
    else:
        QMessageBox.warning(self, 'Kļūda', 'Neizdevās saglabāt izvēlēto mapi.')


def _v58_patch_history_ui(self):
    if getattr(self, '_v58_history_ui_patched', False):
        return
    self._v58_history_ui_patched = True
    try:
        idx = -1
        for i in range(self.tabs.count()):
            if self.tabs.tabText(i) == 'Dokumentu vēsture':
                idx = i
                break
        if idx < 0:
            return
        page = self.tabs.widget(idx)
        layout = page.layout()
        if layout is None:
            return
        row = QHBoxLayout()
        btn_folder = QPushButton('Izvēlēties backup mapi')
        btn_folder.clicked.connect(lambda *_: _v58_choose_history_folder(self))
        self._history_folder_label = QLabel(f"Mape: {_v58_get_history_backup_dir()}")
        self._history_folder_label.setWordWrap(True)
        row.addWidget(btn_folder)
        row.addWidget(self._history_folder_label, 1)
        layout.insertLayout(1, row)
        self._update_history_list()
    except Exception:
        pass


def _v58_record_generated_document(self, pdf_path: str, json_path: str):
    try:
        folder = _v58_get_history_backup_dir()
        os.makedirs(folder, exist_ok=True)
        with open(json_path, 'r', encoding='utf-8') as f:
            payload = json.load(f) or {}
        if isinstance(payload, dict):
            payload['custom_fields'] = _v58_collect_current_custom_fields(self)
            payload['_history_meta'] = {
                'created_at': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
                'source_json': json_path,
                'source_pdf': pdf_path,
                'akta_nr': str(payload.get('akta_nr') or self.in_akta_nr.text().strip()),
                'title': str(payload.get('dokumenta_nosaukums') or self.in_doc_nosaukums.text().strip() or 'Dokuments'),
            }
        safe_doc = drošs_faila_nosaukums(str(payload.get('akta_nr') or self.in_akta_nr.text().strip() or 'dokuments'))
        stamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        dst_json = os.path.join(folder, f'{stamp}_{safe_doc}.json')
        with open(dst_json, 'w', encoding='utf-8') as f:
            json.dump(payload, f, ensure_ascii=False, indent=2)
        self._load_history()
        self._update_history_list()
    except Exception:
        pass


def _v58_add_to_history(self, file_path: str):
    self._load_history()
    self._update_history_list()


def _v58_collect_custom_field_values(self):
    return _v54_collect_custom_field_values(self)


def _v58_savakt_datus_wrapper(old_func):
    def wrapper(self):
        d = old_func(self)
        try:
            setattr(d, 'custom_fields', _v58_collect_current_custom_fields(self))
        except Exception:
            pass
        try:
            _v58_save_ui_state(self)
        except Exception:
            pass
        return d
    return wrapper


def _v58_after_init(self):
    try:
        _v58_orig_after_init(self)
    except Exception:
        pass
    try:
        _v58_patch_history_ui(self)
    except Exception:
        pass
    try:
        for delay in (0, 50, 250):
            QTimer.singleShot(delay, lambda s=self: _v58_load_ui_state(s))
    except Exception:
        pass


AktaLogs._load_history = _v58_load_history
AktaLogs._save_history = _v58_save_history
AktaLogs._update_history_list = _v58_update_history_list
AktaLogs._load_history_entry = _v58_load_history_entry
AktaLogs._clear_history = _v58_clear_history
AktaLogs._open_document_folder = _v58_open_document_folder
AktaLogs._record_generated_document = _v58_record_generated_document
AktaLogs._add_to_history = _v58_add_to_history
AktaLogs.savākt_datus = _v58_savakt_datus_wrapper(AktaLogs.savākt_datus)
AktaLogs._v54_collect_custom_field_values = _v58_collect_custom_field_values

try:
    AktaLogs.ieladet_noklusejuma_iestatijumus = _v58_wrap_after_load(AktaLogs.ieladet_noklusejuma_iestatijumus)
except Exception:
    pass
try:
    AktaLogs.ieladet_projektu = _v58_wrap_after_load(AktaLogs.ieladet_projektu)
except Exception:
    pass
try:
    AktaLogs._load_selected_template = _v58_wrap_after_load(AktaLogs._load_selected_template)
except Exception:
    pass
try:
    AktaLogs._auto_load_test_template_first_run = _v58_wrap_after_load(AktaLogs._auto_load_test_template_first_run)
except Exception:
    pass

_v54_load_ui_state = _v58_load_ui_state
_v54_save_ui_state = _v58_save_ui_state
_v54_bind_optional_date_state = _v58_bind_optional_date_state
_v54_after_init = _v58_after_init



# ==============================
# v59 hotfixes: reliable optional date checkbox persistence, JSON backup history, custom fields preview refresh
# ==============================

def _v59_ui_state_payload(self):
    def _date_text(w):
        try:
            return w.date().toString('yyyy-MM-dd')
        except Exception:
            return ''
    try:
        return {
            'optional_dates': {
                'izpildes': {
                    'checked': bool(getattr(self, 'ck_ieklaut_izpildes_terminu').isChecked()),
                    'date': _date_text(getattr(self, 'in_izpildes_termins')),
                },
                'pieņemšanas': {
                    'checked': bool(getattr(self, 'ck_ieklaut_pienemsanas_datumu').isChecked()),
                    'date': _date_text(getattr(self, 'in_pieņemšanas_datums')),
                },
                'nodošanas': {
                    'checked': bool(getattr(self, 'ck_ieklaut_nodosanas_datumu').isChecked()),
                    'date': _date_text(getattr(self, 'in_nodošanas_datums')),
                },
            },
            'custom_field_definitions': getattr(self, '_custom_field_definitions', {}),
            'custom_field_values': self._v54_collect_custom_field_values() if hasattr(self, '_v54_collect_custom_field_values') else {},
        }
    except Exception:
        return {'optional_dates': {}}


def _v59_save_ui_state(self):
    try:
        st = _v54_load_json(UI_STATE_FILE, {})
        if not isinstance(st, dict):
            st = {}
        st.update(_v59_ui_state_payload(self))
        _v54_save_json(UI_STATE_FILE, st)
    except Exception:
        pass


def _v59_restore_optional_dates(self):
    try:
        st = _v54_load_json(UI_STATE_FILE, {})
        dates = st.get('optional_dates', {}) if isinstance(st, dict) else {}
        mapping = [
            ('izpildes', getattr(self, 'ck_ieklaut_izpildes_terminu', None), getattr(self, 'in_izpildes_termins', None)),
            ('pieņemšanas', getattr(self, 'ck_ieklaut_pienemsanas_datumu', None), getattr(self, 'in_pieņemšanas_datums', None)),
            ('nodošanas', getattr(self, 'ck_ieklaut_nodosanas_datumu', None), getattr(self, 'in_nodošanas_datums', None)),
        ]
        for key, ck, de in mapping:
            if ck is None or de is None:
                continue
            info = dates.get(key, {}) if isinstance(dates, dict) else {}
            if not isinstance(info, dict):
                continue
            try:
                ck.blockSignals(True)
                de.blockSignals(True)
                if 'checked' in info:
                    ck.setChecked(bool(info.get('checked')))
                dt = str(info.get('date') or '').strip()
                if dt:
                    qd = QDate.fromString(dt, 'yyyy-MM-dd')
                    if qd.isValid():
                        de.setDate(qd)
            finally:
                try:
                    ck.blockSignals(False)
                    de.blockSignals(False)
                except Exception:
                    pass
        try:
            # force refresh of enable/disable state if original handler exists
            if hasattr(self, '_refresh_optional_date_widgets_state'):
                self._refresh_optional_date_widgets_state()
        except Exception:
            pass
    except Exception:
        pass


def _v59_load_ui_state(self):
    try:
        _v59_restore_optional_dates(self)
    except Exception:
        pass
    try:
        st = _v54_load_json(UI_STATE_FILE, {})
        defs = st.get('custom_field_definitions', {}) if isinstance(st, dict) else {}
        vals = st.get('custom_field_values', {}) if isinstance(st, dict) else {}
        if isinstance(defs, dict) and hasattr(self, '_custom_field_definitions'):
            merged = {'pie': [], 'nod': [], 'rek': []}
            for k in merged:
                v = defs.get(k)
                if isinstance(v, list):
                    merged[k] = [str(x) for x in v if str(x).strip()]
            self._custom_field_definitions = merged
            for sec in ('pie', 'nod', 'rek'):
                try:
                    _v54_rebuild_custom_fields(self, sec)
                except Exception:
                    pass
        if isinstance(vals, dict):
            for sec, mp in vals.items():
                if not isinstance(mp, dict):
                    continue
                for name, value in mp.items():
                    w = getattr(self, '_custom_field_widgets', {}).get(sec, {}).get(name)
                    if isinstance(w, QLineEdit):
                        w.setText(str(value or ''))
    except Exception:
        pass


def _v59_bind_optional_date_state(self):
    seen = getattr(self, '_v59_optional_bound', False)
    if seen:
        return
    self._v59_optional_bound = True
    widgets = [
        getattr(self, 'ck_ieklaut_izpildes_terminu', None),
        getattr(self, 'ck_ieklaut_pienemsanas_datumu', None),
        getattr(self, 'ck_ieklaut_nodosanas_datumu', None),
        getattr(self, 'in_izpildes_termins', None),
        getattr(self, 'in_pieņemšanas_datums', None),
        getattr(self, 'in_nodošanas_datums', None),
    ]
    for obj in widgets:
        if obj is None:
            continue
        try:
            if hasattr(obj, 'stateChanged'):
                obj.stateChanged.connect(lambda *_args, s=self: _v59_save_ui_state(s))
            if hasattr(obj, 'dateChanged'):
                obj.dateChanged.connect(lambda *_args, s=self: _v59_save_ui_state(s))
        except Exception:
            pass
    try:
        _v59_save_ui_state(self)
    except Exception:
        pass


def _v59_connect_custom_field_refresh(self):
    if getattr(self, '_v59_custom_refresh_bound', False):
        return
    self._v59_custom_refresh_bound = True
    for sec in ('pie', 'nod', 'rek'):
        for widget in getattr(self, '_custom_field_widgets', {}).get(sec, {}).values():
            if isinstance(widget, QLineEdit):
                try:
                    widget.textChanged.connect(lambda *_args, s=self: ( _v59_save_ui_state(s), s._update_preview()))
                except Exception:
                    pass


def _v59_history_copy_json_to_backup(self, json_path: str):
    json_path = str(json_path or '').strip()
    if not json_path or not os.path.exists(json_path):
        return None
    try:
        folder = _v58_get_history_backup_dir()
        os.makedirs(folder, exist_ok=True)
        with open(json_path, 'r', encoding='utf-8') as f:
            payload = json.load(f) or {}
        if not isinstance(payload, dict):
            payload = {'data': payload}
        payload['custom_fields'] = _v58_collect_current_custom_fields(self)
        payload['ieklaut_izpildes_terminu'] = bool(getattr(self, 'ck_ieklaut_izpildes_terminu', None).isChecked()) if hasattr(self, 'ck_ieklaut_izpildes_terminu') else payload.get('ieklaut_izpildes_terminu', True)
        payload['ieklaut_pienemsanas_datumu'] = bool(getattr(self, 'ck_ieklaut_pienemsanas_datumu', None).isChecked()) if hasattr(self, 'ck_ieklaut_pienemsanas_datumu') else payload.get('ieklaut_pienemsanas_datumu', True)
        payload['ieklaut_nodosanas_datumu'] = bool(getattr(self, 'ck_ieklaut_nodosanas_datumu', None).isChecked()) if hasattr(self, 'ck_ieklaut_nodosanas_datumu') else payload.get('ieklaut_nodosanas_datumu', True)
        akta_nr = str(payload.get('akta_nr') or getattr(getattr(self, 'in_akta_nr', None), 'text', lambda: '')() or 'dokuments').strip()
        title = str(payload.get('dokumenta_nosaukums') or getattr(getattr(self, 'in_doc_nosaukums', None), 'text', lambda: '')() or 'Dokuments').strip()
        created_at = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        payload['_history_meta'] = {
            'created_at': created_at,
            'source_json': json_path,
            'akta_nr': akta_nr,
            'title': title,
        }
        safe_doc = drošs_faila_nosaukums(akta_nr or title or 'dokuments')
        stamp = datetime.now().strftime('%Y%m%d_%H%M%S_%f')
        dst_json = os.path.join(folder, f'{stamp}_{safe_doc}.json')
        with open(dst_json, 'w', encoding='utf-8') as f:
            json.dump(payload, f, ensure_ascii=False, indent=2)
        return dst_json
    except Exception:
        return None


def _v59_add_to_history(self, file_path: str):
    try:
        path = str(file_path or '').strip()
        if path.lower().endswith('.json') and os.path.exists(path):
            _v59_history_copy_json_to_backup(self, path)
    except Exception:
        pass
    try:
        self._load_history()
        self._update_history_list()
    except Exception:
        pass


def _v59_record_generated_document(self, pdf_path: str, json_path: str):
    try:
        _v59_history_copy_json_to_backup(self, json_path)
    except Exception:
        pass
    try:
        self._load_history()
        self._update_history_list()
    except Exception:
        pass


def _v59_after_init(self):
    try:
        _v58_after_init(self)
    except Exception:
        pass
    try:
        _v59_bind_optional_date_state(self)
        _v59_connect_custom_field_refresh(self)
        for delay in (0, 200, 800, 1500):
            QTimer.singleShot(delay, lambda s=self: _v59_load_ui_state(s))
    except Exception:
        pass


def _v59_closeEvent_wrapper(old_func):
    def wrapper(self, event):
        try:
            _v59_save_ui_state(self)
        except Exception:
            pass
        return old_func(self, event)
    return wrapper


def _v59_patch_custom_field_rebuild():
    old_rebuild = _v54_rebuild_custom_fields
    def patched(self, section):
        res = old_rebuild(self, section)
        try:
            _v59_connect_custom_field_refresh(self)
        except Exception:
            pass
        return res
    return patched


# apply v59 patches
_v54_rebuild_custom_fields = _v59_patch_custom_field_rebuild()
_v54_load_ui_state = _v59_load_ui_state
_v54_save_ui_state = _v59_save_ui_state
_v54_bind_optional_date_state = _v59_bind_optional_date_state
_v54_after_init = _v59_after_init
AktaLogs._add_to_history = _v59_add_to_history
AktaLogs._record_generated_document = _v59_record_generated_document
AktaLogs.closeEvent = _v59_closeEvent_wrapper(AktaLogs.closeEvent)

# ensure history list is refreshed from backup dir and folder selector remains visible
try:
    _v58_set_history_backup_dir(_v58_get_history_backup_dir())
except Exception:
    pass


# ==============================
# v60 hard fixes: in-tab document history with password, stable optional date defaults,
# template/project persistence for custom fields, and preview/PDF restore support.
# ==============================
V60_HISTORY_PASSWORD_KEY = 'history_tab_password'
V60_HISTORY_FOLDER_KEY = 'history_backup_dir'


def _v60_now_text():
    return datetime.now().strftime('%Y-%m-%d %H:%M:%S')


def _v60_safe_json_load(path, default=None):
    try:
        with open(path, 'r', encoding='utf-8') as f:
            return json.load(f)
    except Exception:
        return default


def _v60_hash_password(password: str, salt: str = '') -> str:
    password = str(password or '')
    salt = str(salt or '')
    return hashlib.sha256((salt + password).encode('utf-8')).hexdigest()


def _v60_get_history_cfg():
    st = load_settings() or {}
    cfg = st.get(V60_HISTORY_PASSWORD_KEY) or {}
    if not isinstance(cfg, dict):
        cfg = {}
    return cfg


def _v60_set_history_cfg(cfg: dict):
    st = load_settings() or {}
    st[V60_HISTORY_PASSWORD_KEY] = cfg or {}
    save_settings(st)


def _v60_get_history_folder():
    st = load_settings() or {}
    folder = str(st.get(V60_HISTORY_FOLDER_KEY) or '').strip()
    if not folder:
        folder = os.path.join(SETTINGS_DIR, 'document_history_json')
    try:
        os.makedirs(folder, exist_ok=True)
    except Exception:
        pass
    return folder


def _v60_set_history_folder(folder: str):
    folder = str(folder or '').strip()
    if not folder:
        return False
    try:
        os.makedirs(folder, exist_ok=True)
        st = load_settings() or {}
        st[V60_HISTORY_FOLDER_KEY] = folder
        save_settings(st)
        return True
    except Exception:
        return False


def _v60_collect_custom_field_bundle(self):
    defs = getattr(self, '_custom_field_definitions', {}) or {'pie': [], 'nod': [], 'rek': []}
    vals = _v54_collect_custom_field_values(self) if hasattr(self, '_custom_field_widgets') else {'pie': {}, 'nod': {}, 'rek': {}}
    return {
        'definitions': {
            'pie': list(defs.get('pie', []) or []),
            'nod': list(defs.get('nod', []) or []),
            'rek': list(defs.get('rek', []) or []),
        },
        'values': vals,
    }


def _v60_apply_custom_field_bundle(self, bundle):
    if not isinstance(bundle, dict):
        return
    defs = bundle.get('definitions') or {}
    vals = bundle.get('values') or {}
    if isinstance(defs, dict):
        merged = {'pie': [], 'nod': [], 'rek': []}
        for sec in merged:
            v = defs.get(sec)
            if isinstance(v, list):
                merged[sec] = [str(x) for x in v if str(x).strip()]
        self._custom_field_definitions = merged
        for sec in ('pie', 'nod', 'rek'):
            try:
                _v54_rebuild_custom_fields(self, sec)
            except Exception:
                pass
    if isinstance(vals, dict):
        for sec, mp in vals.items():
            if not isinstance(mp, dict):
                continue
            for name, value in mp.items():
                w = getattr(self, '_custom_field_widgets', {}).get(sec, {}).get(name)
                if isinstance(w, QLineEdit):
                    w.setText(str(value or ''))
    try:
        self._update_preview()
    except Exception:
        pass


def _v60_ui_optional_dates_payload(self):
    def _d(widget_name):
        w = getattr(self, widget_name, None)
        try:
            return w.date().toString('yyyy-MM-dd')
        except Exception:
            return ''
    def _c(name, default=True):
        w = getattr(self, name, None)
        try:
            return bool(w.isChecked())
        except Exception:
            return default
    return {
        'izpildes': {'checked': _c('ck_ieklaut_izpildes_terminu', True), 'date': _d('in_izpildes_termins')},
        'pieņemšanas': {'checked': _c('ck_ieklaut_pienemsanas_datumu', True), 'date': _d('in_pieņemšanas_datums')},
        'nodošanas': {'checked': _c('ck_ieklaut_nodosanas_datumu', True), 'date': _d('in_nodošanas_datums')},
    }


def _v60_save_ui_state(self):
    try:
        st = _v54_load_json(UI_STATE_FILE, {})
        if not isinstance(st, dict):
            st = {}
        st['optional_dates'] = _v60_ui_optional_dates_payload(self)
        st['custom_field_definitions'] = (getattr(self, '_custom_field_definitions', {}) or {})
        st['custom_field_values'] = _v54_collect_custom_field_values(self) if hasattr(self, '_custom_field_widgets') else {}
        _v54_save_json(UI_STATE_FILE, st)
    except Exception:
        pass


def _v60_restore_optional_dates(self, payload=None):
    try:
        if payload is None:
            st = _v54_load_json(UI_STATE_FILE, {})
            payload = st.get('optional_dates', {}) if isinstance(st, dict) else {}
        mapping = [
            ('izpildes', 'ck_ieklaut_izpildes_terminu', 'in_izpildes_termins'),
            ('pieņemšanas', 'ck_ieklaut_pienemsanas_datumu', 'in_pieņemšanas_datums'),
            ('nodošanas', 'ck_ieklaut_nodosanas_datumu', 'in_nodošanas_datums'),
        ]
        for key, ck_name, dt_name in mapping:
            info = payload.get(key, {}) if isinstance(payload, dict) else {}
            ck = getattr(self, ck_name, None)
            de = getattr(self, dt_name, None)
            if ck is None or de is None or not isinstance(info, dict):
                continue
            try:
                ck.blockSignals(True)
                de.blockSignals(True)
            except Exception:
                pass
            try:
                if 'checked' in info:
                    ck.setChecked(bool(info.get('checked')))
                dt = str(info.get('date') or '').strip()
                if dt:
                    qd = QDate.fromString(dt, 'yyyy-MM-dd')
                    if qd.isValid():
                        de.setDate(qd)
            finally:
                try:
                    ck.blockSignals(False)
                    de.blockSignals(False)
                except Exception:
                    pass
        try:
            if hasattr(self, '_refresh_optional_date_widgets_state'):
                self._refresh_optional_date_widgets_state()
        except Exception:
            pass
    except Exception:
        pass


def _v60_bind_optional_date_state(self):
    if getattr(self, '_v60_optional_bound', False):
        return
    self._v60_optional_bound = True
    for obj in [
        getattr(self, 'ck_ieklaut_izpildes_terminu', None),
        getattr(self, 'ck_ieklaut_pienemsanas_datumu', None),
        getattr(self, 'ck_ieklaut_nodosanas_datumu', None),
        getattr(self, 'in_izpildes_termins', None),
        getattr(self, 'in_pieņemšanas_datums', None),
        getattr(self, 'in_nodošanas_datums', None),
    ]:
        if obj is None:
            continue
        try:
            if hasattr(obj, 'stateChanged'):
                obj.stateChanged.connect(lambda *_a, s=self: _v60_save_ui_state(s))
            if hasattr(obj, 'dateChanged'):
                obj.dateChanged.connect(lambda *_a, s=self: _v60_save_ui_state(s))
        except Exception:
            pass


def _v60_history_backup_snapshot(self, source_json_path: str = ''):
    try:
        folder = _v60_get_history_folder()
        os.makedirs(folder, exist_ok=True)
        d = self.savākt_datus()
        payload = asdict(d)
        payload['pieņēmējs'] = asdict(d.pieņēmējs)
        payload['nodevējs'] = asdict(d.nodevējs)
        payload['rekviziti'] = asdict(getattr(d, 'rekviziti', Persona()))
        payload['custom_fields'] = _v60_collect_custom_field_bundle(self)
        payload['optional_dates'] = _v60_ui_optional_dates_payload(self)
        payload['_history_meta'] = {
            'created_at': _v60_now_text(),
            'source_json': str(source_json_path or ''),
            'akta_nr': str(payload.get('akta_nr') or ''),
            'title': str(payload.get('dokumenta_nosaukums') or 'Dokuments'),
        }
        safe_base = drošs_faila_nosaukums(str(payload.get('akta_nr') or payload.get('dokumenta_nosaukums') or 'dokuments')) or 'dokuments'
        stamp = datetime.now().strftime('%Y%m%d_%H%M%S_%f')
        out_path = os.path.join(folder, f'{stamp}_{safe_base}.json')
        with open(out_path, 'w', encoding='utf-8') as f:
            json.dump(payload, f, ensure_ascii=False, indent=2, default=str)
        return out_path
    except Exception as e:
        print(f'History snapshot save failed: {e}')
        return None


def _v60_history_items():
    items = []
    folder = _v60_get_history_folder()
    try:
        for name in sorted(os.listdir(folder), reverse=True):
            if not str(name).lower().endswith('.json'):
                continue
            path = os.path.join(folder, name)
            data = _v60_safe_json_load(path, {}) or {}
            meta = data.get('_history_meta', {}) if isinstance(data, dict) else {}
            title = str((data.get('dokumenta_nosaukums') if isinstance(data, dict) else '') or meta.get('title') or os.path.splitext(name)[0])
            akta_nr = str((data.get('akta_nr') if isinstance(data, dict) else '') or meta.get('akta_nr') or '')
            created = str(meta.get('created_at') or datetime.fromtimestamp(os.path.getmtime(path)).strftime('%Y-%m-%d %H:%M:%S'))
            items.append({'json': path, 'title': title, 'akta_nr': akta_nr, 'created': created})
    except Exception:
        pass
    items.sort(key=lambda x: str(x.get('created') or ''), reverse=True)
    return items


def _v60_refresh_history_list(self):
    try:
        self.history = _v60_history_items()
        if hasattr(self, 'history_list') and self.history_list is not None:
            self.history_list.clear()
            for it in self.history:
                label = str(it.get('title') or os.path.basename(it.get('json') or '')).strip()
                if it.get('akta_nr'):
                    label += f" | {it.get('akta_nr')}"
                if it.get('created'):
                    label += f" | {it.get('created')}"
                item = QListWidgetItem(label)
                item.setData(Qt.UserRole, it)
                self.history_list.addItem(item)
        if hasattr(self, '_history_folder_label'):
            self._history_folder_label.setText(f"Mape: {_v60_get_history_folder()}")
        if hasattr(self, '_history_count_label'):
            self._history_count_label.setText(f"Ieraksti: {len(self.history)}")
    except Exception:
        pass


def _v60_apply_history_access_ui(self):
    cfg = _v60_get_history_cfg()
    locked = bool(cfg.get('hash')) and not getattr(self, '_history_access_granted', False)
    try:
        if hasattr(self, '_history_auth_widget'):
            self._history_auth_widget.setVisible(locked)
        if hasattr(self, '_history_content_widget'):
            self._history_content_widget.setVisible(not locked)
        if hasattr(self, '_history_status_label'):
            self._history_status_label.setText('Tabs aizsargāts ar paroli.' if bool(cfg.get('hash')) else 'Tabs nav aizsargāts ar paroli.')
        if hasattr(self, '_history_unlock_btn'):
            self._history_unlock_btn.setVisible(bool(cfg.get('hash')))
        if hasattr(self, '_history_remove_pw_btn'):
            self._history_remove_pw_btn.setVisible(bool(cfg.get('hash')))
        if hasattr(self, '_history_set_pw_btn'):
            self._history_set_pw_btn.setText('Mainīt paroli' if bool(cfg.get('hash')) else 'Uzlikt paroli')
    except Exception:
        pass
    if not locked:
        _v60_refresh_history_list(self)


def _v60_history_unlock(self):
    cfg = _v60_get_history_cfg()
    if not cfg.get('hash'):
        self._history_access_granted = True
        _v60_apply_history_access_ui(self)
        return
    pw = self._history_password_input.text() if hasattr(self, '_history_password_input') else ''
    if _v60_hash_password(pw, cfg.get('salt', '')) == cfg.get('hash'):
        self._history_access_granted = True
        if hasattr(self, '_history_password_input'):
            self._history_password_input.clear()
        _v60_apply_history_access_ui(self)
    else:
        QMessageBox.warning(self, 'Kļūda', 'Nepareiza parole.')


def _v60_history_set_password(self):
    pw1 = self._history_new_password.text().strip() if hasattr(self, '_history_new_password') else ''
    pw2 = self._history_new_password_confirm.text().strip() if hasattr(self, '_history_new_password_confirm') else ''
    if not pw1:
        QMessageBox.warning(self, 'Kļūda', 'Ievadiet paroli.')
        return
    if pw1 != pw2:
        QMessageBox.warning(self, 'Kļūda', 'Paroles nesakrīt.')
        return
    salt = secrets.token_hex(8)
    _v60_set_history_cfg({'salt': salt, 'hash': _v60_hash_password(pw1, salt)})
    self._history_access_granted = False
    if hasattr(self, '_history_new_password'):
        self._history_new_password.clear()
    if hasattr(self, '_history_new_password_confirm'):
        self._history_new_password_confirm.clear()
    QMessageBox.information(self, 'Saglabāts', 'Parole dokumentu vēstures tabam saglabāta.')
    _v60_apply_history_access_ui(self)


def _v60_history_remove_password(self):
    _v60_set_history_cfg({})
    self._history_access_granted = True
    QMessageBox.information(self, 'Saglabāts', 'Parole noņemta.')
    _v60_apply_history_access_ui(self)


def _v60_history_choose_folder(self):
    current = _v60_get_history_folder()
    folder = QFileDialog.getExistingDirectory(self, 'Izvēlieties dokumentu vēstures backup mapi', current)
    if folder and _v60_set_history_folder(folder):
        _v60_refresh_history_list(self)


def _v60_history_clear(self):
    folder = _v60_get_history_folder()
    if QMessageBox.question(self, 'Notīrīt vēsturi', f'Dzēst visus JSON backup failus no mapes?\n\n{folder}', QMessageBox.Yes | QMessageBox.No, QMessageBox.No) != QMessageBox.Yes:
        return
    for name in list(os.listdir(folder)):
        if str(name).lower().endswith('.json'):
            try:
                os.remove(os.path.join(folder, name))
            except Exception:
                pass
    _v60_refresh_history_list(self)


def _v60_history_open_folder(self):
    folder = _v60_get_history_folder()
    try:
        if sys.platform.startswith('win'):
            os.startfile(folder)
        elif sys.platform == 'darwin':
            subprocess.Popen(['open', folder])
        else:
            subprocess.Popen(['xdg-open', folder])
    except Exception as e:
        QMessageBox.warning(self, 'Kļūda', f'Neizdevās atvērt mapi:\n{e}')


def _v60_history_load_selected(self):
    item = self.history_list.currentItem() if hasattr(self, 'history_list') else None
    if not item:
        return
    payload = item.data(Qt.UserRole) or {}
    path = str(payload.get('json') or '').strip()
    if path and os.path.exists(path):
        self.ieladet_projektu(path)


def _v60_rebuild_history_tab(self):
    try:
        idx = -1
        for i in range(self.tabs.count()):
            if self.tabs.tabText(i) == 'Dokumentu vēsture':
                idx = i
                break
        if idx < 0:
            return
        page = self.tabs.widget(idx)
        old_layout = page.layout()
        if old_layout is not None:
            while old_layout.count():
                item = old_layout.takeAt(0)
                w = item.widget()
                if w is not None:
                    w.deleteLater()
        else:
            old_layout = QVBoxLayout(page)
        root = old_layout

        self._history_auth_widget = QWidget(page)
        auth = QVBoxLayout(self._history_auth_widget)
        self._history_status_label = QLabel('')
        auth.addWidget(self._history_status_label)
        self._history_password_input = QLineEdit(); self._history_password_input.setEchoMode(QLineEdit.Password)
        self._history_password_input.setPlaceholderText('Ievadiet paroli, lai atvērtu dokumentu vēsturi')
        self._history_unlock_btn = QPushButton('Atvērt vēsturi')
        self._history_unlock_btn.clicked.connect(lambda *_: _v60_history_unlock(self))
        auth_row = QHBoxLayout(); auth_row.addWidget(self._history_password_input, 1); auth_row.addWidget(self._history_unlock_btn)
        auth.addLayout(auth_row)
        self._history_new_password = QLineEdit(); self._history_new_password.setEchoMode(QLineEdit.Password); self._history_new_password.setPlaceholderText('Jauna parole')
        self._history_new_password_confirm = QLineEdit(); self._history_new_password_confirm.setEchoMode(QLineEdit.Password); self._history_new_password_confirm.setPlaceholderText('Atkārtojiet paroli')
        auth.addWidget(self._history_new_password)
        auth.addWidget(self._history_new_password_confirm)
        auth_btns = QHBoxLayout()
        self._history_set_pw_btn = QPushButton('Uzlikt paroli')
        self._history_set_pw_btn.clicked.connect(lambda *_: _v60_history_set_password(self))
        self._history_remove_pw_btn = QPushButton('Noņemt paroli')
        self._history_remove_pw_btn.clicked.connect(lambda *_: _v60_history_remove_password(self))
        btn_choose_folder_locked = QPushButton('Izvēlēties backup mapi')
        btn_choose_folder_locked.clicked.connect(lambda *_: _v60_history_choose_folder(self))
        auth_btns.addWidget(self._history_set_pw_btn)
        auth_btns.addWidget(self._history_remove_pw_btn)
        auth_btns.addWidget(btn_choose_folder_locked)
        auth_btns.addStretch(1)
        auth.addLayout(auth_btns)

        self._history_content_widget = QWidget(page)
        content = QVBoxLayout(self._history_content_widget)
        top = QHBoxLayout()
        self._history_folder_label = QLabel(f'Mape: {_v60_get_history_folder()}')
        self._history_folder_label.setWordWrap(True)
        self._history_count_label = QLabel('Ieraksti: 0')
        top.addWidget(self._history_folder_label, 1)
        top.addWidget(self._history_count_label)
        content.addLayout(top)
        self.history_list = QListWidget(); self.history_list.itemDoubleClicked.connect(lambda *_: _v60_history_load_selected(self))
        content.addWidget(self.history_list)
        btns = QHBoxLayout()
        for txt, fn in [
            ('Ielādēt izvēlēto projektu', _v60_history_load_selected),
            ('Atjaunot sarakstu', _v60_refresh_history_list),
            ('Notīrīt vēsturi', _v60_history_clear),
            ('Atvērt mapi', _v60_history_open_folder),
            ('Izvēlēties backup mapi', _v60_history_choose_folder),
        ]:
            b = QPushButton(txt)
            b.clicked.connect(lambda *_a, f=fn: f(self))
            btns.addWidget(b)
        btn_lock = QPushButton('Aizslēgt tab')
        btn_lock.clicked.connect(lambda *_: (setattr(self, '_history_access_granted', False), _v60_apply_history_access_ui(self)))
        btns.addWidget(btn_lock)
        btns.addStretch(1)
        content.addLayout(btns)

        root.addWidget(self._history_auth_widget)
        root.addWidget(self._history_content_widget)
        _v60_apply_history_access_ui(self)
    except Exception as e:
        print(f'History tab rebuild failed: {e}')


def _v60_after_init(self):
    try:
        _v59_after_init(self)
    except Exception:
        pass
    try:
        _v60_bind_optional_date_state(self)
    except Exception:
        pass
    try:
        for delay in (0, 100, 400, 1200):
            QTimer.singleShot(delay, lambda s=self: _v60_restore_optional_dates(s))
    except Exception:
        pass
    try:
        _v60_rebuild_history_tab(self)
    except Exception:
        pass


def _v60_close_event_wrapper(old_func):
    def wrapper(self, event):
        try:
            _v60_save_ui_state(self)
        except Exception:
            pass
        return old_func(self, event)
    return wrapper


def _v60_savakt_wrapper(old_func):
    def wrapper(self):
        d = old_func(self)
        try:
            setattr(d, 'custom_fields', _v54_collect_custom_field_values(self))
        except Exception:
            pass
        return d
    return wrapper


def _v60_save_project_wrapper(old_func):
    def wrapper(self):
        # original method writes JSON chosen by user and calls _add_to_history
        res = old_func(self)
        try:
            _v60_save_ui_state(self)
            if getattr(self, '_ceļš_projekts', None) and os.path.exists(self._ceļš_projekts):
                payload = _v60_safe_json_load(self._ceļš_projekts, {}) or {}
                if isinstance(payload, dict):
                    payload['custom_fields'] = _v60_collect_custom_field_bundle(self)
                    payload['optional_dates'] = _v60_ui_optional_dates_payload(self)
                    with open(self._ceļš_projekts, 'w', encoding='utf-8') as f:
                        json.dump(payload, f, ensure_ascii=False, indent=2, default=str)
                _v60_history_backup_snapshot(self, self._ceļš_projekts)
        except Exception as e:
            print(f'Project post-save patch failed: {e}')
        return res
    return wrapper


def _v60_save_template_wrapper(old_func):
    def wrapper(self):
        res = old_func(self)
        try:
            item = self.sablonu_list.currentItem() if hasattr(self, 'sablonu_list') else None
            current_name = item.text() if item else ''
            path = self._resolve_template_path(current_name) if current_name else None
            if path and os.path.exists(path):
                payload = _v60_safe_json_load(path, {}) or {}
                if isinstance(payload, dict):
                    payload['custom_fields'] = _v60_collect_custom_field_bundle(self)
                    payload['optional_dates'] = _v60_ui_optional_dates_payload(self)
                    with open(path, 'w', encoding='utf-8') as f:
                        json.dump(payload, f, ensure_ascii=False, indent=2, default=str)
        except Exception as e:
            print(f'Template post-save patch failed: {e}')
        return res
    return wrapper


def _v60_load_project_wrapper(old_func):
    def wrapper(self, path=None):
        res = old_func(self, path)
        try:
            actual_path = path or getattr(self, '_ceļš_projekts', None)
            actual_path = _coerce_path(actual_path)
            if actual_path and os.path.exists(actual_path):
                payload = _v60_safe_json_load(actual_path, {}) or {}
                if isinstance(payload, dict):
                    _v60_restore_optional_dates(self, payload.get('optional_dates'))
                    _v60_apply_custom_field_bundle(self, payload.get('custom_fields'))
                    _v60_save_ui_state(self)
        except Exception as e:
            print(f'Project post-load patch failed: {e}')
        return res
    return wrapper


def _v60_load_template_wrapper(old_func):
    def wrapper(self, *args, **kwargs):
        res = old_func(self, *args, **kwargs)
        try:
            item = self.sablonu_list.currentItem() if hasattr(self, 'sablonu_list') else None
            current_name = item.text() if item else ''
            current_name = current_name.replace(' [Aizsargāts]', '').replace(' [Kļūda]', '').strip()
            path = self._resolve_template_path(current_name) if current_name else None
            if path and os.path.exists(path):
                payload = _v60_safe_json_load(path, {}) or {}
                if isinstance(payload, dict):
                    _v60_restore_optional_dates(self, payload.get('optional_dates'))
                    _v60_apply_custom_field_bundle(self, payload.get('custom_fields'))
                    _v60_save_ui_state(self)
        except Exception as e:
            print(f'Template post-load patch failed: {e}')
        return res
    return wrapper


def _v60_record_generated_document(self, pdf_path: str, json_path: str):
    try:
        if json_path and os.path.exists(json_path):
            payload = _v60_safe_json_load(json_path, {}) or {}
            if isinstance(payload, dict):
                payload['custom_fields'] = _v60_collect_custom_field_bundle(self)
                payload['optional_dates'] = _v60_ui_optional_dates_payload(self)
                with open(json_path, 'w', encoding='utf-8') as f:
                    json.dump(payload, f, ensure_ascii=False, indent=2, default=str)
        _v60_history_backup_snapshot(self, json_path)
        _v60_refresh_history_list(self)
    except Exception as e:
        print(f'Record generated document patch failed: {e}')


def _v60_add_to_history(self, file_path: str):
    try:
        if str(file_path or '').lower().endswith('.json') and os.path.exists(str(file_path)):
            _v60_history_backup_snapshot(self, str(file_path))
        _v60_refresh_history_list(self)
    except Exception as e:
        print(f'Add to history patch failed: {e}')


# apply v60 patches
_v54_save_ui_state = _v60_save_ui_state
_v54_bind_optional_date_state = _v60_bind_optional_date_state
_v54_after_init = _v60_after_init
AktaLogs.closeEvent = _v60_close_event_wrapper(AktaLogs.closeEvent)
AktaLogs.savākt_datus = _v60_savakt_wrapper(AktaLogs.savākt_datus)
AktaLogs.saglabat_projektu = _v60_save_project_wrapper(AktaLogs.saglabat_projektu)
AktaLogs.saglabat_ka_sablonu = _v60_save_template_wrapper(AktaLogs.saglabat_ka_sablonu)
AktaLogs.ieladet_projektu = _v60_load_project_wrapper(AktaLogs.ieladet_projektu)
try:
    AktaLogs._load_selected_template = _v60_load_template_wrapper(AktaLogs._load_selected_template)
except Exception:
    pass
AktaLogs._record_generated_document = _v60_record_generated_document
AktaLogs._add_to_history = _v60_add_to_history
AktaLogs._load_history = lambda self: setattr(self, 'history', _v60_history_items())
AktaLogs._save_history = lambda self: None
AktaLogs._update_history_list = _v60_refresh_history_list
AktaLogs._load_history_entry = lambda self, item: _v60_history_load_selected(self)
AktaLogs._clear_history = _v60_history_clear
AktaLogs._open_document_folder = _v60_history_open_folder




# ==============================
# v62 real fixes: polished history tab with inline auth, reliable optional-date
# persistence, and robust custom-field save/load + PDF/preview feed.
# ==============================
V62_UI_STATE_FILE = os.path.join(SETTINGS_DIR, "ui_state_v62.json")
V62_HISTORY_CFG_FILE = os.path.join(SETTINGS_DIR, "document_history_v62.json")
V62_HISTORY_DEFAULT_DIR = os.path.join(SETTINGS_DIR, "document_history_json")


def _v62_safe_load_json(path, default=None):
    try:
        if path and os.path.exists(path):
            with open(path, 'r', encoding='utf-8') as f:
                return json.load(f)
    except Exception:
        pass
    return {} if default is None else default


def _v62_safe_save_json(path, payload):
    try:
        os.makedirs(os.path.dirname(path), exist_ok=True)
        with open(path, 'w', encoding='utf-8') as f:
            json.dump(payload, f, ensure_ascii=False, indent=2, default=str)
        return True
    except Exception as e:
        print(f'JSON save failed for {path}: {e}')
        return False


def _v62_hash_password(password: str, salt: str) -> str:
    try:
        return hashlib.pbkdf2_hmac('sha256', (password or '').encode('utf-8'), (salt or '').encode('utf-8'), 120000).hex()
    except Exception:
        return ''


def _v62_history_cfg():
    cfg = _v62_safe_load_json(V62_HISTORY_CFG_FILE, {}) or {}
    if not isinstance(cfg, dict):
        cfg = {}
    folder = _coerce_path(cfg.get('folder')) or V62_HISTORY_DEFAULT_DIR
    cfg['folder'] = folder
    return cfg


def _v62_set_history_cfg(new_values: dict):
    cfg = _v62_history_cfg()
    cfg.update(new_values or {})
    folder = _coerce_path(cfg.get('folder')) or V62_HISTORY_DEFAULT_DIR
    cfg['folder'] = folder
    os.makedirs(folder, exist_ok=True)
    _v62_safe_save_json(V62_HISTORY_CFG_FILE, cfg)
    return cfg


def _v62_history_folder():
    folder = _coerce_path(_v62_history_cfg().get('folder')) or V62_HISTORY_DEFAULT_DIR
    os.makedirs(folder, exist_ok=True)
    return folder


def _v62_history_items():
    folder = _v62_history_folder()
    items = []
    try:
        for name in os.listdir(folder):
            if not str(name).lower().endswith('.json'):
                continue
            path = os.path.join(folder, name)
            if os.path.isfile(path):
                items.append(path)
    except Exception:
        pass
    items.sort(key=lambda p: os.path.getmtime(p) if os.path.exists(p) else 0, reverse=True)
    return items


def _v62_collect_optional_dates(self):
    def _date(widget_name):
        w = getattr(self, widget_name, None)
        try:
            return w.date().toString('yyyy-MM-dd') if w is not None else ''
        except Exception:
            return ''
    def _checked(name, default=True):
        w = getattr(self, name, None)
        try:
            return bool(w.isChecked())
        except Exception:
            return default
    return {
        'izpildes': {'checked': _checked('ck_ieklaut_izpildes_terminu', True), 'date': _date('in_izpildes_termins')},
        'pieņemšanas': {'checked': _checked('ck_ieklaut_pienemsanas_datumu', True), 'date': _date('in_pieņemšanas_datums')},
        'nodošanas': {'checked': _checked('ck_ieklaut_nodosanas_datumu', True), 'date': _date('in_nodošanas_datums')},
    }


def _v62_apply_optional_dates(self, payload=None):
    payload = payload or (_v62_safe_load_json(V62_UI_STATE_FILE, {}) or {}).get('optional_dates') or {}
    if not isinstance(payload, dict):
        return
    mapping = [
        ('izpildes', 'ck_ieklaut_izpildes_terminu', 'in_izpildes_termins'),
        ('pieņemšanas', 'ck_ieklaut_pienemsanas_datumu', 'in_pieņemšanas_datums'),
        ('nodošanas', 'ck_ieklaut_nodosanas_datumu', 'in_nodošanas_datums'),
    ]
    for key, ck_name, dt_name in mapping:
        ck = getattr(self, ck_name, None)
        dt = getattr(self, dt_name, None)
        info = payload.get(key) or {}
        if ck is not None and 'checked' in info:
            try:
                ck.blockSignals(True)
                ck.setChecked(bool(info.get('checked')))
            finally:
                ck.blockSignals(False)
        date_str = str(info.get('date') or '').strip()
        if dt is not None and date_str:
            try:
                qd = QDate.fromString(date_str, 'yyyy-MM-dd')
                if qd.isValid():
                    dt.blockSignals(True)
                    dt.setDate(qd)
                    dt.blockSignals(False)
            except Exception:
                pass
        try:
            if ck is not None and dt is not None:
                dt.setEnabled(bool(ck.isChecked()))
        except Exception:
            pass
    try:
        self._update_preview()
    except Exception:
        pass


def _v62_collect_custom_bundle(self):
    defs = getattr(self, '_custom_field_definitions', {}) or {'pie': [], 'nod': [], 'rek': []}
    values = {'pie': {}, 'nod': {}, 'rek': {}}
    for sec, mapping in (getattr(self, '_custom_field_widgets', {}) or {}).items():
        values.setdefault(sec, {})
        if isinstance(mapping, dict):
            for field_name, widget in mapping.items():
                try:
                    values[sec][str(field_name)] = widget.text().strip()
                except Exception:
                    values[sec][str(field_name)] = ''
    return {'definitions': defs, 'values': values}


def _v62_apply_custom_bundle(self, bundle):
    if not isinstance(bundle, dict):
        return
    defs = bundle.get('definitions') or {}
    vals = bundle.get('values') or {}
    merged = {'pie': [], 'nod': [], 'rek': []}
    for sec in merged:
        merged[sec] = [str(x).strip() for x in (defs.get(sec) or []) if str(x).strip()]
    self._custom_field_definitions = merged
    for sec in ('pie', 'nod', 'rek'):
        try:
            _v54_rebuild_custom_fields(self, sec)
        except Exception:
            pass
    for sec, sec_vals in (vals.items() if isinstance(vals, dict) else []):
        if not isinstance(sec_vals, dict):
            continue
        for name, value in sec_vals.items():
            try:
                w = self._custom_field_widgets.get(sec, {}).get(name)
                if w is not None:
                    w.setText(str(value or ''))
            except Exception:
                pass
    try:
        self._update_preview()
    except Exception:
        pass


def _v62_save_ui_state(self):
    state = _v62_safe_load_json(V62_UI_STATE_FILE, {}) or {}
    if not isinstance(state, dict):
        state = {}
    state['optional_dates'] = _v62_collect_optional_dates(self)
    state['custom_fields'] = _v62_collect_custom_bundle(self)
    _v62_safe_save_json(V62_UI_STATE_FILE, state)


def _v62_bind_live_state(self):
    widgets = [
        getattr(self, 'ck_ieklaut_izpildes_terminu', None),
        getattr(self, 'ck_ieklaut_pienemsanas_datumu', None),
        getattr(self, 'ck_ieklaut_nodosanas_datumu', None),
        getattr(self, 'in_izpildes_termins', None),
        getattr(self, 'in_pieņemšanas_datums', None),
        getattr(self, 'in_nodošanas_datums', None),
    ]
    for w in widgets:
        if w is None:
            continue
        try:
            if hasattr(w, 'toggled'):
                w.toggled.connect(lambda *_: _v62_save_ui_state(self))
            if hasattr(w, 'dateChanged'):
                w.dateChanged.connect(lambda *_: _v62_save_ui_state(self))
        except Exception:
            pass
    for sec in ('pie', 'nod', 'rek'):
        for widget in (getattr(self, '_custom_field_widgets', {}).get(sec, {}) or {}).values():
            try:
                widget.textChanged.connect(lambda *_: (_v62_save_ui_state(self), self._update_preview()))
            except Exception:
                pass


def _v62_history_verify(password: str) -> bool:
    cfg = _v62_history_cfg()
    salt = cfg.get('salt') or ''
    pw_hash = cfg.get('hash') or ''
    if not salt or not pw_hash:
        return True
    return _v62_hash_password(password or '', salt) == pw_hash


def _v62_history_has_password():
    cfg = _v62_history_cfg()
    return bool(cfg.get('salt') and cfg.get('hash'))


def _v62_history_enrich_payload(self, payload: dict):
    payload = dict(payload or {})
    payload['optional_dates'] = _v62_collect_optional_dates(self)
    bundle = _v62_collect_custom_bundle(self)
    payload['custom_fields_bundle'] = bundle
    payload['custom_fields'] = bundle.get('values', {})
    payload.setdefault('_history_meta', {})
    meta = payload['_history_meta'] if isinstance(payload['_history_meta'], dict) else {}
    meta.update({
        'saved_at': datetime.now().isoformat(timespec='seconds'),
        'app_version': 'v62',
        'akta_nr': payload.get('akta_nr', ''),
        'datums': payload.get('datums', ''),
        'title': payload.get('virsraksts', '') or payload.get('nosaukums', '') or 'Dokuments',
    })
    payload['_history_meta'] = meta
    return payload


def _v62_history_backup_snapshot(self, json_path: str):
    json_path = _coerce_path(json_path)
    if not json_path or not os.path.exists(json_path):
        return None
    payload = _v62_safe_load_json(json_path, {}) or {}
    if not isinstance(payload, dict):
        return None
    payload = _v62_history_enrich_payload(self, payload)
    folder = _v62_history_folder()
    ts = datetime.now().strftime('%Y%m%d_%H%M%S')
    doc_no = drošs_faila_nosaukums(str(payload.get('akta_nr') or 'dokuments')) or 'dokuments'
    out_path = os.path.join(folder, f'{ts}_{doc_no}.json')
    base, ext = os.path.splitext(out_path)
    counter = 2
    while os.path.exists(out_path):
        out_path = f'{base}_{counter}{ext}'
        counter += 1
    if _v62_safe_save_json(out_path, payload):
        return out_path
    return None


def _v62_history_display_label(path: str):
    data = _v62_safe_load_json(path, {}) or {}
    meta = data.get('_history_meta') if isinstance(data, dict) else {}
    if not isinstance(meta, dict):
        meta = {}
    akta_nr = str(meta.get('akta_nr') or data.get('akta_nr') or '').strip()
    datums = str(meta.get('datums') or data.get('datums') or '').strip()
    title = str(meta.get('title') or data.get('virsraksts') or data.get('nosaukums') or '').strip()
    saved_at = str(meta.get('saved_at') or '').strip().replace('T', ' ')
    parts = []
    if akta_nr:
        parts.append(akta_nr)
    if title:
        parts.append(title)
    text = ' • '.join(parts) if parts else os.path.basename(path)
    if datums:
        text += f'  |  Datums: {datums}'
    if saved_at:
        text += f'  |  Saglabāts: {saved_at}'
    return text


def _v62_apply_history_auth_ui(self):
    locked = _v62_history_has_password() and not bool(getattr(self, '_history_access_granted', False))
    if hasattr(self, '_history_auth_frame'):
        self._history_auth_frame.setVisible(locked or not _v62_history_has_password())
    if hasattr(self, '_history_content_frame'):
        self._history_content_frame.setVisible(not locked)
    if hasattr(self, '_history_status_badge'):
        self._history_status_badge.setText('Aizsargāts' if _v62_history_has_password() else 'Bez paroles')
    if hasattr(self, '_history_hint_label'):
        if locked:
            self._history_hint_label.setText('Lai redzētu dokumentu vēsturi, ievadiet paroli šajā tabā.')
        else:
            self._history_hint_label.setText('Šeit glabājas visi JSON backup dokumenti no izvēlētās mapes.')
    if hasattr(self, '_history_folder_label'):
        self._history_folder_label.setText(f'Backup mape: {_v62_history_folder()}')
    if not locked:
        _v62_refresh_history_list(self)


def _v62_history_unlock(self):
    pw = self._history_password_input.text().strip() if hasattr(self, '_history_password_input') else ''
    if _v62_history_has_password() and not _v62_history_verify(pw):
        if hasattr(self, '_history_hint_label'):
            self._history_hint_label.setText('Nepareiza parole. Mēģiniet vēlreiz.')
        return
    self._history_access_granted = True
    if hasattr(self, '_history_password_input'):
        self._history_password_input.clear()
    _v62_apply_history_auth_ui(self)


def _v62_history_set_password(self):
    pw1 = self._history_new_password.text().strip() if hasattr(self, '_history_new_password') else ''
    pw2 = self._history_new_password_confirm.text().strip() if hasattr(self, '_history_new_password_confirm') else ''
    if not pw1 or pw1 != pw2:
        if hasattr(self, '_history_hint_label'):
            self._history_hint_label.setText('Paroles nesakrīt vai ir tukšas.')
        return
    salt = secrets.token_hex(8)
    _v62_set_history_cfg({'salt': salt, 'hash': _v62_hash_password(pw1, salt)})
    self._history_access_granted = True
    if hasattr(self, '_history_new_password'):
        self._history_new_password.clear()
    if hasattr(self, '_history_new_password_confirm'):
        self._history_new_password_confirm.clear()
    _v62_apply_history_auth_ui(self)


def _v62_history_remove_password(self):
    if _v62_history_has_password():
        current = self._history_password_input.text().strip() if hasattr(self, '_history_password_input') else ''
        if not getattr(self, '_history_access_granted', False) and not _v62_history_verify(current):
            if hasattr(self, '_history_hint_label'):
                self._history_hint_label.setText('Lai noņemtu paroli, ievadiet pareizo esošo paroli.')
            return
    cfg = _v62_history_cfg()
    cfg.pop('salt', None)
    cfg.pop('hash', None)
    _v62_safe_save_json(V62_HISTORY_CFG_FILE, cfg)
    self._history_access_granted = True
    _v62_apply_history_auth_ui(self)


def _v62_history_choose_folder(self):
    current = _v62_history_folder()
    folder = QFileDialog.getExistingDirectory(self, 'Izvēlēties dokumentu backup mapi', current)
    if not folder:
        return
    _v62_set_history_cfg({'folder': folder})
    self._history_access_granted = True if not _v62_history_has_password() else getattr(self, '_history_access_granted', False)
    _v62_apply_history_auth_ui(self)


def _v62_refresh_history_list(self):
    if not hasattr(self, 'history_list'):
        return
    self.history_list.clear()
    items = _v62_history_items()
    self.history = [{'json': p, 'created': datetime.fromtimestamp(os.path.getmtime(p)).isoformat(timespec='seconds')} for p in items]
    for path in items:
        item = QListWidgetItem(_v62_history_display_label(path))
        item.setData(Qt.UserRole, path)
        self.history_list.addItem(item)
    if hasattr(self, '_history_count_label'):
        self._history_count_label.setText(f'Ieraksti: {len(items)}')
    if hasattr(self, '_history_folder_label'):
        self._history_folder_label.setText(f'Backup mape: {_v62_history_folder()}')


def _v62_history_load_selected(self):
    item = self.history_list.currentItem() if hasattr(self, 'history_list') else None
    if not item:
        return
    path = _coerce_path(item.data(Qt.UserRole)) or ''
    if path and os.path.exists(path):
        self.ieladet_projektu(path)


def _v62_history_clear(self):
    folder = _v62_history_folder()
    try:
        for path in _v62_history_items():
            try:
                os.remove(path)
            except Exception:
                pass
    finally:
        _v62_refresh_history_list(self)


def _v62_history_open_folder(self):
    folder = _v62_history_folder()
    try:
        if sys.platform == 'win32':
            os.startfile(folder)
        elif sys.platform == 'darwin':
            os.system(f'open "{folder}"')
        else:
            os.system(f'xdg-open "{folder}"')
    except Exception:
        pass


def _v62_clear_layout(layout):
    if layout is None:
        return
    while layout.count():
        item = layout.takeAt(0)
        if item.widget() is not None:
            item.widget().deleteLater()
        child = item.layout()
        if child is not None:
            _v62_clear_layout(child)


def _v62_rebuild_history_tab(self):
    idx = -1
    for i in range(self.tabs.count()):
        if self.tabs.tabText(i) == 'Dokumentu vēsture':
            idx = i
            break
    if idx < 0:
        return
    page = self.tabs.widget(idx)
    layout = page.layout() or QVBoxLayout(page)
    _v62_clear_layout(layout)
    layout.setContentsMargins(14, 14, 14, 14)
    layout.setSpacing(10)

    header = QFrame(page)
    header.setStyleSheet('QFrame{background:#111827;border:1px solid #253047;border-radius:14px;} QLabel{color:#E5E7EB;}')
    header_l = QVBoxLayout(header)
    row = QHBoxLayout()
    title = QLabel('Dokumentu vēsture')
    title.setStyleSheet('font-size:20px;font-weight:700;color:#F9FAFB;')
    self._history_status_badge = QLabel('')
    self._history_status_badge.setStyleSheet('padding:4px 10px;background:#1F2937;border:1px solid #334155;border-radius:10px;color:#D1FAE5;font-weight:600;')
    row.addWidget(title)
    row.addStretch(1)
    row.addWidget(self._history_status_badge)
    header_l.addLayout(row)
    self._history_hint_label = QLabel('Šeit glabājas visi JSON backup dokumenti no izvēlētās mapes.')
    self._history_hint_label.setWordWrap(True)
    self._history_hint_label.setStyleSheet('color:#CBD5E1;')
    header_l.addWidget(self._history_hint_label)
    layout.addWidget(header)

    self._history_auth_frame = QFrame(page)
    self._history_auth_frame.setStyleSheet('QFrame{background:#0F172A;border:1px solid #233048;border-radius:14px;} QLabel{color:#E5E7EB;} QLineEdit{min-height:34px;}')
    auth = QVBoxLayout(self._history_auth_frame)
    auth.addWidget(QLabel('Piekļuve un iestatījumi'))
    self._history_password_input = QLineEdit(); self._history_password_input.setEchoMode(QLineEdit.Password); self._history_password_input.setPlaceholderText('Esošā parole')
    self._history_new_password = QLineEdit(); self._history_new_password.setEchoMode(QLineEdit.Password); self._history_new_password.setPlaceholderText('Jauna parole')
    self._history_new_password_confirm = QLineEdit(); self._history_new_password_confirm.setEchoMode(QLineEdit.Password); self._history_new_password_confirm.setPlaceholderText('Atkārtojiet jauno paroli')
    auth_row1 = QHBoxLayout(); auth_row1.addWidget(self._history_password_input, 1)
    btn_unlock = QPushButton('Atvērt')
    btn_unlock.clicked.connect(lambda *_: _v62_history_unlock(self))
    auth_row1.addWidget(btn_unlock)
    auth.addLayout(auth_row1)
    auth_row2 = QHBoxLayout(); auth_row2.addWidget(self._history_new_password, 1); auth_row2.addWidget(self._history_new_password_confirm, 1)
    auth.addLayout(auth_row2)
    auth_btns = QHBoxLayout()
    for txt, fn in [
        ('Uzlikt / nomainīt paroli', _v62_history_set_password),
        ('Noņemt paroli', _v62_history_remove_password),
        ('Izvēlēties backup mapi', _v62_history_choose_folder),
    ]:
        b = QPushButton(txt); b.clicked.connect(lambda *_a, f=fn: f(self)); auth_btns.addWidget(b)
    auth_btns.addStretch(1)
    auth.addLayout(auth_btns)
    layout.addWidget(self._history_auth_frame)

    self._history_content_frame = QFrame(page)
    self._history_content_frame.setStyleSheet('QFrame{background:#0B1220;border:1px solid #233048;border-radius:14px;} QLabel{color:#E5E7EB;} QListWidget{border:1px solid #233048;border-radius:12px;padding:6px;}')
    content = QVBoxLayout(self._history_content_frame)
    info_row = QHBoxLayout()
    self._history_folder_label = QLabel('')
    self._history_folder_label.setWordWrap(True)
    self._history_count_label = QLabel('Ieraksti: 0')
    info_row.addWidget(self._history_folder_label, 1)
    info_row.addWidget(self._history_count_label)
    content.addLayout(info_row)
    self.history_list = QListWidget()
    self.history_list.itemDoubleClicked.connect(lambda *_: _v62_history_load_selected(self))
    content.addWidget(self.history_list, 1)
    btns = QHBoxLayout()
    for txt, fn in [
        ('Ielādēt izvēlēto', _v62_history_load_selected),
        ('Atjaunot', _v62_refresh_history_list),
        ('Atvērt mapi', _v62_history_open_folder),
        ('Notīrīt JSON backup', _v62_history_clear),
        ('Aizslēgt', lambda s: (setattr(s, '_history_access_granted', False), _v62_apply_history_auth_ui(s))),
    ]:
        b = QPushButton(txt)
        if txt == 'Aizslēgt':
            b.clicked.connect(lambda *_: (setattr(self, '_history_access_granted', False), _v62_apply_history_auth_ui(self)))
        else:
            b.clicked.connect(lambda *_a, f=fn: f(self))
        btns.addWidget(b)
    btns.addStretch(1)
    content.addLayout(btns)
    layout.addWidget(self._history_content_frame, 1)

    _v62_apply_history_auth_ui(self)


def _v62_save_project_payload_enhancements(path, self):
    path = _coerce_path(path)
    if not path or not os.path.exists(path):
        return
    payload = _v62_safe_load_json(path, {}) or {}
    if not isinstance(payload, dict):
        return
    payload['optional_dates'] = _v62_collect_optional_dates(self)
    bundle = _v62_collect_custom_bundle(self)
    payload['custom_fields_bundle'] = bundle
    payload['custom_fields'] = bundle.get('values', {})
    _v62_safe_save_json(path, payload)


def _v62_restore_project_payload_enhancements(path, self):
    path = _coerce_path(path)
    if not path or not os.path.exists(path):
        return
    payload = _v62_safe_load_json(path, {}) or {}
    if not isinstance(payload, dict):
        return
    _v62_apply_optional_dates(self, payload.get('optional_dates'))
    bundle = payload.get('custom_fields_bundle')
    if not bundle:
        bundle = {'definitions': getattr(self, '_custom_field_definitions', {'pie': [], 'nod': [], 'rek': []}), 'values': payload.get('custom_fields', {}) or {}}
    _v62_apply_custom_bundle(self, bundle)
    _v62_save_ui_state(self)


def _v62_no_popup_history_require_access(self):
    return True


def _v62_no_popup_on_tab_changed(self, index):
    try:
        if self.tabs.tabText(index) == 'Dokumentu vēsture':
            _v62_apply_history_auth_ui(self)
    except Exception:
        pass


def _v62_after_init(self):
    try:
        _v60_after_init(self)
    except Exception:
        pass
    try:
        _v62_bind_live_state(self)
    except Exception:
        pass
    try:
        _v62_apply_optional_dates(self)
    except Exception:
        pass
    try:
        _v62_rebuild_history_tab(self)
    except Exception as e:
        print(f'v62 history rebuild failed: {e}')
    try:
        state = _v62_safe_load_json(V62_UI_STATE_FILE, {}) or {}
        _v62_apply_custom_bundle(self, state.get('custom_fields') or {})
    except Exception:
        pass


def _v62_close_wrapper(old_func):
    def wrapper(self, event):
        try:
            _v62_save_ui_state(self)
        except Exception:
            pass
        return old_func(self, event)
    return wrapper


def _v62_savakt_wrapper(old_func):
    def wrapper(self):
        d = old_func(self)
        try:
            bundle = _v62_collect_custom_bundle(self)
            setattr(d, 'custom_fields', bundle.get('values', {}))
        except Exception:
            pass
        try:
            setattr(d, 'ieklaut_izpildes_terminu', bool(getattr(self, 'ck_ieklaut_izpildes_terminu').isChecked()))
            setattr(d, 'ieklaut_pienemsanas_datumu', bool(getattr(self, 'ck_ieklaut_pienemsanas_datumu').isChecked()))
            setattr(d, 'ieklaut_nodosanas_datumu', bool(getattr(self, 'ck_ieklaut_nodosanas_datumu').isChecked()))
        except Exception:
            pass
        return d
    return wrapper


def _v62_save_project_wrapper(old_func):
    def wrapper(self):
        res = old_func(self)
        try:
            _v62_save_ui_state(self)
            _v62_save_project_payload_enhancements(getattr(self, '_ceļš_projekts', None), self)
            if getattr(self, '_ceļš_projekts', None):
                _v62_history_backup_snapshot(self, getattr(self, '_ceļš_projekts', None))
                _v62_refresh_history_list(self)
        except Exception as e:
            print(f'v62 save project wrapper failed: {e}')
        return res
    return wrapper


def _v62_load_project_wrapper(old_func):
    def wrapper(self, path=None):
        res = old_func(self, path)
        try:
            _v62_restore_project_payload_enhancements(path or getattr(self, '_ceļš_projekts', None), self)
        except Exception as e:
            print(f'v62 load project wrapper failed: {e}')
        return res
    return wrapper


def _v62_save_template_wrapper(old_func):
    def wrapper(self):
        res = old_func(self)
        try:
            item = self.sablonu_list.currentItem() if hasattr(self, 'sablonu_list') else None
            current_name = item.text() if item else ''
            current_name = current_name.replace(' [Aizsargāts]', '').replace(' [Kļūda]', '').strip()
            path = self._resolve_template_path(current_name) if current_name else None
            _v62_save_project_payload_enhancements(path, self)
            _v62_save_ui_state(self)
        except Exception as e:
            print(f'v62 save template wrapper failed: {e}')
        return res
    return wrapper


def _v62_load_template_wrapper(old_func):
    def wrapper(self, *args, **kwargs):
        res = old_func(self, *args, **kwargs)
        try:
            item = self.sablonu_list.currentItem() if hasattr(self, 'sablonu_list') else None
            current_name = item.text() if item else ''
            current_name = current_name.replace(' [Aizsargāts]', '').replace(' [Kļūda]', '').strip()
            path = self._resolve_template_path(current_name) if current_name else None
            _v62_restore_project_payload_enhancements(path, self)
        except Exception as e:
            print(f'v62 load template wrapper failed: {e}')
        return res
    return wrapper


def _v62_add_to_history(self, file_path: str):
    try:
        if str(file_path or '').lower().endswith('.json') and os.path.exists(str(file_path)):
            _v62_history_backup_snapshot(self, str(file_path))
        _v62_refresh_history_list(self)
    except Exception as e:
        print(f'v62 add to history failed: {e}')


def _v62_record_generated_document(self, pdf_path: str, json_path: str):
    try:
        if json_path and os.path.exists(json_path):
            _v62_save_project_payload_enhancements(json_path, self)
            _v62_history_backup_snapshot(self, json_path)
        _v62_refresh_history_list(self)
    except Exception as e:
        print(f'v62 record generated document failed: {e}')


# apply v62 overrides
_v54_history_require_access = _v62_no_popup_history_require_access
_v54_on_tab_changed = _v62_no_popup_on_tab_changed
_v54_after_init = _v62_after_init
AktaLogs.closeEvent = _v62_close_wrapper(AktaLogs.closeEvent)
AktaLogs.savākt_datus = _v62_savakt_wrapper(AktaLogs.savākt_datus)
AktaLogs.saglabat_projektu = _v62_save_project_wrapper(AktaLogs.saglabat_projektu)
AktaLogs.ieladet_projektu = _v62_load_project_wrapper(AktaLogs.ieladet_projektu)
AktaLogs.saglabat_ka_sablonu = _v62_save_template_wrapper(AktaLogs.saglabat_ka_sablonu)
try:
    AktaLogs._load_selected_template = _v62_load_template_wrapper(AktaLogs._load_selected_template)
except Exception:
    pass
AktaLogs._add_to_history = _v62_add_to_history
AktaLogs._record_generated_document = _v62_record_generated_document
AktaLogs._load_history = lambda self: setattr(self, 'history', [{'json': p} for p in _v62_history_items()])
AktaLogs._save_history = lambda self: None
AktaLogs._update_history_list = _v62_refresh_history_list
AktaLogs._load_history_entry = lambda self, item: _v62_history_load_selected(self)
AktaLogs._clear_history = _v62_history_clear
AktaLogs._open_document_folder = _v62_history_open_folder



# v63 visual cleanup for Dokumentu vēsture tab

def _v63_history_details_html(path: str):
    data = _v62_safe_load_json(path, {}) or {}
    meta = data.get('_history_meta') if isinstance(data, dict) else {}
    if not isinstance(meta, dict):
        meta = {}
    def g(*keys):
        for k in keys:
            v = meta.get(k) if isinstance(meta, dict) else None
            if v in (None, '') and isinstance(data, dict):
                v = data.get(k)
            if v not in (None, ''):
                return str(v)
        return ''
    akta_nr = g('akta_nr')
    title = g('title', 'virsraksts', 'nosaukums')
    datums = g('datums')
    saved_at = g('saved_at').replace('T', ' ')
    parties = []
    if isinstance(data, dict):
        for key in ('nodevejs', 'pieņēmējs', 'pienemejs', 'rekviziti'):
            val = data.get(key)
            if isinstance(val, dict):
                name = val.get('vards') or val.get('nosaukums') or val.get('uznemums') or val.get('uzņēmums')
                if name:
                    parties.append(str(name))
    parties = [x for i, x in enumerate(parties) if x and x not in parties[:i]]
    rows = []
    for label, value in [
        ('Akta Nr.', akta_nr),
        ('Nosaukums', title),
        ('Datums', datums),
        ('Saglabāts', saved_at),
        ('Fails', os.path.basename(path)),
        ('Puses', ' / '.join(parties) if parties else ''),
    ]:
        if value:
            rows.append(f'<tr><td style="padding:4px 10px 4px 0;color:#94A3B8;white-space:nowrap;">{label}</td><td style="padding:4px 0;color:#E5E7EB;">{html.escape(str(value))}</td></tr>')
    extra = ''
    if isinstance(data, dict):
        positions = data.get('pozicijas') or []
        if isinstance(positions, list):
            extra = f'<p style="color:#94A3B8;margin-top:10px;">Pozīciju skaits: <span style="color:#E5E7EB;">{len(positions)}</span></p>'
    return (
        '<div style="font-family:Segoe UI,Arial,sans-serif;font-size:13px;">'
        '<div style="font-size:15px;font-weight:700;color:#F8FAFC;margin-bottom:8px;">Dokumenta informācija</div>'
        f'<table cellspacing="0" cellpadding="0">{"".join(rows)}</table>{extra}'
        '</div>'
    )


def _v63_history_selection_changed(self):
    item = self.history_list.currentItem() if hasattr(self, 'history_list') else None
    path = _coerce_path(item.data(Qt.UserRole)) if item else ''
    if hasattr(self, '_history_preview_browser'):
        if path and os.path.exists(path):
            try:
                self._history_preview_browser.setHtml(_v63_history_details_html(path))
            except Exception:
                self._history_preview_browser.setPlainText(path)
        else:
            self._history_preview_browser.setHtml('<div style="color:#94A3B8;">Izvēlieties ierakstu, lai redzētu detaļas.</div>')


def _v63_apply_history_auth_ui(self):
    has_pw = _v62_history_has_password()
    locked = has_pw and not bool(getattr(self, '_history_access_granted', False))
    if hasattr(self, '_history_auth_state'):
        if has_pw:
            self._history_auth_state.setText('Aizsargāts')
            self._history_auth_state.setStyleSheet('padding:4px 10px;background:#3F1D1D;border:1px solid #7F1D1D;border-radius:999px;color:#FECACA;font-weight:700;')
        else:
            self._history_auth_state.setText('Bez paroles')
            self._history_auth_state.setStyleSheet('padding:4px 10px;background:#0F2C22;border:1px solid #14532D;border-radius:999px;color:#BBF7D0;font-weight:700;')
    if hasattr(self, '_history_unlock_box'):
        self._history_unlock_box.setVisible(locked)
    if hasattr(self, '_history_content_wrap'):
        self._history_content_wrap.setVisible(not locked)
    if hasattr(self, '_history_hint_label'):
        self._history_hint_label.setText('Ievadiet paroli zemāk, lai atvērtu dokumentu vēsturi.' if locked else 'Šeit redzami visi JSON backup dokumenti no izvēlētās mapes.')
    if hasattr(self, '_history_folder_label'):
        self._history_folder_label.setText(f'Backup mape: {_v62_history_folder()}')
    if not locked:
        _v62_refresh_history_list(self)
        _v63_history_selection_changed(self)


def _v63_rebuild_history_tab(self):
    idx = -1
    for i in range(self.tabs.count()):
        if self.tabs.tabText(i) == 'Dokumentu vēsture':
            idx = i
            break
    if idx < 0:
        return
    page = self.tabs.widget(idx)
    layout = page.layout() or QVBoxLayout(page)
    _v62_clear_layout(layout)
    layout.setContentsMargins(10, 10, 10, 10)
    layout.setSpacing(10)

    root = QFrame(page)
    root.setObjectName('historyRoot')
    root.setStyleSheet(
        'QFrame#historyRoot{background:transparent;border:none;}'
        'QFrame.historyPanel{background:#0B1220;border:1px solid #22314A;border-radius:12px;}'
        'QLabel{color:#E5E7EB;background:transparent;border:none;}'
        'QPushButton{padding:8px 12px;min-height:34px;}'
        'QLineEdit{min-height:34px;padding:6px 10px;}'
        'QListWidget{background:#0F172A;border:1px solid #233048;border-radius:10px;padding:4px;}'
        'QListWidget::item{padding:8px;border-radius:8px;}'
        'QListWidget::item:selected{background:#1D4ED8;color:white;}'
        'QTextBrowser{background:#0F172A;border:1px solid #233048;border-radius:10px;padding:10px;}'
    )
    root_l = QVBoxLayout(root)
    root_l.setContentsMargins(0, 0, 0, 0)
    root_l.setSpacing(10)

    top = QFrame(); top.setProperty('class', 'historyPanel')
    top_l = QHBoxLayout(top); top_l.setContentsMargins(14, 12, 14, 12); top_l.setSpacing(12)
    info_col = QVBoxLayout(); info_col.setSpacing(2)
    title = QLabel('Dokumentu vēsture')
    title.setStyleSheet('font-size:22px;font-weight:800;color:#F8FAFC;')
    info_col.addWidget(title)
    self._history_hint_label = QLabel('Šeit redzami visi JSON backup dokumenti no izvēlētās mapes.')
    self._history_hint_label.setStyleSheet('font-size:12px;color:#94A3B8;')
    self._history_hint_label.setWordWrap(True)
    info_col.addWidget(self._history_hint_label)
    top_l.addLayout(info_col, 1)
    self._history_auth_state = QLabel('')
    top_l.addWidget(self._history_auth_state, 0, Qt.AlignRight | Qt.AlignVCenter)
    root_l.addWidget(top)

    self._history_unlock_box = QFrame(); self._history_unlock_box.setProperty('class', 'historyPanel')
    unlock_l = QVBoxLayout(self._history_unlock_box); unlock_l.setContentsMargins(14, 12, 14, 12); unlock_l.setSpacing(10)
    unlock_title = QLabel('Atvērt aizsargāto vēsturi')
    unlock_title.setStyleSheet('font-size:14px;font-weight:700;color:#F8FAFC;')
    unlock_l.addWidget(unlock_title)
    unlock_row = QHBoxLayout(); unlock_row.setSpacing(8)
    self._history_password_input = QLineEdit(); self._history_password_input.setEchoMode(QLineEdit.Password); self._history_password_input.setPlaceholderText('Ievadiet paroli')
    self._history_password_input.returnPressed.connect(lambda: _v62_history_unlock(self))
    unlock_row.addWidget(self._history_password_input, 1)
    btn_unlock = QPushButton('Atvērt')
    btn_unlock.clicked.connect(lambda *_: _v62_history_unlock(self))
    unlock_row.addWidget(btn_unlock)
    unlock_l.addLayout(unlock_row)
    root_l.addWidget(self._history_unlock_box)

    settings = QFrame(); settings.setProperty('class', 'historyPanel')
    sett_l = QVBoxLayout(settings); sett_l.setContentsMargins(14, 12, 14, 12); sett_l.setSpacing(10)
    sett_head = QHBoxLayout()
    sh = QLabel('Piekļuve un iestatījumi'); sh.setStyleSheet('font-size:14px;font-weight:700;color:#F8FAFC;')
    sett_head.addWidget(sh)
    sett_head.addStretch(1)
    self._history_folder_label = QLabel('')
    self._history_folder_label.setStyleSheet('font-size:12px;color:#94A3B8;')
    sett_head.addWidget(self._history_folder_label)
    sett_l.addLayout(sett_head)
    pw_row = QHBoxLayout(); pw_row.setSpacing(8)
    self._history_new_password = QLineEdit(); self._history_new_password.setEchoMode(QLineEdit.Password); self._history_new_password.setPlaceholderText('Jauna parole')
    self._history_new_password_confirm = QLineEdit(); self._history_new_password_confirm.setEchoMode(QLineEdit.Password); self._history_new_password_confirm.setPlaceholderText('Atkārtojiet jauno paroli')
    pw_row.addWidget(self._history_new_password, 1)
    pw_row.addWidget(self._history_new_password_confirm, 1)
    sett_l.addLayout(pw_row)
    btn_row = QHBoxLayout(); btn_row.setSpacing(8)
    for txt, fn in [
        ('Uzlikt / nomainīt paroli', _v62_history_set_password),
        ('Noņemt paroli', _v62_history_remove_password),
        ('Izvēlēties backup mapi', _v62_history_choose_folder),
    ]:
        b = QPushButton(txt)
        b.clicked.connect(lambda *_a, f=fn: f(self))
        btn_row.addWidget(b)
    btn_row.addStretch(1)
    sett_l.addLayout(btn_row)
    root_l.addWidget(settings)

    self._history_content_wrap = QFrame(); self._history_content_wrap.setProperty('class', 'historyPanel')
    content_l = QVBoxLayout(self._history_content_wrap); content_l.setContentsMargins(14, 12, 14, 12); content_l.setSpacing(10)
    toolbar = QHBoxLayout(); toolbar.setSpacing(8)
    left_head = QLabel('Saglabātie dokumenti'); left_head.setStyleSheet('font-size:14px;font-weight:700;color:#F8FAFC;')
    toolbar.addWidget(left_head)
    self._history_count_label = QLabel('Ieraksti: 0')
    self._history_count_label.setStyleSheet('font-size:12px;color:#94A3B8;')
    toolbar.addWidget(self._history_count_label)
    toolbar.addStretch(1)
    for txt, fn in [
        ('Atjaunot', _v62_refresh_history_list),
        ('Ielādēt izvēlēto', _v62_history_load_selected),
        ('Atvērt mapi', _v62_history_open_folder),
        ('Notīrīt JSON backup', _v62_history_clear),
        ('Aizslēgt', lambda s: (setattr(s, '_history_access_granted', False), _v63_apply_history_auth_ui(s))),
    ]:
        b = QPushButton(txt)
        if txt == 'Aizslēgt':
            b.clicked.connect(lambda *_: (setattr(self, '_history_access_granted', False), _v63_apply_history_auth_ui(self)))
        else:
            b.clicked.connect(lambda *_a, f=fn: f(self))
        toolbar.addWidget(b)
    content_l.addLayout(toolbar)

    splitter = QSplitter(Qt.Horizontal)
    splitter.setChildrenCollapsible(False)
    left = QWidget(); left_l = QVBoxLayout(left); left_l.setContentsMargins(0,0,0,0); left_l.setSpacing(8)
    self.history_list = QListWidget()
    self.history_list.itemDoubleClicked.connect(lambda *_: _v62_history_load_selected(self))
    self.history_list.currentItemChanged.connect(lambda *_: _v63_history_selection_changed(self))
    left_l.addWidget(self.history_list, 1)
    right = QWidget(); right_l = QVBoxLayout(right); right_l.setContentsMargins(0,0,0,0); right_l.setSpacing(8)
    right_head = QLabel('Dokumenta detaļas'); right_head.setStyleSheet('font-size:14px;font-weight:700;color:#F8FAFC;')
    right_l.addWidget(right_head)
    self._history_preview_browser = QTextBrowser()
    right_l.addWidget(self._history_preview_browser, 1)
    splitter.addWidget(left)
    splitter.addWidget(right)
    splitter.setSizes([560, 300])
    content_l.addWidget(splitter, 1)
    root_l.addWidget(self._history_content_wrap, 1)

    layout.addWidget(root, 1)
    _v63_apply_history_auth_ui(self)


def _v63_after_init(self):
    try:
        _v62_after_init(self)
    except Exception:
        pass
    try:
        _v63_rebuild_history_tab(self)
    except Exception as e:
        print(f'v63 history rebuild failed: {e}')


_v62_apply_history_auth_ui = _v63_apply_history_auth_ui
_v62_rebuild_history_tab = _v63_rebuild_history_tab
_v54_after_init = _v63_after_init



# ==============================
# v64: Grāmatvedības sistēmu integrācija (Jumis + Generic REST)
# Pievienota bez esošās biznesa loģikas pārrakstīšanas – tikai ar wrapperiem un jaunām metodēm.
# ==============================
from xml.sax.saxutils import escape as _xml_escape

ACCOUNTING_SETTINGS_KEY = 'accounting_integrations_v64'
JUMIS_IMPORT_URL = 'https://vadiba.mansjumis.lv/cloudapi/JumisImportExportService.ImportExportService.svc/import'
JUMIS_EXPORT_URL = 'https://vadiba.mansjumis.lv/cloudapi/JumisImportExportService.ImportExportService.svc/export'
JUMIS_DEFAULT_API_KEY = '2BFC1C2B748D4C04BB0ECABA7FBFB1A6'


def _acc_safe_str(v, default=''):
    try:
        if v is None:
            return default
        return str(v)
    except Exception:
        return default


def _acc_clean(v):
    return _acc_safe_str(v, '').strip()


def _acc_to_amount(v, default='0.00'):
    try:
        if isinstance(v, Decimal):
            return f'{v.quantize(Decimal("0.01"))}'
        s = _acc_clean(v).replace(' ', '').replace(',', '.')
        if not s:
            return default
        return f'{Decimal(s).quantize(Decimal("0.01"))}'
    except Exception:
        return default


def _acc_to_qty(v, default='0'):
    try:
        s = _acc_clean(v).replace(' ', '').replace(',', '.')
        if not s:
            return default
        d = Decimal(s)
        s2 = format(d.normalize(), 'f')
        return s2.rstrip('0').rstrip('.') if '.' in s2 else s2
    except Exception:
        return default


def _acc_xml(tag, value=None, allow_empty=False):
    if value is None:
        return f'<{tag} />' if allow_empty else ''
    s = _acc_safe_str(value, '')
    if s == '' and not allow_empty:
        return ''
    return f'<{tag}>{_xml_escape(s)}</{tag}>'


def _acc_iso_date(value):
    s = _acc_clean(value)
    if not s:
        return ''
    for fmt in ('%Y-%m-%d', '%d.%m.%Y', '%d/%m/%Y', '%Y-%m-%d %H:%M:%S'):
        try:
            return datetime.strptime(s, fmt).strftime('%Y-%m-%d')
        except Exception:
            pass
    return s[:10]


def _acc_now_ts():
    return datetime.now().strftime('%Y-%m-%d %H:%M:%S')


@dataclass
class AccountingIntegrationConfig:
    enabled: bool = False
    system_type: str = 'jumis'          # jumis | generic_rest
    profile_name: str = 'Noklusējuma profils'

    # Jumis
    jumis_username: str = ''
    jumis_password: str = ''
    jumis_database: str = ''
    jumis_api_key: str = JUMIS_DEFAULT_API_KEY
    jumis_import_url: str = JUMIS_IMPORT_URL
    jumis_export_url: str = JUMIS_EXPORT_URL

    # Generic REST
    generic_base_url: str = ''
    generic_auth_type: str = 'bearer'   # none | bearer | basic | api_key
    generic_username: str = ''
    generic_password: str = ''
    generic_token: str = ''
    generic_api_key_header: str = 'X-API-Key'
    generic_api_key_value: str = ''
    generic_timeout_sec: int = 45
    generic_verify_ssl: bool = True
    generic_test_endpoint: str = '/health'
    generic_partners_endpoint: str = '/partners'
    generic_products_endpoint: str = '/products'
    generic_documents_endpoint: str = '/documents'
    generic_inventory_documents_endpoint: str = '/inventory-documents'

    # Auto-sync uzvedība
    auto_sync_on_pdf_generate: bool = False
    auto_sync_on_docx_generate: bool = False
    auto_sync_on_project_save: bool = False

    # Ko sinhronizēt komplektajā darbībā
    sync_partners_with_document: bool = True
    sync_products_with_inventory: bool = True
    sync_store_doc_with_document: bool = True
    sync_financial_doc_with_document: bool = True


@dataclass
class AccountingSyncResult:
    ok: bool
    operation: str
    system: str
    request_payload: str = ''
    response_status: int = 0
    response_text: str = ''
    extra: dict = field(default_factory=dict)


class AccountingIntegrationError(Exception):
    pass


class BaseAccountingConnector:
    system_name = 'Base'

    def __init__(self, config: AccountingIntegrationConfig):
        self.config = config
        self.session = requests.Session()

    def test_connection(self) -> AccountingSyncResult:
        raise NotImplementedError

    def upsert_partner(self, persona: Persona) -> AccountingSyncResult:
        raise NotImplementedError

    def upsert_product(self, item: NoliktavasPrece) -> AccountingSyncResult:
        raise NotImplementedError

    def export_store_document(self, akta_dati: AktaDati) -> AccountingSyncResult:
        raise NotImplementedError

    def export_financial_document(self, akta_dati: AktaDati) -> AccountingSyncResult:
        raise NotImplementedError

    def fetch_partners(self, limit: int = 100) -> AccountingSyncResult:
        raise NotImplementedError

    def fetch_products(self, limit: int = 100) -> AccountingSyncResult:
        raise NotImplementedError


class JumisConnector(BaseAccountingConnector):
    system_name = 'Jumis'

    def _credentials_payload(self, xmlrequest: str) -> dict:
        cfg = self.config
        username = _acc_clean(cfg.jumis_username)
        password = _acc_clean(cfg.jumis_password)
        database = _acc_clean(cfg.jumis_database)
        if not username or not password or not database:
            raise AccountingIntegrationError('Jumis pieslēgumam obligāti jānorāda lietotājvārds, speciālā parole un datubāzes nosaukums.')
        return {
            'username': username,
            'password': password,
            'database': database,
            'apikey': _acc_clean(cfg.jumis_api_key) or JUMIS_DEFAULT_API_KEY,
            'XMLrequest': xmlrequest,
        }

    def _post_import(self, xmlrequest: str, operation: str) -> AccountingSyncResult:
        payload = self._credentials_payload(xmlrequest)
        r = self.session.post(
            _acc_clean(self.config.jumis_import_url) or JUMIS_IMPORT_URL,
            json=payload,
            timeout=max(int(getattr(self.config, 'jumis_timeout_sec', 45) or 45), 10),
        )
        return AccountingSyncResult(
            ok=r.ok,
            operation=operation,
            system=self.system_name,
            request_payload=xmlrequest,
            response_status=r.status_code,
            response_text=r.text,
        )

    def _post_export(self, xmlrequest: str, operation: str) -> AccountingSyncResult:
        payload = self._credentials_payload(xmlrequest)
        r = self.session.post(
            _acc_clean(self.config.jumis_export_url) or JUMIS_EXPORT_URL,
            json=payload,
            timeout=max(int(getattr(self.config, 'jumis_timeout_sec', 45) or 45), 10),
        )
        return AccountingSyncResult(
            ok=r.ok,
            operation=operation,
            system=self.system_name,
            request_payload=xmlrequest,
            response_status=r.status_code,
            response_text=r.text,
        )

    def test_connection(self) -> AccountingSyncResult:
        xmlrequest = ('<?xml version="1.0" ?>'
                      '<dataroot>'
                      '<tjDocument Version="TJ5.5.101"/>'
                      '<tjRequest Name="Partner" Operation="Read" Version="TJ7.0.112" Structure="Tree">'
                      '<tjFields><Field Name="PartnerName"/></tjFields>'
                      '</tjRequest>'
                      '</dataroot>')
        return self._post_export(xmlrequest, 'test_connection')

    def _partner_xml(self, persona: Persona) -> str:
        p = persona or Persona()
        reg = _acc_clean(getattr(p, 'reģ_nr', ''))
        vat = reg
        country = 'LV'
        name = _acc_clean(getattr(p, 'nosaukums', '')) or _acc_clean(getattr(p, 'kontaktpersona', '')) or 'Nezināms partneris'
        return (
            '<?xml version="1.0" encoding="utf-8" ?>'
            '<dataroot>'
            '<tjDocument Version="TJ5.5.101"/>'
            '<tjResponse Name="Partner" Operation="Insert" Version="TJ7.0.112" Structure="Tree">'
            '<Partner>'
            f'{_acc_xml("PartnerKindName", "Juridiska persona")}'
            f'{_acc_xml("PartnerName", name)}'
            f'{_acc_xml("PartnerRegistrationNo", reg)}'
            f'{_acc_xml("PartnerPhone", _acc_clean(getattr(p, "tālrunis", "")))}'
            f'{_acc_xml("PartnerEMail", _acc_clean(getattr(p, "epasts", "")))}'
            f'{_acc_xml("PartnerWWW", _acc_clean(getattr(p, "web_lapa", "")))}'
            + (
                '<PartnerAddress>'
                '<AddressDefaultNoticeID>1</AddressDefaultNoticeID>'
                f'{_acc_xml("AddressStreet", _acc_clean(getattr(p, "adrese", "")))}'
                f'{_acc_xml("AddressCountryCode", country)}'
                '</PartnerAddress>'
                if _acc_clean(getattr(p, 'adrese', '')) else ''
            )
            + (
                '<PartnerVatNo>'
                f'{_acc_xml("VatNo", vat if vat.upper().startswith("LV") else f"LV{vat}" if vat else "")}'
                '<VatNoCountryCode>LV</VatNoCountryCode>'
                '<VatNoDefaultNoticeID>1</VatNoDefaultNoticeID>'
                '</PartnerVatNo>'
                if vat else ''
            )
            + '</Partner></tjResponse></dataroot>'
        )

    def upsert_partner(self, persona: Persona) -> AccountingSyncResult:
        return self._post_import(self._partner_xml(persona), 'upsert_partner')

    def _product_xml(self, item: NoliktavasPrece) -> str:
        it = item or NoliktavasPrece()
        code = _acc_clean(getattr(it, 'sku', '')) or _acc_clean(getattr(it, 'inventory_id', '')) or f'ITEM-{secrets.token_hex(4)}'
        price = _acc_to_amount(getattr(it, 'cena', '0'))
        return (
            '<?xml version="1.0" encoding="utf-8" ?>'
            '<dataroot>'
            '<tjDocument Version="TJ5.5.101"/>'
            '<tjResponse Name="Product" Operation="Insert" Version="TJ5.5.101" Structure="Tree">'
            '<Product>'
            f'{_acc_xml("ProductCode", code)}'
            f'{_acc_xml("ProductName", _acc_clean(getattr(it, "nosaukums", "")) or code)}'
            f'{_acc_xml("ProductBarCode", _acc_clean(getattr(it, "svitrkods", "")))}'
            f'{_acc_xml("ProductUnit", _acc_clean(getattr(it, "vieniba", "")) or "gab.")}'
            f'{_acc_xml("ProductPrice", price)}'
            f'{_acc_xml("ProductPurchasePrice", price)}'
            f'{_acc_xml("ProductVATRate", _acc_to_amount(getattr(it, "pvn_likme", "21"), default="21.00"))}'
            f'{_acc_xml("ProductGroupName", _acc_clean(getattr(it, "kategorija", "")))}'
            f'{_acc_xml("ProductDescription", _acc_clean(getattr(it, "piezimes", "")))}'
            '</Product></tjResponse></dataroot>'
        )

    def upsert_product(self, item: NoliktavasPrece) -> AccountingSyncResult:
        return self._post_import(self._product_xml(item), 'upsert_product')

    def _store_doc_xml(self, akta_dati: AktaDati) -> str:
        d = akta_dati or AktaDati()
        lines = []
        for idx, p in enumerate(list(getattr(d, 'pozīcijas', []) or []), start=1):
            code = _acc_clean(getattr(p, 'seriālais_nr', '')) or f'LINE-{idx}'
            name = _acc_clean(getattr(p, 'apraksts', '')) or code
            qty = _acc_to_qty(getattr(p, 'daudzums', '1'))
            price = _acc_to_amount(getattr(p, 'cena', '0'))
            total = _acc_to_amount(getattr(p, 'summa', '0'))
            line = (
                '<StoreDocRow>'
                f'{_acc_xml("RowNo", idx)}'
                f'{_acc_xml("ProductCode", code)}'
                f'{_acc_xml("ProductName", name)}'
                f'{_acc_xml("RowDescription", name)}'
                f'{_acc_xml("RowUnit", _acc_clean(getattr(p, "vienība", "")) or "gab.")}'
                f'{_acc_xml("RowQuantity", qty)}'
                f'{_acc_xml("RowPrice", price)}'
                f'{_acc_xml("RowAmount", total)}'
                '</StoreDocRow>'
            )
            lines.append(line)
        partner = getattr(d, 'pieņēmējs', None) or Persona()
        xmlrequest = (
            '<?xml version="1.0" encoding="utf-8" ?>'
            '<dataroot>'
            '<tjDocument Version="TJ5.5.101"/>'
            '<tjResponse Name="StoreDoc" Operation="Insert" Version="TJ7.0.109" Structure="Tree">'
            '<StoreDoc>'
            f'{_acc_xml("DocNo", _acc_clean(getattr(d, "akta_nr", "")) or f"DOC-{datetime.now().strftime("%Y%m%d%H%M%S")}")}'
            f'{_acc_xml("DocDate", _acc_iso_date(getattr(d, "datums", "")) or datetime.now().strftime("%Y-%m-%d"))}'
            f'{_acc_xml("DocDisbursementTerm", _acc_iso_date(getattr(d, "apmaksas_termins", "")))}'
            f'{_acc_xml("DocCurrencyCode", _acc_clean(getattr(d, "valūta", "")) or "EUR")}'
            f'{_acc_xml("DocComments", _acc_clean(getattr(d, "piezīmes", "")))}'
            f'{_acc_xml("PartnerName", _acc_clean(getattr(partner, "nosaukums", "")))}'
            f'{_acc_xml("PartnerRegistrationNo", _acc_clean(getattr(partner, "reģ_nr", "")))}'
            + ''.join(lines) +
            '</StoreDoc></tjResponse></dataroot>'
        )
        return xmlrequest

    def export_store_document(self, akta_dati: AktaDati) -> AccountingSyncResult:
        return self._post_import(self._store_doc_xml(akta_dati), 'export_store_document')

    def _financial_doc_xml(self, akta_dati: AktaDati) -> str:
        d = akta_dati or AktaDati()
        partner = getattr(d, 'pieņēmējs', None) or Persona()
        amount = _acc_to_amount(d.summa_ar_pvn() if getattr(d, 'iekļaut_pvn', False) else d.kopējā_summma())
        xmlrequest = (
            '<?xml version="1.0" encoding="utf-8" ?>'
            '<dataroot>'
            '<tjDocument Version="TJ5.5.101"/>'
            '<tjResponse Name="FinancialDoc" Operation="Insert" Version="TJ7.0.112" Structure="Tree">'
            '<FinancialDoc>'
            f'{_acc_xml("DocNo", _acc_clean(getattr(d, "akta_nr", "")) or f"FIN-{datetime.now().strftime("%Y%m%d%H%M%S")}")}'
            f'{_acc_xml("DocDate", _acc_iso_date(getattr(d, "datums", "")) or datetime.now().strftime("%Y-%m-%d"))}'
            f'{_acc_xml("DocDisbursementDate", _acc_iso_date(getattr(d, "apmaksas_termins", "")))}'
            f'{_acc_xml("DocCurrencyCode", _acc_clean(getattr(d, "valūta", "")) or "EUR")}'
            f'{_acc_xml("PartnerName", _acc_clean(getattr(partner, "nosaukums", "")))}'
            f'{_acc_xml("PartnerRegistrationNo", _acc_clean(getattr(partner, "reģ_nr", "")))}'
            f'{_acc_xml("DocAmount", amount)}'
            f'{_acc_xml("DocComments", _acc_clean(getattr(d, "piezīmes", "")))}'
            '</FinancialDoc></tjResponse></dataroot>'
        )
        return xmlrequest

    def export_financial_document(self, akta_dati: AktaDati) -> AccountingSyncResult:
        return self._post_import(self._financial_doc_xml(akta_dati), 'export_financial_document')

    def fetch_partners(self, limit: int = 100) -> AccountingSyncResult:
        xmlrequest = (
            '<?xml version="1.0" ?>'
            '<dataroot>'
            '<tjDocument Version="TJ5.5.101"/>'
            '<tjRequest Name="Partner" Operation="Read" Version="TJ7.0.112" Structure="Tree">'
            '<tjFields>'
            '<Field Name="PartnerName"/><Field Name="PartnerRegistrationNo"/><Field Name="PartnerEMail"/>'
            '<Field Name="PartnerPhone"/><Field Name="PartnerWWW"/>'
            '</tjFields>'
            '</tjRequest>'
            '</dataroot>'
        )
        res = self._post_export(xmlrequest, 'fetch_partners')
        res.extra = {'limit': limit}
        return res

    def fetch_products(self, limit: int = 100) -> AccountingSyncResult:
        xmlrequest = (
            '<?xml version="1.0" ?>'
            '<dataroot>'
            '<tjDocument Version="TJ5.5.101"/>'
            '<tjRequest Name="Product" Operation="Read" Version="TJ5.5.101" Structure="Tree">'
            '<tjFields>'
            '<Field Name="ProductCode"/><Field Name="ProductName"/><Field Name="ProductBarCode"/>'
            '<Field Name="ProductUnit"/><Field Name="ProductPrice"/>'
            '</tjFields>'
            '</tjRequest>'
            '</dataroot>'
        )
        res = self._post_export(xmlrequest, 'fetch_products')
        res.extra = {'limit': limit}
        return res


class GenericRestConnector(BaseAccountingConnector):
    system_name = 'Generic REST'

    def _base(self):
        base = _acc_clean(self.config.generic_base_url).rstrip('/')
        if not base:
            raise AccountingIntegrationError('Generic REST pieslēgumam jānorāda Base URL.')
        return base

    def _url(self, endpoint: str) -> str:
        ep = _acc_clean(endpoint)
        if ep.startswith('http://') or ep.startswith('https://'):
            return ep
        if not ep.startswith('/'):
            ep = '/' + ep
        return self._base() + ep

    def _headers(self):
        headers = {'Accept': 'application/json', 'Content-Type': 'application/json'}
        at = _acc_clean(self.config.generic_auth_type).lower()
        if at == 'bearer' and _acc_clean(self.config.generic_token):
            headers['Authorization'] = f'Bearer {_acc_clean(self.config.generic_token)}'
        elif at == 'api_key' and _acc_clean(self.config.generic_api_key_header):
            headers[_acc_clean(self.config.generic_api_key_header)] = _acc_clean(self.config.generic_api_key_value)
        return headers

    def _auth(self):
        at = _acc_clean(self.config.generic_auth_type).lower()
        if at == 'basic' and (_acc_clean(self.config.generic_username) or _acc_clean(self.config.generic_password)):
            return (_acc_clean(self.config.generic_username), _acc_clean(self.config.generic_password))
        return None

    def _post_json(self, endpoint: str, payload: dict, operation: str) -> AccountingSyncResult:
        r = self.session.post(
            self._url(endpoint),
            json=payload,
            headers=self._headers(),
            auth=self._auth(),
            timeout=max(int(self.config.generic_timeout_sec or 45), 5),
            verify=bool(self.config.generic_verify_ssl),
        )
        return AccountingSyncResult(
            ok=r.ok,
            operation=operation,
            system=self.system_name,
            request_payload=json.dumps(payload, ensure_ascii=False, indent=2),
            response_status=r.status_code,
            response_text=r.text,
        )

    def _get_json(self, endpoint: str, params: dict, operation: str) -> AccountingSyncResult:
        r = self.session.get(
            self._url(endpoint),
            params=params,
            headers=self._headers(),
            auth=self._auth(),
            timeout=max(int(self.config.generic_timeout_sec or 45), 5),
            verify=bool(self.config.generic_verify_ssl),
        )
        return AccountingSyncResult(
            ok=r.ok,
            operation=operation,
            system=self.system_name,
            request_payload=json.dumps(params, ensure_ascii=False, indent=2),
            response_status=r.status_code,
            response_text=r.text,
        )

    def test_connection(self) -> AccountingSyncResult:
        return self._get_json(self.config.generic_test_endpoint, {}, 'test_connection')

    def upsert_partner(self, persona: Persona) -> AccountingSyncResult:
        p = persona or Persona()
        payload = {
            'name': _acc_clean(getattr(p, 'nosaukums', '')),
            'registration_no': _acc_clean(getattr(p, 'reģ_nr', '')),
            'address': _acc_clean(getattr(p, 'adrese', '')),
            'contact_person': _acc_clean(getattr(p, 'kontaktpersona', '')),
            'position': _acc_clean(getattr(p, 'amats', '')),
            'phone': _acc_clean(getattr(p, 'tālrunis', '')),
            'email': _acc_clean(getattr(p, 'epasts', '')),
            'website': _acc_clean(getattr(p, 'web_lapa', '')),
            'bank_account': _acc_clean(getattr(p, 'bankas_konts', '')),
            'legal_status': _acc_clean(getattr(p, 'juridiskais_statuss', '')),
        }
        return self._post_json(self.config.generic_partners_endpoint, payload, 'upsert_partner')

    def upsert_product(self, item: NoliktavasPrece) -> AccountingSyncResult:
        it = item or NoliktavasPrece()
        payload = {
            'sku': _acc_clean(getattr(it, 'sku', '')),
            'inventory_id': _acc_clean(getattr(it, 'inventory_id', '')),
            'name': _acc_clean(getattr(it, 'nosaukums', '')),
            'unit': _acc_clean(getattr(it, 'vieniba', '')) or 'gab.',
            'price': _acc_to_amount(getattr(it, 'cena', '0')),
            'vat_rate': _acc_to_amount(getattr(it, 'pvn_likme', '21')),
            'stock': _acc_to_qty(getattr(it, 'atlikums', '0')),
            'barcode': _acc_clean(getattr(it, 'svitrkods', '')),
            'category': _acc_clean(getattr(it, 'kategorija', '')),
            'subcategory': _acc_clean(getattr(it, 'apakskategorija', '')),
            'notes': _acc_clean(getattr(it, 'piezimes', '')),
        }
        return self._post_json(self.config.generic_products_endpoint, payload, 'upsert_product')

    def export_store_document(self, akta_dati: AktaDati) -> AccountingSyncResult:
        d = akta_dati or AktaDati()
        payload = {
            'document_type': 'store_document',
            'document_no': _acc_clean(getattr(d, 'akta_nr', '')),
            'date': _acc_iso_date(getattr(d, 'datums', '')),
            'payment_due': _acc_iso_date(getattr(d, 'apmaksas_termins', '')),
            'currency': _acc_clean(getattr(d, 'valūta', '')) or 'EUR',
            'notes': _acc_clean(getattr(d, 'piezīmes', '')),
            'partner': {
                'name': _acc_clean(getattr(getattr(d, 'pieņēmējs', Persona()), 'nosaukums', '')),
                'registration_no': _acc_clean(getattr(getattr(d, 'pieņēmējs', Persona()), 'reģ_nr', '')),
            },
            'lines': [
                {
                    'description': _acc_clean(getattr(p, 'apraksts', '')),
                    'quantity': _acc_to_qty(getattr(p, 'daudzums', '0')),
                    'unit': _acc_clean(getattr(p, 'vienība', '')),
                    'price': _acc_to_amount(getattr(p, 'cena', '0')),
                    'total': _acc_to_amount(getattr(p, 'summa', '0')),
                    'serial_no': _acc_clean(getattr(p, 'seriālais_nr', '')),
                    'warranty': _acc_clean(getattr(p, 'garantija', '')),
                    'notes': _acc_clean(getattr(p, 'piezīmes_pozīcijai', '')),
                }
                for p in (getattr(d, 'pozīcijas', []) or [])
            ]
        }
        return self._post_json(self.config.generic_inventory_documents_endpoint or self.config.generic_documents_endpoint, payload, 'export_store_document')

    def export_financial_document(self, akta_dati: AktaDati) -> AccountingSyncResult:
        d = akta_dati or AktaDati()
        payload = {
            'document_type': 'financial_document',
            'document_no': _acc_clean(getattr(d, 'akta_nr', '')),
            'date': _acc_iso_date(getattr(d, 'datums', '')),
            'payment_due': _acc_iso_date(getattr(d, 'apmaksas_termins', '')),
            'currency': _acc_clean(getattr(d, 'valūta', '')) or 'EUR',
            'subtotal': _acc_to_amount(d.kopējā_summma()),
            'vat_amount': _acc_to_amount(d.pvn_summa()),
            'total': _acc_to_amount(d.summa_ar_pvn()),
            'vat_enabled': bool(getattr(d, 'iekļaut_pvn', False)),
            'partner': {
                'name': _acc_clean(getattr(getattr(d, 'pieņēmējs', Persona()), 'nosaukums', '')),
                'registration_no': _acc_clean(getattr(getattr(d, 'pieņēmējs', Persona()), 'reģ_nr', '')),
            },
            'notes': _acc_clean(getattr(d, 'piezīmes', '')),
        }
        return self._post_json(self.config.generic_documents_endpoint, payload, 'export_financial_document')

    def fetch_partners(self, limit: int = 100) -> AccountingSyncResult:
        return self._get_json(self.config.generic_partners_endpoint, {'limit': limit}, 'fetch_partners')

    def fetch_products(self, limit: int = 100) -> AccountingSyncResult:
        return self._get_json(self.config.generic_products_endpoint, {'limit': limit}, 'fetch_products')


class AccountingIntegrationDialog(QDialog):
    def __init__(self, parent, cfg: AccountingIntegrationConfig):
        super().__init__(parent)
        self.setWindowTitle('Grāmatvedības integrācijas iestatījumi')
        self.resize(760, 560)
        self.setMinimumSize(700, 480)
        self._cfg = cfg
        root = QVBoxLayout(self)
        root.setContentsMargins(10, 10, 10, 10)
        root.setSpacing(8)

        self.tabs = QTabWidget()
        root.addWidget(self.tabs, 1)

        # Vispārīgi
        tab_general = QWidget(); fg = QFormLayout(tab_general)
        self.ck_enabled = QCheckBox('Ieslēgt integrāciju')
        self.ck_enabled.setChecked(bool(cfg.enabled))
        fg.addRow(self.ck_enabled)
        self.in_profile_name = QLineEdit(cfg.profile_name)
        fg.addRow('Profila nosaukums', self.in_profile_name)
        self.cmb_system_type = QComboBox()
        self.cmb_system_type.addItem('Jumis', 'jumis')
        self.cmb_system_type.addItem('Generic REST / cita sistēma', 'generic_rest')
        ix = self.cmb_system_type.findData(cfg.system_type)
        self.cmb_system_type.setCurrentIndex(ix if ix >= 0 else 0)
        fg.addRow('Sistēma', self.cmb_system_type)
        self.tabs.addTab(tab_general, 'Vispārīgi')

        # Jumis
        tab_j = QWidget(); fj = QFormLayout(tab_j)
        self.in_j_username = QLineEdit(cfg.jumis_username)
        self.in_j_password = QLineEdit(cfg.jumis_password); self.in_j_password.setEchoMode(QLineEdit.Password)
        self.in_j_database = QLineEdit(cfg.jumis_database)
        self.in_j_apikey = QLineEdit(cfg.jumis_api_key or JUMIS_DEFAULT_API_KEY)
        self.in_j_import_url = QLineEdit(cfg.jumis_import_url or JUMIS_IMPORT_URL)
        self.in_j_export_url = QLineEdit(cfg.jumis_export_url or JUMIS_EXPORT_URL)
        fj.addRow('Lietotājvārds (e-pasts)', self.in_j_username)
        fj.addRow('Speciālā parole', self.in_j_password)
        fj.addRow('Datubāzes nosaukums', self.in_j_database)
        fj.addRow('API atslēga', self.in_j_apikey)
        fj.addRow('Import URL', self.in_j_import_url)
        fj.addRow('Export URL', self.in_j_export_url)
        info = QLabel('Jumis izmanto REST API ar JSON POST pieprasījumiem uz import/export servisu. Nepieciešama speciālā parole un datubāzes nosaukums.')
        info.setWordWrap(True)
        fj.addRow(info)
        self.tabs.addTab(tab_j, 'Jumis')

        # Generic REST
        tab_g = QWidget(); fr = QFormLayout(tab_g)
        self.in_g_base_url = QLineEdit(cfg.generic_base_url)
        self.cmb_g_auth = QComboBox()
        for label, value in [('Bez autentifikācijas', 'none'), ('Bearer', 'bearer'), ('Basic', 'basic'), ('API key', 'api_key')]:
            self.cmb_g_auth.addItem(label, value)
        ix = self.cmb_g_auth.findData(cfg.generic_auth_type)
        self.cmb_g_auth.setCurrentIndex(ix if ix >= 0 else 1)
        self.in_g_username = QLineEdit(cfg.generic_username)
        self.in_g_password = QLineEdit(cfg.generic_password); self.in_g_password.setEchoMode(QLineEdit.Password)
        self.in_g_token = QLineEdit(cfg.generic_token); self.in_g_token.setEchoMode(QLineEdit.Password)
        self.in_g_key_header = QLineEdit(cfg.generic_api_key_header)
        self.in_g_key_value = QLineEdit(cfg.generic_api_key_value); self.in_g_key_value.setEchoMode(QLineEdit.Password)
        self.sp_timeout = QSpinBox(); self.sp_timeout.setRange(5, 600); self.sp_timeout.setValue(int(cfg.generic_timeout_sec or 45))
        self.ck_verify_ssl = QCheckBox('Verificēt SSL')
        self.ck_verify_ssl.setChecked(bool(cfg.generic_verify_ssl))
        self.in_g_test = QLineEdit(cfg.generic_test_endpoint)
        self.in_g_partners = QLineEdit(cfg.generic_partners_endpoint)
        self.in_g_products = QLineEdit(cfg.generic_products_endpoint)
        self.in_g_docs = QLineEdit(cfg.generic_documents_endpoint)
        self.in_g_inv_docs = QLineEdit(cfg.generic_inventory_documents_endpoint)
        fr.addRow('Base URL', self.in_g_base_url)
        fr.addRow('Auth veids', self.cmb_g_auth)
        fr.addRow('Lietotājvārds', self.in_g_username)
        fr.addRow('Parole', self.in_g_password)
        fr.addRow('Bearer token', self.in_g_token)
        fr.addRow('API key header', self.in_g_key_header)
        fr.addRow('API key value', self.in_g_key_value)
        fr.addRow('Timeout (sek.)', self.sp_timeout)
        fr.addRow(self.ck_verify_ssl)
        fr.addRow('Test endpoint', self.in_g_test)
        fr.addRow('Partners endpoint', self.in_g_partners)
        fr.addRow('Products endpoint', self.in_g_products)
        fr.addRow('Documents endpoint', self.in_g_docs)
        fr.addRow('Inventory docs endpoint', self.in_g_inv_docs)
        self.tabs.addTab(tab_g, 'Generic REST')

        # Auto sync
        tab_auto = QWidget(); fa = QFormLayout(tab_auto)
        self.ck_auto_pdf = QCheckBox('Automātiski sinhronizēt pēc PDF ģenerēšanas')
        self.ck_auto_docx = QCheckBox('Automātiski sinhronizēt pēc DOCX ģenerēšanas')
        self.ck_auto_project = QCheckBox('Automātiski sinhronizēt pēc projekta saglabāšanas')
        self.ck_auto_pdf.setChecked(bool(cfg.auto_sync_on_pdf_generate))
        self.ck_auto_docx.setChecked(bool(cfg.auto_sync_on_docx_generate))
        self.ck_auto_project.setChecked(bool(cfg.auto_sync_on_project_save))
        self.ck_sync_partners = QCheckBox('Dokumenta sinhronizācijā sūtīt partnerus')
        self.ck_sync_products = QCheckBox('Pilnajā sinhronizācijā sūtīt noliktavas preces')
        self.ck_sync_store = QCheckBox('Dokumenta sinhronizācijā sūtīt noliktavas dokumentu')
        self.ck_sync_fin = QCheckBox('Dokumenta sinhronizācijā sūtīt finanšu dokumentu')
        self.ck_sync_partners.setChecked(bool(cfg.sync_partners_with_document))
        self.ck_sync_products.setChecked(bool(cfg.sync_products_with_inventory))
        self.ck_sync_store.setChecked(bool(cfg.sync_store_doc_with_document))
        self.ck_sync_fin.setChecked(bool(cfg.sync_financial_doc_with_document))
        for w in [self.ck_auto_pdf, self.ck_auto_docx, self.ck_auto_project, self.ck_sync_partners, self.ck_sync_products, self.ck_sync_store, self.ck_sync_fin]:
            fa.addRow(w)
        self.tabs.addTab(tab_auto, 'Automātika')

        self.out_info = QTextBrowser()
        self.out_info.setMaximumHeight(70)
        self.out_info.setMinimumHeight(54)
        self.out_info.setPlainText('Šeit vari konfigurēt Jumis vai citu REST grāmatvedības sistēmu. Iestatījumus saglabā lokāli settings.json failā.')
        root.addWidget(self.out_info, 0)

        btns = QDialogButtonBox(QDialogButtonBox.Save | QDialogButtonBox.Cancel)
        btns.accepted.connect(self.accept)
        btns.rejected.connect(self.reject)
        root.addWidget(btns)

    def get_config(self) -> AccountingIntegrationConfig:
        cfg = AccountingIntegrationConfig()
        cfg.enabled = self.ck_enabled.isChecked()
        cfg.profile_name = _acc_clean(self.in_profile_name.text()) or 'Noklusējuma profils'
        cfg.system_type = self.cmb_system_type.currentData() or 'jumis'

        cfg.jumis_username = _acc_clean(self.in_j_username.text())
        cfg.jumis_password = _acc_clean(self.in_j_password.text())
        cfg.jumis_database = _acc_clean(self.in_j_database.text())
        cfg.jumis_api_key = _acc_clean(self.in_j_apikey.text()) or JUMIS_DEFAULT_API_KEY
        cfg.jumis_import_url = _acc_clean(self.in_j_import_url.text()) or JUMIS_IMPORT_URL
        cfg.jumis_export_url = _acc_clean(self.in_j_export_url.text()) or JUMIS_EXPORT_URL

        cfg.generic_base_url = _acc_clean(self.in_g_base_url.text())
        cfg.generic_auth_type = self.cmb_g_auth.currentData() or 'bearer'
        cfg.generic_username = _acc_clean(self.in_g_username.text())
        cfg.generic_password = _acc_clean(self.in_g_password.text())
        cfg.generic_token = _acc_clean(self.in_g_token.text())
        cfg.generic_api_key_header = _acc_clean(self.in_g_key_header.text()) or 'X-API-Key'
        cfg.generic_api_key_value = _acc_clean(self.in_g_key_value.text())
        cfg.generic_timeout_sec = int(self.sp_timeout.value())
        cfg.generic_verify_ssl = self.ck_verify_ssl.isChecked()
        cfg.generic_test_endpoint = _acc_clean(self.in_g_test.text()) or '/health'
        cfg.generic_partners_endpoint = _acc_clean(self.in_g_partners.text()) or '/partners'
        cfg.generic_products_endpoint = _acc_clean(self.in_g_products.text()) or '/products'
        cfg.generic_documents_endpoint = _acc_clean(self.in_g_docs.text()) or '/documents'
        cfg.generic_inventory_documents_endpoint = _acc_clean(self.in_g_inv_docs.text()) or '/inventory-documents'

        cfg.auto_sync_on_pdf_generate = self.ck_auto_pdf.isChecked()
        cfg.auto_sync_on_docx_generate = self.ck_auto_docx.isChecked()
        cfg.auto_sync_on_project_save = self.ck_auto_project.isChecked()
        cfg.sync_partners_with_document = self.ck_sync_partners.isChecked()
        cfg.sync_products_with_inventory = self.ck_sync_products.isChecked()
        cfg.sync_store_doc_with_document = self.ck_sync_store.isChecked()
        cfg.sync_financial_doc_with_document = self.ck_sync_fin.isChecked()
        return cfg


def _acc_cfg_from_settings() -> AccountingIntegrationConfig:
    st = load_settings() or {}
    raw = st.get(ACCOUNTING_SETTINGS_KEY) or {}
    if not isinstance(raw, dict):
        raw = {}
    cfg = AccountingIntegrationConfig()
    for k in cfg.__dataclass_fields__.keys():
        if k in raw:
            setattr(cfg, k, raw.get(k))
    return cfg


def _acc_cfg_to_settings(cfg: AccountingIntegrationConfig):
    st = load_settings() or {}
    st[ACCOUNTING_SETTINGS_KEY] = asdict(cfg)
    save_settings(st)


def _acc_make_connector(cfg: AccountingIntegrationConfig) -> BaseAccountingConnector:
    system_type = _acc_clean(getattr(cfg, 'system_type', 'jumis')).lower() or 'jumis'
    if system_type == 'jumis':
        return JumisConnector(cfg)
    if system_type == 'generic_rest':
        return GenericRestConnector(cfg)
    raise AccountingIntegrationError(f'Neatbalstīts integrācijas tips: {system_type}')


def _acc_log(self, event: str, details: dict | None = None):
    try:
        if hasattr(self, '_audit_logger') and self._audit_logger:
            self._audit_logger.write(event, details or {}, getattr(self, '_current_user', ''))
        if hasattr(self, '_append_audit_row'):
            try:
                self._append_audit_row(_acc_now_ts(), event, json.dumps(details or {}, ensure_ascii=False))
            except Exception:
                pass
    except Exception:
        pass


def _acc_get_cfg(self) -> AccountingIntegrationConfig:
    cfg = getattr(self, '_accounting_cfg', None)
    if cfg is None:
        cfg = _acc_cfg_from_settings()
        self._accounting_cfg = cfg
    return cfg


def _acc_set_cfg(self, cfg: AccountingIntegrationConfig):
    self._accounting_cfg = cfg
    _acc_cfg_to_settings(cfg)


def _acc_connector(self) -> BaseAccountingConnector:
    cfg = self._acc_get_cfg()
    if not cfg.enabled:
        raise AccountingIntegrationError('Grāmatvedības integrācija nav ieslēgta.')
    return _acc_make_connector(cfg)


def _acc_show_result(self, results, title='Sinhronizācijas rezultāts'):
    if not isinstance(results, list):
        results = [results]
    lines = []
    ok_count = 0
    for idx, res in enumerate(results, start=1):
        if isinstance(res, Exception):
            lines.append(f'{idx}. KĻŪDA: {res}')
            continue
        ok_count += 1 if getattr(res, 'ok', False) else 0
        lines.append(
            f'{idx}. [{"OK" if res.ok else "FAIL"}] {res.system} / {res.operation} / HTTP {res.response_status}\n'
            f'Pieprasījums:\n{res.request_payload[:3000]}\n\nAtbilde:\n{(res.response_text or "")[:4000]}\n'
        )
    text = '\n' + ('-' * 80) + '\n'.join(lines) if lines else 'Nav rezultātu.'
    msg = QMessageBox(self)
    msg.setWindowTitle(title)
    msg.setIcon(QMessageBox.Information if ok_count == len(results) and results else QMessageBox.Warning)
    msg.setText(f'Pabeigts: {ok_count}/{len(results)} veiksmīgi.')
    msg.setDetailedText(text)
    msg.setStandardButtons(QMessageBox.Ok)
    msg.exec()


def _acc_open_settings_dialog(self):
    dlg = AccountingIntegrationDialog(self, self._acc_get_cfg())
    if dlg.exec() == QDialog.Accepted:
        cfg = dlg.get_config()
        self._acc_set_cfg(cfg)
        _acc_log(self, 'accounting_settings_saved', {'system_type': cfg.system_type, 'enabled': cfg.enabled, 'profile_name': cfg.profile_name})
        QMessageBox.information(self, 'Saglabāts', 'Grāmatvedības integrācijas iestatījumi saglabāti.')


def _acc_test_connection(self):
    try:
        res = self._acc_connector().test_connection()
        _acc_log(self, 'accounting_test_connection', {'ok': res.ok, 'system': res.system, 'status': res.response_status})
        _acc_show_result(self, res, 'Pieslēguma pārbaude')
    except Exception as e:
        _acc_log(self, 'accounting_test_connection_error', {'error': str(e)})
        QMessageBox.warning(self, 'Kļūda', f'Neizdevās pārbaudīt pieslēgumu.\n\n{e}')


def _acc_sync_partners_from_current_doc(self):
    try:
        d = self.savākt_datus()
        conn = self._acc_connector()
        results = []
        people = []
        for obj in [getattr(d, 'pieņēmējs', None), getattr(d, 'nodevējs', None), getattr(d, 'rekviziti', None)]:
            if obj and _persona_has_content(obj):
                key = (_acc_clean(getattr(obj, 'nosaukums', '')), _acc_clean(getattr(obj, 'reģ_nr', '')))
                if key not in people:
                    people.append(key)
                    results.append(conn.upsert_partner(obj))
        _acc_log(self, 'accounting_sync_partners', {'count': len(results)})
        _acc_show_result(self, results, 'Partneru sinhronizācija')
    except Exception as e:
        _acc_log(self, 'accounting_sync_partners_error', {'error': str(e)})
        QMessageBox.warning(self, 'Kļūda', f'Partneru sinhronizācija neizdevās.\n\n{e}')


def _acc_sync_inventory_products(self):
    try:
        conn = self._acc_connector()
        db = getattr(self, '_noliktava', None)
        if db is None:
            raise AccountingIntegrationError('Noliktava nav inicializēta.')
        items = list(getattr(db, 'items', []) or [])
        if not items:
            raise AccountingIntegrationError('Noliktavā nav preču, ko sinhronizēt.')
        results = [conn.upsert_product(it) for it in items]
        _acc_log(self, 'accounting_sync_inventory', {'count': len(results)})
        _acc_show_result(self, results, 'Noliktavas preču sinhronizācija')
    except Exception as e:
        _acc_log(self, 'accounting_sync_inventory_error', {'error': str(e)})
        QMessageBox.warning(self, 'Kļūda', f'Noliktavas preču sinhronizācija neizdevās.\n\n{e}')


def _acc_sync_current_document(self, silent: bool = False):
    try:
        cfg = self._acc_get_cfg()
        conn = self._acc_connector()
        d = self.savākt_datus()
        results = []
        if cfg.sync_partners_with_document:
            for obj in [getattr(d, 'pieņēmējs', None), getattr(d, 'nodevējs', None), getattr(d, 'rekviziti', None)]:
                if obj and _persona_has_content(obj):
                    results.append(conn.upsert_partner(obj))
        if cfg.sync_store_doc_with_document:
            results.append(conn.export_store_document(d))
        if cfg.sync_financial_doc_with_document:
            results.append(conn.export_financial_document(d))
        _acc_log(self, 'accounting_sync_current_document', {'count': len(results), 'silent': silent})
        if not silent:
            _acc_show_result(self, results, 'Dokumenta sinhronizācija')
        return results
    except Exception as e:
        _acc_log(self, 'accounting_sync_current_document_error', {'error': str(e), 'silent': silent})
        if not silent:
            QMessageBox.warning(self, 'Kļūda', f'Dokumenta sinhronizācija neizdevās.\n\n{e}')
        else:
            print(f'Accounting sync error: {e}')
        return []


def _acc_sync_everything(self):
    try:
        cfg = self._acc_get_cfg()
        conn = self._acc_connector()
        d = self.savākt_datus()
        results = []
        # partneri
        if cfg.sync_partners_with_document:
            seen = set()
            for obj in [getattr(d, 'pieņēmējs', None), getattr(d, 'nodevējs', None), getattr(d, 'rekviziti', None)]:
                if obj and _persona_has_content(obj):
                    key = (_acc_clean(getattr(obj, 'nosaukums', '')), _acc_clean(getattr(obj, 'reģ_nr', '')))
                    if key in seen:
                        continue
                    seen.add(key)
                    results.append(conn.upsert_partner(obj))
        # noliktava
        if cfg.sync_products_with_inventory and getattr(self, '_noliktava', None) is not None:
            for it in list(getattr(self._noliktava, 'items', []) or []):
                results.append(conn.upsert_product(it))
        # dokuments
        if cfg.sync_store_doc_with_document:
            results.append(conn.export_store_document(d))
        if cfg.sync_financial_doc_with_document:
            results.append(conn.export_financial_document(d))
        _acc_log(self, 'accounting_sync_everything', {'count': len(results)})
        _acc_show_result(self, results, 'Pilnā sinhronizācija')
    except Exception as e:
        _acc_log(self, 'accounting_sync_everything_error', {'error': str(e)})
        QMessageBox.warning(self, 'Kļūda', f'Pilnā sinhronizācija neizdevās.\n\n{e}')


def _acc_fetch_remote_partners(self):
    try:
        conn = self._acc_connector()
        res = conn.fetch_partners(limit=100)
        _acc_log(self, 'accounting_fetch_partners', {'ok': res.ok, 'status': res.response_status})
        _acc_show_result(self, res, 'Attālināto partneru nolasīšana')
    except Exception as e:
        _acc_log(self, 'accounting_fetch_partners_error', {'error': str(e)})
        QMessageBox.warning(self, 'Kļūda', f'Partneru nolasīšana neizdevās.\n\n{e}')


def _acc_fetch_remote_products(self):
    try:
        conn = self._acc_connector()
        res = conn.fetch_products(limit=100)
        _acc_log(self, 'accounting_fetch_products', {'ok': res.ok, 'status': res.response_status})
        _acc_show_result(self, res, 'Attālināto preču nolasīšana')
    except Exception as e:
        _acc_log(self, 'accounting_fetch_products_error', {'error': str(e)})
        QMessageBox.warning(self, 'Kļūda', f'Preču nolasīšana neizdevās.\n\n{e}')


def _acc_build_menu(self):
    try:
        menubar = self.menuBar()
        existing = getattr(self, '_accounting_menu', None)
        if existing is not None:
            return
        m = menubar.addMenu('Grāmatvedība')
        self._accounting_menu = m
        actions = [
            ('Integrācijas iestatījumi…', _acc_open_settings_dialog),
            ('Pārbaudīt pieslēgumu', _acc_test_connection),
            None,
            ('Sinhronizēt pašreizējo dokumentu', _acc_sync_current_document),
            ('Sinhronizēt partnerus no dokumenta', _acc_sync_partners_from_current_doc),
            ('Sinhronizēt visas noliktavas preces', _acc_sync_inventory_products),
            ('Pilnā sinhronizācija', _acc_sync_everything),
            None,
            ('Nolasīt partnerus no attālās sistēmas', _acc_fetch_remote_partners),
            ('Nolasīt preces no attālās sistēmas', _acc_fetch_remote_products),
        ]
        for spec in actions:
            if spec is None:
                m.addSeparator()
                continue
            title, fn = spec
            a = QAction(title, self)
            a.triggered.connect(lambda _=False, f=fn: f(self))
            m.addAction(a)
    except Exception as e:
        print(f'Accounting menu build failed: {e}')


def _acc_after_init(self):
    try:
        self._accounting_cfg = _acc_cfg_from_settings()
    except Exception:
        self._accounting_cfg = AccountingIntegrationConfig()
    _acc_build_menu(self)


def _acc_maybe_auto_sync(self, trigger_name: str):
    try:
        cfg = self._acc_get_cfg()
        if not cfg.enabled:
            return
        enabled = {
            'pdf': bool(cfg.auto_sync_on_pdf_generate),
            'docx': bool(cfg.auto_sync_on_docx_generate),
            'project_save': bool(cfg.auto_sync_on_project_save),
        }.get(trigger_name, False)
        if not enabled:
            return
        _acc_sync_current_document(self, silent=True)
        _acc_log(self, 'accounting_auto_sync_done', {'trigger': trigger_name})
    except Exception as e:
        _acc_log(self, 'accounting_auto_sync_error', {'trigger': trigger_name, 'error': str(e)})


# Monkey-patch uz AktaLogs – tikai pievienojot jaunas iespējas
_AKTA_V64_OLD_INIT = AktaLogs.__init__

def _akta_v64_init_wrapper(self, *args, **kwargs):
    _AKTA_V64_OLD_INIT(self, *args, **kwargs)
    _acc_after_init(self)

AktaLogs.__init__ = _akta_v64_init_wrapper
AktaLogs._acc_get_cfg = _acc_get_cfg
AktaLogs._acc_set_cfg = _acc_set_cfg
AktaLogs._acc_connector = _acc_connector
AktaLogs._acc_open_settings_dialog = _acc_open_settings_dialog
AktaLogs._acc_test_connection = _acc_test_connection
AktaLogs._acc_sync_current_document = _acc_sync_current_document
AktaLogs._acc_sync_partners_from_current_doc = _acc_sync_partners_from_current_doc
AktaLogs._acc_sync_inventory_products = _acc_sync_inventory_products
AktaLogs._acc_sync_everything = _acc_sync_everything
AktaLogs._acc_fetch_remote_partners = _acc_fetch_remote_partners
AktaLogs._acc_fetch_remote_products = _acc_fetch_remote_products
AktaLogs._acc_show_result = _acc_show_result
AktaLogs._acc_maybe_auto_sync = _acc_maybe_auto_sync
AktaLogs._acc_build_menu = _acc_build_menu


# Wrapperi esošajām darbībām, lai automātiskā sinhronizācija ir pieejama bez vecā koda pārtaisīšanas
_AKTA_V64_OLD_PDF = AktaLogs.ģenerēt_pdf_dialogs
_AKTA_V64_OLD_DOCX = AktaLogs.ģenerēt_docx_dialogs
_AKTA_V64_OLD_SAVE_PROJECT = AktaLogs.saglabat_projektu


def _akta_v64_pdf_wrapper(self, *args, **kwargs):
    res = _AKTA_V64_OLD_PDF(self, *args, **kwargs)
    _acc_maybe_auto_sync(self, 'pdf')
    return res


def _akta_v64_docx_wrapper(self, *args, **kwargs):
    res = _AKTA_V64_OLD_DOCX(self, *args, **kwargs)
    _acc_maybe_auto_sync(self, 'docx')
    return res


def _akta_v64_save_project_wrapper(self, *args, **kwargs):
    res = _AKTA_V64_OLD_SAVE_PROJECT(self, *args, **kwargs)
    _acc_maybe_auto_sync(self, 'project_save')
    return res


AktaLogs.ģenerēt_pdf_dialogs = _akta_v64_pdf_wrapper
AktaLogs.ģenerēt_docx_dialogs = _akta_v64_docx_wrapper
AktaLogs.saglabat_projektu = _akta_v64_save_project_wrapper



# ==============================
# v67: salabota divvirzienu produktu sinhronizācija (kods izpildās PIRMS __main__),
# robustāka Jumis atbilžu apstrāde un stabilāki noliktavas hooki.
# ==============================
import xml.etree.ElementTree as _ET_V67
from dataclasses import asdict as _dc_asdict_v67


def _acc_cfg_get_v67(cfg, name, default=None):
    try:
        val = getattr(cfg, name)
    except Exception:
        return default
    return default if val is None else val


def _acc_json_from_response_text_v67(text: str):
    try:
        return json.loads(text or '')
    except Exception:
        return None


def _acc_extract_xml_from_payload_v67(payload) -> str:
    if payload is None:
        return ''
    if isinstance(payload, str):
        s = payload.strip()
        if s.startswith('<'):
            return s
        js = _acc_json_from_response_text_v67(s)
        if js is not None:
            return _acc_extract_xml_from_payload_v67(js)
        start = s.find('<?xml')
        if start < 0:
            start = s.find('<dataroot')
        return s[start:] if start >= 0 else ''
    if isinstance(payload, dict):
        for key in ['XMLresponse', 'xmlResponse', 'XML', 'xml', 'ResponseData', 'responseData', 'Result', 'result', 'Data', 'data']:
            if key in payload:
                xml = _acc_extract_xml_from_payload_v67(payload.get(key))
                if xml:
                    return xml
        for v in payload.values():
            xml = _acc_extract_xml_from_payload_v67(v)
            if xml:
                return xml
        return ''
    if isinstance(payload, list):
        for v in payload:
            xml = _acc_extract_xml_from_payload_v67(v)
            if xml:
                return xml
        return ''
    return ''


def _acc_extract_xml_from_response_text_v67(text: str) -> str:
    return _acc_extract_xml_from_payload_v67(text or '')


def _acc_response_is_success_v67(status_code: int, text: str) -> tuple[bool, str]:
    ok = 200 <= int(status_code or 0) < 300
    payload = _acc_json_from_response_text_v67(text or '')
    msg = ''
    hay = (text or '').lower()
    if isinstance(payload, dict):
        for key in ['Success', 'success', 'IsSuccess', 'isSuccess', 'Succeeded', 'succeeded']:
            if key in payload:
                try:
                    ok = ok and bool(payload.get(key))
                except Exception:
                    pass
        for key in ['Message', 'message', 'ErrorMessage', 'errorMessage', 'Description', 'description']:
            if payload.get(key):
                msg = str(payload.get(key))
                break
        try:
            hay = json.dumps(payload, ensure_ascii=False).lower()
        except Exception:
            hay = str(payload).lower()
    fail_markers = [
        'neizdevās pieslēgties', 'kluda', 'kļūda', 'error', 'exception', 'unauthorized', 'forbidden',
        'invalid', 'not allowed', 'nav tiesību', 'authentication failed'
    ]
    if any(m in hay for m in fail_markers):
        ok = False
    return bool(ok), msg


def _acc_result_with_status_v67(res):
    ok, msg = _acc_response_is_success_v67(getattr(res, 'response_status', 0), getattr(res, 'response_text', ''))
    res.ok = ok
    if msg:
        extra = dict(getattr(res, 'extra', {}) or {})
        extra['message'] = msg
        res.extra = extra
    return res


def _acc_parse_xml_root_v67(xml_text: str):
    xml_text = (xml_text or '').strip()
    if not xml_text:
        return None
    try:
        return _ET_V67.fromstring(xml_text)
    except Exception:
        try:
            xml_text = xml_text.replace('encoding="utf-16"', 'encoding="utf-8"')
            return _ET_V67.fromstring(xml_text)
        except Exception:
            return None


def _acc_local_name_v67(tag) -> str:
    tag = str(tag or '')
    return tag.split('}', 1)[-1]


def _acc_findall_any_v67(root, names):
    if root is None:
        return []
    if isinstance(names, str):
        names = {names}
    else:
        names = set(names)
    out = []
    for el in root.iter():
        if _acc_local_name_v67(getattr(el, 'tag', '')) in names:
            out.append(el)
    return out


def _acc_child_text_v67(el, tag_names, default=''):
    if el is None:
        return default
    if isinstance(tag_names, str):
        tag_names = {tag_names}
    else:
        tag_names = set(tag_names)
    for ch in list(el):
        if _acc_local_name_v67(getattr(ch, 'tag', '')) in tag_names:
            return (ch.text or '').strip()
    return default


def _acc_collect_table_rows_v67(payload, rows=None):
    if rows is None:
        rows = []
    if isinstance(payload, dict):
        if payload and all(not isinstance(v, (dict, list)) for v in payload.values()):
            rows.append({str(k): '' if v is None else str(v) for k, v in payload.items()})
        else:
            for v in payload.values():
                _acc_collect_table_rows_v67(v, rows)
    elif isinstance(payload, list):
        for v in payload:
            _acc_collect_table_rows_v67(v, rows)
    return rows


def _acc_parse_jumis_products_response_v67(text: str) -> list[dict]:
    products = []
    payload = _acc_json_from_response_text_v67(text or '')
    if payload is not None:
        for row in _acc_collect_table_rows_v67(payload, []):
            if any(k in row for k in ['ProductCode', 'ProductName', 'ProductBarCode', 'ProductUnit', 'ProductPrice']):
                products.append({
                    'ProductCode': row.get('ProductCode', ''),
                    'ProductName': row.get('ProductName', ''),
                    'ProductBarCode': row.get('ProductBarCode', ''),
                    'ProductUnit': row.get('ProductUnit', row.get('ProductUnitCode', 'gab.')),
                    'ProductPrice': row.get('ProductPrice', row.get('Price', '0')),
                    'ProductPurchasePrice': row.get('ProductPurchasePrice', row.get('PurchasePrice', '')),
                    'ProductVATRate': row.get('ProductVATRate', row.get('VatRate', '')),
                    'ProductGroupName': row.get('ProductGroupName', row.get('GroupName', '')),
                    'ProductDescription': row.get('ProductDescription', row.get('Description', '')),
                    'StockQuantity': row.get('StockQuantity', row.get('Quantity', row.get('Balance', '0'))),
                    'Status': row.get('Status', row.get('ProductStatus', '')),
                })
    xml_text = _acc_extract_xml_from_response_text_v67(text)
    root = _acc_parse_xml_root_v67(xml_text)
    if root is not None:
        row_like_tags = {'Product', 'Row', 'ProductRow', 'Item', 'Record'}
        for p in _acc_findall_any_v67(root, row_like_tags):
            code = _acc_child_text_v67(p, ['ProductCode', 'Code'])
            name = _acc_child_text_v67(p, ['ProductName', 'Name'])
            barcode = _acc_child_text_v67(p, ['ProductBarCode', 'BarCode', 'Barcode'])
            unit = _acc_child_text_v67(p, ['ProductUnit', 'Unit', 'UnitCode'], 'gab.')
            price = _acc_child_text_v67(p, ['ProductPrice', 'Price'], '0')
            if not any([code, name, barcode, price]):
                continue
            products.append({
                'ProductCode': code,
                'ProductName': name,
                'ProductBarCode': barcode,
                'ProductUnit': unit,
                'ProductPrice': price,
                'ProductPurchasePrice': _acc_child_text_v67(p, ['ProductPurchasePrice', 'PurchasePrice']),
                'ProductVATRate': _acc_child_text_v67(p, ['ProductVATRate', 'VatRate']),
                'ProductGroupName': _acc_child_text_v67(p, ['ProductGroupName', 'GroupName']),
                'ProductDescription': _acc_child_text_v67(p, ['ProductDescription', 'Description']),
                'StockQuantity': _acc_child_text_v67(p, ['StockQuantity', 'Quantity', 'Balance'], '0'),
                'Status': _acc_child_text_v67(p, ['Status', 'ProductStatus']),
            })
    # dedupe
    dedup = {}
    for p in products:
        key = (_acc_clean(p.get('ProductCode')), _acc_clean(p.get('ProductBarCode')), _acc_clean(p.get('ProductName')))
        if any(key):
            dedup[key] = p
    return list(dedup.values())


def _acc_parse_jumis_partners_response_v67(text: str) -> list[dict]:
    out = []
    payload = _acc_json_from_response_text_v67(text or '')
    if payload is not None:
        for row in _acc_collect_table_rows_v67(payload, []):
            if any(k in row for k in ['PartnerName', 'PartnerRegistrationNo', 'PartnerEMail']):
                out.append({
                    'PartnerName': row.get('PartnerName', ''),
                    'PartnerRegistrationNo': row.get('PartnerRegistrationNo', ''),
                    'PartnerEMail': row.get('PartnerEMail', ''),
                    'PartnerPhone': row.get('PartnerPhone', ''),
                    'PartnerWWW': row.get('PartnerWWW', ''),
                })
    xml_text = _acc_extract_xml_from_response_text_v67(text)
    root = _acc_parse_xml_root_v67(xml_text)
    for p in _acc_findall_any_v67(root, {'Partner', 'Row', 'Record'}):
        name = _acc_child_text_v67(p, ['PartnerName', 'Name'])
        reg = _acc_child_text_v67(p, ['PartnerRegistrationNo', 'RegistrationNo'])
        if not name and not reg:
            continue
        out.append({
            'PartnerName': name,
            'PartnerRegistrationNo': reg,
            'PartnerEMail': _acc_child_text_v67(p, ['PartnerEMail', 'EMail', 'Email']),
            'PartnerPhone': _acc_child_text_v67(p, ['PartnerPhone', 'Phone']),
            'PartnerWWW': _acc_child_text_v67(p, ['PartnerWWW', 'WWW']),
        })
    dedup = {}
    for p in out:
        key = (_acc_clean(p.get('PartnerRegistrationNo')), _acc_clean(p.get('PartnerName')))
        dedup[key] = p
    return list(dedup.values())


def _acc_make_item_from_remote_product_v67(raw: dict, existing=None):
    base = _dc_asdict_v67(existing) if existing is not None else _dc_asdict_v67(NoliktavasPrece())
    code = _acc_clean(raw.get('ProductCode', base.get('sku', '')))
    base['sku'] = code or _acc_clean(base.get('sku', ''))
    base['inventory_id'] = _acc_clean(base.get('inventory_id', '')) or code or f"acc-{secrets.token_hex(8)}"
    base['nosaukums'] = _acc_clean(raw.get('ProductName', base.get('nosaukums', '')))
    base['svitrkods'] = _acc_clean(raw.get('ProductBarCode', base.get('svitrkods', '')))
    base['vieniba'] = _acc_clean(raw.get('ProductUnit', base.get('vieniba', 'gab.'))) or 'gab.'
    base['cena'] = _acc_clean(raw.get('ProductPrice', base.get('cena', '')))
    base['pvn_likme'] = _acc_clean(raw.get('ProductVATRate', base.get('pvn_likme', '')))
    base['kategorija'] = _acc_clean(raw.get('ProductGroupName', base.get('kategorija', '')))
    base['piezimes'] = _acc_clean(raw.get('ProductDescription', base.get('piezimes', '')))
    status = _acc_clean(raw.get('Status', ''))
    if status:
        base['statuss'] = status
    try:
        base['atlikums'] = float(str(raw.get('StockQuantity', base.get('atlikums', 0.0)) or 0.0).replace(',', '.'))
    except Exception:
        pass
    base['last_sync_at'] = _acc_now_ts()
    return NoliktavasPrece(**base)


def _acc_find_item_by_remote_product_v67(db, raw: dict):
    code = _acc_clean(raw.get('ProductCode'))
    barcode = _acc_clean(raw.get('ProductBarCode'))
    name = _acc_clean(raw.get('ProductName'))
    for it in list(getattr(db, 'items', []) or []):
        if code and _acc_clean(getattr(it, 'sku', '')) == code:
            return it
        if code and _acc_clean(getattr(it, 'inventory_id', '')) == code:
            return it
        if barcode and _acc_clean(getattr(it, 'svitrkods', '')) == barcode:
            return it
    for it in list(getattr(db, 'items', []) or []):
        if name and _acc_clean(getattr(it, 'nosaukums', '')) == name:
            return it
    return None


def _acc_item_remote_key_v67(it) -> str:
    return _acc_clean(getattr(it, 'sku', '')) or _acc_clean(getattr(it, 'inventory_id', '')) or _acc_clean(getattr(it, 'svitrkods', ''))


def _acc_inventory_item_signature_v67(it) -> str:
    if it is None:
        return ''
    data = _dc_asdict_v67(it)
    for k in ['last_sync_at']:
        data.pop(k, None)
    try:
        return hashlib.sha256(json.dumps(data, sort_keys=True, ensure_ascii=False).encode('utf-8')).hexdigest()
    except Exception:
        return repr(sorted(data.items()))


def _acc_appdata_file_v67(name: str) -> str:
    try:
        root = APP_DATA_DIR
    except Exception:
        root = os.path.join(os.path.expanduser('~'), '.akta_generators')
    os.makedirs(root, exist_ok=True)
    return os.path.join(root, name)


def _acc_load_live_state_v67() -> dict:
    path = _acc_appdata_file_v67('accounting_live_sync_state_v67.json')
    try:
        with open(path, 'r', encoding='utf-8') as f:
            data = json.load(f) or {}
        if isinstance(data, dict):
            return data
    except Exception:
        pass
    return {'tombstones': {}, 'last_remote_products': {}, 'last_live_sync_at': ''}


def _acc_save_live_state_v67(state: dict):
    path = _acc_appdata_file_v67('accounting_live_sync_state_v67.json')
    try:
        with open(path, 'w', encoding='utf-8') as f:
            json.dump(state or {}, f, ensure_ascii=False, indent=2)
    except Exception:
        pass


def _acc_live_state_get_v67(self):
    st = getattr(self, '_accounting_live_state', None)
    if not isinstance(st, dict):
        st = _acc_load_live_state_v67()
        self._accounting_live_state = st
    st.setdefault('tombstones', {})
    st.setdefault('last_remote_products', {})
    st.setdefault('last_live_sync_at', '')
    return st


def _acc_live_state_save_v67(self):
    _acc_save_live_state_v67(_acc_live_state_get_v67(self))


def _jumis_post_import_v67(self, xmlrequest: str, operation: str):
    payload = self._credentials_payload(xmlrequest)
    r = self.session.post(_acc_clean(self.config.jumis_import_url) or JUMIS_IMPORT_URL, json=payload, timeout=max(int(getattr(self.config, 'jumis_timeout_sec', 45) or 45), 10))
    return _acc_result_with_status_v67(AccountingSyncResult(ok=r.ok, operation=operation, system=self.system_name, request_payload=xmlrequest, response_status=r.status_code, response_text=r.text))


def _jumis_post_export_v67(self, xmlrequest: str, operation: str):
    payload = self._credentials_payload(xmlrequest)
    r = self.session.post(_acc_clean(self.config.jumis_export_url) or JUMIS_EXPORT_URL, json=payload, timeout=max(int(getattr(self.config, 'jumis_timeout_sec', 45) or 45), 10))
    return _acc_result_with_status_v67(AccountingSyncResult(ok=r.ok, operation=operation, system=self.system_name, request_payload=xmlrequest, response_status=r.status_code, response_text=r.text))


def _jumis_test_connection_v67(self):
    xmlrequest = ('<?xml version="1.0" ?><dataroot><tjDocument Version="TJ5.5.101"/>'
                  '<tjRequest Name="Partner" Operation="Read" Version="TJ7.0.112" Structure="Tree">'
                  '<tjData Read="All"/><tjFields><Field Name="PartnerName"/></tjFields>'
                  '</tjRequest></dataroot>')
    res = self._post_export(xmlrequest, 'test_connection')
    extra = dict(getattr(res, 'extra', {}) or {})
    extra['partners'] = _acc_parse_jumis_partners_response_v67(res.response_text)
    res.extra = extra
    return res


def _jumis_fetch_partners_v67(self, limit: int = 1000):
    xmlrequest = ('<?xml version="1.0" ?><dataroot><tjDocument Version="TJ5.5.101"/>'
                  '<tjRequest Name="Partner" Operation="Read" Version="TJ7.0.112" Structure="Sheet">'
                  '<tjData Read="All"/><tjFields>'
                  '<Field Name="PartnerName"/><Field Name="PartnerRegistrationNo"/><Field Name="PartnerEMail"/>'
                  '<Field Name="PartnerPhone"/><Field Name="PartnerWWW"/>'
                  '</tjFields></tjRequest></dataroot>')
    res = self._post_export(xmlrequest, 'fetch_partners')
    extra = dict(getattr(res, 'extra', {}) or {})
    extra['limit'] = limit
    extra['partners'] = _acc_parse_jumis_partners_response_v67(res.response_text)[:max(int(limit or 1000),1)]
    res.extra = extra
    return res


def _jumis_fetch_products_v67(self, limit: int = 5000):
    requests_xml = [
        ('Sheet', '<Field Name="ProductCode"/><Field Name="ProductName"/><Field Name="ProductBarCode"/><Field Name="ProductUnit"/><Field Name="ProductPrice"/><Field Name="ProductPurchasePrice"/><Field Name="ProductVATRate"/><Field Name="ProductGroupName"/><Field Name="ProductDescription"/><Field Name="Quantity"/><Field Name="StockQuantity"/><Field Name="Status"/>'),
        ('Tree', '<Field Name="ProductCode"/><Field Name="ProductName"/><Field Name="ProductBarCode"/><Field Name="ProductUnit"/><Field Name="ProductPrice"/><Field Name="ProductPurchasePrice"/><Field Name="ProductVATRate"/><Field Name="ProductGroupName"/><Field Name="ProductDescription"/><Field Name="Quantity"/><Field Name="StockQuantity"/><Field Name="Status"/>')
    ]
    final_res = None
    combined = []
    for structure, fields in requests_xml:
        xmlrequest = f'<?xml version="1.0" ?><dataroot><tjDocument Version="TJ5.5.101"/><tjRequest Name="Product" Operation="Read" Version="TJ5.5.101" Structure="{structure}"><tjData Read="All"/><tjFields>{fields}</tjFields></tjRequest></dataroot>'
        res = self._post_export(xmlrequest, f'fetch_products_{structure.lower()}')
        products = _acc_parse_jumis_products_response_v67(res.response_text)
        if products:
            combined.extend(products)
            final_res = res
            break
        final_res = res
    dedup = {}
    for p in combined:
        key = (_acc_clean(p.get('ProductCode')), _acc_clean(p.get('ProductBarCode')), _acc_clean(p.get('ProductName')))
        if any(key):
            dedup[key] = p
    products = list(dedup.values())[:max(int(limit or 5000), 1)]
    extra = dict(getattr(final_res, 'extra', {}) or {})
    extra['limit'] = limit
    extra['products'] = products
    final_res.extra = extra
    if not products and final_res.ok:
        # atstājam ok, bet norādām, ka parsēšana neko neatdeva
        extra['message'] = (extra.get('message') or 'Preču saraksts netika atrasts atbildē vai ir tukšs.')
        final_res.extra = extra
    return final_res


def _jumis_product_xml_v67(self, item):
    it = item or NoliktavasPrece()
    code = _acc_clean(getattr(it, 'sku', '')) or _acc_clean(getattr(it, 'inventory_id', '')) or _acc_clean(getattr(it, 'svitrkods', '')) or f'ITEM-{secrets.token_hex(4)}'
    try:
        if not _acc_clean(getattr(it, 'sku', '')):
            it.sku = code
    except Exception:
        pass
    try:
        if not _acc_clean(getattr(it, 'inventory_id', '')):
            it.inventory_id = code
    except Exception:
        pass
    price = _acc_to_amount(getattr(it, 'cena', '0'))
    qty = _acc_to_qty(getattr(it, 'atlikums', 0))
    return (
        '<?xml version="1.0" encoding="utf-8" ?>'
        '<dataroot><tjDocument Version="TJ5.5.101"/>'
        '<tjResponse Name="Product" Operation="Insert" Version="TJ5.5.101" Structure="Tree">'
        '<Product>'
        f'{_acc_xml("ProductCode", code)}'
        f'{_acc_xml("ProductName", _acc_clean(getattr(it, "nosaukums", "")) or code)}'
        f'{_acc_xml("ProductBarCode", _acc_clean(getattr(it, "svitrkods", "")))}'
        f'{_acc_xml("ProductUnit", _acc_clean(getattr(it, "vieniba", "")) or "gab.")}'
        f'{_acc_xml("ProductPrice", price)}'
        f'{_acc_xml("ProductPurchasePrice", price)}'
        f'{_acc_xml("ProductVATRate", _acc_to_amount(getattr(it, "pvn_likme", "21"), default="21.00"))}'
        f'{_acc_xml("ProductGroupName", _acc_clean(getattr(it, "kategorija", "")))}'
        f'{_acc_xml("ProductDescription", _acc_clean(getattr(it, "piezimes", "")))}'
        f'{_acc_xml("Quantity", qty)}'
        '</Product></tjResponse></dataroot>'
    )


def _jumis_upsert_product_v67(self, item):
    return self._post_import(_jumis_product_xml_v67(self, item), 'upsert_product')


def _jumis_delete_product_v67(self, item):
    code = _acc_item_remote_key_v67(item)
    xmlrequest = ('<?xml version="1.0" encoding="utf-8" ?><dataroot><tjDocument Version="TJ5.5.101"/>'
                  '<tjResponse Name="Product" Operation="Delete" Version="TJ5.5.101" Structure="Tree"><Product>'
                  f'{_acc_xml("ProductCode", code)}'
                  '</Product></tjResponse></dataroot>')
    res = self._post_import(xmlrequest, 'delete_product')
    if not getattr(res, 'ok', False):
        clone = NoliktavasPrece(**_dc_asdict_v67(item))
        clone.atlikums = 0.0
        clone.statuss = 'Dzēsta'
        clone.piezimes = ((_acc_clean(getattr(clone, 'piezimes', '')) + '\n[SYNC] Atzīmēta kā dzēsta.').strip())[:4000]
        fb = self.upsert_product(clone)
        fb.operation = 'delete_product_fallback_upsert'
        extra = dict(getattr(fb, 'extra', {}) or {})
        extra['delete_fallback'] = True
        fb.extra = extra
        return fb
    return res


def _acc_show_result_v67(self, results, title='Sinhronizācijas rezultāts'):
    if not isinstance(results, list):
        results = [results]
    lines = []
    ok_count = 0
    for idx, res in enumerate(results, start=1):
        if isinstance(res, Exception):
            lines.append(f'{idx}. KĻŪDA: {res}')
            continue
        ok = bool(getattr(res, 'ok', False))
        ok_count += 1 if ok else 0
        extra = dict(getattr(res, 'extra', {}) or {})
        bits = []
        if extra.get('message'):
            bits.append(f'Ziņa: {extra.get("message")}')
        if extra.get('products') is not None:
            bits.append(f'Atrasto preču skaits: {len(extra.get("products") or [])}')
        if extra.get('partners') is not None:
            bits.append(f'Atrasto partneru skaits: {len(extra.get("partners") or [])}')
        if extra.get('apply_stats') is not None:
            bits.append(f'Pielietots lokāli: {json.dumps(extra.get("apply_stats"), ensure_ascii=False)}')
        meta = ('\n' + '\n'.join(bits)) if bits else ''
        lines.append(f'{idx}. [{'OK' if ok else 'FAIL'}] {getattr(res, "system", "")} / {getattr(res, "operation", "")} / HTTP {getattr(res, "response_status", 0)}{meta}\nPieprasījums:\n{(getattr(res, "request_payload", "") or "")[:3000]}\n\nAtbilde:\n{(getattr(res, "response_text", "") or "")[:4000]}\n')
    text = '\n' + ('-' * 80) + '\n'.join(lines) if lines else 'Nav rezultātu.'
    msg = QMessageBox(self)
    msg.setWindowTitle(title)
    msg.setIcon(QMessageBox.Information if results and ok_count == len(results) else QMessageBox.Warning)
    msg.setText(f'Pabeigts: {ok_count}/{len(results)} veiksmīgi.')
    msg.setDetailedText(text)
    msg.setStandardButtons(QMessageBox.Ok)
    msg.exec()


def _acc_fetch_remote_products_v67(self, quiet: bool = False):
    try:
        conn = self._acc_connector()
        limit = max(int(_acc_cfg_get_v67(self._acc_get_cfg(), 'live_sync_fetch_limit', 5000) or 5000), 1)
        res = conn.fetch_products(limit=limit)
        _acc_log(self, 'accounting_fetch_products', {'ok': getattr(res, 'ok', False), 'status': getattr(res, 'response_status', 0), 'quiet': quiet})
        if not quiet:
            _acc_show_result_v67(self, res, 'Attālināto preču nolasīšana')
        return res
    except Exception as e:
        _acc_log(self, 'accounting_fetch_products_error', {'error': str(e), 'quiet': quiet})
        if not quiet:
            QMessageBox.warning(self, 'Kļūda', f'Preču nolasīšana neizdevās.\n\n{e}')
        return AccountingSyncResult(ok=False, operation='fetch_products', system='N/A', response_status=0, response_text=str(e))


def _acc_apply_remote_products_to_local_v67(self, products: list, remove_missing: bool = False):
    db = getattr(self, '_noliktava', None)
    if db is None:
        raise AccountingIntegrationError('Noliktava nav inicializēta.')
    products = list(products or [])
    state = _acc_live_state_get_v67(self)
    tombstones = dict(state.get('tombstones', {}) or {})
    remote_keys = set()
    created = updated = removed = 0
    self._accounting_sync_muted = True
    try:
        for raw in products:
            key = _acc_clean(raw.get('ProductCode')) or _acc_clean(raw.get('ProductBarCode')) or _acc_clean(raw.get('ProductName'))
            if key:
                remote_keys.add(key)
            if key and key in tombstones:
                continue
            existing = _acc_find_item_by_remote_product_v67(db, raw)
            new_item = _acc_make_item_from_remote_product_v67(raw, existing)
            before_sig = _acc_inventory_item_signature_v67(existing) if existing is not None else ''
            after_sig = _acc_inventory_item_signature_v67(new_item)
            db.upsert(new_item)
            if existing is None:
                created += 1
            elif before_sig != after_sig:
                updated += 1
        if remove_missing and remote_keys:
            remove_ids = []
            for it in list(getattr(db, 'items', []) or []):
                key = _acc_item_remote_key_v67(it)
                if key and key not in remote_keys and key not in tombstones:
                    remove_ids.append(getattr(it, 'inventory_id', ''))
            if remove_ids:
                removed = db.delete_by_inventory_ids(remove_ids)
        state['last_remote_products'] = {(_acc_clean(p.get('ProductCode')) or _acc_clean(p.get('ProductBarCode')) or _acc_clean(p.get('ProductName'))): p for p in products}
        state['last_live_sync_at'] = _acc_now_ts()
        self._accounting_live_state = state
        _acc_live_state_save_v67(self)
    finally:
        self._accounting_sync_muted = False
    return {'created': created, 'updated': updated, 'removed': removed, 'total_remote': len(products)}


def _acc_pull_remote_inventory_now_v67(self, silent: bool = False):
    res = _acc_fetch_remote_products_v67(self, quiet=True)
    if not getattr(res, 'ok', False):
        if not silent:
            _acc_show_result_v67(self, res, 'Attālinātās noliktavas ielāde')
        return res
    products = list((getattr(res, 'extra', {}) or {}).get('products') or [])
    stats = _acc_apply_remote_products_to_local_v67(self, products, remove_missing=bool(_acc_cfg_get_v67(self._acc_get_cfg(), 'live_sync_remove_missing_local_products', False)))
    extra = dict(getattr(res, 'extra', {}) or {})
    extra['apply_stats'] = stats
    res.extra = extra
    _acc_log(self, 'accounting_pull_remote_inventory', stats)
    try:
        if hasattr(self, '_refresh_noliktava_table'):
            self._refresh_noliktava_table()
    except Exception:
        pass
    try:
        if hasattr(self, '_update_preview'):
            self._update_preview()
    except Exception:
        pass
    if not silent:
        _acc_show_result_v67(self, res, 'Attālinātās noliktavas ielāde')
    return res


def _acc_prepare_local_item_for_push_v67(self, item):
    changed = False
    code = _acc_item_remote_key_v67(item)
    if code and not _acc_clean(getattr(item, 'sku', '')):
        try:
            item.sku = code
            changed = True
        except Exception:
            pass
    if code and not _acc_clean(getattr(item, 'inventory_id', '')):
        try:
            item.inventory_id = code
            changed = True
        except Exception:
            pass
    if changed:
        try:
            self._noliktava.save()
        except Exception:
            pass
    return item


def _acc_push_single_local_item_v67(self, item, reason: str = ''):
    if getattr(self, '_accounting_sync_muted', False):
        return None
    cfg = self._acc_get_cfg()
    if not cfg.enabled or not bool(_acc_cfg_get_v67(cfg, 'live_sync_enabled', True)):
        return None
    if not bool(_acc_cfg_get_v67(cfg, 'live_sync_push_local_changes', True)):
        return None
    item = _acc_prepare_local_item_for_push_v67(self, item)
    conn = self._acc_connector()
    res = conn.upsert_product(item)
    _acc_log(self, 'accounting_live_push_product', {'ok': getattr(res, 'ok', False), 'sku': getattr(item, 'sku', ''), 'inventory_id': getattr(item, 'inventory_id', ''), 'reason': reason, 'status': getattr(res, 'response_status', 0)})
    st = _acc_live_state_get_v67(self)
    key = _acc_item_remote_key_v67(item)
    if key and key in st.get('tombstones', {}):
        st['tombstones'].pop(key, None)
        _acc_live_state_save_v67(self)
    return res


def _acc_delete_remote_item_v67(self, item, reason: str = ''):
    if getattr(self, '_accounting_sync_muted', False):
        return None
    cfg = self._acc_get_cfg()
    if not cfg.enabled or not bool(_acc_cfg_get_v67(cfg, 'live_sync_enabled', True)):
        return None
    if not bool(_acc_cfg_get_v67(cfg, 'live_sync_delete_remote_on_local_delete', True)):
        return None
    conn = self._acc_connector()
    res = conn.delete_product(item)
    _acc_log(self, 'accounting_live_delete_product', {'ok': getattr(res, 'ok', False), 'sku': getattr(item, 'sku', ''), 'reason': reason, 'status': getattr(res, 'response_status', 0)})
    st = _acc_live_state_get_v67(self)
    key = _acc_item_remote_key_v67(item)
    if key:
        st.setdefault('tombstones', {})[key] = _acc_now_ts()
        _acc_live_state_save_v67(self)
    return res


def _acc_attach_inventory_owner_v67(self):
    try:
        db = getattr(self, '_noliktava', None)
        if db is not None:
            setattr(db, '_accounting_owner', self)
    except Exception:
        pass


def _acc_poll_remote_inventory_v67(self):
    if getattr(self, '_accounting_poll_in_progress', False):
        return
    self._accounting_poll_in_progress = True
    try:
        _acc_pull_remote_inventory_now_v67(self, silent=True)
    except Exception as e:
        _acc_log(self, 'accounting_live_poll_error', {'error': str(e)})
    finally:
        self._accounting_poll_in_progress = False


def _acc_start_live_sync_v67(self):
    cfg = self._acc_get_cfg()
    if not cfg.enabled or not bool(_acc_cfg_get_v67(cfg, 'live_sync_enabled', True)):
        return
    _acc_attach_inventory_owner_v67(self)
    timer = getattr(self, '_accounting_live_sync_timer', None)
    if timer is None:
        timer = QTimer(self)
        timer.timeout.connect(lambda: _acc_poll_remote_inventory_v67(self))
        self._accounting_live_sync_timer = timer
    interval_ms = max(int(float(_acc_cfg_get_v67(cfg, 'live_sync_poll_seconds', 15) or 15) * 1000), 3000)
    timer.start(interval_ms)


def _acc_stop_live_sync_v67(self):
    timer = getattr(self, '_accounting_live_sync_timer', None)
    if timer is not None:
        timer.stop()


def _acc_sync_inventory_products_v67(self):
    try:
        conn = self._acc_connector()
        db = getattr(self, '_noliktava', None)
        if db is None:
            raise AccountingIntegrationError('Noliktava nav inicializēta.')
        items = list(getattr(db, 'items', []) or [])
        if not items:
            raise AccountingIntegrationError('Noliktavā nav preču, ko sinhronizēt.')
        results = []
        for it in items:
            _acc_prepare_local_item_for_push_v67(self, it)
            results.append(conn.upsert_product(it))
        _acc_log(self, 'accounting_sync_inventory', {'count': len(results)})
        _acc_show_result_v67(self, results, 'Noliktavas preču sinhronizācija')
    except Exception as e:
        _acc_log(self, 'accounting_sync_inventory_error', {'error': str(e)})
        QMessageBox.warning(self, 'Kļūda', f'Noliktavas preču sinhronizācija neizdevās.\n\n{e}')


def _acc_sync_everything_v67(self):
    try:
        _acc_sync_inventory_products_v67(self)
        _acc_pull_remote_inventory_now_v67(self, silent=False)
    except Exception as e:
        _acc_log(self, 'accounting_sync_everything_error', {'error': str(e)})
        QMessageBox.warning(self, 'Kļūda', f'Pilnā sinhronizācija neizdevās.\n\n{e}')


def _acc_build_menu_v67(self):
    try:
        menubar = self.menuBar()
        existing = getattr(self, '_accounting_menu', None)
        if existing is not None:
            try:
                existing.clear()
            except Exception:
                return
            m = existing
        else:
            m = menubar.addMenu('Grāmatvedība')
            self._accounting_menu = m
        actions = [
            ('Integrācijas iestatījumi…', _acc_open_settings_dialog),
            ('Pārbaudīt pieslēgumu', _acc_test_connection),
            None,
            ('Sinhronizēt preces no šīs programmas uz grāmatvedību', _acc_sync_inventory_products_v67),
            ('Ielādēt preces no grāmatvedības uz šo programmu', _acc_pull_remote_inventory_now_v67),
            ('Pilnā produktu sinhronizācija abos virzienos', _acc_sync_everything_v67),
            None,
            ('Sinhronizēt pašreizējo dokumentu', _acc_sync_current_document),
            ('Sinhronizēt partnerus no dokumenta', _acc_sync_partners_from_current_doc),
            ('Nolasīt partnerus no attālās sistēmas', _acc_fetch_remote_partners),
            ('Nolasīt preces no attālās sistēmas', _acc_fetch_remote_products_v67),
            None,
            ('Ieslēgt dzīvo sinhronizāciju', _acc_start_live_sync_v67),
            ('Izslēgt dzīvo sinhronizāciju', _acc_stop_live_sync_v67),
        ]
        for spec in actions:
            if spec is None:
                m.addSeparator()
                continue
            title, fn = spec
            a = QAction(title, self)
            a.triggered.connect(lambda _=False, f=fn: f(self))
            m.addAction(a)
    except Exception as e:
        print(f'Accounting menu build failed: {e}')


# saglabājam oriģinālās NoliktavaDB metodes pirms live-sync monkey-patchiem
_NOLIKTAVA_V66_OLD_UPSERT = getattr(NoliktavaDB, 'upsert')
_NOLIKTAVA_V66_OLD_DELETE_IDS = getattr(NoliktavaDB, 'delete_by_inventory_ids')
_NOLIKTAVA_V66_OLD_DELETE_ROW = getattr(NoliktavaDB, 'delete_by_row')
_NOLIKTAVA_V66_OLD_DELETE_SKU = getattr(NoliktavaDB, 'delete_by_sku')

def _acc_after_init_v67(self):
    try:
        self._accounting_cfg = _acc_cfg_from_settings()
    except Exception:
        self._accounting_cfg = AccountingIntegrationConfig()
    cfg = self._accounting_cfg
    defaults = {
        'live_sync_enabled': True,
        'live_sync_poll_seconds': 15,
        'live_sync_fetch_limit': 5000,
        'live_sync_remove_missing_local_products': False,
        'live_sync_push_local_changes': True,
        'live_sync_delete_remote_on_local_delete': True,
    }
    for k, v in defaults.items():
        if not hasattr(cfg, k):
            setattr(cfg, k, v)
    self._accounting_sync_muted = False
    self._accounting_poll_in_progress = False
    _acc_attach_inventory_owner_v67(self)
    _acc_build_menu_v67(self)
    _acc_start_live_sync_v67(self)
    QTimer.singleShot(1200, lambda: _acc_poll_remote_inventory_v67(self))


def _akta_v67_init_wrapper(self, *args, **kwargs):
    _AKTA_V64_OLD_INIT(self, *args, **kwargs)
    _acc_after_init_v67(self)


def _noliktava_upsert_v67(self, item):
    base_upsert = globals().get('_NOLIKTAVA_V66_OLD_UPSERT')
    if base_upsert is None:
        raise NameError('NoliktavaDB oriģinālā upsert metode nav atrasta')
    base_upsert(self, item)
    owner = getattr(self, '_accounting_owner', None)
    if owner is not None:
        try:
            _acc_push_single_local_item_v67(owner, item, reason='local_upsert')
        except Exception as e:
            _acc_log(owner, 'accounting_live_push_product_error', {'error': str(e), 'sku': getattr(item, 'sku', '')})


def _noliktava_delete_by_inventory_ids_v67(self, inventory_ids):
    ids = [(x or '').strip() for x in (inventory_ids or []) if (x or '').strip()]
    removed_items = []
    for inv_id in ids:
        try:
            it = self.get_by_inventory_id(inv_id) if hasattr(self, 'get_by_inventory_id') else None
            if it is not None:
                removed_items.append(NoliktavasPrece(**_dc_asdict_v67(it)))
        except Exception:
            pass
    base_delete_ids = globals().get('_NOLIKTAVA_V66_OLD_DELETE_IDS')
    if base_delete_ids is None:
        raise NameError('NoliktavaDB oriģinālā delete_by_inventory_ids metode nav atrasta')
    removed = base_delete_ids(self, inventory_ids)
    owner = getattr(self, '_accounting_owner', None)
    if owner is not None:
        for it in removed_items:
            try:
                _acc_delete_remote_item_v67(owner, it, reason='local_delete')
            except Exception as e:
                _acc_log(owner, 'accounting_live_delete_product_error', {'error': str(e), 'sku': getattr(it, 'sku', '')})
    return removed


def _noliktava_delete_by_row_v67(self, row: int):
    it = None
    try:
        if 0 <= int(row) < len(getattr(self, 'items', []) or []):
            it = NoliktavasPrece(**_dc_asdict_v67(self.items[int(row)]))
    except Exception:
        it = None
    base_delete_row = globals().get('_NOLIKTAVA_V66_OLD_DELETE_ROW')
    if base_delete_row is None:
        raise NameError('NoliktavaDB oriģinālā delete_by_row metode nav atrasta')
    removed = base_delete_row(self, row)
    owner = getattr(self, '_accounting_owner', None)
    if owner is not None and it is not None:
        try:
            _acc_delete_remote_item_v67(owner, it, reason='local_delete_row')
        except Exception as e:
            _acc_log(owner, 'accounting_live_delete_product_error', {'error': str(e), 'sku': getattr(it, 'sku', '')})
    return removed


def _noliktava_delete_by_sku_v67(self, sku: str):
    it = None
    try:
        for x in list(getattr(self, 'items', []) or []):
            if _acc_clean(getattr(x, 'sku', '')) == _acc_clean(sku):
                it = NoliktavasPrece(**_dc_asdict_v67(x))
                break
    except Exception:
        it = None
    base_delete_sku = globals().get('_NOLIKTAVA_V66_OLD_DELETE_SKU')
    if base_delete_sku is None:
        raise NameError('NoliktavaDB oriģinālā delete_by_sku metode nav atrasta')
    removed = base_delete_sku(self, sku)
    owner = getattr(self, '_accounting_owner', None)
    if owner is not None and it is not None:
        try:
            _acc_delete_remote_item_v67(owner, it, reason='local_delete_sku')
        except Exception as e:
            _acc_log(owner, 'accounting_live_delete_product_error', {'error': str(e), 'sku': getattr(it, 'sku', '')})
    return removed


# aktivizējam salabotos monkey-patchus PIRMS __main__
JumisConnector._post_import = _jumis_post_import_v67
JumisConnector._post_export = _jumis_post_export_v67
JumisConnector.test_connection = _jumis_test_connection_v67
JumisConnector.fetch_partners = _jumis_fetch_partners_v67
JumisConnector.fetch_products = _jumis_fetch_products_v67
JumisConnector.upsert_product = _jumis_upsert_product_v67
JumisConnector.delete_product = _jumis_delete_product_v67
BaseAccountingConnector.delete_product = lambda self, item: (_ for _ in ()).throw(NotImplementedError)
GenericRestConnector.delete_product = lambda self, item: AccountingSyncResult(ok=False, operation='delete_product', system=getattr(self, 'system_name', 'Generic REST'), request_payload='', response_status=501, response_text='Delete operation is not configured for Generic REST connector.')

AktaLogs.__init__ = _akta_v67_init_wrapper
AktaLogs._acc_show_result = _acc_show_result_v67
AktaLogs._acc_fetch_remote_products = _acc_fetch_remote_products_v67
AktaLogs._acc_pull_remote_inventory_now = _acc_pull_remote_inventory_now_v67
AktaLogs._acc_start_live_sync = _acc_start_live_sync_v67
AktaLogs._acc_stop_live_sync = _acc_stop_live_sync_v67
AktaLogs._acc_build_menu = _acc_build_menu_v67
AktaLogs._acc_sync_inventory_products = _acc_sync_inventory_products_v67
AktaLogs._acc_sync_everything = _acc_sync_everything_v67
AktaLogs._acc_delete_remote_item = _acc_delete_remote_item_v67
AktaLogs._acc_push_single_local_item = _acc_push_single_local_item_v67
NoliktavaDB.upsert = _noliktava_upsert_v67
NoliktavaDB.delete_by_inventory_ids = _noliktava_delete_by_inventory_ids_v67
NoliktavaDB.delete_by_row = _noliktava_delete_by_row_v67
NoliktavaDB.delete_by_sku = _noliktava_delete_by_sku_v67


# ==============================
# v68: precizēts Jumis pieslēgums un preču XML struktūra.
# - test_connection izmanto oficiālajam paraugam tuvāku export XML
# - product upsert izmanto ligzdotu ProductPrice/ProductWarehouse struktūru
# - fetch_products mēģina vairākas XML veidnes un robustāk parsē ligzdotus laukus
# - pie auth kļūdām rāda skaidrāku diagnostiku
# ==============================

def _acc_hintify_jumis_message_v68(msg: str, cfg=None) -> str:
    base = _acc_clean(msg)
    if not base:
        return ''
    low = base.lower()
    if 'neizdevās pieslēgties' in low or 'authentication failed' in low or 'unauthorized' in low:
        hints = [
            'Jumis atgrieza autentifikācijas kļūdu. Pārbaudi, vai ievadīta tieši speciālā parole, nevis parastā lietotāja parole.',
            'API pieslēgumam jānorāda datubāzes nosaukums, nevis serveris/ports vai pilna SQL pieslēguma rinda.',
            'Lietotājvārdam jābūt tam pašam Jumis mākoņa e-pastam, kuram izveidota speciālā parole.',
        ]
        try:
            db = _acc_clean(getattr(cfg, 'jumis_database', ''))
            if db and any(x in db for x in [';', '=', ',', '\\', '/']):
                hints.append('Datubāzes lauks izskatās pēc pilnas pieslēguma rindas. API parasti vajag tikai pašu datubāzes nosaukumu.')
        except Exception:
            pass
        return base + '\n\n' + '\n'.join(f'- {x}' for x in hints)
    return base


def _acc_response_is_success_v68(status_code: int, text: str, cfg=None):
    ok, msg = _acc_response_is_success_v67(status_code, text)
    return ok, _acc_hintify_jumis_message_v68(msg or '', cfg)


def _acc_result_with_status_v68(self, res):
    ok, msg = _acc_response_is_success_v68(getattr(res, 'response_status', 0), getattr(res, 'response_text', ''), getattr(self, 'config', None))
    res.ok = ok
    extra = dict(getattr(res, 'extra', {}) or {})
    if msg:
        extra['message'] = msg
    res.extra = extra
    return res


def _jumis_post_import_v68(self, xmlrequest: str, operation: str):
    payload = self._credentials_payload(xmlrequest)
    r = self.session.post(
        _acc_clean(self.config.jumis_import_url) or JUMIS_IMPORT_URL,
        json=payload,
        timeout=max(int(getattr(self.config, 'jumis_timeout_sec', 45) or 45), 10),
        headers={'Accept': 'application/json', 'Content-Type': 'application/json; charset=utf-8'}
    )
    return _acc_result_with_status_v68(self, AccountingSyncResult(ok=r.ok, operation=operation, system=self.system_name, request_payload=xmlrequest, response_status=r.status_code, response_text=r.text))


def _jumis_post_export_v68(self, xmlrequest: str, operation: str):
    payload = self._credentials_payload(xmlrequest)
    r = self.session.post(
        _acc_clean(self.config.jumis_export_url) or JUMIS_EXPORT_URL,
        json=payload,
        timeout=max(int(getattr(self.config, 'jumis_timeout_sec', 45) or 45), 10),
        headers={'Accept': 'application/json', 'Content-Type': 'application/json; charset=utf-8'}
    )
    return _acc_result_with_status_v68(self, AccountingSyncResult(ok=r.ok, operation=operation, system=self.system_name, request_payload=xmlrequest, response_status=r.status_code, response_text=r.text))


def _acc_nested_child_text_v68(el, path_options, default=''):
    if el is None:
        return default
    if isinstance(path_options, str):
        path_options = [path_options]
    for path in path_options:
        cur = el
        ok = True
        for part in [p for p in str(path).split('/') if p]:
            found = None
            for ch in list(cur):
                if _acc_local_name_v67(getattr(ch, 'tag', '')) == part:
                    found = ch
                    break
            if found is None:
                ok = False
                break
            cur = found
        if ok:
            txt = (getattr(cur, 'text', '') or '').strip()
            if txt:
                return txt
    return default


def _acc_parse_jumis_products_response_v68(text: str):
    products = []
    payload = _acc_json_from_response_text_v67(text or '')
    if payload is not None:
        for row in _acc_collect_table_rows_v67(payload, []):
            has_product_shape = any(k in row for k in ['ProductCode', 'ProductName', 'ProductBarCode', 'ProductUnit', 'Price', 'ProductComments', 'WarehouseName'])
            if not has_product_shape:
                continue
            price = row.get('ProductPrice') or row.get('Price') or row.get('SalePrice') or '0'
            qty = row.get('StockQuantity') or row.get('Quantity') or row.get('Balance') or row.get('WarehouseNecessaryQuantity') or '0'
            products.append({
                'ProductCode': row.get('ProductCode', ''),
                'ProductName': row.get('ProductName', ''),
                'ProductBarCode': row.get('ProductBarCode', ''),
                'ProductUnit': row.get('ProductUnit', row.get('ProductUnitCode', 'gab.')),
                'ProductPrice': price,
                'ProductPurchasePrice': row.get('ProductPurchasePrice', row.get('PurchasePrice', '')),
                'ProductVATRate': row.get('ProductVATRate', row.get('VatRate', '')),
                'ProductGroupName': row.get('ProductGroupName', row.get('ProductClassName', row.get('GroupName', ''))),
                'ProductDescription': row.get('ProductDescription', row.get('ProductComments', row.get('Description', ''))),
                'StockQuantity': qty,
                'WarehouseName': row.get('WarehouseName', ''),
                'Status': row.get('Status', row.get('ProductStatus', '')),
            })
    xml_text = _acc_extract_xml_from_response_text_v67(text)
    root = _acc_parse_xml_root_v67(xml_text)
    if root is not None:
        row_like_tags = {'Product', 'Row', 'ProductRow', 'Item', 'Record'}
        for p in _acc_findall_any_v67(root, row_like_tags):
            code = _acc_child_text_v67(p, ['ProductCode', 'Code'])
            name = _acc_child_text_v67(p, ['ProductName', 'Name'])
            barcode = _acc_child_text_v67(p, ['ProductBarCode', 'BarCode', 'Barcode'])
            unit = _acc_child_text_v67(p, ['ProductUnit', 'Unit', 'UnitCode'], 'gab.')
            price = _acc_nested_child_text_v68(p, ['ProductPrice/Price']) or _acc_child_text_v67(p, ['ProductPrice', 'Price'], '0')
            qty = _acc_nested_child_text_v68(p, ['ProductWarehouse/WarehouseNecessaryQuantity', 'ProductWarehouse/WarehouseQuantity']) or _acc_child_text_v67(p, ['StockQuantity', 'Quantity', 'Balance', 'WarehouseNecessaryQuantity'], '0')
            group_name = _acc_child_text_v67(p, ['ProductGroupName', 'ProductClassName', 'GroupName'])
            descr = _acc_child_text_v67(p, ['ProductDescription', 'ProductComments', 'Description'])
            warehouse_name = _acc_nested_child_text_v68(p, ['ProductWarehouse/WarehouseName']) or _acc_child_text_v67(p, ['WarehouseName'])
            if not any([code, name, barcode, price, qty, group_name, descr]):
                continue
            products.append({
                'ProductCode': code,
                'ProductName': name,
                'ProductBarCode': barcode,
                'ProductUnit': unit,
                'ProductPrice': price,
                'ProductPurchasePrice': _acc_nested_child_text_v68(p, ['ProductPurchasePrice/Price']) or _acc_child_text_v67(p, ['ProductPurchasePrice', 'PurchasePrice']),
                'ProductVATRate': _acc_child_text_v67(p, ['ProductVATRate', 'VatRate']),
                'ProductGroupName': group_name,
                'ProductDescription': descr,
                'StockQuantity': qty,
                'WarehouseName': warehouse_name,
                'Status': _acc_child_text_v67(p, ['Status', 'ProductStatus']),
            })
    dedup = {}
    for p in products:
        key = (_acc_clean(p.get('ProductCode')), _acc_clean(p.get('ProductBarCode')), _acc_clean(p.get('ProductName')))
        if any(key):
            dedup[key] = p
    return list(dedup.values())


def _jumis_test_connection_v68(self):
    # Tuvāk oficiālajam eksporta paraugam: bez tjData, tikai vienkāršs Partner Read.
    xml_candidates = [
        '<?xml version="1.0" ?><dataroot><tjDocument Version="TJ5.5.101"/><tjRequest Name="Partner" Operation="Read" Version="TJ7.0.112" Structure="Tree"><tjFields><Field Name="PartnerKindName"/><Field Name="PartnerName"/></tjFields></tjRequest></dataroot>',
        '<?xml version="1.0" ?><dataroot><tjDocument Version="TJ5.5.101"/><tjRequest Name="Partner" Operation="Read" Version="TJ7.0.112" Structure="Sheet"><tjFields><Field Name="PartnerName"/></tjFields></tjRequest></dataroot>',
    ]
    last = None
    for i, xmlrequest in enumerate(xml_candidates, start=1):
        res = self._post_export(xmlrequest, f'test_connection_try_{i}')
        extra = dict(getattr(res, 'extra', {}) or {})
        extra['partners'] = _acc_parse_jumis_partners_response_v67(res.response_text)
        res.extra = extra
        last = res
        if getattr(res, 'ok', False):
            res.operation = 'test_connection'
            return res
    if last is not None:
        last.operation = 'test_connection'
        return last
    return AccountingSyncResult(ok=False, operation='test_connection', system=self.system_name, request_payload='', response_status=0, response_text='Testa pieprasījumu neizdevās izpildīt.')


def _jumis_fetch_products_v68(self, limit: int = 5000):
    fields_variants = [
        # Plakans variants: bieži strādā eksporta veidnēs.
        '<Field Name="ProductCode"/><Field Name="ProductName"/><Field Name="ProductBarCode"/><Field Name="ProductUnit"/><Field Name="Price"/><Field Name="ProductComments"/><Field Name="ProductClassName"/><Field Name="WarehouseName"/><Field Name="WarehouseNecessaryQuantity"/>',
        '<Field Name="ProductCode"/><Field Name="ProductName"/><Field Name="ProductBarCode"/><Field Name="ProductUnit"/><Field Name="ProductPrice"/><Field Name="ProductComments"/><Field Name="ProductClassName"/><Field Name="ProductWarehouse"/>',
        '<Field Name="ProductCode"/><Field Name="ProductName"/><Field Name="ProductBarCode"/><Field Name="ProductUnit"/><Field Name="ProductPrice"/><Field Name="Quantity"/><Field Name="Status"/>',
    ]
    xml_candidates = []
    for structure in ('Tree', 'Sheet'):
        for fields in fields_variants:
            xml_candidates.append(
                f'<?xml version="1.0" ?><dataroot><tjDocument Version="TJ5.5.101"/><tjRequest Name="Product" Operation="Read" Version="TJ5.5.101" Structure="{structure}"><tjFields>{fields}</tjFields></tjRequest></dataroot>'
            )
            xml_candidates.append(
                f'<?xml version="1.0" ?><dataroot><tjDocument Version="TJ5.5.101"/><tjRequest Name="Product" Operation="Read" Version="TJ5.5.101" Structure="{structure}"><tjData Read="All"/><tjFields>{fields}</tjFields></tjRequest></dataroot>'
            )
    best = None
    best_products = []
    for i, xmlrequest in enumerate(xml_candidates, start=1):
        res = self._post_export(xmlrequest, f'fetch_products_try_{i}')
        products = _acc_parse_jumis_products_response_v68(res.response_text)
        extra = dict(getattr(res, 'extra', {}) or {})
        extra['products'] = products[:max(int(limit or 5000), 1)]
        res.extra = extra
        if best is None or (len(products) > len(best_products)):
            best = res
            best_products = products
        if getattr(res, 'ok', False) and products:
            break
        if not getattr(res, 'ok', False):
            # ja autentifikācija nenostrādā, nav jēgas turpināt pārējos XML variantus
            txt = (getattr(res, 'response_text', '') or '').lower()
            if 'neizdevās pieslēgties' in txt or 'unauthorized' in txt or 'authentication failed' in txt:
                best = res
                best_products = []
                break
    final_res = best or AccountingSyncResult(ok=False, operation='fetch_products', system=self.system_name, request_payload='', response_status=0, response_text='Neizdevās iegūt atbildi no Jumis.')
    final_res.operation = 'fetch_products'
    extra = dict(getattr(final_res, 'extra', {}) or {})
    # dedupe
    dedup = {}
    for p in (best_products or []):
        key = (_acc_clean(p.get('ProductCode')), _acc_clean(p.get('ProductBarCode')), _acc_clean(p.get('ProductName')))
        if any(key):
            dedup[key] = p
    products = list(dedup.values())[:max(int(limit or 5000), 1)]
    extra['products'] = products
    extra['limit'] = limit
    if not products and getattr(final_res, 'ok', False):
        extra['message'] = (extra.get('message') or 'Jumis atbildēja, bet preču saraksts netika atrasts vai ir tukšs.')
    final_res.extra = extra
    return final_res


def _jumis_product_xml_v68(self, item):
    it = item or NoliktavasPrece()
    code = _acc_clean(getattr(it, 'sku', '')) or _acc_clean(getattr(it, 'inventory_id', '')) or _acc_clean(getattr(it, 'svitrkods', '')) or f'ITEM-{secrets.token_hex(4)}'
    try:
        if not _acc_clean(getattr(it, 'sku', '')):
            it.sku = code
    except Exception:
        pass
    try:
        if not _acc_clean(getattr(it, 'inventory_id', '')):
            it.inventory_id = code
    except Exception:
        pass
    price = _acc_to_amount(getattr(it, 'cena', '0'))
    qty = _acc_to_qty(getattr(it, 'atlikums', 0))
    vat_rate = _acc_to_amount(getattr(it, 'pvn_likme', '21'), default='21.00')
    return (
        '<?xml version="1.0" encoding="utf-8" ?>'
        '<dataroot><tjDocument Version="TJ5.5.101"/>'
        '<tjResponse Name="Product" Operation="Insert" Version="TJ5.5.101" Structure="Tree">'
        '<Product>'
        f'{_acc_xml("ProductCode", code)}'
        f'{_acc_xml("ProductName", _acc_clean(getattr(it, "nosaukums", "")) or code)}'
        f'{_acc_xml("ProductBarCode", _acc_clean(getattr(it, "svitrkods", "")))}'
        f'{_acc_xml("ProductUnit", _acc_clean(getattr(it, "vieniba", "")) or "gab.")}'
        f'{_acc_xml("ProductComments", _acc_clean(getattr(it, "piezimes", "")))}'
        f'{_acc_xml("ProductClassName", _acc_clean(getattr(it, "kategorija", "")))}'
        f'{_acc_xml("ProductCnCode", _acc_clean(getattr(it, "hs_kods", "")))}'
        f'{_acc_xml("ProductOriginCountryCode", _acc_clean(getattr(it, "izcelsmes_valsts", "")))}'
        f'{_acc_xml("ProductVATRate", vat_rate)}'
        '<ProductPrice>'
        f'{_acc_xml("Price", price)}'
        f'{_acc_xml("PriceCurrency", _acc_clean(getattr(it, "iepirkuma_valuta", "")) or "EUR")}'
        '</ProductPrice>'
        '<ProductWarehouse>'
        f'{_acc_xml("WarehouseName", _acc_clean(getattr(it, "noliktavas_nosaukums", "")) or "Pamatnoliktava")}'
        f'{_acc_xml("WarehouseNecessaryQuantity", qty)}'
        f'{_acc_xml("WarehouseQuantityUnit", _acc_clean(getattr(it, "vieniba", "")) or "gab.")}'
        '</ProductWarehouse>'
        '</Product></tjResponse></dataroot>'
    )


def _jumis_upsert_product_v68(self, item):
    return self._post_import(_jumis_product_xml_v68(self, item), 'upsert_product')


# Aktivizējam v68 patchus
JumisConnector._post_import = _jumis_post_import_v68
JumisConnector._post_export = _jumis_post_export_v68
JumisConnector.test_connection = _jumis_test_connection_v68
JumisConnector.fetch_products = _jumis_fetch_products_v68
JumisConnector.upsert_product = _jumis_upsert_product_v68

# ==============================
# v69 PATCH: Jumis URL normalizācija + skaidrāka autentifikācijas diagnostika
# Svarīgi: ja Jumis pats atgriež "Neizdevās pieslēgties!", sinhronizācija nevar strādāt,
# tāpēc šeit fokuss ir uz drošu URL normalizāciju un nepārprotamu kļūdas avota noteikšanu.
# ==============================

def _acc_normalize_http_url_v69(url: str, default_url: str = '') -> str:
    raw = _acc_clean(url) or _acc_clean(default_url)
    if not raw:
        return ''
    if raw.startswith('://'):
        raw = 'https' + raw
    elif raw.startswith('//'):
        raw = 'https:' + raw
    elif '://' not in raw:
        raw = 'https://' + raw.lstrip('/')
    return raw


def _acc_mask_secret_v69(value: str, show: int = 4) -> str:
    s = _acc_clean(value)
    if not s:
        return ''
    if len(s) <= show:
        return '*' * len(s)
    return '*' * (len(s) - show) + s[-show:]


def _acc_result_with_status_v69(self, res):
    ok, msg = _acc_response_is_success_v68(getattr(res, 'response_status', 0), getattr(res, 'response_text', ''), getattr(self, 'config', None))
    res.ok = ok
    extra = dict(getattr(res, 'extra', {}) or {})
    cfg = getattr(self, 'config', None)
    if cfg is not None:
        extra['normalized_import_url'] = _acc_normalize_http_url_v69(getattr(cfg, 'jumis_import_url', ''), JUMIS_IMPORT_URL)
        extra['normalized_export_url'] = _acc_normalize_http_url_v69(getattr(cfg, 'jumis_export_url', ''), JUMIS_EXPORT_URL)
        extra['database_preview'] = _acc_clean(getattr(cfg, 'jumis_database', ''))
        extra['username_preview'] = _acc_clean(getattr(cfg, 'jumis_username', ''))
        extra['apikey_preview'] = _acc_mask_secret_v69(getattr(cfg, 'jumis_api_key', ''))
    if msg:
        extra['message'] = msg
    txt = _acc_clean(getattr(res, 'response_text', ''))
    low = txt.lower()
    if 'neizdevās pieslēgties' in low:
        extra['auth_failed'] = True
        extra['message'] = (
            'Jumis serviss sasniedzams, bet autentifikācija neizdevās. '
            'Tā nav produktu parsera kļūda — kamēr šī kļūda nav novērsta, produktu sinhronizācija nestrādās.\n\n'
            'Pārbaudi tieši šos laukus Jumis pusē:\n'
            '1) speciālā parole (garā parole no Lietotāja iestatījumi),\n'
            '2) datubāzes nosaukums no Uzņēmumu pārvaldība / Informācija,\n'
            '3) tas pats Jumis mākoņa e-pasts, kuram izveidota speciālā parole.'
        )
    res.extra = extra
    return res


def _jumis_credentials_payload_v69(self, xmlrequest: str) -> dict:
    cfg = self.config
    username = _acc_clean(cfg.jumis_username)
    password = _acc_clean(cfg.jumis_password)
    database = _acc_clean(cfg.jumis_database)
    if not username or not password or not database:
        raise AccountingIntegrationError('Jumis pieslēgumam obligāti jānorāda lietotājvārds, speciālā parole un datubāzes nosaukums.')
    return {
        'username': username,
        'password': password,
        'database': database,
        'apikey': _acc_clean(cfg.jumis_api_key) or JUMIS_DEFAULT_API_KEY,
        'XMLrequest': xmlrequest,
    }


def _jumis_post_import_v69(self, xmlrequest: str, operation: str):
    payload = _jumis_credentials_payload_v69(self, xmlrequest)
    url = _acc_normalize_http_url_v69(_acc_clean(self.config.jumis_import_url), JUMIS_IMPORT_URL)
    r = self.session.post(
        url,
        json=payload,
        timeout=max(int(getattr(self.config, 'jumis_timeout_sec', 45) or 45), 10),
        headers={'Accept': 'application/json', 'Content-Type': 'application/json; charset=utf-8'}
    )
    return _acc_result_with_status_v69(self, AccountingSyncResult(ok=r.ok, operation=operation, system=self.system_name, request_payload=xmlrequest, response_status=r.status_code, response_text=r.text))


def _jumis_post_export_v69(self, xmlrequest: str, operation: str):
    payload = _jumis_credentials_payload_v69(self, xmlrequest)
    url = _acc_normalize_http_url_v69(_acc_clean(self.config.jumis_export_url), JUMIS_EXPORT_URL)
    r = self.session.post(
        url,
        json=payload,
        timeout=max(int(getattr(self.config, 'jumis_timeout_sec', 45) or 45), 10),
        headers={'Accept': 'application/json', 'Content-Type': 'application/json; charset=utf-8'}
    )
    return _acc_result_with_status_v69(self, AccountingSyncResult(ok=r.ok, operation=operation, system=self.system_name, request_payload=xmlrequest, response_status=r.status_code, response_text=r.text))


def _jumis_test_connection_v69(self):
    # Paturam ļoti vienkāršu eksportu. Ja šeit ir "Neizdevās pieslēgties!", tā ir autentifikācija.
    xmlrequest = ('<?xml version="1.0" ?>'
                  '<dataroot>'
                  '<tjDocument Version="TJ5.5.101"/>'
                  '<tjRequest Name="Partner" Operation="Read" Version="TJ7.0.112" Structure="Tree">'
                  '<tjFields><Field Name="PartnerName"/></tjFields>'
                  '</tjRequest>'
                  '</dataroot>')
    res = self._post_export(xmlrequest, 'test_connection')
    extra = dict(getattr(res, 'extra', {}) or {})
    extra['partners'] = _acc_parse_jumis_partners_response_v67(getattr(res, 'response_text', '') or '')
    if not getattr(res, 'ok', False):
        extra['next_step'] = 'Novērs autentifikācijas kļūdu Jumis pusē. Tikai pēc tam var strādāt produktu sinhronizācija.'
    res.extra = extra
    return res


def _acc_save_accounting_settings_v69(self):
    cfg = getattr(self, 'cfg', None)
    if cfg is None:
        cfg = load_accounting_integration_config()
    cfg.enabled = bool(self.chk_enabled.isChecked())
    cfg.provider = self.cmb_provider.currentData() or self.cmb_provider.currentText() or 'jumis'
    cfg.auto_sync_on_pdf = bool(self.chk_pdf.isChecked())
    cfg.auto_sync_on_docx = bool(self.chk_docx.isChecked())
    cfg.auto_sync_on_project_save = bool(self.chk_project_save.isChecked())
    cfg.live_sync_enabled = bool(self.chk_live_sync.isChecked())
    try:
        cfg.live_sync_interval_sec = max(15, int(self.in_live_sync_interval.text() or '60'))
    except Exception:
        cfg.live_sync_interval_sec = 60
    try:
        cfg.jumis_timeout_sec = max(10, int(self.in_timeout.text() or '45'))
    except Exception:
        cfg.jumis_timeout_sec = 45
    cfg.jumis_username = _acc_clean(self.in_j_username.text())
    cfg.jumis_password = _acc_clean(self.in_j_password.text())
    cfg.jumis_database = _acc_clean(self.in_j_database.text())
    cfg.jumis_api_key = _acc_clean(self.in_j_api_key.text()) or JUMIS_DEFAULT_API_KEY
    cfg.jumis_import_url = _acc_normalize_http_url_v69(self.in_j_import_url.text(), JUMIS_IMPORT_URL)
    cfg.jumis_export_url = _acc_normalize_http_url_v69(self.in_j_export_url.text(), JUMIS_EXPORT_URL)
    cfg.generic_base_url = _acc_clean(self.in_g_base.text())
    cfg.generic_token = _acc_clean(self.in_g_token.text())
    cfg.generic_partners_endpoint = _acc_clean(self.in_g_partners.text())
    cfg.generic_products_endpoint = _acc_clean(self.in_g_products.text())
    cfg.generic_store_docs_endpoint = _acc_clean(self.in_g_store_docs.text())
    cfg.generic_financial_docs_endpoint = _acc_clean(self.in_g_fin_docs.text())
    save_accounting_integration_config(cfg)
    self.cfg = cfg
    try:
        self.in_j_import_url.setText(cfg.jumis_import_url)
        self.in_j_export_url.setText(cfg.jumis_export_url)
    except Exception:
        pass
    QMessageBox.information(self, 'Saglabāts', 'Iestatījumi saglabāti.')


# Aktivizējam v69 patchus
JumisConnector._credentials_payload = _jumis_credentials_payload_v69
JumisConnector._post_import = _jumis_post_import_v69
JumisConnector._post_export = _jumis_post_export_v69
JumisConnector.test_connection = _jumis_test_connection_v69
# v70: salabots neesošas klases nosaukums, lai aplikācija startētos korekti
try:
    AccountingIntegrationDialog._save = _acc_save_accounting_settings_v69
except Exception:
    pass



# ==============================
# v71 PATCH: automātiska Jumis URL salabošana + skaidrāks speciālās paroles paskaidrojums
# ==============================

def _acc_cfg_from_settings_v71() -> AccountingIntegrationConfig:
    st = load_settings() or {}
    raw = st.get(ACCOUNTING_SETTINGS_KEY) or {}
    if not isinstance(raw, dict):
        raw = {}
    cfg = AccountingIntegrationConfig()
    for k in cfg.__dataclass_fields__.keys():
        if k in raw:
            setattr(cfg, k, raw.get(k))
    try:
        cfg.jumis_import_url = _acc_normalize_http_url_v69(getattr(cfg, 'jumis_import_url', ''), JUMIS_IMPORT_URL)
        cfg.jumis_export_url = _acc_normalize_http_url_v69(getattr(cfg, 'jumis_export_url', ''), JUMIS_EXPORT_URL)
    except Exception:
        pass
    return cfg


def _acc_dialog_get_config_v71(self):
    cfg = AccountingIntegrationConfig()
    cfg.enabled = self.ck_enabled.isChecked()
    cfg.profile_name = _acc_clean(self.in_profile_name.text()) or 'Noklusējuma profils'
    cfg.system_type = self.cmb_system_type.currentData() or 'jumis'

    cfg.jumis_username = _acc_clean(self.in_j_username.text())
    cfg.jumis_password = _acc_clean(self.in_j_password.text())
    cfg.jumis_database = _acc_clean(self.in_j_database.text())
    cfg.jumis_api_key = _acc_clean(self.in_j_apikey.text()) or JUMIS_DEFAULT_API_KEY
    cfg.jumis_import_url = _acc_normalize_http_url_v69(self.in_j_import_url.text(), JUMIS_IMPORT_URL)
    cfg.jumis_export_url = _acc_normalize_http_url_v69(self.in_j_export_url.text(), JUMIS_EXPORT_URL)

    cfg.generic_base_url = _acc_clean(self.in_g_base_url.text())
    cfg.generic_auth_type = self.cmb_g_auth.currentData() or 'bearer'
    cfg.generic_username = _acc_clean(self.in_g_username.text())
    cfg.generic_password = _acc_clean(self.in_g_password.text())
    cfg.generic_token = _acc_clean(self.in_g_token.text())
    cfg.generic_api_key_header = _acc_clean(self.in_g_key_header.text()) or 'X-API-Key'
    cfg.generic_api_key_value = _acc_clean(self.in_g_key_value.text())
    cfg.generic_timeout_sec = int(self.sp_timeout.value())
    cfg.generic_verify_ssl = self.ck_verify_ssl.isChecked()
    cfg.generic_test_endpoint = _acc_clean(self.in_g_test.text()) or '/health'
    cfg.generic_partners_endpoint = _acc_clean(self.in_g_partners.text()) or '/partners'
    cfg.generic_products_endpoint = _acc_clean(self.in_g_products.text()) or '/products'
    cfg.generic_documents_endpoint = _acc_clean(self.in_g_docs.text()) or '/documents'
    cfg.generic_inventory_documents_endpoint = _acc_clean(self.in_g_inv_docs.text()) or '/inventory-documents'

    cfg.auto_sync_on_pdf_generate = self.ck_auto_pdf.isChecked()
    cfg.auto_sync_on_docx_generate = self.ck_auto_docx.isChecked()
    cfg.auto_sync_on_project_save = self.ck_auto_project.isChecked()
    cfg.sync_partners_with_document = self.ck_sync_partners.isChecked()
    cfg.sync_products_with_inventory = self.ck_sync_products.isChecked()
    cfg.sync_store_doc_with_document = self.ck_sync_store.isChecked()
    cfg.sync_financial_doc_with_document = self.ck_sync_fin.isChecked()

    try:
        self.in_j_import_url.setText(cfg.jumis_import_url)
        self.in_j_export_url.setText(cfg.jumis_export_url)
    except Exception:
        pass
    return cfg


def _acc_open_settings_dialog_v71(self):
    dlg = AccountingIntegrationDialog(self, self._acc_get_cfg())
    try:
        dlg.out_info.setPlainText(
            'Jumis laukā “Speciālā parole” jāievada garā parole no vadiba.mansjumis.lv → '
            'lietotāja izvēlne → Lietotāja iestatījumi → Skatīt garo paroli. '
            'Import/Export URL tiks salaboti automātiski uz pilniem https URL.'
        )
    except Exception:
        pass
    if dlg.exec() == QDialog.Accepted:
        cfg = dlg.get_config()
        self._acc_set_cfg(cfg)
        _acc_log(self, 'accounting_settings_saved', {'system_type': cfg.system_type, 'enabled': cfg.enabled, 'profile_name': cfg.profile_name})
        QMessageBox.information(
            self,
            'Saglabāts',
            'Iestatījumi saglabāti.\n\n'
            'Jumis parolei jābūt speciālajai/garajai parolei no Lietotāja iestatījumi, '
            'nevis parastajai pieslēgšanās parolei.'
        )


def _acc_after_init_v71(self):
    try:
        self._accounting_cfg = _acc_cfg_from_settings_v71()
        _acc_cfg_to_settings(self._accounting_cfg)
    except Exception:
        self._accounting_cfg = AccountingIntegrationConfig()
    _acc_build_menu_v67(self)
    try:
        _acc_start_live_sync_v67(self)
    except Exception:
        pass


_acc_cfg_from_settings = _acc_cfg_from_settings_v71
AccountingIntegrationDialog.get_config = _acc_dialog_get_config_v71
AktaLogs._acc_open_settings_dialog = _acc_open_settings_dialog_v71
_acc_after_init = _acc_after_init_v71


# ==============================
# v74 PATCH: verifikēta Jumis produktu sinhronizācija
# - neuzskata importu par veiksmīgu tikai pēc HTTP 200
# - pēc produkta sūtīšanas pārbauda, vai prece tiešām redzama eksportā
# - produktu nolasīšanai izmanto plašāku XML/JSON parsēšanu
# - attālās ielādes rezultāts uzreiz tiek pielietots lokālajai noliktavai
# ==============================

def _acc_norm_code_v74(value):
    return ''.join(ch for ch in _acc_clean(value).lower() if ch.isalnum())

def _acc_collect_scalar_rows_v74(payload, rows=None):
    if rows is None:
        rows = []
    if isinstance(payload, dict):
        scalar = {}
        complex_values = []
        for k, v in payload.items():
            if isinstance(v, (dict, list)):
                complex_values.append(v)
            else:
                scalar[str(k)] = '' if v is None else str(v)
        if scalar and any(k.lower().startswith('product') or k in {'Price', 'Quantity', 'Balance', 'WarehouseName'} for k in scalar):
            rows.append(scalar)
        for v in complex_values:
            _acc_collect_scalar_rows_v74(v, rows)
    elif isinstance(payload, list):
        for v in payload:
            _acc_collect_scalar_rows_v74(v, rows)
    return rows

def _acc_parse_jumis_products_response_v74(text: str) -> list[dict]:
    products = []
    payload = _acc_json_from_response_text_v67(text or '')
    if payload is not None:
        for row in _acc_collect_scalar_rows_v74(payload, []):
            code = row.get('ProductCode', row.get('Code', ''))
            name = row.get('ProductName', row.get('Name', ''))
            barcode = row.get('ProductBarCode', row.get('BarCode', row.get('Barcode', '')))
            unit = row.get('ProductUnit', row.get('Unit', row.get('UnitCode', 'gab.')))
            price = row.get('Price', row.get('ProductPrice', row.get('SalePrice', '0')))
            qty = row.get('WarehouseNecessaryQuantity', row.get('StockQuantity', row.get('Quantity', row.get('Balance', '0'))))
            notes = row.get('ProductComments', row.get('ProductDescription', row.get('Description', '')))
            group = row.get('ProductClassName', row.get('ProductGroupName', row.get('GroupName', '')))
            if any([code, name, barcode, notes, group, price not in ('', None)]):
                products.append({
                    'ProductCode': code,
                    'ProductName': name,
                    'ProductBarCode': barcode,
                    'ProductUnit': unit,
                    'ProductPrice': price,
                    'ProductPurchasePrice': row.get('ProductPurchasePrice', row.get('PurchasePrice', '')),
                    'ProductVATRate': row.get('ProductVATRate', row.get('VatRate', '')),
                    'ProductGroupName': group,
                    'ProductDescription': notes,
                    'StockQuantity': qty,
                    'Status': row.get('Status', row.get('ProductStatus', '')),
                    'WarehouseName': row.get('WarehouseName', row.get('ProductWarehouse/WarehouseName', '')),
                })
    xml_text = _acc_extract_xml_from_response_text_v67(text)
    root = _acc_parse_xml_root_v67(xml_text)
    if root is not None:
        for p in _acc_findall_any_v67(root, {'Product', 'Row', 'ProductRow', 'Item', 'Record'}):
            code = _acc_child_text_v67(p, ['ProductCode', 'Code'])
            name = _acc_child_text_v67(p, ['ProductName', 'Name'])
            barcode = _acc_child_text_v67(p, ['ProductBarCode', 'BarCode', 'Barcode'])
            unit = _acc_child_text_v67(p, ['ProductUnit', 'Unit', 'UnitCode'], 'gab.')
            price = _acc_nested_child_text_v68(p, ['ProductPrice/Price', 'Price'], '0')
            qty = _acc_nested_child_text_v68(p, ['ProductWarehouse/WarehouseNecessaryQuantity', 'StockQuantity', 'Quantity', 'Balance'], '0')
            notes = _acc_child_text_v67(p, ['ProductComments', 'ProductDescription', 'Description'])
            group = _acc_child_text_v67(p, ['ProductClassName', 'ProductGroupName', 'GroupName'])
            warehouse = _acc_nested_child_text_v68(p, ['ProductWarehouse/WarehouseName', 'WarehouseName'])
            if not any([code, name, barcode, notes, group, price, qty]):
                continue
            products.append({
                'ProductCode': code,
                'ProductName': name,
                'ProductBarCode': barcode,
                'ProductUnit': unit,
                'ProductPrice': price,
                'ProductPurchasePrice': _acc_nested_child_text_v68(p, ['ProductPurchasePrice/PurchasePrice', 'PurchasePrice']),
                'ProductVATRate': _acc_child_text_v67(p, ['ProductVATRate', 'VatRate']),
                'ProductGroupName': group,
                'ProductDescription': notes,
                'StockQuantity': qty,
                'Status': _acc_child_text_v67(p, ['Status', 'ProductStatus']),
                'WarehouseName': warehouse,
            })
    dedup = {}
    for p in products:
        key = _acc_norm_code_v74(p.get('ProductCode')) or _acc_norm_code_v74(p.get('ProductBarCode')) or _acc_norm_code_v74(p.get('ProductName'))
        if key:
            dedup[key] = p
    return list(dedup.values())

def _jumis_fetch_products_v74(self, limit: int = 5000):
    field_variants = [
        '<Field Name="ProductCode"/><Field Name="ProductName"/><Field Name="ProductBarCode"/><Field Name="ProductUnit"/><Field Name="ProductComments"/><Field Name="ProductClassName"/><Field Name="ProductPrice"/><Field Name="ProductWarehouse"/>',
        '<Field Name="ProductCode"/><Field Name="ProductName"/><Field Name="ProductBarCode"/><Field Name="ProductUnit"/><Field Name="Price"/><Field Name="ProductComments"/><Field Name="ProductClassName"/><Field Name="WarehouseName"/><Field Name="WarehouseNecessaryQuantity"/>',
        '<Field Name="ProductCode"/><Field Name="ProductName"/><Field Name="ProductBarCode"/><Field Name="ProductUnit"/><Field Name="ProductPrice"/><Field Name="Quantity"/><Field Name="Status"/>'
    ]
    candidates = []
    for structure in ('Tree', 'Sheet'):
        for fields in field_variants:
            candidates.append(f'<?xml version="1.0" ?><dataroot><tjDocument Version="TJ5.5.101"/><tjRequest Name="Product" Operation="Read" Version="TJ5.5.101" Structure="{structure}"><tjData Read="All"/><tjFields>{fields}</tjFields></tjRequest></dataroot>')
            candidates.append(f'<?xml version="1.0" ?><dataroot><tjDocument Version="TJ5.5.101"/><tjRequest Name="Product" Operation="Read" Version="TJ5.5.101" Structure="{structure}"><tjFields>{fields}</tjFields></tjRequest></dataroot>')
    best = None
    best_products = []
    for i, xmlrequest in enumerate(candidates, start=1):
        res = self._post_export(xmlrequest, f'fetch_products_try_{i}')
        products = _acc_parse_jumis_products_response_v74(getattr(res, 'response_text', '') or '')
        extra = dict(getattr(res, 'extra', {}) or {})
        extra['products'] = products[:max(int(limit or 5000), 1)]
        extra['candidate_index'] = i
        res.extra = extra
        if best is None or len(products) > len(best_products) or (getattr(res, 'ok', False) and not getattr(best, 'ok', False)):
            best = res
            best_products = products
        if getattr(res, 'ok', False) and products:
            break
    final_res = best or AccountingSyncResult(ok=False, operation='fetch_products', system=self.system_name, request_payload='', response_status=0, response_text='Neizdevās saņemt preču atbildi no Jumis.')
    final_res.operation = 'fetch_products'
    extra = dict(getattr(final_res, 'extra', {}) or {})
    extra['products'] = best_products[:max(int(limit or 5000), 1)]
    extra['limit'] = limit
    if getattr(final_res, 'ok', False) and not best_products:
        extra['message'] = (extra.get('message') or 'Jumis atbildēja, bet eksportā netika atrasta neviena prece.')
    final_res.extra = extra
    return final_res

def _jumis_find_product_in_list_v74(products, item):
    target_codes = {
        _acc_norm_code_v74(getattr(item, 'sku', '')),
        _acc_norm_code_v74(getattr(item, 'inventory_id', '')),
        _acc_norm_code_v74(getattr(item, 'svitrkods', '')),
        _acc_norm_code_v74(getattr(item, 'nosaukums', '')),
    }
    target_codes.discard('')
    for p in products or []:
        cand = {
            _acc_norm_code_v74(p.get('ProductCode')),
            _acc_norm_code_v74(p.get('ProductBarCode')),
            _acc_norm_code_v74(p.get('ProductName')),
        }
        if target_codes & cand:
            return p
    return None

def _jumis_verify_product_visible_v74(self, item):
    fetch_res = _jumis_fetch_products_v74(self, limit=max(int(getattr(self.config, 'live_sync_fetch_limit', 5000) or 5000), 500))
    products = list((getattr(fetch_res, 'extra', {}) or {}).get('products') or [])
    return fetch_res, _jumis_find_product_in_list_v74(products, item)

def _jumis_product_xml_candidates_v74(self, item):
    it = item or NoliktavasPrece()
    code = _acc_clean(getattr(it, 'sku', '')) or _acc_clean(getattr(it, 'inventory_id', '')) or _acc_clean(getattr(it, 'svitrkods', '')) or f'ITEM-{secrets.token_hex(4)}'
    name = _acc_clean(getattr(it, 'nosaukums', '')) or code
    unit = _acc_clean(getattr(it, 'vieniba', '')) or 'gab.'
    price = _acc_to_amount(getattr(it, 'cena', '0'))
    qty = _acc_to_qty(getattr(it, 'atlikums', 0))
    notes = _acc_clean(getattr(it, 'piezimes', ''))
    group = _acc_clean(getattr(it, 'kategorija', ''))
    barcode = _acc_clean(getattr(it, 'svitrkods', ''))
    vat_rate = _acc_to_amount(getattr(it, 'pvn_likme', '21'), default='21.00')
    wh = _acc_clean(getattr(it, 'noliktavas_nosaukums', '')) or 'Pamatnoliktava'
    common_open = '<?xml version="1.0" encoding="utf-8" ?><dataroot><tjDocument Version="TJ5.5.101"/><tjResponse Name="Product" Operation="Insert" Version="TJ5.5.101" Structure="Tree"><Product>'
    common_close = '</Product></tjResponse></dataroot>'
    cand1 = common_open + ''.join([
        _acc_xml('ProductCode', code),
        _acc_xml('ProductName', name),
        _acc_xml('ProductBarCode', barcode),
        _acc_xml('ProductUnit', unit),
        _acc_xml('ProductComments', notes),
        _acc_xml('ProductClassName', group),
        _acc_xml('ProductVATRate', vat_rate),
        '<ProductPrice>' + _acc_xml('Price', price) + '</ProductPrice>',
        '<ProductWarehouse>' + _acc_xml('WarehouseName', wh) + _acc_xml('WarehouseNecessaryQuantity', qty) + '</ProductWarehouse>'
    ]) + common_close
    cand2 = common_open + ''.join([
        _acc_xml('ProductCode', code),
        _acc_xml('ProductName', name),
        _acc_xml('ProductBarCode', barcode),
        _acc_xml('ProductUnit', unit),
        _acc_xml('ProductComments', notes),
        _acc_xml('ProductClassName', group),
        _acc_xml('ProductVATRate', vat_rate),
        _acc_xml('ProductPrice', price),
    ]) + common_close
    cand3 = common_open + ''.join([
        _acc_xml('ProductCode', code),
        _acc_xml('ProductName', name),
        _acc_xml('ProductBarCode', barcode),
        _acc_xml('ProductUnit', unit),
        _acc_xml('ProductNick', name[:20]),
        _acc_xml('ProductComments', notes),
        _acc_xml('ProductClassName', group),
        '<ProductPrice>' + _acc_xml('Price', price) + '</ProductPrice>'
    ]) + common_close
    return code, [cand1, cand2, cand3]

def _jumis_upsert_product_v74(self, item):
    code, xml_candidates = _jumis_product_xml_candidates_v74(self, item)
    last_res = None
    for idx, xmlrequest in enumerate(xml_candidates, start=1):
        res = self._post_import(xmlrequest, f'upsert_product_try_{idx}')
        extra = dict(getattr(res, 'extra', {}) or {})
        extra['product_code'] = code
        extra['xml_variant'] = idx
        res.extra = extra
        last_res = res
        if not getattr(res, 'ok', False):
            continue
        verify_res, found = _jumis_verify_product_visible_v74(self, item)
        extra = dict(getattr(res, 'extra', {}) or {})
        extra['verification_http_status'] = getattr(verify_res, 'response_status', 0)
        extra['verification_found'] = bool(found)
        if found:
            extra['message'] = f'Prece verificēta Jumis eksportā ar variantu #{idx}.'
            res.extra = extra
            res.ok = True
            res.operation = 'upsert_product'
            return res
        extra['message'] = (extra.get('message') or '') + ' Imports tika nosūtīts, bet prece pēc pārbaudes Jumis eksportā netika atrasta.'
        res.extra = extra
        res.ok = False
        last_res = res
    if last_res is None:
        last_res = AccountingSyncResult(ok=False, operation='upsert_product', system=self.system_name, request_payload='', response_status=0, response_text='Neizdevās nosūtīt produktu uz Jumis.')
    else:
        last_res.operation = 'upsert_product'
    return last_res

def _acc_fetch_remote_products_v74(self, quiet: bool = False):
    try:
        conn = self._acc_connector()
        limit = max(int(_acc_cfg_get_v67(self._acc_get_cfg(), 'live_sync_fetch_limit', 5000) or 5000), 1)
        res = conn.fetch_products(limit=limit)
        products = list((getattr(res, 'extra', {}) or {}).get('products') or [])
        _acc_log(self, 'accounting_fetch_products_v74', {'ok': getattr(res, 'ok', False), 'status': getattr(res, 'response_status', 0), 'count': len(products), 'quiet': quiet})
        if not quiet:
            _acc_show_result_v67(self, res, 'Attālināto preču nolasīšana')
        return res
    except Exception as e:
        _acc_log(self, 'accounting_fetch_products_error_v74', {'error': str(e), 'quiet': quiet})
        if not quiet:
            QMessageBox.warning(self, 'Kļūda', f'Preču nolasīšana neizdevās.\n\n{e}')
        return AccountingSyncResult(ok=False, operation='fetch_products', system='N/A', request_payload='', response_status=0, response_text=str(e))

def _acc_pull_remote_inventory_now_v74(self, silent: bool = False):
    res = _acc_fetch_remote_products_v74(self, quiet=True)
    if not getattr(res, 'ok', False):
        if not silent:
            _acc_show_result_v67(self, res, 'Attālinātās noliktavas ielāde')
        return res
    products = list((getattr(res, 'extra', {}) or {}).get('products') or [])
    stats = _acc_apply_remote_products_to_local_v67(self, products, remove_missing=bool(_acc_cfg_get_v67(self._acc_get_cfg(), 'live_sync_remove_missing_local_products', False)))
    extra = dict(getattr(res, 'extra', {}) or {})
    extra['apply_stats'] = stats
    if not products:
        extra['message'] = (extra.get('message') or 'Jumis atbildēja, bet eksportā netika atrasta neviena prece, ko ielasīt lokāli.')
    res.extra = extra
    try:
        if hasattr(self, '_refresh_noliktava_table'):
            self._refresh_noliktava_table()
    except Exception:
        pass
    try:
        if hasattr(self, '_update_preview'):
            self._update_preview()
    except Exception:
        pass
    if not silent:
        _acc_show_result_v67(self, res, 'Attālinātās noliktavas ielāde')
    return res

JumisConnector.fetch_products = _jumis_fetch_products_v74
JumisConnector.upsert_product = _jumis_upsert_product_v74
AktaLogs._acc_fetch_remote_products = _acc_fetch_remote_products_v74
AktaLogs._acc_pull_remote_inventory_now = _acc_pull_remote_inventory_now_v74

if __name__ == "__main__":
    set_windows_app_id("kulinics.akta_generators")
    app = QApplication(sys.argv)
    try:
        app.setWindowIcon(QIcon(resource_path("Akta_Generators_Icon.ico")))
    except Exception:
        pass
    apply_modern_theme(app, dark=True)
    window = AktaLogs()
    window.show()
    sys.exit(app.exec())
