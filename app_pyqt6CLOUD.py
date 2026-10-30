from PyQt6.QtCore import Qt, QThread, QObject, pyqtSignal, QUrl, QTimer, QPropertyAnimation, QDate, QEvent, QRectF
import sys
import io
import os
import re
import unicodedata
import socket
import sqlite3
import time
import datetime
import shutil
import json
import urllib.request
import barcode
import openpyxl
import pandas as pd
import qrcode
import requests
import threading
import functools
import hashlib
from barcode.writer import ImageWriter
from PIL import Image, ImageDraw, ImageFont, ImageWin


from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, QLabel,
    QLineEdit, QPushButton, QRadioButton, QButtonGroup, QTabWidget,
    QFormLayout, QGroupBox, QSpinBox, QDialog, QScrollArea, QFrame,
    QFileDialog, QCheckBox, QMessageBox, QTableWidget, QTableWidgetItem, QHeaderView,
    QGraphicsOpacityEffect, QTextEdit, QDateEdit, QListWidget, QInputDialog, QAbstractItemView,
    QStyledItemDelegate, QAbstractItemDelegate, QComboBox, QPlainTextEdit, QAbstractSpinBox, QStyle,
    QToolButton, QMenu, QSplitter, QGridLayout
)
from PyQt6.QtGui import QColor, QPixmap, QFont, QDesktopServices, QCloseEvent, QShortcut, QKeySequence, QIntValidator, QPainter, QImageReader, QImage
from PyQt6.QtWebEngineWidgets import QWebEngineView
from PyQt6.QtWebEngineCore import QWebEnginePage

# -------------------------------------------------------------------
# JS KONZOL ÜZENETEK ELNÉMÍTÁSA
# -------------------------------------------------------------------
class SilentWebEnginePage(QWebEnginePage):
    def javaScriptConsoleMessage(self, level, message, lineNumber, sourceID):
        pass

# -------------------------------------------------------------------
# GLOBÁLIS JELZŐBUSZ: háttérszálról (flush_sync_outbox) biztonságosan
# eljuttatja az R18-ütközéseket a főablak UI szálára (QMessageBox csak
# a fő szálon hívható, de egy pyqtSignal emit bármelyik szálról jöhet).
# -------------------------------------------------------------------
class OutboxSignalBus(QObject):
    conflicts_detected = pyqtSignal(list)

outbox_signal_bus = OutboxSignalBus()

# -------------------------------------------------------------------
# 📷 TERMÉKFOTÓK a NAS-ról (Synology Drive szinkronizált mappa)
# A fotó akkor tartozik egy termékhez, ha a FÁJLNEVÉBEN vagy valamelyik MAPPANEVÉBEN szerepel az
# R18-szám (R18-12960, R18_12960, R18 12960, r18/12960 ...). A mappák a beallitasok.json-ban
# ("foto_mappak") vannak, a kártyán a "📁 Fotómappa" gombbal állíthatók.
# -------------------------------------------------------------------
# Fotók be/ki - ha egy gépen nem kellenek (pl. nagyon régi gép), írd át False-ra
PHOTOS_ENABLED = True
# Egyszerre legfeljebb ennyi kép töltődik a háttérben (régebbi gépen se terhelje a processzort)
PHOTO_MAX_PARALLEL_LOADS = 2
_photo_load_slots = threading.Semaphore(PHOTO_MAX_PARALLEL_LOADS)
PHOTO_EXTENSIONS = (".jpg", ".jpeg", ".png", ".webp", ".bmp", ".gif")
# Az előnézet mérete a kártyán (pixel)
PHOTO_THUMB_W, PHOTO_THUMB_H = 80, 70
PHOTO_INDEX_REFRESH_SEC = 600          # ennyi időnként újraolvassa a mappákat (új fotók)
_PHOTO_CODE_RE = re.compile(r"R18\s*[-_/ ]?\s*0*(\d{1,6})(?!\d)", re.IGNORECASE)

class PhotoBus(QObject):
    thumb_ready = pyqtSignal(int, object)      # (kérés-azonosító, QImage)
    index_ready = pyqtSignal(int)              # fotók száma

photo_bus = PhotoBus()

class PhotoIndex:
    """R18-kód -> fotófájlok listája (háttérszálon építve, a kártyák csak kiolvassák)."""
    def __init__(self):
        self.by_code = {}
        self.built_at = 0.0
        self.building = False
        self.lock = threading.Lock()

    @staticmethod
    def default_dirs():
        dirs = []
        home = os.path.expanduser("~")
        for name in ("SynologyDrive", "Synology Drive"):
            p = os.path.join(home, name)
            if os.path.isdir(p):
                dirs.append(p)
        return dirs

    def dirs(self):
        if not PHOTOS_ENABLED:
            return []
        saved = load_config_value("foto_mappak")
        if isinstance(saved, list) and saved:
            return [d for d in saved if d]
        return self.default_dirs()

    def build(self):
        if self.building:
            return
        self.building = True
        try:
            idx = {}
            seen = 0
            for root_dir in self.dirs():
                if not os.path.isdir(root_dir):
                    continue
                for dirpath, dirnames, filenames in os.walk(root_dir):
                    dirnames[:] = [d for d in dirnames if not d.startswith(("#", "@", "."))]   # Synology rendszermappák
                    folder_codes = {f"R18-{int(m)}" for m in _PHOTO_CODE_RE.findall(os.path.relpath(dirpath, root_dir))}
                    seen += len(filenames)
                    if seen > 400:
                        seen = 0
                        time.sleep(0.005)   # "szusszanás": a felület ne akadjon meg nagy mappánál
                    for fn in filenames:
                        if not fn.lower().endswith(PHOTO_EXTENSIONS):
                            continue
                        codes = {f"R18-{int(m)}" for m in _PHOTO_CODE_RE.findall(fn)} or folder_codes
                        for c in codes:
                            idx.setdefault(c, []).append(os.path.join(dirpath, fn))
            for c in idx:
                idx[c].sort()
            with self.lock:
                self.by_code = idx
                self.built_at = time.monotonic()
            photo_bus.index_ready.emit(sum(len(v) for v in idx.values()))
        except Exception as e:
            print(f"[Fotó] Index hiba: {e}")
        finally:
            self.building = False

    def build_async(self, force=False):
        if not PHOTOS_ENABLED:
            return
        if force or time.monotonic() - self.built_at > PHOTO_INDEX_REFRESH_SEC:
            threading.Thread(target=self.build, daemon=True).start()

    def photos(self, code):
        m = _PHOTO_CODE_RE.search(str(code or ""))
        key = f"R18-{int(m.group(1))}" if m else norm_code(code)
        with self.lock:
            return list(self.by_code.get(key, []))

photo_index = PhotoIndex()
_thumb_cache = {}
_thumb_req = [0]

# Windows "csak a felhőben" fájl (Synology Drive / OneDrive igény szerinti szinkron) jelzői
_CLOUD_ONLY_ATTRS = 0x00400000 | 0x00040000 | 0x00001000   # RECALL_ON_DATA_ACCESS | RECALL_ON_OPEN | OFFLINE

def is_cloud_only(path):
    """Igaz, ha a fájl csak a felhőben van (a gépen nincs letöltve). A vizsgálat NEM indít letöltést."""
    try:
        attrs = getattr(os.stat(path), "st_file_attributes", 0)
        return bool(attrs & _CLOUD_ONLY_ATTRS)
    except OSError:
        return True

def load_image_scaled(path, max_w, max_h):
    """Kép betöltése MÁR KICSINYÍTVE (gyors, kevés memória) - háttérszálon is hívható (QImage).
    A fájlt a Python olvassa be (közben a többi szál - a felület - fut tovább, akkor is, ha a
    Synology Drive épp letölti vagy engedélyre vár), a dekódolás már a memóriából megy."""
    from PyQt6.QtCore import QBuffer, QByteArray, QIODevice
    with open(path, "rb") as f:
        data = f.read()
    buf = QBuffer()
    buf.setData(QByteArray(data))
    buf.open(QIODevice.OpenModeFlag.ReadOnly)
    reader = QImageReader(buf)
    reader.setAutoTransform(True)                       # telefonos fotók elforgatása (EXIF)
    size = reader.size()
    if size.isValid() and (size.width() > max_w or size.height() > max_h):
        size.scale(max_w, max_h, Qt.AspectRatioMode.KeepAspectRatio)
        reader.setScaledSize(size)
    img = reader.read()
    buf.close()
    return img

def request_thumbnail(path, max_w=120, max_h=120, allow_download=False):
    """Előnézet kérése háttérszálon; a kész QImage a photo_bus.thumb_ready jelzésen jön. Visszaad: kérés-azonosító.
    allow_download=False: ha a fájl csak a felhőben van, NEM tölti le, hanem "CLOUD" jelzést küld."""
    _thumb_req[0] += 1
    rid = _thumb_req[0]
    key = (path, max_w, max_h)
    if key in _thumb_cache:
        QTimer.singleShot(0, lambda: photo_bus.thumb_ready.emit(rid, _thumb_cache[key]))
        return rid

    def job():
        if not allow_download and is_cloud_only(path):
            photo_bus.thumb_ready.emit(rid, "CLOUD")
            return
        with _photo_load_slots:              # egyszerre legfeljebb PHOTO_MAX_PARALLEL_LOADS kép
            try:
                img = load_image_scaled(path, max_w, max_h)
            except Exception:
                img = QImage()
        # Csak a kis előnézeteket tartjuk meg (a nagy nézet képei sok memóriát foglalnának)
        if not img.isNull() and max_w <= 200 and max_h <= 200:
            if len(_thumb_cache) > 500:
                _thumb_cache.clear()
            _thumb_cache[key] = img
        photo_bus.thumb_ready.emit(rid, img)
    threading.Thread(target=job, daemon=True).start()
    return rid

# -------------------------------------------------------------------
# ALAPÉRTÉKEK ÉS BEÁLLÍTÁSOK
# -------------------------------------------------------------------
VPS_IP = "217.144.56.10"
VPS_URL = f"http://{VPS_IP}:5000"
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DB_FILE = os.path.join(BASE_DIR, "raktar_adatbazis.db")
CONFIG_FILE = os.path.join(BASE_DIR, "beallitasok.json")
# ⚙️ Folyamat (technikai) napló napi fájljai - ennyi napig őrzi meg őket
PROCESS_LOG_DIR = os.path.join(BASE_DIR, "naplok")
PROCESS_LOG_KEEP_DAYS = 30

TEST_SHEETS_ID = "1ZoB2yl_GvaY0kiICQN10dEKghAE7b9Z3hDlwtv9nMdc"

# -------------------------------------------------------------------
# ELADÁSI MUNKALAPOK (3 db) - PDA / Eladás fül 3 gombja ide írja az eladásokat.
# TESZT UTÁN ITT KELL ÁTÍRNI az élő táblázatra: "spreadsheet_id" = a Google táblázat ID-ja
# (a docs.google.com/spreadsheets/d/<EZ>/edit részből), "worksheet" = a munkalap (fül) neve.
# Ha a munkalap még nem létezik, a program létrehozza és fejlécet ír bele.
# (A credentials.json szolgáltatásfiókjának szerkesztési joga kell legyen a táblázathoz!)
# -------------------------------------------------------------------
SALES_SHEETS = {
    "EBAY_KIVULI": {"cim": "eBay-en kívüli eladás", "spreadsheet_id": TEST_SHEETS_ID, "worksheet": "Eladas_EbayKivul"},
    "EBAY":        {"cim": "eBay eladás",           "spreadsheet_id": TEST_SHEETS_ID, "worksheet": "Eladas_Ebay"},
    "WEBSHOP":     {"cim": "Webshop eladás",        "spreadsheet_id": TEST_SHEETS_ID, "worksheet": "Eladas_Webshop"},
}
# A Zebra PDA-n rögzített eladásokat a program ilyen gyakran nézi meg a VPS-en (és veszi át)
PDA_SALES_POLL_MS = 20000
# Hibakódok, amiket az asztali program a PDA-nak visszaküld (a PDA Napló / Folyamat napló fülén látszanak):
#   A-21  a termék nincs meg ezen a gépen        A-22  helyi adatbázis / mentési hiba
#   A-23  ismeretlen platform / mező             A-24  váratlan hiba az átvétel közben
#   A-25  a készletet nem lehetett javítani (a termék már törölve) - Admin fülön visszaállítható
# Az eladási munkalapok oszlopai (sorrendben) - (belső kulcs, fejléc)
SALES_COLUMNS = [
    ("eladas_id", "Eladás ID"), ("datum", "Eladás dátuma"), ("platform", "Platform"),
    ("A", "R18 kód"), ("C", "Márka"), ("D", "Típus"), ("E", "Cikkszám"), ("G", "Állapot"),
    ("L", "Raklap"), ("M", "Súly"), ("N", "eBay ID"),
    ("vevo", "Vevő"), ("ar", "Eladási ár"), ("darab", "Darab"), ("tracking", "Tracking number"),
    ("megjegyzes", "Megjegyzés"),
]
# Ezek a mezők "később" is kitölthetők - az Eladottak fülön narancs háttérrel jelöli, ha még üresek
SALES_LATER_FIELDS = ("vevo", "ar", "tracking")

MAP_SHEETS_ID = "13uw_-HNPHep94j-Ra9TsJrUYXADa8qcuBYyBuFdL1Fw"
MAP_EDIT_URL = f"https://docs.google.com/spreadsheets/d/{MAP_SHEETS_ID}/edit"

CREDENTIALS_FILE = "credentials.json"

# Google Sheets KIKAPCSOLVA: a program a helyi SQLite adatbázissal + VPS-sel dolgozik.
# (Ha valaha vissza kell kapcsolni: SHEETS_ENABLED = True)
SHEETS_ENABLED = True

# Excel Live munkalap automatikus mentési késleltetése (ms) - ennyi idő után indul a mentés/nyomtatás
AUTO_SAVE_INTERVAL_MS = 4000

# -------------------------------------------------------------------
# TÖRÖLT SOROK FIGYELÉSE
# A program a FIGYELT (élő) táblázatot CSAK OLVASSA (a Google felé is csak olvasási joggal
# kapcsolódik, így véletlenül sem módosíthatja). Ha onnan eltűnik egy R18-as termék sora,
# vagy kiürítik a Márka/Típus/Cikkszám mezőit, a korábbi teljes adatai a ToroltSor munkalapra kerülnek.
# -------------------------------------------------------------------
# FIGYELEM: a figyelés a VPS-en fut folyamatosan (/opt/raktarapp/torolt_figyelo.py, lásd vps/ mappa).
# Ezért ITT hagyd üresen - különben a gépen futó program is figyelne, feleslegesen.
# (Csak akkor töltsd ki, ha a VPS-es figyelő nem fut.)
# A figyelt (élő) táblázat ID-ja: a docs.google.com/spreadsheets/d/<EZ>/edit részből.
# Üresen hagyva a figyelés KI van kapcsolva.
DELETED_WATCH_SPREADSHEET_ID = ""
# A figyelt munkalap (fül) neve - üresen hagyva a táblázat ELSŐ munkalapja
DELETED_WATCH_WORKSHEET = ""
# Milyen gyakran nézze meg a figyelt táblázatot (másodperc)
DELETED_WATCH_INTERVAL_SEC = 60

# A ToroltSor munkalap helye (ide ÍR a program) - NEM lehet ugyanaz a táblázat, mint a figyelt!
DELETED_SHEET_SPREADSHEET_ID = TEST_SHEETS_ID
DELETED_SHEET_NAME = "ToroltSor"
DELETED_SHEET_HEADER = ["A: F-Kód", "B: Vonalkód", "C: Márka", "D: Típus", "E: Cikkszám", "F: DB",
                        "G: Állapot", "H: Ár", "I: Páter", "J: Szabad", "K: Polc", "L: Raklap",
                        "M: Súly", "N: eBay ID", "O: Megjegyzés", "P: Állapot részlet",
                        "Q: Törlés ideje", "R: Honnan törölték"]
# Ha egy ellenőrzéskor ennél több kitöltött termék tűnik el egyszerre a figyelt táblázatból, az valószínűleg
# nem kézi törlés (hanem pl. hibás letöltés / szűrés) - ilyenkor nem archivál automatikusan.
DELETED_MISSING_SAFETY_LIMIT = 20
# A programon BELÜLI törlés / eladás / visszavonás / Excel Live kiürítés is a ToroltSor lapra írja a
# termék adatait (ez a program saját táblázatára vonatkozik, nem a figyeltre). False = kikapcsolva.
ARCHIVE_PROGRAM_DELETIONS = True

def get_deleted_sheet():
    """A 'ToroltSor' munkalap (ha nincs, létrehozza; üres lapra fejlécet ír)."""
    if not SHEETS_ENABLED:
        return None
    import gspread
    from google.oauth2.service_account import Credentials
    scopes = ["https://www.googleapis.com/auth/spreadsheets"]
    creds = Credentials.from_service_account_file(CREDENTIALS_FILE, scopes=scopes)
    if DELETED_WATCH_SPREADSHEET_ID and DELETED_SHEET_SPREADSHEET_ID.strip() == DELETED_WATCH_SPREADSHEET_ID.strip():
        # Biztonsági zár: a figyelt (élő) táblázatba SOHA nem írunk
        raise RuntimeError("A ToroltSor nem lehet a figyelt (élő) táblázatban - állíts be másik DELETED_SHEET_SPREADSHEET_ID-t!")
    client = gspread.authorize(creds)
    book = client.open_by_key(DELETED_SHEET_SPREADSHEET_ID)
    try:
        ws = book.worksheet(DELETED_SHEET_NAME)
    except gspread.exceptions.WorksheetNotFound:
        ws = book.add_worksheet(title=DELETED_SHEET_NAME, rows=1000, cols=len(DELETED_SHEET_HEADER))
    if not ws.row_values(1):
        ws.update(range_name="A1:R1", values=[DELETED_SHEET_HEADER])
    return ws

def _read_watched_sheet_values():
    """A figyelt (élő) táblázat összes cellája - KIZÁRÓLAG OLVASÁSI JOGGAL (spreadsheets.readonly),
    így a program semmilyen módon nem tudja módosítani."""
    import gspread
    from google.oauth2.service_account import Credentials
    scopes = ["https://www.googleapis.com/auth/spreadsheets.readonly"]
    creds = Credentials.from_service_account_file(CREDENTIALS_FILE, scopes=scopes)
    book = gspread.authorize(creds).open_by_key(DELETED_WATCH_SPREADSHEET_ID.strip())
    ws = book.worksheet(DELETED_WATCH_WORKSHEET) if DELETED_WATCH_WORKSHEET.strip() else book.sheet1
    return ws.get_all_values()

_watch_lock = threading.Lock()

def check_watched_sheet_for_deletions(log=print):
    """Összeveti a figyelt táblázat mostani állapotát az előző ellenőrzéskor eltárolt pillanatképpel
    (helyi 'figyelt_snapshot' tábla - így a program bezárása alatt történt törléseket is észreveszi).
    Ami eltűnt / kiürült, annak a KORÁBBI teljes adata a ToroltSor munkalapra kerül.
    Visszaadja az új archív bejegyzések számát."""
    if not DELETED_WATCH_SPREADSHEET_ID.strip() or not SHEETS_ENABLED:
        return 0
    if not _watch_lock.acquire(blocking=False):
        return 0
    try:
        try:
            values = _read_watched_sheet_values()
        except Exception as e:
            log(f"⚠️ Figyelt táblázat nem olvasható (később újra próbálja): {e}", "ERROR")
            return 0

        current = {}
        for r in values:
            r = (list(r) + [""] * 16)[:16]
            code = str(r[0]).strip().upper()
            if code.startswith("R18-"):
                current[code] = {c: str(v).strip() for c, v in zip(ROW_COLS, r)}

        conn = get_db_connection()
        cursor = conn.cursor()
        watch_key = f"{DELETED_WATCH_SPREADSHEET_ID.strip()}|{DELETED_WATCH_WORKSHEET.strip()}"
        cursor.execute("SELECT code, payload FROM figyelt_snapshot WHERE forras = ?", (watch_key,))
        previous = {row["code"]: json.loads(row["payload"]) for row in cursor.fetchall()}

        # Hibás / üres letöltés elleni védelem: ilyenkor se nem archivál, se nem írja felül a pillanatképet
        if previous and len(current) < len(previous) * 0.5:
            log(f"⚠️ Figyelt táblázat: csak {len(current)} sor jött le (előzőleg {len(previous)}) - "
                f"valószínűleg hibás letöltés, kihagyva.", "ERROR")
            conn.close()
            return 0

        gone = []
        for code, old_row in previous.items():
            if not _row_has_product_data(old_row):
                continue
            new_row = current.get(code)
            if new_row is None:
                gone.append((old_row, "Élő tábla: sor törölve"))
            elif not _row_has_product_data(new_row):
                gone.append((old_row, "Élő tábla: kiürítve"))

        archived = 0
        if len(gone) > DELETED_MISSING_SAFETY_LIMIT:
            log(f"⚠️ Figyelt táblázat: {len(gone)} termék tűnt el egyszerre - túl sok, automatikus "
                f"archiválás kihagyva (ellenőrizd a táblázatot!)", "ERROR")
        else:
            for old_row, forras in gone:
                if archive_deleted_row(old_row, forras, cursor, force=True):
                    archived += 1

        # Új pillanatkép mentése
        cursor.execute("DELETE FROM figyelt_snapshot WHERE forras = ?", (watch_key,))
        cursor.executemany("INSERT INTO figyelt_snapshot (forras, code, payload) VALUES (?, ?, ?)",
                           [(watch_key, c, json.dumps(r, ensure_ascii=False)) for c, r in current.items()])
        conn.commit()
        conn.close()

        if archived:
            log(f"🗃️ Figyelt táblázat: {archived} törölt termék teljes adatai a '{DELETED_SHEET_NAME}' munkalapra kerülnek", "INFO")
            flush_outbox_async()
        return archived
    except Exception as e:
        log(f"❌ Törölt sorok figyelési hiba: {e}", "ERROR")
        return 0
    finally:
        _watch_lock.release()

def get_sheet():
    """Google Sheets munkalap (csak ha SHEETS_ENABLED = True)"""
    if not SHEETS_ENABLED:
        return None
    import gspread
    from google.oauth2.service_account import Credentials
    scopes = ["https://www.googleapis.com/auth/spreadsheets"]
    creds = Credentials.from_service_account_file(CREDENTIALS_FILE, scopes=scopes)
    client = gspread.authorize(creds)
    return client.open_by_key(TEST_SHEETS_ID).sheet1
COLOR_CORNFLOWER = {"red": 0.26, "green": 0.52, "blue": 0.96}
COLOR_ORANGE = {"red": 0.93, "green": 0.49, "blue": 0.19}
COLOR_RED = {"red": 1.0, "green": 0.0, "blue": 0.0}
COLOR_WHITE = {"red": 1.0, "green": 1.0, "blue": 1.0}

try:
    import win32print
    import win32ui
    WIN32_AVAILABLE = True
except ImportError:
    WIN32_AVAILABLE = False

# -------------------------------------------------------------------
# CÍMKENYOMTATÓ: helyi vagy távoli (a nyomtató nélküli gép a nyomtatós gépre küldi a címkét a VPS-en át)
# -------------------------------------------------------------------
class PrintBus(QObject):
    msg = pyqtSignal(str)     # értesítés a felületre (toast) bármelyik szálból


print_bus = PrintBus()
_label_printer_cache = {"t": 0.0, "nev": None}
_print_queue_lock = threading.Lock()


def printer_mode():
    """'auto' (ha van Zebra nyomtató ezen a gépen, itt nyomtat, különben a nyomtatós gépre küldi),
    'helyi' (ehhez a géphez van kötve), 'tavoli' (nincs nyomtató – mindig a nyomtatós gépre küldi)."""
    m = str(load_config_value("nyomtato_mod", "auto") or "auto")
    return m if m in ("auto", "helyi", "tavoli") else "auto"


def local_label_printer(force=False):
    """A gépre kötött címkenyomtató neve (Zebra / ZDesigner / GK420 ...), vagy None. 60 mp-ig megjegyzi."""
    if not WIN32_AVAILABLE:
        return None
    if not force and time.time() - _label_printer_cache["t"] < 60:
        return _label_printer_cache["nev"]
    nev = None
    try:
        for pr in win32print.EnumPrinters(win32print.PRINTER_ENUM_LOCAL | win32print.PRINTER_ENUM_CONNECTIONS):
            n = str(pr[2])
            if any(k in n.upper() for k in ("ZEBRA", "ZDESIGNER", "GK420", "GX420", "ZD4", "ZD2", "ZT2", "ZT4")):
                nev = n
                break
    except Exception:
        nev = None
    _label_printer_cache.update(t=time.time(), nev=nev)
    return nev


def is_print_host():
    """Ez a gép nyomtat-e címkét (és nyomtatja-e ki a többi gépről érkezőket)."""
    m = printer_mode()
    if m == "tavoli":
        return False
    if m == "helyi":
        return WIN32_AVAILABLE
    return local_label_printer() is not None


def _send_zpl_remote(zpl_data):
    """A címkét a VPS-en át a nyomtatós gépre küldi; a végeredményről (kinyomtatva / hiba) később értesítést küld."""
    m = re.search(r"R18[-/][0-9A-Za-z]+", zpl_data or "")
    code = m.group(0).upper() if m else ""
    pq = re.search(r"\^PQ(\d+)", zpl_data or "")
    try:
        res = vps_request("POST", f"{VPS_URL}/print/jobs", timeout=8,
                          json={"zpl": zpl_data, "A": code, "peldany": int(pq.group(1)) if pq else 1, "kuldo": GEP_NEV})
        if res.status_code == 404:
            return False, ("Ehhez a géphez nincs címkenyomtató kötve, és a VPS-en még a régi server_api.py fut "
                           "(nem tud a nyomtatós gépre küldeni)")
        if res.status_code != 200:
            return False, f"A címkét nem sikerült a nyomtatós gépre küldeni (HTTP {res.status_code})"
        js = res.json()
    except Exception as e:
        return False, f"A címkét nem sikerült a nyomtatós gépre küldeni (nincs kapcsolat a szerverrel: {e.__class__.__name__})"
    jid, gepek = js.get("id"), js.get("nyomtato_gepek") or []

    def figyel():
        for _ in range(45):
            time.sleep(2)
            try:
                r = vps_request("GET", f"{VPS_URL}/print/jobs/{jid}", timeout=8).json()
            except Exception:
                continue
            if r.get("allapot") == "kesz":
                print_bus.msg.emit(f"✅ Kinyomtatva a(z) {r.get('nyomtatta') or 'nyomtatós'} gépen: {code}")
                return
            if r.get("allapot") in ("hiba", "lejart"):
                print_bus.msg.emit(f"❌ A nyomtatós gépen nem sikerült ({code}): {r.get('uzenet', '')}")
                return
        print_bus.msg.emit(f"⏳ {code}: a címke még vár – a nyomtatós gépen fut a program? (2 órán belül még kinyomtatja)")
    threading.Thread(target=figyel, daemon=True).start()
    if gepek:
        return True, f"📨 Elküldve a nyomtatós gépre ({', '.join(gepek)}) – pár másodperc és kinyomtatja"
    return True, "📨 Elküldve – most egyik nyomtatós gép sem online; amint ott elindul a program, kinyomtatja (2 órán belül)"


def process_print_queue():
    """(nyomtatós gép, háttérszál) A többi gépről küldött címkék kinyomtatása."""
    if not is_print_host() or not _print_queue_lock.acquire(blocking=False):
        return
    try:
        res = vps_request("GET", f"{VPS_URL}/print/jobs", params={"pending": 1}, timeout=8)
        if res.status_code != 200:
            return
        for job in res.json().get("jobs") or []:
            r = vps_request("POST", f"{VPS_URL}/print/jobs/{job['id']}/claim", json={"gep": GEP_NEV}, timeout=8)
            if r.status_code != 200:
                continue
            ok, msg = _send_zpl_local(local_label_printer() or get_zebra_printer_name(), job.get("zpl", ""))
            try:
                vps_request("POST", f"{VPS_URL}/print/jobs/{job['id']}/result", json={"ok": ok, "uzenet": msg}, timeout=8)
            except Exception:
                pass
            kuldo, code, db = job.get("kuldo", ""), job.get("A", ""), job.get("peldany", 1)
            print_bus.msg.emit(f"🖨️ Címke nyomtatva a(z) {kuldo} gépről: {code} ({db} db)" if ok
                               else f"❌ A(z) {kuldo} gépről küldött címke ({code}) nyomtatása nem sikerült: {msg}")
            log_event(code, "CÍMKE NYOMTATÁS", f"Távoli nyomtatás a(z) {kuldo} gépről - " + (f"OK ({db} db)" if ok else f"HIBA: {msg}"))
    except Exception as e:
        print(f"[Távoli nyomtatás] Hiba: {e}")
    finally:
        _print_queue_lock.release()


# -------------------------------------------------------------------
# HELPERES ÉS ADATBÁZIS SEGÉDFÜGGVÉNYEK
# -------------------------------------------------------------------
def load_saved_zoom_factor():
    if os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                data = json.load(f)
                return float(data.get("zoom_factor", 1.0))
        except Exception:
            pass
    return 1.0

# -------------------------------------------------------------------
# VPS JELSZÓ (TOKEN): a szerver csak a jó tokennel érkező kéréseket fogadja el.
# A token NEM lehet a programkódban (a kód a GitHubon van) - a program mellett lévő
# api_token.txt fájlból olvassa (egyetlen sor). Ugyanez van a VPS-en: /opt/raktarapp/api_token.txt
# -------------------------------------------------------------------
API_TOKEN_FILE = os.path.join(BASE_DIR, "api_token.txt")


def _load_api_token():
    try:
        with open(API_TOKEN_FILE, "r", encoding="utf-8") as f:
            return f.read().strip()
    except Exception:
        return ""


API_TOKEN = _load_api_token()


def vps_request(method, url, **kwargs):
    """Kérés a VPS felé, a tokennel együtt (X-Raktar-Token fejléc)."""
    headers = dict(kwargs.pop("headers", None) or {})
    if API_TOKEN:
        headers["X-Raktar-Token"] = API_TOKEN
    return requests.request(method, url, headers=headers, **kwargs)


# A program által használt Python csomagok: (import név, pip csomag neve, mire kell, csak Windows-on kell-e).
# Ha egy új verzióhoz új csomag kell, IDE kell felvenni – a VERZIÓK fül ebből tudja, mit kell telepíteni.
REQUIRED_PACKAGES = [
    ("PyQt6", "PyQt6", "a program felülete", False),
    ("PyQt6.QtWebEngineWidgets", "PyQt6-WebEngine", "beágyazott oldalak (térkép, Sheets)", False),
    ("requests", "requests", "kapcsolat a VPS szerverrel", False),
    ("pandas", "pandas", "táblázatok kezelése", False),
    ("openpyxl", "openpyxl", "Excel fájlok, raktártérkép", False),
    ("barcode", "python-barcode", "vonalkód a címkén", False),
    ("qrcode", "qrcode", "QR kód a címkén", False),
    ("PIL", "Pillow", "képek, címke rajzolása, fotók", False),
    ("gspread", "gspread", "Google Sheets", False),
    ("google.oauth2", "google-auth", "Google bejelentkezés (Sheets)", False),
]

# Változásnapló (a legújabb legfelül). A VERZIÓK fülön ebből látszik, mi változott.
VALTOZASNAPLO = [
    ("v2026.10.30-nyomt1", "2026-10-30", [
        "Távoli nyomtatás: azon a gépen, amelyikhez nincs címkenyomtató kötve, a Nyomtatás gomb a nyomtatós gépre küldi "
        "a címkét, és az pár másodpercen belül kinyomtatja (visszajelzés: ✅ kinyomtatva a … gépen)",
        "A pywin32 csomag csak a nyomtatós gépen kell – a többi gépen nem akadályozza a frissítést",
        "VERZIÓK fül: 🖨️ Címkenyomtató beállítás (automatikus / ehhez a géphez van kötve / nincs nyomtató)",
    ]),
    ("v2026.10.29-web5", "2026-10-29", [
        "Excel Live: a 🖨️ Nyomtatás oszlop a sorszámokkal együtt külön, rögzített sávban a bal szélen – oldalra görgetéskor is "
        "mindig látszik, és a táblázat az A oszloppal kezdődik (semmit nem takar el)",
    ]),
    ("v2026.10.29-web3", "2026-10-29", [
        "Excel Live: a 🖨️ Nyomtatás gomb a táblázat fölé került, mindig látható (keskeny képernyőn sem csúszik ki)",
        "A „✍️ most tölti” jelzés kikerült",
    ]),
    ("v2026.10.29-web2", "2026-10-29", [
        "Excel Live javítás: a másik gépen felvitt sor akkor is azonnal megjelenik, ha ezen a gépen éppen azon a soron áll a kijelölés "
        "(eddig üres maradt, és üres címkét nyomtathatott)",
        "Nyomtatás előtt mindig a legfrissebb adatot tölti be a sorba",
    ]),
    ("v2026.10.29-web1", "2026-10-29", [
        "Gépnév: egyedi név gépenként (VERZIÓK fülön átírható)",
        "(Webes rendelések – Shopify / eBay – előkészítve, egyelőre kikapcsolva)",
    ]),
    ("v2026.10.29-elad1", "2026-10-29", [
        "Eladottak fül: a régi és a más gépen rögzített eladások is látszanak – a program behozza őket a Google táblázat "
        "Eladas_Ebay / Eladas_EbayKivul / Eladas_Webshop lapjairól (fül megnyitásakor, 5 percenként és gombbal)",
    ]),
    ("v2026.10.29-upd1", "2026-10-29", [
        "Új 🆕 VERZIÓK fül: automatikus frissítés a VPS-ről, egy gombnyomással",
        "Ha új verzió van, villogó piros sáv a program tetején + felugró ablak",
        "Python csomagok ellenőrzése, hiányzó csomag telepítése lépésről lépésre vagy automatikusan",
        "A régi verzió mentése: a gépen (regi_verziok mappa), a VPS-en és a saját Google Drive-on",
        "Korábbi verzió visszaállítása egy gombbal",
    ]),
    ("v2026.10.29-gep2", "2026-10-29", ["Leltár páter / polc szerint is (I / K oszlop)"]),
    ("v2026.10.29-gep1", "2026-10-29", ["Excel Live: a másik gépen felvitt sorok azonnal látszanak (ki, mikor), kattintható nyomtatással",
                                        "Excel Live: látszik, ki melyik soron dolgozik éppen"]),
    ("v2026.10.16-pda6", "2026-10-16", ["Új 📋 Leltár fül (régi: raklaponként, új: kézi / csippantós leltár, kártyák, visszavonás)"]),
]


# 🛒 Webes rendelések (Shopify + eBay) fül: egyelőre KIKAPCSOLVA. Bekapcsolás: True (és a VPS-en webshop_api.json,
# lásd vps/TELEPITES_webrendelesek.md).
WEBES_RENDELESEK_BE = False


def missing_packages(pkgs=None):
    """A ezen a gépen hiányzó csomagok (ugyanabban a formában, mint a REQUIRED_PACKAGES)."""
    import importlib.util
    out = []
    for imp, pip_name, mire, csak_win in (pkgs if pkgs is not None else REQUIRED_PACKAGES):
        if csak_win and os.name != "nt":
            continue
        try:
            ok = importlib.util.find_spec(imp) is not None
        except Exception:
            ok = False
        if not ok:
            out.append((imp, pip_name, mire, csak_win))
    return out


def _parse_module_const(text, name):
    """Egy másik programverzió szövegéből kiolvas egy egyszerű konstanst (pl. REQUIRED_PACKAGES) – futtatás nélkül."""
    import ast
    try:
        for node in ast.parse(text).body:
            if isinstance(node, ast.Assign) and any(isinstance(t, ast.Name) and t.id == name for t in node.targets):
                return ast.literal_eval(node.value)
    except Exception:
        return None
    return None


def python_console_exe():
    """A pip futtatásához a python.exe (a pythonw.exe nem ír ki semmit)."""
    exe = sys.executable or "python"
    if os.path.basename(exe).lower() == "pythonw.exe":
        cand = os.path.join(os.path.dirname(exe), "python.exe")
        if os.path.isfile(cand):
            return cand
    return exe


def _ver_key(v):
    """'v2026.10.16-pda2' -> összehasonlítható kulcs (a számokat számként kezeli)."""
    return [int(x) if x.isdigit() else x for x in re.split(r"(\d+)", str(v or ""))]


def send_heartbeat():
    """Jelentkezés a szerveren (gép neve + verzió). Visszaadja a legfrissebb verziókat, vagy None-t.
    Háttérszálból hívandó."""
    try:
        import getpass
        user = getpass.getuser()
    except Exception:
        user = ""
    res = vps_request("POST", f"{VPS_URL}/clients/heartbeat", timeout=8,
                      json={"tipus": "gep", "nev": GEP_NEV, "verzio": APP_VERSION,
                            "felhasznalo": user, "nyomtato": (local_label_printer() or "igen") if is_print_host() else ""})
    if res.status_code == 200:
        return res.json().get("legfrissebb") or {}
    return None


def push_naplo_to_server(max_rows=300):
    """A helyi Napló új bejegyzéseit feltölti a szerverre, hogy a PDA 'Előzmények' ablaka is lássa
    (minden gép a sajátját). Első alkalommal csak az utolsó 3000 bejegyzést. Háttérszálból hívandó."""
    conn = get_db_connection()
    try:
        last = load_config_value("naplo_push_id")
        if last is None:
            mx = conn.execute("SELECT COALESCE(MAX(id), 0) FROM naplo").fetchone()[0]
            last = max(0, mx - 3000)
        rows = conn.execute("SELECT id, idopont, r18_kod, esemeny, reszletek, forras FROM naplo WHERE id > ? "
                            "ORDER BY id LIMIT ?", (last, max_rows)).fetchall()
    finally:
        conn.close()
    if not rows:
        return 0
    sorok = [{k: safe_str(r[k]) if k != "id" else r["id"] for k in ("id", "idopont", "r18_kod", "esemeny", "reszletek", "forras")}
             for r in rows]
    res = vps_request("POST", f"{VPS_URL}/pda/history/push", timeout=15,
                      json={"gep": GEP_NEV, "sorok": sorok})
    if res.status_code == 200:
        save_config_value("naplo_push_id", rows[-1]["id"])
        return len(rows)
    return 0


def serve_pda_photo_requests():
    """A PDA által kért termékfotók feltöltése a szerverre (a NAS-ról, kicsinyítve). Háttérszálból hívandó.
    Visszaad: hány termék képeit töltötte fel."""
    if not PHOTOS_ENABLED:
        return 0
    res = vps_request("GET", f"{VPS_URL}/pda/photo_requests", timeout=8)
    if res.status_code != 200:
        return 0
    keresek = res.json().get("keresek") or []
    if not keresek:
        return 0
    if not photo_index.built_at:
        photo_index.build()            # első alkalommal megvárjuk a fotó-mappák beolvasását
    done = 0
    for k in keresek:
        code = norm_code(k.get("kod"))
        paths = photo_index.photos(code)[:8]
        files, hibak = [], 0
        for i, path in enumerate(paths):
            try:
                with Image.open(path) as im:
                    im = im.convert("RGB")
                    im.thumbnail((1280, 1280))
                    buf = io.BytesIO()
                    im.save(buf, "JPEG", quality=80)
                files.append(("fotok", (f"{i + 1}.jpg", buf.getvalue(), "image/jpeg")))
            except Exception:
                hibak += 1
        uzenet = "" if files else ("Nincs fénykép a NAS-on ehhez a termékhez" if not paths else "A képek nem olvashatók")
        if hibak and files:
            uzenet = f"{hibak} kép nem volt olvasható"
        try:
            vps_request("POST", f"{VPS_URL}/pda/photos/upload", timeout=60, files=files or None,
                        data={"kod": code, "uzenet": uzenet, "gep": GEP_NEV})
            done += 1
        except Exception as e:
            print(f"[PDA fotó] Feltöltési hiba ({code}): {e}")
    return done


def report_pda_result(tipus, rid, ok, uzenet, kod=""):
    """A PDA-ról jött eladás / módosítás eredményének visszaküldése a szerverre (háttérszálon).
    A PDA a Napló / Folyamat napló fülön mutatja (hiba esetén a kóddal, pl. A-21)."""
    def _send():
        try:
            vps_request("POST", f"{VPS_URL}/pda/result", timeout=8,
                        json={"tipus": tipus, "id": rid, "ok": bool(ok), "uzenet": uzenet, "kod": kod,
                              "gep": GEP_NEV})
        except Exception as e:
            print(f"[PDA] Eredmény visszaküldési hiba ({rid}): {e}")
    threading.Thread(target=_send, daemon=True).start()


def send_to_vps_async(endpoint, data):
    """Adatok beküldése a VPS API-ra külön háttérszálon (Zéró fagyás)"""
    def _send():
        try:
            url = f"{VPS_URL}{endpoint}"
            vps_request("POST", url, json=data, timeout=3)
        except Exception as e:
            print(f"VPS Szinkron hiba ({endpoint}): {e}")

    threading.Thread(target=_send, daemon=True).start()

_config_lock = threading.Lock()

def load_config_value(key, default=None):
    try:
        with _config_lock:
            if os.path.exists(CONFIG_FILE):
                with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                    return json.load(f).get(key, default)
    except Exception:
        pass
    return default

def save_config_value(key, value):
    """Egy kulcs mentése a beallitasok.json-ba (a többi kulcs megmarad). Háttérszálból is hívható."""
    try:
        with _config_lock:
            data = {}
            if os.path.exists(CONFIG_FILE):
                try:
                    with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                        data = json.load(f)
                except Exception:
                    data = {}
            data[key] = value
            with open(CONFIG_FILE, "w", encoding="utf-8") as f:
                json.dump(data, f, indent=4)
    except Exception:
        pass

def save_zoom_factor(zoom_val):
    save_config_value("zoom_factor", zoom_val)

def create_database_backup(tag="auto"):
    try:
        os.makedirs("backups", exist_ok=True)
        now_str = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        backup_path = f"backups/raktar_backup_{tag}_{now_str}.db"
        if os.path.exists(DB_FILE):
            shutil.copy2(DB_FILE, backup_path)
            return backup_path
    except Exception as e:
        print(f"Hiba a mentés készítésekor: {e}")
    return None

def get_db_connection():
    conn = sqlite3.connect(DB_FILE, timeout=30.0)
    conn.text_factory = lambda x: str(x, 'utf-8', 'ignore')
    conn.row_factory = sqlite3.Row
    cursor = conn.cursor()
    
    cursor.execute("PRAGMA journal_mode=WAL;")
    cursor.execute("PRAGMA busy_timeout=30000;")
    
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS termekek (
            rowid INTEGER PRIMARY KEY,
            A TEXT, B TEXT, C TEXT, D TEXT, E TEXT, F TEXT, G TEXT, H TEXT,
            I TEXT, J TEXT, K TEXT, L TEXT, M TEXT, N TEXT, O TEXT, P TEXT,
            szin TEXT
        )
    """)
    cursor.execute("PRAGMA table_info(termekek)")
    cols = [c[1] for c in cursor.fetchall()]
    if "szin" not in cols:
        cursor.execute("ALTER TABLE termekek ADD COLUMN szin TEXT")

    cursor.execute("""
        CREATE TABLE IF NOT EXISTS naplo (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            idopont TEXT,
            r18_kod TEXT,
            esemeny TEXT,
            reszletek TEXT
        )
    """)
    # A termék TELJES adata (A-P) az esemény előtt és után (JSON) - a Napló fül ebből mutatja,
    # mivel mi történt. Régi bejegyzéseknél üres.
    cursor.execute("PRAGMA table_info(naplo)")
    naplo_cols = [c[1] for c in cursor.fetchall()]
    for col in ("elotte", "utana", "forras"):
        if col not in naplo_cols:
            cursor.execute(f"ALTER TABLE naplo ADD COLUMN {col} TEXT")

    # "Kimenő postaláda": ha nincs net, ide kerül minden módosítás, amíg fel nem ment
    # a VPS-re / Google Sheets-re. Semmi nem vész el - csak várakozik a következő
    # automatikus újrapróbálkozásra (lásd flush_sync_outbox()).
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS sync_outbox (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            payload TEXT NOT NULL,
            vps_synced INTEGER NOT NULL DEFAULT 0,
            sheets_synced INTEGER NOT NULL DEFAULT 0,
            attempts INTEGER NOT NULL DEFAULT 0,
            created_at TEXT,
            last_attempt TEXT
        )
    """)

    # Törölt / kiürített termékek helyi archívuma: a teljes sor megmarad akkor is, ha a
    # Sheets-ből vagy a programból törölték. A 'alairas' (teljes A-P tartalom) alapján
    # ugyanazt a törlést csak egyszer küldjük ki a ToroltSor munkalapra.
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS torolt_archiv (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            idopont TEXT,
            r18_kod TEXT,
            forras TEXT,
            alairas TEXT UNIQUE,
            payload TEXT
        )
    """)

    # A figyelt (élő) táblázat utolsó ismert állapota - ebből derül ki, mi tűnt el azóta
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS figyelt_snapshot (
            forras TEXT NOT NULL,
            code TEXT NOT NULL,
            payload TEXT,
            PRIMARY KEY (forras, code)
        )
    """)

    # Eladások (PDA / Eladás fül 3 gombja) - az Eladottak fül ebből dolgozik, és innen megy ki
    # a 3 eladási munkalapra. valtozat: minden szerkesztés növeli (a szinkron csak a legutolsó
    # állapotot jelöli késznek); synced_platform: melyik munkalapon van most fent.
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS eladasok (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            eladas_id TEXT UNIQUE,
            datum TEXT, platform TEXT,
            A TEXT, C TEXT, D TEXT, E TEXT, G TEXT, L TEXT, M TEXT, N TEXT,
            vevo TEXT, ar TEXT, darab TEXT, tracking TEXT, megjegyzes TEXT,
            sheets_synced INTEGER NOT NULL DEFAULT 0,
            synced_platform TEXT,
            torolve INTEGER NOT NULL DEFAULT 0,
            valtozat INTEGER NOT NULL DEFAULT 0
        )
    """)

    conn.commit()
    return conn

NAPLO_COLS = "ABCDEFGHIJKLMNOP"

def get_row_snapshot(code, cursor=None):
    """A termék aktuális teljes adata (A-P) dict-ként, vagy None, ha nincs ilyen sor."""
    if not code:
        return None
    own = cursor is None
    try:
        if own:
            conn = get_db_connection()
            cursor = conn.cursor()
        cursor.execute(f"SELECT {', '.join(NAPLO_COLS)} FROM termekek WHERE UPPER(TRIM(A)) = ? OR UPPER(TRIM(B)) = ?",
                       (str(code).strip().upper(), str(code).strip().upper()))
        row = cursor.fetchone()
        if own:
            conn.close()
    except Exception:
        return None
    if not row:
        return None
    return {c: ("" if v is None or str(v).lower() in ("none", "nan") else str(v)) for c, v in zip(NAPLO_COLS, tuple(row))}

# Másoláskor (Excel / Word / weboldal) a cellába láthatatlan karakterek vagy "hosszú" kötőjelek
# kerülhetnek - ezek miatt az R18-kód nem ismerhető fel, pedig ugyanúgy néz ki.
_INVISIBLE_CHARS = dict.fromkeys(map(ord, "\ufeff\u200b\u200c\u200d\u2060\u00ad"), None)
_INVISIBLE_CHARS.update({ord(ch): "-" for ch in "\u2010\u2011\u2012\u2013\u2014\u2212"})
_INVISIBLE_CHARS[0xa0] = " "


def _gep_nev():
    """Ennek a gépnek a neve: a többi gépen ("ki vitte fel"), a PDA-n, a Gépek listában és a naplóban ez látszik.
    Ha két gépnek ugyanaz a Windows-neve, akkor is különbözzön: első indításkor a gépnév mögé egy rövid,
    véletlen azonosító kerül (beallitasok.json: gep_nev). A VERZIÓK fülön átírható (pl. "Raktár gép")."""
    n = load_config_value("gep_nev")
    if not n:
        n = f"{socket.gethostname() or 'gep'}-{os.urandom(2).hex()}"
        save_config_value("gep_nev", n)
    return str(n)


GEP_NEV = _gep_nev()
REMOTE_MARK_SEC = 2 * 3600   # ennyi ideig jelöli az Excel Live a másik gépen felvitt / módosított sorokat


def norm_pallet(v):
    """Egységes hely: nagybetű, egy szóköz, "KISRAKTÁR80" -> "KISRAKTÁR 80"."""
    v = re.sub(r"\s+", " ", str(v or "").upper()).strip()
    return re.sub(r"([^\W\d_])\s*(\d)", r"\1 \2", v)


def pallet_tokens(v):
    """Egy raklap-cellában több raklap is lehet: "493 494" = a 493-as ÉS a 494-es raklapon is van belőle.
    A "KISRAKTÁR 80" / "ENCS" típusú helyek EGY helynek számítanak."""
    return [norm_pallet(t) for t in re.findall(r"[^\W\d_]+\s*\d+|[^\s,;/+]+", str(v or "").upper())]


# A leltár helye: raklap (L) vagy páter / polc (I / K): "PATER 1 / POLC 12", egész páter: "PATER 1"
_PATER_RE = re.compile(r"^P[AÁ]TER\s*([^/]+?)\s*(?:/?\s*(?:POLC|OSZLOP)\s*(.+))?$")


def _norm_ik(v):
    return re.sub(r"\s+", " ", str(v or "").upper()).strip()


def norm_hely(v):
    m = _PATER_RE.match(_norm_ik(v))
    if m:
        return "PATER " + m.group(1).strip() + (" / POLC " + m.group(2).strip() if m.group(2) else "")
    return norm_pallet(v)


def on_hely(row, hely):
    m = _PATER_RE.match(hely)
    if m:
        return _norm_ik(row.get("I")) == m.group(1).strip() and (not m.group(2) or _norm_ik(row.get("K")) == m.group(2).strip())
    return hely in pallet_tokens(row.get("L"))


def hely_txt(hely):
    """'573' -> 'Raklap 573', 'PATER 1 / POLC 12' -> 'Páter 1 · Polc 12'"""
    h = str(hely or "")
    if h.startswith("PATER "):
        return h.replace("PATER ", "Páter ", 1).replace(" / POLC ", " · Polc ")
    return f"Raklap {h}"


def row_hely_txt(row):
    return " · ".join(x for x in (row.get("L") or "", f"Páter {row['I']}" if row.get("I") else "",
                                  f"Polc {row['K']}" if row.get("K") else "") if x) or "—"


def stock_of(f):
    m = re.search(r"\d+", str(f or ""))
    return int(m.group(0)) if m else 1


def norm_code(code):
    """R18-kód összehasonlításhoz: szóközök és láthatatlan karakterek nélkül, nagybetűvel."""
    return str(code or "").translate(_INVISIBLE_CHARS).strip().upper()


# A USB-s vonalkód-olvasó amerikai billentyűzetként "gépel", de a gépen magyar
# kiosztás van, ezért a '-' helyett 'ü', a '0' helyett 'ö' érkezik.
APP_VERSION = "v2026.10.30-nyomt1"

_SCANNER_MAP = str.maketrans({"ü": "-", "Ü": "-", "ö": "0", "Ö": "0"})


def _scanner_translate(text):
    """ü → -, ö → 0 csere (a felbontott „u + ¨” alakot is kezeli)."""
    return unicodedata.normalize("NFC", str(text or "")).translate(_SCANNER_MAP)


def fix_scanner_text(text):
    """Vonalkód-olvasó elgépeléseinek javítása (ü → -, ö → 0).

    Csak kód-jellegű szövegnél cserél (nincs benne szóköz, és a csere után
    nem marad ékezetes betű), így a kézzel beírt magyar szavak érintetlenek.
    """
    text = str(text or "").strip()
    if not text:
        return text
    fixed = _scanner_translate(text)
    if " " in fixed or not fixed.isascii():
        return text
    return fixed


def attach_scanner_fix(line_edit):
    """Vonalkód-olvasónak szánt mezőnél már gépelés közben kicseréli az ü → -, ö → 0 karaktereket."""
    def _fix(text):
        fixed = _scanner_translate(text)
        if fixed != text:
            pos = line_edit.cursorPosition()
            line_edit.blockSignals(True)
            line_edit.setText(fixed)
            line_edit.setCursorPosition(pos)
            line_edit.blockSignals(False)
    line_edit.textChanged.connect(_fix)

def dedupe_local_products():
    """DUPLIKÁLT R18-SZÁMOK összevonása a helyi adatbázisban: ugyanarra a kódra csak egy sor maradhat.
    Megmarad a legtöbb adattal kitöltött (egyenlőség esetén a legújabb) sor; a többi törlődik, de a
    teljes adatuk a Naplóba kerül. Utána egy egyedi index (zár) gondoskodik róla, hogy ne ismétlődjön.
    Visszaad: [(kód, törölt sor adata, megmaradt sor adata), ...]"""
    removed = []
    try:
        conn = get_db_connection()
        cursor = conn.cursor()
        cursor.execute(f"SELECT rowid AS rid, {', '.join(ROW_COLS)} FROM termekek "
                       f"WHERE A IS NOT NULL AND TRIM(A) != ''")
        groups = {}
        for r in cursor.fetchall():
            d = {c: safe_str(r[c]) for c in ROW_COLS}
            groups.setdefault(norm_code(d["A"]), []).append((r["rid"], d))
        for code, rows in groups.items():
            if len(rows) < 2:
                continue
            rows.sort(key=lambda x: (sum(1 for c in "CDEFGHIJKLMNOP" if x[1][c].strip()), x[0]), reverse=True)
            keep_rid, keep = rows[0]
            for rid, d in rows[1:]:
                cursor.execute("DELETE FROM termekek WHERE rowid = ?", (rid,))
                removed.append((code, d, keep))
        conn.commit()
        # Minden kód egységes alakra (szóköz / kisbetű nélkül) - így az egyedi index mindenhol hat.
        # (A duplikátumok már összevonva, ezért ez nem ütközhet.)
        cursor.execute("UPDATE termekek SET A = UPPER(TRIM(A)) WHERE A IS NOT NULL AND A != UPPER(TRIM(A))")
        conn.commit()
        try:
            cursor.execute("CREATE UNIQUE INDEX IF NOT EXISTS ux_termekek_a ON termekek(A) WHERE A IS NOT NULL AND A != ''")
            conn.commit()
        except Exception as e:
            print(f"[Duplikáció] Egyedi index nem hozható létre: {e}")
        conn.close()
    except Exception as e:
        print(f"[Duplikáció] Összevonási hiba: {e}")
        return removed
    for code, d, keep in removed:
        log_event(code, "DUPLIKÁCIÓ JAVÍTVA (kétszer szereplő R18-szám)",
                  "Ugyanez az R18-szám kétszer volt a helyi adatbázisban - a kevesebb adatot tartalmazó sor törölve, "
                  "a teljes adata az ELŐTTE oszlopban van.", d, keep, "Automatikus: duplikáció-javító")
    return removed

def describe_changes(elotte, utana):
    """'L (Raklap): 573 → 574 | M (Súly): 19 → 20' - csak a megváltozott oszlopok."""
    b, a = elotte or {}, utana or {}
    parts = []
    for c in NAPLO_COLS:
        ov, nv = str(b.get(c, "") or "").strip(), str(a.get(c, "") or "").strip()
        if ov != nv:
            parts.append(f"{c} ({COLUMN_LABELS.get(c, c)}): {ov or '∅'} → {nv or '∅'}")
    return " | ".join(parts)

# -------------------------------------------------------------------
# ÉRTHETŐ MAGYARÁZATOK a naplókhoz: mit jelent, mi a teendő
# -------------------------------------------------------------------
EVENT_EXPLANATIONS = [
    ("ELADÁS:", "A terméket eladták a PDA / Eladás fülön. A teljes eladás (vevő, ár, tracking) az 💰 Eladottak fülön "
                "látható és javítható. Ha a készlet elfogyott, a terméksor törlődött (külön TÖRLÉS bejegyzés)."),
    ("TÖRLÉS (EXCEL LIVE", "Az Excel Live munkalapon a Márka / Típus / Cikkszám cellákat kitörölték (Delete / Backspace "
                           "gombbal vagy felülírással). A régi adatok a ToroltSor lapra is felkerültek. Visszaállítás: "
                           "írd vissza az ELŐTTE oszlop értékeit a sorba."),
    ("TÖRLÉS (FELVÉTEL VISSZAVONVA", "Az Új termék fülön megnyomták az '↩️ Utolsó felvétel visszavonása' gombot. "
                                     "Az R18-kód megmaradt, az adatai kiürültek."),
    ("TÖRLÉS (ADATOK KIÜRÍTVE", "Egy kártyán megnyomták az '↺ Adatok kiürítése' gombot. Az R18-kód megmaradt."),
    ("TÖRLÉS (GOOGLE SHEETS", "Valaki KÖZVETLENÜL a Google táblázatban kitörölte a termék adatait, a program a letöltéskor "
                              "átvette. A régi adatok az ELŐTTE oszlopban vannak."),
    ("TÖRLÉS (MÁSIK GÉP", "A terméket egy MÁSIK gépen törölték / adták el, és a VPS-en keresztül ide is átjött."),
    ("TÖRLÉS", "A termék sora teljesen törlődött (törlés gomb vagy elfogyott eladás után). A teljes régi adat a "
               "ToroltSor lapon is megvan."),
    ("FELVÉTEL", "Új termék adatai kerültek egy R18-kód alá."),
    ("MÓDOSÍTÁS (GOOGLE SHEETS", "Valaki KÖZVETLENÜL a Google táblázatban írta át, a program a letöltéskor átvette."),
    ("MÓDOSÍTÁS (MÁSIK GÉP", "Egy MÁSIK gépen módosították, és a VPS-en keresztül ide is átjött."),
    ("MÓDOSÍTÁS (FÉLIG MENTETT", "Az automatikus mentés korábban csak egy részét (pl. csak a Márkát) küldte fel a szerverre. "
                                 "Ez NEM ütközés: a program a teljes adattal magától kiegészítette."),
    ("MÓDOSÍTÁS (R18-ÜTKÖZÉS", "Ezen az R18-számon a szerveren MÁSIK termék volt; az ütközés-ablakban a felülírást "
                               "választották, így most az új adat van fent. A régi adat a Részleteknél látszik."),
    ("MÓDOSÍTÁS", "A termék adatai megváltoztak. Az alsó részen sárga háttérrel látszik, mi változott."),
    ("TÖRÖLT SOR ARCHIVÁLVA", "A törölt termék teljes adata felkerült a ToroltSor munkalapra (biztonsági mentés)."),
    ("R18 ÜTKÖZÉS", "Ezt a terméket NEM sikerült elmenteni a szerverre, mert ott ugyanez az R18-szám már egy MÁSIK "
                    "termékhez tartozik. A program ablakban megkérdezi, felülírja-e az új adattal. Ha nem írtad "
                    "felül, vedd fel egy új R18-szám alatt - a teljes adatai az UTÁNA oszlopban vannak."),
    ("FIGYELMEZTETÉS: SORKONFLIKTUS", "Miközben te ezt a sort szerkesztetted, valaki más a Google táblázatban is "
                                      "átírta. Ellenőrizd, hogy a végleges adat a jó-e (ELŐTTE / UTÁNA)."),
    ("DUPLIKÁCIÓ JAVÍTVA", "Ugyanez az R18-szám kétszer szerepelt a helyi adatbázisban (ezért látszott duplán az Excel Live "
                           "munkalapon). A program összevonta: a több adatot tartalmazó sor maradt meg (UTÁNA), a törölt "
                           "sor teljes adata az ELŐTTE oszlopban van."),
    ("CÍMKE NYOMTATÁS", "Címkét nyomtattak ehhez a termékhez."),
    ("SZÍN MÓDOSÍTÁSA", "Az A-B cellák színjelölését változtatták meg."),
    ("ELADÁS MÓDOSÍTÁSA", "Egy korábbi eladás adatát (vevő, ár, tracking...) javították az 💰 Eladottak fülön."),
    ("ELADÁS TÖRLÉSE", "Egy eladást töröltek az 💰 Eladottak fülön (a raktárkészlet ettől NEM áll vissza)."),
]

def explain_event(esemeny):
    e = str(esemeny or "").upper()
    for key, text in EVENT_EXPLANATIONS:
        if e.startswith(key.upper()):
            return text
    return ""

# Folyamat napló: hibák / figyelmeztetések mellé automatikus magyarázat (minta -> magyarázat)
TECH_HINTS = [
    (r"Max retries|ConnectionError|Connection refused|timed out|Timeout|getaddrinfo|NameResolution|nincs net|Network is unreachable",
     "Nincs internet, vagy a szerver épp nem érhető el. Nem vész el semmi: a módosítás helyben el van mentve, "
     "és a program 20 mp-enként magától újrapróbálja."),
    (r"PERMISSION_DENIED|\b403\b",
     "A Google megtagadta a hozzáférést: a credentials.json szolgáltatásfiókja nincs megosztva ezzel a táblázattal "
     "(Szerkesztő jog kell)."),
    (r"\b404\b|not found|SpreadsheetNotFound|WorksheetNotFound",
     "A táblázat vagy munkalap nem található - ellenőrizd az ID-t / munkalap nevét a fájl elején."),
    (r"\b429\b|Quota|RATE_LIMIT",
     "Túl sok Google-kérés rövid idő alatt. A program kicsit később magától újrapróbálja, nincs teendő."),
    (r"database is locked",
     "Az adatbázist épp egy másik művelet írta. A program újrapróbálja, nincs teendő."),
    (r"ÜTKÖZÉS|ütközés|\b409\b",
     "Ugyanaz az R18-szám már egy MÁSIK termékhez tartozik a szerveren. A részletek a 📜 Napló fülön vannak."),
    (r"túl sok egyszerre|hibás letöltés",
     "Biztonsági fék: egyszerre gyanúsan sok termék tűnt el, ezért a program NEM archivált semmit. "
     "Nézd meg a táblázatot - ha valóban törölték őket, ez nem hiba."),
]

def technical_hint(text):
    for pattern, hint in TECH_HINTS:
        if re.search(pattern, str(text), re.IGNORECASE):
            return hint
    return ""

def log_event(r18_kod, esemeny, reszletek="", elotte=None, utana=None, forras=""):
    """Raktári esemény a Napló fülre. elotte / utana: a termék teljes A-P adata az esemény előtt / után.
    forras: HONNAN és HOGYAN történt (pl. 'Excel Live - billentyűzet', 'Másik gép (VPS)')."""
    try:
        conn = get_db_connection()
        cursor = conn.cursor()
        now_str = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        cursor.execute("INSERT INTO naplo (idopont, r18_kod, esemeny, reszletek, elotte, utana, forras) VALUES (?, ?, ?, ?, ?, ?, ?)",
                       (now_str, r18_kod, esemeny, reszletek,
                        json.dumps(elotte, ensure_ascii=False) if elotte is not None else None,
                        json.dumps(utana, ensure_ascii=False) if utana is not None else None,
                        forras or ""))
        conn.commit()
        conn.close()
    except Exception as e:
        print(f"Naplózási hiba: {e}")

# Az Excel-táblázat oszlopbetűinek (A-P) emberi olvasásra alkalmas neve - ugyanaz a
# leképezés, mint amit a táblázat fejléce (setHorizontalHeaderLabels) is használ.
COLUMN_LABELS = {
    "A": "F-Kód", "B": "Vonalkód", "C": "Márka", "D": "Típus", "E": "Cikkszám",
    "F": "DB", "G": "Állapot", "H": "Ár", "I": "Páter", "J": "Szabad",
    "K": "Polc", "L": "Raklap", "M": "Súly", "N": "eBay ID", "O": "Megjegyzés",
    "P": "Állapot részlet"
}

def format_row_for_log(row_data: dict) -> str:
    """Egy A-P oszlopkódos sor-dict-et teljes, emberi olvasásra alkalmas szöveggé alakít
    (pl. 'Márka: Bosch | Típus: Alternátor | Cikkszám: 1234') - a Napló ezt jeleníti meg."""
    parts = []
    for col in "ABCDEFGHIJKLMNOP":
        val = str(row_data.get(col, "")).strip()
        if val:
            parts.append(f"{COLUMN_LABELS.get(col, col)}: {val}")
    return " | ".join(parts) if parts else "(nincs adat)"

# -------------------------------------------------------------------
# OFFLINE-BIZTOS KIMENŐ SOR (OUTBOX): ha nincs net, a módosítás itt vár,
# amíg a flush_sync_outbox() sikeresen fel nem küldi a VPS-re / Sheets-re.
#
# Műveletek (a payload "_op" kulcsa):
#   - "upsert" (alapértelmezett): teljes A-P sor mentése ("_force": szándékos javítás,
#     a VPS ne vegye R18-ütközésnek, ha a Márka/Típus/Cikkszám megváltozik)
#   - "color":  {"A", "szin", "_color": Sheets színkód} - színjelölés
#   - "delete": {"A"} - termék törlése / eladás
# -------------------------------------------------------------------
ROW_COLS = "ABCDEFGHIJKLMNOP"
_flush_lock = threading.Lock()

def enqueue_outbox(row_data: dict):
    """Egy termék-sor módosítását betuszkolja a helyi kimenő sorba (mindig sikerül, mert helyi művelet)."""
    try:
        conn = get_db_connection()
        cursor = conn.cursor()
        now_str = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        cursor.execute(
            "INSERT INTO sync_outbox (payload, vps_synced, sheets_synced, attempts, created_at) VALUES (?, 0, 0, 0, ?)",
            (json.dumps(row_data, ensure_ascii=False), now_str)
        )
        conn.commit()
        conn.close()
    except Exception as e:
        print(f"[Outbox] Sorbaállítási hiba: {e}")

def enqueue_row_from_db(code, force=False):
    """A helyi adatbázisban lévő AKTUÁLIS sort (A-P) teszi az outboxba - a helyi mentés után hívandó."""
    try:
        conn = get_db_connection()
        cursor = conn.cursor()
        cursor.execute(f"SELECT {', '.join(ROW_COLS)} FROM termekek WHERE A = ? OR B = ?", (code, code))
        row = cursor.fetchone()
        conn.close()
    except Exception as e:
        print(f"[Outbox] Sor beolvasási hiba ({code}): {e}")
        return False
    if not row:
        return False
    payload = {c: safe_str(row[c]) for c in ROW_COLS}
    if force:
        payload["_force"] = True
    enqueue_outbox(payload)
    return True

def enqueue_color(code, szin, color_dict):
    enqueue_outbox({"_op": "color", "A": str(code).strip().upper(), "szin": szin, "_color": color_dict})

def enqueue_delete(code):
    enqueue_outbox({"_op": "delete", "A": str(code).strip().upper()})

def _row_has_product_data(row_data):
    """Van-e a sorban valódi termékadat (Márka / Típus / Cikkszám)."""
    return any(safe_str(row_data.get(c, "")).strip() for c in ("C", "D", "E"))

def archive_deleted_row(row_data, forras, cursor=None, force=False):
    """Egy törölt / kiürített termék TELJES (A-P) adatait elmenti a helyi archívumba és
    sorba állítja a Sheets 'ToroltSor' munkalapjára (offline-biztos, az outboxon át).
    Ugyanazt a tartalmat csak egyszer archiválja. True, ha új archív bejegyzés készült."""
    if not force and not ARCHIVE_PROGRAM_DELETIONS:
        return False
    data = {c: safe_str(row_data.get(c, "")).strip() for c in ROW_COLS}
    code = data["A"].upper()
    if not code or not _row_has_product_data(data):
        return False
    alairas = json.dumps(data, ensure_ascii=False, sort_keys=True)
    now_str = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    own_conn = cursor is None
    try:
        if own_conn:
            conn = get_db_connection()
            cursor = conn.cursor()
        cursor.execute(
            "INSERT OR IGNORE INTO torolt_archiv (idopont, r18_kod, forras, alairas, payload) VALUES (?, ?, ?, ?, ?)",
            (now_str, code, forras, alairas, alairas)
        )
        is_new = cursor.rowcount > 0
        if is_new:
            payload = dict(data)
            payload.update({"_op": "archive", "_ido": now_str, "_forras": forras})
            cursor.execute(
                "INSERT INTO sync_outbox (payload, vps_synced, sheets_synced, attempts, created_at) VALUES (?, 1, 0, 0, ?)",
                (json.dumps(payload, ensure_ascii=False), now_str)
            )
            cursor.execute("INSERT INTO naplo (idopont, r18_kod, esemeny, reszletek, elotte, forras) VALUES (?, ?, ?, ?, ?, ?)",
                           (now_str, code, "TÖRÖLT SOR ARCHIVÁLVA", "", json.dumps(data, ensure_ascii=False), forras))
        if own_conn:
            conn.commit()
            conn.close()
        return is_new
    except Exception as e:
        print(f"[Archív] Hiba ({code}): {e}")
        return False

# -------------------------------------------------------------------
# ELADÁSOK -> 3 ELADÁSI MUNKALAP (offline-biztos: ami nem ment ki, sheets_synced=0 marad)
# -------------------------------------------------------------------
_sales_sync_lock = threading.Lock()

def _sales_last_col_letter():
    return chr(ord("A") + len(SALES_COLUMNS) - 1)

def get_sales_worksheet(platform, cache):
    """A platformhoz tartozó eladási munkalap (cache: {platform: ws}); ha nincs, létrehozza fejléccel."""
    if platform in cache:
        return cache[platform]
    import gspread
    from google.oauth2.service_account import Credentials
    cfg = SALES_SHEETS[platform]
    if "_client" not in cache:
        scopes = ["https://www.googleapis.com/auth/spreadsheets"]
        creds = Credentials.from_service_account_file(CREDENTIALS_FILE, scopes=scopes)
        cache["_client"] = gspread.authorize(creds)
    book = cache["_client"].open_by_key(cfg["spreadsheet_id"])
    try:
        ws = book.worksheet(cfg["worksheet"])
    except gspread.exceptions.WorksheetNotFound:
        ws = book.add_worksheet(title=cfg["worksheet"], rows=1000, cols=len(SALES_COLUMNS))
    if not ws.row_values(1):
        ws.update(range_name=f"A1:{_sales_last_col_letter()}1", values=[[h for _, h in SALES_COLUMNS]])
    cache[platform] = ws
    return ws

def _sales_row_values(row):
    vals = []
    for key, _ in SALES_COLUMNS:
        if key == "platform":
            vals.append(SALES_SHEETS.get(row["platform"], {}).get("cim", row["platform"]))
        else:
            vals.append(safe_str(row[key]))
    return vals

def sync_sales_to_sheets():
    """A még fel nem töltött eladásokat (új / szerkesztett / törölt / platformot váltott)
    kiírja a megfelelő eladási munkalapra. Visszaadja a sikeresen szinkronizált sorok számát."""
    if not SHEETS_ENABLED:
        return 0
    if not _sales_sync_lock.acquire(blocking=False):
        return 0
    done = 0
    try:
        conn = get_db_connection()
        cursor = conn.cursor()
        cursor.execute("SELECT * FROM eladasok WHERE sheets_synced = 0 ORDER BY id")
        pending = [dict(r) for r in cursor.fetchall()]
        conn.close()
        if not pending:
            return 0

        cache, id_cols = {}, {}

        def ids_of(platform):
            if platform not in id_cols:
                id_cols[platform] = [str(x).strip() for x in get_sales_worksheet(platform, cache).col_values(1)]
            return id_cols[platform]

        for row in pending:
            try:
                eid = row["eladas_id"]
                old_pf = row.get("synced_platform") or ""
                new_pf = row["platform"]
                # 1. A régi helyéről törlés (platformváltás vagy törölt eladás)
                if old_pf in SALES_SHEETS and (old_pf != new_pf or row["torolve"]):
                    ids = ids_of(old_pf)
                    if eid in ids:
                        idx = ids.index(eid)
                        get_sales_worksheet(old_pf, cache).delete_rows(idx + 1)
                        ids.pop(idx)
                conn2 = get_db_connection()
                cur2 = conn2.cursor()
                if row["torolve"]:
                    cur2.execute("DELETE FROM eladasok WHERE id = ? AND valtozat = ?", (row["id"], row["valtozat"]))
                    conn2.commit()
                    conn2.close()
                    done += 1
                    continue
                conn2.close()
                if new_pf not in SALES_SHEETS:
                    continue
                # 2. Beírás / frissítés az aktuális platform munkalapján (Eladás ID alapján)
                ws = get_sales_worksheet(new_pf, cache)
                ids = ids_of(new_pf)
                vals = _sales_row_values(row)
                if eid in ids:
                    r = ids.index(eid) + 1
                    ws.update(range_name=f"A{r}:{_sales_last_col_letter()}{r}", values=[vals], value_input_option="RAW")
                else:
                    ws.append_row(vals, value_input_option="RAW")
                    ids.append(eid)
                conn2 = get_db_connection()
                cur2 = conn2.cursor()
                cur2.execute("UPDATE eladasok SET sheets_synced = 1, synced_platform = ? WHERE id = ? AND valtozat = ?",
                             (new_pf, row["id"], row["valtozat"]))
                conn2.commit()
                conn2.close()
                done += 1
            except Exception as e:
                print(f"[Eladás szinkron] Hiba ({row.get('eladas_id')}): {e}")
        return done
    except Exception as e:
        print(f"[Eladás szinkron] Hiba: {e}")
        return done
    finally:
        _sales_sync_lock.release()

def _norm_hdr(h):
    return re.sub(r"[^a-z0-9áéíóöőúüű]", "", str(h or "").lower())


def pull_sales_from_sheets():
    """A 3 eladási munkalapon lévő eladásokat behozza a helyi adatbázisba (az Eladottak fül ebből dolgozik):
    a régi (a program előtti) és a MÁS gépen rögzített eladások is látszanak így.
    - új eladás a lapon -> felvesszük (fent van, nem kell feltölteni; a készlet NEM változik)
    - meglévő, itt nem szerkesztett eladás -> a lap szerinti adatra frissítjük (pl. másik gépen javították)
    - itt szerkesztett, még fel nem töltött eladást NEM írunk felül
    - ha egy eladás egyik lapon sincs már (másik gépen törölték), itt is eltűnik
    - ha a lapon egy sornak nincs Eladás ID-je (régi, kézzel beírt sor), kap egyet (az A oszlopba írjuk)
    Visszaad: {"uj", "frissitett", "torolt", "id_potolva", "hibak"}"""
    out = {"uj": 0, "frissitett": 0, "torolt": 0, "id_potolva": 0, "hibak": []}
    if not SHEETS_ENABLED:
        out["hibak"].append("A Google Sheets kapcsolat ki van kapcsolva ezen a gépen")
        return out
    if not _sales_sync_lock.acquire(timeout=60):
        out["hibak"].append("Éppen fut egy eladás-feltöltés – próbáld újra pár másodperc múlva")
        return out
    try:
        cache = {}
        by_hdr = {_norm_hdr(h): k for k, h in SALES_COLUMNS}
        lapon = {}          # eladas_id -> (platform, {mező: érték})
        minden_lap_ok = True
        for pf in SALES_SHEETS:
            try:
                ws = get_sales_worksheet(pf, cache)
                values = ws.get_all_values()
            except Exception as e:
                minden_lap_ok = False
                out["hibak"].append(f"{SALES_SHEETS[pf]['worksheet']}: {e}")
                continue
            if not values:
                continue
            hdr = values[0]
            colmap = {}
            for i, h in enumerate(hdr):
                k = by_hdr.get(_norm_hdr(h))
                if k and k not in colmap.values():
                    colmap[i] = k
            if "eladas_id" not in colmap.values():          # nincs felismerhető fejléc -> a szokásos sorrend
                colmap = {i: k for i, (k, _) in enumerate(SALES_COLUMNS)}
            id_col = [i for i, k in colmap.items() if k == "eladas_id"][0]
            potlas = []
            for ri, row in enumerate(values[1:], start=2):
                rec = {k: (row[i].strip() if i < len(row) else "") for i, k in colmap.items()}
                if not any(rec.get(k) for k in ("A", "C", "D", "E", "datum")):
                    continue                                   # üres sor
                eid = rec.get("eladas_id", "")
                if not eid:
                    base = "|".join(rec.get(k, "") for k in ("datum", "A", "C", "D", "E", "vevo", "ar"))
                    eid = f"R{hashlib.sha1((pf + base).encode('utf-8')).hexdigest()[:12].upper()}-{rec.get('A') or 'X'}"
                    rec["eladas_id"] = eid
                    potlas.append((ri, eid))
                rec["platform"] = pf
                lapon.setdefault(eid, (pf, rec))
            if potlas:
                try:
                    col = chr(ord("A") + id_col)
                    ws.batch_update([{"range": f"{col}{ri}", "values": [[eid]]} for ri, eid in potlas], value_input_option="RAW")
                    out["id_potolva"] += len(potlas)
                except Exception as e:
                    out["hibak"].append(f"{SALES_SHEETS[pf]['worksheet']}: az Eladás ID-k beírása nem sikerült ({e})")

        mezok = [k for k, _ in SALES_COLUMNS if k != "eladas_id"]
        conn = get_db_connection()
        cur = conn.cursor()
        cur.execute("SELECT * FROM eladasok")
        helyi = {r["eladas_id"]: dict(r) for r in cur.fetchall()}
        for eid, (pf, rec) in lapon.items():
            h = helyi.get(eid)
            if h is None:
                cur.execute(f"INSERT INTO eladasok (eladas_id, {', '.join(mezok)}, sheets_synced, synced_platform) "
                            f"VALUES (?, {', '.join('?' * len(mezok))}, 1, ?)",
                            (eid, *[rec.get(k, "") for k in mezok], pf))
                out["uj"] += 1
            elif h["sheets_synced"] and not h["torolve"]:
                if any(safe_str(h.get(k)) != rec.get(k, "") for k in mezok) or h.get("synced_platform") != pf:
                    cur.execute(f"UPDATE eladasok SET {', '.join(k + '=?' for k in mezok)}, synced_platform=? "
                                f"WHERE eladas_id=? AND sheets_synced=1",
                                (*[rec.get(k, "") for k in mezok], pf, eid))
                    out["frissitett"] += cur.rowcount
        if minden_lap_ok:
            for eid, h in helyi.items():
                if eid not in lapon and h["sheets_synced"] and not h["torolve"] and h.get("synced_platform"):
                    cur.execute("DELETE FROM eladasok WHERE eladas_id=? AND sheets_synced=1", (eid,))
                    out["torolt"] += cur.rowcount
        conn.commit()
        conn.close()
        return out
    except Exception as e:
        out["hibak"].append(str(e))
        return out
    finally:
        _sales_sync_lock.release()


def sync_sales_async():
    threading.Thread(target=sync_sales_to_sheets, daemon=True).start()

def archive_row_from_db(code, forras):
    """A helyi adatbázisban lévő sort archiválja (törlés / kiürítés ELŐTT hívandó)."""
    try:
        conn = get_db_connection()
        cursor = conn.cursor()
        cursor.execute(f"SELECT {', '.join(ROW_COLS)} FROM termekek WHERE A = ? OR B = ?", (code, code))
        row = cursor.fetchone()
        conn.close()
    except Exception as e:
        print(f"[Archív] Sor beolvasási hiba ({code}): {e}")
        return False
    if not row:
        return False
    return archive_deleted_row({c: row[c] for c in ROW_COLS}, forras)

def flush_outbox_async():
    threading.Thread(target=flush_sync_outbox, daemon=True).start()

def get_pending_outbox_codes(cursor, csak_vps=False):
    """Azok az R18-kódok, amelyeknek van még fel nem töltött módosítása az outboxban.
    Ezeket a letöltések (VPS / Sheets) NEM írhatják felül a régebbi felhős állapottal.
    csak_vps=True: csak a VPS-re még fel NEM ment módosítások számítanak (ami a VPS-en már fent van, de a Sheets-be
    még nem ment fel, az ne tartsa vissza a másik gépről érkező újabb adatot)."""
    cursor.execute("SELECT payload FROM sync_outbox" + (" WHERE vps_synced = 0" if csak_vps else ""))
    codes = set()
    for row in cursor.fetchall():
        try:
            payload = json.loads(row[0])
            if payload.get("_op") == "archive":
                continue  # a ToroltSor-ra váró archív bejegyzés nem védi a Raktar-sort
            codes.add(str(payload.get("A", "")).strip().upper())
        except Exception:
            pass
    codes.discard("")
    return codes

def flush_sync_outbox():
    """Megpróbálja kiküldeni a még el nem küldött sorokat a VPS-re és (ha be van kapcsolva) a Sheets-re.
    Egy sort csak akkor töröl a sorból, ha MINDEN cél felé sikerült. Ha nincs net, a sor egyszerűen
    a helyén marad a következő hívásig (auto_save_timer / outbox_flush_timer) - semmi nem vész el.
    A műveletek a felvitel sorrendjében mennek ki; egyszerre csak egy küldés fut (_flush_lock),
    így két párhuzamos hívás nem fűzhet a Sheets végére kétszer ugyanolyan új sort.

    ÜTKÖZÉS-KEZELÉS: ha a szerver 409-et ad vissza (mert ez az R18-szám már MÁS termékkel foglalt),
    a tételt NEM próbáljuk újra automatikusan - úgysem menne át, amíg valaki kézzel nem rendel
    hozzá új R18-számot. Helyette a teljes adatait naplózzuk (Napló fül) és jelezzük az UI-n.

    Visszaadja az ütközésbe futott tételek listáját (dict-ek)."""
    with _flush_lock:
        return _flush_sync_outbox_locked()

def _server_conflict_existing(server_msg):
    """A szerver 409-es üzenetéből: '... foglalt egy MASIK termekkel (REXROTH ).' -> 'REXROTH'
    (a szerveren most lévő Márka / Típus / Cikkszám)."""
    m = re.findall(r"\(([^()]*)\)", str(server_msg or ""))
    return re.sub(r"\s+", " ", m[-1]).strip() if m else None

def _is_fill_in_of(server_existing, row_data):
    """Igaz, ha a szerveren lévő sor csak egy FÉLIG kitöltött (vagy üres) változata a mostaninak:
    a szerveren lévő Márka / Típus / Cikkszám mind üres vagy egyezik a most mentendővel.
    (Tipikus eset: az automatikus mentés már akkor felküldte a sort, amikor még csak a Márka volt beírva.)"""
    if server_existing is None:
        return False
    target = server_existing.upper()
    vals = [re.sub(r"\s+", " ", str(row_data.get(c, "") or "")).strip().upper() for c in ("C", "D", "E")]
    for mask in range(8):
        cand = " ".join(v for i, v in enumerate(vals) if mask & (1 << i) and v)
        if cand == target:
            return True
    return False

def _send_outbox_to_vps(op, row_data):
    """Visszaad: 'ok', 'conflict' vagy 'retry' (nincs net / szerverhiba -> később újra)."""
    code_a = str(row_data.get("A", "")).strip().upper()
    if op == "archive":
        return "ok"  # csak a Sheets ToroltSor munkalapjára megy
    try:
        if op == "delete":
            res = vps_request("POST", f"{VPS_URL}/delete_item", json={"A": code_a}, timeout=3)
        elif op == "color":
            res = vps_request("POST", f"{VPS_URL}/update_color", json={"A": code_a, "szin": row_data.get("szin", "")}, timeout=3)
        else:
            body = {k: v for k, v in row_data.items() if not k.startswith("_")}
            body["gep"] = GEP_NEV
            if row_data.get("_force"):
                body["force"] = True
            res = vps_request("POST", f"{VPS_URL}/update_item", json=body, timeout=3)
    except Exception:
        return "retry"

    if res.status_code == 200:
        return "ok"
    if res.status_code == 409 and op == "upsert":
        details = format_row_for_log(row_data)
        server_msg, existing_dict = "", None
        try:
            js = res.json()
            server_msg = js.get("message", "")
            existing_dict = js.get("existing")
        except Exception:
            pass
        if isinstance(existing_dict, dict):
            existing = " ".join(str(existing_dict.get(k) or "").strip() for k in ("C", "D", "E")
                                if str(existing_dict.get(k) or "").strip())
        else:
            existing = _server_conflict_existing(server_msg)
        new_row = {c: str(row_data.get(c, "")) for c in ROW_COLS}

        # 1) A szerveren csak egy FÉLIG kitöltött / üres sor van ezzel az R18-számmal (pl. a Márka már
        #    felment, a Típus / Cikkszám még nem): ez nem másik termék -> automatikusan kiegészítjük.
        if _is_fill_in_of(existing, row_data):
            try:
                body = {k: v for k, v in row_data.items() if not k.startswith("_")}
                body["force"] = True
                body["gep"] = GEP_NEV
                res2 = vps_request("POST", f"{VPS_URL}/update_item", json=body, timeout=3)
            except Exception:
                return "retry"
            if res2.status_code == 200:
                log_event(code_a, "MÓDOSÍTÁS (félig mentett sor kiegészítve)",
                          f"A szerveren eddig csak ez volt: '{existing or '(üres)'}' - most a teljes adat felment.",
                          None, new_row, "Automatikus: szerver (VPS) - nem ütközés, csak kiegészítés")
                return "ok"
            return "retry"

        # 2) Valódi ütközés: másik termék van ott -> a felhasználó dönt (felülírás / kihagyás)
        row_data["_server_existing"] = existing or server_msg
        log_event(code_a, "R18 ÜTKÖZÉS - NEM MENTVE", f"A szerveren most: '{existing or server_msg}' | Új (nem mentett): {details}",
                  None, new_row, "Szerver (VPS) elutasította - döntésre vár (felülírás / kihagyás)")
        return "conflict"
    return "retry"

# -------------------------------------------------------------------
# IGAZÍTÁS a Google táblázatban és az Excel Live munkalapon:
#   M (Súly): mindig JOBBRA | L (Raklap): JOBBRA, ha csak szám van benne (pl. 573, 8,65), különben BALRA
# -------------------------------------------------------------------
_PURE_NUMBER = re.compile(r"^\s*\d+([.,]\d+)?\s*$")

def is_pure_number(val):
    return bool(_PURE_NUMBER.match(str(val or "")))

def sheet_alignment_requests(row_num, l_val):
    """batch_format kérések egy Sheets-sorhoz (L feltételesen, M mindig jobbra)."""
    return [
        {"range": f"L{row_num}", "format": {"horizontalAlignment": "RIGHT" if is_pure_number(l_val) else "LEFT"}},
        {"range": f"M{row_num}", "format": {"horizontalAlignment": "RIGHT"}},
    ]

def fix_sheet_alignment_all(log=print):
    """EGYSZERI javítás: a Google táblázat összes sorában beállítja az L / M igazítást.
    Az egymást követő azonos igazítású L-cellákat egy tartományba vonja (kevés Google-kérés)."""
    sheet = get_sheet()
    if sheet is None:
        return False
    codes = sheet.col_values(1)
    l_vals = sheet.col_values(12)
    n = len(codes)
    if n < 2:
        return True
    reqs = [{"range": f"M2:M{n}", "format": {"horizontalAlignment": "RIGHT"}}]
    start, cur = 2, None
    for r in range(2, n + 2):
        al = None if r > n else ("RIGHT" if is_pure_number(l_vals[r - 1] if r - 1 < len(l_vals) else "") else "LEFT")
        if al != cur:
            if cur is not None:
                reqs.append({"range": f"L{start}:L{r - 1}", "format": {"horizontalAlignment": cur}})
            start, cur = r, al
    for i in range(0, len(reqs), 500):
        sheet.batch_format(reqs[i:i + 500])
    log(f"✔ Google táblázat igazítás javítva: {n - 1} sor (M jobbra, L számnál jobbra / szövegnél balra)", "SUCCESS")
    return True

def _send_outbox_to_sheets(op, row_data, sheet, all_a, deleted_ws_holder=None):
    """A Sheets-be írja a műveletet. Az all_a (A oszlop) listát helyben frissíti. True, ha sikerült."""
    code_a = str(row_data.get("A", "")).strip().upper()
    if not code_a:
        return True
    try:
        if op == "archive":
            # Törölt termék teljes adatai a 'ToroltSor' munkalap végére (A-P ugyanott, mint a Raktar lapon)
            if deleted_ws_holder is None:
                deleted_ws_holder = {}
            if deleted_ws_holder.get("ws") is None:
                ws = get_deleted_sheet()
                # A már archivált sorok (A-P) - egy másik gép se írhassa be ugyanazt kétszer
                existing = set()
                for r in ws.get_all_values()[1:]:
                    existing.add(tuple((list(r) + [""] * 16)[:16]))
                deleted_ws_holder["ws"], deleted_ws_holder["existing"] = ws, existing
            data_vals = [str(row_data.get(k, "")) for k in ROW_COLS]
            if tuple(data_vals) not in deleted_ws_holder["existing"]:
                row_vals = data_vals + [row_data.get("_ido", ""), row_data.get("_forras", "")]
                deleted_ws_holder["ws"].append_row(row_vals, value_input_option="RAW")
                deleted_ws_holder["existing"].add(tuple(data_vals))
        elif op == "delete":
            if code_a in all_a:
                idx = all_a.index(code_a)
                sheet.delete_rows(idx + 1)
                all_a.pop(idx)
        elif op == "color":
            if code_a in all_a:
                target_row = all_a.index(code_a) + 1
                sheet.format(f"A{target_row}:B{target_row}", {"backgroundColor": row_data.get("_color", COLOR_WHITE)})
        else:
            if code_a in all_a:
                target_row = all_a.index(code_a) + 1
            else:
                target_row = len(all_a) + 1
                all_a.append(code_a)
            row_vals = [row_data.get(k, "") for k in ROW_COLS]
            sheet.update(range_name=f"A{target_row}:P{target_row}", values=[row_vals])
            # L / M igazítás - a flush végén egyetlen kérésben megy ki (lásd _flush_sync_outbox_locked)
            if deleted_ws_holder is not None:
                deleted_ws_holder.setdefault("_fmt", []).extend(sheet_alignment_requests(target_row, row_data.get("L", "")))
        return True
    except Exception as e:
        print(f"[Outbox] Sheets írási hiba ({op} {code_a}): {e}")
        return False

def _flush_sync_outbox_locked():
    conflicts = []
    try:
        conn = get_db_connection()
        conn.row_factory = sqlite3.Row
        cursor = conn.cursor()
        cursor.execute("SELECT id, payload, vps_synced, sheets_synced FROM sync_outbox WHERE vps_synced = 0 OR sheets_synced = 0 ORDER BY id")
        pending = cursor.fetchall()
        conn.close()
    except Exception as e:
        print(f"[Outbox] Olvasási hiba: {e}")
        return conflicts

    if not pending:
        return conflicts

    sheet, all_a = None, []
    deleted_ws_holder = {}
    if SHEETS_ENABLED:
        try:
            sheet = get_sheet()
            all_a = [str(x).strip().upper() for x in sheet.col_values(1)]
        except Exception as e:
            print(f"[Outbox] Sheets kapcsolat még nem elérhető, később újra próbálja: {e}")

    for row in pending:
        try:
            row_data = json.loads(row["payload"])
        except Exception:
            continue

        op = row_data.get("_op", "upsert")
        vps_ok = bool(row["vps_synced"])
        sheets_ok = bool(row["sheets_synced"]) or not SHEETS_ENABLED
        is_conflict = False

        def do_vps():
            nonlocal vps_ok, is_conflict
            if not vps_ok:
                result = _send_outbox_to_vps(op, row_data)
                vps_ok = result == "ok"
                if result == "conflict":
                    is_conflict = True
                    conflicts.append(row_data)

        def do_sheets():
            nonlocal sheets_ok
            if not sheets_ok and not is_conflict and (sheet is not None or op == "archive"):
                sheets_ok = _send_outbox_to_sheets(op, row_data, sheet, all_a, deleted_ws_holder)

        # Törlésnél előbb a Sheets-ből töröljük: így mire a másik gép a VPS-től megtudja
        # a törlést, a Sheets-ből letöltött adatok már nem hozzák vissza a tételt.
        if op == "delete":
            do_sheets()
            do_vps()
        else:
            do_vps()
            do_sheets()

        if is_conflict:
            sheets_ok = True  # ütközésnél a Sheets-küldést is elengedjük, ez a tétel úgyis kikerül a sorból

        try:
            conn2 = get_db_connection()
            cur2 = conn2.cursor()
            if is_conflict or (vps_ok and sheets_ok):
                cur2.execute("DELETE FROM sync_outbox WHERE id = ?", (row["id"],))
            else:
                cur2.execute(
                    "UPDATE sync_outbox SET vps_synced=?, sheets_synced=?, attempts=attempts+1, last_attempt=? WHERE id=?",
                    (int(vps_ok), int(sheets_ok), datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"), row["id"])
                )
            conn2.commit()
            conn2.close()
        except Exception as e:
            print(f"[Outbox] Állapot-frissítési hiba: {e}")

    if sheet is not None and deleted_ws_holder.get("_fmt"):
        try:
            sheet.batch_format(deleted_ws_holder["_fmt"])
        except Exception as e:
            print(f"[Outbox] Igazítási hiba (nem kritikus): {e}")

    if conflicts:
        try:
            outbox_signal_bus.conflicts_detected.emit(conflicts)
        except Exception as e:
            print(f"[Outbox] Jelzési hiba: {e}")

    return conflicts

def safe_str(val):
    # Ha egy lekérdezésben két azonos nevű oszlop van, a pandas Series-t ad vissza egyetlen érték helyett
    if isinstance(val, pd.Series):
        val = val.iloc[0] if len(val) else None
    elif isinstance(val, (list, tuple)):
        val = val[0] if val else None
    if val is None:
        return ""
    try:
        if pd.isna(val):
            return ""
    except (TypeError, ValueError):
        pass
    s = str(val).strip()
    if s.lower() in ("nan", "none", "null"):
        return ""
    if s.endswith(".0"):
        s = s[:-2]
    return s

def clean_part_number(text):
    if not text:
        return ""
    return re.sub(r'[\s\-_/.,]+', '', str(text)).upper()

def get_best_code(row):
    for col_name in ["A", "B"]:
        val = safe_str(row.get(col_name) if hasattr(row, 'get') else row[col_name])
        if val and val.startswith("R18-"):
            return val
    return safe_str(row.get("A") if hasattr(row, 'get') else row["A"]) or "ISMERETLEN"

# -------------------------------------------------------------------
# OKOS CÍMKE SZÖVEG VÁLASZTÓ ÉS CIKKSZÁM KINYERŐ ALGORITMUSOK
# -------------------------------------------------------------------
def get_smart_part_number(d_val, e_val):
    e_str = safe_str(e_val).strip()
    d_str = safe_str(d_val).strip()

    if '-' in e_str and bool(re.search(r'\d', e_str)):
        return e_str[:25]

    if '-' in d_str and bool(re.search(r'\d', d_str)):
        return d_str[:25]

    if e_str and bool(re.search(r'\d', e_str)):
        return e_str[:25]

    return (d_str or e_str)[:25]

def get_best_label_text(tipus_text="", cikkszam_text=""):
    return get_smart_part_number(tipus_text, cikkszam_text)

# -------------------------------------------------------------------
# CÍMKE ALSÓ SORA: a D (Típus) / E (Cikkszám) mezőkből a lényeges alkatrészszám kinyerése
# -------------------------------------------------------------------
LABEL_SUBTEXT_MAX = 19          # LEGFELJEBB ennyi karakter kerül a címke alsó sorába (minden címkén)
_LABEL_UNITS = {"V", "A", "MA", "KA", "W", "KW", "VA", "VAC", "VDC", "AC", "DC", "HZ", "KHZ",
                "BAR", "MM", "CM", "M", "KG", "G", "NM", "RPM", "S", "MS", "UF", "OHM"}
# Ezek után álló szám nem alkatrészszám (szabvány), ill. ezek nem előtagjai az alkatrészszámnak
_LABEL_STANDARDS = {"IEC", "EN", "DIN", "VDE", "ISO", "UL"}
_LABEL_NOT_PREFIX = _LABEL_STANDARDS | {"TYPE", "TYP", "ID", "NO", "NR", "ART", "PN", "P/N", "MNR",
                                        "REV", "SN", "S/N", "ORD", "PART", "UA", "CE", "UJ", "NEW",
                                        "AND", "ES", "ÉS", "VAN", "DB"}
_SIEMENS_PREFIX = re.compile(r"^\d[A-Z]{2}[0-9A-Z]{0,2}$")   # pl. 6ES7, 6EP1, 6FX, 6SC, 6FC5
_MEASUREMENT = re.compile(r"^(DC|AC)?[\d.,]+[A-Z]*([-/~]+(DC|AC)?[\d.,]*[A-Z]*)*$")

def _is_measurement(tok):
    """'0-100BAR', 'DC12-24V', '24-28.8V/20A' - mértékegységes érték, nem alkatrészszám."""
    if not _MEASUREMENT.match(tok) or re.search(r"\d{5,}", tok):
        return False
    letters = re.findall(r"[A-Z]+", tok)
    return bool(letters) and all(l in _LABEL_UNITS for l in letters)

def _is_dashed_part(tok, prev):
    """Kötőjeles alkatrészszám: van benne kötőjel és szám (pl. A06B-6087-H130, 77-06-70, S-K65)."""
    if "-" not in tok or not re.search(r"\d", tok) or prev in _LABEL_STANDARDS or _is_measurement(tok):
        return False
    core = tok.strip("-/")
    if not re.search(r"[A-Z]", core):
        return len(core) >= 6            # csak számos: '77-06-70' igen, '01-03' nem
    return len(core) >= 5                # 'S-K65', 'NTC-1' igen, '1-6L' nem

def _is_mixed_part(tok):
    """Betű + szám vegyesen (pl. BN624A652G2, DZ4000W, FX53B)."""
    return bool(len(tok) >= 3 and re.search(r"\d", tok) and re.search(r"[A-Z]", tok)
                and not _is_measurement(tok))

def _label_tokens(text):
    """Szavakra bontás (szóköz, vessző, '+', ';', '|' mentén). Visszaad: [(megjelenő szöveg, vizsgált rész, előző szó)].
    Kötőjeles, számmal kezdődő részhez hozzáveszi az előtagját: '6ES7 214-1AE30-0XB0',
    'FN 3258H-180-40', 'BEG 0800-05', Heidenhain ID '358 701-05', Bosch '0.652 620-91'."""
    raw = [t.strip("()[]{}:;,.'\"").strip("/") for t in re.split(r"[\s,+;|]+", str(text or "").upper())]
    raw = [t for t in raw if t]
    out = []
    for i, t in enumerate(raw):
        prev = raw[i - 1] if i else ""
        prev2 = raw[i - 2] if i > 1 else ""
        joinable_prefix = (
            _SIEMENS_PREFIX.match(prev)
            or (re.fullmatch(r"[A-Z]{2,4}", prev) and prev not in _LABEL_NOT_PREFIX)
            or (re.fullmatch(r"\d{3}|\d\.\d{3}", prev) and re.fullmatch(r"\d{3}-\d{2}", t))
        )
        if out and "-" in t and t[:1].isdigit() and joinable_prefix and out[-1][0] == prev:
            out[-1] = (f"{prev} {t}", t, prev2)
        else:
            out.append((t, t, prev))
    return out

def _join_fitting(parts):
    out = []
    for p in parts:
        if p in out:
            continue
        if out and len(" ".join(out + [p])) > LABEL_SUBTEXT_MAX:
            break
        out.append(p)
    return " ".join(out)

def get_label_subtext(d_val, e_val):
    """A címke alsó sora - LEGFELJEBB LABEL_SUBTEXT_MAX (19) karakter. Minden címkegenerálás
    (ZPL, PNG, előnézetek) ezt használja. Hosszú kötőjeles cikkszámnál a 19. karakter után levágja,
    a végén maradó kötőjelet / perjelet / szóközt pedig eltávolítja (pl. 'A06B-6087-H130-XYZ-' -> 'A06B-6087-H130-XYZ')."""
    text = _label_subtext_full(d_val, e_val)
    if len(text) > LABEL_SUBTEXT_MAX:
        text = text[:LABEL_SUBTEXT_MAX].rstrip(" -/.,_+")
    return text

def _label_subtext_full(d_val, e_val):
    """A címke alsó sorába (az R18-kód alá) kerülő szöveg a D (Típus) és E (Cikkszám) mezőből.
    Sorrend: 1) kötőjeles alkatrészszámok (előbb E, aztán D), 2) ha a mező rövid és van benne
    szám: a teljes mező, 3) betű+szám vegyes részek, 4) legalább 4 jegyű számok, 5) a mező eleje."""
    fields = [re.sub(r"\s+", " ", str(v or "").strip().upper()) for v in (e_val, d_val)]
    fields = [f for f in fields if f and f not in ("NONE", "NAN", "-")]
    if not fields:
        return ""
    tokenized = [_label_tokens(f) for f in fields]

    for toks in tokenized:
        dashed = [shown for shown, core, prev in toks if _is_dashed_part(core, prev)]
        if dashed:
            return _join_fitting(dashed)
    for f in fields:
        f = re.sub(r"\s*\|\s*", " ", f).strip("+-,;|/ ")
        if re.search(r"\d", f) and len(f) <= LABEL_SUBTEXT_MAX:
            return f
    for toks in tokenized:
        mixed = [shown for shown, core, _ in toks if _is_mixed_part(core)]
        if mixed:
            return _join_fitting(mixed)
    for toks in tokenized:
        nums = [shown for shown, core, _ in toks if re.fullmatch(r"[\d.]{4,}", core)]
        if nums:
            return _join_fitting(nums)
    return fields[0][:LABEL_SUBTEXT_MAX]

def get_last_processed_info():
    try:
        conn = get_db_connection()
        cursor = conn.cursor()
        
        cursor.execute("""
            SELECT COUNT(*) FROM termekek 
            WHERE (C IS NOT NULL AND TRIM(C) != '' AND LOWER(TRIM(C)) != 'none' AND LOWER(TRIM(C)) != 'nan')
               OR (D IS NOT NULL AND TRIM(D) != '' AND LOWER(TRIM(D)) != 'none' AND LOWER(TRIM(D)) != 'nan')
        """)
        row = cursor.fetchone()
        total_count = row[0] if row else 0

        cursor.execute("""
            SELECT A, B FROM termekek 
            WHERE (C IS NOT NULL AND TRIM(C) != '' AND LOWER(TRIM(C)) != 'none' AND LOWER(TRIM(C)) != 'nan')
               OR (D IS NOT NULL AND TRIM(D) != '' AND LOWER(TRIM(D)) != 'none' AND LOWER(TRIM(D)) != 'nan')
        """)
        rows = cursor.fetchall()
        conn.close()

        highest_num = -1
        highest_code = "R18-0"

        for r in rows:
            code = str(r["A"] or r["B"] or "")
            m = re.search(r'R18-(\d+)', code, re.IGNORECASE)
            if m:
                num = int(m.group(1))
                if 0 < num < 100000:
                    if num > highest_num:
                        highest_num = num
                        highest_code = f"R18-{num}"

        return highest_code, total_count, []
    except Exception:
        return "R18-0", 0, []

# "Összesen raktáron" összesítő: korábban MINDEN kártya végigolvasta a teljes adatbázist.
# Most egyszer épül fel egy index (alkatrészszám -> összes db), és pár másodpercig újrahasznosítjuk.
_STOCK_SUM_TTL_SEC = 5.0
_stock_sum_cache = {"ts": 0.0, "by_clean": {}, "by_upper": {}}
_stock_sum_lock = threading.Lock()

def get_stock_sum(search_target):
    """Az azonos (okos) alkatrészszámú tételek összes darabszáma a raktárban."""
    with _stock_sum_lock:
        if time.monotonic() - _stock_sum_cache["ts"] > _STOCK_SUM_TTL_SEC:
            by_clean, by_upper = {}, {}
            conn = get_db_connection()
            cursor = conn.cursor()
            cursor.execute("SELECT D, E, F FROM termekek WHERE (D IS NOT NULL AND D != '') OR (E IS NOT NULL AND E != '')")
            for r in cursor.fetchall():
                r_smart = get_smart_part_number(safe_str(r[0]), safe_str(r[1]))
                f_match = re.search(r'\d+', safe_str(r[2]).strip())
                qty = int(f_match.group(0)) if f_match else 1
                up = r_smart.upper()
                by_upper[up] = by_upper.get(up, 0) + qty
                cl = clean_part_number(r_smart)
                if cl:
                    by_clean[cl] = by_clean.get(cl, 0) + qty
            conn.close()
            _stock_sum_cache.update(ts=time.monotonic(), by_clean=by_clean, by_upper=by_upper)
        clean_target = clean_part_number(search_target)
        if clean_target:
            return _stock_sum_cache["by_clean"].get(clean_target, 0)
        return _stock_sum_cache["by_upper"].get(search_target.upper(), 0)

def get_next_suggested_r18_code(last_code):
    match = re.search(r'R18-(\d+)', last_code, re.IGNORECASE)
    if match:
        next_num = int(match.group(1)) + 1
        return f"R18-{next_num}"
    return "R18-"

# -------------------------------------------------------------------
# RAKTÁRTÉRKÉP
# -------------------------------------------------------------------
def load_warehouse_map_data():
    xlsx_url = f"https://docs.google.com/spreadsheets/d/{MAP_SHEETS_ID}/export?format=xlsx"
    try:
        req = urllib.request.Request(xlsx_url, headers={'User-Agent': 'Mozilla/5.0'})
        with urllib.request.urlopen(req) as response:
            data = response.read()
        wb = openpyxl.load_workbook(io.BytesIO(data), data_only=True)
        ws = wb.active
        map_data = []
        for row in ws.iter_rows():
            row_data = []
            for cell in row:
                val = cell.value if cell.value is not None else ""
                val_str = str(val).strip()
                if val_str.endswith(".0"): val_str = val_str[:-2]
                bg_color = "transparent"
                if cell.fill and getattr(cell.fill, 'start_color', None):
                    color_rgb = getattr(cell.fill.start_color, 'rgb', None)
                    if isinstance(color_rgb, str) and len(color_rgb) >= 6 and color_rgb != "00000000":
                        bg_color = "#" + color_rgb[-6:]
                row_data.append({"value": val_str, "color": bg_color})
            map_data.append(row_data)
        return map_data
    except Exception as e:
        print(f"Hiba a térkép XLSX letöltésekor: {e}")
        return None

def match_location(cell_val, target_val):
    if not cell_val or not target_val: return False
    c_raw = str(cell_val).strip().upper()
    t_raw = str(target_val).strip().upper()
    if not c_raw or not t_raw: return False

    c_clean = re.sub(r'[^\d]', '', c_raw)
    t_clean = re.sub(r'[^\d]', '', t_raw)

    if t_clean and c_clean == t_clean: return True
    return c_raw == t_raw

def _norm_map_text(val):
    """Pontos egyezéshez: kis/nagybetű, felesleges szóközök és a '112.0' alak nem számít."""
    t = re.sub(r"\s+", " ", str(val or "")).strip().upper()
    if re.fullmatch(r"\d+\.0", t):
        t = t[:-2]
    return t

def match_location_exact(cell_val, target_val):
    """Raktártérkép kereső: CSAK a pontosan egyező cella (a '112' nem találja meg a '112B'-t vagy az 'A112'-t)."""
    t = _norm_map_text(target_val)
    return bool(t) and _norm_map_text(cell_val) == t

def render_interactive_map_html(map_data, target_location, exact=False):
    if not map_data:
        return "<html><body style='background:#121212; color:red;'><h2>⚠️ A térkép letöltése sikertelen!</h2></body></html>"

    html_code = """
    <!DOCTYPE html>
    <html>
    <head>
    <style>
        html, body { margin: 0; padding: 2px; background-color: #121212; color: white; font-family: Arial, sans-serif; overflow: auto; }
        .map-table { border-collapse: collapse; font-size: 11px; text-align: center; margin: 0 auto; background-color: white; }
        .map-cell { border: 1px solid #444; padding: 3px 5px; min-width: 28px; height: 20px; color: black; font-weight: bold; }
        
        .highlighted-cell { 
            background-color: #ff0055 !important; 
            color: white !important; 
            font-size: 15px !important; 
            border: 2px solid #ffffff !important; 
            box-shadow: 0 0 15px #ff0055; 
            animation: pulse 0.8s infinite alternate; 
        }
        @keyframes pulse { 
            0% { transform: scale(1); box-shadow: 0 0 8px #ff0055; } 
            100% { transform: scale(1.15); box-shadow: 0 0 20px #ff0055; } 
        }
    </style>
    </head>
    <body>
    <table class="map-table">
    """
    match_found = False
    for row in map_data:
        html_code += "<tr>"
        for cell in row:
            val = cell["value"]
            bg_color = cell["color"]
            is_match = match_location_exact(val, target_location) if exact else match_location(val, target_location)
            if is_match:
                match_found = True
                cell_class = "map-cell highlighted-cell"
                inline_style = "" 
            else:
                cell_class = "map-cell"
                inline_style = f"background-color: {bg_color};" if bg_color != "transparent" else ""
            html_code += f'<td class="{cell_class}" style="{inline_style}">{val}</td>'
        html_code += "</tr>"
    html_code += "</table>"

    if match_found and exact:
        # Több találatnál: az elsőhöz ugrik; nextMatch() a következőhöz lép (újabb keresés ugyanarra)
        html_code += """
        <style>
            .current-match { outline: 4px solid #ffeb3b !important; outline-offset: 2px; z-index: 10; position: relative; }
        </style>
        <script>
            var matches = [], idx = -1;
            function showMatch(i) {
                matches = document.querySelectorAll('.highlighted-cell');
                if (!matches.length) return 0;
                idx = ((i % matches.length) + matches.length) % matches.length;
                matches.forEach(function (m) { m.classList.remove('current-match'); });
                var t = matches[idx];
                t.classList.add('current-match');
                t.scrollIntoView({ behavior: 'smooth', block: 'center', inline: 'center' });
                return idx + 1;
            }
            function nextMatch() { return showMatch(idx + 1); }
            window.addEventListener('load', function () { showMatch(0); });
            setTimeout(function () { if (idx < 0) showMatch(0); }, 150);
        </script>
        """
    elif match_found:
        html_code += """
        <script>
            function forceScroll() {
                var target = document.querySelector('.highlighted-cell');
                if (target) {
                    target.scrollIntoView({ behavior: 'auto', block: 'center', inline: 'center' });
                }
            }
            window.addEventListener('load', forceScroll);
            setTimeout(forceScroll, 150);
        </script>
        """
    html_code += "</body></html>"
    return html_code

class MapWorker(QThread):
    finished_signal = pyqtSignal(list)
    error_signal = pyqtSignal(str)

    def run(self):
        try:
            map_data = load_warehouse_map_data()
            if map_data: self.finished_signal.emit(map_data)
            else: self.error_signal.emit("Nem sikerült letölteni a térkép adatait.")
        except Exception as e:
            self.error_signal.emit(f"Hiba a térkép betöltésekor: {e}")

class MapDialog(QDialog):
    def __init__(self, target_loc, parent=None):
        super().__init__(parent)
        self.setWindowTitle(f"🗺️ Raktártérkép Pozíció: [{target_loc}]")
        self.resize(1100, 650)
        self.setStyleSheet("QDialog { background-color: #121212; }")
        self.target_loc = target_loc
        
        layout = QVBoxLayout(self)
        layout.setContentsMargins(8, 8, 8, 8)
        layout.setSpacing(4)

        self.lbl_status = QLabel(f"📍 Keresett pozíció kiemelése: <b>{target_loc}</b>")
        self.lbl_status.setStyleSheet("color: #29b6f6; font-size: 15px; font-weight: bold; margin-bottom: 2px;")
        layout.addWidget(self.lbl_status)

        self.web = QWebEngineView()
        self.web.setPage(SilentWebEnginePage(self.web))
        layout.addWidget(self.web, stretch=1)

        self.start_async_map_load()

    def start_async_map_load(self):
        self.map_worker = MapWorker()
        self.map_worker.finished_signal.connect(self.on_map_loaded)
        self.map_worker.error_signal.connect(self.on_map_error)
        self.map_worker.start()

    def on_map_loaded(self, map_data):
        self.lbl_status.setText(f"📍 Keresett pozíció kiemelése: <b>{self.target_loc}</b>")
        self.lbl_status.setStyleSheet("color: #29b6f6; font-size: 15px; font-weight: bold; margin-bottom: 2px;")
        self.web.setHtml(render_interactive_map_html(map_data, self.target_loc))

    def on_map_error(self, err_msg):
        self.lbl_status.setText(f"❌ {err_msg}")
        self.lbl_status.setStyleSheet("color: #ff1744; font-size: 15px; font-weight: bold; margin-bottom: 2px;")

class ToastNotification(QFrame):
    def __init__(self, message, parent=None):
        super().__init__(parent)
        self.setStyleSheet("""
            QFrame {
                background-color: #00e676;
                color: #000000;
                border-radius: 8px;
                padding: 10px 16px;
                font-weight: bold;
                font-size: 15px;
                border: 1px solid #ffffff;
            }
        """)
        layout = QHBoxLayout(self)
        layout.setContentsMargins(10, 8, 10, 8)
        lbl = QLabel(message)
        lbl.setStyleSheet("color: black; font-weight: bold;")
        layout.addWidget(lbl)
        self.opacity_effect = QGraphicsOpacityEffect(self)
        self.setGraphicsEffect(self.opacity_effect)

    def show_toast(self, duration=2500):
        self.show()
        self.raise_()
        if self.parent():
            parent_rect = self.parent().rect()
            self.adjustSize()
            x = parent_rect.width() - self.width() - 25
            y = 70
            self.move(x, y)
        QTimer.singleShot(duration, self.fade_out)

    def fade_out(self):
        self.anim = QPropertyAnimation(self.opacity_effect, b"opacity")
        self.anim.setDuration(400)
        self.anim.setStartValue(1.0)
        self.anim.setEndValue(0.0)
        self.anim.finished.connect(self.close)
        self.anim.start()

class SilentConfirmDialog(QDialog):
    def __init__(self, title, message, parent=None):
        super().__init__(parent)
        self.setWindowTitle(title)
        self.setFixedWidth(420)
        self.setStyleSheet("QDialog { background-color: #1E1E1E; color: white; border: 2px solid #ff1744; border-radius: 8px; } QLabel { font-size: 15px; color: white; font-weight: bold; padding: 12px; } QPushButton { font-weight: bold; padding: 8px; border-radius: 4px; font-size: 14px; }")
        layout = QVBoxLayout(self)
        lbl = QLabel(message)
        lbl.setWordWrap(True)
        lbl.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(lbl)
        btn_layout = QHBoxLayout()
        yes_btn = QPushButton("Igen")
        yes_btn.setStyleSheet("background-color: #ff1744; color: white;")
        yes_btn.clicked.connect(self.accept)
        no_btn = QPushButton("Mégse")
        no_btn.setStyleSheet("background-color: #444; color: white;")
        no_btn.clicked.connect(self.reject)
        btn_layout.addWidget(yes_btn)
        btn_layout.addWidget(no_btn)
        layout.addLayout(btn_layout)


class RowPrintDialog(QDialog):
    def __init__(self, code, marka="", tipus="", cikkszam="", parent=None):
        super().__init__(parent)
        self.code = code
        self.marka = marka
        self.tipus = tipus
        self.cikkszam = cikkszam

        # Okos alkatrészszám választó algoritmus lefuttatása
        self.smart_part_number = get_label_subtext(self.tipus, self.cikkszam)

        self.setWindowTitle(f"🖨️ Címke Nyomtatása: [{self.code}]")
        self.setFixedWidth(380)
        self.setStyleSheet("QDialog { background-color: #121212; color: white; } QLabel { color: white; font-weight: bold; }")

        layout = QVBoxLayout(self)
        layout.setSpacing(10)

        # Információs kártya kiemelve az okos alkatrészszámmal
        lbl_info = QLabel(
            f"🏷️ <b>Kód:</b> {self.code}<br>"
            f"<b>Márka:</b> {self.marka or '-'}<br>"
            f"<b>Típus:</b> {self.tipus or '-'}<br>"
            f"<b>Cikkszám:</b> {self.cikkszam or '-'}<br>"
            f"<b>🎯 Címkére kerülő szöveg:</b> <span style='color: #00e676;'>{self.smart_part_number or '-'}</span>"
        )
        lbl_info.setStyleSheet("background: #1e1e1e; padding: 10px; border: 1px solid #333; border-radius: 6px; color: #ffffff; font-size: 13px;")
        layout.addWidget(lbl_info)

        # Címke előnézeti képek generálása a smart szöveggel
        codes_box = QWidget()
        codes_layout = QHBoxLayout(codes_box)
        codes_layout.setContentsMargins(0, 0, 0, 0)

        qr_bytes, bar_bytes = generate_codes_in_memory(self.code, cikkszam_text=self.cikkszam, tipus_text=self.tipus)
        
        if qr_bytes and bar_bytes:
            lbl_qr = QLabel()
            pix_qr = QPixmap()
            pix_qr.loadFromData(qr_bytes)
            lbl_qr.setPixmap(pix_qr.scaled(100, 100, Qt.AspectRatioMode.KeepAspectRatio))

            lbl_bar = QLabel()
            pix_bar = QPixmap()
            pix_bar.loadFromData(bar_bytes)
            lbl_bar.setPixmap(pix_bar.scaled(190, 85, Qt.AspectRatioMode.KeepAspectRatio))

            codes_layout.addWidget(lbl_qr)
            codes_layout.addWidget(lbl_bar)
            layout.addWidget(codes_box)

        # Nyomtatási Mód (ZPL / PNG)
        mode_box = QGroupBox("Nyomtatási Mód")
        mode_box.setStyleSheet("QGroupBox { color: #29b6f6; font-weight: bold; border: 1px solid #333; border-radius: 6px; margin-top: 6px; padding-top: 10px; }")
        mode_lay = QHBoxLayout(mode_box)
        self.r_zpl = QRadioButton("ZPL Nyomtatás")
        self.r_png = QRadioButton("PNG Nyomtatás")
        self.r_zpl.setChecked(True)
        mode_lay.addWidget(self.r_zpl)
        mode_lay.addWidget(self.r_png)
        layout.addWidget(mode_box)

        # Példányszám
        spin_box = QHBoxLayout()
        spin_box.addWidget(QLabel("Példányszám:"))
        self.spin_copies = QSpinBox()
        self.spin_copies.setMinimum(1)
        self.spin_copies.setValue(1)
        self.spin_copies.setStyleSheet("background-color: #2b2b2b; color: white; padding: 4px; font-weight: bold;")
        spin_box.addWidget(self.spin_copies)
        layout.addLayout(spin_box)

        # Gombok
        btn_lay = QHBoxLayout()
        
        btn_print = QPushButton("🖨️ Nyomtatás")
        btn_print.setStyleSheet("background-color: #29b6f6; color: black; font-weight: bold; padding: 8px; font-size: 14px; border-radius: 4px;")
        btn_print.clicked.connect(self.execute_print)
        btn_lay.addWidget(btn_print, stretch=2)

        btn_cancel = QPushButton("Mégse")
        btn_cancel.setStyleSheet("background-color: #444; color: white; padding: 8px; border-radius: 4px;")
        btn_cancel.clicked.connect(self.reject)
        btn_lay.addWidget(btn_cancel, stretch=1)

        layout.addLayout(btn_lay)

    def execute_print(self):
        copies = self.spin_copies.value()
        is_zpl = self.r_zpl.isChecked()
        zebra_printer = get_zebra_printer_name()

        if is_zpl:
            zpl_data = generate_zpl_code(self.code, cikkszam=self.cikkszam, tipus=self.tipus, copies=copies)
            ok, msg = send_zpl_to_printer(zebra_printer, zpl_data)
            if ok and hasattr(self.parent(), 'trigger_toast'):
                self.parent().trigger_toast(f"🖨️ ZPL Címke kiküldve [{self.code}]!")
            elif not ok and hasattr(self.parent(), 'trigger_toast'):
                self.parent().trigger_toast(f"❌ Nyomtatási hiba: {msg}")
        else:
            _, bar_bytes = generate_codes_in_memory(self.code, cikkszam_text=self.cikkszam, tipus_text=self.tipus)
            if bar_bytes:
                ok, msg = send_png_bytes_to_printer(zebra_printer, bar_bytes, copies=copies)
                if ok and hasattr(self.parent(), 'trigger_toast'):
                    self.parent().trigger_toast(f"🖨️ PNG Címke kinyomtatva [{self.code}]!")
                elif not ok and hasattr(self.parent(), 'trigger_toast'):
                    self.parent().trigger_toast(f"❌ Nyomtatási hiba: {msg}")

        self.accept()



class CopiesDialog(QDialog):
    """Gyors példányszám-kérő ablak: alapból 1, Enter = nyomtatás, Esc = mégse.
    A szám ki van jelölve, így azonnal felülírható (pl. 5 + Enter)."""
    def __init__(self, code, parent=None):
        super().__init__(parent)
        self.setWindowTitle(f"🖨️ Nyomtatás: {code}")
        self.setFixedWidth(300)
        self.setStyleSheet("QDialog { background-color: #121212; color: white; } QLabel { color: white; font-weight: bold; }")

        layout = QVBoxLayout(self)
        layout.addWidget(QLabel(f"🏷️ {code}\nHány darab címke legyen?"))

        self.spin_copies = QSpinBox()
        self.spin_copies.setRange(1, 999)
        self.spin_copies.setValue(1)
        self.spin_copies.setStyleSheet("background-color: #2b2b2b; color: white; padding: 6px; font-weight: bold; font-size: 18px;")
        layout.addWidget(self.spin_copies)

        btn_lay = QHBoxLayout()
        btn_print = QPushButton("🖨️ Nyomtatás")
        btn_print.setDefault(True)
        btn_print.setStyleSheet("background-color: #29b6f6; color: black; font-weight: bold; padding: 8px; font-size: 14px; border-radius: 4px;")
        btn_print.clicked.connect(self.accept)
        btn_lay.addWidget(btn_print, stretch=2)

        btn_cancel = QPushButton("Mégse")
        btn_cancel.setStyleSheet("background-color: #444; color: white; padding: 8px; border-radius: 4px;")
        btn_cancel.clicked.connect(self.reject)
        btn_lay.addWidget(btn_cancel, stretch=1)
        layout.addLayout(btn_lay)

        self.spin_copies.setFocus()
        self.spin_copies.selectAll()

    def copies(self):
        return self.spin_copies.value()


class SaleDialog(QDialog):
    """Eladás rögzítése: a mezők egymás után jönnek (Enter = következő mező).
    Vevő / Eladási ár / Tracking number: '⏭ Később' gombbal kihagyható, az Eladottak fülön pótolható.
    A dátum automatikus."""
    def __init__(self, platform, product, stock_qty, parent=None):
        super().__init__(parent)
        self.platform = platform
        self.product = product
        cim = SALES_SHEETS[platform]["cim"]
        self.datum = datetime.datetime.now().strftime("%Y-%m-%d %H:%M")
        self.setWindowTitle(f"💰 {cim}: {product.get('A', '')}")
        self.setMinimumWidth(520)
        self.setStyleSheet("""
            QDialog { background-color: #121212; color: white; }
            QLabel { color: white; font-weight: bold; }
            QLineEdit { background-color: #2b2b2b; color: white; padding: 6px; font-size: 15px; border: 1px solid #444; border-radius: 4px; }
            QLineEdit:focus { border: 2px solid #00e676; }
        """)

        layout = QVBoxLayout(self)
        info = QLabel(
            f"🏷️ <b>{product.get('A', '')}</b> &nbsp;|&nbsp; {product.get('C', '') or '-'} &nbsp;|&nbsp; "
            f"{product.get('D', '') or '-'} &nbsp;|&nbsp; {product.get('E', '') or '-'}<br>"
            f"📦 Raktáron: <b>{stock_qty} db</b> &nbsp;|&nbsp; 📅 Eladás dátuma (automatikus): <b>{self.datum}</b>"
        )
        info.setStyleSheet("background: #1e1e1e; padding: 10px; border: 1px solid #333; border-radius: 6px; font-size: 13px;")
        layout.addWidget(info)

        form = QFormLayout()
        self.inp_vevo = QLineEdit()
        self.inp_vevo.setPlaceholderText("Vevő neve / felhasználóneve")
        self.inp_ar = QLineEdit()
        self.inp_ar.setPlaceholderText("pl. 120 EUR")
        self.inp_darab = QLineEdit("1")
        self.inp_darab.setValidator(QIntValidator(1, 100000, self))
        self.inp_tracking = QLineEdit()
        self.inp_tracking.setPlaceholderText("Tracking number")
        attach_scanner_fix(self.inp_tracking)
        self.inp_megj = QLineEdit()
        self.inp_megj.setPlaceholderText("(nem kötelező)")

        form.addRow("👤 Vevő:", self._with_later(self.inp_vevo))
        form.addRow("💶 Eladási ár:", self._with_later(self.inp_ar))
        form.addRow("🔢 Eladott darab:", self.inp_darab)
        form.addRow("🚚 Tracking number:", self._with_later(self.inp_tracking))
        form.addRow("📝 Megjegyzés:", self.inp_megj)
        layout.addLayout(form)

        self.chk_stock = QCheckBox("Eladott darabszám levonása a raktárkészletből (0 db esetén a termék törlődik és a ToroltSor lapra kerül)")
        self.chk_stock.setChecked(True)
        self.chk_stock.setStyleSheet("color: #ffb74d; font-weight: bold;")
        layout.addWidget(self.chk_stock)

        btn_lay = QHBoxLayout()
        self.btn_ok = QPushButton(f"✅ Eladás rögzítése ({cim})")
        self.btn_ok.setStyleSheet("background-color: #00e676; color: black; font-weight: bold; padding: 10px; font-size: 15px; border-radius: 4px;")
        self.btn_ok.clicked.connect(self._try_accept)
        btn_cancel = QPushButton("Mégse")
        btn_cancel.setStyleSheet("background-color: #444; color: white; padding: 10px; border-radius: 4px;")
        btn_cancel.clicked.connect(self.reject)
        for b in (self.btn_ok, btn_cancel):
            b.setAutoDefault(False)
            b.setDefault(False)
        btn_lay.addWidget(self.btn_ok, stretch=3)
        btn_lay.addWidget(btn_cancel, stretch=1)
        layout.addLayout(btn_lay)

        # Enter = következő mező, az utolsó után rögzítés
        order = [self.inp_vevo, self.inp_ar, self.inp_darab, self.inp_tracking, self.inp_megj]
        for cur, nxt in zip(order, order[1:]):
            cur.returnPressed.connect(lambda n=nxt: (n.setFocus(), n.selectAll()))
        self.inp_megj.returnPressed.connect(self._try_accept)
        self.inp_vevo.setFocus()

    def _with_later(self, line_edit):
        box = QWidget()
        lay = QHBoxLayout(box)
        lay.setContentsMargins(0, 0, 0, 0)
        lay.addWidget(line_edit, stretch=1)
        btn = QPushButton("⏭ Később")
        btn.setAutoDefault(False)
        btn.setToolTip("Most kihagyja - az Eladottak fülön később kitölthető")
        btn.setStyleSheet("background-color: #ff9800; color: black; font-weight: bold; padding: 5px 10px; border-radius: 4px;")

        def later():
            line_edit.clear()
            line_edit.setPlaceholderText("⏭ később kitöltendő (Eladottak fül)")
            line_edit.focusNextChild()
        btn.clicked.connect(later)
        lay.addWidget(btn)
        return box

    def _try_accept(self):
        if not self.inp_darab.text().strip().isdigit() or int(self.inp_darab.text()) < 1:
            self.inp_darab.setFocus()
            self.inp_darab.selectAll()
            return
        self.accept()

    def values(self):
        return {
            "vevo": self.inp_vevo.text().strip(),
            "ar": self.inp_ar.text().strip(),
            "darab": int(self.inp_darab.text().strip() or "1"),
            "tracking": self.inp_tracking.text().strip(),
            "megjegyzes": self.inp_megj.text().strip(),
            "reduce_stock": self.chk_stock.isChecked(),
            "datum": self.datum,
        }


class ConflictDialog(QDialog):
    """R18-ütközés: a szerveren ugyanazzal az R18-számmal MÁSIK termék van. Egymás mellett mutatja a
    RÉGI (szerveren lévő) és az ÚJ adatot; a bejelölt sorokat felülírja az újjal."""
    def __init__(self, conflicts, parent=None):
        super().__init__(parent)
        self.conflicts = conflicts
        self.setWindowTitle("⚠️ R18-szám ütközés - felülírod?")
        self.resize(1000, 420)
        self.setStyleSheet("QDialog { background-color: #121212; color: white; } QLabel { color: white; }")
        lay = QVBoxLayout(self)
        info = QLabel(
            "<b>Ezekkel az R18-számokkal a szerveren már MÁSIK termék van, ezért NEM lettek elmentve.</b><br>"
            "Pipáld be, amelyiket <b>felül szeretnéd írni az új adattal</b>. A régi adat a 📜 Naplóban megmarad "
            "(ELŐTTE / UTÁNA). Ami nincs bepipálva, az nem mentődik - azt egy új R18-szám alatt vedd fel.")
        info.setWordWrap(True)
        lay.addWidget(info)

        self.table = QTableWidget(len(conflicts), 4)
        self.table.setHorizontalHeaderLabels(["Felülírás?", "R18-kód", "⬅️ RÉGI (most a szerveren)", "➡️ ÚJ (amit beírtál)"])
        self.table.verticalHeader().setVisible(False)
        self.table.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        self.table.setStyleSheet("QTableWidget { background-color: #21252b; color: #e0e0e0; gridline-color: #3b4048; }"
                                 "QHeaderView::section { background-color: #1b1e23; color: #61afef; font-weight: bold; padding: 5px; }")
        self.checks = []
        for i, c in enumerate(conflicts):
            chk = QCheckBox()
            chk.setChecked(True)
            chk.setStyleSheet("margin-left: 18px;")
            self.checks.append(chk)
            self.table.setCellWidget(i, 0, chk)
            self.table.setItem(i, 1, QTableWidgetItem(str(c.get("A", ""))))
            old = QTableWidgetItem(str(c.get("_server_existing", "") or "(ismeretlen)"))
            old.setBackground(QColor("#7f1d1d"))
            self.table.setItem(i, 2, old)
            new = QTableWidgetItem(" | ".join(str(c.get(k, "")) for k in ("C", "D", "E") if str(c.get(k, "")).strip())
                                   + "   (" + format_row_for_log(c) + ")")
            new.setBackground(QColor("#1b5e20"))
            new.setToolTip(format_row_for_log(c))
            self.table.setItem(i, 3, new)
        self.table.setColumnWidth(0, 80)
        self.table.setColumnWidth(1, 100)
        self.table.setColumnWidth(2, 260)
        self.table.horizontalHeader().setStretchLastSection(True)
        lay.addWidget(self.table)

        btns = QHBoxLayout()
        ok = QPushButton("✅ Bepipáltak felülírása az ÚJ adattal")
        ok.setStyleSheet("background-color: #00e676; color: black; font-weight: bold; padding: 10px;")
        ok.clicked.connect(self.accept)
        cancel = QPushButton("Most egyiket sem")
        cancel.setStyleSheet("background-color: #444; color: white; padding: 10px;")
        cancel.clicked.connect(self.reject)
        btns.addWidget(ok, stretch=2)
        btns.addWidget(cancel, stretch=1)
        lay.addLayout(btns)

    def selected(self):
        return [c for c, chk in zip(self.conflicts, self.checks) if chk.isChecked()]


class PhotoDialog(QDialog):
    """📷 Termékfotók nagyban, lapozással (← → / gombok)."""
    def __init__(self, code, paths, start=0, parent=None):
        super().__init__(parent)
        self.code, self.paths, self.i = code, paths, start
        self.setWindowTitle(f"📷 {code} - fotók")
        self.resize(1000, 760)
        self.setStyleSheet("QDialog { background-color: #121212; color: white; } QLabel { color: white; }")
        lay = QVBoxLayout(self)
        self.info = QLabel("")
        self.info.setStyleSheet("color: #29b6f6; font-weight: bold; font-size: 14px;")
        lay.addWidget(self.info)
        self.img = QLabel("⏳")
        self.img.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.img.setMinimumSize(400, 300)
        lay.addWidget(self.img, stretch=1)
        bar = QHBoxLayout()
        b_prev = QPushButton("◀ Előző"); b_prev.clicked.connect(lambda: self.show_i(self.i - 1)); bar.addWidget(b_prev)
        b_next = QPushButton("Következő ▶"); b_next.clicked.connect(lambda: self.show_i(self.i + 1)); bar.addWidget(b_next)
        bar.addStretch()
        b_file = QPushButton("🖼️ Megnyitás képnézegetőben")
        b_file.clicked.connect(lambda: QDesktopServices.openUrl(QUrl.fromLocalFile(self.paths[self.i])))
        bar.addWidget(b_file)
        b_dir = QPushButton("📂 Mappa megnyitása")
        b_dir.clicked.connect(lambda: QDesktopServices.openUrl(QUrl.fromLocalFile(os.path.dirname(self.paths[self.i]))))
        bar.addWidget(b_dir)
        lay.addLayout(bar)
        self._req = None
        photo_bus.thumb_ready.connect(self._on_img)
        self.show_i(start)

    def show_i(self, i):
        if not self.paths:
            return
        self.i = i % len(self.paths)
        self.info.setText(f"{self.code}  |  {self.i + 1} / {len(self.paths)}  |  {os.path.basename(self.paths[self.i])}")
        cloud = is_cloud_only(self.paths[self.i])
        self.img.setText("☁️ Letöltés a NAS-ról... (a program közben használható)" if cloud else "⏳ Betöltés...")
        self._req = request_thumbnail(self.paths[self.i], 1600, 1200, allow_download=True)

    def _on_img(self, rid, img):
        if rid != self._req:
            return
        try:
            if not isinstance(img, QImage) or img.isNull():
                self.img.setText("❌ A kép nem olvasható (lehet, hogy a Synology Drive még nem töltötte le).")
                return
            pm = QPixmap.fromImage(img).scaled(self.img.width(), self.img.height(), Qt.AspectRatioMode.KeepAspectRatio,
                                               Qt.TransformationMode.SmoothTransformation)
            self.img.setPixmap(pm)
        except RuntimeError:
            pass

    def keyPressEvent(self, e):
        if e.key() == Qt.Key.Key_Left:
            self.show_i(self.i - 1)
        elif e.key() == Qt.Key.Key_Right:
            self.show_i(self.i + 1)
        else:
            super().keyPressEvent(e)

    def done(self, r):
        try:
            photo_bus.thumb_ready.disconnect(self._on_img)
        except Exception:
            pass
        super().done(r)


class RestoreDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("↩️ Adatbázis Visszaállítása Pillanatképből")
        self.setMinimumSize(500, 350)
        self.setStyleSheet("""
            QDialog { background-color: #121212; color: white; border: 1px solid #29b6f6; border-radius: 8px; }
            QLabel { font-size: 15px; font-weight: bold; color: #29b6f6; margin-bottom: 8px; }
            QListWidget { background-color: #1E1E1E; color: white; border: 1px solid #444; border-radius: 6px; font-size: 14px; }
            QPushButton { background-color: #29b6f6; color: black; font-weight: bold; padding: 8px; border-radius: 4px; }
        """)

        layout = QVBoxLayout(self)
        layout.addWidget(QLabel("Válaszd ki a visszaállítani kívánt biztonsági pillanatképet:"))

        self.file_list = QListWidget()
        self.load_backup_files()
        layout.addWidget(self.file_list)

        btn_layout = QHBoxLayout()
        
        restore_btn = QPushButton("↩️ Kiválasztott Állapot Visszaállítása")
        restore_btn.setStyleSheet("background-color: #ff9800; color: black;")
        
        restore_btn.clicked.connect(
            lambda: self.run_button_action(
                restore_btn, 
                "↩️ Kiválasztott Állapot Visszaállítása", 
                self.do_restore
            )
        )
        
        cancel_btn = QPushButton("Mégse")
        cancel_btn.setStyleSheet("background-color: #444; color: white;")
        cancel_btn.clicked.connect(self.reject)

        btn_layout.addWidget(restore_btn)
        btn_layout.addWidget(cancel_btn)
        layout.addLayout(btn_layout)
        self.selected_file = None

    def load_backup_files(self):
        self.file_list.clear()
        if os.path.exists("backups"):
            files = [f for f in os.listdir("backups") if f.endswith(".db")]
            files.sort(reverse=True)
            for f in files:
                self.file_list.addItem(f)

    def do_restore(self):
        curr_item = self.file_list.currentItem()
        if not curr_item: return
        filename = curr_item.text()
        self.selected_file = os.path.join("backups", filename)
        
        confirm_dlg = SilentConfirmDialog("Visszaállítás Megerősítése", f"Biztosan visszaállítod az adatbázist?\n\n[{filename}]", self)
        if confirm_dlg.exec() == QDialog.DialogCode.Accepted:
            try:
                shutil.copy2(self.selected_file, DB_FILE)
                self.accept()
            except Exception as e:
                print(f"Hiba: {e}")

# -------------------------------------------------------------------
# NYOMTATÁS SEGÉDFÜGGVÉNYEK
# -------------------------------------------------------------------
# GYORSÍTÓTÁR: ugyanarra a kódra/szövegre nem rajzolja újra a QR + vonalkód képet (kártyák, előnézetek)
@functools.lru_cache(maxsize=1024)
def generate_codes_in_memory(code_text, cikkszam_text="", tipus_text=""):
    if not code_text or code_text in ("ISMERETLEN", "None", "nan", ""):
        return None, None

    qr_engine = qrcode.QRCode(version=1, box_size=8, border=2)
    qr_engine.add_data(code_text)
    qr_engine.make(fit=True)
    img_qr = qr_engine.make_image(fill_color="black", back_color="white")
    qr_buf = io.BytesIO()
    img_qr.save(qr_buf, format="PNG")

    bar_bytes = None
    try:
        code128 = barcode.get('code128', code_text, writer=ImageWriter())
        raw_barcode = code128.render(writer_options={'module_height': 5.0, 'module_width': 0.20, 'quiet_zone': 0.0, 'write_text': False, 'dpi': 300})
        bw, bh = raw_barcode.size
        canvas_width = 450
        canvas_height = bh + 120
        label_img = Image.new("RGB", (canvas_width, canvas_height), "white")
        label_img.paste(raw_barcode, (max(0, (canvas_width - bw) // 2), 12))
        draw = ImageDraw.Draw(label_img)
        try:
            font_r18 = ImageFont.truetype("arialbd.ttf", 50)
            font_cikk = ImageFont.truetype("arialbd.ttf", 28)
        except Exception:
            font_r18 = font_cikk = ImageFont.load_default()

        text_center_x = int(canvas_width * 0.49)
        draw.text((text_center_x, bh + 34), code_text, fill="black", font=font_r18, anchor="mm")
        
        label_subtext = get_label_subtext(tipus_text, cikkszam_text)
        if label_subtext and label_subtext not in (code_text, "None", "nan", "-"):
            # Hosszú szövegnél kisebb betű, hogy ne lógjon le a címkéről
            size = 28
            while size > 14 and draw.textlength(label_subtext, font=font_cikk) > canvas_width - 16:
                size -= 1
                try:
                    font_cikk = ImageFont.truetype("arialbd.ttf", size)
                except Exception:
                    break
            draw.text((text_center_x, bh + 82), label_subtext, fill="black", font=font_cikk, anchor="mm")
            
        bar_buf = io.BytesIO()
        label_img.save(bar_buf, format="PNG", dpi=(300, 300))
        bar_bytes = bar_buf.getvalue()
    except Exception:
        pass

    return qr_buf.getvalue(), bar_bytes

def generate_zpl_code(r18_code, cikkszam="", tipus="", copies=1):
    # '^' és '~' ZPL vezérlőkarakter - a szövegben nem maradhat
    clean_subtext = re.sub(r"[\^~]", " ", get_label_subtext(tipus, cikkszam))
    # Hosszú szövegnél kisebb betű (A0 font: legfeljebb kb. 0,66 × magasság egy karakter), max 32
    font_h = max(22, min(32, int(600 / (max(len(clean_subtext), 1) * 0.66))))
    return f"""^XA
^CI28
^PW600
^LL250
^LH0,0
^FO141,20^BY2,2.7,80^B3N,N,80,N,N^FD{r18_code}^FS
^FO7,115^A0N,50,50^FB600,1,0,C^FD{r18_code}^FS
^FO0,168^A0N,{font_h},{font_h}^FB600,1,0,C^FD{clean_subtext}^FS
^PQ{copies}
^XZ"""

def get_zebra_printer_name():
    if not WIN32_AVAILABLE:
        return None
    try:
        printers = [p[2] for p in win32print.EnumPrinters(win32print.PRINTER_ENUM_LOCAL | win32print.PRINTER_ENUM_CONNECTIONS)]
        for p in printers:
            if "ZEBRA" in p.upper() or "GK420" in p.upper():
                return p
        return win32print.GetDefaultPrinter()
    except Exception:
        return None

def send_zpl_to_printer(printer_name, zpl_data):
    """Címke nyomtatása: ha ehhez a géphez van kötve címkenyomtató, itt; különben a nyomtatós gépre küldi (VPS)."""
    if is_print_host():
        return _send_zpl_local(printer_name or local_label_printer(), zpl_data)
    return _send_zpl_remote(zpl_data)


def _send_zpl_local(printer_name, zpl_data):
    if not WIN32_AVAILABLE: 
        return False, "A 'pywin32' nincs telepítve!"
    
    target_printer = printer_name or get_zebra_printer_name()
    if not target_printer:
        return False, "Nem található elérhető nyomtató!"

    try:
        h_printer = win32print.OpenPrinter(target_printer)
        try:
            win32print.StartDocPrinter(h_printer, 1, ("Raktar_ZPL", None, "RAW"))
            win32print.StartPagePrinter(h_printer)
            win32print.WritePrinter(h_printer, zpl_data.encode('utf-8'))
            win32print.EndPagePrinter(h_printer)
            win32print.EndDocPrinter(h_printer)
            return True, f"ZPL kiküldve [{target_printer}]"
        finally:
            win32print.ClosePrinter(h_printer)
    except Exception as e:
        return False, str(e)

def send_png_bytes_to_printer(printer_name, bar_bytes, copies=1):
    if not WIN32_AVAILABLE:
        return False, "A 'pywin32' modul nincs telepítve!"
    
    target_printer = printer_name or get_zebra_printer_name()
    if not target_printer:
        return False, "Nem található elérhető nyomtató!"

    try:
        h_printer = win32print.OpenPrinter(target_printer)
        try:
            img = Image.open(io.BytesIO(bar_bytes))
            h_dc = win32ui.CreateDC()
            h_dc.CreatePrinterDC(target_printer)

            pw = h_dc.GetDeviceCaps(8)
            ph = h_dc.GetDeviceCaps(10)

            img_w, img_h = img.size
            scale = min(pw / img_w, ph / img_h)
            new_w = int(img_w * scale)
            new_h = int(img_h * scale)

            for _ in range(copies):
                h_dc.StartDoc("Raktar_PNG_Cimke")
                h_dc.StartPage()
                dib = ImageWin.Dib(img)
                dib.draw(h_dc.GetHandleOutput(), (0, 0, new_w, new_h))
                h_dc.EndPage()
                h_dc.EndDoc()

            return True, "PNG Címke kinyomtatva!"
        finally:
            win32print.ClosePrinter(h_printer)
    except Exception as e:
        return False, f"PNG Nyomtatási hiba: {str(e)}"

# -------------------------------------------------------------------
# VALÓDI HÁTTÉRSZÁLAK
# -------------------------------------------------------------------
class VPSCheckWorker(QThread):
    status_signal = pyqtSignal(bool, str)

    def run(self):
        try:
            res = vps_request("GET", f"{VPS_URL}/sync/status", timeout=3)
            if res.status_code == 200:
                self.status_signal.emit(True, "🟢 VPS Csatlakozva")
            elif res.status_code == 401:
                self.status_signal.emit(False, "🔑 VPS: hibás / hiányzó jelszó (api_token.txt)")
            else:
                self.status_signal.emit(False, f"⚠️ VPS Hiba: {res.status_code}")
        except Exception:
            self.status_signal.emit(False, "🔴 VPS Nem elérhető")
            
class AutoRowSyncWorker(QThread):
    sync_finished_signal = pyqtSignal(str)

    def __init__(self, dirty_data_list):
        super().__init__()
        # "_forras" (honnan / hogyan) csak a naplóba kell, a mentésbe / felhőbe nem
        self.sources = {str(d.get("A", "")).strip().upper(): d.get("_forras", "") for d in dirty_data_list}
        self.dirty_data_list = [{k: v for k, v in d.items() if k != "_forras"} for d in dirty_data_list]

    def run(self):
        if not self.dirty_data_list:
            return

        # 1. BIZTONSÁGOS HELYI SQLITE MENTÉS A DB_FILE GLOBÁLIS VÁLTOZÓVAL
        naplo_entries = []   # (kód, előtte, utána) - a tranzakció lezárása UTÁN írjuk a naplóba
        try:
            conn = sqlite3.connect(DB_FILE)
            cursor = conn.cursor()

            for row_data in self.dirty_data_list:
                code_a = str(row_data.get("A", "")).strip().upper()
                if not code_a:
                    continue
                elotte = get_row_snapshot(code_a, cursor)

                # Ha a sorból kitörölték a terméket (Márka/Típus/Cikkszám mind üres lett),
                # a régi teljes adatok a ToroltSor munkalapra kerülnek
                if not _row_has_product_data(row_data):
                    cursor.execute(f"SELECT {', '.join(ROW_COLS)} FROM termekek WHERE UPPER(TRIM(A)) = ? OR UPPER(TRIM(B)) = ?", (code_a, code_a))
                    old = cursor.fetchone()
                    if old:
                        old_row = dict(zip(ROW_COLS, old))
                        if _row_has_product_data(old_row):
                            archive_deleted_row(old_row, "Excel Live: kiürítve", cursor)

                # Keresés a pontos A oszlop vagy B oszlop alapján
                cursor.execute("""
                    UPDATE termekek 
                    SET B=?, C=?, D=?, E=?, F=?, G=?, H=?, I=?, J=?, K=?, L=?, M=?, N=?, O=?, P=?
                    WHERE UPPER(TRIM(A)) = ? OR UPPER(TRIM(B)) = ?
                """, (
                    row_data.get("B", code_a), row_data.get("C", ""), row_data.get("D", ""),
                    row_data.get("E", ""), row_data.get("F", "1"), row_data.get("G", "HASZNÁLT"),
                    row_data.get("H", ""), row_data.get("I", ""), row_data.get("J", ""),
                    row_data.get("K", ""), row_data.get("L", ""), row_data.get("M", ""),
                    row_data.get("N", ""), row_data.get("O", ""), row_data.get("P", ""),
                    code_a, code_a
                ))

                # Ha még nem volt ilyen A kód az SQLite-ban, beszúrja!
                if cursor.rowcount == 0:
                    cursor.execute("""
                        INSERT INTO termekek (A, B, C, D, E, F, G, H, I, J, K, L, M, N, O, P, szin)
                        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 'FEHÉR')
                    """, (
                        code_a, row_data.get("B", code_a), row_data.get("C", ""),
                        row_data.get("D", ""), row_data.get("E", ""), row_data.get("F", "1"),
                        row_data.get("G", "HASZNÁLT"), row_data.get("H", ""), row_data.get("I", ""),
                        row_data.get("J", ""), row_data.get("K", ""), row_data.get("L", ""),
                        row_data.get("M", ""), row_data.get("N", ""), row_data.get("O", ""),
                        row_data.get("P", "")
                    ))

                utana = get_row_snapshot(code_a, cursor)
                if elotte != utana:
                    naplo_entries.append((code_a, elotte, utana))

            conn.commit()
            conn.close()
            print(f"[AutoRowSync] SQLite sikeresen frissítve ({len(self.dirty_data_list)} sor).")
            for code_a, elotte, utana in naplo_entries:
                src = self.sources.get(code_a, "") or "Excel Live"
                if elotte is None:
                    log_event(code_a, "FELVÉTEL (Excel Live)", "", None, utana, src)
                elif _row_has_product_data(elotte) and not _row_has_product_data(utana):
                    log_event(code_a, "TÖRLÉS (Excel Live: kiürítve)", describe_changes(elotte, utana), elotte, utana, src)
                elif not _row_has_product_data(elotte) and _row_has_product_data(utana):
                    log_event(code_a, "FELVÉTEL (Excel Live)", describe_changes(elotte, utana), elotte, utana, src)
                else:
                    log_event(code_a, "MÓDOSÍTÁS (Excel Live)", describe_changes(elotte, utana), elotte, utana, src)

        except Exception as e:
            print(f"[AutoRowSync] Helyi SQL hiba: {e}")

        # 2-3. SORBA ÁLLÍTÁS + AZONNALI KÜLDÉSI KÍSÉRLET (offline-biztos)
        # Ha van net, ez rögtön kimegy a VPS-re/Sheets-re, mint eddig. Ha nincs net,
        # a sor a sync_outbox-ban marad, és a periodikus outbox_flush_timer automatikusan
        # kiküldi, amint visszajön a kapcsolat - nem kell semmit kézzel újraküldeni.
        for row_data in self.dirty_data_list:
            enqueue_outbox(row_data)
        flush_sync_outbox()

        target_txt = "Local DB (+ VPS/Sheets, amint van internet)"
        self.sync_finished_signal.emit(f"✅ {len(self.dirty_data_list)} sor mentve: {target_txt}!")

class FastCopyLiveToTestWorker(QThread):
    # Az előző letöltött CSV ujjlenyomata: ha a táblázat azóta nem változott, nincs mit feldolgozni
    _last_csv_hash = None
    _first_pull_done = False
    _warned_sheet_dups = set()
    finished_signal = pyqtSignal(bool, str, float, list) # Szignál kiegészítve a konfliktusos kódok listájával
    process_log_signal = pyqtSignal(str, str)

    def __init__(self, pending_rows_data=None):
        super().__init__()
        self.pending_rows_data = pending_rows_data or {} # { "R18-12878": {"C": "Márka", ...} }
        self.unchanged = False

    def run(self):
        start_time = time.perf_counter()
        conflicts = []
        try:
            test_csv_url = f"https://docs.google.com/spreadsheets/d/{TEST_SHEETS_ID}/export?format=csv"
            res = requests.get(test_csv_url, timeout=6)
            res.encoding = 'utf-8'

            if res.status_code != 200:
                self.finished_signal.emit(False, "❌ Nem sikerült letölteni a Raktár táblázatot!", 0.0, [])
                return

            csv_hash = hashlib.md5(res.content).hexdigest()
            if csv_hash == FastCopyLiveToTestWorker._last_csv_hash and not self.pending_rows_data:
                self.unchanged = True
                elapsed = round(time.perf_counter() - start_time, 2)
                self.finished_signal.emit(True, f"✔ Sheets változatlan ({elapsed} mp)", elapsed, [])
                return

            df_test = pd.read_csv(io.StringIO(res.text), header=None, low_memory=False, dtype=str, keep_default_na=False)

            conn = get_db_connection()
            cursor = conn.cursor()
            cursor.execute("BEGIN TRANSACTION")

            cursor.execute(f"SELECT {', '.join(ROW_COLS)} FROM termekek")
            db_full = {}
            for row in cursor.fetchall():
                d = {c: safe_str(row[c]) for c in ROW_COLS}
                # A és B kód alapján is (ha a helyi sorban a kód kicsit másképp van írva, vagy csak a B-ben van meg)
                for k in (norm_code(d["A"]), norm_code(d["B"])):
                    if k and k not in db_full:
                        db_full[k] = d
            db_existing = db_full
            sheet_duplicates = set()
            # A még fel nem töltött (outboxban váró) tételeket a Sheets régebbi állapota nem írhatja
            # felül - és egy helyben törölt tételt sem hozhat vissza, amíg a törlés ki nem ment.
            pending_outbox = get_pending_outbox_codes(cursor)

            update_data = []
            insert_data = []
            sheet_changes = []   # (kód, előtte, utána) a Napló fülre - a Google táblázatban közvetlenül módosítva

            # Ha egy R18-szám TÖBBSZÖR szerepel a táblázatban (pl. sorok bemásolásakor az üres R18-sor is
            # megmaradt), akkor az ADATOT TARTALMAZÓ sor számít - nem az üres. Különben a bemásolt termék
            # nem látszana az Excel Live-on, a szerveren viszont ott lenne (-> "felülírod?" kérdés).
            sheet_rows = {}     # kód -> a táblázat 16 értéke (A-P)
            for row in df_test.itertuples(index=False, name=None):
                code = norm_code(safe_str(row[0]) if len(row) > 0 else "")
                if not code or not code.startswith("R18-"):
                    continue
                vals = [safe_str(row[i]) if i < len(row) else "" for i in range(16)]
                vals[0] = code
                if code in sheet_rows:
                    sheet_duplicates.add(code)
                    if _row_has_product_data(dict(zip(ROW_COLS, sheet_rows[code]))) \
                            or not _row_has_product_data(dict(zip(ROW_COLS, vals))):
                        continue
                sheet_rows[code] = vals

            for code, vals in sheet_rows.items():
                cloud_c, cloud_d, cloud_e = vals[2], vals[3], vals[4]

                # ÜTKÖZÉS-VIZSGÁLAT: Ha a felhasználónál pending (szerkesztés alatti) módosítás van erre a sorra
                # Csak VALÓDI ütközés: a táblázatban MÁS van, mint ami ezen a gépen utoljára el volt mentve
                # (vagyis valaki más írta át), ÉS más, mint amit most gépelsz. Az, hogy a táblázatban még a
                # régi érték van, amíg a te gépelésed fel nem ment, NEM ütközés.
                if code in self.pending_rows_data:
                    local_p = self.pending_rows_data[code]
                    saved = db_full.get(code) or {}
                    cloud_cd = (cloud_c.strip().upper(), cloud_d.strip().upper())
                    saved_cd = (str(saved.get("C", "")).strip().upper(), str(saved.get("D", "")).strip().upper())
                    typed_cd = (str(local_p.get("C", "")).strip().upper(), str(local_p.get("D", "")).strip().upper())
                    if saved and cloud_c and cloud_cd != saved_cd and cloud_cd != typed_cd:
                        conflicts.append(code)

                if code.upper() in pending_outbox:
                    continue

                if code in db_existing:
                    # Ha nem szerkesztés alatti sor, frissítjük a helyit a felhőből
                    # (a színjelölést - szin - nem bántjuk, az a CSV-ben nincs benne)
                    if code not in self.pending_rows_data:
                        local_row = db_full[code]
                        # Csak a ténylegesen megváltozott sorokat írjuk (nem mind a több ezret 15 mp-enként)
                        if [local_row[c] for c in ROW_COLS] != vals[:16]:
                            update_data.append((*vals[:16], code, code))
                            sheet_changes.append((code, dict(local_row), dict(zip(ROW_COLS, vals[:16]))))
                else:
                    insert_data.append((*vals[:16], 'FEHÉR'))
                    db_existing[code] = dict(zip(ROW_COLS, vals[:16]))

            if sheet_duplicates:
                new_dups = sheet_duplicates - FastCopyLiveToTestWorker._warned_sheet_dups
                if new_dups:
                    FastCopyLiveToTestWorker._warned_sheet_dups |= new_dups
                    self.process_log_signal.emit(
                        f"⚠️ A Google táblázatban ezek az R18-számok TÖBBSZÖR szerepelnek: {', '.join(sorted(new_dups)[:20])}"
                        f"{' ...' if len(new_dups) > 20 else ''} - a program az adatot tartalmazó sort használja. "
                        f"Töröld a felesleges sort a táblázatban.", "ERROR")

            if update_data:
                cursor.executemany("""
                    UPDATE termekek SET A=?, B=?, C=?, D=?, E=?, F=?, G=?, H=?, I=?, J=?, K=?, L=?, M=?, N=?, O=?, P=?
                    WHERE UPPER(TRIM(A))=? OR UPPER(TRIM(B))=?
                """, update_data)

            if insert_data:
                cursor.executemany("""
                    INSERT OR IGNORE INTO termekek (A, B, C, D, E, F, G, H, I, J, K, L, M, N, O, P, szin)
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                """, insert_data)

            conn.commit()
            conn.close()

            FastCopyLiveToTestWorker._last_csv_hash = csv_hash

            # Napló: a Google táblázatban KÖZVETLENÜL végzett módosítások (az első letöltésnél nem - az csak
            # a kiinduló állapot; tömeges változásnál csak összesítés, hogy ne árassza el a naplót)
            if FastCopyLiveToTestWorker._first_pull_done and sheet_changes:
                if len(sheet_changes) <= 50:
                    for code, elotte, utana in sheet_changes:
                        kind = ("TÖRLÉS (Google Sheets-ben kiürítve)" if _row_has_product_data(elotte) and not _row_has_product_data(utana)
                                else "FELVÉTEL (Google Sheets-ben)" if not _row_has_product_data(elotte) and _row_has_product_data(utana)
                                else "MÓDOSÍTÁS (Google Sheets-ben)")
                        log_event(code, kind, describe_changes(elotte, utana), elotte, utana,
                                  "📊 Google Sheets - közvetlenül a táblázatban írták át")
                else:
                    log_event("TÖBB", "MÓDOSÍTÁS (Google Sheets-ben, tömeges)",
                              f"{len(sheet_changes)} sor változott egyszerre a Google táblázatban "
                              f"(pl. {', '.join(c for c, _, _ in sheet_changes[:10])} ...)", forras="📊 Google Sheets")
            FastCopyLiveToTestWorker._first_pull_done = True

            elapsed = round(time.perf_counter() - start_time, 2)
            self.finished_signal.emit(True, f"✔ Háttér szinkron kész ({elapsed} mp)!", elapsed, conflicts)

        except Exception as e:
            elapsed = round(time.perf_counter() - start_time, 2)
            self.finished_signal.emit(False, f"❌ Hiba a szinkronizálás során: {e}", elapsed, [])

class SaveWorker(QThread):
    finished_signal = pyqtSignal(str, float)
    process_log_signal = pyqtSignal(str, str)

    def __init__(self, code_final, uj_sor_adatok):
        super().__init__()
        self.code_final = code_final
        self.uj_sor_adatok = [str(x).upper() if isinstance(x, str) else x for x in uj_sor_adatok]

    def run(self):
        start_time = time.perf_counter()
        if not SHEETS_ENABLED:
            self.finished_signal.emit("Helyi mentés kész (Google Sheets kikapcsolva)", 0.0)
            return
        try:
            self.process_log_signal.emit(f"🔌 Kapcsolódás a Google Sheets API-hoz [{self.code_final}]...", "INFO")
            import gspread
            from google.oauth2.service_account import Credentials
            scopes = ["https://www.googleapis.com/auth/spreadsheets"]
            creds = Credentials.from_service_account_file(CREDENTIALS_FILE, scopes=scopes)
            client = gspread.authorize(creds)
            sheet = client.open_by_key(TEST_SHEETS_ID).sheet1

            self.process_log_signal.emit(f"🔎 Sor keresése az A oszlopban [{self.code_final}]...", "INFO")
            
            all_a = sheet.col_values(1)
            target_row = None

            for idx, val in enumerate(all_a):
                if str(val).strip().upper() == self.code_final.upper():
                    target_row = idx + 1
                    break

            if not target_row:
                target_row = len(all_a) + 1

            if target_row > sheet.row_count:
                sheet.add_rows(100)

            cell_range = f"A{target_row}:P{target_row}"
            self.process_log_signal.emit(f"➕ Sor ÍRÁSA a felhőben (Pontos Sor: {target_row})...", "INFO")
            sheet.update(range_name=cell_range, values=[self.uj_sor_adatok[:16]])

            elapsed = round(time.perf_counter() - start_time, 2)
            self.process_log_signal.emit(f"✔ A(z) {self.code_final} tétel sikeresen elmentve a felhőbe! (Sor: {target_row}, Idő: {elapsed} mp)", "SUCCESS")
            self.finished_signal.emit(f"Sikeres felhő szinkronizáció! ({elapsed} mp)", elapsed)
        except Exception as e:
            elapsed = round(time.perf_counter() - start_time, 2)
            self.process_log_signal.emit(f"❌ Hiba a felhő mentés során: {e}", "ERROR")
            self.finished_signal.emit(f"Hiba: {e}", elapsed)

class ClearWorker(QThread):
    finished_signal = pyqtSignal(str)
    process_log_signal = pyqtSignal(str, str)

    def __init__(self, code):
        super().__init__()
        self.code = code

    def run(self):
        try:
            # 0. Teljes adatok archiválása a ToroltSor munkalapra (kiürítés előtt!)
            archive_row_from_db(self.code, "Program: felvétel visszavonva")

            # 1. Helyi mentés (net nélkül is azonnal megtörténik)
            conn = get_db_connection()
            cursor = conn.cursor()
            cursor.execute("""
                UPDATE termekek
                SET C='', D='', E='', F='', G='', H='', I='', J='', K='', L='', M='', N='', O='', P='', szin='FEHÉR'
                WHERE A=? OR B=?
            """, (self.code, self.code))
            conn.commit()
            conn.close()

            # 2. Felhő (VPS + Sheets) az outboxon át: ha nincs net, később automatikusan kimegy
            enqueue_row_from_db(self.code, force=True)
            enqueue_color(self.code, "FEHÉR", COLOR_WHITE)
            flush_outbox_async()
            self.process_log_signal.emit(f"✔ A(z) {self.code} adatai kiürítve helyben (felhő: amint van internet)", "SUCCESS")

            self.finished_signal.emit(f"↩️ A(z) {self.code} tétel sikeresen visszavonva (A/B kód megmaradt)!")
        except Exception as e:
            err_msg = f"❌ Visszavonási hiba: {e}"
            self.process_log_signal.emit(err_msg, "ERROR")
            self.finished_signal.emit(err_msg)

class AutoVPSPollWorker(QThread):
    data_updated_signal = pyqtSignal(list)
    presence_signal = pyqtSignal(list)     # ki melyik Excel Live soron dolgozik most (másik gépek)

    def run(self):
        # Az utoljára feldolgozott VPS-verzió a beallitasok.json-ban marad meg, így
        # újraindítás után a kimaradt változásokat is letölti. Ha még nincs, az első
        # ('since' nélküli) lekérés csak a VPS aktuális verzióját adja vissza.
        last_verzio = load_config_value("vps_verzio")
        while True:
            try:
                # 1 másodpercenként lekérdezi a VPS-ről az azóta változott termékeket
                time.sleep(1)
                params = {} if last_verzio is None else {"since": last_verzio}
                res = vps_request("GET", f"{VPS_URL}/get_updates", params=params, timeout=2)
                if res.status_code != 200:
                    continue
                data = res.json()
                if data.get("items"):
                    self.data_updated_signal.emit(data["items"])
                if data.get("nyomtatas_var") and is_print_host() and not _print_queue_lock.locked():
                    threading.Thread(target=process_print_queue, daemon=True).start()
                jel = sorted((j.get("gep", ""), j.get("A", "")) for j in (data.get("jelenlet") or []) if j.get("gep") != GEP_NEV)
                if jel != getattr(self, "_last_jel", None):
                    self._last_jel = jel
                    self.presence_signal.emit([{"gep": g, "A": a} for g, a in jel])
                new_verzio = data.get("verzio")
                if new_verzio is not None and new_verzio != last_verzio:
                    last_verzio = new_verzio
                    save_config_value("vps_verzio", last_verzio)
            except Exception:
                pass
            
class VPSDownloadWorker(QThread):
    finished_signal = pyqtSignal(bool, str)
    process_log_signal = pyqtSignal(str, str)

    def run(self):
        try:
            self.process_log_signal.emit("🔌 VPS Adatbázis letöltése folyamatban...", "INFO")
            
            # 1. VPS Adatbázis lekérése / Sync indítása
            url = f"{VPS_URL}/sync/full"
            res = vps_request("POST", url, timeout=10)
            
            if res.status_code == 200:
                msg = "✅ VPS Adatbázis sikeresen frissítve a helyi gépen!"
                self.process_log_signal.emit(msg, "SUCCESS")
                self.finished_signal.emit(True, msg)
            else:
                raise Exception(f"VPS Válasz hiba: {res.status_code}")

        except Exception as e:
            err_msg = f"❌ VPS Szinkron hiba: {str(e)}"
            self.process_log_signal.emit(err_msg, "ERROR")
            self.finished_signal.emit(False, err_msg)
            
class CardUpdateWorker(QThread):
    finished_signal = pyqtSignal(str)
    process_log_signal = pyqtSignal(str, str)

    def __init__(self, row_id, code, marka, tipus, cikkszam, sulij, ar_val, qty_val, pater_val="", polc_val="", hely_val=""):
        super().__init__()
        self.row_id = row_id
        self.code = code
        self.marka = marka
        self.tipus = tipus
        self.cikkszam = cikkszam
        self.sulij = sulij
        self.ar_val = ar_val
        self.qty_val = qty_val
        self.pater_val = pater_val
        self.polc_val = polc_val
        self.hely_val = hely_val

    def run(self):
        start_time = time.perf_counter()
        try:
            self.process_log_signal.emit(f"⚙️ Kártya adatai és helyszíne frissítése az SQLite adatbázisban [{self.code}]...", "INFO")
            elotte = get_row_snapshot(self.code)
            conn = get_db_connection()
            cursor = conn.cursor()
            cursor.execute("""
                UPDATE termekek 
                SET C = ?, D = ?, E = ?, M = ?, H = ?, F = ?, I = ?, K = ?, L = ? 
                WHERE A = ? OR B = ?
            """, (self.marka, self.tipus, self.cikkszam, self.sulij, self.ar_val, self.qty_val, self.pater_val, self.polc_val, self.hely_val, self.code, self.code))
            conn.commit()
            conn.close()

            utana = get_row_snapshot(self.code)
            if elotte != utana:
                log_event(self.code, "MÓDOSÍTÁS (kártya)", describe_changes(elotte, utana), elotte, utana,
                          "Kártya (Gyorskereső / PDA) - 💾 Adatok mentése gomb")

            # Felhő (VPS + Sheets) az outboxon át: ha nincs net, később automatikusan kimegy.
            # force: a kártyán szándékosan javítjuk a Márkát/Típust/Cikkszámot, ez nem R18-ütközés.
            enqueue_row_from_db(self.code, force=True)
            flush_outbox_async()

            elapsed = round(time.perf_counter() - start_time, 2)
            self.process_log_signal.emit(f"✔ Adatok és helyszín mentve [{self.code}] ({elapsed} mp)!", "SUCCESS")
            self.finished_signal.emit(f"✔ Mentés sikeres [{self.code}]!")
        except Exception as e:
            self.process_log_signal.emit(f"❌ Mentési hiba: {e}", "ERROR")
            self.finished_signal.emit(f"Hiba: {e}")

class PreciseManualSyncWorker(QThread):
    finished_signal = pyqtSignal(bool, str)
    process_log_signal = pyqtSignal(str, str)

    def __init__(self, selected_local):
        super().__init__()
        self.selected_local = selected_local

    def run(self):
        if not SHEETS_ENABLED:
            self.finished_signal.emit(False, "ℹ️ A Google Sheets szinkron ki van kapcsolva.")
            return
        try:
            self.process_log_signal.emit("🔌 Kapcsolódás a Google Sheets-hez...", "INFO")

            import gspread
            from google.oauth2.service_account import Credentials
            scopes = ["https://www.googleapis.com/auth/spreadsheets"]
            creds = Credentials.from_service_account_file(CREDENTIALS_FILE, scopes=scopes)
            client = gspread.authorize(creds)
            sheet = client.open_by_key(TEST_SHEETS_ID).sheet1

            self.process_log_signal.emit("🔎 A oszlop sorszámainak beolvasása...", "INFO")
            all_a_codes = [str(x).strip().upper() for x in sheet.col_values(1)]

            updated_rows_count = 0
            format_requests = []

            for r in self.selected_local:
                code_key = safe_str(r["A"]).strip().upper()
                row_vals = [
                    safe_str(r["A"]), safe_str(r["B"]), safe_str(r["C"]), safe_str(r["D"]),
                    safe_str(r["E"]), safe_str(r["F"]), safe_str(r["G"]), safe_str(r["H"]),
                    safe_str(r["I"]), "",               safe_str(r["K"]), safe_str(r["L"]),
                    safe_str(r["M"]), safe_str(r["N"]), safe_str(r["O"]), safe_str(r["P"])
                ]

                if code_key in all_a_codes:
                    target_row = all_a_codes.index(code_key) + 1
                else:
                    target_row = len(all_a_codes) + 1
                    all_a_codes.append(code_key)

                if target_row > sheet.row_count:
                    sheet.add_rows(50)

                # 1. A-P SOR FELTÖLTÉSE
                cell_range = f"A{target_row}:P{target_row}"
                sheet.update(range_name=cell_range, values=[row_vals])

                # 2. IGAZÍTÁSI FORMÁZÁSOK ELŐKÉSZÍTÉSE A SHEETS FELÜLETRE
                val_l = safe_str(r["L"]).strip()
                is_pure_number_l = bool(val_l and re.match(r'^\d+([.,]\d+)?$', val_l))

                # I (Páter: 9. oszlop / col 8) -> JOBBRA
                format_requests.append(sheet.format(f"I{target_row}", {"horizontalAlignment": "RIGHT"}))
                # K (Polc: 11. oszlop / col 10) -> JOBBRA
                format_requests.append(sheet.format(f"K{target_row}", {"horizontalAlignment": "RIGHT"}))
                # M (Súly: 13. oszlop / col 12) -> JOBBRA
                format_requests.append(sheet.format(f"M{target_row}", {"horizontalAlignment": "RIGHT"}))

                # L (Raklap/Hely: 12. oszlop / col 11) -> CSAK HA TISZTA SZÁM
                align_l = "RIGHT" if is_pure_number_l else "LEFT"
                format_requests.append(sheet.format(f"L{target_row}", {"horizontalAlignment": align_l}))

                updated_rows_count += 1
                self.process_log_signal.emit(f"⬆️ [{updated_rows_count}] {code_key} adatai és igazításai frissítve (Sor: {target_row})", "INFO")

            msg = f"✅ {updated_rows_count} tétel adatai és igazításai (I, K, M jobbra, L feltételesen) sikeresen feltöltve a Google Sheets-be!"
            self.process_log_signal.emit(msg, "SUCCESS")
            self.finished_signal.emit(True, msg)

        except Exception as e:
            err_msg = f"❌ Szinkronizálási hiba: {str(e)}"
            self.process_log_signal.emit(err_msg, "ERROR")
            self.finished_signal.emit(False, err_msg)

class DeleteWorker(QThread):
    finished_signal = pyqtSignal(str)
    process_log_signal = pyqtSignal(str, str)

    def __init__(self, code, reason="", forras=""):
        super().__init__()
        self.code = code
        self.reason = reason
        self.forras = forras

    def run(self):
        try:
            elotte = get_row_snapshot(self.code)
            # 0. Teljes adatok archiválása a ToroltSor munkalapra (törlés előtt!)
            archive_row_from_db(self.code, "Program: törlés / eladás")

            # 1. Helyi törlés (net nélkül is azonnal megtörténik)
            conn = get_db_connection()
            cursor = conn.cursor()
            cursor.execute("DELETE FROM termekek WHERE A = ? OR B = ?", (self.code, self.code))
            conn.commit()
            conn.close()

            # 2. Felhő (Sheets + VPS, a másik gép is törli) az outboxon át
            enqueue_delete(self.code)
            flush_outbox_async()
            self.process_log_signal.emit(f"✔ A(z) {self.code} törölve helyben (felhő: amint van internet)", "SUCCESS")

            log_event(self.code, "TÖRLÉS", self.reason or "A termék sora törölve", elotte, None,
                      self.forras or "Kártya - 🗑️ Teljes törlés / eladás gomb")
            self.finished_signal.emit(f"🗑️ A(z) {self.code} termék sikeresen törölve!")
        except Exception as e:
            err_msg = f"❌ Törlési hiba: {e}"
            self.process_log_signal.emit(err_msg, "ERROR")
            self.finished_signal.emit(err_msg)

class ColorWorker(QThread):
    finished_signal = pyqtSignal(str)

    def __init__(self, code, color_dict, color_name):
        super().__init__()
        self.code = code
        self.color_dict = color_dict
        self.color_name = color_name

    def run(self):
        try:
            # 1. Helyi mentés, 2. felhő (Sheets színezés + VPS) az outboxon át
            conn = get_db_connection()
            cursor = conn.cursor()
            cursor.execute("UPDATE termekek SET szin = ? WHERE A = ? OR B = ?", (self.color_name.upper(), self.code, self.code))
            conn.commit()
            conn.close()
            enqueue_color(self.code, self.color_name.upper(), self.color_dict)
            flush_outbox_async()
            log_event(self.code, "SZÍN MÓDOSÍTÁSA", f"Új szín: {self.color_name}")
            self.finished_signal.emit(f"🎨 Átszínezve {self.color_name} színre!")
        except Exception as e:
            self.finished_signal.emit(f"Színezési hiba: {e}")

class SummaryDialog(QDialog):
    def __init__(self, tipus_val, cikk_val, code_val, total_sum, parent=None):
        super().__init__(parent)
        
        self.search_term = get_smart_part_number(tipus_val, cikk_val)
        if not self.search_term or self.search_term == "ISMERETLEN":
            self.search_term = code_val

        self.setWindowTitle(f"📦 Azonos termékek összesítése: [{self.search_term}]")
        self.resize(850, 500)
        self.setStyleSheet("QDialog { background-color: #121212; color: white; } QLabel { color: white; font-weight: bold; }")
        
        layout = QVBoxLayout(self)
        
        self.lbl_title = QLabel(f"🔍 Azonos azonosítójú (<b>{self.search_term}</b>) tételek")
        self.lbl_title.setStyleSheet("font-size: 16px; color: #29b6f6; margin-bottom: 10px;")
        layout.addWidget(self.lbl_title)

        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        container = QWidget()
        lay_container = QVBoxLayout(container)
        lay_container.setAlignment(Qt.AlignmentFlag.AlignTop)

        try:
            conn = get_db_connection()
            cursor = conn.cursor()
            cursor.execute("SELECT * FROM termekek")
            rows = cursor.fetchall()
            conn.close()

            matched_rows = []
            clean_target = clean_part_number(self.search_term)

            for r in rows:
                r_dict = dict(r)
                d_txt = safe_str(r_dict.get("D"))
                e_txt = safe_str(r_dict.get("E"))
                
                item_smart_code = get_smart_part_number(d_txt, e_txt)
                
                if item_smart_code.upper() == self.search_term.upper() or \
                   (clean_target and clean_part_number(item_smart_code) == clean_target):
                    matched_rows.append(r_dict)

            active_sum = 0
            sold_count = 0

            for r_dict in matched_rows:
                color_val = safe_str(r_dict.get("szin")).upper()
                c_val = safe_str(r_dict.get("C"))
                is_sold = (color_val == "LILA") or (not c_val)

                db_str = safe_str(r_dict.get("F"))
                f_num = int(re.search(r'\d+', db_str).group(0)) if re.search(r'\d+', db_str) else 1

                if is_sold:
                    sold_count += 1
                else:
                    active_sum += f_num

            self.lbl_title.setText(f"🔍 Azonos azonosítójú (<b>{self.search_term}</b>) tételek | 📊 Összesen raktáron: <span style='color: #00e676;'>{active_sum} db</span> | 🟣 Eladva/Törölve: <span style='color: #ab47bc;'>{sold_count} db</span>")

            if not matched_rows:
                lay_container.addWidget(QLabel("Nincs más azonos elem a raktárban."))
            else:
                for r_dict in matched_rows:
                    code = get_best_code(r_dict)
                    db = safe_str(r_dict.get("F")) or "1"
                    allapot = safe_str(r_dict.get("G")) or "HASZNÁLT"
                    ar = safe_str(r_dict.get("H")) or "Nincs"
                    hely = safe_str(r_dict.get("L")) or "Ismeretlen"
                    _pater, _polc = safe_str(r_dict.get("I")), safe_str(r_dict.get("K"))
                    hely_txt = hely + (f" · Páter: {_pater}" if _pater else "") + (f" · Oszlop: {_polc}" if _polc else "")
                    
                    color_val = safe_str(r_dict.get("szin")).upper()
                    c_val = safe_str(r_dict.get("C"))
                    is_sold = (color_val == "LILA") or (not c_val)

                    box = QFrame()
                    if is_sold:
                        box.setStyleSheet("QFrame { background-color: rgba(171, 71, 188, 0.25); border: 2px solid #ab47bc; border-radius: 6px; margin-bottom: 6px; padding: 8px; }")
                        info_text = f"🏷️ <b>{code}</b> | 🟣 <b style='color: #ab47bc;'>[ELADVA / TÖRÖLVE]</b> | 📍 Utolsó Hely: <b>{hely_txt}</b>"
                    else:
                        box.setStyleSheet("QFrame { background-color: #1e1e1e; border: 1px solid #333; border-radius: 6px; margin-bottom: 6px; padding: 8px; }")
                        info_text = f"🏷️ <b>{code}</b> | 📦 Mennyiség: <b>{db} db</b> | 📍 Hely: <b>{hely_txt}</b> | Állapot: <b>{allapot}</b> | Ár: <b>{ar} EUR</b>"

                    box_lay = QHBoxLayout(box)
                    lbl_info = QLabel(info_text)
                    box_lay.addWidget(lbl_info)

                    if hely and hely != "Ismeretlen" and not re.search(r'[A-Za-z]', hely):
                        btn_m = QPushButton(f"🗺️ Térkép ({hely})")
                        btn_m.setStyleSheet("background-color: #00bcd4; color: black; font-weight: bold; padding: 4px 8px;")
                        
                        btn_m.clicked.connect(
                            lambda checked, b=btn_m, h=hely: self.run_button_action(
                                b, 
                                f"🗺️ Térkép ({h})", 
                                lambda: MapDialog(h, self).exec()
                            )
                        )
                        box_lay.addWidget(btn_m)

                    lay_container.addWidget(box)

        except Exception as e:
            lay_container.addWidget(QLabel(f"Hiba a lekérdezéskor: {e}"))

        scroll.setWidget(container)
        layout.addWidget(scroll)

        btn_close = QPushButton("Bezárás")
        btn_close.clicked.connect(self.accept)
        layout.addWidget(btn_close)
        

# -------------------------------------------------------------------
# EXCEL-SZERŰ TÁBLÁZAT (Excel Live munkalap)
#  - Delete / Backspace: a kijelölt cellák tartalma törlődik (nem kell belépni a cellába)
#  - Gépelés közben a nyilak elmentik az értéket és átlépnek a szomszéd cellába
# -------------------------------------------------------------------
# Excel Live munkalap KÉPERNYŐS oszlopsorrendje (logikai indexek: 0-15 = A-P, 16 = 🖨️ gomb).
# G (Állapot) után jön L (Raklap) és M (Súly), mert ezeket írjuk a leggyakrabban.
EXCEL_LIVE_VISUAL_ORDER = [16, 0, 1, 2, 3, 4, 5, 6, 11, 12, 8, 9, 10, 7, 13, 14, 15]
# A J (Szabad, logikai index 9) oszlop nincs használatban - az Excel Live munkalapon rejtve van
EXCEL_LIVE_HIDDEN_COL_J = 9
# Excel Live munkalap sormagassága (px) - a C/D/E hosszú szövege így két sorban kifér
EXCEL_LIVE_ROW_HEIGHT = 46


class PrintButtonDelegate(QStyledItemDelegate):
    """Az Excel Live 🖨️ oszlopának gombjait csak RAJZOLJA. Korábban minden sor (több ezer!) kapott egy
    valódi QPushButton-t, ami fülváltáskor / görgetéskor / frissítéskor nagyon lassított régebbi gépen.
    Kattintás: on_click(sor) - pontosan az a sor, amelyikre kattintottak."""

    LABEL = "🖨️ Nyomtatás"

    def __init__(self, on_click, parent=None):
        super().__init__(parent)
        self.on_click = on_click

    def paint(self, painter, option, index):
        painter.save()
        painter.setRenderHint(QPainter.RenderHint.Antialiasing)
        rect = option.rect.adjusted(3, 3, -3, -3)
        hover = bool(option.state & QStyle.StateFlag.State_MouseOver)
        painter.setPen(Qt.PenStyle.NoPen)
        painter.setBrush(QColor("#039be5" if hover else "#0288d1"))
        painter.drawRoundedRect(QRectF(rect), 3, 3)
        font = QFont(option.font)
        font.setBold(True)
        painter.setFont(font)
        painter.setPen(QColor("#ffffff"))
        painter.drawText(rect, Qt.AlignmentFlag.AlignCenter, self.LABEL)
        painter.restore()

    def createEditor(self, parent, option, index):
        return None

    def helpEvent(self, event, view, option, index):
        from PyQt6.QtWidgets import QToolTip
        QToolTip.showText(event.globalPos(),
                          "Kattintás: példányszám megadása (alap: 1, Enter) majd nyomtatás | "
                          "Shift+kattintás: nyomtatási ablak (PNG / előnézet)\nBillentyűvel: Ctrl+P a kijelölt sorra", view)
        return True

    def editorEvent(self, event, model, option, index):
        et = event.type()
        if et == QEvent.Type.MouseButtonRelease and event.button() == Qt.MouseButton.LeftButton:
            if option.rect.contains(event.position().toPoint()):
                row = index.row()
                QTimer.singleShot(0, lambda: self.on_click(row))
            return True
        if et in (QEvent.Type.MouseButtonPress, QEvent.Type.MouseButtonDblClick):
            return True
        return super().editorEvent(event, model, option, index)


class FrozenColumn(QObject):
    """Rögzített (mindig látható) oszlop a táblázat BAL OLDALÁN, külön sávban – mint az Excelben az „ablaktábla rögzítése”.
    A sáv (self.v) a sorszámokat és a rögzített oszlopot mutatja, ugyanazzal az adattal és kijelöléssel; fel-le együtt
    görög a táblázattal, oldalra nem. A táblázatban az oszlop és a sorszám el van rejtve, így az az A oszloppal kezdődik
    és semmit nem takar el. A sávot a hívó teszi a táblázat elé a layoutban."""

    def __init__(self, table, col, delegate):
        super().__init__(table)
        from PyQt6.QtWidgets import QTableView
        self.t, self.col = table, col
        v = QTableView()
        self.v = v
        v.setModel(table.model())
        v.setSelectionModel(table.selectionModel())
        v.setFocusPolicy(Qt.FocusPolicy.NoFocus)
        v.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Fixed)
        v.horizontalHeader().setStretchLastSection(False)
        v.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOn)   # ugyanolyan magas legyen, mint a táblázat
        v.setVerticalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        v.setVerticalScrollMode(table.verticalScrollMode())
        v.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        v.setSelectionBehavior(table.selectionBehavior())
        v.setSelectionMode(table.selectionMode())
        v.setMouseTracking(True)
        v.setStyleSheet(table.styleSheet().replace("QTableWidget", "QTableView"))
        v.setItemDelegateForColumn(col, delegate)
        for c in range(table.columnCount()):
            v.setColumnHidden(c, c != col)
        v.setColumnWidth(col, max(table.columnWidth(col), 140))
        vh, tvh = v.verticalHeader(), table.verticalHeader()
        vh.setMinimumSectionSize(tvh.minimumSectionSize())
        vh.setDefaultSectionSize(tvh.defaultSectionSize())
        for r in range(table.rowCount()):
            v.setRowHeight(r, table.rowHeight(r))
        # a táblázatban: a rögzített oszlop és a sorszám rejtve (a sávban látszanak)
        table.setColumnHidden(col, True)
        tvh.hide()
        table.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOn)
        table.horizontalHeader().sectionResized.connect(self._col_resized)
        tvh.sectionResized.connect(lambda i, o, n: v.setRowHeight(i, n))
        vh.geometriesChanged.connect(self.update_width)
        v.verticalScrollBar().valueChanged.connect(table.verticalScrollBar().setValue)
        table.verticalScrollBar().valueChanged.connect(v.verticalScrollBar().setValue)
        self.update_width()

    def _col_resized(self, idx, old, new):
        if idx == self.col and new > 0:
            self.v.setColumnWidth(self.col, new)
            self.update_width()

    def update_width(self):
        v = self.v
        v.setFixedWidth(v.verticalHeader().width() + v.columnWidth(self.col) + 2 * v.frameWidth())


# G (Állapot) oszlop választható értékei az Excel Live munkalapon
ALLAPOT_OPTIONS = ["ÚJ", "HASZNÁLT", "FELÚJÍTOTT"]
ALLAPOT_COL = 6   # G


def allapot_dropdown_enabled():
    """Be van-e kapcsolva a G (Állapot) legördülő lista ezen a gépen (Excel Live: ☑ jelölőnégyzet)."""
    return bool(load_config_value("allapot_legordulo", True))


class ExcelCellDelegate(QStyledItemDelegate):
    """Cellaszerkesztő: a nyilak szerkesztés közben is átléptetnek a szomszéd cellába.
    A G (Állapot) oszlopban legördülő lista: ÚJ / HASZNÁLT / FELÚJÍTOTT."""

    _ARROWS = (Qt.Key.Key_Left, Qt.Key.Key_Right, Qt.Key.Key_Up, Qt.Key.Key_Down)
    _MODS = (Qt.KeyboardModifier.ControlModifier | Qt.KeyboardModifier.AltModifier |
             Qt.KeyboardModifier.ShiftModifier | Qt.KeyboardModifier.MetaModifier)

    def createEditor(self, parent, option, index):
        if index.column() == ALLAPOT_COL and allapot_dropdown_enabled():
            combo = QComboBox(parent)
            combo.addItems(ALLAPOT_OPTIONS)
            combo.setStyleSheet("QComboBox { background-color: #2b2b2b; color: white; font-weight: bold; } "
                                "QComboBox QAbstractItemView { background-color: #1e1e1e; color: white; font-size: 15px; "
                                "selection-background-color: #29b6f6; selection-color: black; }")
            # Választás után azonnal ment és bezárul (nem kell Entert nyomni)
            def chosen(_i, c=combo):
                table = self.parent()
                table._edit_source = "🖱️ Állapot választólista (G)"
                try:
                    self.commitData.emit(c)
                finally:
                    table._edit_source = None
                self.closeEditor.emit(c, QAbstractItemDelegate.EndEditHint.NoHint)
            combo.activated.connect(chosen)
            QTimer.singleShot(0, combo.showPopup)   # a lista rögtön lenyílik
            return combo
        editor = super().createEditor(parent, option, index)
        table = self.parent()
        if editor is not None:
            # gépeléssel indult a szerkesztés (nem dupla kattintással / F2-vel)?
            editor.setProperty("typing_mode", bool(getattr(table, "_typing_edit", False)))
        return editor

    def setEditorData(self, editor, index):
        if isinstance(editor, QComboBox):
            cur = str(index.data() or "").strip().upper()
            if cur and editor.findText(cur) < 0:
                editor.insertItem(0, cur)          # régi, listán kívüli érték is megmarad választhatónak
            editor.setCurrentIndex(max(editor.findText(cur), 0) if cur else editor.findText("HASZNÁLT"))
            return
        super().setEditorData(editor, index)

    def setModelData(self, editor, model, index):
        if isinstance(editor, QComboBox):
            if str(index.data() or "") != editor.currentText():
                model.setData(index, editor.currentText())
            return
        super().setModelData(editor, model, index)

    @staticmethod
    def _should_leave_cell(editor, key):
        if key in (Qt.Key.Key_Up, Qt.Key.Key_Down):
            return True
        if editor.property("typing_mode"):
            return True                       # gépelés után Jobbra/Balra mindig lép
        if isinstance(editor, QLineEdit):
            # dupla kattintással / F2-vel indított szerkesztés: a nyíl a szövegben mozog,
            # és csak a szöveg szélén lép át a szomszéd cellába
            if editor.hasSelectedText():
                return False
            if key == Qt.Key.Key_Left:
                return editor.cursorPosition() == 0
            if key == Qt.Key.Key_Right:
                return editor.cursorPosition() == len(editor.text())
        return False

    def eventFilter(self, editor, event):
        if isinstance(editor, QComboBox):
            return super().eventFilter(editor, event)   # a lista nyilakkal a saját elemei között lép
        if event.type() == QEvent.Type.KeyPress and not (event.modifiers() & self._MODS):
            key = event.key()
            if key == Qt.Key.Key_F2:
                # F2: átváltás szövegszerkesztő módba (a nyilak a kurzort mozgatják)
                editor.setProperty("typing_mode", False)
                return True
            if key in self._ARROWS and self._should_leave_cell(editor, key):
                table = self.parent()
                self.commitData.emit(editor)                                   # érték mentése
                self.closeEditor.emit(editor, QAbstractItemDelegate.EndEditHint.NoHint)
                if hasattr(table, "move_current_cell"):
                    table.move_current_cell(key)
                return True
        return super().eventFilter(editor, event)


class ExcelTableWidget(QTableWidget):
    """Excel-szerű táblázat: Delete/Backspace törlés, nyilakkal lépkedés és Ctrl+Z visszavonás"""

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self._typing_edit = False
        self.undo_stack = [] # Visszavonási előzmények: [(row, col, old_text, new_text)]
        self.setItemDelegate(ExcelCellDelegate(self))

    def move_current_cell(self, key):
        # Jobbra/balra a KÉPERNYŐN látható sorrend szerint lép (az oszlopok át vannak rendezve),
        # a nyomtatás-gomb oszlopát (16) átugorja
        hdr = self.horizontalHeader()
        r, c = self.currentRow(), self.currentColumn()
        if key in (Qt.Key.Key_Left, Qt.Key.Key_Right):
            step = -1 if key == Qt.Key.Key_Left else 1
            data_visuals = sorted(hdr.visualIndex(col) for col in range(min(16, self.columnCount()))
                                  if not self.isColumnHidden(col))
            cur_v = hdr.visualIndex(c) if c >= 0 else data_visuals[0]
            if step > 0:
                nxt = [v for v in data_visuals if v > cur_v]
                v = nxt[0] if nxt else data_visuals[-1]
            else:
                prv = [v for v in data_visuals if v < cur_v]
                v = prv[-1] if prv else data_visuals[0]
            c = hdr.logicalIndex(v)
        elif key == Qt.Key.Key_Up:
            r -= 1
        elif key == Qt.Key.Key_Down:
            r += 1
        r = max(0, min(self.rowCount() - 1, r))
        c = max(0, min(15, c))
        self.setCurrentCell(r, c)

    # ---- Excel-szerű másolás / beillesztés ----
    DATA_COLS = 16   # 0-15 = A-P (logikai sorrend = a Google táblázat oszlopsorrendje)

    def _selected_rows_cols(self):
        idx = [i for i in self.selectedIndexes() if i.column() < self.DATA_COLS]
        if not idx and self.currentRow() >= 0 and 0 <= self.currentColumn() < self.DATA_COLS:
            idx = [self.model().index(self.currentRow(), self.currentColumn())]
        rows = sorted({i.row() for i in idx})
        cells = {(i.row(), i.column()) for i in idx}
        return rows, cells

    def full_selected_rows(self):
        """A teljesen kijelölt sorok (sorszámra kattintva / Shift+Szóköz)."""
        sm = self.selectionModel()
        return sorted(r for r in {i.row() for i in self.selectedIndexes()} if sm.isRowSelected(r))

    def copy_selection(self):
        """Ctrl+C: a kijelölt cellák tabulátorral elválasztva, a GOOGLE TÁBLÁZAT oszlopsorrendjében (A-P).
        Teljes sor kijelölésénél mind a 16 oszlop (A-P, a rejtett J-vel együtt) - így a Sheets-be
        beillesztve minden a saját oszlopába kerül."""
        rows, cells = self._selected_rows_cols()
        if not rows:
            return 0
        full = set(self.full_selected_rows())
        cols = sorted({c for _, c in cells} | (set(range(self.DATA_COLS)) if full else set()))
        lines = []
        for r in rows:
            use = range(self.DATA_COLS) if r in full else cols
            vals = []
            for c in use:
                it = self.item(r, c)
                vals.append(it.text() if it and ((r, c) in cells or r in full) else "")
            lines.append("\t".join(vals))
        QApplication.clipboard().setText("\n".join(lines))
        w = self.window()
        if hasattr(w, "trigger_toast"):
            w.trigger_toast(f"📋 {len(rows)} sor másolva (A-P sorrendben, Sheets-be beilleszthető)")
        return len(rows)

    def paste_clipboard(self):
        """Ctrl+V: a vágólap tartalma a kijelölt cellától kezdve (A-P sorrendben).
        Teljes sor beillesztésekor (16 oszlop / R18-kóddal kezdődő sor) az A-B (R18-kód) NEM íródik felül."""
        text = QApplication.clipboard().text()
        if not text or self.currentRow() < 0:
            return 0
        lines = [l for l in text.replace("\r\n", "\n").replace("\r", "\n").split("\n")]
        while lines and lines[-1] == "":
            lines.pop()
        start_r, start_c = self.currentRow(), max(self.currentColumn(), 0)
        if start_c >= self.DATA_COLS:
            start_c = 0
        group = []
        self._edit_source = "📋 Beillesztés (Ctrl+V)"
        self._suppress_undo = True
        try:
            for dr, line in enumerate(lines):
                r = start_r + dr
                if r >= self.rowCount():
                    break
                vals = line.split("\t")
                full_row = len(vals) >= self.DATA_COLS or (start_c == 0 and re.fullmatch(r"R18[-/]\d+", vals[0].strip(), re.I))
                base = 0 if full_row else start_c
                for dc, v in enumerate(vals):
                    c = base + dc
                    if c >= self.DATA_COLS:
                        break
                    if c in (0, 1):          # A-B: az R18-kód védett
                        continue
                    it = self.item(r, c)
                    old = it.text() if it else ""
                    if old == v:
                        continue
                    group.append((r, c, old, v))
                    if it:
                        it.setText(v)
                    else:
                        self.setItem(r, c, QTableWidgetItem(v))
        finally:
            self._edit_source = None
            self._suppress_undo = False
        if group:
            self.undo_stack.append(group)
        return len(group)

    def keyPressEvent(self, event):
        editing = self.state() == QAbstractItemView.State.EditingState

        if not editing and event.matches(QKeySequence.StandardKey.Copy):
            self.copy_selection()
            event.accept()
            return
        if not editing and event.matches(QKeySequence.StandardKey.Paste):
            self.paste_clipboard()
            event.accept()
            return

        # CTRL + Z: UTOLSÓ MŰVELET (CELLAMÓDOSÍTÁS/TÖRLÉS) VISSZAVONÁSA
        if event.key() == Qt.Key.Key_Z and (event.modifiers() & Qt.KeyboardModifier.ControlModifier):
            if self.undo_stack:
                action = self.undo_stack.pop()
                self.blockSignals(True)
                for r, c, old_val, _ in action:
                    item = self.item(r, c)
                    if not item:
                        item = QTableWidgetItem(old_val)
                        self.setItem(r, c, item)
                    else:
                        item.setText(old_val)
                self.blockSignals(False)

                # A szülő ablakban megjelöljük a sorokat mentésre a visszavonás után
                parent_window = self.window()
                if hasattr(parent_window, 'pending_dirty_rows'):
                    for r, c, old_val, new_val in action:
                        parent_window.pending_dirty_rows.add(r)
                        if hasattr(parent_window, "note_excel_source"):
                            parent_window.note_excel_source(r, c, new_val, old_val, "↩️ Ctrl+Z visszavonás")
                if hasattr(parent_window, 'trigger_toast'):
                    parent_window.trigger_toast("↩️ Művelet visszavonva (Ctrl+Z)!")
            event.accept()
            return

        # A nyomtatás-gomb oszlopa (16) képernyőn a legelső: balra nyíllal ne lehessen belelépni
        if (not editing and self.currentColumn() <= 0 and not event.modifiers()
                and event.key() in (Qt.Key.Key_Left, Qt.Key.Key_Home)):
            self.setCurrentCell(max(self.currentRow(), 0), 0)
            event.accept()
            return

        # DELETE / BACKSPACE: kijelölt cellák/sorok törlése + mentés az Undo verembe
        if not editing and event.key() in (Qt.Key.Key_Delete, Qt.Key.Key_Backspace):
            # Teljes sor(ok) kijelölve: megerősítéssel, az R18-kód megmarad (Adat törlés gomb)
            if self.full_selected_rows() and hasattr(self.window(), "clear_selected_excel_rows"):
                self.window().clear_selected_excel_rows()
                event.accept()
                return
            items = self.selectedItems()
            if not items and self.currentItem() is not None:
                items = [self.currentItem()]

            deleted_group = []
            self._edit_source = "⌫ Delete / Backspace gomb (cella törlése)"
            self._suppress_undo = True
            try:
                for item in items:
                    if item.column() > 0 and item.text(): # A oszlop (F-Kód) védve marad
                        deleted_group.append((item.row(), item.column(), item.text(), ""))
                        item.setText("")
            finally:
                self._edit_source = None
                self._suppress_undo = False

            if deleted_group:
                self.undo_stack.append(deleted_group)

            event.accept()
            return

        txt = event.text()
        if (not editing and txt and txt.isprintable() and
                not (event.modifiers() & (Qt.KeyboardModifier.ControlModifier |
                                          Qt.KeyboardModifier.AltModifier |
                                          Qt.KeyboardModifier.MetaModifier))):
            self._typing_edit = True
            try:
                super().keyPressEvent(event)
            finally:
                self._typing_edit = False
            return

        super().keyPressEvent(event)


# -------------------------------------------------------------------
# FŐ ABLAK (PYQT6)
# -------------------------------------------------------------------
def align_excel_item(item, col):
    """Excel Live / Helyi DB cella: M (12) mindig jobbra, L (11) jobbra ha csak szám, különben balra."""
    if item is None:
        return
    if col == 12 or (col == 11 and is_pure_number(item.text())):
        item.setTextAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
    elif col == 11:
        item.setTextAlignment(Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignVCenter)

def _normalize_excel_row(row):
    """Egy A-P adatbázis-sort az Excel Live táblában megjelenő szövegekké alakít (tuple, gyors összevetéshez)."""
    return tuple(str(v) if v is not None and str(v).lower() != "none" else "" for v in row[:16])


class RaktarApp(QMainWindow):
    excel_rows_signal = pyqtSignal(object)  # háttérszálon beolvasott Excel Live sorok -> fő szál
    server_items_signal = pyqtSignal(object)  # a VPS összes sora (Szerver adatbázis fül) -> fő szál
    server_status_signal = pyqtSignal(bool, str)
    counter_update_signal = pyqtSignal(str, int, list)
    process_log_signal = pyqtSignal(str)
    print_result_signal = pyqtSignal(str)   # nyomtatás eredménye (háttérszálból) -> toast
    pda_sales_signal = pyqtSignal(list)     # a Zebra PDA-n rögzített, ez a gép által átvett eladások -> fő szál
    pda_changes_signal = pyqtSignal(list)   # a PDA-n kért módosítások (pl. áthelyezés) -> fő szál
    pda_poll_status_signal = pyqtSignal(str, bool, bool)   # PDA-átvétel állapota: (szöveg, hiba?, kézi indítás?)
    version_info_signal = pyqtSignal(object)   # a szerver szerinti legfrissebb verziók -> fő szál
    clients_signal = pyqtSignal(object)        # gépek és verziók listája (Admin fül) -> fő szál
    leltar_signal = pyqtSignal(str, object)    # leltár szerver-válaszok (fajta, adat) -> fő szál
    sold_pull_signal = pyqtSignal(object, bool)   # eladások behozása a táblázatból: (eredmény, kézi?) -> fő szál
    web_signal = pyqtSignal(str, object)          # webes rendelések (Shopify / eBay): (fajta, adat) -> fő szál
    upd_signal = pyqtSignal(str, object)       # VERZIÓK: frissítés-ellenőrzés / letöltés / pip (fajta, adat) -> fő szál

    def __init__(self):
        super().__init__()
        self.setWindowTitle(f"📦 Raktári Rendszer (PyQt6 Ultra-Fast) – {APP_VERSION}")
        self.zoom_factor = load_saved_zoom_factor()

        self.active_workers = []

        self.server_status_signal.connect(self.update_server_status_ui)
        self.counter_update_signal.connect(self.update_counter_ui)
        self.process_log_signal.connect(self.append_process_log)
        self.print_result_signal.connect(self.trigger_toast)

        main_widget = QWidget()
        self.setCentralWidget(main_widget)
        main_layout = QVBoxLayout(main_widget)

        # FEJLÉC
        header_layout = QHBoxLayout()
        title_label = QLabel("📦 Raktári Rendszer")
        title_label.setFont(QFont("Arial", 18, QFont.Weight.Bold))
        header_layout.addWidget(title_label)

        self.lbl_vps_status = QLabel("⏳ VPS Ellenőrzése...")
        self.lbl_vps_status.setStyleSheet("font-weight: bold; color: #ffb74d; padding: 4px 8px; border-radius: 4px; background: #333;")
        header_layout.addWidget(self.lbl_vps_status)

        header_layout.addStretch()

        zoom_box = QWidget()
        zoom_layout = QHBoxLayout(zoom_box)
        zoom_layout.setContentsMargins(0, 0, 0, 0)
        lbl_zoom = QLabel("🔍 Méret:")
        lbl_zoom.setStyleSheet("font-weight: bold; color: #29b6f6;")
        zoom_layout.addWidget(lbl_zoom)

        btn_zoom_out = QPushButton("-")
        btn_zoom_out.setFixedSize(28, 28)
        btn_zoom_out.clicked.connect(self.zoom_out)
        zoom_layout.addWidget(btn_zoom_out)

        percent_init = int(round(self.zoom_factor * 100))
        self.lbl_zoom_val = QLabel(f"{percent_init}%")
        zoom_layout.addWidget(self.lbl_zoom_val)

        btn_zoom_in = QPushButton("+")
        btn_zoom_in.setFixedSize(28, 28)
        btn_zoom_in.clicked.connect(self.zoom_in)
        zoom_layout.addWidget(btn_zoom_in)

        header_layout.addWidget(zoom_box)

        # 1. Google Sheets Szinkron gomb (1 mp)
        self.btn_quick_live = QPushButton("⚡ Google Sheets Szinkron (1 mp)")
        self.btn_quick_live.setStyleSheet("background-color: #00e676; color: black; font-weight: bold;")
        self.btn_quick_live.clicked.connect(
            lambda: self.run_button_action(
                self.btn_quick_live,
                "⚡ Google Sheets Szinkron (1 mp)",
                self.trigger_fast_live_copy
            )
        )
        header_layout.addWidget(self.btn_quick_live)

        # 2. Helyi Adatok Feltöltése / Hiányzó R18 Adatok Beillesztése gomb
        self.btn_manual_sync = QPushButton("☁️ Helyi Adatok Kézi Feltöltése a Felhőbe")
        self.btn_manual_sync.setStyleSheet("background-color: #ff9800; color: black; font-weight: bold;")
        self.btn_manual_sync.clicked.connect(
            lambda: self.run_button_action(
                self.btn_manual_sync,
                "☁️ Helyi Adatok Kézi Feltöltése a Felhőbe",
                self.trigger_manual_cloud_sync
            )
        )
        header_layout.addWidget(self.btn_manual_sync)
        if not SHEETS_ENABLED:
            self.btn_quick_live.hide()
            self.btn_manual_sync.hide()
        
        # --- ÚJ 3. GOMB: VPS SZERVER SZINKRON (KÉK GOMB) ---
        self.btn_vps_sync = QPushButton("🌐 VPS Szerver Szinkron")
        self.btn_vps_sync.setStyleSheet("background-color: #0288d1; color: white; font-weight: bold; padding: 6px 12px; border-radius: 4px;")
        self.btn_vps_sync.clicked.connect(self.start_vps_sync)
        
        header_layout.addWidget(self.btn_vps_sync)

        # Figyelmeztetés, ha egy másik gépen már frissebb programverzió fut
        self.lbl_version_warn = QLabel("")
        self.lbl_version_warn.setStyleSheet("background-color: #e65100; color: white; font-weight: bold; padding: 5px 10px; border-radius: 4px;")
        self.lbl_version_warn.setToolTip("Egy másik gépen már újabb programverzió fut. Cseréld le ezen a gépen is az app_pyqt6CLOUD.py-t.\n"
                                         "A gépek listája: Admin fül → 🖥️ Gépek és verziók")
        self.lbl_version_warn.hide()
        header_layout.addWidget(self.lbl_version_warn)
        
        # --- 1 MP-ES AUTOMATIKUS LÁTHATATLAN VPS SZINKRON INDÍTÁSA ---
        self.auto_vps_worker = AutoVPSPollWorker()
        self.auto_vps_worker.data_updated_signal.connect(self.on_auto_vps_data_received)
        self.auto_vps_worker.presence_signal.connect(self.on_remote_presence)
        self._remote_edits = {}      # R18 -> {"gep", "ido", "t", "nev"}: másik gépen felvitt / módosított sorok
        self._remote_presence = []   # [{"gep", "A"}]: másik gépen most szerkesztett sorok
        self._remote_printed = set()
        self._remote_marked_rows = set()
        self.auto_vps_worker.start()

        # --- ZEBRA PDA ELADÁSOK ÁTVÉTELE (20 mp-enként) ---
        self.pda_sales_signal.connect(self.apply_pda_sales)
        self.pda_changes_signal.connect(self.apply_pda_changes)
        self.pda_poll_status_signal.connect(self.on_pda_poll_status)
        self.version_info_signal.connect(self.on_version_info)
        self.clients_signal.connect(self.fill_clients_table)
        self.leltar_signal.connect(self.on_leltar_signal)
        print_bus.msg.connect(lambda t: (self.trigger_toast(t, 4500), self.append_process_log(t)))
        self._pda_last_status = None
        self._pda_poll_busy = False
        self.pda_poll_timer = QTimer(self)
        self.pda_poll_timer.timeout.connect(self.poll_pda_sales)
        self.pda_poll_timer.start(PDA_SALES_POLL_MS)
        QTimer.singleShot(5000, self.poll_pda_sales)

        self.close_btn = QPushButton("🚪 Bezárás")
        self.close_btn.setStyleSheet("background-color: #ff5252; color: white;")
        self.close_btn.clicked.connect(self.close)
        header_layout.addWidget(self.close_btn)

        main_layout.addLayout(header_layout)

        # ⬆️ ÚJ VERZIÓ – nagy, villogó sáv (kattintásra a VERZIÓK fülre visz)
        self.update_banner = QPushButton("")
        self.update_banner.setCursor(Qt.CursorShape.PointingHandCursor)
        self.update_banner.clicked.connect(lambda: self.tabs.setCurrentWidget(self.tab_versions))
        self.update_banner.hide()
        main_layout.addWidget(self.update_banner)
        self._banner_blink = False
        self.banner_timer = QTimer(self)
        self.banner_timer.timeout.connect(self._blink_banner)

        # FÜLEK LÉTREHOZÁSA ÉS SORRENDJE (Excel Live Munkalap a legelső!)
        self.tabs = QTabWidget()
        # A fülek tartalma ne szabja az ablak legkisebb méretét a képernyőnél nagyobbra
        self.tabs.setMinimumSize(700, 300)
        main_layout.addWidget(self.tabs)

        # 0. Duplikált R18-számok összevonása (mielőtt bármi betöltődik)
        removed = dedupe_local_products()
        self._startup_dedupe_count = len(removed)

        # 1. Először létrehozzuk az Excel Live Munkalap fület legelöl
        self.setup_excel_style_tab()

        # 2. Utána hozzuk létre a többi fület
        self.tab_search = QWidget()
        self.tab_add = QWidget()
        self.tab_local_db = QWidget()
        self.tab_pda = QWidget()
        self.tab_sold = QWidget()
        self.tab_leltar = QWidget()
        self.tab_server = QWidget()
        self.tab_excel = QWidget()
        self.tab_single = QWidget()
        self.tab_map = QWidget()
        self.tab_log = QWidget()
        self.tab_process_log = QWidget()
        self.tab_admin = QWidget()
        self.tab_versions = QWidget()
        self.tab_web = QWidget()

        # Fülek hozzáadása a kívánt sorrendben
        # FŐ FÜLEK (mindig látszanak): Excel Live, Gyorskereső, PDA / Eladás, Eladottak
        self.tabs.addTab(self.tab_search, "⚡ Gyorskereső")
        self.tabs.addTab(self.tab_pda, "📱 PDA / Eladás")
        self.tabs.addTab(self.tab_sold, "💰 Eladottak")
        if WEBES_RENDELESEK_BE:
            self.tabs.addTab(self.tab_web, "🛒 Webes rendelések")
        self.tabs.addTab(self.tab_leltar, "📋 Leltár")
        self.tabs.addTab(self.tab_versions, "🆕 VERZIÓK")
        # TOVÁBBI FÜLEK (a "☰ További fülek" menüből nyílnak)
        self.tabs.addTab(self.tab_add, "📝 Új Termék Felvitele")
        self.tabs.addTab(self.tab_local_db, "💾 Helyi Adatbázis (Élő)")
        self.tabs.addTab(self.tab_server, "🖥️ Szerver adatbázis")
        self.tabs.addTab(self.tab_excel, "📊 Excel Kezelés")
        self.tabs.addTab(self.tab_single, "🖨️ Egyedi Címke")
        self.tabs.addTab(self.tab_map, "🗺️ Raktártérkép")
        self.tabs.addTab(self.tab_log, "📜 Napló")
        self.tabs.addTab(self.tab_process_log, "⚙️ Folyamat & Szinkron Napló")
        self.tabs.addTab(self.tab_admin, "🛡️ Admin")

        # Fülek tartalmának inicializálása
        self.init_search_tab()
        self.init_add_tab()
        self.init_local_db_tab()
        self.init_pda_tab()
        self.init_sold_tab()
        self.init_leltar_tab()
        self.init_versions_tab()
        if WEBES_RENDELESEK_BE:
            self.init_web_tab()
        self.init_server_tab()
        self.init_excel_tab()
        self.init_single_tab()
        self.init_map_tab()
        self.init_log_tab()
        self.init_process_log_tab()
        self.init_admin_tab()
        self.setup_tab_menu()

        # 📷 Termékfotók: a fotómappa beolvasása háttérben, 10 percenként frissítve
        photo_bus.index_ready.connect(self._on_photo_index_ready)
        photo_index.build_async(force=True)
        self.photo_timer = QTimer()
        self.photo_timer.setInterval(PHOTO_INDEX_REFRESH_SEC * 1000)
        self.photo_timer.timeout.connect(lambda: photo_index.build_async(force=True))
        self.photo_timer.start()

        # GARANTÁLTAN AZ 1. FÜLRE (Excel Live Munkalap) FÓKUSZÁL INDÍTÁSKOR:
        self.tabs.setCurrentIndex(0)
        # Elhalasztott (nem látható fülre szóló) frissítések elvégzése fülváltáskor
        self.tabs.currentChanged.connect(self.on_tab_changed)

        # Google táblázat L / M igazításának EGYSZERI javítása a régi sorokon (háttérben)
        if SHEETS_ENABLED and not load_config_value("sheets_igazitas_javitva_v1"):
            def _fix_align():
                try:
                    if fix_sheet_alignment_all(lambda m, t="INFO": self.process_log_signal.emit(m)):
                        save_config_value("sheets_igazitas_javitva_v1", True)
                except Exception as e:
                    self.process_log_signal.emit(f"⚠️ Google táblázat igazítás javítása sikertelen (később újra próbálja): {e}")
            QTimer.singleShot(20000, lambda: threading.Thread(target=_fix_align, daemon=True).start())

        # TÖRÖLT SOROK FIGYELÉSE a figyelt (élő) táblázatban - háttérszálon, csak olvasással
        if DELETED_WATCH_SPREADSHEET_ID.strip() and SHEETS_ENABLED:
            def _watch_log(msg, log_type="INFO"):
                self.process_log_signal.emit(msg)
            self.deleted_watch_timer = QTimer()
            self.deleted_watch_timer.setInterval(DELETED_WATCH_INTERVAL_SEC * 1000)
            self.deleted_watch_timer.timeout.connect(
                lambda: threading.Thread(target=check_watched_sheet_for_deletions, args=(_watch_log,), daemon=True).start())
            self.deleted_watch_timer.start()
            QTimer.singleShot(5000, lambda: threading.Thread(target=check_watched_sheet_for_deletions,
                                                             args=(_watch_log,), daemon=True).start())

        self.update_app_stylesheet()

        self.check_vps_connection()
        self.refresh_counters()

        QTimer.singleShot(10, self.load_latest_200_items)
        if SHEETS_ENABLED:
            QTimer.singleShot(300, self.trigger_fast_live_copy)
        
        # 15 másodpercenkénti automatikus felhő-szinkronizáció (Google Sheets -> Helyi DB)
        if SHEETS_ENABLED:
            self.sheets_pull_timer = QTimer()
            self.sheets_pull_timer.setInterval(15000) # 15000 ms = 15 mp
            self.sheets_pull_timer.timeout.connect(self.trigger_fast_live_copy)
            self.sheets_pull_timer.start()

    def reset_item_keep_code(self, r18_code):
        dlg = SilentConfirmDialog(
            "Termék adatok kiürítése",
            f"Biztosan kiüríted a(z) {r18_code} kódhoz tartozó ÖSSZES termékadatot?\n\n(A márka, típus, cikkszám, darabszám és állapot is törlődik, kizárólag a(z) {r18_code} kód marad meg!)",
            self
        )
        if dlg.exec() != QDialog.DialogCode.Accepted:
            return

        try:
            elotte = get_row_snapshot(r18_code)
            conn = get_db_connection()
            cursor = conn.cursor()
            cursor.execute("""
                UPDATE termekek 
                SET C='', D='', E='', F='', G='', H='', I='', J='', K='', L='', M='', N='', O='', P='', szin='FEHÉR'
                WHERE A=? OR B=?
            """, (r18_code, r18_code))
            conn.commit()
            conn.close()
            log_event(r18_code, "TÖRLÉS (adatok kiürítve, R18 megmarad)", "", elotte, get_row_snapshot(r18_code),
                      "Kártya - ↺ Adatok kiürítése gomb")

            # Felhő (VPS + Sheets) az outboxon át: ha nincs net, később automatikusan kimegy
            enqueue_row_from_db(r18_code, force=True)
            enqueue_color(r18_code, "FEHÉR", COLOR_WHITE)
            flush_outbox_async()

            self.trigger_toast(f"↺ {r18_code} adatai (DB és Állapot is) teljesen kiürítve!")
            self.append_process_log(f"↺ {r18_code} adatai teljesen kiürítve (R18 kód megtartva)", "#ff9800")

            search_text = self.search_input.text().strip()
            if search_text and search_text.upper() != "R18-":
                self.execute_search()
            else:
                self.load_latest_200_items()

            self.refresh_counters()

        except Exception as e:
            self.trigger_toast(f"❌ Hiba a kiürítés során: {e}")
            
    def start_vps_sync(self):
            """Új kék gomb: Teljes manuális letöltés a VPS-ről"""
            self.btn_vps_sync.setEnabled(False)
            self.trigger_toast("🌐 VPS Szinkronizáció indítása...")
    
            self.vps_worker = VPSDownloadWorker()
            self.vps_worker.process_log_signal.connect(self.append_process_log)
            self.vps_worker.finished_signal.connect(self.on_vps_sync_finished)
            self.vps_worker.start()
            
        
    def setup_excel_style_tab(self):
        """Excel Live Munkalap: Adatok betöltése az SQLite-ból és fókusz az első hiányos (C oszlop üres) sorra"""
        self.tab_excel_live = QWidget()
        layout = QVBoxLayout(self.tab_excel_live)

        # Tájékoztató sáv
        info_label = QLabel("💡 A rendszer automatikusan az első kitöltendő (C oszlop üres) sorhoz ugrott. Nyilakkal/Tabbal navigálhatsz.")
        info_label.setStyleSheet("color: #b0bec5; font-size: 13px; margin-bottom: 5px;")
        layout.addWidget(info_label)

        # Excel-szerű táblázat
        self.excel_table = ExcelTableWidget()
        self.excel_table.setColumnCount(17)
        headers = ["A: F-Kód", "B: Vonalkód", "C: Márka", "D: Típus", "E: Cikkszám", "F: DB", 
                   "G: Állapot", "H: Ár", "I: Páter", "J: Szabad", "K: Polc", "L: Raklap", 
                   "M: Súly", "N: eBay ID", "O: Megjegyzés", "P: Állapot részlet", "🖨️ Akció"]
        self.excel_table.setHorizontalHeaderLabels(headers)

        # Adatok lekérése a helyi adatbázisból (R18- számsorrendben csökkenő)
        conn = get_db_connection()
        cursor = conn.cursor()
        cursor.execute("""
            SELECT A, B, C, D, E, F, G, H, I, J, K, L, M, N, O, P 
            FROM termekek 
            ORDER BY 
                CASE 
                    WHEN A LIKE 'R18-%' THEN CAST(SUBSTR(A, 5) AS INTEGER)
                    ELSE 0 
                END DESC,
                rowid DESC
        """)
        rows = cursor.fetchall()
        conn.close()

        self.excel_table.setRowCount(len(rows))
        
        first_unfilled_row = -1  # Az első olyan sor, ahol C (Márka) üres

        self.excel_table.blockSignals(True) # Szignálok letiltása a betöltés idejére

        for r_idx, row in enumerate(rows):
            # A-P oszlopok feltöltése
            for c_idx in range(16):
                val = str(row[c_idx]) if row[c_idx] is not None else ""
                item = QTableWidgetItem(val)
                self.excel_table.setItem(r_idx, c_idx, item)

            # Ellenőrizzük, hogy a C oszlop (index 2) üres-e
            c_val = str(row[2]).strip() if row[2] is not None else ""
            if not c_val and first_unfilled_row == -1:
                first_unfilled_row = r_idx

            # Sorvégi 🖨️ Nyomtatás gomb beágyazása
            btn_print = QPushButton("🖨️ Nyomtatás")
            btn_print.setStyleSheet("background-color: #0288d1; color: white; font-weight: bold; padding: 2px 6px; border-radius: 3px;")
            btn_print.clicked.connect(lambda ch, row=r_idx: self.print_row_label(row))
            self.excel_table.setCellWidget(r_idx, 16, btn_print)

        self.excel_table.blockSignals(False)

        # Cella módosítás esemény bekötése
        self.excel_table.itemChanged.connect(self.on_excel_cell_changed)
        layout.addWidget(self.excel_table)

        # AUTOMATIKUS UGRÁS ÉS FÓKUSZ az első olyan sorra, ahol a C oszlop még üres!
        if first_unfilled_row != -1:
            self.excel_table.setCurrentCell(first_unfilled_row, 2) # Fókusz a C oszlopra
            self.excel_table.scrollToItem(self.excel_table.item(first_unfilled_row, 2))
            self.excel_table.horizontalScrollBar().setValue(0)

        # Módosított sorok gyűjtője
        self.pending_dirty_rows = set()

        # 4 mp-es időzítő az automatikus mentéshez
        self.auto_save_timer = QTimer()
        self.auto_save_timer.setInterval(AUTO_SAVE_INTERVAL_MS)
        self.auto_save_timer.timeout.connect(self.process_auto_row_sync)
        self.auto_save_timer.start()

        # Új lapfül hozzáadása
        #if hasattr(self, 'tabs'):
        #    self.tabs.addTab(self.tab_excel_live, "📊 Excel Live Munkalap")
    
    
    
    def on_auto_vps_data_received(self, updated_items):
        """Háttérben beérkező VPS frissítések feldolgozása zéró fagyással"""
        try:
            conn = get_db_connection()
            cursor = conn.cursor()

            # Amit ez a gép még nem töltött fel (outbox-ban vár), azt NEM írjuk felül a
            # VPS régebbi állapotával - a helyi módosítás úgyis felmegy és az nyer.
            pending_codes = get_pending_outbox_codes(cursor, csak_vps=True)

            cols = list(ROW_COLS) + ["szin"]
            changed = 0
            naplo_entries = []   # a tranzakció lezárása után írjuk a naplóba
            for item in updated_items:
                code = (safe_str(item.get("A")) or safe_str(item.get("B"))).upper()
                if not code or code in pending_codes:
                    continue
                gep = safe_str(item.get("_gep"))
                if gep and gep != GEP_NEV and not item.get("_torolt") and any(safe_str(item.get(c)) for c in "CDE"):
                    self._remote_edits[code] = {"gep": gep, "ido": safe_str(item.get("_ido"))[11:16] or datetime.datetime.now().strftime("%H:%M"),
                                                "t": time.time(), "nev": " ".join(safe_str(item.get(c)) for c in "CD" if safe_str(item.get(c)))}
                    self._remote_printed.discard(code)

                # A másik gépen törölt (eladott) tétel
                if item.get("_torolt"):
                    elotte = get_row_snapshot(code, cursor)
                    cursor.execute("DELETE FROM termekek WHERE UPPER(TRIM(A))=? OR UPPER(TRIM(B))=?", (code, code))
                    if cursor.rowcount:
                        naplo_entries.append((code, "TÖRLÉS (másik gépen)", elotte, None))
                    changed += cursor.rowcount
                    continue
                elotte = get_row_snapshot(code, cursor)

                vals = [safe_str(item.get(c)) for c in ROW_COLS]
                szin = safe_str(item.get("szin")).upper()
                vals.append("FEHÉR" if szin in ("", "FEHER") else szin)

                # Helyi adatbázis frissítése csendben
                cursor.execute(
                    f"SELECT rowid, {', '.join(cols)} FROM termekek WHERE UPPER(TRIM(A))=? OR UPPER(TRIM(B))=?", (code, code)
                )
                existing = cursor.fetchone()
                if existing:
                    old_vals = [safe_str(existing[c]) for c in ROW_COLS]
                    old_vals.append(safe_str(existing["szin"]).upper() or "FEHÉR")
                    if old_vals == vals:
                        continue  # a saját módosításunk visszhangja, nincs teendő
                    cursor.execute(
                        f"UPDATE termekek SET {', '.join(c + '=?' for c in cols)} WHERE rowid=?",
                        (*vals, existing["rowid"])
                    )
                else:
                    cursor.execute(
                        f"INSERT OR IGNORE INTO termekek ({', '.join(cols)}) VALUES ({', '.join('?' * len(cols))})",
                        vals
                    )
                changed += 1
                utana = get_row_snapshot(code, cursor)
                if elotte != utana:
                    kind = ("FELVÉTEL (másik gépen)" if not (elotte and _row_has_product_data(elotte)) and _row_has_product_data(utana)
                            else "TÖRLÉS (másik gépen kiürítve)" if elotte and _row_has_product_data(elotte) and not _row_has_product_data(utana)
                            else "MÓDOSÍTÁS (másik gépen)")
                    naplo_entries.append((code, kind, elotte, utana))

            conn.commit()
            conn.close()
            for code, kind, elotte, utana in naplo_entries:
                log_event(code, kind, describe_changes(elotte, utana) if utana is not None else "A termék sora törölve a másik gépen",
                          elotte, utana, "🌐 Másik gép - VPS élő szinkron")

            if changed:
                self.append_process_log(f"🔄 VPS élő szinkron: {changed} tétel frissítve a másik gépről")
                # CSANAK A KÁRTYÁKAT FRISSÍTJÜK (Nem építjük újra a teljes táblázatot -> ZÉRÓ FAGYÁS!)
                self.update_cards_only()
                # Az Excel Live-on is azonnal látszódjon, mit vittek fel a másik gépen (csak a változott sorok íródnak át)
                self.silent_update_excel_table()
            self.render_remote_marks()

        except Exception as e:
            print(f"Auto VPS frissítési hiba: {e}")
    
    

    
    
    
    

    def on_vps_sync_finished(self, success, msg):
        self.btn_vps_sync.setEnabled(True)
        self.trigger_toast(msg)
        if success:
            # Táblázat és kártyák frissítése a letöltött VPS adatokkal
            self.refresh_local_db_table_view()
            self.update_cards_only()
            
    def on_local_table_cell_edited(self, item):
        """Cella módosítása a Helyi DB táblázatban – Azonnali mentés SQLite-ba és automatikusan a VPS-re"""
        if not item:
            return

        row_idx = item.row()
        col_idx = item.column()

        if col_idx == 0:
            return

        code_item = self.local_db_table.item(row_idx, 0)
        if not code_item:
            return

        r18_code = code_item.text().strip().upper()
        if not r18_code or not r18_code.startswith("R18-"):
            return

        col_letters = ["A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P"]
        target_col_letter = col_letters[col_idx]
        new_value = item.text().strip()

        try:
            # 1. Helyi SQLite frissítése
            elotte = get_row_snapshot(r18_code)
            conn = get_db_connection()
            cursor = conn.cursor()
            query = f"UPDATE termekek SET {target_col_letter} = ? WHERE A = ? OR B = ?"
            cursor.execute(query, (new_value, r18_code, r18_code))
            conn.commit()
            conn.close()

            # 2. FELHŐ (VPS + Sheets) az outboxon át - ha nincs net, később automatikusan kimegy.
            # force: szándékos javítás, a VPS ne vegye R18-ütközésnek.
            enqueue_row_from_db(r18_code, force=True)
            flush_outbox_async()

            # 3. Cella igazítása helyben (I, K, M mindig jobbra, L csak ha szám)
            self.local_db_table.blockSignals(True)
            if col_idx in (8, 10, 12): 
                item.setTextAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
            elif col_idx == 11: 
                if new_value and re.match(r'^\d+([.,]\d+)?$', new_value):
                    item.setTextAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
                else:
                    item.setTextAlignment(Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignVCenter)
            self.local_db_table.blockSignals(False)

            utana = get_row_snapshot(r18_code)
            if elotte != utana:
                log_event(r18_code, "MÓDOSÍTÁS (Helyi adatbázis)", describe_changes(elotte, utana), elotte, utana,
                          f"Helyi adatbázis fül - cella szerkesztés billentyűzettel ({target_col_letter} oszlop)")
            self.trigger_toast(f"💾 {r18_code} [{target_col_letter}] frissítve (Local + VPS)!")

            self.update_cards_only()

        except Exception as e:
            self.trigger_toast(f"❌ Hiba a cella mentésekor: {e}")
            
    def update_cards_only(self):
        """Kizárólag a Gyorskereső kártyáit frissíti újra a Helyi DB táblázat bántása nélkül"""
        if not self._tab_is_current("tab_search"):
            self._cards_dirty = True
            return
        self._cards_dirty = False
        try:
            self.results_area.setUpdatesEnabled(False)
            while self.results_layout.count():
                item = self.results_layout.takeAt(0)
                widget = item.widget()
                if widget:
                    widget.deleteLater()

            info_lbl = QLabel("💡 Pittyents a pisztollyal a kereséshez, vagy nyomtass közvetlenül az utoljára felvitt alkatrészek közül:")
            info_lbl.setStyleSheet("background-color: #1565c0; color: white; padding: 10px; border-radius: 6px; font-size: 15px; font-weight: bold;")
            self.results_layout.addWidget(info_lbl)

            header_lbl = QLabel("🆕 Legutóbb felvitt alkatrészek (A legfrissebb felül):")
            header_lbl.setStyleSheet("font-size: 16px; font-weight: bold; margin-top: 10px; color: #29b6f6;")
            self.results_layout.addWidget(header_lbl)

            conn = get_db_connection()
            query = """
                SELECT * FROM termekek 
                WHERE C IS NOT NULL AND TRIM(C) != '' AND LOWER(TRIM(C)) != 'none' AND LOWER(TRIM(C)) != 'nan'
            """
            cursor = conn.cursor()
            cursor.execute(query)
            rows = cursor.fetchall()
            conn.close()

            filled_c_items = []
            for row in rows:
                row_dict = dict(row)
                code = get_best_code(row_dict)
                m = re.search(r'R18-(\d+)', str(code), re.IGNORECASE)
                if m:
                    num = int(m.group(1))
                    if 0 < num < 100000:
                        filled_c_items.append((num, row_dict))

            filled_c_items.sort(key=lambda x: x[0], reverse=True)

            seen_codes = set()
            valid_count = 0
            
            for num, row_dict in filled_c_items:
                code_raw = get_best_code(row_dict)
                if code_raw in seen_codes:
                    continue
                seen_codes.add(code_raw)
                
                self.results_layout.addWidget(self.create_item_card(row_dict))
                valid_count += 1
                if valid_count == 15:
                    break

            if valid_count == 0:
                self.results_layout.addWidget(QLabel("Nincs megjeleníthető friss alkatrész."))
        except Exception as e:
            print(f"Hiba a kártyák frissítésekor: {e}")
        finally:
            self.results_area.setUpdatesEnabled(True)

    def print_add_label(self):
        code = getattr(self, 'last_added_code', None) or self.input_a.text().strip().upper()
        cikk = getattr(self, 'last_added_cikk', None) or self.input_e.text().strip().upper()
        tipus = getattr(self, 'last_added_tipus', None) or self.input_d.text().strip().upper()

        if not code or code == "R18-":
            self.trigger_toast("⚠️ Nincs nyomtatható termék adat!")
            return

        is_zpl = self.radio_add_zpl.isChecked() if hasattr(self, 'radio_add_zpl') else True
        copies = self.spin_add_copies.value() if hasattr(self, 'spin_add_copies') else 1

        zebra_printer = get_zebra_printer_name()

        if is_zpl:
            zpl = generate_zpl_code(code, cikkszam=cikk, tipus=tipus, copies=copies)
            ok, msg = send_zpl_to_printer(zebra_printer, zpl)
            if ok:
                self.trigger_toast(f"🖨️ ZPL Címke kinyomtatva [{code}]!")
            else:
                self.trigger_toast(f"❌ Nyomtatási hiba: {msg}")
        else:
            _, bar_bytes = generate_codes_in_memory(code, cikkszam_text=cikk, tipus_text=tipus)
            if bar_bytes:
                ok, msg = send_png_bytes_to_printer(zebra_printer, bar_bytes, copies)
                if ok:
                    self.trigger_toast(f"🖨️ PNG Címke kinyomtatva [{code}]!")
                else:
                    self.trigger_toast(f"❌ Nyomtatási hiba: {msg}")

    def run_button_action(self, button, original_text, action_func, *args, **kwargs):
        try:
            button.setText("⏳ Processing...")
            button.setEnabled(False)
            QApplication.processEvents()
            
            action_func(*args, **kwargs)
        except Exception as e:
            print(f"Hiba a gomb művelet futtatása közben: {e}")
            raise e
        finally:
            try:
                if button and not button.signalsBlocked():
                    button.setText(original_text)
                    button.setEnabled(True)
            except RuntimeError:
                pass

    def trigger_manual_cloud_sync(self):
        """A kiválasztott legújabb sorok TELJES A-P adatainak kézi feltöltése a felhőbe"""
        if not SHEETS_ENABLED:
            self.trigger_toast("ℹ️ A Google Sheets szinkron ki van kapcsolva - az adatok a helyi adatbázisban és a VPS-en vannak.")
            return
        try:
            conn = get_db_connection()
            cursor = conn.cursor()
            cursor.execute("""
                SELECT *,
                       CAST(SUBSTR(A, 5) AS INTEGER) as num_code
                FROM termekek
                WHERE C IS NOT NULL 
                  AND TRIM(C) != '' 
                  AND LOWER(TRIM(C)) != 'none' 
                  AND LOWER(TRIM(C)) != 'nan'
                ORDER BY num_code DESC
            """)
            local_rows = cursor.fetchall()
            conn.close()

            if not local_rows:
                QMessageBox.information(self, "Információ", "Nincs kitöltött tétel az adatbázisban.")
                return

            latest_valid_code = local_rows[0]["A"]

            count, ok = QInputDialog.getInt(
                self, 
                "Teljes Soros Felhő Szinkronizáció", 
                f"Legfrissebb alkatrész: {latest_valid_code}\n\n"
                f"Összesen {len(local_rows)} kitöltött elem van.\nHány legújabb sort töltsünk fel A-tól P-ig a felhőbe?", 
                value=10, min=1, max=len(local_rows)
            )

            if not ok:
                return

            selected_local = local_rows[:count]
            first_code = selected_local[0]["A"]
            last_code = selected_local[-1]["A"]

            confirm = QMessageBox.question(
                self,
                "Feltöltés Megerősítése",
                f"Biztosan feltöltöd a felhőbe ezen {count} elem TELJES adatát (A-P oszlopok)?\n\n"
                f"📌 Tartomány: {first_code} ➔ {last_code}",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
            )

            if confirm != QMessageBox.StandardButton.Yes:
                return

            # Háttérszál indítása (Fagyásmentesen frissít A-tól P-ig mindent)
            self.manual_sync_worker = PreciseManualSyncWorker(selected_local)
            self.manual_sync_worker.process_log_signal.connect(self.handle_log_from_thread)
            self.manual_sync_worker.finished_signal.connect(self.on_precise_sync_finished)
            self.manual_sync_worker.finished.connect(lambda: self.track_worker_finished(self.manual_sync_worker))
            self.active_workers.append(self.manual_sync_worker)
            self.manual_sync_worker.start()

            self.trigger_toast("⚡ Teljes soros felhőfeltöltés elindítva...")

        except Exception as e:
            QMessageBox.critical(self, "Hiba", f"Hiba az indítás során: {str(e)}")

    def on_precise_sync_finished(self, success, message):
        if success:
            self.trigger_toast(message)
            QMessageBox.information(self, "Siker", message)
        else:
            QMessageBox.critical(self, "Hiba", message)

    def on_manual_sync_finished(self, msg):
        self.trigger_toast(msg)
        self.refresh_counters()
        self.load_latest_200_items()
        
        
    def force_save_excel_table(self):
        """Kézi gombnyomásra végigpásztázza a táblázatot és elment minden kitöltött sort"""
        col_keys = ["A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P"]
        dirty_data = []

        for r_idx in range(self.excel_table.rowCount()):
            row_dict = {}
            for col_idx, col_name in enumerate(col_keys):
                item = self.excel_table.item(r_idx, col_idx)
                row_dict[col_name] = item.text().strip() if item else ""

            # Csak azokat a sorokat mentjük, ahol van F-kód ÉS van legalább Márka vagy Típus beírva!
            if row_dict.get("A") and (row_dict.get("C") or row_dict.get("D")):
                notes = self._excel_row_sources.pop(r_idx, [])
                row_dict["_forras"] = "Excel Live - 💾 MENTÉS MOST gomb" + (" | " + " | ".join(notes) if notes else "")
                dirty_data.append(row_dict)

        if dirty_data:
            self.trigger_toast(f"⏳ {len(dirty_data)} sor mentése folyamatban...")
            self.row_sync_worker = AutoRowSyncWorker(dirty_data)
            self.row_sync_worker.sync_finished_signal.connect(self.on_auto_row_sync_finished)
            self.row_sync_worker.start()
        else:
            self.trigger_toast("⚠️ Nincs menthető/módosított adat a táblázatban!")
    

    def log_system_heartbeat(self):
        vps_status = self.lbl_vps_status.text()
        self.append_process_log(f"💓 [ÉLETJEL] Háttérfolyamatok aktívak | Szinkron fut | VPS Állapot: {vps_status}", "#29b6f6")

    def trigger_fast_live_copy(self):
        if not SHEETS_ENABLED:
            return
        if hasattr(self, 'fast_copy_worker') and self.fast_copy_worker.isRunning():
            return

        # Összegyűjtjük az éppen szerkesztett, még ki nem mentett sorokat
        pending_data = {}
        if hasattr(self, 'pending_dirty_rows') and hasattr(self, 'excel_table'):
            for r_idx in self.pending_dirty_rows:
                item_a = self.excel_table.item(r_idx, 0)
                item_c = self.excel_table.item(r_idx, 2)
                item_d = self.excel_table.item(r_idx, 3)
                if item_a and item_a.text().strip():
                    code = item_a.text().strip().upper()
                    pending_data[code] = {
                        "C": item_c.text().strip() if item_c else "",
                        "D": item_d.text().strip() if item_d else ""
                    }

        self.fast_copy_worker = FastCopyLiveToTestWorker(pending_rows_data=pending_data)
        self.fast_copy_worker.finished_signal.connect(self.on_fast_copy_finished)
        self.fast_copy_worker.process_log_signal.connect(self.handle_log_from_thread)
        self.fast_copy_worker.finished.connect(lambda: self.track_worker_finished(self.fast_copy_worker))
        self.active_workers.append(self.fast_copy_worker)
        self.fast_copy_worker.start()

    def on_fast_copy_finished(self, success, message, elapsed, conflicts):
        # A Sheets nem változott: nincs mit frissíteni a felületen
        if success and getattr(self.sender(), "unchanged", False):
            return
        if success and not getattr(self, "pending_dirty_rows", None):
            # Biztonsági háló: ha mégis kétszer került be egy R18-szám, összevonjuk (csak amíg nem gépelnek)
            removed = dedupe_local_products()
            if removed:
                self.append_process_log(f"🧹 {len(removed)} duplikált R18-sor összevonva: "
                                        f"{', '.join(sorted({c for c, _, _ in removed})[:15])} (részletek a 📜 Naplóban)", "#ffb74d")
        if success:
            # HA ÜTKÖZÉS VAN A GOOGLE SHEETS ÉS A TE BEÍRÁSOD KÖZÖTT:
            # Ugyanarra a sorra 2 percen belül csak egyszer szólunk
            now_t = time.monotonic()
            if not hasattr(self, "_conflict_notified"):
                self._conflict_notified = {}
            conflicts = [cc for cc in conflicts if now_t - self._conflict_notified.get(cc, -1e9) > 120]
            if conflicts:
                for cc in conflicts:
                    self._conflict_notified[cc] = now_t
                conflict_codes = ", ".join(conflicts)
                # Nem felugró ablak: csak értesítés, nem kell rákattintani
                self.trigger_toast(f"⚠️ Valaki más is ír ebbe a sorba: {conflict_codes} - ellenőrizd mentés után!", duration=6000)
                self.append_process_log(f"⚠️ Párhuzamos módosítás: a Google táblázatban is átírták, miközben itt "
                                        f"szerkesztették: {conflict_codes}", "#ffb74d")
                for cc in conflicts:
                    log_event(cc, "FIGYELMEZTETÉS: SORKONFLIKTUS", "Szerkesztés közben a Google táblázatban is átírták",
                              None, get_row_snapshot(cc), "📊 Google Sheets + Excel Live egyszerre")

            # TELJES ÚJRAÉPÍTÉS HELYETT CSENDES, NEM-BLOKKOLÓ FRISSÍTÉS:
            if hasattr(self, 'silent_update_excel_table'):
                self.silent_update_excel_table()
                
            self.refresh_counters()
            
    def silent_update_excel_table(self):
        """Csendes háttérfrissítés: az adatbázist HÁTTÉRSZÁLON olvassa, a felületen pedig csak a ténylegesen
        megváltozott sorokat írja át (nem veszi el a fókuszt, nem szakítja meg a gépelést, nem akaszt)."""
        if not hasattr(self, 'excel_table'):
            return
        if not self.excel_table.isVisible():
            self._excel_dirty = True      # majd ha a fülre váltanak (on_tab_changed)
            return
        if self._excel_loading:
            return
        self._excel_loading = True
        self._excel_dirty = False

        def job():
            rows = None
            try:
                conn = sqlite3.connect(DB_FILE, timeout=10)
                cursor = conn.cursor()
                cursor.execute("""
                    SELECT A, B, C, D, E, F, G, H, I, J, K, L, M, N, O, P
                    FROM termekek
                    WHERE A LIKE 'R18-%'
                    ORDER BY CAST(SUBSTR(A, 5) AS INTEGER) ASC
                """)
                rows = [_normalize_excel_row(r) for r in cursor.fetchall()]
                conn.close()
            except Exception as e:
                print(f"Csendes frissítési hiba: {e}")
            self.excel_rows_signal.emit(rows)

        threading.Thread(target=job, daemon=True).start()

    def _apply_excel_rows(self, rows):
        """(fő szál) A háttérben beolvasott sorokat összeveti az előzővel és csak a változást írja ki."""
        self._excel_loading = False
        if rows is None:
            return

        current_row = self.excel_table.currentRow()
        # Csak azt a sort nem bántjuk, amibe ezen a gépen ÉPPEN gépelnek (nyitott cellaszerkesztő vagy még nem mentett
        # módosítás). Ha csak rajta áll a kijelölés, frissítjük (különben a másik gépen felvitt sor üres maradna).
        active_dirty_rows = getattr(self, 'pending_dirty_rows', set())
        editing = self.excel_table.state() == QAbstractItemView.State.EditingState
        self._excel_deferred = False
        old = self._excel_rows_cache

        self.excel_table.blockSignals(True)
        if len(rows) != self.excel_table.rowCount():
            self.excel_table.setRowCount(len(rows))

        new_cache = list(rows)
        for r_idx, row in enumerate(rows):
            if r_idx < len(old) and old[r_idx] == row:
                continue
            if r_idx in active_dirty_rows or (r_idx == current_row and editing):
                # a régi állapot marad a cache-ben: a következő körben újra összeveti (és az R18-index sem veszíti el)
                new_cache[r_idx] = old[r_idx] if r_idx < len(old) and old[r_idx] else None
                self._excel_deferred = True
                continue
            for c_idx in range(16):
                val = row[c_idx]
                item = self.excel_table.item(r_idx, c_idx)
                if item:
                    if item.text() != val:
                        item.setText(val)
                else:
                    item = QTableWidgetItem(val)
                    self.excel_table.setItem(r_idx, c_idx, item)
                if c_idx in (11, 12):
                    align_excel_item(item, c_idx)
        self.excel_table.blockSignals(False)

        if len(new_cache) != len(old) or any(a is None or a[0] != (b[0] if b else None) for a, b in zip(new_cache, old)):
            self._excel_rows_cache = new_cache
            self._rebuild_excel_code_index()
        else:
            self._excel_rows_cache = new_cache
        self.render_remote_marks()

    def update_print_selected_label(self, row, *_):
        if not hasattr(self, "btn_print_selected"):
            return
        item = self.excel_table.item(row, 0) if row >= 0 else None
        code = item.text().strip() if item and item.text().strip() else "–"
        self.btn_print_selected.setText(f"🖨️ Nyomtatás: {code} (Ctrl+P)")

    # ---- 🆕 Másik gépen felvitt sorok / ki min dolgozik most ----
    def on_remote_presence(self, lst):
        self._remote_presence = lst or []
        self.render_remote_marks()

    def clear_remote_marks(self):
        self._remote_edits.clear()
        self._remote_printed.clear()
        self.render_remote_marks()

    def _presence_tick(self):
        if not hasattr(self, "excel_table"):
            return
        code = ""
        t = self.excel_table
        if self._tab_is_current("tab_excel_live") and self.isActiveWindow():
            editing = t.state() == QAbstractItemView.State.EditingState
            if editing or time.time() - self._last_excel_edit < 20:
                r = t.currentRow()
                it = t.item(r, 0) if r >= 0 else None
                code = it.text().strip().upper() if it else ""
        now = time.time()
        if code == self._presence_sent and (not code or now - self._presence_ts < 10):
            return
        self._presence_sent, self._presence_ts = code, now

        def _send():
            try:
                vps_request("POST", f"{VPS_URL}/presence", json={"gep": GEP_NEV, "A": code}, timeout=3)
            except Exception:
                pass
        threading.Thread(target=_send, daemon=True).start()

    def _remote_jump(self, code, print_it=False):
        row = self._excel_code_index.get(code.upper(), -1) if hasattr(self, "_excel_code_index") else -1
        if row < 0:
            self.trigger_toast(f"⚠️ {code} még nincs az Excel Live-on – pár másodperc és megjelenik")
            return
        item = self.excel_table.item(row, 2) or self.excel_table.item(row, 0)
        if item:
            self.excel_table.setCurrentItem(item)
            self.excel_table.scrollToItem(item, self.excel_table.ScrollHint.PositionAtCenter)
            self.excel_table.horizontalScrollBar().setValue(0)
        if print_it:
            self._remote_printed.add(code.upper())
            self.print_row_label(row)
            self.render_remote_marks()

    def render_remote_marks(self):
        """Excel Live: a sor számánál 🆕 GÉP (másik gépen felvitt) / ✍️ GÉP (most tölti), fölötte kattintható lista."""
        if not hasattr(self, "remote_bar"):
            return
        now = time.time()
        for k in [k for k, v in self._remote_edits.items() if now - v["t"] > REMOTE_MARK_SEC]:
            self._remote_edits.pop(k, None)
        idx = getattr(self, "_excel_code_index", {})
        marks = {}
        for code, v in self._remote_edits.items():
            r = idx.get(code, -1)
            if r >= 0:
                marks[r] = (f"🆕 {r + 1} · {v['gep']} {v['ido']}", "#00e676")
        working = {}
        for p in []:            # a „most tölti” jelzés kikapcsolva
            working.setdefault(p["A"], []).append(p["gep"])
            r = idx.get(p["A"], -1)
            if r >= 0:
                marks[r] = (f"✍️ {r + 1} · {', '.join(working[p['A']])} most tölti", "#ffb74d")
        t = self.excel_table
        for r in self._remote_marked_rows - set(marks):
            if r < t.rowCount():
                t.takeVerticalHeaderItem(r)
        for r, (txt, col) in marks.items():
            h = QTableWidgetItem(txt)
            h.setForeground(QColor(col))
            h.setToolTip("Másik gépen most szerkesztik" if txt.startswith("✍️") else "Másik gépen felvitt / módosított sor")
            t.setVerticalHeaderItem(r, h)
        self._remote_marked_rows = set(marks)

        # kattintható lista
        while self.remote_chips.count():
            w = self.remote_chips.takeAt(0).widget()
            if w:
                w.setParent(None)
                w.deleteLater()
        n = 0
        for code, geps in working.items():
            b = QPushButton(f"✍️ {', '.join(geps)} most tölti: {code}")
            b.setStyleSheet("background-color: #ffb74d; color: black; font-weight: bold; padding: 4px 10px; border-radius: 4px;")
            b.setToolTip("Kattints: odaugrik a sorra")
            b.clicked.connect(lambda _=False, c=code: self._remote_jump(c))
            self.remote_chips.addWidget(b)
            n += 1
        for code, v in sorted(self._remote_edits.items(), key=lambda kv: -kv[1]["t"])[:12]:
            done = code in self._remote_printed
            b = QPushButton(f"{'✔' if done else '🖨️'} {code} {v['nev'][:28]} · {v['gep']} {v['ido']}")
            b.setStyleSheet("background-color: %s; color: black; font-weight: bold; padding: 4px 10px; border-radius: 4px;"
                            % ("#9e9e9e" if done else "#00e676"))
            b.setToolTip("Kattints: odaugrik a sorra és nyomtatja a címkét (előtte megkérdezi a darabszámot)")
            b.clicked.connect(lambda _=False, c=code: self._remote_jump(c, print_it=True))
            self.remote_chips.addWidget(b)
            n += 1
        if not n:
            nincs = QLabel("nincs új felvitel")
            nincs.setStyleSheet("color: #777;")
            self.remote_chips.addWidget(nincs)
        self.remote_chips.addStretch()

    def print_current_excel_row(self):
        """Ctrl+P: a kijelölt (éppen szerkesztett) sor címkéjét nyomtatja."""
        # Ha épp gépel egy cellában, a fókusz elvétele elmenti a beírt értéket a nyomtatás előtt
        self.excel_table.setFocus()
        row = self.excel_table.currentRow()
        if row < 0:
            self.trigger_toast("⚠️ Nincs kijelölt sor a nyomtatáshoz!")
            return
        self.print_row_label(row)

    def load_data_from_db(self):
        self.load_latest_200_items()

    def track_worker_finished(self, worker):
        if worker in self.active_workers:
            self.active_workers.remove(worker)

    def check_vps_connection(self):
        self.vps_worker = VPSCheckWorker()
        self.vps_worker.status_signal.connect(self.update_server_status_ui)
        self.vps_worker.finished.connect(lambda: self.track_worker_finished(self.vps_worker))
        self.active_workers.append(self.vps_worker)
        self.vps_worker.start()

    def zoom_in(self):
        if self.zoom_factor < 1.5:
            self.zoom_factor += 0.1
            self.update_zoom()

    def zoom_out(self):
        if self.zoom_factor > 0.5:
            self.zoom_factor -= 0.1
            self.update_zoom()

    def update_zoom(self):
        percent = int(round(self.zoom_factor * 100))
        self.lbl_zoom_val.setText(f"{percent}%")
        save_zoom_factor(round(self.zoom_factor, 2))
        self.update_app_stylesheet()

    def update_app_stylesheet(self):
        base_font_size = int(14 * self.zoom_factor)
        input_font_size = int(15 * self.zoom_factor)
        btn_padding = int(8 * self.zoom_factor)

        style = f"""
            QMainWindow {{ background-color: #121212; color: #FFFFFF; }}
            QWidget {{ background-color: #1E1E1E; color: #FFFFFF; font-size: {base_font_size}px; }}
            QLineEdit {{ background-color: #2D2D2D; border: 1px solid #444; border-radius: 6px; padding: 6px; font-size: {input_font_size}px; color: #FFF; }}
            
            QPushButton {{ 
                background-color: #29b6f6; 
                color: #000; 
                font-weight: bold; 
                border-radius: 6px; 
                padding: {btn_padding}px; 
                border: 1px solid #1a83b8;
            }}
            QPushButton:hover {{ 
                background-color: #4fc3f7; 
                border: 1px solid #ffffff;
            }}
            QPushButton:pressed {{ 
                background-color: #0288d1; 
                color: #ffffff; 
            }}

            QTabWidget::pane {{ border: 1px solid #333; }}
            QTabBar::tab {{ background: #2D2D2D; color: #AAA; padding: 8px 14px; font-weight: bold; }}
            QTabBar::tab:selected {{ background: #29b6f6; color: #000; }}
            QGroupBox {{ border: 1px solid #444; border-radius: 8px; margin-top: 12px; padding-top: 12px; font-weight: bold; color: #29b6f6; }}
            
            QRadioButton {{ color: #A0A0A0; font-weight: bold; padding: 3px; }}
            QRadioButton::indicator {{ width: 16px; height: 16px; border-radius: 9px; border: 2px solid #555555; background-color: #2A2A2A; }}
            QRadioButton::indicator:checked {{ border: 2px solid #00e676; background-color: #00e676; }}
            QRadioButton:checked {{ color: #00e676; font-weight: bold; }}
        """
        self.setStyleSheet(style)

    def undo_last_added_item(self):
        last_code = None
        try:
            conn = get_db_connection()
            cursor = conn.cursor()
            cursor.execute("""
                SELECT r18_kod FROM naplo 
                WHERE esemeny = 'TERMÉK FELVÉTELE' 
                ORDER BY id DESC LIMIT 1
            """)
            row = cursor.fetchone()
            if row:
                last_code = row[0]
            conn.close()
        except Exception:
            pass

        if not last_code:
            last_code, _, _ = get_last_processed_info()

        if not last_code or last_code == "R18-0":
            self.trigger_toast("⚠️ Nincs visszavonható felvétel!")
            return

        dlg = SilentConfirmDialog(
            "Visszavonás Megerősítése",
            f"Biztosan VISSZAVONOD az utoljára felvitt terméket?\n\nKód: {last_code}\n\nEz a művelet kiüríti a termék adatait, de az R18 kód megmarad a táblázatban!",
            self
        )
        if dlg.exec() == QDialog.DialogCode.Accepted:
            elotte = get_row_snapshot(last_code)
            utana = dict(elotte, **{c: "" for c in "CDEFGHIJKLMNOP"}) if elotte else None
            log_event(last_code, "TÖRLÉS (felvétel visszavonva)", f"A(z) {last_code} adatai kiürítve.", elotte, utana,
                      "Új termék fül - ↩️ Utolsó felvétel visszavonása gomb")
            if getattr(self, 'last_added_code', None) == last_code:
                self.last_added_code = None
            
            self.clear_worker = ClearWorker(last_code)
            self.clear_worker.process_log_signal.connect(self.handle_log_from_thread)
            self.clear_worker.finished_signal.connect(self.on_delete_finished)
            self.clear_worker.finished.connect(lambda: self.track_worker_finished(self.clear_worker))
            self.active_workers.append(self.clear_worker)
            self.clear_worker.start()

    def update_server_status_ui(self, connected, message):
        self.lbl_vps_status.setText(message)
        if connected:
            self.lbl_vps_status.setStyleSheet("font-weight: bold; color: #ffffff; padding: 4px 8px; border-radius: 4px; background: #2e7d32;")
        else:
            self.lbl_vps_status.setStyleSheet("font-weight: bold; color: #ffffff; padding: 4px 8px; border-radius: 4px; background: #c62828;")

    def refresh_counters(self):
        def async_calc():
            code, count, gaps = get_last_processed_info()
            self.counter_update_signal.emit(code, count, gaps)
        threading.Thread(target=async_calc, daemon=True).start()

    def update_counter_ui(self, code, count, gaps):
        self.last_code_lbl.setText(f"📌 Utolsó feldolgozott R18 kód: {code}")
        self.total_count_lbl.setText(f"📊 Kitöltött alkatrészek száma: {count} db")
        
        next_code = get_next_suggested_r18_code(code)
        self.input_a.blockSignals(True)
        self.input_a.setText(next_code)
        self.input_a.blockSignals(False)
        if not self.input_b.text().strip() or self.input_b.text().startswith("R18-"):
            self.input_b.setText(next_code)

        if gaps:
            gap_txts = [f"{g[0]}" if g[0] == g[1] else f"{g[0]} és {g[1]} között" for g in gaps[:3]]
            self.gap_warning_lbl.setText(f"⚠️ FIGYELEM! Üres sorok találhatók: {', '.join(gap_txts)}!")
            self.gap_warning_box.show()
        else:
            self.gap_warning_box.hide()

    def handle_log_from_thread(self, msg, log_type="INFO"):
        color = "#00e676"
        if log_type == "ERROR": color = "#ff5252"
        elif log_type == "SUCCESS": color = "#29b6f6"
        elif log_type == "INFO": color = "#aaaaaa"
        self.append_process_log(msg, color)

    def trigger_toast(self, message, duration=2500):
        toast = ToastNotification(message, self)
        toast.show_toast(duration=duration)

    def append_process_log(self, text, color="#00e676"):
        """⚙️ Folyamat napló: CSAK technikai üzenetek (szinkron, VPS, hibák). Naponta külön fájlba
        is menti (naplok/folyamat_ÉÉÉÉ-HH-NN.log), így a régebbi napok is visszanézhetők."""
        now = datetime.datetime.now()
        try:
            os.makedirs(PROCESS_LOG_DIR, exist_ok=True)
            with open(os.path.join(PROCESS_LOG_DIR, f"folyamat_{now:%Y-%m-%d}.log"), "a", encoding="utf-8") as f:
                f.write(f"{now:%H:%M:%S}\t{color}\t{str(text).replace(chr(10), ' ')}\n")
        except Exception:
            pass
        # A nézetbe csak akkor, ha éppen a mai (élő) napot nézzük
        if hasattr(self, 'process_log_text') and getattr(self, "_process_log_live", True):
            self.process_log_text.append(self._process_log_line_html(f"{now:%H:%M:%S}", color, text))
        # Hiba / figyelmeztetés: érthető magyarázat a sor alá (fájlba is)
        is_problem = color.lower() in ("#ff5252", "#ff1744", "#ffb74d", "#ff9800") or any(m in str(text) for m in ("❌", "⚠️"))
        hint = technical_hint(text) if is_problem else ""
        if hint and not str(text).startswith("💡"):
            self.append_process_log(f"💡 {hint}", "#ffe082")

    @staticmethod
    def _process_log_line_html(time_str, color, text):
        import html as _html
        return f"<span style='color: #888;'>[{time_str}]</span> <span style='color: {color};'>{_html.escape(str(text))}</span>"

    # -------------------------------------------------------------------
    # TAB 1: GYORSKERESŐ TAB
    # -------------------------------------------------------------------
    def init_search_tab(self):
        layout = QVBoxLayout(self.tab_search)
        search_layout = QHBoxLayout()
        self.search_input = QLineEdit("R18-")
        self.search_input.setPlaceholderText("Pittyents ide a pisztollyal...")
        self.search_input.returnPressed.connect(self.execute_search)
        search_layout.addWidget(self.search_input)

        self.search_btn = QPushButton("🔍 Keresés")
        self.search_btn.clicked.connect(
            lambda: self.run_button_action(
                self.search_btn, 
                "🔍 Keresés", 
                self.execute_search
            )
        )
        
        search_layout.addWidget(self.search_btn)
        layout.addLayout(search_layout)

        self.results_area = QWidget()
        self.results_layout = QVBoxLayout(self.results_area)
        self.results_layout.setAlignment(Qt.AlignmentFlag.AlignTop)

        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setWidget(self.results_area)
        layout.addWidget(scroll)

    def create_item_card(self, row, sale_mode=False):
        card = QFrame()
        card.setStyleSheet("QFrame#MainCard { background-color: #1e1e1e; border: 1px solid #333; border-radius: 8px; margin-bottom: 10px; } QLineEdit { background-color: #2b2b2b; border: 1px solid #444; color: #fff; padding: 4px; border-radius: 4px; } QLabel { color: #fff; }")
        card.setObjectName("MainCard")
        main_card_layout = QHBoxLayout(card)
        main_card_layout.setContentsMargins(10, 10, 10, 10)
        main_card_layout.setSpacing(12)

        row_id = row.get("rowid")
        raw_code = get_best_code(row)
        code_match = re.search(r'R18-\d+', raw_code, re.IGNORECASE)
        code = code_match.group(0).upper() if code_match else raw_code.split()[0].upper()

        marka = safe_str(row.get('C'))
        tipus = safe_str(row.get('D'))
        cikkszam = safe_str(row.get('E'))
        
        cikk_smart = get_smart_part_number(tipus, cikkszam)
        
        raw_f = safe_str(row.get('F')).strip()
        f_match = re.search(r'\d+', raw_f)
        db_val = int(f_match.group(0)) if f_match else 1
        allapot = safe_str(row.get('G')) or "HASZNÁLT"
        suly = safe_str(row.get('M'))
        
        raw_ar = safe_str(row.get('H'))
        ar = "" if raw_ar.lower() in ("nincs", "none", "nan") else raw_ar
        
        pater = safe_str(row.get('I'))
        polc = safe_str(row.get('K'))
        hely = safe_str(row.get('L'))       # a Páter (I) NEM hely - azt külön mutatjuk (a térképen nincs rajta)
        ebay_nr = safe_str(row.get('N')).strip()
        
        left_box = QWidget()
        left_layout = QVBoxLayout(left_box)
        left_layout.setContentsMargins(0, 0, 0, 0)
        left_layout.setSpacing(8)

        row1 = QHBoxLayout()
        lbl_code = QLabel(f"🏷️ <b>KÓD:</b> {code}")
        lbl_code.setStyleSheet("font-size: 15px; font-weight: bold; color: #fff;")
        row1.addWidget(lbl_code)

        # 🛒 eBay item number (N oszlop): jól látható, egy kattintással másolható
        ebay_box = QWidget()
        ebay_box.setStyleSheet("background-color: #3a2f00; border: 1px solid #ffd600; border-radius: 4px;")
        ebay_lay = QHBoxLayout(ebay_box)
        ebay_lay.setContentsMargins(6, 2, 4, 2)
        ebay_lay.setSpacing(4)
        lbl_ebay = QLabel(f"🛒 eBay: <b>{ebay_nr}</b>" if ebay_nr else "🛒 eBay: <i>nincs</i>")
        lbl_ebay.setStyleSheet("color: #ffd600; border: none; font-size: 14px;")
        lbl_ebay.setTextInteractionFlags(Qt.TextInteractionFlag.TextSelectableByMouse)
        ebay_lay.addWidget(lbl_ebay)
        if ebay_nr:
            btn_copy_ebay = QPushButton("📋 Másolás")
            btn_copy_ebay.setToolTip("Az eBay item number vágólapra másolása (Ctrl+V-vel beilleszthető)")
            btn_copy_ebay.setCursor(Qt.CursorShape.PointingHandCursor)
            btn_copy_ebay.setStyleSheet("background-color: #ffd600; color: black; font-weight: bold; padding: 2px 8px; border-radius: 3px;")
            def copy_ebay(checked=False, nr=ebay_nr, c=code):
                QApplication.clipboard().setText(nr)
                self.trigger_toast(f"📋 eBay szám másolva: {nr} [{c}]")
            btn_copy_ebay.clicked.connect(copy_ebay)
            ebay_lay.addWidget(btn_copy_ebay)
            if re.fullmatch(r"\d{9,15}", ebay_nr):
                btn_open_ebay = QPushButton("🌐")
                btn_open_ebay.setToolTip("Hirdetés megnyitása az eBay-en")
                btn_open_ebay.setFixedWidth(30)
                btn_open_ebay.setStyleSheet("background-color: #444; color: white; padding: 2px; border-radius: 3px;")
                btn_open_ebay.clicked.connect(lambda checked=False, nr=ebay_nr: QDesktopServices.openUrl(QUrl(f"https://www.ebay.com/itm/{nr}")))
                ebay_lay.addWidget(btn_open_ebay)
        row1.addWidget(ebay_box)

        db_box = QWidget()
        db_box.setStyleSheet("background-color: #29b6f6; border-radius: 4px; padding: 2px;")
        db_lay = QHBoxLayout(db_box)
        db_lay.setContentsMargins(4, 2, 4, 2)
        lbl_db_title = QLabel("📦 DB:")
        lbl_db_title.setStyleSheet("color: #000; font-weight: bold;")
        db_lay.addWidget(lbl_db_title)

        btn_minus = QPushButton("-")
        btn_minus.setFixedSize(22, 22)
        btn_minus.setStyleSheet("background-color: #ff5252; color: #fff; font-weight: bold; border-radius: 3px;")
        db_lay.addWidget(btn_minus)

        inp_db = QLineEdit(str(db_val))
        inp_db.setFixedWidth(35)
        inp_db.setAlignment(Qt.AlignmentFlag.AlignCenter)
        inp_db.setStyleSheet("background: #000; color: #fff; font-weight: bold; border: none;")
        db_lay.addWidget(inp_db)

        btn_plus = QPushButton("+")
        btn_plus.setFixedSize(22, 22)
        btn_plus.setStyleSheet("background-color: #00e676; color: #000; font-weight: bold; border-radius: 3px;")
        db_lay.addWidget(btn_plus)

        btn_minus.clicked.connect(lambda: inp_db.setText(str(max(0, int(inp_db.text() or "0") - 1))))
        btn_plus.clicked.connect(lambda: inp_db.setText(str(int(inp_db.text() or "0") + 1)))

        row1.addWidget(db_box)

        search_target = get_smart_part_number(tipus, cikkszam) or code
        clean_target = clean_part_number(search_target)
        
        try:
            total_db_sum = get_stock_sum(search_target)
        except Exception:
            total_db_sum = db_val

        orig_total_txt = f"📊 Összesen raktáron: {total_db_sum} db"
        btn_total = QPushButton(orig_total_txt)
        btn_total.setCursor(Qt.CursorShape.PointingHandCursor)
        btn_total.setStyleSheet("background-color: #00e676; color: #000; font-weight: bold; padding: 4px 8px; border-radius: 4px; border: none;")
        
        btn_total.clicked.connect(
            lambda checked, b=btn_total, txt=orig_total_txt, t=tipus, c=cikkszam, cd=code, s=total_db_sum: 
            self.run_button_action(
                b, 
                txt, 
                lambda: SummaryDialog(t, c, cd, s, self).exec()
            )
        )
        row1.addWidget(btn_total)

        row1.addStretch()
        left_layout.addLayout(row1)

        row2 = QHBoxLayout()
        row2.addWidget(QLabel("Márka (C):"))
        inp_marka = QLineEdit(marka)
        inp_marka.setFixedWidth(110)
        row2.addWidget(inp_marka)

        row2.addWidget(QLabel("Típus (D):"))
        inp_tipus = QLineEdit(tipus)
        row2.addWidget(inp_tipus, stretch=1)

        row2.addWidget(QLabel("Cikkszám (E):"))
        inp_cikk = QLineEdit(cikkszam)
        row2.addWidget(inp_cikk, stretch=1)
        left_layout.addLayout(row2)

        row3 = QHBoxLayout()
        row3.setSpacing(10)
        row3.addWidget(QLabel(f"Állapot: <b>{allapot}</b> |"))
        
        row3.addWidget(QLabel("Súly (kg):"))
        # M (Súly) - Mindig jobbra igazítva
        inp_suly = QLineEdit(suly)
        inp_suly.setAlignment(Qt.AlignmentFlag.AlignRight)
        inp_suly.setFixedWidth(60)
        row3.addWidget(inp_suly)

        row3.addWidget(QLabel("Ár (EUR):"))
        inp_ar = QLineEdit(ar)
        inp_ar.setPlaceholderText("Üres")
        inp_ar.setFixedWidth(70)
        row3.addWidget(inp_ar)

        row3.addStretch(1)

        # 4. SOR: Páter (I), Polc (K) és Raklap/Hely (L) mezők egymás mellett
        row4 = QHBoxLayout()

        row4.addWidget(QLabel("Páter (I):"))
        inp_pater = QLineEdit(str(pater or ""))
        inp_pater.setAlignment(Qt.AlignmentFlag.AlignRight)
        inp_pater.setPlaceholderText("Páter száma")
        inp_pater.setFixedWidth(90)
        row4.addWidget(inp_pater)

        row4.addWidget(QLabel("Oszlop (K):"))
        inp_polc = QLineEdit(str(polc or ""))
        inp_polc.setAlignment(Qt.AlignmentFlag.AlignRight)
        inp_polc.setPlaceholderText("Oszlop száma")
        inp_polc.setFixedWidth(90)
        row4.addWidget(inp_polc)

        # L (Raklap/Hely) - Csak akkor jobbra, ha kizárólag szám van benne
        row4.addWidget(QLabel("Hely (L):"))
        inp_hely = QLineEdit(str(hely or ""))
        if hely and re.match(r'^\d+([.,]\d+)?$', str(hely).strip()):
            inp_hely.setAlignment(Qt.AlignmentFlag.AlignRight)
        else:
            inp_hely.setAlignment(Qt.AlignmentFlag.AlignLeft)
        inp_hely.setPlaceholderText("pl. KISRAKTÁR 80")
        row4.addWidget(inp_hely, stretch=1)

        has_letters = bool(re.search(r'[A-Za-zÁÉÍÓÖŐÚÜŰáéíóöőúüű]', hely))
        if not has_letters and hely != "":
            map_txt = f"🗺️ Térkép ({hely})"
            btn_map = QPushButton(map_txt)
            btn_map.setStyleSheet("background-color: #00bcd4; color: black; font-weight: bold; padding: 4px 8px;")
            btn_map.clicked.connect(
                lambda checked, b=btn_map, txt=map_txt: self.run_button_action(
                    b, 
                    txt, 
                    lambda: MapDialog(inp_hely.text(), self).exec()
                )
            )
            row4.addWidget(btn_map)

        left_layout.addLayout(row4)

        # EGYETLEN MENTÉS LOGIKA ÉS GOMB
        def confirm_and_save_data():
            new_m = inp_marka.text().strip()
            new_d = inp_tipus.text().strip()
            new_e = inp_cikk.text().strip()
            new_s = inp_suly.text().strip() or suly # Ha a kártyán üres, megtartja az meglévőt!
            new_a = inp_ar.text().strip()
            clean_a = "" if new_a.lower() in ("nincs", "") else new_a
            new_db = inp_db.text().strip()
            new_i = inp_pater.text().strip() or pater
            new_k = inp_polc.text().strip() or polc
            new_l = inp_hely.text().strip() or hely

            diffs = []
            if new_m != marka: diffs.append(f"• C (Márka): '{marka}' ➔ '{new_m}'")
            if new_d != tipus: diffs.append(f"• D (Típus): '{tipus}' ➔ '{new_d}'")
            if new_e != cikkszam: diffs.append(f"• E (Cikkszám): '{cikkszam}' ➔ '{new_e}'")
            if new_s != suly: diffs.append(f"• M (Súly): '{suly}' ➔ '{new_s}' kg")
            if clean_a != ar: diffs.append(f"• H (Ár): '{ar}' ➔ '{clean_a}' EUR")
            if new_db != str(db_val): diffs.append(f"• F (DB): '{db_val}' ➔ '{new_db}' db")
            if new_i != pater: diffs.append(f"• I (Páter): '{pater}' ➔ '{new_i}'")
            if new_k != polc: diffs.append(f"• K (Polc): '{polc}' ➔ '{new_k}'")
            if new_l != hely: diffs.append(f"• L (Hely): '{hely}' ➔ '{new_l}'")

            msg = f"Biztosan mented a változtatásokat a(z) {code} terméknél?\n\n"
            msg += "\n".join(diffs) if diffs else "Nincs változás az adatokban."

            dlg = SilentConfirmDialog("Adatok Módosítása Megerősítés", msg, self)
            if dlg.exec() == QDialog.DialogCode.Accepted:
                self.save_card_data_action(row_id, code, new_m, new_d, new_e, new_s, clean_a, new_db, new_i, new_k, new_l)

        btn_save_data = QPushButton("💾 Adatok Mentése")
        btn_save_data.setStyleSheet("background-color: #29b6f6; color: #000; font-weight: bold; padding: 5px 12px;")
        
        btn_save_data.clicked.connect(
            lambda: self.run_button_action(
                btn_save_data, 
                "💾 Adatok Mentése", 
                confirm_and_save_data
            )
        )
        row3.addWidget(btn_save_data)
        left_layout.addLayout(row3)

        row5 = QHBoxLayout()
        row5.addWidget(QLabel("🎨 A-B Szín:"))
        btn_orange = QPushButton("Narancs")
        btn_orange.setStyleSheet("background-color: #ff9800; color: #000; font-weight: bold; border-radius: 4px;")
        btn_red = QPushButton("Piros")
        btn_red.setStyleSheet("background-color: #ff2a2a; color: #fff; font-weight: bold; border-radius: 4px;")
        btn_white = QPushButton("Fehér")
        btn_white.setStyleSheet("background-color: #ffffff; color: #000; font-weight: bold; border-radius: 4px;")

        btn_orange.clicked.connect(
            lambda checked, b=btn_orange, c=code: self.run_button_action(
                b, "Narancs", self.confirm_and_set_color, c, COLOR_ORANGE, "Narancs"
            )
        )
        btn_red.clicked.connect(
            lambda checked, b=btn_red, c=code: self.run_button_action(
                b, "Piros", self.confirm_and_set_color, c, COLOR_RED, "Piros"
            )
        )
        btn_white.clicked.connect(
            lambda checked, b=btn_white, c=code: self.run_button_action(
                b, "Fehér", self.confirm_and_set_color, c, COLOR_WHITE, "Fehér"
            )
        )

        row5.addWidget(btn_orange)
        row5.addWidget(btn_red)
        row5.addWidget(btn_white)
        left_layout.addLayout(row5)

        btn_layout = QHBoxLayout()
        btn_layout.setSpacing(10)

        btn_reset = QPushButton("↺ Adatok Kiürítése (R18 megmarad)")
        btn_reset.setCursor(Qt.CursorShape.PointingHandCursor)
        btn_reset.setStyleSheet("""
            QPushButton { background-color: #ff9800; color: black; font-weight: bold; padding: 6px; border-radius: 4px; font-size: 13px; }
            QPushButton:hover { background-color: #ffb74d; }
        """)
        code_val = get_best_code(row)
        
        btn_reset.clicked.connect(
            lambda checked, b=btn_reset, c=code_val: self.run_button_action(
                b, "↺ Adatok Kiürítése (R18 megmarad)", self.reset_item_keep_code, c
            )
        )
        btn_layout.addWidget(btn_reset, stretch=1)

        btn_delete = QPushButton("🗑️ Teljes Törlés / Eladás")
        btn_delete.setCursor(Qt.CursorShape.PointingHandCursor)
        btn_delete.setStyleSheet("""
            QPushButton { background-color: #ff2a2a; color: white; font-weight: bold; padding: 6px; border-radius: 4px; font-size: 13px; }
            QPushButton:hover { background-color: #ff5252; }
        """)
        
        btn_delete.clicked.connect(
            lambda checked, b=btn_delete, c=code_val: self.run_button_action(
                b, "🗑️ Teljes Törlés / Eladás", self.delete_item_action, c
            )
        )
        btn_layout.addWidget(btn_delete, stretch=1)

        left_layout.addLayout(btn_layout)

        # PDA / Eladás fül: 3 eladás gomb - mindegyik a saját eladási munkalapjára ír
        if sale_mode:
            sale_layout = QHBoxLayout()
            sale_layout.setSpacing(10)
            sale_styles = {
                "EBAY_KIVULI": "background-color: #7e57c2; color: white;",
                "EBAY": "background-color: #ffd600; color: black;",
                "WEBSHOP": "background-color: #00c853; color: black;",
            }
            for pf, cfg in SALES_SHEETS.items():
                b = QPushButton(f"💰 {cfg['cim']}")
                b.setCursor(Qt.CursorShape.PointingHandCursor)
                b.setStyleSheet(f"QPushButton {{ {sale_styles.get(pf, '')} font-weight: bold; padding: 10px; border-radius: 4px; font-size: 14px; }}")
                b.clicked.connect(lambda checked, p=pf, c=code_val: self.open_sale_dialog(p, c))
                sale_layout.addWidget(b, stretch=1)
            left_layout.addLayout(sale_layout)

        main_card_layout.addWidget(left_box, stretch=3)

        if PHOTOS_ENABLED:
            main_card_layout.addWidget(self.create_photo_widget(code))

        qr_bytes, bar_bytes = generate_codes_in_memory(code, cikkszam_text=cikkszam, tipus_text=tipus)
        if qr_bytes and bar_bytes:
            codes_box = QWidget()
            codes_layout = QHBoxLayout(codes_box)
            codes_layout.setContentsMargins(0, 0, 0, 0)

            lbl_qr = QLabel()
            pix_qr = QPixmap()
            pix_qr.loadFromData(qr_bytes)
            lbl_qr.setPixmap(pix_qr.scaled(85, 85, Qt.AspectRatioMode.KeepAspectRatio))

            lbl_bar = QLabel()
            pix_bar = QPixmap()
            pix_bar.loadFromData(bar_bytes)
            lbl_bar.setPixmap(pix_bar.scaled(150, 85, Qt.AspectRatioMode.KeepAspectRatio))

            codes_layout.addWidget(lbl_qr)
            codes_layout.addWidget(lbl_bar)
            main_card_layout.addWidget(codes_box)

        print_box = QGroupBox("🖨️ Nyomtatás")
        print_box.setFixedWidth(180)
        print_lay = QVBoxLayout(print_box)

        r_zpl = QRadioButton("ZPL")
        r_png = QRadioButton("PNG")
        r_zpl.setChecked(True)
        print_lay.addWidget(r_zpl)
        print_lay.addWidget(r_png)

        print_lay.addWidget(QLabel("Példány:"))
        spin_copies = QSpinBox()
        spin_copies.setMinimum(1)
        spin_copies.setValue(1)
        print_lay.addWidget(spin_copies)

        btn_card_print = QPushButton("🖨️ Nyomtatás")
        btn_card_print.setStyleSheet("background-color: #29b6f6; color: #000; font-weight: bold; padding: 6px;")
        btn_card_print.clicked.connect(
            lambda checked, b=btn_card_print, c=code, cs=cikkszam, t=tipus, z=r_zpl, sp=spin_copies: 
            self.run_button_action(
                b, 
                "🖨️ Nyomtatás", 
                self.print_card_label, 
                c, cs, t, z.isChecked(), sp.value()
            )
        )
        print_lay.addWidget(btn_card_print)

        main_card_layout.addWidget(print_box)

        return card

    def create_photo_widget(self, code):
        """📷 Termékfotó a kártyán: kicsinyített előnézet (háttérben töltve), kattintásra nagy nézet."""
        box = QWidget()
        box.setFixedWidth(PHOTO_THUMB_W + 6)
        lay = QVBoxLayout(box)
        lay.setContentsMargins(0, 0, 0, 0)
        lay.setSpacing(2)
        thumb = QPushButton()
        thumb.setFixedSize(PHOTO_THUMB_W, PHOTO_THUMB_H)
        thumb.setCursor(Qt.CursorShape.PointingHandCursor)
        thumb.setStyleSheet("QPushButton { background-color: #111; border: 1px solid #444; border-radius: 6px; color: #888; "
                            "font-size: 10px; padding: 0px; }"
                            "QPushButton:hover { border: 2px solid #29b6f6; }")
        lay.addWidget(thumb)
        cap = QLabel("")
        cap.setAlignment(Qt.AlignmentFlag.AlignCenter)
        cap.setStyleSheet("color: #aaa; font-size: 11px;")
        lay.addWidget(cap)

        if not photo_index.dirs():
            thumb.setText("📁 Fotó-\nmappa")
            thumb.clicked.connect(self.choose_photo_dir)
            return box
        paths = photo_index.photos(code)
        if not paths:
            thumb.setText("📷\nnincs" if photo_index.built_at else "📷\n⏳")
            thumb.setToolTip("Nincs fotó ehhez az R18-számhoz a fotómappában\n(a fájl- vagy mappanévben kell szerepelnie, pl. "
                             f"{code}.jpg)")
            thumb.clicked.connect(lambda: QDesktopServices.openUrl(QUrl.fromLocalFile(photo_index.dirs()[0])))
            return box
        thumb.setText("⏳")
        cap.setText(f"📷 {len(paths)} fotó")
        thumb.setToolTip(f"{len(paths)} fotó - kattints a nagy nézethez")
        thumb.clicked.connect(lambda: PhotoDialog(code, paths, 0, self).exec())
        holder = {}

        def on_ready(r, img, t=thumb):
            if r != holder.get("rid"):
                return
            try:
                photo_bus.thumb_ready.disconnect(on_ready)
            except Exception:
                pass
            try:
                if img == "CLOUD":
                    t.setText("☁️ NAS\nkattints")
                    t.setToolTip("A fotó csak a NAS-on van, ezen a gépen nincs letöltve.\nKattints, és letölti / megmutatja.\n"
                                 "Tipp: a Synology Drive-ban a fotómappára 'Mindig a gépen tartás' -> előnézet is lesz.")
                elif isinstance(img, QImage) and not img.isNull():
                    from PyQt6.QtGui import QIcon
                    t.setText("")
                    t.setIcon(QIcon(QPixmap.fromImage(img)))
                    t.setIconSize(img.size())
                else:
                    t.setText("📷 ❌")
            except RuntimeError:
                pass   # a kártya közben megszűnt
        # ELŐBB feliratkozunk, csak utána kérjük a képet (különben a gyors válasz elveszhet)
        photo_bus.thumb_ready.connect(on_ready)
        holder["rid"] = request_thumbnail(paths[0], PHOTO_THUMB_W, PHOTO_THUMB_H)
        return box

    def choose_photo_dir(self):
        start = (photo_index.dirs() or [os.path.expanduser("~")])[0]
        d = QFileDialog.getExistingDirectory(self, "Válaszd ki a fotók mappáját (Synology Drive)", start)
        if not d:
            return
        save_config_value("foto_mappak", [d])
        self.trigger_toast("📷 Fotómappa beállítva - fotók beolvasása...")
        photo_index.build_async(force=True)

    def _on_photo_index_ready(self, count):
        self.append_process_log(f"📷 Fotók beolvasva: {count} kép, {len(photo_index.by_code)} termékhez "
                                f"({'; '.join(photo_index.dirs())})", "#29b6f6")
        if self._tab_is_current("tab_search"):
            self.update_cards_only()

    def save_card_data_action(self, row_id, code, new_m, new_d, new_e, new_s, clean_a, new_db, new_i="", new_k="", new_l=""):
        self.card_worker = CardUpdateWorker(row_id, code, new_m, new_d, new_e, new_s, clean_a, new_db, new_i, new_k, new_l)
        self.card_worker.process_log_signal.connect(self.handle_log_from_thread)
        self.card_worker.finished_signal.connect(lambda msg: self.on_card_saved(msg))
        self.card_worker.finished.connect(lambda: self.track_worker_finished(self.card_worker))
        self.active_workers.append(self.card_worker)
        self.card_worker.start()

    def on_card_saved(self, msg):
        self.trigger_toast(msg)
        self.load_latest_200_items()

    def confirm_and_set_color(self, code, color_dict, color_name):
        dlg = SilentConfirmDialog("Szín Módosítása Megerősítés", f"Biztosan átszínezed a(z) {code} kód A és B oszlopát {color_name} színre?", self)
        if dlg.exec() == QDialog.DialogCode.Accepted:
            self.color_worker = ColorWorker(code, color_dict, color_name)
            self.color_worker.finished_signal.connect(lambda msg: self.trigger_toast(msg))
            self.color_worker.finished.connect(lambda: self.track_worker_finished(self.color_worker))
            self.active_workers.append(self.color_worker)
            self.color_worker.start()

    def delete_item_action(self, code, reason="", forras=""):
        self.del_worker = DeleteWorker(code, reason, forras)
        self.del_worker.process_log_signal.connect(self.handle_log_from_thread)
        self.del_worker.finished_signal.connect(self.on_delete_finished)
        self.del_worker.finished.connect(lambda: self.track_worker_finished(self.del_worker))
        self.active_workers.append(self.del_worker)
        self.del_worker.start()

    def on_delete_finished(self, msg):
        self.trigger_toast(msg)
        self.load_latest_200_items()
        self.refresh_counters()

    def print_card_label(self, code, cikkszam, tipus="", is_zpl=True, copies=1):
        zebra_printer = get_zebra_printer_name()
        
        if is_zpl:
            zpl = generate_zpl_code(code, cikkszam=cikkszam, tipus=tipus, copies=copies)
            ok, msg = send_zpl_to_printer(zebra_printer, zpl)
            if ok:
                self.trigger_toast(f"🖨️ ZPL Címke kiküldve a Zebra nyomtatóra [{code}]!")
            else:
                self.trigger_toast(f"❌ Nyomtatási hiba: {msg}")
        else:
            _, bar_bytes = generate_codes_in_memory(code, cikkszam_text=cikkszam, tipus_text=tipus)
            if bar_bytes:
                ok, msg = send_png_bytes_to_printer(zebra_printer, bar_bytes, copies)
                if ok:
                    self.trigger_toast(f"🖨️ PNG Címke kinyomtatva [{code}]!")
                else:
                    self.trigger_toast(f"❌ Nyomtatási hiba: {msg}")
                
    def load_latest_200_items(self):
        self.refresh_local_db_table_view()
        # A kártyákat csak akkor építjük újra, ha a Gyorskereső fül látszik; különben fülváltáskor
        if not self._tab_is_current("tab_search"):
            self._cards_dirty = True
            return
        self._cards_dirty = False
        self.results_area.setUpdatesEnabled(False)
        
        while self.results_layout.count():
            item = self.results_layout.takeAt(0)
            widget = item.widget()
            if widget:
                widget.deleteLater()

        info_lbl = QLabel("💡 Pittyents a pisztollyal a kereséshez, vagy nyomtass közvetlenül az utoljára felvitt alkatrészek közül:")
        info_lbl.setStyleSheet("background-color: #1565c0; color: white; padding: 10px; border-radius: 6px; font-size: 15px; font-weight: bold;")
        self.results_layout.addWidget(info_lbl)

        header_lbl = QLabel("🆕 Legutóbb felvitt alkatrészek (A legfrissebb felül):")
        header_lbl.setStyleSheet("font-size: 16px; font-weight: bold; margin-top: 10px; color: #29b6f6;")
        self.results_layout.addWidget(header_lbl)

        try:
            conn = get_db_connection()
            query = """
                SELECT * FROM termekek 
                WHERE C IS NOT NULL AND TRIM(C) != '' AND LOWER(TRIM(C)) != 'none' AND LOWER(TRIM(C)) != 'nan'
            """
            cursor = conn.cursor()
            cursor.execute(query)
            rows = cursor.fetchall()
            conn.close()

            filled_c_items = []
            for row in rows:
                row_dict = dict(row)
                code = get_best_code(row_dict)
                m = re.search(r'R18-(\d+)', str(code), re.IGNORECASE)
                if m:
                    num = int(m.group(1))
                    if 0 < num < 100000:
                        filled_c_items.append((num, row_dict))

            filled_c_items.sort(key=lambda x: x[0], reverse=True)

            seen_codes = set()
            valid_count = 0
            
            for num, row_dict in filled_c_items:
                code_raw = get_best_code(row_dict)
                
                if code_raw in seen_codes:
                    continue
                
                seen_codes.add(code_raw)
                
                self.results_layout.addWidget(self.create_item_card(row_dict))
                valid_count += 1
                if valid_count == 15:
                    break

            if valid_count == 0:
                self.results_layout.addWidget(QLabel("Nincs megjeleníthető friss alkatrész."))
        except Exception as e:
            self.results_layout.addWidget(QLabel(f"Hiba a legújabb termékek betöltésekor: {e}"))
        finally:
            self.results_area.setUpdatesEnabled(True)

    def execute_search(self):
        query_text = fix_scanner_text(self.search_input.text())
        if query_text != self.search_input.text().strip():
            self.search_input.setText(query_text)
        if not query_text or query_text.strip().upper() == "R18-": 
            self.load_latest_200_items()
            return

        clean_query = clean_part_number(query_text)

        self.results_area.setUpdatesEnabled(False)
        while self.results_layout.count():
            item = self.results_layout.takeAt(0)
            widget = item.widget()
            if widget:
                widget.deleteLater()

        conn = get_db_connection()
        df = pd.read_sql_query("SELECT * FROM termekek WHERE C IS NOT NULL AND C != ''", conn)
        conn.close()

        matched_rows = []
        for idx, row in df.iterrows():
            row_text = " ".join([safe_str(row.get(col)) for col in df.columns])
            clean_row = clean_part_number(row_text)
            
            if query_text.lower() in row_text.lower() or (clean_query and clean_query in clean_row):
                matched_rows.append(dict(row))
                if len(matched_rows) >= 30: break

        if not matched_rows:
            self.results_layout.addWidget(QLabel("❌ Nincs találat!"))
        else:
            for row_dict in matched_rows:
                self.results_layout.addWidget(self.create_item_card(row_dict))

        self.results_area.setUpdatesEnabled(True)

    # -------------------------------------------------------------------
    # TAB 2: ÚJ TERMÉK TAB
    # -------------------------------------------------------------------
    def init_add_tab(self):
        main_add_layout = QVBoxLayout(self.tab_add)
        main_add_layout.setSpacing(6)
        main_add_layout.setContentsMargins(8, 8, 8, 8)

        self.counter_box = QFrame()
        self.counter_box.setStyleSheet("background-color: rgba(255, 0, 0, 0.15); border: 2px solid #ff4b4b; border-radius: 6px; padding: 6px;")
        counter_layout = QVBoxLayout(self.counter_box)

        top_cnt_layout = QHBoxLayout()
        self.last_code_lbl = QLabel("📌 Utolsó feldolgozott R18 kód: Betöltés...")
        self.last_code_lbl.setStyleSheet("font-size: 16px; font-weight: bold; color: #ff4b4b;")
        self.total_count_lbl = QLabel("📊 Kitöltött alkatrészek száma: 0 db")
        self.total_count_lbl.setStyleSheet("font-size: 16px; font-weight: bold; color: #ffffff;")

        top_cnt_layout.addWidget(self.last_code_lbl)
        top_cnt_layout.addStretch()
        top_cnt_layout.addWidget(self.total_count_lbl)
        counter_layout.addLayout(top_cnt_layout)

        self.gap_warning_box = QFrame()
        self.gap_warning_box.setStyleSheet("background-color: rgba(255, 152, 0, 0.2); border: 1px solid #ff9800; border-radius: 6px; padding: 4px; margin-top: 4px;")
        gap_lay = QHBoxLayout(self.gap_warning_box)
        self.gap_warning_lbl = QLabel("")
        self.gap_warning_lbl.setStyleSheet("color: #ffb74d; font-weight: bold; font-size: 13px;")
        gap_lay.addWidget(self.gap_warning_lbl)
        self.gap_warning_box.hide()
        counter_layout.addWidget(self.gap_warning_box)

        main_add_layout.addWidget(self.counter_box)

        content_hbox = QHBoxLayout()

        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll_content = QWidget()
        scroll.setWidget(scroll_content)
        scroll_vlayout = QVBoxLayout(scroll_content)

        form_layout = QHBoxLayout()

        col1_box = QGroupBox("A - E Oszlopok")
        col1_layout = QFormLayout(col1_box)
        self.input_a = QLineEdit("R18-")
        self.input_b = QLineEdit()
        self.input_c = QLineEdit()
        self.input_d = QLineEdit()
        self.input_e = QLineEdit()
        col1_layout.addRow("A: F-Kód (R18-xxx):", self.input_a)
        col1_layout.addRow("B: Vonalkód:", self.input_b)
        col1_layout.addRow("C: Márka:", self.input_c)
        col1_layout.addRow("D: Típus / Megnevezés:", self.input_d)
        col1_layout.addRow("E: Cikkszám:", self.input_e)
        form_layout.addWidget(col1_box)

        col2_box = QGroupBox("Készlet & Pozíciók")
        col2_layout = QFormLayout(col2_box)
        self.input_f = QSpinBox()
        self.input_f.setMinimum(1)

        radio_box = QWidget()
        radio_layout = QVBoxLayout(radio_box)
        radio_layout.setContentsMargins(0, 0, 0, 0)
        self.radio_group = QButtonGroup(self)
        self.radio_used = QRadioButton("HASZNÁLT")
        self.radio_new = QRadioButton("ÚJ")
        self.radio_rep = QRadioButton("JAVÍTOTT")
        self.radio_def = QRadioButton("HIBÁS")
        self.radio_used.setChecked(True)

        for btn in [self.radio_used, self.radio_new, self.radio_rep, self.radio_def]:
            radio_layout.addWidget(btn)
            self.radio_group.addButton(btn)

        self.input_l = QLineEdit()
        self.input_i = QLineEdit()
        self.input_k = QLineEdit()

        col2_layout.addRow("F: Darabszám (DB):", self.input_f)
        col2_layout.addRow("G: Állapot:", radio_box)
        col2_layout.addRow("L: Raklap Száma:", self.input_l)
        col2_layout.addRow("I: Páter szám:", self.input_i)
        col2_layout.addRow("K: Páter Polc Szám:", self.input_k)
        form_layout.addWidget(col2_box)

        col3_box = QGroupBox("Egyéb Adatok (Súly & Ár)")
        col3_layout = QFormLayout(col3_box)
        self.input_m = QLineEdit()
        self.input_m.setAlignment(Qt.AlignmentFlag.AlignRight)
        self.input_n = QLineEdit()
        self.input_n.setPlaceholderText("eBay item number")
        self.input_o = QLineEdit()
        self.input_p = QLineEdit()
        self.input_h = QLineEdit()

        col3_layout.addRow("M: SÚLY (KG):", self.input_m)
        col3_layout.addRow("H: Ár (EUR):", self.input_h)
        col3_layout.addRow("N: Ebay Item Number:", self.input_n)
        col3_layout.addRow("O: Megjegyzés:", self.input_o)
        col3_layout.addRow("P: Állapot részletei:", self.input_p)
        form_layout.addWidget(col3_box)

        scroll_vlayout.addLayout(form_layout)

        btn_layout = QHBoxLayout()
        
        self.save_btn = QPushButton("⚡ Termék Felvétele (Azonnali)")
        self.save_btn.setStyleSheet("background-color: #00e676; color: #000; font-size: 16px; font-weight: bold; padding: 10px;")
        
        self.save_btn.clicked.connect(
            lambda: self.run_button_action(
                self.save_btn, 
                "⚡ Termék Felvétele (Azonnali)", 
                self.generate_and_save
            )
        )
        btn_layout.addWidget(self.save_btn, stretch=3)

        self.undo_btn = QPushButton("↩️ Utolsó Felvétel Visszavonása")
        self.undo_btn.setStyleSheet("background-color: #ff9800; color: #000; font-size: 15px; font-weight: bold; padding: 10px;")
        
        self.undo_btn.clicked.connect(
            lambda: self.run_button_action(
                self.undo_btn, 
                "↩️ Utolsó Felvétel Visszavonása", 
                self.undo_last_added_item
            )
        )
        btn_layout.addWidget(self.undo_btn, stretch=1)

        scroll_vlayout.addLayout(btn_layout)

        content_hbox.addWidget(scroll, stretch=3)

        self.preview_box = QGroupBox("🖨️ Címke Előnézet & Nyomtatás")
        self.preview_box.setFixedWidth(320)
        preview_vlayout = QVBoxLayout(self.preview_box)
        preview_vlayout.setSpacing(6)

        self.lbl_add_qr_img = QLabel()
        self.lbl_add_qr_img.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.lbl_add_bar_img = QLabel()
        self.lbl_add_bar_img.setAlignment(Qt.AlignmentFlag.AlignCenter)

        self.lbl_add_qr_img.hide()
        self.lbl_add_bar_img.hide()
        
        preview_vlayout.addWidget(self.lbl_add_qr_img)
        preview_vlayout.addWidget(self.lbl_add_bar_img)

        self.lbl_add_info_summary = QLabel("<b>Kód:</b> R18-<br><b>Márka:</b> -<br><b>Típus:</b> -")
        self.lbl_add_info_summary.setStyleSheet("background: #2b2b2b; padding: 6px; border-radius: 4px; font-size: 13px; color: #00e676;")
        preview_vlayout.addWidget(self.lbl_add_info_summary)

        mode_box = QWidget()
        mode_lay = QHBoxLayout(mode_box)
        mode_lay.setContentsMargins(0, 0, 0, 0)
        self.add_r_zpl = QRadioButton("ZPL Nyomtatás")
        self.add_r_png = QRadioButton("PNG Nyomtatás")
        self.add_r_zpl.setChecked(True)
        mode_lay.addWidget(self.add_r_zpl)
        mode_lay.addWidget(self.add_r_png)
        preview_vlayout.addWidget(mode_box)

        preview_vlayout.addWidget(QLabel("Példány:"))
        self.add_copies_spin = QSpinBox()
        self.add_copies_spin.setMinimum(1)
        self.add_copies_spin.setValue(1)
        preview_vlayout.addWidget(self.add_copies_spin)


        self.print_btn = QPushButton("🖨️ Nyomtatás")
        self.print_btn.setStyleSheet("background-color: #29b6f6; color: black; font-weight: bold; padding: 10px; font-size: 15px;")
        
        self.print_btn.clicked.connect(
            lambda: self.run_button_action(
                self.print_btn, 
                "🖨️ Nyomtatás", 
                self.print_current_label
            )
        )
        preview_vlayout.addWidget(self.print_btn)

        preview_vlayout.addStretch()
        content_hbox.addWidget(self.preview_box, stretch=1)

        main_add_layout.addLayout(content_hbox)

    def update_add_preview(self):
        code = self.input_a.text().strip().upper()
        marka = self.input_c.text().strip().upper() or "-"
        tipus = self.input_d.text().strip().upper() or "-"
        cikk = self.input_e.text().strip().upper() or "-"
        cikk_smart = get_label_subtext(tipus, cikk)

        self.lbl_add_info_summary.setText(f"<b>Kód:</b> {code}<br><b>Márka:</b> {marka}<br><b>Típus:</b> {tipus}<br><b>Cikkszám:</b> {cikk_smart}")

        qr_b, bar_b = generate_codes_in_memory(code, cikkszam_text=cikk, tipus_text=tipus)
        if qr_b:
            pix_qr = QPixmap()
            pix_qr.loadFromData(qr_b)
            self.lbl_add_qr_img.setPixmap(pix_qr.scaled(90, 90, Qt.AspectRatioMode.KeepAspectRatio))
        if bar_b:
            pix_bar = QPixmap()
            pix_bar.loadFromData(bar_b)
            self.lbl_add_bar_img.setPixmap(pix_bar.scaled(180, 75, Qt.AspectRatioMode.KeepAspectRatio))

    def generate_and_save(self):
        """Termék felvitele - villámgyors, fagyásmentes végrehajtás"""
        code_final = self.input_a.text().strip().upper() or "R18-TESZT"
        code_b = self.input_b.text().strip().upper() or code_final
        c_marka = self.input_c.text().strip().upper()
        d_tipus = self.input_d.text().strip().upper()
        e_cikk = self.input_e.text().strip().upper()
        
        self.last_added_code = code_final
        self.last_added_marka = c_marka
        self.last_added_tipus = d_tipus
        self.last_added_cikk = e_cikk

        allapot = "HASZNÁLT"
        if self.radio_new.isChecked(): allapot = "ÚJ"
        elif self.radio_rep.isChecked(): allapot = "JAVÍTOTT"
        elif self.radio_def.isChecked(): allapot = "HIBÁS"

        val_f = self.input_f.value()
        val_h = self.input_h.text().strip().upper()
        val_i = self.input_i.text().strip().upper()
        val_k = self.input_k.text().strip().upper()
        val_l = self.input_l.text().strip().upper()
        val_m = self.input_m.text().strip().upper()
        val_n = self.input_n.text().strip().upper()
        val_o = self.input_o.text().strip().upper()
        val_p = self.input_p.text().strip().upper()

        # 1. CÍMKE ELŐNÉZET MEGGENERÁLÁSA
        qr_bytes, bar_bytes = generate_codes_in_memory(code_final, cikkszam_text=e_cikk, tipus_text=d_tipus)
        if qr_bytes and bar_bytes:
            pix_qr = QPixmap()
            pix_qr.loadFromData(qr_bytes)
            self.lbl_add_qr_img.setPixmap(pix_qr.scaled(110, 110, Qt.AspectRatioMode.KeepAspectRatio))

            pix_bar = QPixmap()
            pix_bar.loadFromData(bar_bytes)
            self.lbl_add_bar_img.setPixmap(pix_bar.scaled(180, 110, Qt.AspectRatioMode.KeepAspectRatio))

            self.lbl_add_qr_img.show()
            self.lbl_add_bar_img.show()

        summary_txt = (f"🖨️ <b>Nyomtatásra kész (utoljára felvett):</b><br>"
                       f"<b>Kód:</b> {code_final}<br><b>Márka:</b> {c_marka or '-'}<br>"
                       f"<b>Típus:</b> {d_tipus or '-'}<br><b>Cikkszám:</b> {get_label_subtext(d_tipus, e_cikk) or '-'}")
        if hasattr(self, 'lbl_add_info_summary'):
            self.lbl_add_info_summary.setText(summary_txt)

        # 2. HELYI SQLITE MENTÉS (Azonnali, tranzakciós)
        try:
            elotte = get_row_snapshot(code_final)
            conn = get_db_connection()
            cursor = conn.cursor()
            cursor.execute("SELECT rowid FROM termekek WHERE UPPER(TRIM(A))=? OR UPPER(TRIM(B))=?", (code_final, code_final))
            existing_row = cursor.fetchone()

            if existing_row:
                cursor.execute("""
                    UPDATE termekek 
                    SET A=?, B=?, C=?, D=?, E=?, F=?, G=?, H=?, I=?, J='', K=?, L=?, M=?, N=?, O=?, P=?
                    WHERE rowid=?
                """, (code_final, code_b, c_marka, d_tipus, e_cikk, val_f, allapot, val_h, val_i, val_k, val_l, val_m, val_n, val_o, val_p, existing_row["rowid"]))
            else:
                cursor.execute("""
                    INSERT INTO termekek (A, B, C, D, E, F, G, H, I, J, K, L, M, N, O, P, szin)
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, '', ?, ?, ?, ?, ?, ?, 'FEHÉR')
                """, (code_final, code_b, c_marka, d_tipus, e_cikk, val_f, allapot, val_h, val_i, val_k, val_l, val_m, val_n, val_o, val_p))

            conn.commit()
            conn.close()

            # AUTOMATIKUS VPS FRISSÍTÉS (offline-biztos: sorba áll, majd azonnal próbál is küldeni;
            # ha nincs net, a periodikus outbox_flush_timer később automatikusan kiküldi)
            vps_payload = {
                "A": code_final, "B": code_b, "C": c_marka, "D": d_tipus,
                "E": e_cikk, "F": val_f, "G": allapot, "H": val_h,
                "I": val_i, "J": "", "K": val_k, "L": val_l,
                "M": val_m, "N": val_n, "O": val_o, "P": val_p
            }
            enqueue_outbox(vps_payload)
            threading.Thread(target=flush_sync_outbox, daemon=True).start()
            
            self.trigger_toast(f"⚡ {code_final} elmentve!")
            log_event(code_final, "FELVÉTEL", f"Új termék: {c_marka} {d_tipus}", elotte, get_row_snapshot(code_final),
                      "Új termék fül - ⚡ Termék felvétele gomb")
        except Exception as e:
            self.trigger_toast(f"❌ Mentési hiba: {e}")
            
        # 3. MEZŐK ÜRÍTÉSE ÉS KÓD LÉPTETÉSE
        self.clear_add_form()
        
        try:
            m = re.search(r'R18-(\d+)', code_final)
            if m:
                next_num = int(m.group(1)) + 1
                self.input_a.setText(f"R18-{next_num}")
        except Exception:
            pass

        # 4. FAGYÁSMENTES, CSENDESTETT NÉZETFRISSÍTÉS
        # A teljes load_latest_200_items() helyett csak a számlálót és az Excel táblát frissítjük csendben!
        if hasattr(self, 'silent_update_excel_table'):
            self.silent_update_excel_table()

        if hasattr(self, 'refresh_counters'):
            self.refresh_counters()

    def on_save_worker_finished(self, msg, elapsed):
        self.append_process_log(f"☁️ Felhő szinkron kész: {msg}", "#29b6f6")

    def clear_add_form(self):
        self.input_b.clear()
        self.input_c.clear()
        self.input_d.clear()
        self.input_e.clear()
        self.input_f.setValue(1)
        self.input_l.clear()
        self.input_i.clear()
        self.input_k.clear()
        self.input_m.clear()
        self.input_h.clear()
        self.input_n.clear()
        self.input_o.clear()
        self.input_p.clear()
        
    def print_current_label(self):
        """Címke nyomtatása az Új Termék fülön.
        A 'Termék Felvétele' után az A mező már a KÖVETKEZŐ (üres) R18-számot mutatja, ezért:
        - ha az űrlapba még nem írtak új adatot (Márka/Típus/Cikkszám üres), az UTOLJÁRA FELVETT
          termék címkéje jön ki a felvételkor megadott adatokkal;
        - ha már gépelnek új terméket, az űrlapon lévő adatok kerülnek a címkére."""
        form_code = self.input_a.text().strip().upper()
        form_marka = self.input_c.text().strip().upper()
        form_tipus = self.input_d.text().strip().upper()
        form_cikk = self.input_e.text().strip().upper()
        copies = self.add_copies_spin.value()

        last_code = getattr(self, 'last_added_code', None)
        if (form_marka or form_tipus or form_cikk) or not last_code:
            code, tipus, cikk = form_code, form_tipus, form_cikk
        else:
            code = last_code
            tipus = getattr(self, 'last_added_tipus', "") or ""
            cikk = getattr(self, 'last_added_cikk', "") or ""

        if not code or code == "R18-":
            self.trigger_toast("⚠️ Nincs nyomtatható termék adat!")
            return

        printers = [p[2] for p in win32print.EnumPrinters(win32print.PRINTER_ENUM_LOCAL | win32print.PRINTER_ENUM_CONNECTIONS)] if WIN32_AVAILABLE else []
        if printers:
            selected = printers[0]
            for p in printers:
                if "zebra" in p.lower() or "gk420" in p.lower():
                    selected = p
                    break

            if self.add_r_zpl.isChecked():
                zpl_data = generate_zpl_code(code, cikkszam=cikk, tipus=tipus, copies=copies)
                ok, msg = send_zpl_to_printer(selected, zpl_data)
                self.trigger_toast(f"🖨️ ZPL Címke kiküldve [{code}] - {copies} db" if ok else f"❌ Nyomtatási hiba: {msg}")
            else:
                _, bar_b = generate_codes_in_memory(code, cikkszam_text=cikk, tipus_text=tipus)
                if bar_b:
                    ok, msg = send_png_bytes_to_printer(selected, bar_b, copies=copies)
                    self.trigger_toast(f"🖨️ PNG Címke kinyomtatva [{code}] - {copies} db" if ok else f"❌ Nyomtatási hiba: {msg}")
            log_event(code, "CÍMKE NYOMTATÁS", f"Új Termék fül - {copies} db")
        else:
            self.trigger_toast("❌ Nem található nyomtató!")

    # -------------------------------------------------------------------
    # TAB 3: PDA / ELADÁS TAB
    # -------------------------------------------------------------------
    def init_pda_tab(self):
        layout = QVBoxLayout(self.tab_pda)
        pda_layout = QHBoxLayout()
        self.pda_input = QLineEdit()
        self.pda_input.setPlaceholderText("Pittyents ide a PDA olvasóval...")
        self.pda_input.returnPressed.connect(self.execute_pda_search)
        attach_scanner_fix(self.pda_input)
        pda_layout.addWidget(self.pda_input)

        pda_btn = QPushButton("📱 PDA Keresés")
        pda_btn.clicked.connect(
            lambda: self.run_button_action(
                pda_btn, 
                "📱 PDA Keresés", 
                self.execute_pda_search
            )
        )
        pda_layout.addWidget(pda_btn)
        layout.addLayout(pda_layout)

        self.pda_results_area = QWidget()
        self.pda_results_layout = QVBoxLayout(self.pda_results_area)
        self.pda_results_layout.setAlignment(Qt.AlignmentFlag.AlignTop)

        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setWidget(self.pda_results_area)
        layout.addWidget(scroll)

    def _is_text_input_widget(self, w):
        if w is None:
            return False
        if isinstance(w, (QLineEdit, QTextEdit, QPlainTextEdit, QAbstractSpinBox)):
            return True
        if isinstance(w, QComboBox) and w.isEditable():
            return True
        return False

    def eventFilter(self, obj, event):
        """PDA / Eladás fül: a vonalkód-olvasó (billentyűzetként gépel) akkor is a keresőmezőbe ír,
        ha éppen nem abban áll a kurzor (pl. egy gombra kattintottunk előtte)."""
        try:
            if (event.type() == QEvent.Type.KeyPress and hasattr(self, "pda_input")
                    and self.tabs.currentWidget() is self.tab_pda
                    and QApplication.activeModalWidget() is None
                    and QApplication.activeWindow() is self):
                focus = QApplication.focusWidget()
                if focus is not self.pda_input and not self._is_text_input_widget(focus):
                    txt = event.text()
                    mods = event.modifiers() & (Qt.KeyboardModifier.ControlModifier | Qt.KeyboardModifier.AltModifier |
                                                Qt.KeyboardModifier.MetaModifier)
                    if txt and txt.isprintable() and not mods:
                        # Új csippantás: a mező tartalmát lecseréli, a többi karakter már közvetlenül ide jön
                        self.pda_input.setFocus()
                        self.pda_input.setText(txt)
                        return True
        except Exception:
            pass
        return super().eventFilter(obj, event)

    def open_sale_dialog(self, platform, code):
        """💰 Eladás gomb a PDA kártyán: bekéri a vevő / ár / darab / tracking adatokat és rögzíti az eladást."""
        try:
            conn = get_db_connection()
            cursor = conn.cursor()
            cursor.execute(f"SELECT {', '.join(ROW_COLS)} FROM termekek WHERE A = ? OR B = ?", (code, code))
            r = cursor.fetchone()
            conn.close()
        except Exception as e:
            self.trigger_toast(f"❌ Adatbázis hiba: {e}")
            return
        if not r:
            self.trigger_toast(f"⚠️ {code} nem található a raktárban!")
            return
        product = {c: safe_str(r[c]) for c in ROW_COLS}
        f_match = re.search(r"\d+", product.get("F", ""))
        stock = int(f_match.group(0)) if f_match else 1

        dlg = SaleDialog(platform, product, stock, self)
        if dlg.exec() != QDialog.DialogCode.Accepted:
            self.pda_input.setFocus()
            return
        v = dlg.values()
        self.record_sale(platform, product, stock, v)

    def store_sale(self, platform, product, stock, v, eladas_id=None, sale_src=""):
        """Eladás mentése (eladasok tábla -> Sheets), készlet levonása és naplózás.
        Visszaad: a készletről szóló üzenet (pl. ' | Maradt: 2 db'), hiba esetén None."""
        code = product.get("A", "").strip().upper()
        eladas_id = eladas_id or f"E{datetime.datetime.now().strftime('%Y%m%d%H%M%S')}-{code}-{os.urandom(2).hex()}"
        try:
            conn = get_db_connection()
            cursor = conn.cursor()
            cursor.execute("""
                INSERT INTO eladasok (eladas_id, datum, platform, A, C, D, E, G, L, M, N,
                                      vevo, ar, darab, tracking, megjegyzes, sheets_synced, valtozat)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 0, 0)
            """, (eladas_id, v["datum"], platform, code, product.get("C", ""), product.get("D", ""),
                  product.get("E", ""), product.get("G", ""), product.get("L", ""), product.get("M", ""),
                  product.get("N", ""), v["vevo"], v["ar"], str(v["darab"]), v["tracking"], v["megjegyzes"]))
            conn.commit()
            conn.close()
        except sqlite3.IntegrityError:
            return ""          # ez az eladás már rögzítve van ezen a gépen
        except Exception as e:
            self.trigger_toast(f"❌ Eladás mentési hiba: {e}")
            return None

        cim = SALES_SHEETS[platform]["cim"]
        sale_details = (f"{v['darab']} db | Vevő: {v['vevo'] or '(később)'} | Ár: {v['ar'] or '(később)'} | "
                        f"Tracking: {v['tracking'] or '(később)'}")
        elotte = get_row_snapshot(code) or {c: product.get(c, "") for c in NAPLO_COLS}

        # Készlet levonása
        stock_msg = ""
        if v["reduce_stock"] and stock is not None:
            remaining = stock - v["darab"]
            if remaining > 0:
                try:
                    conn = get_db_connection()
                    cursor = conn.cursor()
                    cursor.execute("UPDATE termekek SET F = ? WHERE A = ? OR B = ?", (str(remaining), code, code))
                    conn.commit()
                    conn.close()
                    enqueue_row_from_db(code, force=True)
                    flush_outbox_async()
                    stock_msg = f" | Maradt: {remaining} db"
                except Exception as e:
                    stock_msg = f" | ❌ Készlet hiba: {e}"
                log_event(code, f"ELADÁS: {cim}", sale_details + stock_msg, elotte, get_row_snapshot(code), sale_src)
            else:
                log_event(code, f"ELADÁS: {cim}", sale_details + " | Elfogyott - törölve a raktárból", elotte, None, sale_src)
                self.delete_item_action(code, reason=f"Eladás ({cim}) után elfogyott - 0 db maradt", forras=sale_src)
                stock_msg = " | Elfogyott - törölve a raktárból"
        elif v["reduce_stock"]:
            stock_msg = " | A termék nincs meg ezen a gépen - készlet nem változott"
            log_event(code, f"ELADÁS: {cim}", sale_details + stock_msg, elotte, elotte, sale_src)
        else:
            log_event(code, f"ELADÁS: {cim}", sale_details + " | Készlet nem változott (a levonás ki volt kapcsolva)",
                      elotte, elotte, sale_src)
        return stock_msg

    def poll_pda_sales(self, manual=False):
        """A PDA-n (Zebra TC27) rögzített eladásokat a VPS-ről átveszi (háttérszálon). Egy eladást csak
        egy gép vehet át, így a készlet akkor sem vonódik le kétszer, ha több gépen fut a program."""
        if self._pda_poll_busy:
            return
        self._pda_poll_busy = True

        def job():
            claimed, claimed_changes, waiting = [], [], 0
            status, is_error = "", False
            try:
                res = vps_request("GET", f"{VPS_URL}/pda/sales", params={"pending": 1, "limit": 50}, timeout=8)
                if res.status_code == 200:
                    sales = res.json().get("sales", [])
                    waiting = len(sales)
                    gep = GEP_NEV
                    for sale in sales:
                        r = vps_request("POST", f"{VPS_URL}/pda/sales/claim",
                                        json={"eladas_id": sale.get("eladas_id"), "gep": gep}, timeout=8)
                        if r.status_code == 200:
                            claimed.append(sale)
                    # Áthelyezések / módosítások
                    res2 = vps_request("GET", f"{VPS_URL}/pda/changes", params={"pending": 1, "limit": 50}, timeout=8)
                    if res2.status_code == 200:
                        changes = res2.json().get("changes", [])
                        waiting += len(changes)
                        for ch in changes:
                            r = vps_request("POST", f"{VPS_URL}/pda/changes/claim",
                                            json={"valtozas_id": ch.get("valtozas_id"), "gep": gep}, timeout=8)
                            if r.status_code == 200:
                                claimed_changes.append(ch)
                    status = (f"📱 PDA: {len(claimed)} eladás, {len(claimed_changes)} áthelyezés átvéve" if waiting
                              else "📱 PDA: nincs új átvennivaló (eladás / áthelyezés)")
                elif res.status_code == 401:
                    status, is_error = (f"🔑 PDA eladások: a szerver nem fogadja el a jelszót - nézd meg az "
                                        f"api_token.txt-t ({API_TOKEN_FILE})" if API_TOKEN else
                                        f"🔑 PDA eladások: a szerver jelszót kér, de nincs api_token.txt ({API_TOKEN_FILE})"), True
                elif res.status_code == 404:
                    status, is_error = ("❌ PDA eladások: a VPS-en még a régi server_api.py fut "
                                        "(nincs /pda/sales) - töltsd fel az újat és indítsd újra a szervert"), True
                else:
                    status, is_error = f"❌ PDA eladások: szerver hiba ({res.status_code})", True
            except Exception as e:
                status, is_error = f"❌ PDA eladások: nem érem el a szervert ({e.__class__.__name__})", True
            try:
                push_naplo_to_server()
            except Exception:
                pass
            try:
                serve_pda_photo_requests()
            except Exception as e:
                print(f"[PDA fotó] Hiba: {e}")
            try:
                leg = send_heartbeat()
                if leg is not None:
                    self.version_info_signal.emit(leg)
            except Exception:
                pass
            finally:
                self._pda_poll_busy = False
            self.pda_poll_status_signal.emit(status, is_error, manual)
            if claimed:
                self.pda_sales_signal.emit(claimed)
            if claimed_changes:
                self.pda_changes_signal.emit(claimed_changes)

        threading.Thread(target=job, daemon=True).start()

    def on_pda_poll_status(self, text, is_error, manual):
        """A PDA-átvétel állapota a Folyamat naplóba (csak ha változott) és kézi indításnál felugró üzenetként."""
        if manual:
            self.trigger_toast(text)
        if text != self._pda_last_status:
            self._pda_last_status = text
            if is_error or manual:
                self.append_process_log(text, "#ff5252" if is_error else "#00e676")
        if hasattr(self, "lbl_pda_status"):
            self.lbl_pda_status.setText(f"{text}  ({datetime.datetime.now().strftime('%H:%M:%S')})")
            self.lbl_pda_status.setStyleSheet(f"color: {'#ff5252' if is_error else '#aaa'}; padding: 2px 4px;")

    def apply_pda_sales(self, sales):
        """Az átvett PDA-eladások rögzítése ezen a gépen: Eladottak + Sheets eladási lap, készlet levonása."""
        for sale in sales:
            try:
                platform = str(sale.get("platform", "")).strip().upper()
                if platform not in SALES_SHEETS:
                    report_pda_result("eladas", sale.get("eladas_id"), False, f"Ismeretlen platform: {platform}", "A-23")
                    continue
                code = norm_code(sale.get("A"))
                local = get_row_snapshot(code)
                if local:
                    product = dict(local)
                    f_match = re.search(r"\d+", product.get("F", ""))
                    stock = int(f_match.group(0)) if f_match else 1
                else:
                    product = {c: safe_str(sale.get(c, "")) for c in ROW_COLS}
                    stock = None
                product["A"] = code
                try:
                    darab = max(1, int(str(sale.get("darab", "1")).strip() or 1))
                except ValueError:
                    darab = 1
                v = {"datum": safe_str(sale.get("datum")), "vevo": safe_str(sale.get("vevo")),
                     "ar": safe_str(sale.get("ar")), "darab": darab, "tracking": safe_str(sale.get("tracking")),
                     "megjegyzes": safe_str(sale.get("megjegyzes")),
                     "reduce_stock": bool(int(sale.get("keszlet_levonas", 1) or 0))}
                cim = SALES_SHEETS[platform]["cim"]
                eszkoz = safe_str(sale.get("eszkoz")) or "PDA"
                stock_msg = self.store_sale(platform, product, stock, v, eladas_id=sale.get("eladas_id"),
                                            sale_src=f"📱 Zebra PDA ({eszkoz}) - 💰 {cim} gomb")
                if stock_msg is None:
                    report_pda_result("eladas", sale.get("eladas_id"), False,
                                      "Az eladást nem sikerült elmenteni a gépen (helyi adatbázis hiba)", "A-22")
                    continue
                if stock is None and v["reduce_stock"]:
                    report_pda_result("eladas", sale.get("eladas_id"), False,
                                      "Az eladás rögzítve, de a készlet NEM változott: a termék nincs meg ezen a gépen", "A-21")
                else:
                    report_pda_result("eladas", sale.get("eladas_id"), True,
                                      f"Rögzítve ({cim}, {darab} db){stock_msg}")
                self.trigger_toast(f"📱 PDA eladás: {code} ({cim}, {darab} db){stock_msg}")
                self.append_process_log(f"📱 PDA eladás átvéve: {code} - {cim} - {darab} db{stock_msg}", "#00e676")
            except Exception as e:
                report_pda_result("eladas", sale.get("eladas_id"), False, f"Váratlan hiba: {e}", "A-24")
                self.append_process_log(f"❌ [A-24] PDA eladás átvételi hiba ({sale.get('eladas_id')}): {e}", "#ff5252")
        sync_sales_async()
        if hasattr(self, "refresh_sold_table"):
            self.refresh_sold_table()

    # A PDA-ról módosítható termék-mezők (a szerver is ugyanezt engedi)
    PDA_EDITABLE_COLS = ("C", "D", "E", "F", "G", "H", "I", "K", "L", "M", "N", "O", "P")

    def apply_pda_changes(self, changes):
        """A PDA-n kért módosítások végrehajtása ezen a gépen, majd az eredmény visszaküldése a PDA-nak:
          ATHELYEZES / MODOSITAS  - termék mezői (helyi adatbázis -> Napló -> outbox: VPS + Google Sheets)
          ELADAS_MODOSITAS        - PDA-s eladás pótlása / javítása (darabszám-változásnál készlet-korrekció)
          ELADAS_TORLES           - PDA-s eladás visszavonása (kérésre a készlet visszaírásával)"""
        done = 0
        for ch in changes:
            vid = ch.get("valtozas_id")
            tipus = str(ch.get("tipus", "")).upper()
            try:
                if tipus in ("ELADAS_MODOSITAS", "ELADAS_TORLES"):
                    ok, msg, kod = self._apply_pda_sale_change(ch)
                elif tipus == "TORLES":
                    ok, msg, kod = self._apply_pda_delete(ch)
                elif tipus == "VISSZAALLITAS":
                    ok, msg, kod = self._apply_pda_restore(ch)
                else:
                    ok, msg, kod = self._apply_pda_product_change(ch)
                report_pda_result("valtozas", vid, ok, msg, kod)
                if ok:
                    done += 1
                    self.append_process_log(f"📱 PDA {tipus.lower()}: {ch.get('A')} - {msg}", "#00e676")
                else:
                    self.append_process_log(f"❌ [{kod}] PDA {tipus.lower()}: {ch.get('A')} - {msg}", "#ff5252")
            except Exception as e:
                report_pda_result("valtozas", vid, False, f"Váratlan hiba: {e}", "A-24")
                self.append_process_log(f"❌ [A-24] PDA módosítás átvételi hiba ({vid}): {e}", "#ff5252")
        if done:
            flush_outbox_async()
            sync_sales_async()
            self.trigger_toast(f"📱 PDA: {done} módosítás végrehajtva")
            self.silent_update_excel_table()
            if hasattr(self, "refresh_sold_table"):
                self.refresh_sold_table()

    @staticmethod
    def _pda_json(val):
        try:
            d = json.loads(val) if isinstance(val, str) else (val or {})
            return d if isinstance(d, dict) else {}
        except Exception:
            return {}

    def _apply_pda_product_change(self, ch):
        """Termék mező(k) módosítása. Visszaad: (sikerült?, üzenet, hibakód)."""
        code = norm_code(ch.get("A"))
        eszkoz = safe_str(ch.get("eszkoz")) or "PDA"
        adat = self._pda_json(ch.get("adat"))
        valt = adat.get("valtozasok") or {str(ch.get("mezo", "")).strip().upper(): safe_str(ch.get("uj"))}
        regi_pda = adat.get("regi") or {str(ch.get("mezo", "")).strip().upper(): safe_str(ch.get("regi"))}
        valt = {str(k).upper(): safe_str(v) for k, v in valt.items()}
        bad = [k for k in valt if k not in self.PDA_EDITABLE_COLS]
        if bad or not valt:
            return False, f"Ez a mező nem módosítható a PDA-ról: {', '.join(bad) or '-'}", "A-23"
        elotte = get_row_snapshot(code)
        if not elotte:
            return False, f"{code} nincs meg ezen a gépen - nem módosítottam", "A-21"
        try:
            conn = get_db_connection()
            sets = ", ".join(f"{k} = ?" for k in valt)
            conn.execute(f"UPDATE termekek SET {sets} WHERE UPPER(TRIM(A)) = ? OR UPPER(TRIM(B)) = ?",
                         (*valt.values(), code, code))
            conn.commit()
            conn.close()
        except Exception as e:
            return False, f"Helyi adatbázis hiba: {e}", "A-22"
        utana = get_row_snapshot(code)
        elteres = [k for k in valt if safe_str(elotte.get(k)) != safe_str(regi_pda.get(k, elotte.get(k)))]
        megj = (f" (figyelem: a gépen közben más volt itt: "
                f"{', '.join(f'{k}={safe_str(elotte.get(k))!r}' for k in elteres)})") if elteres else ""
        athelyezes = list(valt) == ["L"]
        visszavonas = str(ch.get("tipus", "")).upper() == "VISSZAVONAS"
        log_event(code, "VISSZAVONÁS (PDA)" if visszavonas else ("ÁTHELYEZÉS (PDA)" if athelyezes else "MÓDOSÍTÁS (PDA)"),
                  describe_changes(elotte, utana) + megj, elotte, utana,
                  f"📱 Zebra PDA ({eszkoz}) - " + ("↩️ Visszavonás (véletlen nyomás javítása)" if visszavonas
                                                   else ('Áthelyezés' if athelyezes else 'Szerkesztés')))
        enqueue_row_from_db(code, force=True)
        return True, " | ".join(f"{COLUMN_LABELS.get(k, k)}: {safe_str(elotte.get(k)) or '-'} → {v or '-'}"
                                for k, v in valt.items()) + megj, ""

    def _apply_pda_delete(self, ch):
        """Termék törlése a PDA kérésére: ToroltSor-ra archiválja, majd törli a gépen, a szerveren és a táblázatban."""
        code = norm_code(ch.get("A"))
        eszkoz = safe_str(ch.get("eszkoz")) or "PDA"
        adat = self._pda_json(ch.get("adat"))
        ok_txt = safe_str(adat.get("ok")) or "PDA-ról törölve"
        src = f"📱 Zebra PDA ({eszkoz}) - 🗑️ Törlés"
        elotte = get_row_snapshot(code)
        try:
            if elotte:
                archive_row_from_db(code, f"PDA törlés ({eszkoz}): {ok_txt}")
                conn = get_db_connection()
                conn.execute("DELETE FROM termekek WHERE UPPER(TRIM(A)) = ? OR UPPER(TRIM(B)) = ?", (code, code))
                conn.commit()
                conn.close()
                msg = "Törölve: gép + szerver + Google táblázat (adatai a ToroltSor lapra kerültek)"
            else:
                sor = adat.get("sor") or {}
                msg = "Ezen a gépen nem volt meg – a szerverről és a táblázatból törölve"
                if _row_has_product_data(sor):
                    archive_deleted_row({c: safe_str(sor.get(c, "")) for c in ROW_COLS}, f"PDA törlés ({eszkoz}): {ok_txt}")
                    msg += " (adatai a ToroltSor lapra kerültek)"
            enqueue_delete(code)
        except Exception as e:
            return False, f"Törlési hiba: {e}", "A-22"
        log_event(code, "TÖRLÉS (PDA)", ok_txt, elotte, None, src)
        return True, msg, ""

    def _apply_pda_restore(self, ch):
        """Egy PDA-s törlés visszavonása: a termék visszakerül ugyanazokkal az adatokkal."""
        code = norm_code(ch.get("A"))
        eszkoz = safe_str(ch.get("eszkoz")) or "PDA"
        sor = self._pda_json(ch.get("adat")).get("sor") or {}
        if not sor:
            return False, "Nincs mentett adat a visszaállításhoz", "A-23"
        data = {c: safe_str(sor.get(c, "")) for c in NAPLO_COLS}
        data["A"] = code
        now = get_row_snapshot(code)
        try:
            conn = get_db_connection()
            cur = conn.cursor()
            cur.execute(f"UPDATE termekek SET {', '.join(c + '=?' for c in NAPLO_COLS)} WHERE UPPER(TRIM(A))=? OR UPPER(TRIM(B))=?",
                        (*[data[c] for c in NAPLO_COLS], code, code))
            if cur.rowcount == 0:
                cur.execute(f"INSERT INTO termekek ({', '.join(NAPLO_COLS)}, szin) VALUES ({', '.join('?' * 16)}, 'FEHÉR')",
                            [data[c] for c in NAPLO_COLS])
            conn.commit()
            conn.close()
        except Exception as e:
            return False, f"Helyi adatbázis hiba: {e}", "A-22"
        payload = dict(data); payload["_force"] = True
        enqueue_outbox(payload)
        log_event(code, "FELVÉTEL (PDA-s törlés visszavonva)", "", now, get_row_snapshot(code),
                  f"📱 Zebra PDA ({eszkoz}) - ↩️ Törlés visszavonása")
        return True, "Visszaállítva: gép + szerver + Google táblázat", ""

    def _change_stock(self, code, delta, src, reason):
        """Készlet (F) módosítása delta-val. 0 alá / 0-ra csökkenésnél a terméket törli (mint az eladás).
        Visszaad: üzenet, vagy None, ha a termék nincs meg ezen a gépen."""
        elotte = get_row_snapshot(code)
        if not elotte:
            return None
        m = re.search(r"\d+", elotte.get("F", ""))
        stock = int(m.group(0)) if m else 1
        remaining = stock + delta
        if remaining <= 0:
            log_event(code, "KÉSZLET (PDA)", f"{reason} | {stock} → 0 db - elfogyott, törölve a raktárból", elotte, None, src)
            self.delete_item_action(code, reason=f"{reason} - 0 db maradt", forras=src)
            return f"Készlet: {stock} → 0 db (elfogyott, törölve)"
        conn = get_db_connection()
        conn.execute("UPDATE termekek SET F = ? WHERE UPPER(TRIM(A)) = ? OR UPPER(TRIM(B)) = ?", (str(remaining), code, code))
        conn.commit()
        conn.close()
        log_event(code, "KÉSZLET (PDA)", f"{reason} | {stock} → {remaining} db", elotte, get_row_snapshot(code), src)
        enqueue_row_from_db(code, force=True)
        return f"Készlet: {stock} → {remaining} db"

    def _apply_pda_sale_change(self, ch):
        """PDA-s eladás javítása / visszavonása ezen a gépen. Visszaad: (sikerült?, üzenet, hibakód)."""
        eid = safe_str(ch.get("ref"))
        code = norm_code(ch.get("A"))
        eszkoz = safe_str(ch.get("eszkoz")) or "PDA"
        adat = self._pda_json(ch.get("adat"))
        src = f"📱 Zebra PDA ({eszkoz}) - Napló fül"
        conn = get_db_connection()
        r = conn.execute("SELECT * FROM eladasok WHERE eladas_id = ?", (eid,)).fetchone()
        conn.close()
        if not r:
            return False, f"Az eladás ({eid}) nincs meg ezen a gépen", "A-21"
        sale = dict(r)
        levont = bool(int(adat.get("keszlet_levonas", 1) or 0))

        if str(ch.get("tipus")).upper() == "ELADAS_TORLES":
            try:
                conn = get_db_connection()
                conn.execute("UPDATE eladasok SET torolve = 1, sheets_synced = 0, valtozat = valtozat + 1 WHERE eladas_id = ?", (eid,))
                conn.execute("DELETE FROM eladasok WHERE eladas_id = ? AND (synced_platform IS NULL OR synced_platform = '')", (eid,))
                conn.commit()
                conn.close()
            except Exception as e:
                return False, f"Helyi adatbázis hiba: {e}", "A-22"
            log_event(eid, "ELADÁS VISSZAVONÁSA (PDA)", f"{code} | {sale.get('darab')} db | {sale.get('platform')}", forras=src)
            msg = "Eladás visszavonva (az eladási lapról is törlődik)"
            if adat.get("keszlet_vissza") and levont:
                try:
                    darab = int(str(sale.get("darab") or adat.get("darab") or 1))
                except ValueError:
                    darab = 1
                st = self._change_stock(code, +darab, src, f"Eladás visszavonása ({eid})")
                if st is None:
                    return False, (msg + f" - DE a készletet nem tudtam visszaírni: {code} már nincs a raktárban "
                                   "(elfogyott és törlődött). Állítsd vissza az Admin fülön: Törölt termékek → Visszaállítás."), "A-25"
                msg += f" | {st}"
            return True, msg, ""

        # ELADAS_MODOSITAS
        valt = {k: safe_str(v) for k, v in (adat.get("valtozasok") or {}).items()
                if k in ("vevo", "ar", "tracking", "megjegyzes", "darab", "platform")}
        if not valt:
            return True, "Nincs változás", ""
        if "platform" in valt and valt["platform"].upper() not in SALES_SHEETS:
            return False, f"Ismeretlen platform: {valt['platform']}", "A-23"
        try:
            conn = get_db_connection()
            sets = ", ".join(f"{k} = ?" for k in valt)
            conn.execute(f"UPDATE eladasok SET {sets}, sheets_synced = 0, valtozat = valtozat + 1 WHERE eladas_id = ?",
                         (*valt.values(), eid))
            conn.commit()
            conn.close()
        except Exception as e:
            return False, f"Helyi adatbázis hiba: {e}", "A-22"
        labels = dict(SALES_COLUMNS)
        reszl = " | ".join(f"{labels.get(k, k)}: {safe_str(sale.get(k)) or '∅'} → {v or '∅'}" for k, v in valt.items())
        log_event(eid, "ELADÁS MÓDOSÍTÁSA (PDA)", f"{code} | {reszl}", forras=src)
        msg = reszl
        if "darab" in valt and levont:
            try:
                diff = int(valt["darab"]) - int(safe_str(sale.get("darab")) or 1)
            except ValueError:
                diff = 0
            if diff:
                st = self._change_stock(code, -diff, src, f"Eladott darabszám javítása ({eid})")
                if st is None:
                    return False, (msg + f" - DE a készletet nem tudtam javítani: {code} már nincs a raktárban. "
                                   "Ha kell, állítsd vissza az Admin fülön (Törölt termékek)."), "A-25"
                msg += f" | {st}"
        return True, msg, ""

    def record_sale(self, platform, product, stock, v):
        code = product.get("A", "").strip().upper()
        cim = SALES_SHEETS[platform]["cim"]
        stock_msg = self.store_sale(platform, product, stock, v, sale_src=f"PDA / Eladás fül - 💰 {cim} gomb")
        if stock_msg is None:
            return

        sync_sales_async()
        self.trigger_toast(f"💰 {code} eladva ({cim}, {v['darab']} db){stock_msg}")

        # PDA fül: visszaigazolás, és a kurzor kész a következő csippantásra
        while self.pda_results_layout.count():
            item = self.pda_results_layout.takeAt(0)
            if item.widget():
                item.widget().deleteLater()
        lbl = QLabel(f"✅ <b>{code}</b> eladva — {cim} — {v['darab']} db{stock_msg}<br>"
                     f"<span style='color:#aaa'>Az adatok az 💰 Eladottak fülön módosíthatók.</span>")
        lbl.setStyleSheet("background: #1b5e20; color: white; padding: 14px; border-radius: 6px; font-size: 15px;")
        self.pda_results_layout.addWidget(lbl)
        self.pda_input.setFocus()
        self.pda_input.selectAll()

        if hasattr(self, "refresh_sold_table"):
            self.refresh_sold_table()

    # -------------------------------------------------------------------
    # 💰 ELADOTTAK FÜL: minden programban rögzített eladás - látható és szerkeszthető
    # -------------------------------------------------------------------
    def init_sold_tab(self):
        layout = QVBoxLayout(self.tab_sold)

        bar = QHBoxLayout()
        self.sold_search = QLineEdit()
        self.sold_search.setPlaceholderText("🔍 Keresés (R18, vevő, típus, tracking...)")
        self.sold_search.textChanged.connect(self.refresh_sold_table)
        bar.addWidget(self.sold_search, stretch=3)

        self.sold_platform_filter = QComboBox()
        self.sold_platform_filter.addItem("Minden platform", "")
        for pf, cfg in SALES_SHEETS.items():
            self.sold_platform_filter.addItem(cfg["cim"], pf)
        self.sold_platform_filter.currentIndexChanged.connect(self.refresh_sold_table)
        bar.addWidget(self.sold_platform_filter, stretch=1)

        self.sold_only_missing = QCheckBox("Csak hiányos (⏭ később)")
        self.sold_only_missing.toggled.connect(self.refresh_sold_table)
        bar.addWidget(self.sold_only_missing)

        btn_refresh = QPushButton("🔄 Frissítés")
        btn_refresh.clicked.connect(self.refresh_sold_table)
        bar.addWidget(btn_refresh)

        btn_pull = QPushButton("⬇️ Behozás a táblázatból")
        btn_pull.setStyleSheet("background-color: #ffb74d; color: black; font-weight: bold;")
        btn_pull.setToolTip("A Google táblázat Eladas_Ebay / Eladas_EbayKivul / Eladas_Webshop lapjairól behozza\n"
                            "a régi és a más gépen rögzített eladásokat (magától is megteszi 5 percenként).\n"
                            "A készlet ettől NEM változik.")
        btn_pull.clicked.connect(lambda: self.pull_sales(manual=True))
        bar.addWidget(btn_pull)

        btn_sync = QPushButton("☁️ Szinkron most")
        btn_sync.setStyleSheet("background-color: #00e676; color: black; font-weight: bold;")
        btn_sync.clicked.connect(lambda: (sync_sales_async(), self.trigger_toast("☁️ Eladások feltöltése elindítva...")))
        bar.addWidget(btn_sync)

        btn_pda = QPushButton("📱 PDA eladások átvétele most")
        btn_pda.setStyleSheet("background-color: #29b6f6; color: black; font-weight: bold;")
        btn_pda.setToolTip("A Zebra PDA-n rögzített eladásokat azonnal lekéri a szerverről\n"
                           "(magától is megteszi 20 másodpercenként).")
        btn_pda.clicked.connect(lambda: self.poll_pda_sales(manual=True))
        bar.addWidget(btn_pda)

        btn_del = QPushButton("🗑️ Kijelölt eladás törlése")
        btn_del.setStyleSheet("background-color: #ff5252; color: white; font-weight: bold;")
        btn_del.clicked.connect(self.delete_selected_sale)
        bar.addWidget(btn_del)
        layout.addLayout(bar)

        self.sold_status_lbl = QLabel("")
        self.sold_status_lbl.setStyleSheet("color: #ffb74d; font-weight: bold;")
        layout.addWidget(self.sold_status_lbl)
        self.lbl_sold_pull = QLabel("⬇️ Eladások a táblázatból: még nem volt behozás")
        self.lbl_sold_pull.setStyleSheet("color: #aaa; padding: 2px 4px;")
        layout.addWidget(self.lbl_sold_pull)
        self._sold_pull_busy = False
        self._sold_pull_last = 0.0
        self.sold_pull_signal.connect(self.on_sales_pulled)
        self.sold_pull_timer = QTimer(self)
        self.sold_pull_timer.timeout.connect(self.pull_sales)
        self.sold_pull_timer.start(5 * 60 * 1000)
        QTimer.singleShot(25000, self.pull_sales)

        # Zebra PDA: az utolsó átvételi próbálkozás eredménye (hiba esetén pirosan)
        self.lbl_pda_status = QLabel(self._pda_last_status or "📱 PDA eladások: még nem volt ellenőrzés")
        self.lbl_pda_status.setStyleSheet("color: #aaa; padding: 2px 4px;")
        layout.addWidget(self.lbl_pda_status)

        self.sold_table = QTableWidget()
        self.sold_table.setColumnCount(len(SALES_COLUMNS) + 1)
        self.sold_table.setHorizontalHeaderLabels([h for _, h in SALES_COLUMNS] + ["☁️"])
        self.sold_table.setAlternatingRowColors(True)
        self.sold_table.setWordWrap(True)
        self.sold_table.verticalHeader().setDefaultSectionSize(36)
        self.sold_table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Interactive)
        self.sold_table.setStyleSheet("""
            QTableWidget { background-color: #21252b; alternate-background-color: #282c34; color: #e0e0e0;
                           gridline-color: #3b4048; font-size: 13px; }
            QHeaderView::section { background-color: #1b1e23; color: #61afef; font-weight: bold; padding: 6px; border: 1px solid #3b4048; }
            QTableWidget::item:selected { background-color: #2f5f8f; color: #ffffff; }
        """)
        widths = {"eladas_id": 150, "datum": 130, "platform": 170, "A": 100, "C": 110, "D": 200, "E": 180,
                  "G": 90, "L": 100, "M": 60, "N": 110, "vevo": 160, "ar": 90, "darab": 60, "tracking": 170, "megjegyzes": 200}
        for i, (key, _) in enumerate(SALES_COLUMNS):
            self.sold_table.setColumnWidth(i, widths.get(key, 100))
        self.sold_table.setColumnWidth(len(SALES_COLUMNS), 40)
        self.sold_table.itemChanged.connect(self.on_sold_cell_changed)
        layout.addWidget(self.sold_table)

        # Időnként újrapróbálja feltölteni, ami net hiánya miatt nem ment ki
        self.sales_sync_timer = QTimer()
        self.sales_sync_timer.setInterval(30000)
        self.sales_sync_timer.timeout.connect(sync_sales_async)
        self.sales_sync_timer.start()
        QTimer.singleShot(3000, sync_sales_async)

        # A szinkron állapot (☁️ oszlop) frissítése, ha épp ezen a fülön vagyunk
        self.sold_status_timer = QTimer()
        self.sold_status_timer.setInterval(10000)
        self.sold_status_timer.timeout.connect(
            lambda: self.refresh_sold_table() if self.tabs.currentWidget() is self.tab_sold
            and self.sold_table.state() != QAbstractItemView.State.EditingState else None)
        self.sold_status_timer.start()

        # Vonalkód-olvasó elkapása a PDA fülön (lásd eventFilter)
        QApplication.instance().installEventFilter(self)

        self.refresh_sold_table()

    def pull_sales(self, manual=False):
        """A táblázat eladási lapjairól behozza a régi / más gépen rögzített eladásokat (háttérszálon)."""
        if self._sold_pull_busy:
            return
        self._sold_pull_busy = True
        self._sold_pull_last = time.time()
        if manual:
            self.lbl_sold_pull.setText("⏳ Eladások behozása a táblázatból...")

        def job():
            try:
                res = pull_sales_from_sheets()
            except Exception as e:
                res = {"uj": 0, "frissitett": 0, "torolt": 0, "id_potolva": 0, "hibak": [str(e)]}
            self.sold_pull_signal.emit(res, manual)
        threading.Thread(target=job, daemon=True).start()

    def on_sales_pulled(self, res, manual):
        self._sold_pull_busy = False
        valt = res["uj"] + res["frissitett"] + res["torolt"]
        txt = (f"⬇️ Eladások a táblázatból ({datetime.datetime.now().strftime('%H:%M')}): "
               f"{res['uj']} új, {res['frissitett']} frissítve, {res['torolt']} törölve"
               + (f", {res['id_potolva']} régi sor Eladás ID-t kapott" if res["id_potolva"] else ""))
        if res["hibak"]:
            txt += "  |  ❌ " + "; ".join(res["hibak"])[:300]
        self.lbl_sold_pull.setText(txt)
        self.lbl_sold_pull.setStyleSheet("color: %s; padding: 2px 4px;" % ("#ff5252" if res["hibak"] else "#aaa"))
        if valt:
            self.append_process_log(txt)
        if valt or manual:
            self.refresh_sold_table()
        if manual:
            self.trigger_toast(("❌ " + res["hibak"][0]) if res["hibak"] and not valt else
                               f"⬇️ {res['uj']} új eladás behozva, {res['frissitett']} frissítve", 3500)

    def refresh_sold_table(self, *_):
        if not hasattr(self, "sold_table"):
            return
        try:
            conn = get_db_connection()
            cursor = conn.cursor()
            cursor.execute("SELECT * FROM eladasok WHERE torolve = 0 ORDER BY datum DESC, id DESC")
            rows = [dict(r) for r in cursor.fetchall()]
            conn.close()
        except Exception as e:
            self.trigger_toast(f"❌ Eladások betöltési hiba: {e}")
            return

        q = self.sold_search.text().strip().lower()
        pf_filter = self.sold_platform_filter.currentData()
        only_missing = self.sold_only_missing.isChecked()
        missing_total = sum(1 for r in rows if any(not safe_str(r.get(k)).strip() for k in SALES_LATER_FIELDS))

        shown = []
        for r in rows:
            if pf_filter and r["platform"] != pf_filter:
                continue
            if only_missing and not any(not safe_str(r.get(k)).strip() for k in SALES_LATER_FIELDS):
                continue
            if q and q not in " ".join(safe_str(r.get(k)) for k, _ in SALES_COLUMNS).lower():
                continue
            shown.append(r)

        self.sold_table.blockSignals(True)
        self.sold_table.setRowCount(len(shown))
        later_bg = QColor("#7a4a00")
        for ri, r in enumerate(shown):
            for ci, (key, _) in enumerate(SALES_COLUMNS):
                if key == "platform":
                    combo = QComboBox()
                    for pf, cfg in SALES_SHEETS.items():
                        combo.addItem(cfg["cim"], pf)
                    idx = combo.findData(r["platform"])
                    combo.setCurrentIndex(max(idx, 0))
                    combo.currentIndexChanged.connect(
                        lambda _i, c=combo, eid=r["eladas_id"]: self.update_sale_field(eid, "platform", c.currentData()))
                    self.sold_table.setCellWidget(ri, ci, combo)
                    item = QTableWidgetItem("")
                    item.setFlags(item.flags() & ~Qt.ItemFlag.ItemIsEditable)
                else:
                    item = QTableWidgetItem(safe_str(r.get(key)))
                    if key == "eladas_id":
                        item.setFlags(item.flags() & ~Qt.ItemFlag.ItemIsEditable)
                    if key in SALES_LATER_FIELDS and not safe_str(r.get(key)).strip():
                        item.setBackground(later_bg)
                        item.setToolTip("⏭ Később kitöltendő")
                item.setData(Qt.ItemDataRole.UserRole, r["eladas_id"])
                self.sold_table.setItem(ri, ci, item)
            st = QTableWidgetItem("✅" if r["sheets_synced"] else "⏳")
            st.setToolTip("Fent van az eladási munkalapon" if r["sheets_synced"] else "Feltöltésre vár")
            st.setFlags(st.flags() & ~Qt.ItemFlag.ItemIsEditable)
            st.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
            self.sold_table.setItem(ri, len(SALES_COLUMNS), st)
        self.sold_table.blockSignals(False)

        self.sold_status_lbl.setText(
            f"💰 Összes eladás: {len(rows)} db  |  Megjelenítve: {len(shown)} db  |  "
            f"⏭ Hiányos (vevő / ár / tracking később): {missing_total} db  |  Narancs cella = később kitöltendő")

    def on_sold_cell_changed(self, item):
        col = item.column()
        if col >= len(SALES_COLUMNS):
            return
        key = SALES_COLUMNS[col][0]
        if key in ("eladas_id", "platform"):
            return
        eid = item.data(Qt.ItemDataRole.UserRole)
        if eid:
            self.update_sale_field(eid, key, item.text().strip())
            self.sold_table.blockSignals(True)
            if key in SALES_LATER_FIELDS:
                item.setBackground(QColor("#7a4a00") if not item.text().strip() else QColor(0, 0, 0, 0))
            st = self.sold_table.item(item.row(), len(SALES_COLUMNS))
            if st:
                st.setText("⏳")
            self.sold_table.blockSignals(False)

    def update_sale_field(self, eladas_id, key, value):
        allowed = {k for k, _ in SALES_COLUMNS} - {"eladas_id"}
        if key not in allowed:
            return
        old_val = ""
        try:
            conn = get_db_connection()
            cursor = conn.cursor()
            cursor.execute(f"SELECT {key} FROM eladasok WHERE eladas_id = ?", (eladas_id,))
            r = cursor.fetchone()
            old_val = safe_str(r[0]) if r else ""
            cursor.execute(f"UPDATE eladasok SET {key} = ?, sheets_synced = 0, valtozat = valtozat + 1 WHERE eladas_id = ?",
                           (value, eladas_id))
            conn.commit()
            conn.close()
        except Exception as e:
            self.trigger_toast(f"❌ Eladás módosítási hiba: {e}")
            return
        label = dict(SALES_COLUMNS).get(key, key)
        log_event(eladas_id, "ELADÁS MÓDOSÍTÁSA", f"{label}: {old_val or '∅'} → {value or '∅'}", forras="💰 Eladottak fül - cella szerkesztés")
        # Ha PDA-s eladás, a PDA Naplójában is a friss adat látsszon (pl. utólag kitöltött vevő / ár / tracking)
        def _mirror():
            try:
                vps_request("POST", f"{VPS_URL}/pda/sale/mirror", timeout=8,
                            json={"eladas_id": eladas_id, "mezo": key, "ertek": value})
            except Exception:
                pass
        threading.Thread(target=_mirror, daemon=True).start()
        # Gyors egymás utáni javításoknál csak egyszer töltünk fel (2 mp késleltetés)
        if not hasattr(self, "_sales_sync_debounce"):
            self._sales_sync_debounce = QTimer()
            self._sales_sync_debounce.setSingleShot(True)
            self._sales_sync_debounce.setInterval(2000)
            self._sales_sync_debounce.timeout.connect(sync_sales_async)
        self._sales_sync_debounce.start()

    def delete_selected_sale(self):
        row = self.sold_table.currentRow()
        if row < 0:
            self.trigger_toast("⚠️ Nincs kijelölt eladás!")
            return
        item = self.sold_table.item(row, 0)
        eid = item.data(Qt.ItemDataRole.UserRole) if item else None
        if not eid:
            return
        dlg = SilentConfirmDialog("Eladás törlése",
                                  f"Biztosan törlöd ezt az eladást?\n\n{eid}\n\n"
                                  f"Az eladási munkalapról is törlődik. (A raktárkészletet NEM állítja vissza!)", self)
        if dlg.exec() != QDialog.DialogCode.Accepted:
            return
        try:
            conn = get_db_connection()
            cursor = conn.cursor()
            cursor.execute("UPDATE eladasok SET torolve = 1, sheets_synced = 0, valtozat = valtozat + 1 WHERE eladas_id = ?", (eid,))
            # Ha még fel sem ment sehova, helyben azonnal törölhető
            cursor.execute("DELETE FROM eladasok WHERE eladas_id = ? AND (synced_platform IS NULL OR synced_platform = '')", (eid,))
            conn.commit()
            conn.close()
        except Exception as e:
            self.trigger_toast(f"❌ Törlési hiba: {e}")
            return
        log_event(eid, "ELADÁS TÖRLÉSE", "Eladás törölve az Eladottak fülön", forras="💰 Eladottak fül - 🗑️ Kijelölt eladás törlése gomb")
        sync_sales_async()
        self.refresh_sold_table()
        self.trigger_toast("🗑️ Eladás törölve")

    # -------------------------------------------------------------------
    # 🖥️ SZERVER ADATBÁZIS FÜL: a VPS-en lévő adatbázis megtekintése, javítása, törlése
    # -------------------------------------------------------------------
    SERVER_STATUS_COLS = ["Állapot"] + [f"{c}: {COLUMN_LABELS.get(c, c)}" for c in "ABCDEFGHIJKLMNOP"] + ["Verzió"]

    def init_server_tab(self):
        layout = QVBoxLayout(self.tab_server)
        bar = QHBoxLayout()
        title = QLabel("🖥️ VPS szerver adatbázis")
        title.setStyleSheet("font-size: 16px; font-weight: bold; color: #29b6f6;")
        bar.addWidget(title)

        self.server_search = QLineEdit()
        self.server_search.setPlaceholderText("🔍 Keresés (R18, márka, típus, cikkszám...) - Enter")
        self.server_search.returnPressed.connect(self.render_server_table)
        bar.addWidget(self.server_search, stretch=2)

        self.server_filter = QComboBox()
        for label in ["Minden sor", "⚠️ Csak ahol eltér a helyi géptől", "⏳ Feltöltésre vár (még nincs fent)",
                      "❌ Csak helyben van, a szerveren nincs", "🕘 Legutóbb feltöltött 200"]:
            self.server_filter.addItem(label)
        self.server_filter.currentIndexChanged.connect(self.render_server_table)
        bar.addWidget(self.server_filter, stretch=1)

        btn_load = QPushButton("🔄 Betöltés a szerverről")
        btn_load.setStyleSheet("background-color: #00e676; color: black; font-weight: bold; padding: 6px 12px;")
        btn_load.clicked.connect(self.load_server_items)
        bar.addWidget(btn_load)

        btn_del = QPushButton("🗑️ Kijelölt törlése")
        btn_del.setStyleSheet("background-color: #ff5252; color: white; font-weight: bold; padding: 6px 12px;")
        btn_del.setToolTip("A kijelölt termék(ek) törlése a szerverről, a Google táblázatból és a helyi gépről is (megerősítéssel)")
        btn_del.clicked.connect(self.delete_selected_server_items)
        bar.addWidget(btn_del)

        btn_upload = QPushButton("⬆️ Helyi adat feltöltése")
        btn_upload.setToolTip("A kijelölt sor(ok) HELYI (ezen a gépen lévő) adatát felküldi a szerverre és a Google táblázatba")
        btn_upload.clicked.connect(self.upload_selected_local_to_server)
        bar.addWidget(btn_upload)
        layout.addLayout(bar)

        self.server_status_lbl = QLabel("Nyomd meg a 🔄 Betöltés gombot. Dupla kattintással a cella szerkeszthető "
                                        "(a mentés a szerverre ÉS a Google táblázatba is kimegy).")
        self.server_status_lbl.setStyleSheet("color: #aaaaaa;")
        self.server_status_lbl.setWordWrap(True)
        layout.addWidget(self.server_status_lbl)

        self.server_table = QTableWidget()
        self.server_table.setColumnCount(len(self.SERVER_STATUS_COLS))
        self.server_table.setHorizontalHeaderLabels(self.SERVER_STATUS_COLS)
        self.server_table.setAlternatingRowColors(True)
        self.server_table.setEditTriggers(QAbstractItemView.EditTrigger.DoubleClicked | QAbstractItemView.EditTrigger.EditKeyPressed)
        self.server_table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self.server_table.setStyleSheet("""
            QTableWidget { background-color: #21252b; alternate-background-color: #282c34; color: #e0e0e0;
                           gridline-color: #3b4048; font-size: 13px; }
            QHeaderView::section { background-color: #1b1e23; color: #61afef; font-weight: bold; padding: 6px; border: 1px solid #3b4048; }
            QTableWidget::item:selected { background-color: #2f5f8f; color: #ffffff; }
        """)
        widths = [190, 95, 95, 110, 190, 180, 45, 85, 60, 70, 40, 60, 110, 55, 100, 120, 100, 60]
        for i, w in enumerate(widths):
            self.server_table.setColumnWidth(i, w)
        self.server_table.itemChanged.connect(self.on_server_cell_edited)
        layout.addWidget(self.server_table)

        self._server_items = []
        self._server_deleted = []
        self.server_items_signal.connect(self._on_server_items_loaded)

    def load_server_items(self):
        if getattr(self, "_server_loading", False):
            return
        self._server_loading = True
        self.server_status_lbl.setText("⏳ Betöltés a VPS szerverről...")

        def job():
            try:
                res = vps_request("GET", f"{VPS_URL}/items", timeout=30)
                if res.status_code == 404:
                    self.server_items_signal.emit({"error": "A szerveren még a RÉGI server_api.py fut - töltsd fel az újat "
                                                            "(vps/server_api.py) és indítsd újra a szerver szolgáltatást."})
                    return
                js = res.json()
                if js.get("status") != "ok":
                    self.server_items_signal.emit({"error": js.get("message", "Ismeretlen hiba")})
                    return
                self.server_items_signal.emit(js)
            except Exception as e:
                self.server_items_signal.emit({"error": f"A szerver nem érhető el: {e}"})
        threading.Thread(target=job, daemon=True).start()

    def _on_server_items_loaded(self, js):
        self._server_loading = False
        if "error" in js:
            self.server_status_lbl.setText(f"❌ {js['error']}")
            self.append_process_log(f"❌ Szerver adatbázis fül: {js['error']}", "#ff5252")
            return
        self._server_items = js.get("items", [])
        self._server_deleted = js.get("torolt", [])
        self._server_loaded_at = datetime.datetime.now().strftime("%H:%M:%S")
        self.render_server_table()

    def _server_rows_with_status(self):
        """Összeveti a szerver sorait a helyi adatbázissal: (állapot, szín, adat-dict, verzió) lista."""
        try:
            conn = get_db_connection()
            cursor = conn.cursor()
            cursor.execute(f"SELECT {', '.join(ROW_COLS)} FROM termekek")
            local = {}
            for r in cursor.fetchall():
                d = {c: safe_str(r[c]) for c in ROW_COLS}
                if d["A"]:
                    local[d["A"].strip().upper()] = d
            pending = get_pending_outbox_codes(cursor)
            conn.close()
        except Exception:
            local, pending = {}, set()

        out, seen = [], set()
        for it in self._server_items:
            code = str(it.get("A", "")).strip().upper()
            seen.add(code)
            srv = {c: safe_str(it.get(c)) for c in ROW_COLS}
            loc = local.get(code)
            if code in pending:
                st, col = "⏳ Helyi módosítás feltöltésre vár", "#ffb74d"
            elif loc is None:
                st, col = "➕ Csak a szerveren van", "#29b6f6"
            elif all(str(loc[c]).strip() == str(srv[c]).strip() for c in ROW_COLS):
                st, col = "✅ Egyezik a helyivel", "#00e676"
            else:
                diff = [c for c in ROW_COLS if str(loc[c]).strip() != str(srv[c]).strip()]
                st, col = f"⚠️ Eltér a helyitől: {', '.join(diff)}", "#ff9800"
            out.append((st, col, srv, int(it.get("verzio") or 0), loc))
        for code, loc in local.items():
            if code not in seen and code.startswith("R18-") and _row_has_product_data(loc):
                st = "⏳ Feltöltésre vár" if code in pending else "❌ Csak helyben van, a szerveren NINCS"
                out.append((st, "#ff5252", loc, -1, loc))
        return out

    def render_server_table(self, *_):
        if not hasattr(self, "server_table"):
            return
        if not self._server_items:
            return
        rows = self._server_rows_with_status()
        f = self.server_filter.currentIndex()
        if f == 1:
            rows = [r for r in rows if r[0].startswith("⚠️")]
        elif f == 2:
            rows = [r for r in rows if r[0].startswith("⏳")]
        elif f == 3:
            rows = [r for r in rows if r[0].startswith("❌")]
        elif f == 4:
            rows = sorted([r for r in rows if r[3] >= 0], key=lambda r: -r[3])[:200]
        q = self.server_search.text().strip().lower()
        if q:
            rows = [r for r in rows if q in " ".join(r[2].values()).lower() or q in r[0].lower()]
        if f != 4:
            def num(r):
                m = re.search(r"(\d+)", r[2].get("A", ""))
                return int(m.group(1)) if m else 0
            rows.sort(key=num, reverse=True)

        counts = {}
        for r in self._server_rows_with_status():
            k = r[0].split(" ")[0]
            counts[k] = counts.get(k, 0) + 1
        self.server_status_lbl.setText(
            f"🖥️ Szerveren: {len(self._server_items)} sor (betöltve {getattr(self, '_server_loaded_at', '')})  |  "
            f"✅ egyezik: {counts.get('✅', 0)}  |  ⚠️ eltér: {counts.get('⚠️', 0)}  |  ⏳ feltöltésre vár: {counts.get('⏳', 0)}  |  "
            f"➕ csak szerveren: {counts.get('➕', 0)}  |  ❌ csak helyben: {counts.get('❌', 0)}  |  "
            f"Megjelenítve: {len(rows)}  —  Dupla kattintás = szerkesztés")

        self._server_view_rows = rows
        t = self.server_table
        t.blockSignals(True)
        t.setUpdatesEnabled(False)
        t.setRowCount(len(rows))
        for ri, (st, col, data, ver, loc) in enumerate(rows):
            it = QTableWidgetItem(st)
            it.setForeground(QColor(col))
            it.setFlags(it.flags() & ~Qt.ItemFlag.ItemIsEditable)
            if loc and st.startswith("⚠️"):
                it.setToolTip("Helyi (ezen a gépen):\n" + format_row_for_log(loc) + "\n\nSzerveren:\n" + format_row_for_log(data))
            t.setItem(ri, 0, it)
            for ci, c in enumerate(ROW_COLS):
                cell = QTableWidgetItem(data.get(c, ""))
                if c in ("A", "B") or ver < 0:
                    cell.setFlags(cell.flags() & ~Qt.ItemFlag.ItemIsEditable)
                if loc and str(loc.get(c, "")).strip() != str(data.get(c, "")).strip() and ver >= 0:
                    cell.setBackground(QColor("#5c4a00"))
                    cell.setToolTip(f"Helyi érték: {loc.get(c, '') or '∅'}")
                t.setItem(ri, ci + 1, cell)
            vi = QTableWidgetItem(str(ver) if ver >= 0 else "-")
            vi.setFlags(vi.flags() & ~Qt.ItemFlag.ItemIsEditable)
            t.setItem(ri, len(ROW_COLS) + 1, vi)
        t.setUpdatesEnabled(True)
        t.blockSignals(False)

    def _server_row_data(self, ri):
        return dict(self._server_view_rows[ri][2]) if 0 <= ri < len(getattr(self, "_server_view_rows", [])) else None

    def on_server_cell_edited(self, item):
        """Cella javítása a Szerver adatbázis fülön: a teljes sor a megszokott úton megy ki
        (helyi adatbázis + VPS + Google Sheets), így minden hely egyforma marad."""
        ri, ci = item.row(), item.column()
        if ci < 1 or ci > len(ROW_COLS):
            return
        data = self._server_row_data(ri)
        if not data:
            return
        col = ROW_COLS[ci - 1]
        new_val = item.text().strip()
        old_val = data.get(col, "")
        if new_val == old_val:
            return
        code = data["A"].strip().upper()
        elotte = get_row_snapshot(code)
        data[col] = new_val
        try:
            conn = get_db_connection()
            cursor = conn.cursor()
            cursor.execute(f"UPDATE termekek SET {', '.join(c + '=?' for c in ROW_COLS)} WHERE UPPER(TRIM(A)) = ? OR UPPER(TRIM(B)) = ?",
                           (*[data[c] for c in ROW_COLS], code, code))
            if cursor.rowcount == 0:
                cursor.execute(f"INSERT INTO termekek ({', '.join(ROW_COLS)}, szin) VALUES ({', '.join('?' * len(ROW_COLS))}, 'FEHÉR')",
                               [data[c] for c in ROW_COLS])
            conn.commit()
            conn.close()
        except Exception as e:
            self.trigger_toast(f"❌ Mentési hiba: {e}")
            return
        payload = dict(data)
        payload["_force"] = True
        enqueue_outbox(payload)
        flush_outbox_async()
        log_event(code, "MÓDOSÍTÁS (Szerver adatbázis fül)", f"{col} ({COLUMN_LABELS.get(col, col)}): {old_val or '∅'} → {new_val or '∅'}",
                  elotte, get_row_snapshot(code), f"🖥️ Szerver adatbázis fül - cella szerkesztés ({col} oszlop)")
        self._server_view_rows[ri] = ("⏳ Helyi módosítás feltöltésre vár", "#ffb74d", data, self._server_view_rows[ri][3], data)
        st = self.server_table.item(ri, 0)
        if st:
            self.server_table.blockSignals(True)
            st.setText("⏳ Mentés folyamatban (szerver + Google Sheets)...")
            st.setForeground(QColor("#ffb74d"))
            self.server_table.blockSignals(False)
        self.trigger_toast(f"💾 {code} [{col}] mentve - szerver + Google Sheets")
        if hasattr(self, "silent_update_excel_table"):
            self.silent_update_excel_table()

    def _selected_server_rows(self):
        rows = sorted({i.row() for i in self.server_table.selectedIndexes()})
        return [(r, self._server_row_data(r)) for r in rows if self._server_row_data(r)]

    def delete_selected_server_items(self):
        sel = self._selected_server_rows()
        if not sel:
            self.trigger_toast("⚠️ Jelölj ki legalább egy sort!")
            return
        lista = "\n".join(f"• {d['A']}  |  {d.get('C') or '-'}  |  {d.get('D') or '-'}  |  {d.get('E') or '-'}" for _, d in sel[:12])
        if len(sel) > 12:
            lista += f"\n... és még {len(sel) - 12} sor"
        dlg = SilentConfirmDialog("🗑️ Törlés a szerverről",
                                  f"Biztosan TÖRLÖD ezt a {len(sel)} terméket?\n\nTörlődik a VPS szerverről, a Google táblázatból "
                                  f"és erről a gépről is (a másik gépről is). A teljes adatuk a Naplóba és a ToroltSor lapra kerül.\n\n{lista}",
                                  self)
        if dlg.exec() != QDialog.DialogCode.Accepted:
            return
        for _, d in sel:
            code = d["A"].strip().upper()
            if not get_row_snapshot(code):
                # Helyben nincs meg: a naplóhoz / archívumhoz a szerver adatát használjuk
                archive_deleted_row(d, "Szerver adatbázis fül: törlés")
                log_event(code, "TÖRLÉS", "Törölve a szerverről (helyben nem volt meg)", d, None,
                          "🖥️ Szerver adatbázis fül - 🗑️ Kijelölt törlése gomb")
                enqueue_delete(code)
            else:
                DeleteWorker(code, "Törölve a Szerver adatbázis fülön",
                             "🖥️ Szerver adatbázis fül - 🗑️ Kijelölt törlése gomb").run()
        flush_outbox_async()
        codes = {d["A"].strip().upper() for _, d in sel}
        self._server_items = [it for it in self._server_items if str(it.get("A", "")).strip().upper() not in codes]
        self.render_server_table()
        self.trigger_toast(f"🗑️ {len(sel)} termék törölve (szerver + Google Sheets + helyi)")
        if hasattr(self, "silent_update_excel_table"):
            self.silent_update_excel_table()

    def upload_selected_local_to_server(self):
        sel = self._selected_server_rows()
        done = 0
        for _, d in sel:
            code = d["A"].strip().upper()
            if enqueue_row_from_db(code, force=True):
                done += 1
                log_event(code, "MÓDOSÍTÁS (helyi adat feltöltve a szerverre)", "", None, get_row_snapshot(code),
                          "🖥️ Szerver adatbázis fül - ⬆️ Helyi adat feltöltése gomb")
        if done:
            flush_outbox_async()
            self.trigger_toast(f"⬆️ {done} sor helyi adata feltöltésre küldve (szerver + Google Sheets)")
        else:
            self.trigger_toast("⚠️ A kijelölt sor(ok) nincsenek meg ezen a gépen.")

    def execute_pda_search(self):
        # A PDA mező a vonalkód-olvasóé: itt mindig javítjuk az ü → -, ö → 0 elgépelést
        query_text = _scanner_translate(self.pda_input.text().strip())
        if not query_text: return
        if query_text != self.pda_input.text().strip():
            self.pda_input.setText(query_text)
        clean_query = clean_part_number(query_text)

        while self.pda_results_layout.count():
            item = self.pda_results_layout.takeAt(0)
            widget = item.widget()
            if widget:
                widget.deleteLater()

        conn = get_db_connection()
        df = pd.read_sql_query("SELECT * FROM termekek WHERE C IS NOT NULL AND C != ''", conn)
        conn.close()

        matched_rows = []
        # Leolvasott R18-kód: a pontosan egyező termék kerül előre (R18-129 ne hozza elöl az R18-1290-et)
        q_up = query_text.strip().upper()
        exact = df[(df["A"].astype(str).str.strip().str.upper() == q_up) | (df["B"].astype(str).str.strip().str.upper() == q_up)]
        for _, row in exact.iterrows():
            matched_rows.append(dict(row))
        exact_idx = set(exact.index)
        if not (re.match(r"^R18-\d+$", q_up) and matched_rows):
            for idx, row in df.iterrows():
                if idx in exact_idx:
                    continue
                row_text = " ".join([safe_str(row.get(col)) for col in df.columns])
                if query_text.lower() in row_text.lower() or (clean_query and clean_query in clean_part_number(row_text)):
                    matched_rows.append(dict(row))
                    if len(matched_rows) >= 10: break

        if not matched_rows:
            self.pda_results_layout.addWidget(QLabel("❌ Nincs találat!"))
        else:
            for row_dict in matched_rows:
                self.pda_results_layout.addWidget(self.create_item_card(row_dict, sale_mode=True))

        # A következő csippantás felülírja (nem fűzi hozzá) a keresőmező tartalmát
        self.pda_input.selectAll()

    # -------------------------------------------------------------------
    # TAB 4: EXCEL KEZELÉS TAB
    # -------------------------------------------------------------------
    def init_excel_tab(self):
        layout = QVBoxLayout(self.tab_excel)
        layout.setSpacing(10)
        layout.setContentsMargins(10, 10, 10, 10)

        # 1. FELSŐ GOMBSOR
        top_btn_layout = QHBoxLayout()

        btn_export = QPushButton("📊 Excel Exportálás a Helyi Készletről")
        btn_export.setStyleSheet("background-color: #0288d1; color: white; font-weight: bold; padding: 8px;")
        btn_export.clicked.connect(
            lambda: self.run_button_action(
                btn_export, 
                "📊 Excel Exportálás a Helyi Készletről", 
                self.export_excel
            )
        )
        top_btn_layout.addWidget(btn_export)

        btn_open_history = QPushButton("📜 Google Sheets Verzióelőzmények (Múltbéli állapotok)")
        btn_open_history.setStyleSheet("background-color: #aa00ff; color: white; font-weight: bold; padding: 8px;")
        btn_open_history.clicked.connect(self.open_sheets_version_history)
        top_btn_layout.addWidget(btn_open_history)

        layout.addLayout(top_btn_layout)

        # 2. BIZTONSÁGI MENTÉSEK DOBOZ
        backup_group = QGroupBox("🛡️ Adatbázis Biztonsági Mentés & Visszaállítás (Pillanatképek)")
        backup_lay = QHBoxLayout(backup_group)

        btn_manual_backup = QPushButton("🛡️ Biztonsági Mentés Készítése")
        btn_manual_backup.setStyleSheet("background-color: #00e676; color: black; font-weight: bold; padding: 6px;")
        btn_manual_backup.clicked.connect(
            lambda: self.run_button_action(
                btn_manual_backup, 
                "🛡️ Biztonsági Mentés Készítése", 
                self.make_manual_backup
            )
        )
        backup_lay.addWidget(btn_manual_backup)

        btn_restore_backup = QPushButton("↩️ Visszaállítás Korábbi Állapotra")
        btn_restore_backup.setStyleSheet("background-color: #ff9800; color: black; font-weight: bold; padding: 6px;")
        btn_restore_backup.clicked.connect(
            lambda: self.run_button_action(
                btn_restore_backup, 
                "↩️ Visszaállítás Korábbi Állapotra", 
                self.open_restore_dialog
            )
        )
        backup_lay.addWidget(btn_restore_backup)

        layout.addWidget(backup_group)

        # 3. VÁLTOZÁSKÖVETŐ ÉS MÓDOSÍTÁSI ELŐZMÉNYEK TÁBLÁZAT (Mint az Excel / Sheets verziótörténet)
        history_group = QGroupBox("🔍 Excel Változáskövető & Cella Módosítási Előzmények")
        history_lay = QVBoxLayout(history_group)

        hist_top_bar = QHBoxLayout()
        lbl_hist_title = QLabel("🕒 A legutóbbi adat- és cellamódosítások részletes listája:")
        lbl_hist_title.setStyleSheet("color: #61afef; font-size: 13px; font-weight: bold;")
        hist_top_bar.addWidget(lbl_hist_title)
        hist_top_bar.addStretch()

        btn_refresh_changes = QPushButton("🔄 Előzmények Frissítése")
        btn_refresh_changes.setStyleSheet("background-color: #2b2b2b; color: white; font-weight: bold; padding: 4px 10px;")
        btn_refresh_changes.clicked.connect(self.refresh_excel_history_table)
        hist_top_bar.addWidget(btn_refresh_changes)
        history_lay.addLayout(hist_top_bar)

        self.table_history = QTableWidget()
        self.table_history.setColumnCount(5)
        self.table_history.setHorizontalHeaderLabels(["Időpont", "R18 Kód", "Esemény", "Részletek / Módosított értékek", "Állapot"])
        
        self.table_history.setStyleSheet("""
            QTableWidget {
                background-color: #1e2227;
                color: #abb2bf;
                gridline-color: #2c313a;
                font-size: 13px;
                border: 1px solid #181a1f;
            }
            QHeaderView::section {
                background-color: #21252b;
                color: #61afef;
                font-weight: bold;
                padding: 6px;
                border: 1px solid #2c313a;
            }
        """)

        # Oszlopszélességek beállítása
        self.table_history.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeMode.ResizeToContents)
        self.table_history.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeMode.ResizeToContents)
        self.table_history.horizontalHeader().setSectionResizeMode(2, QHeaderView.ResizeMode.ResizeToContents)
        self.table_history.horizontalHeader().setSectionResizeMode(3, QHeaderView.ResizeMode.Stretch)
        self.table_history.horizontalHeader().setSectionResizeMode(4, QHeaderView.ResizeMode.ResizeToContents)

        history_lay.addWidget(self.table_history)
        layout.addWidget(history_group)

        # Előzmények feltöltése indításkor
        self.refresh_excel_history_table()
        
    def refresh_excel_history_table(self):
        """Betölti és színezve megjeleníti az adatbázisban történt legutóbbi cellamódosításokat"""
        if not hasattr(self, 'table_history'):
            return

        try:
            conn = get_db_connection()
            query = "SELECT * FROM naplo ORDER BY id DESC LIMIT 150"
            df_hist = pd.read_sql_query(query, conn)
            conn.close()

            self.table_history.blockSignals(True)
            self.table_history.setRowCount(len(df_hist))

            for r_idx, row in df_hist.iterrows():
                idopont = safe_str(row.get('idopont'))
                kod = safe_str(row.get('r18_kod'))
                esemeny = safe_str(row.get('esemeny')).upper()
                reszletek = safe_str(row.get('reszletek'))

                item_time = QTableWidgetItem(idopont)
                item_kod = QTableWidgetItem(kod)
                item_esemeny = QTableWidgetItem(esemeny)
                item_reszlet = QTableWidgetItem(reszletek)
                item_status = QTableWidgetItem("✔ Rögzítve")

                # Színezés
                if "TÖRLÉS" in esemeny or "ELADÁS" in esemeny:
                    color_bg = QColor(100, 30, 30)
                elif "VISSZAVONÁS" in esemeny or "KIÜRÍTÉS" in esemeny:
                    color_bg = QColor(110, 70, 20)
                elif "CELLA MÓDOSÍTÁSA" in esemeny or "MÓDOSÍTÁS" in esemeny:
                    color_bg = QColor(25, 70, 100)
                elif "FELVÉTEL" in esemeny:
                    color_bg = QColor(20, 80, 40)
                else:
                    color_bg = QColor(35, 40, 48)

                for item in [item_time, item_kod, item_esemeny, item_reszlet, item_status]:
                    item.setBackground(color_bg)
                    item.setForeground(QColor("#ffffff"))

                self.table_history.setItem(r_idx, 0, item_time)
                self.table_history.setItem(r_idx, 1, item_kod)
                self.table_history.setItem(r_idx, 2, item_esemeny)
                self.table_history.setItem(r_idx, 3, item_reszlet)
                self.table_history.setItem(r_idx, 4, item_status)

            self.table_history.blockSignals(False)

        except Exception as e:
            print(f"Hiba a változáskövető táblázat frissítésekor: {e}")

    def make_manual_backup(self):
        path = create_database_backup(tag="manual")
        if path:
            self.trigger_toast("🛡️ Biztonsági pillanatkép elkészült!")
            log_event("SYSTEM", "MANUÁLIS BIZTONSÁGI MENTÉS", f"Fájl: {path}")

    def open_restore_dialog(self):
        dlg = RestoreDialog(self)
        if dlg.exec() == QDialog.DialogCode.Accepted:
            self.trigger_toast("↩️ Adatbázis sikeresen visszaállítva!")
            self.load_latest_200_items()
            self.refresh_counters()

    def export_excel(self):
        try:
            conn = get_db_connection()
            df = pd.read_sql_query("SELECT * FROM termekek", conn)
            conn.close()
            path, _ = QFileDialog.getSaveFileName(self, "Excel Mentése", "Raktar_Adatbazis.xlsx", "Excel Files (*.xlsx)")
            if path:
                df.to_excel(path, index=False)
                self.trigger_toast("✅ Excel sikeresen kimentve!")
        except Exception as e:
            self.trigger_toast(f"❌ Hiba: {e}")

    # -------------------------------------------------------------------
    # TAB 5: EGYEDI CÍMKE TAB
    # -------------------------------------------------------------------
    def init_single_tab(self):
        layout = QFormLayout(self.tab_single)
        self.single_code = QLineEdit("R18-")
        self.single_sub = QLineEdit()
        layout.addRow("R18 Kód:", self.single_code)
        layout.addRow("Cikkszám / Megnevezés:", self.single_sub)

        btn_print = QPushButton("🖨️ Egyedi Címke Nyomtatása")
        btn_print.clicked.connect(
            lambda: self.run_button_action(
                btn_print, 
                "🖨️ Egyedi Címke Nyomtatása", 
                self.print_single_label
            )
        )
        layout.addRow(btn_print)

    def print_single_label(self):
        code = self.single_code.text().strip().upper()
        sub = self.single_sub.text().strip().upper()
        if not code: return
        zpl = generate_zpl_code(code, cikkszam=sub)
        printers = [p[2] for p in win32print.EnumPrinters(win32print.PRINTER_ENUM_LOCAL | win32print.PRINTER_ENUM_CONNECTIONS)] if WIN32_AVAILABLE else []
        if printers:
            send_zpl_to_printer(printers[0], zpl)
            self.trigger_toast(f"🖨️ Egyedi címke kiküldve [{code}]!")

    # -------------------------------------------------------------------
    # TAB 6: RAKTÁRTÉRKÉP TAB
    # -------------------------------------------------------------------
    # ==========================================================
    # 📋 LELTÁR – Régi (raklaponként mi és mennyi) + Új leltár (gépen kézzel / csippantva, PDA-ról is)
    # ==========================================================
    LELTAR_TIP = {"egyezik": ("✅", "egyezik", "#00e676"), "elter": ("❌", "eltér", "#ff5252"),
                  "hiany": ("❌", "hiányzik", "#ff5252"), "tobb": ("⚠️", "több raklapon", "#ffb74d"),
                  "idegen": ("📦", "máshonnan ide", "#ffb74d"), "ismeretlen": ("❓", "nincs a rendszerben", "#ff5252")}

    def init_leltar_tab(self):
        layout = QVBoxLayout(self.tab_leltar)
        self.lel_tabs = QTabWidget()
        layout.addWidget(self.lel_tabs)

        # ---- RÉGI: raklaponként ----
        w1 = QWidget()
        l1 = QVBoxLayout(w1)
        bar = QHBoxLayout()
        self.lel_old_filter = QLineEdit()
        self.lel_old_filter.setPlaceholderText("🔍 Raklap szűrés (pl. 573, KISRAKTÁR) vagy R18-szám")
        attach_scanner_fix(self.lel_old_filter)
        self.lel_old_filter.textChanged.connect(self.leltar_fill_pallets)
        bar.addWidget(self.lel_old_filter, stretch=2)
        b = QPushButton("🔄 Frissítés")
        b.clicked.connect(self.leltar_refresh_old)
        bar.addWidget(b)
        self.lel_old_info = QLabel("")
        self.lel_old_info.setStyleSheet("color: #29b6f6; font-weight: bold;")
        bar.addWidget(self.lel_old_info, stretch=2)
        l1.addLayout(bar)
        split = QSplitter(Qt.Orientation.Horizontal)
        self.lel_pallets = self._admin_table(["Raklap / hely", "Termékfajta", "Összes db"], [150, 100, 90])
        self.lel_pallets.itemSelectionChanged.connect(self.leltar_show_pallet)
        split.addWidget(self.lel_pallets)
        right = QWidget()
        rl = QVBoxLayout(right)
        rl.setContentsMargins(0, 0, 0, 0)
        top = QHBoxLayout()
        self.lel_old_title = QLabel("Válassz egy raklapot bal oldalt.")
        self.lel_old_title.setStyleSheet("font-size: 18px; font-weight: bold; color: #ffb74d;")
        top.addWidget(self.lel_old_title, stretch=1)
        b2 = QPushButton("📋 Új leltár ezen a raklapon")
        b2.setStyleSheet("background-color: #29b6f6; color: black; font-weight: bold;")
        b2.clicked.connect(self.leltar_start_from_old)
        top.addWidget(b2)
        rl.addLayout(top)
        self.lel_old_items = self._admin_table(["R18", "Márka", "Típus", "Cikkszám", "DB", "Állapot", "Raklap (L)", "Páter (I)", "Oszlop (K)"],
                                               [100, 120, 160, 120, 50, 110, 110, 70])
        rl.addWidget(self.lel_old_items)
        split.addWidget(right)
        split.setSizes([360, 900])
        l1.addWidget(split, stretch=1)
        self.lel_tabs.addTab(w1, "📦 Régi – raklaponként")

        # ---- ÚJ LELTÁR ----
        w2 = QWidget()
        l2 = QHBoxLayout(w2)
        split2 = QSplitter(Qt.Orientation.Horizontal)
        left = QGroupBox("✍️ Új leltár a gépen (kézzel vagy vonalkód-olvasóval)")
        ll = QVBoxLayout(left)
        r1 = QHBoxLayout()
        self.lel_pallet_in = QLineEdit()
        self.lel_pallet_in.setPlaceholderText("Raklap (pl. 573) vagy páter / polc (pl. Páter 1 / Polc 12) – vagy csippants egy terméket")
        attach_scanner_fix(self.lel_pallet_in)
        self.lel_pallet_in.returnPressed.connect(self.leltar_start)
        r1.addWidget(self.lel_pallet_in, stretch=1)
        b = QPushButton("▶️ Leltár indítása")
        b.setStyleSheet("background-color: #00e676; color: black; font-weight: bold;")
        b.clicked.connect(self.leltar_start)
        r1.addWidget(b)
        ll.addLayout(r1)
        self.lel_run_info = QLabel("Írd be a raklap számát és indítsd a leltárt. Utána csippantsd a termékeket (minden darabot), "
                                   "vagy írd be kézzel a darabszámot a „Számolt” oszlopba.")
        self.lel_run_info.setWordWrap(True)
        self.lel_run_info.setStyleSheet("color: #29b6f6; font-weight: bold; padding: 4px;")
        ll.addWidget(self.lel_run_info)
        r2 = QHBoxLayout()
        self.lel_scan_in = QLineEdit()
        self.lel_scan_in.setPlaceholderText("R18 csippantás / beírás + Enter")
        attach_scanner_fix(self.lel_scan_in)
        self.lel_scan_in.returnPressed.connect(self.leltar_scan)
        r2.addWidget(self.lel_scan_in, stretch=1)
        r2.addWidget(QLabel("db:"))
        self.lel_qty = QSpinBox()
        self.lel_qty.setRange(1, 999)
        r2.addWidget(self.lel_qty)
        b = QPushButton("➕ Hozzáad")
        b.clicked.connect(self.leltar_scan)
        r2.addWidget(b)
        ll.addLayout(r2)
        self.lel_run_table = self._admin_table(["R18", "Termék", "Előtte", "Számolt", "Megjegyzés"], [100, 190, 70, 90])
        ll.addWidget(self.lel_run_table)
        r3 = QHBoxLayout()
        b = QPushButton("↩️ Utolsó csippantás visszavonása")
        b.clicked.connect(self.leltar_undo_scan)
        r3.addWidget(b)
        b = QPushButton("✖ Leltár elvetése")
        b.setStyleSheet("background-color: #555; color: white;")
        b.clicked.connect(self.leltar_discard)
        r3.addWidget(b)
        ll.addLayout(r3)
        self.lel_btn_done = QPushButton("✅ Kész – összegzés (előtte / most)")
        self.lel_btn_done.setStyleSheet("background-color: #00e676; color: black; font-weight: bold; font-size: 15px; padding: 8px;")
        self.lel_btn_done.clicked.connect(self.leltar_finish)
        ll.addWidget(self.lel_btn_done)
        split2.addWidget(left)

        right2 = QWidget()
        rl2 = QVBoxLayout(right2)
        rl2.setContentsMargins(0, 0, 0, 0)
        top2 = QHBoxLayout()
        t = QLabel("📋 Leltárak (PDA + gép)")
        t.setStyleSheet("font-size: 16px; font-weight: bold; color: #29b6f6;")
        top2.addWidget(t)
        self.lel_list_filter = QLineEdit()
        self.lel_list_filter.setPlaceholderText("Raklap szűrés")
        self.lel_list_filter.returnPressed.connect(self.leltar_load_list)
        top2.addWidget(self.lel_list_filter)
        b = QPushButton("🔄 Frissítés")
        b.clicked.connect(self.leltar_load_list)
        top2.addWidget(b)
        rl2.addLayout(top2)
        self.lel_list_info = QLabel("")
        self.lel_list_info.setStyleSheet("color: #888;")
        rl2.addWidget(self.lel_list_info)
        self.lel_scroll = QScrollArea()
        self.lel_scroll.setWidgetResizable(True)
        self.lel_cards = QWidget()
        self.lel_cards_layout = QGridLayout(self.lel_cards)
        self.lel_cards_layout.setAlignment(Qt.AlignmentFlag.AlignTop)
        self.lel_scroll.setWidget(self.lel_cards)
        rl2.addWidget(self.lel_scroll)
        split2.addWidget(right2)
        split2.setSizes([560, 800])
        l2.addWidget(split2)
        self.lel_tabs.addTab(w2, "📋 Új leltár (PDA + gép)")

        self._lel = None          # a futó gépi leltár: {raklap, kezdes, rows: [{A, nev, elotte, L, megj}], cnt: {A: n}, hist: [A]}
        self._lel_list = {}
        self._lel_old_rows = []
        self.lel_timer = QTimer(self)
        self.lel_timer.timeout.connect(lambda: self._tab_is_current("tab_leltar") and self.leltar_load_list())
        self.lel_timer.start(30000)

    # ---- Régi: raklaponként a helyi adatbázisból ----
    def leltar_refresh_old(self):
        try:
            conn = get_db_connection()
            cur = conn.cursor()
            cur.execute("SELECT A, C, D, E, F, G, I, K, L FROM termekek WHERE C IS NOT NULL AND TRIM(C) != ''")
            self._lel_old_rows = [{k: safe_str(r[k]) for k in r.keys()} for r in cur.fetchall()]
            conn.close()
        except Exception as e:
            self.lel_old_info.setText(f"❌ Adatbázis hiba: {e}")
            return
        self.leltar_fill_pallets()

    def _lel_pallet_map(self):
        m = {}
        for r in self._lel_old_rows:
            keys = pallet_tokens(r["L"])
            if r.get("I"):
                keys.append(norm_hely(f"PATER {r['I']}" + (f" / POLC {r['K']}" if r.get("K") else "")))
            for t in keys or ["(nincs raklap)"]:
                m.setdefault(t, []).append(r)
        return m

    def leltar_fill_pallets(self):
        m = self._lel_pallet_map()
        f = norm_pallet(self.lel_old_filter.text())
        fh = norm_hely(self.lel_old_filter.text())
        if fh.startswith("PATER "):
            f = fh
        if f and re.match(r"^R18", f):
            keys = sorted({t for t, rows in m.items() if any(norm_code(r["A"]) == norm_code(f) for r in rows)})
        else:
            keys = [k for k in m if not f or f in k]
        keys.sort(key=lambda k: (0, int(k), "") if k.isdigit() else (1, 0, k))
        t = self.lel_pallets
        t.setRowCount(len(keys))
        for i, k in enumerate(keys):
            rows = m[k]
            for c, v in enumerate([hely_txt(k) if k.startswith("PATER ") else k, len(rows), sum(stock_of(r["F"]) for r in rows)]):
                it = QTableWidgetItem(str(v))
                it.setData(Qt.ItemDataRole.UserRole, k)
                t.setItem(i, c, it)
        self.lel_old_info.setText(f"{len(keys)} raklap / hely · {len(self._lel_old_rows)} termék")
        if len(keys) == 1:
            t.selectRow(0)

    def leltar_show_pallet(self):
        it = self.lel_pallets.item(self.lel_pallets.currentRow(), 0) if self.lel_pallets.currentRow() >= 0 else None
        if not it:
            return
        k = it.data(Qt.ItemDataRole.UserRole)
        rows = sorted(self._lel_pallet_map().get(k, []), key=lambda r: r["A"])
        self._lel_old_sel = k
        self.lel_old_title.setText(f"{'🗄️' if k.startswith('PATER ') else '📦'} {hely_txt(k)}: {len(rows)} féle termék, összesen {sum(stock_of(r['F']) for r in rows)} db")
        t = self.lel_old_items
        t.setRowCount(len(rows))
        for i, r in enumerate(rows):
            vals = [r["A"], r["C"], r["D"], r["E"], r["F"] or "1", r["G"], r["L"], r["I"], r["K"]]
            for c, v in enumerate(vals):
                cell = QTableWidgetItem(str(v))
                if c == 6 and len(pallet_tokens(r["L"])) > 1:
                    cell.setForeground(QColor("#ffb74d"))
                    cell.setToolTip("Több raklapon van – a DB az összes raklapra együtt vonatkozik")
                t.setItem(i, c, cell)

    def leltar_start_from_old(self):
        k = getattr(self, "_lel_old_sel", "")
        if not k or k == "(nincs raklap)":
            self.trigger_toast("Előbb válassz egy raklapot.")
            return
        self.lel_tabs.setCurrentIndex(1)
        self.lel_pallet_in.setText(k)
        self.leltar_start()

    # ---- Új leltár a gépen ----
    def _lel_find(self, code):
        code = norm_code(code)
        for r in self._lel_old_rows:
            if norm_code(r["A"]) == code:
                return r
        return None

    def leltar_start(self):
        raklap = norm_hely(fix_scanner_text(self.lel_pallet_in.text()))
        if not raklap:
            return
        if raklap.startswith("R18"):
            # termék kódja került a hely mezőbe: a termék helyén indul
            self.lel_pallet_in.clear()
            self.lel_scan_in.setText(raklap)
            self.leltar_scan()
            return
        if self._lel and self._lel["raklap"] != raklap or (self._lel and self._lel["cnt"] and self._lel["raklap"] == raklap):
            if QMessageBox.question(self, "Új leltár?", f"Van egy folyamatban lévő leltár: {self._lel['raklap']} "
                                    f"({sum(self._lel['cnt'].values())} db számolva).\n\nElveted, és új leltárt kezdesz a(z) {hely_txt(raklap)} helyre?"
                                    ) != QMessageBox.StandardButton.Yes:
                return
        self.leltar_refresh_old()
        rows = [r for r in self._lel_old_rows if on_hely(r, raklap)]
        rows.sort(key=lambda r: r["A"])
        self._lel = {"raklap": raklap, "kezdes": datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"), "cnt": {}, "hist": [],
                     "rows": [{"A": r["A"], "nev": " ".join(x for x in (r["C"], r["D"]) if x), "elotte": stock_of(r["F"]), "L": r["L"],
                               "megj": ("több raklapon: " + r["L"]) if len(pallet_tokens(r["L"])) > 1 and not raklap.startswith("PATER ") else ""} for r in rows]}
        self.lel_pallet_in.setText(raklap)
        self.leltar_render_run()
        self.lel_scan_in.setFocus()

    def leltar_render_run(self):
        L = self._lel
        t = self.lel_run_table
        t.blockSignals(True)
        t.setRowCount(0)
        if not L:
            self.lel_run_info.setText("Írd be a raklap számát és indítsd a leltárt.")
            t.blockSignals(False)
            return
        t.setRowCount(len(L["rows"]))
        for i, r in enumerate(L["rows"]):
            n = L["cnt"].get(norm_code(r["A"]), 0)
            vals = [r["A"], r["nev"], "–" if r["elotte"] is None else str(r["elotte"]), None, r["megj"]]
            for c, v in enumerate(vals):
                if c == 3:
                    sp = QSpinBox()
                    sp.setRange(0, 9999)
                    sp.setValue(n)
                    sp.valueChanged.connect(lambda val, code=norm_code(r["A"]): self._leltar_set_cnt(code, val))
                    t.setCellWidget(i, 3, sp)
                    continue
                it = QTableWidgetItem(v)
                if r["elotte"] is not None and c == 2 and n != r["elotte"] and (n or L["cnt"]):
                    it.setForeground(QColor("#ff5252" if n != r["elotte"] else "#00e676"))
                if c == 4 and v:
                    it.setForeground(QColor("#ffb74d"))
                t.setItem(i, c, it)
        found = sum(1 for r in L["rows"] if L["cnt"].get(norm_code(r["A"]), 0) > 0)
        exp = [r for r in L["rows"] if r["elotte"] is not None and not r["megj"].startswith(("más", "nincs"))]
        self.lel_run_info.setText(
            f"{'🗄️' if L['raklap'].startswith('PATER ') else '📦'} {hely_txt(L['raklap'])} – a táblázat szerint {len(exp)} féle termék, összesen {sum(r['elotte'] for r in exp)} db.  "
            f"Számolva: {found} féle, {sum(L['cnt'].values())} db.")
        t.blockSignals(False)

    def _leltar_set_cnt(self, code, val):
        if not self._lel:
            return
        if val:
            self._lel["cnt"][code] = val
        else:
            self._lel["cnt"].pop(code, None)
        QTimer.singleShot(0, self._leltar_update_info)

    def _leltar_update_info(self):
        L = self._lel
        if not L:
            return
        found = sum(1 for r in L["rows"] if L["cnt"].get(norm_code(r["A"]), 0) > 0)
        exp = [r for r in L["rows"] if r["elotte"] is not None and not r["megj"].startswith(("más", "nincs"))]
        self.lel_run_info.setText(
            f"{'🗄️' if L['raklap'].startswith('PATER ') else '📦'} {hely_txt(L['raklap'])} – a táblázat szerint {len(exp)} féle termék, összesen {sum(r['elotte'] for r in exp)} db.  "
            f"Számolva: {found} féle, {sum(L['cnt'].values())} db.")

    def leltar_scan(self):
        code = norm_code(fix_scanner_text(self.lel_scan_in.text()))
        self.lel_scan_in.clear()
        if not code:
            return
        if not self._lel and code.startswith("R18"):
            # termék csippantva leltár nélkül: a termék raklapján / páterén-polcán indul, és ez a darab már számít
            self.leltar_refresh_old()
            p = self._lel_find(code)
            if not p:
                self.trigger_toast(f"❓ {code} nincs a rendszerben – írd be a raklapot / pátert")
                return
            opts = pallet_tokens(p["L"])
            if p.get("I"):
                opts.append(norm_hely(f"PATER {p['I']}" + (f" / POLC {p['K']}" if p.get("K") else "")))
            if not opts:
                self.trigger_toast(f"⚠️ {code}-nak nincs se raklapja, se pátere – írd be a helyet")
                return
            hely = opts[0]
            if len(opts) > 1:
                txt, ok = QInputDialog.getItem(self, "Melyik hely leltára?", f"{code} – hely: {row_hely_txt(p)}\nMelyiknél állsz?",
                                               [hely_txt(o) for o in opts], 0, False)
                if not ok:
                    return
                hely = opts[[hely_txt(o) for o in opts].index(txt)]
            self.lel_pallet_in.setText(hely)
            self.leltar_start()
            if self._lel and self._lel["raklap"] == hely:
                self.lel_scan_in.setText(code)
                self.leltar_scan()
            return
        if not self._lel or not code.startswith("R18"):
            self.lel_pallet_in.setText(code)
            self.leltar_start()
            return
        L = self._lel
        if not any(norm_code(r["A"]) == code for r in L["rows"]):
            p = self._lel_find(code)
            if p:
                L["rows"].append({"A": p["A"], "nev": " ".join(x for x in (p["C"], p["D"]) if x), "elotte": 0, "L": p["L"],
                                  "megj": f"máshonnan (rendszer: {row_hely_txt(p)})"})
                self.trigger_toast(f"⚠️ {code} a rendszer szerint máshol van: {row_hely_txt(p)}")
            else:
                L["rows"].append({"A": code, "nev": "", "elotte": None, "L": "", "megj": "nincs a rendszerben"})
                self.trigger_toast(f"❓ {code} nincs a rendszerben")
        q = self.lel_qty.value()
        L["cnt"][code] = L["cnt"].get(code, 0) + q
        L["hist"].append((code, q))
        self.lel_qty.setValue(1)
        self.leltar_render_run()
        for i, r in enumerate(L["rows"]):
            if norm_code(r["A"]) == code:
                self.lel_run_table.selectRow(i)
                self.lel_run_table.scrollToItem(self.lel_run_table.item(i, 0))
        self.lel_scan_in.setFocus()

    def leltar_undo_scan(self):
        L = self._lel
        if not L or not L["hist"]:
            self.trigger_toast("Nincs visszavonható csippantás.")
            return
        code, q = L["hist"][-1]
        if QMessageBox.question(self, "Visszavonás?", f"Utolsó csippantás visszavonása: {code} ({q} db)?") != QMessageBox.StandardButton.Yes:
            return
        L["hist"].pop()
        n = L["cnt"].get(code, 0) - q
        if n > 0:
            L["cnt"][code] = n
        else:
            L["cnt"].pop(code, None)
        self.leltar_render_run()

    def leltar_discard(self):
        if not self._lel:
            return
        if QMessageBox.question(self, "Leltár elvetése?", f"{hely_txt(self._lel['raklap'])} – a számolás elvész, semmi nem kerül mentésre.\n\nElveted?"
                                ) != QMessageBox.StandardButton.Yes:
            return
        self._lel = None
        self.lel_pallet_in.clear()
        self.leltar_render_run()

    def _lel_job(self, kind, method, path, body=None, extra=None):
        def job():
            try:
                res = vps_request(method, f"{VPS_URL}{path}", json=body, params=None, timeout=15)
                try:
                    data = res.json()
                except Exception:
                    data = {"status": "error", "message": f"HTTP {res.status_code}"}
                if res.status_code != 200 and "message" not in data:
                    data["message"] = f"HTTP {res.status_code}"
                if res.status_code == 404 and not data.get("kod"):
                    data["message"] = "A VPS-en még a régi server_api.py fut (nincs leltár végpont) – töltsd fel az újat"
                data["_ok"] = res.status_code == 200
            except Exception as e:
                data = {"_ok": False, "message": f"Nem érem el a szervert ({e.__class__.__name__})"}
            data["_extra"] = extra
            self.leltar_signal.emit(kind, data)
        threading.Thread(target=job, daemon=True).start()

    def leltar_finish(self):
        L = self._lel
        if not L:
            self.trigger_toast("Nincs folyamatban leltár.")
            return
        self.lel_btn_done.setEnabled(False)
        self._lel_job("osszegzes", "POST", "/leltar/osszegzes", {"raklap": L["raklap"], "szamlalt": L["cnt"]})

    def leltar_load_list(self):
        f = norm_pallet(self.lel_list_filter.text()) if hasattr(self, "lel_list_filter") else ""
        self._lel_job("lista", "GET", "/leltarak?limit=40" + (f"&raklap={urllib.parse.quote(f)}" if f else ""))

    def on_leltar_signal(self, kind, data):
        if kind == "osszegzes":
            self.lel_btn_done.setEnabled(True)
        if not data.get("_ok"):
            msg = f"❌ Leltár: {data.get('message', 'hiba')}" + (f" [{data['kod']}]" if data.get("kod") else "")
            if kind == "lista":
                self.lel_list_info.setText(msg)
            else:
                QMessageBox.warning(self, "Leltár", msg)
            return
        if kind == "osszegzes":
            self.leltar_summary_dialog(data)
        elif kind == "lista":
            self.leltar_fill_cards(data.get("leltarak") or [])
        elif kind in ("mentes", "felulir", "visszavon"):
            hibak = data.get("hibak") or []
            if kind == "mentes":
                self._lel = None
                self.lel_pallet_in.clear()
                self.leltar_render_run()
                txt = f"✅ Leltár lezárva" + (f" – {data.get('modositas')} felülírás" if data.get("modositas") else "")
            elif kind == "felulir":
                txt = f"💾 {data.get('modositas', 0)} felülírás elküldve"
            else:
                txt = f"↩️ Leltár visszavonva" + (f" – {data.get('visszavont')} érték visszaáll" if data.get("visszavont") else "")
            if hibak:
                QMessageBox.warning(self, "Leltár", txt + "\n\nNem sikerült:\n" + "\n".join(hibak))
            else:
                self.trigger_toast(txt, 3500)
            log_event("-", "LELTÁR", f"{txt} ({hely_txt((data.get('leltar') or {}).get('raklap', ''))})", None, None, "Gép")
            # a felülírásokat / visszaállításokat ez a gép rögtön végre is hajtja
            QTimer.singleShot(1200, self.poll_pda_sales)
            QTimer.singleShot(5000, self.leltar_load_list)
            self.leltar_load_list()

    @staticmethod
    def _lel_muv(m):
        return ", ".join(("DB → " + v + " db") if k == "F" else ("Raklap → " + (v or "(üres)")) if k == "L"
                         else f"Páter → {v}" if k == "I" else f"Polc → {v}" if k == "K" else f"{k} → {v}"
                         for k, v in (m or {}).items())

    def leltar_summary_dialog(self, data):
        L = self._lel
        if not L:
            return
        rows = data.get("sorok") or []
        dlg = QDialog(self)
        dlg.setWindowTitle(f"📋 {hely_txt(L['raklap'])} – leltár összegzés")
        dlg.resize(980, 560)
        v = QVBoxLayout(dlg)
        ok_n = sum(1 for r in rows if r["tipus"] == "egyezik")
        head = QLabel(f"✅ Minden egyezik ({ok_n} termék) – nincs mit felülírni." if data.get("egyezik") else
                      f"⚠️ Eltérés: {len(rows) - ok_n} sor · egyezik: {ok_n}.  Pipáld ki, mit írjon felül (előtte → most).")
        head.setStyleSheet("font-size: 15px; font-weight: bold; color: %s;" % ("#00e676" if data.get("egyezik") else "#ffb74d"))
        v.addWidget(head)
        t = self._admin_table(["Felülír", "Állapot", "R18", "Termék", "Előtte", "Most", "Javasolt módosítás"], [70, 150, 100, 220, 60, 60])
        t.setRowCount(len(rows))
        checks = []
        for i, r in enumerate(rows):
            ic, txt, col = self.LELTAR_TIP.get(r["tipus"], ("•", r["tipus"], "#fff"))
            cb = None
            if r.get("muvelet"):
                cb = QCheckBox()
                cb.setChecked(bool(r.get("javasolt")))
                t.setCellWidget(i, 0, cb)
            checks.append((cb, r))
            vals = [None, f"{ic} {txt}", r["A"], r.get("nev", ""), "–" if r.get("elotte") is None else str(r["elotte"]), str(r["utana"]),
                    self._lel_muv(r.get("muvelet")) + (f"  (rendszer: {r.get('L') or '—'})" if r["tipus"] == "idegen" else "")]
            for c, val in enumerate(vals):
                if c == 0:
                    continue
                it = QTableWidgetItem(val)
                if c == 1:
                    it.setForeground(QColor(col))
                t.setItem(i, c, it)
        v.addWidget(t)
        has = any(cb for cb, _ in checks)
        hint = QLabel("A kijelölt sorokat ez a gép írja át mindenhol (Google táblázat is). A leltár a kártyáján később visszavonható."
                      if has else "")
        hint.setStyleSheet("color: #888;")
        v.addWidget(hint)
        bb = QHBoxLayout()
        res = {"v": None}
        if has:
            bw = QPushButton("💾 Felülírás (kijelöltek)")
            bw.setStyleSheet("background-color: #00e676; color: black; font-weight: bold; padding: 8px;")
            bw.clicked.connect(lambda: (res.update(v="ir"), dlg.accept()))
            bb.addWidget(bw)
        bc = QPushButton("✅ Lezárás felülírás nélkül" if has else "✅ Leltár lezárása")
        bc.setStyleSheet("font-weight: bold; padding: 8px;")
        bc.clicked.connect(lambda: (res.update(v="zar"), dlg.accept()))
        bb.addWidget(bc)
        bx = QPushButton("↩️ Vissza a leltárhoz")
        bx.clicked.connect(dlg.reject)
        bb.addWidget(bx)
        v.addLayout(bb)
        if not dlg.exec() or not res["v"]:
            return
        kod = [r["A"] for cb, r in checks if cb and cb.isChecked()] if res["v"] == "ir" else []
        if res["v"] == "ir":
            if not kod:
                QMessageBox.information(self, "Leltár", "Nincs kijelölve semmi – pipáld ki, mit írjon felül.")
                return
            lista = "\n".join(f"{r['A']}: {self._lel_muv(r['muvelet'])}" for cb, r in checks if cb and cb.isChecked())
            q = f"{hely_txt(L['raklap'])} – {len(kod)} termék felülírása:\n\n{lista}\n\nFelülírod? (Később visszavonható.)"
        else:
            q = f"{hely_txt(L['raklap'])} leltárának lezárása felülírás NÉLKÜL?\n(A különbség a leltár kártyáján látszik, később is felülírható.)"
        if QMessageBox.question(self, "Leltár", q) != QMessageBox.StandardButton.Yes:
            return
        self._lel_job("mentes", "POST", "/leltar", {"raklap": L["raklap"], "szamlalt": L["cnt"], "eszkoz": GEP_NEV,
                                                     "forras": "GEP", "kezdes": L["kezdes"], "felulir": kod})

    # ---- Leltár-kártyák ----
    def leltar_fill_cards(self, items):
        lay = self.lel_cards_layout
        while lay.count():
            w = lay.takeAt(0).widget()
            if w:
                w.deleteLater()
        self._lel_list = {x["id"]: x for x in items}
        cols = 2 if self.lel_scroll.width() > 1100 else 1
        for i, l in enumerate(items):
            lay.addWidget(self._leltar_card(l), i // cols, i % cols)
        self.lel_list_info.setText(f"{len(items)} leltár · frissítve {datetime.datetime.now().strftime('%H:%M:%S')}"
                                   if items else "Még nincs lezárt leltár.")

    def _leltar_card(self, l):
        vv = l.get("allapot") == "VISSZAVONVA"
        border = "#555" if vv else "#00c853" if l.get("egyezik") else "#ff9800"
        fr = QFrame()
        fr.setStyleSheet(f"QFrame#lc {{ border: 2px solid {border}; border-radius: 10px; background: #1e1e1e; }}")
        fr.setObjectName("lc")
        v = QVBoxLayout(fr)
        allapot = {"VISSZAVONVA": "↩️ VISSZAVONVA", "FELULIRVA": "💾 FELÜLÍRVA"}.get(l.get("allapot"), "LEZÁRVA")
        egy = "✅ Egyezik" if l.get("egyezik") else "⚠️ Eltérés"
        h = QLabel(f"<span style='font-size:20px;font-weight:bold'>{'🗄️ ' + hely_txt(l.get('raklap')) if str(l.get('raklap')).startswith('PATER ') else '📦 ' + str(l.get('raklap'))}</span>&nbsp;&nbsp;"
                   f"<span style='background:{'#2e7d32' if l.get('egyezik') else '#ff9800'};color:{'#fff' if l.get('egyezik') else '#000'};"
                   f"font-weight:bold'>&nbsp;{egy}&nbsp;</span>&nbsp;&nbsp;<span style='background:#37474f;font-weight:bold'>&nbsp;{allapot}&nbsp;</span>"
                   f"<br><span style='color:#999'>{str(l.get('lezaras') or '')[:16]} · {'🖥️' if l.get('forras') == 'GEP' else '📱'} {l.get('eszkoz', '')}"
                   + (f" · visszavonta: {l.get('visszavonta')}" if vv else "") + "</span>")
        h.setTextFormat(Qt.TextFormat.RichText)
        v.addWidget(h)
        rows_html = []
        for r in l.get("sorok") or []:
            ic, txt, col = self.LELTAR_TIP.get(r.get("tipus"), ("•", r.get("tipus"), "#fff"))
            st = ""
            if r.get("vid"):
                st = {"kesz": " <span style='color:#00e676'>✔ átírva</span>", "visszavonva": " <span style='color:#ffb74d'>↩️ visszavonva</span>",
                      "hiba": f" <span style='color:#ff5252'>❌ {r.get('allapot_uzenet', '')}</span>"}.get(r.get("allapot"), " <span style='color:#ffb74d'>⏳ gépre vár</span>")
            elif r.get("hiba"):
                st = f" <span style='color:#ff5252'>❌ {r['hiba']}</span>"
            muv = self._lel_muv(r.get("muvelet"))
            rows_html.append(
                f"<tr><td width='22' valign='top'>{ic}</td><td><b>{r.get('A')}</b> {r.get('nev', '')}<br><span style='color:{col}'>{txt}</span>"
                + (f"<span style='color:#999'> · {muv}</span>" if muv else "") + st +
                f"</td><td align='right' style='white-space:nowrap'><b>{'–' if r.get('elotte') is None else r.get('elotte')} → {r.get('utana')} db</b></td></tr>")
        body = QLabel("<table width='100%' cellspacing='2'>" + "".join(rows_html) + "</table>" if rows_html else "<i>Üres raklap</i>")
        body.setTextFormat(Qt.TextFormat.RichText)
        body.setWordWrap(True)
        v.addWidget(body)
        if not vv:
            bb = QHBoxLayout()
            if any(r.get("muvelet") and not r.get("vid") for r in l.get("sorok") or []):
                b = QPushButton("💾 Felülírás")
                b.setStyleSheet("background-color: #29b6f6; color: black; font-weight: bold;")
                b.clicked.connect(lambda _=False, lid=l["id"]: self.leltar_write_card(lid))
                bb.addWidget(b)
            b = QPushButton("↩️ Visszavonás")
            b.setStyleSheet("background-color: #ff9800; color: black; font-weight: bold;")
            b.clicked.connect(lambda _=False, lid=l["id"]: self.leltar_undo_card(lid))
            bb.addWidget(b)
            v.addLayout(bb)
        return fr

    def leltar_write_card(self, lid):
        l = self._lel_list.get(lid)
        if not l:
            return
        rows = [r for r in l.get("sorok") or [] if r.get("muvelet") and not r.get("vid")]
        dlg = QDialog(self)
        dlg.setWindowTitle(f"💾 Felülírás – {hely_txt(l['raklap'])}")
        v = QVBoxLayout(dlg)
        v.addWidget(QLabel("Mit írjon felül?"))
        cbs = []
        for r in rows:
            cb = QCheckBox(f"{r['A']} {r.get('nev', '')}  –  {r.get('elotte', '–')} → {r.get('utana')} db  ·  {self._lel_muv(r['muvelet'])}")
            cb.setChecked(bool(r.get("javasolt")))
            v.addWidget(cb)
            cbs.append((cb, r))
        bb = QHBoxLayout()
        b1 = QPushButton("💾 Felülírás")
        b1.clicked.connect(dlg.accept)
        b2 = QPushButton("Mégse")
        b2.clicked.connect(dlg.reject)
        bb.addWidget(b1)
        bb.addWidget(b2)
        v.addLayout(bb)
        if not dlg.exec():
            return
        kod = [r["A"] for cb, r in cbs if cb.isChecked()]
        if not kod:
            return
        if QMessageBox.question(self, "Felülírás?", f"{hely_txt(l['raklap'])} – {len(kod)} termék felülírása:\n\n" +
                                "\n".join(f"{r['A']}: {self._lel_muv(r['muvelet'])}" for cb, r in cbs if cb.isChecked()) +
                                "\n\nBiztosan felülírod?") != QMessageBox.StandardButton.Yes:
            return
        self._lel_job("felulir", "POST", f"/leltar/{lid}/felulir", {"felulir": kod, "eszkoz": GEP_NEV})

    def leltar_undo_card(self, lid):
        l = self._lel_list.get(lid)
        if not l:
            return
        n = sum(1 for r in l.get("sorok") or [] if r.get("vid"))
        if QMessageBox.question(self, "Leltár visszavonása?",
                                f"{hely_txt(l['raklap'])} ({str(l.get('lezaras') or '')[:16]}, {l.get('eszkoz', '')}) leltárának visszavonása?\n\n" +
                                (f"{n} felülírt érték visszaáll a leltár előttire." if n else "Felülírás nem volt, csak a leltár lesz érvénytelen.")
                                ) != QMessageBox.StandardButton.Yes:
            return
        self._lel_job("visszavon", "POST", f"/leltar/{lid}/visszavon", {"eszkoz": GEP_NEV})

    def init_map_tab(self):
        layout = QVBoxLayout(self.tab_map)

        # 🔍 KERESŐSÁV: pontos egyezés kiemelése és odagörgetés
        search_bar = QHBoxLayout()
        self.map_search_input = QLineEdit()
        self.map_search_input.setPlaceholderText("🔍 Keresés a raktártérképen (pontos érték, pl. 112 vagy KISRAKTÁR 80) - Enter")
        self.map_search_input.setStyleSheet("font-size: 15px; padding: 6px;")
        self.map_search_input.returnPressed.connect(self.search_on_map)
        search_bar.addWidget(self.map_search_input, stretch=3)

        btn_map_search = QPushButton("🔍 Keresés")
        btn_map_search.setStyleSheet("background-color: #ff0055; color: white; font-weight: bold; padding: 6px 14px;")
        btn_map_search.clicked.connect(self.search_on_map)
        search_bar.addWidget(btn_map_search)

        btn_map_reload = QPushButton("🔄 Térkép frissítése")
        btn_map_reload.setToolTip("A térképet újra letölti a Google táblázatból (ha közben módosult)")
        btn_map_reload.clicked.connect(lambda: self.search_on_map(force_reload=True))
        search_bar.addWidget(btn_map_reload)

        btn_edit_txt = "✏️ Térkép Szerkesztése (Google Sheets)"
        btn_edit = QPushButton(btn_edit_txt)
        btn_edit.clicked.connect(
            lambda: self.run_button_action(
                btn_edit, 
                btn_edit_txt, 
                lambda: QDesktopServices.openUrl(QUrl(MAP_EDIT_URL))
            )
        )
        search_bar.addWidget(btn_edit)
        layout.addLayout(search_bar)

        self.map_search_status = QLabel("Írd be a keresett értéket, és nyomj Entert. Több találatnál az újabb Enter a következőre ugrik.")
        self.map_search_status.setStyleSheet("color: #aaaaaa; font-size: 13px;")
        layout.addWidget(self.map_search_status)

        self._map_data_cache = None
        self._map_last_query = None
        self._map_pending_query = None

        # A beépített böngésző CSAK a fül első megnyitásakor indul (lásd ensure_map_view): induláskor
        # Windowson újraépítené az ablakot (elveszne a teljes méret), és sok memóriát / időt venne el.
        self._map_layout = layout
        self._map_placeholder = QLabel("🗺️ A térkép betöltése...")
        self._map_placeholder.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(self._map_placeholder, stretch=1)

    def ensure_map_view(self):
        if hasattr(self, "web_view"):
            return True
        try:
            self.web_view = QWebEngineView()
            self.web_view.setPage(SilentWebEnginePage(self.web_view))
            embed_url = f"https://docs.google.com/spreadsheets/d/{MAP_SHEETS_ID}/htmlembed?widget=true&headers=false"
            self.web_view.setUrl(QUrl(embed_url))
            self._map_layout.replaceWidget(self._map_placeholder, self.web_view)
            self._map_placeholder.deleteLater()
        except Exception:
            self._map_placeholder.setText("🗺️ A beágyazott böngésző nem érhető el.")
            return False
        # A böngésző első megjelenése újraépítheti az ablakot - ha emiatt kicsi lett, vissza teljes méretre
        QTimer.singleShot(400, lambda: ensure_maximized(self))
        return True

    def search_on_map(self, *_, force_reload=False):
        """🔍 Raktártérkép kereső: a pontosan egyező cellá(ka)t színesen kiemeli és odagörget."""
        query = self.map_search_input.text().strip()
        if not query:
            self.map_search_status.setText("⚠️ Írj be egy keresett értéket!")
            return
        if not self.ensure_map_view():
            return

        # Ugyanaz a keresés újra (és már látszik): a következő találatra lép
        if (not force_reload and self._map_data_cache is not None
                and _norm_map_text(query) == self._map_last_query and getattr(self, "_map_match_count", 0) > 1):
            def on_next(pos):
                if pos:
                    self.map_search_status.setText(f"📍 '{query}': {pos}. / {self._map_match_count} találat (Enter = következő)")
            self.web_view.page().runJavaScript("nextMatch()", on_next)
            return

        if force_reload or self._map_data_cache is None:
            # A térkép letöltése háttérszálon (a felület nem akad meg)
            self._map_pending_query = query
            if getattr(self, "map_tab_worker", None) and self.map_tab_worker.isRunning():
                return
            self.map_search_status.setText("⏳ Raktártérkép betöltése...")
            self.map_tab_worker = MapWorker()
            self.map_tab_worker.finished_signal.connect(self._on_map_tab_loaded)
            self.map_tab_worker.error_signal.connect(lambda e: self.map_search_status.setText(f"❌ {e}"))
            self.map_tab_worker.start()
            return

        self._render_map_search(query)

    def _on_map_tab_loaded(self, map_data):
        self._map_data_cache = map_data
        if self._map_pending_query:
            self._render_map_search(self._map_pending_query)

    def _render_map_search(self, query):
        data = self._map_data_cache or []
        count = sum(1 for row in data for cell in row if match_location_exact(cell["value"], query))
        self._map_last_query = _norm_map_text(query)
        self._map_match_count = count
        self.web_view.setHtml(render_interactive_map_html(data, query, exact=True))
        if count == 0:
            self.map_search_status.setText(f"❌ '{query}' nem található a raktártérképen (csak pontos egyezést keres).")
        elif count == 1:
            self.map_search_status.setText(f"📍 '{query}': 1 találat")
        else:
            self.map_search_status.setText(f"📍 '{query}': 1. / {count} találat (Enter = következő)")

    # -------------------------------------------------------------------
    # TAB 7: NAPLÓ TAB (ADATBÁZIS ALAPÚ)
    # -------------------------------------------------------------------
    # Napló fül eseménytípus-szűrői: (felirat, SQL feltétel az esemeny oszlopra)
    LOG_FILTERS = [
        ("Minden esemény", None),
        ("➕ Felvétel", "esemeny LIKE '%FELVÉTEL%'"),
        ("✏️ Módosítás", "(esemeny LIKE '%MÓDOSÍTÁS%' AND esemeny NOT LIKE 'ELADÁS%')"),
        ("🗑️ Törlés / kiürítés", "(esemeny LIKE '%TÖRLÉS%' OR esemeny LIKE '%VISSZAVONÁS%') AND esemeny NOT LIKE 'ELADÁS%'"),
        ("💰 Minden eladás", "esemeny LIKE 'ELADÁS: %'"),
        ("⚠️ Figyelmeztetések / ütközések", "(esemeny LIKE 'FIGYELMEZTETÉS%' OR esemeny LIKE '%ÜTKÖZÉS%')"),
        ("🌐 Másik gépről / Google Sheets-ből", "(forras LIKE '%Másik gép%' OR forras LIKE '%Google Sheets%')"),
    ]

    def init_log_tab(self):
        layout = QVBoxLayout(self.tab_log)
        layout.setContentsMargins(10, 10, 10, 10)

        top_bar = QHBoxLayout()
        lbl_title = QLabel("📜 Raktári események")
        lbl_title.setStyleSheet("font-size: 16px; font-weight: bold; color: #29b6f6;")
        top_bar.addWidget(lbl_title)

        self.db_log_type = QComboBox()
        for label, _ in self.LOG_FILTERS:
            self.db_log_type.addItem(label)
        for pf, cfg in SALES_SHEETS.items():
            self.db_log_type.addItem(f"💰 {cfg['cim']}")
        self.db_log_type.currentIndexChanged.connect(self.refresh_db_log_view)
        top_bar.addWidget(self.db_log_type)
        top_bar.addStretch()

        self.db_log_search = QLineEdit()
        self.db_log_search.setPlaceholderText("🔍 R18-kód / szöveg (Enter)")
        self.db_log_search.setFixedWidth(240)
        self.db_log_search.returnPressed.connect(self.refresh_db_log_view)
        top_bar.addWidget(self.db_log_search)

        top_bar.addWidget(QLabel("Ettől a naptól:"))
        self.db_log_date = QDateEdit()
        self.db_log_date.setCalendarPopup(True)
        self.db_log_date.setDate(QDate.currentDate().addDays(-7))
        self.db_log_date.setStyleSheet("background-color: #2b2b2b; color: #fff; padding: 4px; font-weight: bold;")
        self.db_log_date.dateChanged.connect(self.refresh_db_log_view)
        top_bar.addWidget(self.db_log_date)

        btn_refresh_log = QPushButton("🔄 Frissítés")
        btn_refresh_log.setStyleSheet("background-color: #00e676; color: black; font-weight: bold; padding: 6px 12px;")
        btn_refresh_log.clicked.connect(self.refresh_db_log_view)
        top_bar.addWidget(btn_refresh_log)
        layout.addLayout(top_bar)

        self.db_log_count_lbl = QLabel("")
        self.db_log_count_lbl.setStyleSheet("color: #888;")
        layout.addWidget(self.db_log_count_lbl)

        table_style = """
            QTableWidget { background-color: #21252b; alternate-background-color: #282c34; color: #e0e0e0;
                           gridline-color: #3b4048; font-size: 13px; }
            QHeaderView::section { background-color: #1b1e23; color: #61afef; font-weight: bold; padding: 6px; border: 1px solid #3b4048; }
            QTableWidget::item:selected { background-color: #2f5f8f; color: #ffffff; }
        """
        from PyQt6.QtWidgets import QSplitter
        splitter = QSplitter(Qt.Orientation.Vertical)

        self.db_log_table = QTableWidget()
        self.db_log_table.setColumnCount(8)
        self.db_log_table.setHorizontalHeaderLabels(["Időpont", "Esemény", "Honnan / hogyan", "R18-kód", "Márka", "Típus",
                                                     "Cikkszám", "Mi változott / részletek"])
        self.db_log_table.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        self.db_log_table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self.db_log_table.setSelectionMode(QAbstractItemView.SelectionMode.SingleSelection)
        self.db_log_table.setAlternatingRowColors(True)
        self.db_log_table.verticalHeader().setVisible(False)
        self.db_log_table.setStyleSheet(table_style)
        for i, w in enumerate([140, 220, 260, 95, 110, 170, 160]):
            self.db_log_table.setColumnWidth(i, w)
        self.db_log_table.horizontalHeader().setStretchLastSection(True)
        self.db_log_table.currentCellChanged.connect(lambda r, *_: self.show_log_entry_detail(r))
        splitter.addWidget(self.db_log_table)

        detail_box = QWidget()
        detail_lay = QVBoxLayout(detail_box)
        detail_lay.setContentsMargins(0, 6, 0, 0)
        self.db_log_detail_lbl = QLabel("Válassz ki egy eseményt a listában - itt látszik a termék teljes adata ELŐTTE és UTÁNA.")
        self.db_log_detail_lbl.setStyleSheet("color: #29b6f6; font-weight: bold;")
        self.db_log_detail_lbl.setWordWrap(True)
        detail_lay.addWidget(self.db_log_detail_lbl)
        self.db_log_detail = QTableWidget()
        self.db_log_detail.setColumnCount(3)
        self.db_log_detail.setHorizontalHeaderLabels(["Oszlop", "⬅️ ELŐTTE", "➡️ UTÁNA"])
        self.db_log_detail.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        self.db_log_detail.verticalHeader().setVisible(False)
        self.db_log_detail.setStyleSheet(table_style)
        self.db_log_detail.setColumnWidth(0, 160)
        self.db_log_detail.setColumnWidth(1, 380)
        self.db_log_detail.horizontalHeader().setStretchLastSection(True)
        detail_lay.addWidget(self.db_log_detail)
        splitter.addWidget(detail_box)
        splitter.setSizes([450, 330])
        layout.addWidget(splitter)

        self._db_log_rows = []
        self.refresh_db_log_view()

    @staticmethod
    def _log_event_color(esemeny):
        e = esemeny.upper()
        if e.startswith("FIGYELMEZTETÉS") or "ÜTKÖZÉS" in e:
            return "#ff9800"
        if e.startswith("ELADÁS:"):
            return "#ffd54f"
        if "TÖRLÉS" in e or "VISSZAVONÁS" in e:
            return "#ff5252"
        if "FELVÉTEL" in e:
            return "#00e676"
        if "MÓDOSÍTÁS" in e:
            return "#29b6f6"
        return "#bbbbbb"

    def refresh_db_log_view(self, *_):
        if not hasattr(self, 'db_log_table'):
            return
        try:
            date_from = self.db_log_date.date().toString("yyyy-MM-dd")
            search = self.db_log_search.text().strip()
            sql = "SELECT id, idopont, r18_kod, esemeny, reszletek, elotte, utana, forras FROM naplo WHERE idopont >= ?"
            params = [f"{date_from} 00:00:00"]
            idx = self.db_log_type.currentIndex()
            if idx < len(self.LOG_FILTERS):
                cond = self.LOG_FILTERS[idx][1]
                if cond:
                    sql += f" AND ({cond})"
            else:
                pf_cim = list(SALES_SHEETS.values())[idx - len(self.LOG_FILTERS)]["cim"]
                sql += " AND esemeny = ?"
                params.append(f"ELADÁS: {pf_cim}")
            if search:
                sql += " AND (r18_kod LIKE ? OR esemeny LIKE ? OR reszletek LIKE ? OR elotte LIKE ? OR utana LIKE ? OR forras LIKE ?)"
                params += [f"%{search}%"] * 6
            sql += " ORDER BY id DESC LIMIT 1000"
            conn = get_db_connection()
            cursor = conn.cursor()
            cursor.execute(sql, params)
            rows = [dict(r) for r in cursor.fetchall()]
            conn.close()
        except Exception as e:
            self.db_log_count_lbl.setText(f"❌ Hiba a napló betöltésekor: {e}")
            return

        self._db_log_rows = rows
        self.db_log_table.setUpdatesEnabled(False)
        self.db_log_table.setRowCount(len(rows))
        for ri, r in enumerate(rows):
            elotte = json.loads(r["elotte"]) if r.get("elotte") else None
            utana = json.loads(r["utana"]) if r.get("utana") else None
            prod = utana or elotte or {}
            esemeny = safe_str(r.get("esemeny"))
            valtozas = safe_str(r.get("reszletek"))
            if elotte and utana and not esemeny.upper().startswith("ELADÁS"):
                valtozas = describe_changes(elotte, utana) or valtozas
            forras = safe_str(r.get("forras"))
            vals = [safe_str(r.get("idopont")), esemeny, forras, safe_str(r.get("r18_kod")),
                    prod.get("C", ""), prod.get("D", ""), prod.get("E", ""), valtozas]
            color = QColor(self._log_event_color(esemeny))
            warn = "⚠️" in forras
            for ci, v in enumerate(vals):
                item = QTableWidgetItem(v)
                if ci == 1:
                    item.setForeground(color)
                    f = item.font(); f.setBold(True); item.setFont(f)
                    expl = explain_event(esemeny)
                    if expl:
                        item.setToolTip("💡 " + expl)
                if ci in (2, 7):
                    item.setToolTip(v)
                if ci == 2 and warn:
                    item.setForeground(QColor("#ffb74d"))
                self.db_log_table.setItem(ri, ci, item)
        self.db_log_table.setUpdatesEnabled(True)
        more = " - csak a legutóbbi 1000 látszik, szűkítsd a dátumot / keresést" if len(rows) >= 1000 else ""
        self.db_log_count_lbl.setText(f"{len(rows)} esemény {date_from} óta{more}")
        self.db_log_detail.setRowCount(0)
        if rows:
            self.db_log_table.selectRow(0)
            self.show_log_entry_detail(0)

    def show_log_entry_detail(self, row):
        """Az alsó részen: a kiválasztott esemény előtte / utána teljes termékadata, a változások kiemelve."""
        if row < 0 or row >= len(self._db_log_rows):
            return
        self.render_log_detail(self._db_log_rows[row], self.db_log_detail_lbl, self.db_log_detail)

    def render_log_detail(self, r, lbl, table):
        """Egy napló-bejegyzés fejléce + az ELŐTTE / UTÁNA táblázat (a Napló és az Admin fül használja)."""
        elotte = json.loads(r["elotte"]) if r.get("elotte") else None
        utana = json.loads(r["utana"]) if r.get("utana") else None
        esemeny = safe_str(r.get("esemeny"))
        import html as _html
        head = f"<b>{_html.escape(safe_str(r.get('idopont')))}  |  {_html.escape(esemeny)}  |  {_html.escape(safe_str(r.get('r18_kod')))}</b>"
        if r.get("forras"):
            head += f"<br>🧭 <b>Honnan / hogyan:</b> {_html.escape(safe_str(r.get('forras')))}"
        if r.get("reszletek"):
            head += f"<br>📝 {_html.escape(safe_str(r.get('reszletek')))}"
        expl = explain_event(esemeny)
        if expl:
            head += f"<br><span style='color:#ffe082;'>💡 <b>Mit jelent:</b> {_html.escape(expl)}</span>"
        if elotte is None and utana is None:
            head += "<br><span style='color:#888;'>(ehhez a bejegyzéshez nincs előtte/utána termékadat)</span>"
        lbl.setTextFormat(Qt.TextFormat.RichText)
        lbl.setText(head)

        table.setRowCount(len(NAPLO_COLS))
        changed_bg, new_bg, gone_bg = QColor("#5c4a00"), QColor("#1b5e20"), QColor("#7f1d1d")
        for i, c in enumerate(NAPLO_COLS):
            ov = (elotte or {}).get(c, "") if elotte is not None else ""
            nv = (utana or {}).get(c, "") if utana is not None else ""
            items = [QTableWidgetItem(f"{c}: {COLUMN_LABELS.get(c, c)}"),
                     QTableWidgetItem(ov if elotte is not None else "—"),
                     QTableWidgetItem(nv if utana is not None else ("— (törölve)" if elotte is not None else "—"))]
            if elotte is not None and utana is None:
                items[2].setBackground(gone_bg)
            elif elotte is None and utana is not None:
                items[2].setBackground(new_bg)
            elif elotte is not None and str(ov).strip() != str(nv).strip():
                for it in items:
                    it.setBackground(changed_bg)
                f = items[0].font(); f.setBold(True); items[0].setFont(f)
            for ci, it in enumerate(items):
                table.setItem(i, ci, it)

    # -------------------------------------------------------------------
    # 🛡️ ADMIN FÜL: minden változás, törölt termékek (visszaállítással), termék-történet, feltöltésre váró
    # -------------------------------------------------------------------
    ADMIN_TABLE_STYLE = """
        QTableWidget { background-color: #21252b; alternate-background-color: #282c34; color: #e0e0e0;
                       gridline-color: #3b4048; font-size: 13px; }
        QHeaderView::section { background-color: #1b1e23; color: #61afef; font-weight: bold; padding: 6px; border: 1px solid #3b4048; }
        QTableWidget::item:selected { background-color: #2f5f8f; color: #ffffff; }
    """

    def _admin_table(self, headers, widths=None):
        t = QTableWidget()
        t.setColumnCount(len(headers))
        t.setHorizontalHeaderLabels(headers)
        t.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        t.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        t.setAlternatingRowColors(True)
        t.verticalHeader().setVisible(False)
        t.setStyleSheet(self.ADMIN_TABLE_STYLE)
        for i, w in enumerate(widths or []):
            t.setColumnWidth(i, w)
        t.horizontalHeader().setStretchLastSection(True)
        return t

    def _detail_widgets(self):
        box = QWidget()
        lay = QVBoxLayout(box)
        lay.setContentsMargins(0, 6, 0, 0)
        lbl = QLabel("Válassz ki egy sort - itt látszik a termék teljes adata ELŐTTE és UTÁNA.")
        lbl.setWordWrap(True)
        lbl.setStyleSheet("color: #29b6f6; font-weight: bold;")
        lay.addWidget(lbl)
        tbl = self._admin_table(["Oszlop", "⬅️ ELŐTTE", "➡️ UTÁNA"], [160, 380])
        lay.addWidget(tbl)
        return box, lbl, tbl

    def init_admin_tab(self):
        layout = QVBoxLayout(self.tab_admin)
        title = QLabel("🛡️ Admin - minden változás, minden törölt termék teljes adattal, termék-történet")
        title.setStyleSheet("font-size: 16px; font-weight: bold; color: #29b6f6;")
        layout.addWidget(title)
        self.admin_tabs = QTabWidget()
        layout.addWidget(self.admin_tabs)

        # --- 1. Minden változás ---
        w1 = QWidget(); l1 = QVBoxLayout(w1)
        bar = QHBoxLayout()
        bar.addWidget(QLabel("Ettől:"))
        self.adm_from = QDateEdit(); self.adm_from.setCalendarPopup(True); self.adm_from.setDate(QDate.currentDate().addDays(-30))
        bar.addWidget(self.adm_from)
        bar.addWidget(QLabel("Eddig:"))
        self.adm_to = QDateEdit(); self.adm_to.setCalendarPopup(True); self.adm_to.setDate(QDate.currentDate())
        bar.addWidget(self.adm_to)
        self.adm_type = QComboBox()
        for label, _ in self.LOG_FILTERS:
            self.adm_type.addItem(label)
        bar.addWidget(self.adm_type)
        self.adm_search = QLineEdit(); self.adm_search.setPlaceholderText("🔍 R18 / szöveg / forrás (Enter)")
        self.adm_search.returnPressed.connect(self.admin_load_all)
        bar.addWidget(self.adm_search, stretch=1)
        b = QPushButton("🔄 Lekérés"); b.setStyleSheet("background-color: #00e676; color: black; font-weight: bold;")
        b.clicked.connect(self.admin_load_all); bar.addWidget(b)
        b = QPushButton("📤 Export (CSV / Excel)"); b.clicked.connect(self.admin_export_all); bar.addWidget(b)
        l1.addLayout(bar)
        self.adm_all_count = QLabel(""); self.adm_all_count.setStyleSheet("color: #888;")
        l1.addWidget(self.adm_all_count)
        sp = QSplitter(Qt.Orientation.Vertical)
        self.adm_all = self._admin_table(["Időpont", "Esemény", "Honnan / hogyan", "R18-kód", "Márka", "Típus", "Cikkszám", "Mi változott"],
                                         [140, 220, 260, 95, 110, 170, 160])
        self.adm_all.currentCellChanged.connect(lambda r, *_: self._admin_show(self._adm_all_rows, r, self.adm_all_lbl, self.adm_all_det))
        sp.addWidget(self.adm_all)
        box, self.adm_all_lbl, self.adm_all_det = self._detail_widgets()
        sp.addWidget(box); sp.setSizes([420, 320])
        l1.addWidget(sp)
        self._adm_all_rows = []
        self.admin_tabs.addTab(w1, "📜 Minden változás")

        # --- 2. Törölt termékek ---
        w2 = QWidget(); l2 = QVBoxLayout(w2)
        bar = QHBoxLayout()
        self.adm_del_search = QLineEdit(); self.adm_del_search.setPlaceholderText("🔍 R18 / márka / típus / cikkszám (Enter)")
        self.adm_del_search.returnPressed.connect(self.admin_load_deleted)
        bar.addWidget(self.adm_del_search, stretch=1)
        b = QPushButton("🔄 Lekérés"); b.setStyleSheet("background-color: #00e676; color: black; font-weight: bold;")
        b.clicked.connect(self.admin_load_deleted); bar.addWidget(b)
        b = QPushButton("♻️ Kijelölt visszaállítása"); b.setStyleSheet("background-color: #ff9800; color: black; font-weight: bold;")
        b.setToolTip("A törölt termék teljes adatát visszaírja az R18-kódja alá (helyi + szerver + Google Sheets)")
        b.clicked.connect(self.admin_restore_deleted); bar.addWidget(b)
        l2.addLayout(bar)
        self.adm_del_count = QLabel(""); self.adm_del_count.setStyleSheet("color: #888;")
        l2.addWidget(self.adm_del_count)
        self.adm_del = self._admin_table(["Törlés ideje", "Honnan / hogyan"] + [f"{c}: {COLUMN_LABELS.get(c, c)}" for c in NAPLO_COLS],
                                         [140, 240, 95, 95, 110, 190, 170, 45, 85, 60, 70, 40, 60, 110, 55, 100, 120])
        l2.addWidget(self.adm_del)
        self._adm_del_rows = []
        self.admin_tabs.addTab(w2, "🗑️ Törölt termékek")

        # --- 3. Termék története ---
        w3 = QWidget(); l3 = QVBoxLayout(w3)
        bar = QHBoxLayout()
        self.adm_hist_code = QLineEdit(); self.adm_hist_code.setPlaceholderText("R18-kód (pl. R18-12960) - Enter / csippantás")
        self.adm_hist_code.returnPressed.connect(self.admin_load_history)
        bar.addWidget(self.adm_hist_code, stretch=1)
        b = QPushButton("🔎 Történet"); b.setStyleSheet("background-color: #29b6f6; color: black; font-weight: bold;")
        b.clicked.connect(self.admin_load_history); bar.addWidget(b)
        l3.addLayout(bar)
        self.adm_hist_now = QLabel(""); self.adm_hist_now.setWordWrap(True); self.adm_hist_now.setTextFormat(Qt.TextFormat.RichText)
        self.adm_hist_now.setStyleSheet("background: #1e1e1e; padding: 8px; border: 1px solid #333; border-radius: 6px;")
        l3.addWidget(self.adm_hist_now)
        sp = QSplitter(Qt.Orientation.Vertical)
        self.adm_hist = self._admin_table(["Időpont", "Esemény", "Honnan / hogyan", "Mi változott"], [140, 240, 300])
        self.adm_hist.currentCellChanged.connect(lambda r, *_: self._admin_show(self._adm_hist_rows, r, self.adm_hist_lbl, self.adm_hist_det))
        sp.addWidget(self.adm_hist)
        box, self.adm_hist_lbl, self.adm_hist_det = self._detail_widgets()
        sp.addWidget(box); sp.setSizes([360, 320])
        l3.addWidget(sp)
        self._adm_hist_rows = []
        self.admin_tabs.addTab(w3, "🔎 Termék története")

        # --- 4. Feltöltésre vár ---
        w4 = QWidget(); l4 = QVBoxLayout(w4)
        bar = QHBoxLayout()
        b = QPushButton("🔄 Lekérés"); b.setStyleSheet("background-color: #00e676; color: black; font-weight: bold;")
        b.clicked.connect(self.admin_load_outbox); bar.addWidget(b)
        b = QPushButton("🔁 Újrapróbálás most")
        b.clicked.connect(lambda: (flush_outbox_async(), sync_sales_async(), self.trigger_toast("🔁 Feltöltés elindítva"),
                                   QTimer.singleShot(4000, self.admin_load_outbox)))
        bar.addWidget(b)
        bar.addStretch()
        l4.addLayout(bar)
        self.adm_out_count = QLabel(""); self.adm_out_count.setStyleSheet("color: #888;")
        l4.addWidget(self.adm_out_count)
        self.adm_out = self._admin_table(["Felvéve", "Művelet", "R18-kód", "VPS", "Google Sheets", "Próbálkozás", "Utolsó próba", "Adat"],
                                         [140, 90, 95, 60, 110, 90, 140])
        l4.addWidget(self.adm_out)
        self.admin_tabs.addTab(w4, "⏳ Feltöltésre vár")

        # --- 5. Gépek és verziók ---
        w5 = QWidget(); l5 = QVBoxLayout(w5)
        bar = QHBoxLayout()
        b = QPushButton("🔄 Frissítés"); b.setStyleSheet("background-color: #00e676; color: black; font-weight: bold;")
        b.clicked.connect(self.admin_load_clients); bar.addWidget(b)
        bar.addStretch()
        l5.addLayout(bar)
        self.adm_cli_info = QLabel(""); self.adm_cli_info.setStyleSheet("color: #29b6f6; font-weight: bold;")
        self.adm_cli_info.setWordWrap(True)
        l5.addWidget(self.adm_cli_info)
        self.adm_cli = self._admin_table(["Típus", "Név", "Program verzió", "PDA app", "Állapot", "Utoljára jelentkezett",
                                          "Felhasználó", "IP"], [80, 200, 230, 90, 110, 150, 110])
        l5.addWidget(self.adm_cli)
        hint = QLabel("Minden gépi program és PDA kb. 20-30 másodpercenként jelentkezik a szerveren. "
                      "Online = az elmúlt 2 percben jelentkezett. A ⚠️ RÉGI jelölésű gépen cseréld le az app_pyqt6CLOUD.py-t.")
        hint.setWordWrap(True); hint.setStyleSheet("color: #888;")
        l5.addWidget(hint)
        self.admin_tabs.addTab(w5, "🖥️ Gépek és verziók")

        self.admin_tabs.currentChanged.connect(self._admin_tab_changed)

    def _admin_tab_changed(self, i):
        [self.admin_load_all, self.admin_load_deleted, lambda: None, self.admin_load_outbox, self.admin_load_clients][i]()

    # ---- 🖥️ GÉPEK ÉS VERZIÓK ----
    def on_version_info(self, leg):
        """A szerver szerinti legfrissebb gépi verzió; ha ez a gép régebbi, a fejlécben figyelmeztet."""
        latest = str((leg or {}).get("gep") or "")
        if getattr(self, "_upd", None):
            self.lbl_version_warn.hide()       # a nagy villogó sáv már jelzi
        elif latest and _ver_key(latest) > _ver_key(APP_VERSION):
            self.lbl_version_warn.setText(f"⚠️ Van frissebb program: {latest} (ezen a gépen: {APP_VERSION})")
            self.lbl_version_warn.show()
        else:
            self.lbl_version_warn.hide()

    # ==========================================================
    # 🛒 WEBES RENDELÉSEK – Shopify + eBay (a VPS 5 percenként lekéri; itt 1 kattintással eladás)
    # ==========================================================
    WEB_COLS = ["Forrás", "Dátum", "Rendelés", "Vevő", "Cím", "Termék (webes név)", "SKU / eBay ID", "Db", "Ár",
                "R18 (párosítva)", "Raktáron", "Webes állapot", "Rögzítés"]

    def init_web_tab(self):
        lay = QVBoxLayout(self.tab_web)
        bar = QHBoxLayout()
        self.web_filter = QComboBox()
        for txt, val in (("🟠 Rögzítésre vár", "nyitott"), ("Mind", "mind"), ("✅ Rögzítve", "rogzitve"),
                         ("🚫 Kihagyva", "kihagyva"), ("❌ Lemondva", "lemondva")):
            self.web_filter.addItem(txt, val)
        self.web_filter.setMinimumWidth(190)
        self.web_filter.currentIndexChanged.connect(self.web_render)
        bar.addWidget(self.web_filter)
        self.web_search = QLineEdit()
        self.web_search.setPlaceholderText("🔍 Keresés (vevő, termék, R18, rendelésszám...)")
        self.web_search.textChanged.connect(self.web_render)
        bar.addWidget(self.web_search, stretch=2)
        b = QPushButton("🔄 Lekérés most (Shopify + eBay)")
        b.setStyleSheet("background-color: #29b6f6; color: black; font-weight: bold;")
        b.clicked.connect(lambda: self.web_load(refresh=True))
        bar.addWidget(b)
        lay.addLayout(bar)
        self.web_status = QLabel("⏳ Betöltés...")
        self.web_status.setTextFormat(Qt.TextFormat.RichText)
        self.web_status.setStyleSheet("padding: 4px; font-size: 13px;")
        lay.addWidget(self.web_status)

        self.web_table = self._admin_table(self.WEB_COLS, [95, 120, 120, 170, 220, 230, 120, 40, 90, 170, 170, 150])
        self.web_table.setWordWrap(True)
        self.web_table.cellDoubleClicked.connect(lambda r, c: self.web_record())
        lay.addWidget(self.web_table, stretch=1)

        bb = QHBoxLayout()
        self.btn_web_record = QPushButton("💰 Rögzítés eladásként (kijelölt)")
        self.btn_web_record.setStyleSheet("background-color: #00e676; color: black; font-weight: bold; font-size: 15px; padding: 8px 18px;")
        self.btn_web_record.clicked.connect(self.web_record)
        bb.addWidget(self.btn_web_record)
        b = QPushButton("🔗 Kézi párosítás (R18)")
        b.clicked.connect(self.web_match)
        bb.addWidget(b)
        b = QPushButton("🚫 Kihagyás / visszavétel")
        b.setToolTip("Pl. nem raktári termék: nem kell rögzíteni, eltűnik a „Rögzítésre vár” listából")
        b.clicked.connect(self.web_skip)
        bb.addWidget(b)
        bb.addStretch()
        hint = QLabel("Dupla kattintás egy sorra = rögzítés")
        hint.setStyleSheet("color: #888;")
        bb.addWidget(hint)
        lay.addLayout(bb)

        self._web_rows = []
        self._web_info = {}
        self._web_known_open = None
        self._web_busy = False
        self.web_signal.connect(self.on_web_signal)
        self.web_timer = QTimer(self)
        self.web_timer.timeout.connect(self.web_load)
        self.web_timer.start(2 * 60 * 1000)
        QTimer.singleShot(18000, self.web_load)

    def web_load(self, refresh=False):
        if self._web_busy:
            return
        self._web_busy = True
        if refresh:
            self.web_status.setText("⏳ Lekérés a Shopify-ról és az eBay-ről...")

        def job():
            try:
                if refresh:
                    r = vps_request("POST", f"{VPS_URL}/web/orders/refresh", timeout=90)
                    if r.status_code == 404:
                        raise RuntimeError("A VPS-en még a régi server_api.py fut (nincs webes rendelés rész) – töltsd fel az újat")
                res = vps_request("GET", f"{VPS_URL}/web/orders", params={"limit": 500}, timeout=15)
                if res.status_code == 404:
                    raise RuntimeError("A VPS-en még a régi server_api.py fut (nincs webes rendelés rész) – töltsd fel az újat")
                if res.status_code != 200:
                    raise RuntimeError(f"Szerver hiba (HTTP {res.status_code})")
                self.web_signal.emit("lista", dict(res.json(), _kezi=refresh))
            except Exception as e:
                self.web_signal.emit("hiba", {"message": str(e)})
        threading.Thread(target=job, daemon=True).start()

    def on_web_signal(self, kind, d):
        self._web_busy = False
        if kind == "hiba":
            self.web_status.setText(f"<span style='color:#ff5252'>❌ {d.get('message', '')}</span>")
            return
        self._web_rows = d.get("rendelesek") or []
        self._web_info = d
        nyitott = int(d.get("nyitott") or 0)
        idx = self.tabs.indexOf(self.tab_web)
        self.tabs.setTabText(idx, f"🛒 Webes rendelések ({nyitott})" if nyitott else "🛒 Webes rendelések")
        self.tabs.tabBar().setTabTextColor(idx, QColor("#ff9800") if nyitott else QColor("#ffffff"))
        if self._web_known_open is not None and nyitott > self._web_known_open:
            uj = [r for r in self._web_rows if not r.get("rogzitve") and not r.get("kihagyva") and not r.get("lemondva")][:3]
            self.trigger_toast("🛒 ÚJ WEBES RENDELÉS: " + "; ".join(f"{r['forras']} – {r['termek'][:30]}" for r in uj), 6000)
            self.append_process_log(f"🛒 Új webes rendelés érkezett ({nyitott} vár rögzítésre)")
        self._web_known_open = nyitott
        self.web_render()
        if d.get("_kezi"):
            self.trigger_toast("🔄 Webes rendelések lekérve")

    def _web_code(self, r):
        return norm_code(r.get("kezi_A") or r.get("A") or "")

    def web_render(self):
        if not hasattr(self, "web_table"):
            return
        d = self._web_info or {}
        be = d.get("beallitva") or {}
        hibak = d.get("hibak") or {}

        def allapot(nev, kulcs, extra_ok=True, extra_txt=""):
            if hibak.get(kulcs):
                return f"<span style='color:#ff5252'>❌ {nev}: {hibak[kulcs][:140]}</span>"
            if not be.get(kulcs):
                return f"<span style='color:#888'>⚪ {nev}: nincs beállítva</span>"
            if not extra_ok:
                return f"<span style='color:#ffb74d'>⚠️ {nev}: {extra_txt}</span>"
            return f"<span style='color:#00e676'>✅ {nev}</span>"
        self.web_status.setText(
            allapot("Shopify", "shopify") + " &nbsp;·&nbsp; " +
            allapot("eBay", "ebay", be.get("ebay_engedelyezve"), "még nincs engedélyezve (lásd TELEPITES_webrendelesek.md)") +
            f" &nbsp;·&nbsp; Utolsó lekérés: {d.get('utolso_lekeres') or '–'} &nbsp;·&nbsp; "
            f"<b style='color:#ff9800'>Rögzítésre vár: {d.get('nyitott', 0)}</b>")

        flt = self.web_filter.currentData()
        q = self.web_search.text().strip().lower()
        rows = []
        for r in self._web_rows:
            st = ("lemondva" if r.get("lemondva") else "rogzitve" if r.get("rogzitve") else
                  "kihagyva" if r.get("kihagyva") else "nyitott")
            if flt != "mind" and st != flt:
                continue
            if q and q not in " ".join(str(r.get(k, "")) for k in ("vevo", "termek", "sku", "ebay_item", "A", "kezi_A",
                                                                      "rendeles_szam", "cim", "email")).lower():
                continue
            rows.append((r, st))
        t = self.web_table
        t.setRowCount(len(rows))
        for i, (r, st) in enumerate(rows):
            code = self._web_code(r)
            local = get_row_snapshot(code) if code else None
            raktar = ""
            if code:
                raktar = (f"{local.get('F') or '1'} db · {local.get('L') or ''} {('Páter ' + local['I']) if local.get('I') else ''}".strip()
                          if local else "– nincs meg ezen a gépen")
            r18 = (f"{code}{' ✋' if r.get('kezi_A') else ''}  " + (f"{local.get('C', '')} {local.get('D', '')}".strip() if local else "")
                   if code else "❓ nincs párosítva")
            rog = {"rogzitve": f"✅ {r.get('rogzitette', '')} {r.get('rogzitve_ido', '')}", "kihagyva": "🚫 kihagyva",
                   "lemondva": "❌ LEMONDVA", "nyitott": "🟠 rögzítésre vár"}[st]
            vals = ["🛍️ Shopify" if r["forras"] == "SHOPIFY" else "🏷️ eBay", r.get("datum", ""), r.get("rendeles_szam", ""),
                    r.get("vevo", "") + (f"\n{r['email']}" if r.get("email") else ""), r.get("cim", ""), r.get("termek", ""),
                    " / ".join(x for x in (r.get("sku"), r.get("ebay_item")) if x), str(r.get("darab", "")),
                    f"{r.get('ar', '')} {r.get('penznem', '')}".strip(), r18, raktar, r.get("statusz", ""), rog]
            for c, v in enumerate(vals):
                it = QTableWidgetItem(v)
                it.setData(Qt.ItemDataRole.UserRole, r["id"])
                if c == 9 and not code:
                    it.setForeground(QColor("#ff5252"))
                if c == 10 and code and not local:
                    it.setForeground(QColor("#ffb74d"))
                if c == 12:
                    it.setForeground(QColor({"nyitott": "#ff9800", "rogzitve": "#00e676", "lemondva": "#ff5252"}.get(st, "#888")))
                t.setItem(i, c, it)
        t.resizeRowsToContents()

    def _web_selected(self):
        r = self.web_table.currentRow()
        it = self.web_table.item(r, 0) if r >= 0 else None
        if not it:
            self.trigger_toast("Előbb jelölj ki egy rendelést a listában.")
            return None
        wid = it.data(Qt.ItemDataRole.UserRole)
        return next((x for x in self._web_rows if x["id"] == wid), None)

    def web_match(self):
        r = self._web_selected()
        if not r:
            return
        cur = self._web_code(r)
        uj, ok = QInputDialog.getText(self, "Kézi párosítás",
                                      f"{r['forras']} rendelés {r['rendeles_szam']}\n{r['termek']}\n\n"
                                      f"Melyik raktári termék ez? Írd be / csippantsd az R18 kódot (üres = párosítás törlése):", text=cur)
        if not ok:
            return
        code = norm_code(fix_scanner_text(uj))
        if code and not get_row_snapshot(code):
            if QMessageBox.question(self, "Kézi párosítás", f"{code} nincs meg ezen a gépen. Így is párosítod?") != QMessageBox.StandardButton.Yes:
                return
        try:
            vps_request("POST", f"{VPS_URL}/web/orders/{r['id']}/match", json={"A": code}, timeout=8)
        except Exception as e:
            QMessageBox.warning(self, "Kézi párosítás", f"Nem sikerült: {e}")
            return
        r["kezi_A"] = code
        self.web_render()
        self.trigger_toast(f"🔗 Párosítva: {code or '(törölve)'}")

    def web_skip(self):
        r = self._web_selected()
        if not r:
            return
        vissza = bool(r.get("kihagyva"))
        if QMessageBox.question(self, "Kihagyás", (f"Visszaveszed a rögzítésre várók közé?\n\n{r['termek']}" if vissza else
                                f"Kihagyod ezt a rendelés-tételt? (pl. nem raktári termék – nem kerül az eladások közé)\n\n"
                                f"{r['forras']} {r['rendeles_szam']} – {r['termek']}")) != QMessageBox.StandardButton.Yes:
            return
        try:
            vps_request("POST", f"{VPS_URL}/web/orders/{r['id']}/skip", json={"kihagyva": not vissza}, timeout=8)
        except Exception as e:
            QMessageBox.warning(self, "Kihagyás", f"Nem sikerült: {e}")
            return
        self.web_load()

    def web_record(self):
        r = self._web_selected()
        if not r:
            return
        if r.get("rogzitve"):
            QMessageBox.information(self, "Rögzítés", f"Ez már rögzítve van ({r.get('rogzitette')} {r.get('rogzitve_ido')}).")
            return
        if r.get("lemondva"):
            QMessageBox.information(self, "Rögzítés", "Ezt a rendelést lemondták – nem kell rögzíteni.")
            return
        code = self._web_code(r)
        if not code:
            QMessageBox.information(self, "Rögzítés", "Ez a rendelés még nincs párosítva raktári termékkel.\n"
                                                      "Előbb nyomd meg a „🔗 Kézi párosítás” gombot és add meg az R18 kódot.")
            return
        local = get_row_snapshot(code)
        try:
            darab = max(1, int(r.get("darab") or 1))
        except ValueError:
            darab = 1
        platform = "EBAY" if r["forras"] == "EBAY" else "WEBSHOP"
        cim = SALES_SHEETS[platform]["cim"]
        if local:
            product = dict(local)
            stock = stock_of(product.get("F"))
            keszlet_txt = (f"{stock} db → {stock - darab} db" if stock - darab > 0 else
                           f"{stock} db → 0 db (ELFOGY – törlődik a raktárból, a ToroltSor lapra kerül)")
        else:
            product = {c: "" for c in ROW_COLS}
            product.update({"A": code, "C": r.get("termek", "")[:60], "N": r.get("ebay_item", "")})
            stock = None
            keszlet_txt = "⚠️ a termék nincs meg ezen a gépen – a készlet NEM változik"
        product["A"] = code
        if QMessageBox.question(self, "💰 Eladás rögzítése?",
                                f"{'🏷️ eBay' if platform == 'EBAY' else '🛍️ Shopify'} rendelés: {r['rendeles_szam']}  ({r.get('datum', '')})\n\n"
                                f"Termék: {code} – {product.get('C', '')} {product.get('D', '')}\n"
                                f"Webes név: {r.get('termek', '')}\n"
                                f"Vevő: {r.get('vevo', '')}\nCím: {r.get('cim', '')}\n"
                                f"Ár: {r.get('ar', '')} {r.get('penznem', '')}   ·   Darab: {darab}\n\n"
                                f"Készlet: {keszlet_txt}\n\nRögzíted eladásként ({cim})?") != QMessageBox.StandardButton.Yes:
            return
        eladas_id = f"W{r['forras'][0]}{r['id']}-{code}-{os.urandom(2).hex()}"
        try:
            res = vps_request("POST", f"{VPS_URL}/web/orders/{r['id']}/claim", json={"eladas_id": eladas_id, "gep": GEP_NEV}, timeout=8)
        except Exception as e:
            QMessageBox.warning(self, "Rögzítés", f"Nem érem el a szervert ({e}) – próbáld újra.")
            return
        if res.status_code != 200:
            try:
                msg = res.json().get("message", "")
            except Exception:
                msg = f"HTTP {res.status_code}"
            QMessageBox.warning(self, "Rögzítés", f"Nem rögzíthető: {msg}")
            self.web_load()
            return
        v = {"datum": r.get("datum") or datetime.datetime.now().strftime("%Y-%m-%d %H:%M"),
             "vevo": r.get("vevo", "") + (f" | {r['cim']}" if r.get("cim") else ""),
             "ar": f"{r.get('ar', '')} {r.get('penznem', '')}".strip(), "darab": darab, "tracking": "",
             "megjegyzes": f"{'eBay' if platform == 'EBAY' else 'Shopify'} rendelés {r['rendeles_szam']}" + (f" | {r['email']}" if r.get("email") else ""),
             "reduce_stock": True}
        stock_msg = self.store_sale(platform, product, stock, v, eladas_id=eladas_id,
                                    sale_src=f"🛒 Webes rendelés ({'eBay' if platform == 'EBAY' else 'Shopify'} {r['rendeles_szam']})")
        if stock_msg is None:
            try:
                vps_request("POST", f"{VPS_URL}/web/orders/{r['id']}/unclaim", timeout=8)
            except Exception:
                pass
            return
        sync_sales_async()
        if hasattr(self, "refresh_sold_table"):
            self.refresh_sold_table()
        self.trigger_toast(f"💰 {code} eladva ({cim}, {darab} db){stock_msg}", 4000)
        self.append_process_log(f"🛒 Webes rendelés rögzítve: {r['forras']} {r['rendeles_szam']} → {code} ({darab} db){stock_msg}", "#00e676")
        self.web_load()

    # ==========================================================
    # 🆕 VERZIÓK – automatikus frissítés a VPS-ről, Python csomagok, régi verziók
    # ==========================================================
    def init_versions_tab(self):
        outer = QVBoxLayout(self.tab_versions)
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        box = QWidget()
        lay = QVBoxLayout(box)
        scroll.setWidget(box)
        outer.addWidget(scroll)

        self.ver_status = QLabel("⏳ Ellenőrzés...")
        self.ver_status.setWordWrap(True)
        self.ver_status.setTextFormat(Qt.TextFormat.RichText)
        self.ver_status.setStyleSheet("font-size: 20px; font-weight: bold; padding: 14px; border-radius: 8px; background: #333;")
        lay.addWidget(self.ver_status)

        self.ver_info = QLabel("")
        self.ver_info.setTextFormat(Qt.TextFormat.RichText)
        self.ver_info.setTextInteractionFlags(Qt.TextInteractionFlag.TextSelectableByMouse)
        self.ver_info.setStyleSheet("font-size: 14px; padding: 6px;")
        lay.addWidget(self.ver_info)

        row = QHBoxLayout()
        b = QPushButton("🔄 Ellenőrzés most")
        b.setStyleSheet("font-size: 15px; font-weight: bold; padding: 8px 16px;")
        b.clicked.connect(lambda: (self.ver_refresh_local(), self.check_app_update(manual=True)))
        row.addWidget(b)
        b = QPushButton("✏️ Gép nevének átírása")
        b.setStyleSheet("font-size: 15px; padding: 8px 16px;")
        b.setToolTip("Ezen a néven látszik ez a gép a többi gépen, a PDA-n és a Gépek listában.\n"
                     "Ha két gépnek ugyanaz a Windows-neve, itt adj nekik különböző nevet (pl. Raktár gép, Iroda gép).")
        b.clicked.connect(self.ver_rename_machine)
        row.addWidget(b)
        self.btn_do_update = QPushButton("⬆️ FRISSÍTÉS MOST")
        self.btn_do_update.setStyleSheet("QPushButton { background-color: #00e676; color: black; font-size: 20px; font-weight: bold; padding: 10px 30px; border-radius: 6px; }"
                                          "QPushButton:disabled { background-color: #555; color: #999; }")
        self.btn_do_update.clicked.connect(self.do_app_update)
        self.btn_do_update.setEnabled(False)
        row.addWidget(self.btn_do_update)
        row.addStretch()
        lay.addLayout(row)

        # Python csomagok
        g = QGroupBox("🐍 Python csomagok (ami a programhoz kell)")
        gl = QVBoxLayout(g)
        self.ver_pkgs = self._admin_table(["Csomag (pip)", "Mire kell", "Ezen a gépen"], [200, 330, 160])
        self.ver_pkgs.setMinimumHeight(260)
        gl.addWidget(self.ver_pkgs)
        self.ver_pip_help = QLabel("")
        self.ver_pip_help.setWordWrap(True)
        self.ver_pip_help.setTextFormat(Qt.TextFormat.RichText)
        self.ver_pip_help.setTextInteractionFlags(Qt.TextInteractionFlag.TextSelectableByMouse)
        self.ver_pip_help.setStyleSheet("font-size: 14px; padding: 8px;")
        gl.addWidget(self.ver_pip_help)
        prow = QHBoxLayout()
        self.btn_pip_copy = QPushButton("📋 Parancs másolása")
        self.btn_pip_copy.clicked.connect(self.ver_copy_pip)
        prow.addWidget(self.btn_pip_copy)
        self.btn_pip_install = QPushButton("⚙️ Telepítés most (automatikusan)")
        self.btn_pip_install.setStyleSheet("background-color: #ff9800; color: black; font-weight: bold; padding: 6px 14px;")
        self.btn_pip_install.clicked.connect(self.ver_pip_install)
        prow.addWidget(self.btn_pip_install)
        prow.addStretch()
        gl.addLayout(prow)
        self.ver_pip_out = QPlainTextEdit()
        self.ver_pip_out.setReadOnly(True)
        self.ver_pip_out.setMaximumHeight(170)
        self.ver_pip_out.setPlaceholderText("Itt látszik az automatikus telepítés kimenete.")
        gl.addWidget(self.ver_pip_out)
        lay.addWidget(g)

        gp = QGroupBox("🖨️ Címkenyomtató ezen a gépen")
        gpl = QVBoxLayout(gp)
        prow2 = QHBoxLayout()
        self.ver_printer_mode = QComboBox()
        for txt, val in (("Automatikus (ha van Zebra nyomtató, itt nyomtat, különben a nyomtatós gépre küldi)", "auto"),
                         ("Ehhez a géphez van kötve a címkenyomtató", "helyi"),
                         ("Nincs nyomtató – mindig a nyomtatós gépre küldje", "tavoli")):
            self.ver_printer_mode.addItem(txt, val)
        self.ver_printer_mode.setCurrentIndex(max(0, self.ver_printer_mode.findData(printer_mode())))
        self.ver_printer_mode.currentIndexChanged.connect(self.ver_set_printer_mode)
        prow2.addWidget(self.ver_printer_mode, stretch=1)
        gpl.addLayout(prow2)
        self.ver_printer_info = QLabel("")
        self.ver_printer_info.setWordWrap(True)
        self.ver_printer_info.setTextFormat(Qt.TextFormat.RichText)
        gpl.addWidget(self.ver_printer_info)
        lay.addWidget(gp)

        g2 = QGroupBox("📝 Mi változott")
        g2l = QVBoxLayout(g2)
        self.ver_changes = QLabel("")
        self.ver_changes.setWordWrap(True)
        self.ver_changes.setTextFormat(Qt.TextFormat.RichText)
        self.ver_changes.setStyleSheet("font-size: 14px; padding: 6px;")
        g2l.addWidget(self.ver_changes)
        lay.addWidget(g2)

        g3 = QGroupBox("🖥️ Melyik gépen melyik verzió fut")
        g3l = QVBoxLayout(g3)
        self.ver_cli = self._admin_table(["Típus", "Név", "Program verzió", "PDA app", "Állapot", "Utoljára jelentkezett",
                                          "Felhasználó", "IP"], [80, 200, 230, 90, 110, 150, 110])
        self.ver_cli.setMinimumHeight(200)
        g3l.addWidget(self.ver_cli)
        lay.addWidget(g3)

        g4 = QGroupBox("💾 Régi verziók (mentések)")
        g4l = QVBoxLayout(g4)
        self.ver_drive = QLabel("")
        self.ver_drive.setWordWrap(True)
        self.ver_drive.setTextFormat(Qt.TextFormat.RichText)
        self.ver_drive.setOpenExternalLinks(True)
        g4l.addWidget(self.ver_drive)
        g4l.addWidget(QLabel("Ezen a gépen (regi_verziok mappa) – kattints egy sorra, majd a Visszaállítás gombra:"))
        self.ver_backups = QListWidget()
        self.ver_backups.setMaximumHeight(150)
        g4l.addWidget(self.ver_backups)
        b = QPushButton("↩️ A kijelölt régi verzió visszaállítása")
        b.clicked.connect(self.ver_restore_backup)
        g4l.addWidget(b)
        lay.addWidget(g4)

        g5 = QGroupBox("ℹ️ Új verzió kiadása (rendszergazdának)")
        g5l = QVBoxLayout(g5)
        h = QLabel("1. Az új <b>app_pyqt6CLOUD.py</b>-t WinSCP-vel töltsd fel a VPS-en ide: <b>/opt/raktarapp/kliens/</b> (felülírás).<br>"
                   "2. Kész – szerver-újraindítás nem kell. Minden gép legkésőbb 30 percen belül villogó piros sávval jelzi.<br>"
                   "3. Ha az új verzióhoz új Python csomag kell, a programban a <b>REQUIRED_PACKAGES</b> listába kell felvenni – "
                   "a gépek ebből tudják, mit kell telepíteni.<br>"
                   "Részletes leírás (Google Drive mentés beállítása is): vps/TELEPITES_verziok.md")
        h.setWordWrap(True)
        h.setTextFormat(Qt.TextFormat.RichText)
        g5l.addWidget(h)
        lay.addWidget(g5)
        lay.addStretch()

        self._upd = None              # {"info", "data", "text", "missing", "changes", "compile_err"}
        self._upd_check_busy = False
        self._pip_cmd = ""
        self.upd_signal.connect(self.on_upd_signal)
        self.upd_timer = QTimer(self)
        self.upd_timer.timeout.connect(self.check_app_update)
        self.upd_timer.start(30 * 60 * 1000)
        QTimer.singleShot(12000, self.check_app_update)
        self.ver_refresh_local()

    # ---- villogó sáv ----
    def _blink_banner(self):
        self._banner_blink = not self._banner_blink
        bg, fg = ("#d50000", "#ffffff") if self._banner_blink else ("#ffea00", "#000000")
        self.update_banner.setStyleSheet(f"background-color: {bg}; color: {fg}; font-size: 22px; font-weight: bold; "
                                         f"padding: 12px; border: 3px solid #ffffff; border-radius: 8px;")

    def _show_update_banner(self, text):
        self.update_banner.setText(text)
        if not self.update_banner.isVisible():
            self.update_banner.show()
            self._blink_banner()
            self.banner_timer.start(700)
        idx = self.tabs.indexOf(self.tab_versions)
        self.tabs.setTabText(idx, "🔴 VERZIÓK – ÚJ!")
        self.tabs.tabBar().setTabTextColor(idx, QColor("#ff1744"))
        self.lbl_version_warn.hide()

    def _hide_update_banner(self):
        self.banner_timer.stop()
        self.update_banner.hide()
        idx = self.tabs.indexOf(self.tab_versions)
        self.tabs.setTabText(idx, "🆕 VERZIÓK")
        self.tabs.tabBar().setTabTextColor(idx, QColor("#ffffff"))

    # ---- ellenőrzés ----
    def check_app_update(self, manual=False):
        if self._upd_check_busy:
            return
        self._upd_check_busy = True
        cached = self._upd

        def job():
            try:
                res = vps_request("GET", f"{VPS_URL}/app/version", timeout=10)
                if res.status_code == 404:
                    # a mi 404-ünk JSON (kod: S-404) = a szerver új, csak a fájl hiányzik; ha nem JSON, régi a server_api.py
                    try:
                        uj_szerver = res.json().get("kod") == "S-404"
                    except Exception:
                        uj_szerver = False
                    self.upd_signal.emit("nincs" if uj_szerver else "regi_szerver", {"manual": manual})
                    return
                if res.status_code != 200:
                    raise RuntimeError(f"A szerver hibát adott (HTTP {res.status_code})")
                info = res.json()
                uj = str(info.get("verzio") or "")
                if not uj or _ver_key(uj) <= _ver_key(APP_VERSION):
                    self.upd_signal.emit("naprakesz", {"manual": manual, "info": info})
                    return
                if cached and cached["info"].get("sha256") == info.get("sha256"):
                    cached["info"] = info
                    self.upd_signal.emit("uj", dict(cached, manual=manual))
                    return
                # az új verziót előre letöltjük és ellenőrizzük (így látszik, mi változott és milyen csomag kell)
                r2 = vps_request("GET", f"{VPS_URL}/app/download", timeout=90)
                if r2.status_code != 200:
                    raise RuntimeError(f"A letöltés nem sikerült (HTTP {r2.status_code})")
                data = r2.content
                if hashlib.sha256(data).hexdigest() != info.get("sha256"):
                    raise RuntimeError("A letöltött fájl sérült (az ellenőrzőkód nem egyezik)")
                text = data.decode("utf-8")
                compile_err = ""
                try:
                    compile(text, "app_pyqt6CLOUD.py", "exec")
                except SyntaxError as e:
                    compile_err = f"szintaktikai hiba a(z) {e.lineno}. sorban"
                pk = _parse_module_const(text, "REQUIRED_PACKAGES")
                pk = [tuple(x) for x in pk] if isinstance(pk, list) else REQUIRED_PACKAGES
                ch = _parse_module_const(text, "VALTOZASNAPLO") or []
                self.upd_signal.emit("uj", {"manual": manual, "info": info, "data": data, "pkgs": pk,
                                            "missing": missing_packages(pk), "changes": ch, "compile_err": compile_err})
            except Exception as e:
                self.upd_signal.emit("hiba", {"manual": manual, "message": str(e)})
            finally:
                self._upd_check_busy = False
        threading.Thread(target=job, daemon=True).start()

    def on_upd_signal(self, kind, d):
        manual = (d or {}).get("manual")
        if kind == "uj":
            first = not self._upd or self._upd["info"].get("sha256") != d["info"].get("sha256")
            self._upd = d
            uj = d["info"].get("verzio")
            self._show_update_banner(f"⬆️  ÚJ PROGRAMVERZIÓ ÉRHETŐ EL: {uj}   (ezen a gépen: {APP_VERSION})   —   KATTINTS IDE A FRISSÍTÉSHEZ!")
            self.ver_render()
            if first and load_config_value("frissites_jelezve") != d["info"].get("sha256"):
                save_config_value("frissites_jelezve", d["info"].get("sha256"))
                box = QMessageBox(self)
                box.setWindowTitle("⬆️ Új programverzió!")
                box.setIcon(QMessageBox.Icon.Information)
                box.setText(f"<h2>Új programverzió érhető el: {uj}</h2>Ezen a gépen most: {APP_VERSION}"
                            + ("<br><br><b style='color:#ff9800'>⚠️ Ehhez új Python csomag is kell – a VERZIÓK fül leírja, hogyan.</b>"
                               if d.get("missing") else ""))
                bnez = box.addButton("🆕 Megnézem (VERZIÓK fül)", QMessageBox.ButtonRole.AcceptRole)
                box.addButton("Később", QMessageBox.ButtonRole.RejectRole)
                box.exec()
                if box.clickedButton() is bnez:
                    self.tabs.setCurrentWidget(self.tab_versions)
        elif kind == "naprakesz":
            self._upd = None
            self._hide_update_banner()
            self._ver_server_info = d.get("info") or {}
            self.ver_render()
            if manual:
                self.trigger_toast(f"✅ Ez a legfrissebb program ({APP_VERSION})")
        elif kind == "nincs":
            self._ver_server_info = {}
            self.ver_render(extra="ℹ️ A VPS-en még nincs feltöltve program. Pontosan ide kell: /opt/raktarapp/kliens/app_pyqt6CLOUD.py "
                                  "(a mappa neve kliens, kisbetűvel; a fájl neve pontosan app_pyqt6CLOUD.py).")
            if manual:
                self.trigger_toast("ℹ️ A VPS-en még nincs feltöltve program a kliens mappába")
        elif kind == "regi_szerver":
            self._ver_server_info = {}
            self.ver_render(extra="⚠️ A VPS-en még a RÉGI server_api.py fut (nem ismeri a frissítést). Töltsd fel az újat az "
                                  "/opt/raktarapp/ mappába, és indítsd újra: sudo systemctl restart raktarapp")
            if manual:
                self.trigger_toast("⚠️ A VPS-en még a régi server_api.py fut – töltsd fel az újat és indítsd újra a szervert")
        elif kind == "hiba":
            self.ver_render(extra=f"❌ Az ellenőrzés nem sikerült: {d.get('message', '')}")
            if manual:
                self.trigger_toast(f"❌ Ellenőrzés nem sikerült: {d.get('message', '')}")
        elif kind == "pip_sor":
            self.ver_pip_out.appendPlainText(d["sor"])
        elif kind == "pip_kesz":
            self.btn_pip_install.setEnabled(True)
            self.ver_pip_out.appendPlainText("\n✅ KÉSZ – sikerült." if d["ok"] else f"\n❌ NEM SIKERÜLT (kilépési kód: {d['kod']}) – lásd fent, vagy kövesd a kézi lépéseket.")
            if self._upd:
                self._upd["missing"] = missing_packages(self._upd.get("pkgs"))
            self.ver_render()
        elif kind == "frissitve":
            self.append_process_log(f"⬆️ Program frissítve: {APP_VERSION} → {d['verzio']} (régi mentve: {d['mentes']})")
            log_event("-", "PROGRAM FRISSÍTÉS", f"{APP_VERSION} → {d['verzio']} (régi: {d['mentes']})", None, None, "VPS")
            QMessageBox.information(self, "✅ Frissítés kész", f"Az új program ({d['verzio']}) a helyére került.\n\n"
                                    f"A program most újraindul.\n\nA régi verzió mentve: {d['mentes']}")
            self._restart_app()
        elif kind == "frissites_hiba":
            self.btn_do_update.setEnabled(True)
            self.btn_do_update.setText("⬆️ FRISSÍTÉS MOST")
            QMessageBox.warning(self, "❌ Frissítés nem sikerült", f"{d.get('message', '')}\n\nA jelenlegi program változatlanul fut tovább.")

    # ---- megjelenítés ----
    def ver_refresh_local(self):
        self.ver_backups.clear()
        bdir = os.path.join(BASE_DIR, "regi_verziok")
        try:
            files = sorted((f for f in os.listdir(bdir) if f.endswith(".py")),
                           key=lambda f: os.path.getmtime(os.path.join(bdir, f)), reverse=True)
        except Exception:
            files = []
        for f in files:
            ts = datetime.datetime.fromtimestamp(os.path.getmtime(os.path.join(bdir, f))).strftime("%Y-%m-%d %H:%M")
            self.ver_backups.addItem(f"{f}   (mentve: {ts})")
        if not files:
            self.ver_backups.addItem("(még nincs – az első frissítéskor ide kerül a régi verzió)")
        self.ver_render()

    def ver_render(self, extra=""):
        d = self._upd
        srv = (d or {}).get("info") or getattr(self, "_ver_server_info", {}) or {}
        if d:
            uj = d["info"].get("verzio")
            if d.get("compile_err"):
                self.ver_status.setText(f"❌ A VPS-en lévő új verzió ({uj}) HIBÁS ({d['compile_err']}) – NE frissíts, szólj a rendszergazdának.")
                self.ver_status.setStyleSheet("font-size: 20px; font-weight: bold; padding: 14px; border-radius: 8px; background: #b71c1c; color: white;")
                self.btn_do_update.setEnabled(False)
            elif d.get("missing"):
                self.ver_status.setText(f"⚠️ ÚJ VERZIÓ: {uj} – de előbb {len(d['missing'])} Python csomagot telepíteni kell (lásd lent, lépésről lépésre)!")
                self.ver_status.setStyleSheet("font-size: 20px; font-weight: bold; padding: 14px; border-radius: 8px; background: #ff9800; color: black;")
                self.btn_do_update.setEnabled(False)
            else:
                self.ver_status.setText(f"⬆️ ÚJ VERZIÓ ÉRHETŐ EL: {uj} – nyomd meg a zöld FRISSÍTÉS MOST gombot!")
                self.ver_status.setStyleSheet("font-size: 20px; font-weight: bold; padding: 14px; border-radius: 8px; background: #d50000; color: white;")
                self.btn_do_update.setEnabled(True)
        else:
            self.ver_status.setText(extra or f"✅ A legfrissebb verzió fut ezen a gépen ({APP_VERSION}).")
            self.ver_status.setStyleSheet("font-size: 20px; font-weight: bold; padding: 14px; border-radius: 8px; background: %s; color: %s;"
                                          % (("#37474f", "white") if extra else ("#1b5e20", "white")))
            self.btn_do_update.setEnabled(False)
        self.ver_info.setText(
            f"🖥️ Ezen a gépen: <b>{APP_VERSION}</b> &nbsp;·&nbsp; gép neve: <b>{GEP_NEV}</b><br>"
            f"☁️ A VPS-en: <b>{srv.get('verzio') or '–'}</b>" + (f" (feltöltve: {srv.get('feltoltve')})" if srv.get("feltoltve") else "") + "<br>"
            f"🐍 Python: {sys.version.split()[0]} – {python_console_exe()}<br>"
            f"📁 A program helye: {os.path.abspath(__file__)}")

        # címkenyomtató
        if hasattr(self, "ver_printer_info"):
            lp = local_label_printer()
            if is_print_host():
                self.ver_printer_info.setText(f"✅ <b>Ez a nyomtatós gép</b> – itt nyomtat: <b>{lp or get_zebra_printer_name() or '?'}</b>. "
                                              f"A többi gépről küldött címkéket is ez nyomtatja ki (pár másodpercen belül).")
            elif printer_mode() == "helyi":
                self.ver_printer_info.setText("❌ „Ehhez a géphez van kötve” van beállítva, de a <b>pywin32</b> csomag nincs telepítve, "
                                              "így nem tud nyomtatni. Telepítsd (fekete ablakban): <code>"
                                              f"\"{python_console_exe()}\" -m pip install pywin32</code>, majd indítsd újra a programot.")
            else:
                self.ver_printer_info.setText("📨 <b>Ehhez a géphez nincs címkenyomtató kötve</b> – a Nyomtatás gomb a nyomtatós gépre "
                                              "küldi a címkét, és az kinyomtatja (a nyomtatós gépen futnia kell a programnak).")

        # csomagok: a jelenlegi + (ha van) az új verzió listája
        pkgs = list(REQUIRED_PACKAGES)
        if d and d.get("pkgs"):
            names = {x[1] for x in pkgs}
            pkgs += [x for x in d["pkgs"] if x[1] not in names]
        missing = {x[1] for x in missing_packages(pkgs)}
        uj_nevek = {x[1] for x in (d or {}).get("pkgs", [])} - {x[1] for x in REQUIRED_PACKAGES}
        t = self.ver_pkgs
        t.setRowCount(len(pkgs))
        for i, (imp, pip_name, mire, csak_win) in enumerate(pkgs):
            if csak_win and os.name != "nt":
                st, col = "– (csak Windows)", "#888"
            elif pip_name in missing:
                st, col = "❌ HIÁNYZIK", "#ff5252"
            else:
                st, col = "✅ telepítve", "#00e676"
            for c, v in enumerate([pip_name + ("  🆕 (új verzióhoz)" if pip_name in uj_nevek else ""), mire, st]):
                it = QTableWidgetItem(v)
                if c == 2:
                    it.setForeground(QColor(col))
                t.setItem(i, c, it)
        miss = [p for p in pkgs if p[1] in missing]
        exe = python_console_exe()
        self._pip_cmd = f'"{exe}" -m pip install ' + " ".join(p[1] for p in miss) if miss else ""
        self.btn_pip_copy.setVisible(bool(miss))
        self.btn_pip_install.setVisible(bool(miss))
        if miss:
            self.ver_pip_help.setText(
                f"<div style='background:#4e342e; padding:10px; border-radius:6px'>"
                f"<b style='font-size:17px; color:#ffb74d'>⚠️ {len(miss)} csomag hiányzik: {', '.join(p[1] for p in miss)}</b><br><br>"
                f"<b>A) Automatikusan:</b> nyomd meg lent a narancs <b>⚙️ Telepítés most</b> gombot, és várd meg, amíg kiírja: <b>✅ KÉSZ</b> (1-3 perc).<br><br>"
                f"<b>B) Kézzel, lépésről lépésre</b> (ha az automatikus nem sikerült):<br>"
                f"<b>1.</b> Nyomd meg a billentyűzeten a <b>Windows gomb + R</b> billentyűt. Felugrik a „Futtatás” ablak.<br>"
                f"<b>2.</b> Írd be: <b>cmd</b> és nyomj <b>Enter</b>-t. Megnyílik egy fekete ablak.<br>"
                f"<b>3.</b> Itt nyomd meg a <b>📋 Parancs másolása</b> gombot, majd a fekete ablakban <b>jobb egérgomb</b> (= beillesztés), és <b>Enter</b>.<br>"
                f"&nbsp;&nbsp;&nbsp;&nbsp;A parancs: <code style='background:#000; color:#00e676'>{self._pip_cmd}</code><br>"
                f"<b>4.</b> Várd meg, amíg a végén ezt írja: <b>Successfully installed ...</b> (1-3 perc, közben sok sor fut le – ez normális).<br>"
                f"<b>5.</b> Ha ezt írja: <i>Access is denied / Permission denied</i> → zárd be a fekete ablakot, Start menü → írd be: <b>cmd</b> → "
                f"jobb klikk a „Parancssor”-on → <b>Futtatás rendszergazdaként</b>, és ismételd meg a 3. lépést.<br>"
                f"&nbsp;&nbsp;&nbsp;&nbsp;Ha ezt írja: <i>No module named pip</i> → előbb ezt futtasd: <code style='background:#000; color:#00e676'>\"{exe}\" -m ensurepip</code>, utána a 3. lépést.<br>"
                f"<b>6.</b> Ha kész, itt nyomd meg az <b>🔄 Ellenőrzés most</b> gombot: a csomag mellett <b>✅ telepítve</b> lesz.<br>"
                f"<b>7.</b> Utána nyomd meg a zöld <b>⬆️ FRISSÍTÉS MOST</b> gombot.</div>")
        else:
            self.ver_pip_help.setText("✅ Minden szükséges Python csomag telepítve van ezen a gépen.")

        # változások
        ch = (d or {}).get("changes") or VALTOZASNAPLO
        html = []
        for ver, datum, pontok in ch:
            uj_e = _ver_key(ver) > _ver_key(APP_VERSION)
            szin = "#ff1744" if uj_e else "#29b6f6"
            jel = (" <b style='color:#ff1744'>← ÚJ, még nincs ezen a gépen</b>" if uj_e else
                   " <b style='color:#00e676'>← ez fut most</b>" if ver == APP_VERSION else "")
            html.append(f"<p><b style='color:{szin}; font-size:15px'>{ver}</b> "
                        f"<span style='color:#999'>({datum})</span>{jel}<br>"
                        + "<br>".join(f"&nbsp;&nbsp;• {x}" for x in pontok) + "</p>")
        self.ver_changes.setText("".join(html[:12]))

        # Drive / VPS mentések
        arch = srv.get("archivum") or []
        if srv.get("verzio") is None and not arch:
            self.ver_drive.setText("")
        else:
            sor = []
            for a in arch[:8]:
                st = {"ok": "✅ Google Drive-on" + (f" (<a href='{a.get('drive_url')}'>megnyitás</a>)" if a.get("drive_url") else ""),
                      "var": "⏳ Drive-ra mentés folyamatban", "nincs beallitva": "⚪ Drive mentés nincs beállítva",
                      "hiba": f"❌ Drive hiba: {a.get('hiba', '')}"}.get(a.get("drive"), a.get("drive", ""))
                sor.append(f"• <b>{a.get('verzio')}</b> ({a.get('ido')}, {a.get('forras', '')}) – {st}")
            self.ver_drive.setText("☁️ <b>A VPS-en archivált verziók</b> (/opt/raktarapp/kliens/archivum/):<br>" + ("<br>".join(sor) or "–") +
                                   ("" if srv.get("drive_beallitva", True) else
                                    "<br><span style='color:#ffb74d'>A Google Drive mentés még nincs beállítva – lásd vps/TELEPITES_verziok.md (3. pont).</span>"))

    def ver_rename_machine(self):
        global GEP_NEV
        uj, ok = QInputDialog.getText(self, "Gép neve", "Ennek a gépnek a neve (a többi gépen, a PDA-n és a Gépek listában ez látszik).\n"
                                      "Mindegyik gépnek MÁS nevet adj, pl. Raktár gép / Iroda gép:", text=GEP_NEV)
        uj = (uj or "").strip()
        if not ok or not uj or uj == GEP_NEV:
            return
        if QMessageBox.question(self, "Gép neve", f"Átírod a gép nevét?\n\n{GEP_NEV}  →  {uj}") != QMessageBox.StandardButton.Yes:
            return
        GEP_NEV = uj
        save_config_value("gep_nev", uj)
        self.ver_render()
        self.trigger_toast(f"✅ A gép neve mostantól: {uj}")

    def ver_set_printer_mode(self, *_):
        m = self.ver_printer_mode.currentData()
        save_config_value("nyomtato_mod", m)
        local_label_printer(force=True)
        self.ver_render()
        self.trigger_toast("🖨️ Címkenyomtató beállítás mentve")

    def ver_copy_pip(self):
        if self._pip_cmd:
            QApplication.clipboard().setText(self._pip_cmd)
            self.trigger_toast("📋 Parancs kimásolva – a fekete ablakban jobb egérgomb = beillesztés")

    def ver_pip_install(self):
        if not self._pip_cmd:
            return
        pkgs = self._pip_cmd.split(" install ", 1)[1].split()
        if QMessageBox.question(self, "Csomagok telepítése?", f"Telepítsem ezeket a Python csomagokat?\n\n{', '.join(pkgs)}\n\n"
                                "Internet kell hozzá, 1-3 percig tarthat. A program közben használható.") != QMessageBox.StandardButton.Yes:
            return
        self.btn_pip_install.setEnabled(False)
        self.ver_pip_out.clear()
        self.ver_pip_out.appendPlainText(f"> {self._pip_cmd}\n")

        def job():
            import subprocess
            try:
                flags = 0x08000000 if os.name == "nt" else 0   # CREATE_NO_WINDOW
                pr = subprocess.Popen([python_console_exe(), "-m", "pip", "install", *pkgs], stdout=subprocess.PIPE,
                                      stderr=subprocess.STDOUT, text=True, encoding="utf-8", errors="replace", creationflags=flags)
                for line in pr.stdout:
                    self.upd_signal.emit("pip_sor", {"sor": line.rstrip()})
                kod = pr.wait()
                self.upd_signal.emit("pip_kesz", {"ok": kod == 0, "kod": kod})
            except Exception as e:
                self.upd_signal.emit("pip_sor", {"sor": f"HIBA: {e}"})
                self.upd_signal.emit("pip_kesz", {"ok": False, "kod": -1})
        threading.Thread(target=job, daemon=True).start()

    # ---- frissítés ----
    def do_app_update(self):
        d = self._upd
        if not d or not d.get("data"):
            return
        if getattr(sys, "frozen", False):
            QMessageBox.information(self, "Frissítés", "Ez a program .exe-ként fut, így nem tudja magát frissíteni.")
            return
        miss = missing_packages(d.get("pkgs"))
        if miss:
            QMessageBox.warning(self, "Előbb a csomagok!", "Az új verzióhoz előbb ezeket kell telepíteni:\n\n" +
                                "\n".join(p[1] for p in miss) + "\n\nA VERZIÓK fülön lépésről lépésre le van írva.")
            return
        if QMessageBox.question(self, "Program frissítése?",
                                f"Frissítés erre: {d['info'].get('verzio')}\n(ezen a gépen most: {APP_VERSION})\n\n"
                                f"A program utána magától újraindul. A még nem mentett sorok előtte elmentődnek,\n"
                                f"a régi verzió megmarad (regi_verziok mappa + VPS + Google Drive).\n\nMehet?"
                                ) != QMessageBox.StandardButton.Yes:
            return
        try:
            self.process_auto_row_sync()
        except Exception:
            pass
        self.btn_do_update.setEnabled(False)
        self.btn_do_update.setText("⏳ Frissítés...")
        data, uj = d["data"], d["info"].get("verzio")

        def job():
            try:
                target = os.path.abspath(__file__)
                with open(target, "rb") as f:
                    regi = f.read()
                # a régi verzió a VPS-re is felmegy (onnan a Google Drive-ra) – ha nem sikerül, attól még frissítünk
                try:
                    import base64
                    vps_request("POST", f"{VPS_URL}/app/archive", json={"tartalom_b64": base64.b64encode(regi).decode("ascii"),
                                                                        "gep": GEP_NEV}, timeout=30)
                except Exception:
                    pass
                backup_dir = os.path.join(BASE_DIR, "regi_verziok")
                os.makedirs(backup_dir, exist_ok=True)
                mentes = os.path.join(backup_dir, f"app_pyqt6CLOUD_{APP_VERSION}.py")
                with open(mentes, "wb") as f:
                    f.write(regi)
                tmp = target + ".uj"
                with open(tmp, "wb") as f:
                    f.write(data)
                os.replace(tmp, target)
                self.upd_signal.emit("frissitve", {"verzio": uj, "mentes": mentes})
            except Exception as e:
                self.upd_signal.emit("frissites_hiba", {"message": str(e)})
        QTimer.singleShot(1500, lambda: threading.Thread(target=job, daemon=True).start())

    def ver_restore_backup(self):
        it = self.ver_backups.currentItem()
        if not it or not it.text().endswith(")") or it.text().startswith("("):
            self.trigger_toast("Előbb kattints egy régi verzióra a listában.")
            return
        fname = it.text().split("   (")[0]
        src = os.path.join(BASE_DIR, "regi_verziok", fname)
        if QMessageBox.question(self, "Régi verzió visszaállítása?",
                                f"Visszaállítod ezt a verziót?\n\n{fname}\n\nA mostani ({APP_VERSION}) is elmentődik a regi_verziok mappába.\n"
                                f"A program utána újraindul.\n\n(Ha a VPS-en újabb verzió van, a villogó sáv újra megjelenik – ez normális.)"
                                ) != QMessageBox.StandardButton.Yes:
            return
        try:
            with open(src, "rb") as f:
                data = f.read()
            compile(data.decode("utf-8"), fname, "exec")
            target = os.path.abspath(__file__)
            shutil.copy2(target, os.path.join(BASE_DIR, "regi_verziok", f"app_pyqt6CLOUD_{APP_VERSION}.py"))
            tmp = target + ".uj"
            with open(tmp, "wb") as f:
                f.write(data)
            os.replace(tmp, target)
        except Exception as e:
            QMessageBox.warning(self, "Visszaállítás", f"Nem sikerült: {e}")
            return
        log_event("-", "PROGRAM VISSZAÁLLÍTÁS", f"{APP_VERSION} → {fname}", None, None, "Gép")
        QMessageBox.information(self, "Visszaállítás", "Kész. A program most újraindul.")
        self._restart_app()

    def _restart_app(self):
        """Az új program elindítása ugyanúgy, ahogy ez el lett indítva, majd ennek a bezárása."""
        import subprocess
        try:
            subprocess.Popen([sys.executable] + sys.argv, cwd=os.getcwd())
        except Exception as e:
            QMessageBox.warning(self, "Újraindítás", f"Nem sikerült automatikusan újraindítani ({e}).\nIndítsd el kézzel a programot.")
        QApplication.instance().quit()

    def admin_load_clients(self):
        if hasattr(self, "adm_cli_info"):
            self.adm_cli_info.setText("⏳ Betöltés...")

        def job():
            try:
                res = vps_request("GET", f"{VPS_URL}/clients", timeout=10)
                data = res.json() if res.status_code == 200 else {"hiba": f"HTTP {res.status_code}"}
            except Exception as e:
                data = {"hiba": str(e)}
            self.clients_signal.emit(data)
        threading.Thread(target=job, daemon=True).start()

    def fill_clients_table(self, data):
        if not hasattr(self, "adm_cli"):
            return
        if data.get("hiba"):
            self.adm_cli_info.setText(f"❌ Nem sikerült lekérni a szerverről: {data['hiba']}")
            return
        leg = data.get("legfrissebb") or {}
        rows = data.get("kliensek") or []
        self.adm_cli_info.setText(
            f"Legfrissebb verziók → 🖥️ gépi program: {leg.get('gep') or '-'}  ·  📱 PDA felület: {leg.get('pda') or '-'}  ·  "
            f"📱 PDA app: {leg.get('pda_app') or '-'}     |  Ez a gép: {GEP_NEV} – {APP_VERSION}")
        t = self.adm_cli
        t.setRowCount(0)
        hostname = GEP_NEV
        for r in rows:
            i = t.rowCount(); t.insertRow(i)
            tipus = "🖥️ Gép" if r.get("tipus") == "gep" else "📱 PDA"
            ver = r.get("verzio") or "-"
            app_v = r.get("app_verzio") or ""
            latest = leg.get("gep") if r.get("tipus") == "gep" else leg.get("pda")
            regi = bool(latest and ver and _ver_key(ver) < _ver_key(latest))
            mp = int(r.get("mp") or 0)
            if r.get("online"):
                allapot = "🟢 Online"
            elif mp < 3600:
                allapot = f"⚪ {mp // 60} perce"
            elif mp < 86400 * 2:
                allapot = f"⚪ {mp // 3600} órája"
            else:
                allapot = f"⚪ {mp // 86400} napja"
            nev = r.get("nev", "") + ("  (ez a gép)" if r.get("tipus") == "gep" and r.get("nev") == hostname else "")
            vals = [tipus, nev, ver + ("  ⚠️ RÉGI" if regi else "  ✔ legfrissebb" if latest and ver == latest else ""),
                    app_v or ("-" if r.get("tipus") == "gep" else "régi app"), allapot, r.get("utolso", ""),
                    r.get("felhasznalo", ""), r.get("ip", "")]
            for c, v in enumerate(vals):
                it = QTableWidgetItem(str(v))
                if c == 2 and regi:
                    it.setForeground(QColor("#ff9800"))
                if c == 4 and r.get("online"):
                    it.setForeground(QColor("#00e676"))
                t.setItem(i, c, it)
        if hasattr(self, "ver_cli"):
            v = self.ver_cli
            v.setRowCount(t.rowCount())
            for i in range(t.rowCount()):
                for c in range(t.columnCount()):
                    src = t.item(i, c)
                    if src:
                        v.setItem(i, c, src.clone())

    def _admin_show(self, rows, r, lbl, tbl):
        if 0 <= r < len(rows):
            self.render_log_detail(rows[r], lbl, tbl)

    def _admin_fill_events(self, table, rows, cols):
        table.setUpdatesEnabled(False)
        table.setRowCount(len(rows))
        for ri, r in enumerate(rows):
            elotte = json.loads(r["elotte"]) if r.get("elotte") else None
            utana = json.loads(r["utana"]) if r.get("utana") else None
            prod = utana or elotte or {}
            esemeny = safe_str(r.get("esemeny"))
            valt = safe_str(r.get("reszletek"))
            if elotte and utana and not esemeny.upper().startswith("ELADÁS"):
                valt = describe_changes(elotte, utana) or valt
            data = {"idopont": safe_str(r.get("idopont")), "esemeny": esemeny, "forras": safe_str(r.get("forras")),
                    "kod": safe_str(r.get("r18_kod")), "C": prod.get("C", ""), "D": prod.get("D", ""), "E": prod.get("E", ""), "valt": valt}
            for ci, key in enumerate(cols):
                it = QTableWidgetItem(data[key])
                if key == "esemeny":
                    it.setForeground(QColor(self._log_event_color(esemeny)))
                    expl = explain_event(esemeny)
                    if expl:
                        it.setToolTip("💡 " + expl)
                if key in ("valt", "forras"):
                    it.setToolTip(data[key])
                table.setItem(ri, ci, it)
        table.setUpdatesEnabled(True)

    def _admin_query_events(self):
        sql = "SELECT * FROM naplo WHERE idopont >= ? AND idopont <= ?"
        params = [self.adm_from.date().toString("yyyy-MM-dd") + " 00:00:00", self.adm_to.date().toString("yyyy-MM-dd") + " 23:59:59"]
        cond = self.LOG_FILTERS[self.adm_type.currentIndex()][1]
        if cond:
            sql += f" AND ({cond})"
        q = self.adm_search.text().strip()
        if q:
            sql += " AND (r18_kod LIKE ? OR esemeny LIKE ? OR reszletek LIKE ? OR elotte LIKE ? OR utana LIKE ? OR forras LIKE ?)"
            params += [f"%{q}%"] * 6
        sql += " ORDER BY id DESC LIMIT 20000"
        conn = get_db_connection()
        cur = conn.cursor()
        cur.execute(sql, params)
        rows = [dict(r) for r in cur.fetchall()]
        conn.close()
        return rows

    def admin_load_all(self):
        try:
            self._adm_all_rows = self._admin_query_events()
        except Exception as e:
            self.adm_all_count.setText(f"❌ Hiba: {e}")
            return
        self._admin_fill_events(self.adm_all, self._adm_all_rows, ["idopont", "esemeny", "forras", "kod", "C", "D", "E", "valt"])
        self.adm_all_count.setText(f"{len(self._adm_all_rows)} esemény (legfeljebb 20 000 a kiválasztott időszakban)")

    def admin_export_all(self):
        import csv
        rows = self._admin_query_events()
        path, _ = QFileDialog.getSaveFileName(self, "Változások exportálása",
                                              f"raktar_valtozasok_{datetime.date.today().isoformat()}.csv", "CSV (*.csv)")
        if not path:
            return
        with open(path, "w", newline="", encoding="utf-8-sig") as f:   # utf-8-sig: az Excel helyesen nyitja meg
            w = csv.writer(f, delimiter=";")
            w.writerow(["Időpont", "Esemény", "Honnan / hogyan", "R18-kód", "Részletek"]
                       + [f"ELŐTTE {c}" for c in NAPLO_COLS] + [f"UTÁNA {c}" for c in NAPLO_COLS])
            for r in rows:
                b = json.loads(r["elotte"]) if r.get("elotte") else {}
                a = json.loads(r["utana"]) if r.get("utana") else {}
                w.writerow([safe_str(r.get("idopont")), safe_str(r.get("esemeny")), safe_str(r.get("forras")),
                            safe_str(r.get("r18_kod")), safe_str(r.get("reszletek"))]
                           + [b.get(c, "") for c in NAPLO_COLS] + [a.get(c, "") for c in NAPLO_COLS])
        self.trigger_toast(f"📤 {len(rows)} esemény exportálva: {os.path.basename(path)}")

    def admin_load_deleted(self):
        """Minden törölt / kiürített termék teljes adata (ToroltSor archívum + napló törlés-bejegyzések)."""
        rows, seen = [], set()
        try:
            conn = get_db_connection()
            cur = conn.cursor()
            cur.execute("SELECT idopont, r18_kod, forras, payload FROM torolt_archiv ORDER BY id DESC")
            for r in cur.fetchall():
                d = json.loads(r["payload"]) if r["payload"] else {}
                sig = json.dumps(d, sort_keys=True, ensure_ascii=False)
                if d and sig not in seen:
                    seen.add(sig); rows.append((safe_str(r["idopont"]), safe_str(r["forras"]), d))
            cur.execute("SELECT idopont, esemeny, forras, elotte FROM naplo WHERE elotte IS NOT NULL AND "
                        "(esemeny LIKE 'TÖRLÉS%' OR esemeny LIKE 'DUPLIKÁCIÓ%') ORDER BY id DESC")
            for r in cur.fetchall():
                d = json.loads(r["elotte"])
                sig = json.dumps(d, sort_keys=True, ensure_ascii=False)
                if d and _row_has_product_data(d) and sig not in seen:
                    seen.add(sig); rows.append((safe_str(r["idopont"]), f"{safe_str(r['esemeny'])} | {safe_str(r['forras'])}", d))
            conn.close()
        except Exception as e:
            self.adm_del_count.setText(f"❌ Hiba: {e}")
            return
        rows.sort(key=lambda x: x[0], reverse=True)
        q = self.adm_del_search.text().strip().lower()
        if q:
            rows = [r for r in rows if q in (" ".join(r[2].values()) + " " + r[1]).lower()]
        self._adm_del_rows = rows
        t = self.adm_del
        t.setUpdatesEnabled(False)
        t.setRowCount(len(rows))
        for ri, (ido, forras, d) in enumerate(rows):
            for ci, v in enumerate([ido, forras] + [d.get(c, "") for c in NAPLO_COLS]):
                it = QTableWidgetItem(str(v))
                if ci == 1:
                    it.setToolTip(str(v))
                t.setItem(ri, ci, it)
        t.setUpdatesEnabled(True)
        self.adm_del_count.setText(f"{len(rows)} törölt / kiürített termék (teljes adattal). Kijelölés + ♻️ = visszaállítás.")

    def admin_restore_deleted(self):
        sel = sorted({i.row() for i in self.adm_del.selectedIndexes()})
        if not sel:
            self.trigger_toast("⚠️ Jelölj ki egy törölt terméket!")
            return
        ri = sel[0]
        _, _, d = self._adm_del_rows[ri]
        code = norm_code(d.get("A"))
        now = get_row_snapshot(code)
        msg = (f"Visszaállítod a(z) {code} terméket ezekkel az adatokkal?\n\n{format_row_for_log(d)}")
        if now and _row_has_product_data(now):
            msg += f"\n\n⚠️ FIGYELEM: ezen az R18-számon most ez van, ami FELÜLÍRÓDIK:\n{format_row_for_log(now)}"
        if SilentConfirmDialog("♻️ Visszaállítás", msg, self).exec() != QDialog.DialogCode.Accepted:
            return
        data = {c: str(d.get(c, "")) for c in NAPLO_COLS}
        data["A"] = code
        try:
            conn = get_db_connection()
            cur = conn.cursor()
            cur.execute(f"UPDATE termekek SET {', '.join(c + '=?' for c in NAPLO_COLS)} WHERE UPPER(TRIM(A))=? OR UPPER(TRIM(B))=?",
                        (*[data[c] for c in NAPLO_COLS], code, code))
            if cur.rowcount == 0:
                cur.execute(f"INSERT INTO termekek ({', '.join(NAPLO_COLS)}, szin) VALUES ({', '.join('?' * 16)}, 'FEHÉR')",
                            [data[c] for c in NAPLO_COLS])
            conn.commit()
            conn.close()
        except Exception as e:
            self.trigger_toast(f"❌ Visszaállítási hiba: {e}")
            return
        payload = dict(data); payload["_force"] = True
        enqueue_outbox(payload)
        flush_outbox_async()
        log_event(code, "FELVÉTEL (törölt termék visszaállítva)", "", now, get_row_snapshot(code),
                  "🛡️ Admin fül - ♻️ Kijelölt visszaállítása gomb")
        self.trigger_toast(f"♻️ {code} visszaállítva (helyi + szerver + Google Sheets)")
        if hasattr(self, "silent_update_excel_table"):
            self.silent_update_excel_table()

    def admin_load_history(self):
        code = norm_code(fix_scanner_text(self.adm_hist_code.text()))
        if not code:
            return
        self.adm_hist_code.selectAll()
        conn = get_db_connection()
        cur = conn.cursor()
        cur.execute("SELECT * FROM naplo WHERE UPPER(TRIM(r18_kod)) = ? ORDER BY id DESC", (code,))
        rows = [dict(r) for r in cur.fetchall()]
        cur.execute("SELECT datum, platform, vevo, ar, darab, tracking FROM eladasok WHERE UPPER(TRIM(A)) = ? AND torolve = 0", (code,))
        sales = cur.fetchall()
        conn.close()
        now = get_row_snapshot(code)
        import html as _html
        txt = f"<b>{_html.escape(code)}</b> - jelenleg: "
        txt += (_html.escape(format_row_for_log(now)) if now and _row_has_product_data(now)
                else "<span style='color:#ff5252;'>nincs (törölve / üres)</span>")
        if sales:
            txt += "<br>💰 Eladások: " + " | ".join(
                _html.escape(f"{s['datum']} {SALES_SHEETS.get(s['platform'], {}).get('cim', s['platform'])} {s['darab']} db, "
                             f"vevő: {s['vevo'] or '-'}, ár: {s['ar'] or '-'}") for s in sales)
        txt += f"<br>📜 {len(rows)} esemény a naplóban"
        self.adm_hist_now.setText(txt)
        self._adm_hist_rows = rows
        self._admin_fill_events(self.adm_hist, rows, ["idopont", "esemeny", "forras", "valt"])
        if rows:
            self.adm_hist.selectRow(0)

    def admin_load_outbox(self):
        conn = get_db_connection()
        cur = conn.cursor()
        cur.execute("SELECT * FROM sync_outbox ORDER BY id")
        rows = [dict(r) for r in cur.fetchall()]
        cur.execute("SELECT COUNT(*) FROM eladasok WHERE sheets_synced = 0")
        sales_pending = cur.fetchone()[0]
        conn.close()
        t = self.adm_out
        t.setRowCount(len(rows))
        op_names = {"upsert": "Mentés", "delete": "Törlés", "color": "Szín", "archive": "ToroltSor"}
        for ri, r in enumerate(rows):
            try:
                pl = json.loads(r["payload"])
            except Exception:
                pl = {}
            vals = [safe_str(r.get("created_at")), op_names.get(pl.get("_op", "upsert"), pl.get("_op", "")),
                    safe_str(pl.get("A")), "✅" if r.get("vps_synced") else "⏳", "✅" if r.get("sheets_synced") else "⏳",
                    str(r.get("attempts") or 0), safe_str(r.get("last_attempt")),
                    format_row_for_log({k: v for k, v in pl.items() if not str(k).startswith("_")})]
            for ci, v in enumerate(vals):
                t.setItem(ri, ci, QTableWidgetItem(v))
        self.adm_out_count.setText(f"{len(rows)} tétel vár feltöltésre (VPS / Google Sheets)  |  💰 {sales_pending} eladás vár az "
                                   f"eladási munkalapokra. Ha sokáig itt marad valami: nincs internet vagy hiba van - "
                                   f"lásd ⚙️ Folyamat napló.")

    # -------------------------------------------------------------------
    # TAB 8: FOLYAMAT ÉS SZINKRON NAPLÓ TAB
    # -------------------------------------------------------------------
    def init_process_log_tab(self):
        layout = QVBoxLayout(self.tab_process_log)
        layout.setContentsMargins(10, 10, 10, 10)

        top_bar = QHBoxLayout()
        lbl_title = QLabel("⚙️ TECHNIKAI FOLYAMAT NAPLÓ: szinkron, VPS, Google Sheets, hibák")
        lbl_title.setStyleSheet("font-size: 16px; font-weight: bold; color: #ffffff;")
        top_bar.addWidget(lbl_title)
        top_bar.addStretch()

        top_bar.addWidget(QLabel("Nap:"))
        self.log_date_edit = QDateEdit()
        self.log_date_edit.setCalendarPopup(True)
        self.log_date_edit.setDate(QDate.currentDate())
        self.log_date_edit.setStyleSheet("background-color: #2b2b2b; color: #fff; padding: 4px; font-weight: bold;")
        top_bar.addWidget(self.log_date_edit)

        btn_fetch_changes = QPushButton("📂 Nap betöltése")
        btn_fetch_changes.setToolTip("A kiválasztott nap technikai naplóját tölti be (naplok/folyamat_ÉÉÉÉ-HH-NN.log)")
        btn_fetch_changes.setStyleSheet("background-color: #29b6f6; color: #000; font-weight: bold; padding: 6px 12px;")
        btn_fetch_changes.clicked.connect(self.query_history_logs_by_date)
        top_bar.addWidget(btn_fetch_changes)

        btn_live = QPushButton("🟢 Élő (mai) nézet")
        btn_live.setStyleSheet("background-color: #00e676; color: #000; font-weight: bold; padding: 6px 12px;")
        btn_live.clicked.connect(self.show_live_process_log)
        top_bar.addWidget(btn_live)

        btn_clear_view = QPushButton("🗑️ Nézet Törlése")
        btn_clear_view.setStyleSheet("background-color: #ef5350; color: #fff; font-weight: bold; padding: 6px 12px;")
        btn_clear_view.clicked.connect(
            lambda: self.run_button_action(
                btn_clear_view, 
                "🗑️ Nézet Törlése", 
                self.clear_process_log
            )
        )
        top_bar.addWidget(btn_clear_view)

        layout.addLayout(top_bar)

        self.process_log_text = QTextEdit()
        self.process_log_text.setReadOnly(True)
        self.process_log_text.setStyleSheet("""
            QTextEdit {
                background-color: #0d0d0d;
                color: #00e676;
                font-family: 'Consolas', 'Courier New', monospace;
                font-size: 14px;
                border: 1px solid #333;
                border-radius: 6px;
                padding: 10px;
            }
        """)
        layout.addWidget(self.process_log_text)

        # Régi napi fájlok takarítása
        try:
            limit = datetime.date.today() - datetime.timedelta(days=PROCESS_LOG_KEEP_DAYS)
            for fn in os.listdir(PROCESS_LOG_DIR) if os.path.isdir(PROCESS_LOG_DIR) else []:
                m = re.fullmatch(r"folyamat_(\d{4}-\d{2}-\d{2})\.log", fn)
                if m and datetime.date.fromisoformat(m.group(1)) < limit:
                    os.remove(os.path.join(PROCESS_LOG_DIR, fn))
        except Exception:
            pass

        self._process_log_live = True
        self.show_live_process_log()
        self.append_process_log("🚀 Rendszer elindítva. (Itt csak technikai üzenetek vannak - a raktári események a 📜 Napló fülön.)")

    def _load_process_log_day(self, day_str):
        path = os.path.join(PROCESS_LOG_DIR, f"folyamat_{day_str}.log")
        if not os.path.exists(path):
            return None
        lines = []
        with open(path, encoding="utf-8", errors="replace") as f:
            for line in f.read().splitlines()[-3000:]:
                parts = line.split("\t", 2)
                if len(parts) == 3:
                    lines.append(self._process_log_line_html(*parts))
        return lines

    def show_live_process_log(self):
        """A mai nap technikai naplója, és innentől élőben jönnek az új sorok."""
        self._process_log_live = True
        self.log_date_edit.setDate(QDate.currentDate())
        self.process_log_text.clear()
        lines = self._load_process_log_day(datetime.date.today().isoformat()) or []
        if lines:
            self.process_log_text.setHtml("<br>".join(lines))
            self.process_log_text.moveCursor(self.process_log_text.textCursor().MoveOperation.End)

    def query_history_logs_by_date(self):
        """📂 Egy korábbi nap TECHNIKAI naplójának betöltése (nem a raktári események - azok a Napló fülön)."""
        day = self.log_date_edit.date().toString("yyyy-MM-dd")
        if day == datetime.date.today().isoformat():
            self.show_live_process_log()
            return
        self._process_log_live = False
        self.process_log_text.clear()
        lines = self._load_process_log_day(day)
        if not lines:
            self.process_log_text.setHtml(f"<span style='color:#888;'>Nincs technikai napló erre a napra: {day} "
                                          f"(legfeljebb {PROCESS_LOG_KEEP_DAYS} napig őrződnek meg).</span>")
            return
        self.process_log_text.setHtml(f"<p style='color:#29b6f6;'>📂 {day} - {len(lines)} sor (archív nézet; "
                                      f"az élő sorokhoz: 🟢 Élő (mai) nézet)</p>" + "<br>".join(lines))

    def clear_process_log(self):
        """Csak a NÉZETET üríti (a napi fájl megmarad)."""
        if hasattr(self, 'process_log_text'):
            self.process_log_text.clear()

    def closeEvent(self, event: QCloseEvent):
        for worker in list(self.active_workers):
            if worker.isRunning():
                worker.quit()
                worker.wait(1000)
        event.accept()
        
    # -------------------------------------------------------------------
    # TAB 9: HELYI ADATBÁZIS ÉLŐ NÉZET TAB (RÉSZLETES SZŰRÉSSEL)
    # -------------------------------------------------------------------
    def init_local_db_tab(self):
        layout = QVBoxLayout(self.tab_local_db)
        layout.setContentsMargins(10, 10, 10, 10)
        layout.setSpacing(8)

        top_bar = QHBoxLayout()
        lbl_title = QLabel("💾 A Helyi SQLite Adatbázis Élő Tartalma (Legfrissebb felvitel legfelül)")
        lbl_title.setStyleSheet("font-size: 16px; font-weight: bold; color: #00e676;")
        top_bar.addWidget(lbl_title)
        top_bar.addStretch()

        btn_manual_sync = QPushButton("☁️ Helyi Adatok Kézi Feltöltése a Felhőbe")
        btn_manual_sync.setStyleSheet("background-color: #ff9800; color: black; font-weight: bold; padding: 6px 12px;")
        btn_manual_sync.clicked.connect(
            lambda: self.run_button_action(
                btn_manual_sync, 
                "☁️ Helyi Adatok Kézi Feltöltése a Felhőbe", 
                self.trigger_manual_cloud_sync
            )
        )
        top_bar.addWidget(btn_manual_sync)

        self.lbl_local_db_count = QLabel("📊 Találatok: 0 db")
        self.lbl_local_db_count.setStyleSheet("font-weight: bold; color: #ffb74d; font-size: 14px;")
        top_bar.addWidget(self.lbl_local_db_count)

        btn_refresh_table = QPushButton("🔄 Adatbázis Újratöltése")
        btn_refresh_table.setStyleSheet("background-color: #29b6f6; color: black; font-weight: bold; padding: 6px 12px;")
        btn_refresh_table.clicked.connect(
            lambda: self.run_button_action(
                btn_refresh_table, 
                "🔄 Adatbázis Újratöltése", 
                self.refresh_local_db_table_view
            )
        )
        top_bar.addWidget(btn_refresh_table)

        layout.addLayout(top_bar)

        filter_box = QGroupBox("🔍 Részletes Keresés és Szűrés")
        filter_layout = QHBoxLayout(filter_box)
        filter_layout.setContentsMargins(8, 6, 8, 6)
        filter_layout.setSpacing(10)

        filter_layout.addWidget(QLabel("🔍 Keresés:"))
        self.local_filter_text = QLineEdit()
        self.local_filter_text.setPlaceholderText("Kód, Márka, Típus, Cikkszám, Helyszín (Enter a kereséshez)...")
        
        # A textChanged HELYETT csak Enter megnyomásakor vagy a fókusz elvesztésekor frissítjük a táblázatot:
        self.local_filter_text.returnPressed.connect(self.refresh_local_db_table_view)
        self.local_filter_text.editingFinished.connect(self.refresh_local_db_table_view)
        
        filter_layout.addWidget(self.local_filter_text, stretch=2)

        filter_layout.addWidget(QLabel("Állapot:"))
        self.local_filter_status = QRadioButton("Mind")
        self.local_filter_used = QRadioButton("HASZNÁLT")
        self.local_filter_new = QRadioButton("ÚJ")
        self.local_filter_status.setChecked(True)

        status_group = QWidget()
        status_lay = QHBoxLayout(status_group)
        status_lay.setContentsMargins(0, 0, 0, 0)
        status_lay.setSpacing(6)
        for btn in [self.local_filter_status, self.local_filter_used, self.local_filter_new]:
            status_lay.addWidget(btn)
            btn.toggled.connect(self.refresh_local_db_table_view)
        filter_layout.addWidget(status_group)

        filter_layout.addWidget(QLabel("Szín:"))
        self.local_filter_color_all = QRadioButton("Mind")
        self.local_filter_color_blue = QRadioButton("Kék")
        self.local_filter_color_orange = QRadioButton("Narancs")
        self.local_filter_color_red = QRadioButton("Piros")
        self.local_filter_color_all.setChecked(True)

        color_group = QWidget()
        color_lay = QHBoxLayout(color_group)
        color_lay.setContentsMargins(0, 0, 0, 0)
        color_lay.setSpacing(6)
        for btn in [self.local_filter_color_all, self.local_filter_color_blue, self.local_filter_color_orange, self.local_filter_color_red]:
            color_lay.addWidget(btn)
            btn.toggled.connect(self.refresh_local_db_table_view)
        filter_layout.addWidget(color_group)

        btn_clear_filters = QPushButton("🧹 Szűrők Alaphelyzetbe")
        btn_clear_filters.setStyleSheet("background-color: #555; color: white; padding: 4px 8px; border-radius: 4px;")
        btn_clear_filters.clicked.connect(self.reset_local_db_filters)
        filter_layout.addWidget(btn_clear_filters)

        layout.addWidget(filter_box)

        self.local_db_table = QTableWidget()
        self.local_db_table.setColumnCount(16)
        headers = ["A: F-Kód", "B: Vonalkód", "C: Márka", "D: Típus", "E: Cikkszám", "F: DB", 
                   "G: Állapot", "H: Ár (EUR)", "I: Páter", "J: Szabad", "K: Polc", "L: Raklap", 
                   "M: Súly (kg)", "N: eBay ID", "O: Megjegyzés", "P: Állapot részletei"]
        self.local_db_table.setHorizontalHeaderLabels(headers)
        self.local_db_table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Interactive)
        
        # Cella frissítési szignál rákötése a helyi mentésre
        self.local_db_table.itemChanged.connect(self.on_local_table_cell_edited)
        
        self.local_db_table.setStyleSheet("""
            QTableWidget {
                background-color: #121212;
                color: #ffffff;
                gridline-color: #333333;
                font-size: 13px;
                border: 1px solid #333;
            }
            QHeaderView::section {
                background-color: #1e1e1e;
                color: #29b6f6;
                font-weight: bold;
                padding: 6px;
                border: 1px solid #333;
            }
            QTableWidget::item:selected {
                background-color: #29b6f6;
                color: #000000;
            }
        """)
        layout.addWidget(self.local_db_table)

    def reset_local_db_filters(self):
        self.local_filter_text.blockSignals(True)
        self.local_filter_text.clear()
        self.local_filter_text.blockSignals(False)

        self.local_filter_status.setChecked(True)
        self.local_filter_color_all.setChecked(True)
        self.refresh_local_db_table_view()

    # -------------------------------------------------------------------
    # FÜLSOR: 4 fő fül mindig látszik, a többi a "☰ További fülek" menüből nyílik
    # -------------------------------------------------------------------
    def setup_tab_menu(self):
        self._main_tabs = [w for w in (getattr(self, "tab_excel_live", None), self.tab_search, self.tab_pda, self.tab_sold, self.tab_leltar, self.tab_versions) if w]
        self._extra_tabs = [self.tabs.widget(i) for i in range(self.tabs.count()) if self.tabs.widget(i) not in self._main_tabs]

        self.btn_more_tabs = QToolButton()
        self.btn_more_tabs.setText("☰ További fülek ▾")
        self.btn_more_tabs.setPopupMode(QToolButton.ToolButtonPopupMode.InstantPopup)
        self.btn_more_tabs.setStyleSheet("QToolButton { background-color: #2D2D2D; color: #29b6f6; font-weight: bold; "
                                         "padding: 6px 14px; border: 1px solid #29b6f6; border-radius: 4px; }"
                                         "QToolButton::menu-indicator { image: none; }")
        menu = QMenu(self.btn_more_tabs)
        menu.setStyleSheet("QMenu { background-color: #1E1E1E; color: white; font-size: 15px; border: 1px solid #29b6f6; }"
                           "QMenu::item { padding: 8px 24px; } QMenu::item:selected { background-color: #29b6f6; color: black; }")
        for w in self._extra_tabs:
            act = menu.addAction(self.tabs.tabText(self.tabs.indexOf(w)))
            act.triggered.connect(lambda checked=False, tw=w: self.open_extra_tab(tw))
        self.btn_more_tabs.setMenu(menu)
        self.tabs.setCornerWidget(self.btn_more_tabs, Qt.Corner.TopRightCorner)
        self._hide_extra_tabs()

    def _hide_extra_tabs(self, keep=None):
        bar = self.tabs.tabBar()
        for w in self._extra_tabs:
            idx = self.tabs.indexOf(w)
            if idx >= 0 and w is not keep and w is not self.tabs.currentWidget():
                bar.setTabVisible(idx, False)

    def open_extra_tab(self, w):
        """A menüből választott fül: ideiglenesen megjelenik a fülsoron (hogy lásd, hol vagy), és rá vált."""
        self.tabs.tabBar().setTabVisible(self.tabs.indexOf(w), True)
        self.tabs.setCurrentWidget(w)
        self._hide_extra_tabs(keep=w)
        self.btn_more_tabs.setText(f"☰ {self.tabs.tabText(self.tabs.indexOf(w))} ▾")

    def _tab_is_current(self, tab_attr):
        """Igaz, ha a megadott fül látszik (a fülek létrehozása előtt mindig igaz)."""
        if not hasattr(self, "tabs") or not hasattr(self, tab_attr):
            return True
        return self.tabs.currentWidget() is getattr(self, tab_attr)

    def on_tab_changed(self, index):
        """Fülváltás: a háttérben elhalasztott frissítéseket csak most, a fül megjelenése UTÁN végezzük el."""
        w = self.tabs.widget(index)
        # Fő fülre lépve a menüből nyitott fül újra elrejtőzik
        if hasattr(self, "_main_tabs") and w in self._main_tabs:
            QTimer.singleShot(0, self._hide_extra_tabs)
            self.btn_more_tabs.setText("☰ További fülek ▾")

        def later():
            if w is getattr(self, "tab_search", None) and getattr(self, "_cards_dirty", False):
                self.update_cards_only()
            elif w is getattr(self, "tab_local_db", None) and getattr(self, "_local_db_dirty", False):
                self.refresh_local_db_table_view()
            elif w is getattr(self, "tab_excel_live", None) and getattr(self, "_excel_dirty", False):
                self.silent_update_excel_table()
            elif w is getattr(self, "tab_sold", None):
                self.refresh_sold_table()
                if time.time() - getattr(self, "_sold_pull_last", 0) > 60:
                    self.pull_sales()
            elif w is getattr(self, "tab_server", None) and not getattr(self, "_server_items", None):
                self.load_server_items()
            elif w is getattr(self, "tab_map", None):
                self.ensure_map_view()
            elif w is getattr(self, "tab_leltar", None):
                self.leltar_refresh_old()
                self.leltar_load_list()
            elif WEBES_RENDELESEK_BE and w is getattr(self, "tab_web", None):
                self.web_load()
            elif w is getattr(self, "tab_versions", None):
                self.ver_refresh_local()
                self.check_app_update()
                self.admin_load_clients()
            elif w is getattr(self, "tab_admin", None):
                self._admin_tab_changed(self.admin_tabs.currentIndex())
        QTimer.singleShot(30, later)

    def refresh_local_db_table_view(self):
        """Helyi SQLite táblázat frissítése: Legfelül a legutoljára KITÖLTÖTT (C oszlop) legnagyobb R18-as tételek"""
        # A több ezer soros táblát csak akkor építjük újra, ha a fül látszik; különben fülváltáskor
        if not self._tab_is_current("tab_local_db"):
            self._local_db_dirty = True
            return
        self._local_db_dirty = False
        try:
            conn = sqlite3.connect(DB_FILE)
            cursor = conn.cursor()

            base_query = """
                SELECT rowid, A, B, C, D, E, F, G, H, I, J, K, L, M, N, O, P 
                FROM termekek
            """
            
            where_clauses = []
            params = []

            # 1. Keresőmező szűrése
            local_search = self.local_filter_text.text().strip() if hasattr(self, 'local_filter_text') else ""
            if local_search:
                where_clauses.append("(A LIKE ? OR B LIKE ? OR C LIKE ? OR D LIKE ? OR E LIKE ? OR L LIKE ?)")
                p_val = f"%{local_search}%"
                params.extend([p_val] * 6)

            # 2. Állapot szűrők
            if hasattr(self, 'local_filter_used') and self.local_filter_used.isChecked():
                where_clauses.append("G = 'HASZNÁLT'")
            elif hasattr(self, 'local_filter_new') and self.local_filter_new.isChecked():
                where_clauses.append("G = 'ÚJ'")

            if where_clauses:
                base_query += " WHERE " + " AND ".join(where_clauses)

            # 3. OKOS RENDEZÉS:
            # - Először azokat teszi legfelülre, ahol a C oszlop (Márka) KI VAN TÖLTVE!
            # - Azokon belül a legnagyobb R18 szám kerül az 1. sorba (pl. R18-12878)
            base_query += """
                ORDER BY 
                    CASE 
                        WHEN C IS NOT NULL AND TRIM(C) != '' AND LOWER(TRIM(C)) != 'none' THEN 1 
                        ELSE 0 
                    END DESC,
                    CASE 
                        WHEN A LIKE 'R18-%' THEN CAST(SUBSTR(A, 5) AS INTEGER)
                        ELSE 0 
                    END DESC,
                    rowid DESC
            """

            cursor.execute(base_query, params)
            rows = cursor.fetchall()
            conn.close()

            self.local_db_table.blockSignals(True)
            self.local_db_table.setRowCount(len(rows))

            for r_idx, row in enumerate(rows):
                for c_idx in range(16):
                    val = str(row[c_idx + 1]) if row[c_idx + 1] is not None and str(row[c_idx + 1]).lower() != "none" else ""
                    item = QTableWidgetItem(val)

                    # Igazítások
                    if c_idx in (8, 10, 12):
                        item.setTextAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
                    elif c_idx == 11:
                        if val and re.match(r'^\d+([.,]\d+)?$', val):
                            item.setTextAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
                        else:
                            item.setTextAlignment(Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignVCenter)
                    else:
                        item.setTextAlignment(Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignVCenter)

                    self.local_db_table.setItem(r_idx, c_idx, item)

            self.local_db_table.blockSignals(False)

            if hasattr(self, 'lbl_local_db_count'):
                self.lbl_local_db_count.setText(f"📊 Találatok: {len(rows)} db")

        except Exception as e:
            print(f"Hiba a Helyi DB táblázat frissítésekor: {e}")
            
    
    
    def open_sheets_version_history(self):
        """Közvetlenül megnyitja a Google Sheets Verzióelőzmények (Version History) oldalát a böngészőben"""
        version_history_url = f"https://docs.google.com/spreadsheets/d/{TEST_SHEETS_ID}/history"
        QDesktopServices.openUrl(QUrl(version_history_url))
        self.trigger_toast("📜 Google Sheets Verzióelőzmények megnyitása a böngészőben...")
    # =========================================================================
    # ÚJ EXCEL LIVE MUNKALAP FÜL METÓDUSAI
    # =========================================================================

    def setup_excel_style_tab(self):
        """Excel Live Munkalap: Kézi Mentés Gombbal és Automatikus Cellafelismeréssel"""
        self._excel_dirty = False
        self.tab_excel_live = QWidget()
        layout = QVBoxLayout(self.tab_excel_live)

        # 1. VEZÉRLŐSÁV (Kézi mentés gomb + Pozícionálás)
        ctrl_box = QGroupBox("📌 Munkaterület Pozícionálása & Mentés")
        ctrl_layout = QHBoxLayout(ctrl_box)
        ctrl_layout.setContentsMargins(8, 4, 8, 4)
        ctrl_layout.setSpacing(6)

        # 💾 KÉZI MENTÉS GOMB
        btn_manual_save = QPushButton("💾 MENTÉS MOST (SQLite + VPS)")
        btn_manual_save.setStyleSheet("background-color: #ff9800; color: black; font-weight: bold; font-size: 14px; padding: 6px 14px;")
        btn_manual_save.clicked.connect(self.force_save_excel_table)
        ctrl_layout.addWidget(btn_manual_save)

        ctrl_layout.addWidget(QLabel(" |  Kezdő R18 Szám:"))

        self.input_start_row = QLineEdit("12876")
        self.input_start_row.setFixedWidth(90)
        self.input_start_row.setStyleSheet("font-weight: bold; font-size: 14px; padding: 4px;")
        self.input_start_row.returnPressed.connect(self.reload_excel_tab_from_num)
        ctrl_layout.addWidget(self.input_start_row)

        btn_up = QPushButton("▲")
        btn_up.setFixedSize(28, 28)
        btn_up.setStyleSheet("background-color: #333; color: #00e676; font-weight: bold;")
        btn_up.clicked.connect(lambda: self.step_r18_number(1))
        ctrl_layout.addWidget(btn_up)

        btn_down = QPushButton("▼")
        btn_down.setFixedSize(28, 28)
        btn_down.setStyleSheet("background-color: #333; color: #ff5252; font-weight: bold;")
        btn_down.clicked.connect(lambda: self.step_r18_number(-1))
        ctrl_layout.addWidget(btn_down)

        btn_jump = QPushButton("🚀 Ugrás")
        btn_jump.setStyleSheet("background-color: #0288d1; color: white; font-weight: bold; padding: 5px 10px;")
        btn_jump.clicked.connect(self.reload_excel_tab_from_num)
        ctrl_layout.addWidget(btn_jump)

        btn_auto_latest = QPushButton("⚡ Auto (Legújabb)")
        btn_auto_latest.setStyleSheet("background-color: #00e676; color: black; font-weight: bold; padding: 5px 10px;")
        btn_auto_latest.clicked.connect(self.auto_set_latest_and_reload)
        ctrl_layout.addWidget(btn_auto_latest)

        # Mindig látható nyomtatás gomb a kijelölt sorra (a táblázat gördítésétől függetlenül),
        # a felirat mutatja, melyik R18-kód címkéje jön ki
        self.btn_print_selected = QPushButton("🖨️ Nyomtatás: – (Ctrl+P)")
        self.btn_print_selected.setStyleSheet("background-color: #ff9800; color: black; font-weight: bold; font-size: 15px; padding: 5px 14px;")
        self.btn_print_selected.setToolTip("A kijelölt sor címkéjét nyomtatja (ugyanaz, mint a Ctrl+P)")
        self.btn_print_selected.clicked.connect(self.print_current_excel_row)

        btn_clear_rows = QPushButton("🗑️ Adat törlés (R18 marad)")
        btn_clear_rows.setStyleSheet("background-color: #ff5252; color: white; font-weight: bold; font-size: 14px; padding: 5px 12px;")
        btn_clear_rows.setToolTip("A kijelölt sor(ok) MINDEN adatát törli (Márka, Típus, Cikkszám, DB, Raklap ...),\n"
                                  "az R18-kód (A-B) megmarad. Előtte rákérdez.\n"
                                  "Teljes sor kijelölése: kattints a sor számára (bal oldalt), több sor: Shift / Ctrl + kattintás.\n"
                                  "Másolás: Ctrl+C (A-P sorrendben, Sheets-be beilleszthető) | Beillesztés: Ctrl+V")
        btn_clear_rows.clicked.connect(self.clear_selected_excel_rows)
        ctrl_layout.addWidget(btn_clear_rows)

        # G (Állapot) legördülő lista ki/be - gépenként megjegyzi (beallitasok.json)
        self.chk_allapot_dropdown = QCheckBox("▾ G legördülő (Új / Használt / Felújított)")
        self.chk_allapot_dropdown.setStyleSheet("color: white; font-size: 13px; padding-left: 8px;")
        self.chk_allapot_dropdown.setToolTip("Bepipálva: a G (Állapot) cellára kattintva lenyílik a választólista.\n"
                                             "Kikapcsolva: a G oszlopba ugyanúgy gépelni kell, mint a többibe.\n"
                                             "A beállítást ez a gép megjegyzi.")
        self.chk_allapot_dropdown.setChecked(allapot_dropdown_enabled())
        self.chk_allapot_dropdown.toggled.connect(lambda on: save_config_value("allapot_legordulo", bool(on)))
        ctrl_layout.addWidget(self.chk_allapot_dropdown)

        ctrl_layout.addStretch()
        # Keskenyebb képernyőn (pl. Windows 125-150% nagyítás) a gombsor vízszintesen görgethető,
        # így nem ez szabja meg az ablak legkisebb szélességét (különben nem lehetne teljes méretre állítani)
        ctrl_scroll = QScrollArea()
        ctrl_scroll.setWidget(ctrl_box)
        ctrl_scroll.setWidgetResizable(True)
        ctrl_scroll.setFrameShape(QFrame.Shape.NoFrame)
        ctrl_scroll.setVerticalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        ctrl_scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)
        ctrl_scroll.setFixedHeight(ctrl_box.sizeHint().height() + 18)
        ctrl_scroll.setMinimumWidth(300)
        layout.addWidget(ctrl_scroll)

        # 2. Excel-szerű Táblázat testreszabása
        self.excel_table = ExcelTableWidget()
        self.excel_table.setColumnCount(17)
        headers = ["A: F-Kód", "B: Vonalkód", "C: Márka", "D: Típus", "E: Cikkszám", "F: DB",
                   "G: Állapot", "H: Ár", "I: Páter", "J: Szabad", "K: Polc", "L: Raklap",
                   "M: Súly", "N: eBay ID", "O: Megjegyzés", "P: Állapot részlet", "🖨️ Nyomtatás"]
        self.excel_table.setHorizontalHeaderLabels(headers)
        # KÉPERNYŐS oszlopsorrend: 🖨️, A-G, majd a leggyakrabban írt L (Raklap) és M (Súly), utána I, J, K, H, N, O, P.
        # Csak a megjelenítés változik (moveSection): a program és a mentés továbbra is a logikai
        # oszlopokkal dolgozik (0-15 = A-P, 16 = gomb), így minden érték a saját oszlopába kerül.
        hdr = self.excel_table.horizontalHeader()
        for visual_pos, logical_col in enumerate(EXCEL_LIVE_VISUAL_ORDER):
            hdr.moveSection(hdr.visualIndex(logical_col), visual_pos)
        # A J (Szabad) oszlopot nem használjuk: elrejtjük. Az adata megmarad és a mentés változatlan.
        self.excel_table.setColumnHidden(EXCEL_LIVE_HIDDEN_COL_J, True)

        # Ctrl+P: a kijelölt sor címkéjének nyomtatása (gépelés közben is működik)
        self.excel_print_shortcut = QShortcut(QKeySequence("Ctrl+P"), self.excel_table)
        self.excel_print_shortcut.setContext(Qt.ShortcutContext.WidgetWithChildrenShortcut)
        self.excel_print_shortcut.activated.connect(self.print_current_excel_row)
        self.excel_table.currentCellChanged.connect(self.update_print_selected_label)
        # a gépelés miatt kihagyott (másik gépről jött) sorok pótlása, amint már nem gépelnek benne
        self._excel_deferred = False
        self.excel_deferred_timer = QTimer()
        self.excel_deferred_timer.timeout.connect(
            lambda: self._excel_deferred and self.excel_table.state() != QAbstractItemView.State.EditingState
            and self.silent_update_excel_table())
        self.excel_deferred_timer.start(3000)
        # G (Állapot): egy kattintásra lenyílik a választólista (ÚJ / HASZNÁLT / FELÚJÍTOTT)
        self.excel_table.cellClicked.connect(
            lambda r, c: self.excel_table.editItem(self.excel_table.item(r, c))
            if c == ALLAPOT_COL and self.excel_table.item(r, c) is not None
            and self.chk_allapot_dropdown.isChecked() else None)

        # STÍLUS: Szürkés háttér, jól látható rácsvonalak és kontrasztos kijelölés
        self.excel_table.setAlternatingRowColors(True) # Váltakozó sormegvilágítás bekapcsolása
        self.excel_table.setStyleSheet("""
            QTableWidget {
                background-color: #21252b;
                alternate-background-color: #282c34;
                color: #abb2bf;
                gridline-color: #3b4048;
                font-size: 13px;
                border: 1px solid #181a1f;
                selection-background-color: #61afef;
                selection-color: #000000;
            }
            QHeaderView::section {
                background-color: #1b1e23;
                color: #61afef;
                font-weight: bold;
                padding: 6px;
                border: 1px solid #3b4048;
            }
            QTableCornerButton::section {
                background-color: #1b1e23;
                border: 1px solid #3b4048;
            }
            /* Kijelölt cellák: KIFEJEZETTEN meg kell adni (aktív és inaktív ablakra is),
               mert az ::item:focus szabály miatt a Qt a fenti selection-* színeket figyelmen
               kívül hagyja, és a kijelölt cellák szövege olvashatatlan lesz. */
            /* Egér a cella fölött: halvány kiemelés, a szöveg fehér marad */
            QTableWidget::item:hover {
                background-color: #3a4556;
                color: #ffffff;
            }
            QTableWidget::item:selected,
            QTableWidget::item:selected:!active,
            QTableWidget::item:selected:hover {
                background-color: #2f5f8f;
                color: #ffffff;
            }
            /* Aktuális cella (nyilakkal / kattintással): minden állapot-kombinációban
               ugyanúgy néz ki, hogy a fenti szabályok ne írják felül */
            QTableWidget::item:focus,
            QTableWidget::item:focus:hover,
            QTableWidget::item:selected:focus,
            QTableWidget::item:selected:focus:hover,
            QTableWidget::item:selected:focus:!active {
                background-color: #61afef;
                color: #000000;
                font-weight: bold;
                border: 2px solid #ffffff;
            }
        """)

        # MAGASABB SOROK: a hosszú Márka / Típus / Cikkszám (C, D, E) szöveg két sorba tördelve
        # kifér, nem vágódik le "..."-tal, és a cellaszerkesztő is kényelmesebb
        self.excel_table.setWordWrap(True)
        self.excel_table.setTextElideMode(Qt.TextElideMode.ElideNone)
        self.excel_table.verticalHeader().setDefaultSectionSize(EXCEL_LIVE_ROW_HEIGHT)
        self.excel_table.verticalHeader().setMinimumSectionSize(EXCEL_LIVE_ROW_HEIGHT)

        # 🖨️ oszlop: rajzolt gombok (lásd PrintButtonDelegate) - nem kell soronként widget
        self.excel_print_delegate = PrintButtonDelegate(self.print_row_label, self.excel_table)
        self.excel_table.setItemDelegateForColumn(16, self.excel_print_delegate)
        self.excel_table.setMouseTracking(True)
        # a 🖨️ oszlop rögzítve: oldalra görgetéskor is mindig látszik a bal szélen
        self.excel_print_frozen = FrozenColumn(self.excel_table, 16, PrintButtonDelegate(self.print_row_label, self.excel_table))

        # Az Excel Live tábla aktuális tartalma soronként (gyors összehasonlításhoz a háttérfrissítésnél)
        self._excel_rows_cache = []
        self._excel_code_index = {}
        self._excel_loading = False
        self.excel_rows_signal.connect(self._apply_excel_rows)

        self._excel_row_sources = {}   # sor -> ["⌨️ Billentyűzet - L: '573' felülírva → '574'", ...] (Napló)
        self.excel_table.itemChanged.connect(self.on_excel_cell_changed)

        # 🆕 Másik gépen felvitt / most kitöltés alatt lévő sorok (kattintásra odaugrik / nyomtat)
        self.remote_bar = QWidget()
        rb = QHBoxLayout(self.remote_bar)
        rb.setContentsMargins(0, 2, 0, 2)
        # 🖨️ a kijelölt sor nyomtatása – itt mindig látszik (a felső gombsor keskeny képernyőn oldalra kicsúszhat)
        self.btn_print_selected.setMinimumWidth(300)
        rb.addWidget(self.btn_print_selected)
        rb.addSpacing(12)
        lbl = QLabel("🆕 Másik gépen:")
        lbl.setStyleSheet("color: #00e676; font-weight: bold; font-size: 14px;")
        rb.addWidget(lbl)
        self.remote_chips_box = QWidget()
        self.remote_chips = QHBoxLayout(self.remote_chips_box)
        self.remote_chips.setContentsMargins(0, 0, 0, 0)
        chips_scroll = QScrollArea()
        chips_scroll.setWidget(self.remote_chips_box)
        chips_scroll.setWidgetResizable(True)
        chips_scroll.setFrameShape(QFrame.Shape.NoFrame)
        chips_scroll.setVerticalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        chips_scroll.setFixedHeight(46)
        rb.addWidget(chips_scroll, stretch=1)
        btn_clear_remote = QPushButton("✖ Lista ürítése")
        btn_clear_remote.setToolTip("A másik gépen felvitt sorok listáját és jelölését eltünteti (az adat nem változik)")
        btn_clear_remote.clicked.connect(self.clear_remote_marks)
        rb.addWidget(btn_clear_remote)
        self.remote_bar.setFixedHeight(54)
        layout.addWidget(self.remote_bar)
        # a rögzített 🖨️ sáv (sorszám + Nyomtatás) a táblázat bal oldalán
        excel_row = QHBoxLayout()
        excel_row.setSpacing(0)
        excel_row.addWidget(self.excel_print_frozen.v)
        excel_row.addWidget(self.excel_table, stretch=1)
        layout.addLayout(excel_row, stretch=1)

        # Jelenlét: a többi gép látja, melyik soron dolgozunk éppen (2 mp-enként, csak ha tényleg szerkesztünk)
        self._presence_sent, self._presence_ts, self._last_excel_edit = "", 0.0, 0.0
        self.presence_timer = QTimer()
        self.presence_timer.timeout.connect(self._presence_tick)
        # (a „most tölti” jelzés kikapcsolva – a felhasználó kérésére)

        # 3. Módosított sorok mentési időzítője (4 mp)
        self.pending_dirty_rows = set()
        self.auto_save_timer = QTimer()
        self.auto_save_timer.setInterval(AUTO_SAVE_INTERVAL_MS)
        self.auto_save_timer.timeout.connect(self.process_auto_row_sync)
        self.auto_save_timer.start()

        # --- OFFLINE-BIZTOS ÚJRAPRÓBÁLKOZÁS ---
        # 20 másodpercenként megpróbálja kiküldeni, ami korábban net hiánya miatt
        # a sync_outbox-ban rekedt (élő tábla módosítások ÉS új termékfelvitel egyaránt).
        self.outbox_flush_timer = QTimer()
        self.outbox_flush_timer.setInterval(20000)
        self.outbox_flush_timer.timeout.connect(lambda: threading.Thread(target=flush_sync_outbox, daemon=True).start())
        self.outbox_flush_timer.start()

        # R18-ÜTKÖZÉS FIGYELMEZTETÉS: bármelyik háttérszálról jöhet (flush_sync_outbox),
        # a jelzőbusz biztonságosan áthozza a fő szálra ezt a Qt-kapcsolatot.
        outbox_signal_bus.conflicts_detected.connect(self.on_outbox_conflicts)

        if hasattr(self, 'tabs'):
           self.tabs.addTab(self.tab_excel_live, "📊 Excel Live Munkalap")

        self.auto_set_latest_and_reload()

    def step_r18_number(self, delta):
        """A beviteli mezőben lévő R18 szám léptetése fel (1) vagy le (-1) irányba"""
        raw_val = self.input_start_row.text().strip().upper().replace("R18-", "")
        try:
            val = int(raw_val) + delta
            if val < 1:
                val = 1
        except ValueError:
            val = 12876

        self.input_start_row.setText(str(val))
        self.reload_excel_tab_from_num()

    def auto_set_latest_and_reload(self):
        """Automatikusan megkeresi az utolsó kitöltött R18 számot"""
        try:
            conn = sqlite3.connect(DB_FILE)
            cursor = conn.cursor()
            cursor.execute("""
                SELECT CAST(SUBSTR(A, 5) AS INTEGER) 
                FROM termekek 
                WHERE A LIKE 'R18-%' AND C IS NOT NULL AND TRIM(C) != '' AND LOWER(TRIM(C)) != 'none'
                ORDER BY CAST(SUBSTR(A, 5) AS INTEGER) DESC LIMIT 1
            """)
            row = cursor.fetchone()
            conn.close()
            
            if row and row[0]:
                self.input_start_row.setText(str(row[0]))
        except Exception as e:
            print(f"Hiba a legújabb sor megállapításakor: {e}")

        self.reload_excel_tab_from_num()

    def reload_excel_tab_from_num(self, full=False):
        """A kiválasztott R18-kódra görget. A TELJES (több ezer soros) betöltés csak az első alkalommal fut
        (vagy full=True); utána a ▲▼ / Ugrás / Auto gombok már csak görgetnek - az adatokat a háttérben
        futó csendes frissítés (silent_update_excel_table) tartja naprakészen."""
        raw_val = self.input_start_row.text().strip().upper().replace("R18-", "")
        try:
            target_num = int(raw_val)
        except ValueError:
            target_num = 12878

        if not full and self._excel_rows_cache:
            self._scroll_excel_to_num(target_num)
            self.silent_update_excel_table()   # közben felvett új sorok behozása (háttérben)
            return

        try:
            conn = sqlite3.connect(DB_FILE)
            cursor = conn.cursor()
            # Az összes sort beolvassuk az R18-1-től kezdve emelkedő sorrendben!
            cursor.execute("""
                SELECT A, B, C, D, E, F, G, H, I, J, K, L, M, N, O, P 
                FROM termekek 
                WHERE A LIKE 'R18-%'
                ORDER BY CAST(SUBSTR(A, 5) AS INTEGER) ASC
            """)
            rows = cursor.fetchall()
            conn.close()
        except Exception as e:
            print(f"Hiba az összes sor betöltésekor: {e}")
            rows = []

        rows = [_normalize_excel_row(r) for r in rows]

        self.excel_table.setUpdatesEnabled(False)
        self.excel_table.blockSignals(True)
        self.excel_table.setRowCount(len(rows))

        for r_idx, row in enumerate(rows):
            for c_idx in range(16):
                it = QTableWidgetItem(row[c_idx])
                if c_idx in (11, 12):
                    align_excel_item(it, c_idx)
                self.excel_table.setItem(r_idx, c_idx, it)

        self.excel_table.blockSignals(False)
        self.excel_table.setUpdatesEnabled(True)
        self._excel_rows_cache = list(rows)
        self._rebuild_excel_code_index()
        self._remote_marked_rows = set()
        self.render_remote_marks()

        # -----------------------------------------------------------------
        # FIX OSZLOPSZÉLESSÉGEK (Villámgyors töltés, nincs fagyás sok ezres sornál!)
        # -----------------------------------------------------------------
        col_widths = [110, 110, 150, 260, 250, 60, 100, 80, 80, 80, 100, 140, 80, 110, 150, 140, 140]
        for c_idx, w in enumerate(col_widths):
            self.excel_table.setColumnWidth(c_idx, w)

        self._scroll_excel_to_num(target_num)

    def _rebuild_excel_code_index(self):
        self._excel_code_index = {row[0].strip().upper(): i for i, row in enumerate(self._excel_rows_cache) if row}

    def _scroll_excel_to_num(self, target_num):
        """GÖRGETÉS ÉS FÓKUSZ A KÍVÁNT R18-SZÁMRA"""
        rows = self._excel_rows_cache
        target_row_in_view = self._excel_code_index.get(f"R18-{target_num}", self._excel_code_index.get(str(target_num), -1))
        if target_row_in_view != -1:
            target_item = self.excel_table.item(target_row_in_view, 2) # C oszlop (Márka)
            if target_item:
                self.excel_table.setCurrentItem(target_item)
                # A képernyő közepére pozícionálja a megadott sort, így alatta-felette is látszanak a szomszédos tételek
                self.excel_table.scrollToItem(target_item, self.excel_table.ScrollHint.PositionAtCenter)
                self.excel_table.horizontalScrollBar().setValue(0)   # vízszintesen mindig az A oszlopnál kezdődjön
        elif len(rows) > 0:
            # Ha a pontos szám még nem létezik az adatbázisban, az utolsó létező sorra ugrik
            last_idx = len(rows) - 1
            target_item = self.excel_table.item(last_idx, 2)
            if target_item:
                self.excel_table.setCurrentItem(target_item)
                self.excel_table.scrollToItem(target_item, self.excel_table.ScrollHint.PositionAtCenter)
                self.excel_table.horizontalScrollBar().setValue(0)   # vízszintesen mindig az A oszlopnál kezdődjön
            
            
    def clear_selected_excel_rows(self):
        """🗑️ Adat törlés: a kijelölt sorok C-P adatait törli (az R18-kód: A-B megmarad), megerősítés után."""
        t = self.excel_table
        t.setFocus()   # ha épp gépel egy cellában, az elmentődik
        rows = sorted({i.row() for i in t.selectedIndexes()} or ({t.currentRow()} if t.currentRow() >= 0 else set()))
        rows = [r for r in rows if any((t.item(r, c).text().strip() if t.item(r, c) else "") for c in range(2, 16))]
        if not rows:
            self.trigger_toast("⚠️ A kijelölt sor(ok)ban nincs törölhető adat.")
            return

        def desc(r):
            g = lambda c: t.item(r, c).text().strip() if t.item(r, c) else ""
            return f"• {g(0)}  |  {g(2) or '-'}  |  {g(3) or '-'}  |  {g(4) or '-'}"
        lista = "\n".join(desc(r) for r in rows[:12]) + (f"\n... és még {len(rows) - 12} sor" if len(rows) > 12 else "")
        dlg = SilentConfirmDialog(
            "🗑️ Adatok törlése",
            f"Biztosan TÖRLÖD ennek a {len(rows)} sornak az összes adatát?\n"
            f"(Az R18-kód megmarad, a régi adatok a Naplóba és a ToroltSor lapra kerülnek.)\n\n{lista}",
            self)
        if dlg.exec() != QDialog.DialogCode.Accepted:
            return

        group = []
        t._edit_source = "🗑️ Adat törlés gomb (sor adatai törölve, R18 megmarad)"
        t._suppress_undo = True
        try:
            for r in rows:
                for c in range(2, 16):
                    it = t.item(r, c)
                    if it and it.text():
                        group.append((r, c, it.text(), ""))
                        it.setText("")
        finally:
            t._edit_source = None
            t._suppress_undo = False
        if group:
            t.undo_stack.append(group)
        self.trigger_toast(f"🗑️ {len(rows)} sor adatai törölve (Ctrl+Z: visszavonás)")

    def note_excel_source(self, row, col, old_val, new_val, how):
        """Feljegyzi, HOGYAN változott egy Excel Live cella - a Napló 'Honnan / hogyan' oszlopába kerül."""
        if old_val == new_val:
            return
        letter = "ABCDEFGHIJKLMNOP"[col] if 0 <= col < 16 else "?"
        what = (f"{letter}: '{old_val}' törölve" if old_val and not new_val else
                f"{letter}: '{old_val}' felülírva → '{new_val}'" if old_val else
                f"{letter}: üres cellába beírva '{new_val}'")
        note = f"{how} - {what}"
        # Vonalkódolvasó a rossz helyen? (R18-kód egy adatcellában)
        if col >= 2 and re.fullmatch(r"R18[-/]\d+", str(new_val).strip(), re.IGNORECASE):
            note += " ⚠️ R18-kód került egy adatcellába - valószínűleg vonalkódolvasó csippantás rossz helyen!"
        self._excel_row_sources.setdefault(row, []).append(note)

    def on_excel_cell_changed(self, item):
        """Amikor megváltozik egy cella értéke, eltároljuk a Ctrl+Z-hez és megjelöljük a sort mentésre"""
        if item:
            row = item.row()
            col = item.column()
            self.pending_dirty_rows.add(row)

            # L / M igazítás (jelzések nélkül, hogy ne számítson újabb módosításnak)
            if col in (11, 12):
                was = self.excel_table.signalsBlocked()
                self.excel_table.blockSignals(True)
                align_excel_item(item, col)
                self.excel_table.blockSignals(was)

            # Napló: honnan / hogyan változott (a régi érték az utoljára betöltött adatbázis-állapotból)
            if not self.excel_table.signalsBlocked():
                self._last_excel_edit = time.time()
                cache = getattr(self, "_excel_rows_cache", [])
                cached = cache[row] if row < len(cache) else None
                old = cached[col] if cached and col < len(cached) else getattr(item, '_previous_text', "")
                how = getattr(self.excel_table, "_edit_source", None) or "⌨️ Billentyűzet (gépelés a cellába)"
                self.note_excel_source(row, col, old, item.text(), how)

            # Ha a módosítás a táblázatban történt, elmentjük az Undo verembe
            if hasattr(self, 'excel_table') and not self.excel_table.signalsBlocked():
                # Az előzménykezeléshez elmentjük az utolsó egyedi módosítást
                old_val = getattr(item, '_previous_text', "")
                new_val = item.text()
                if old_val != new_val and not getattr(self.excel_table, "_suppress_undo", False):
                    self.excel_table.undo_stack.append([(row, col, old_val, new_val)])
                item._previous_text = new_val

    def _refresh_excel_row_from_db(self, row_idx):
        """Nyomtatás előtt: a sor a helyi adatbázis legfrissebb állapotát mutassa (ha közben nem gépelnek bele)."""
        try:
            if row_idx in getattr(self, "pending_dirty_rows", set()):
                return
            it = self.excel_table.item(row_idx, 0)
            code = it.text().strip().upper() if it else ""
            if not code:
                return
            conn = get_db_connection()
            cur = conn.cursor()
            cur.execute(f"SELECT {', '.join(ROW_COLS)} FROM termekek WHERE UPPER(TRIM(A))=? OR UPPER(TRIM(B))=?", (code, code))
            dbrow = cur.fetchone()
            conn.close()
            if not dbrow:
                return
            vals = _normalize_excel_row(tuple(dbrow))
            self.excel_table.blockSignals(True)
            for c_idx in range(16):
                val = vals[c_idx]
                item = self.excel_table.item(row_idx, c_idx)
                if item is None:
                    item = QTableWidgetItem(val)
                    self.excel_table.setItem(row_idx, c_idx, item)
                elif item.text() != val and not (c_idx == 0 and not val):
                    item.setText(val)
            self.excel_table.blockSignals(False)
        except Exception:
            self.excel_table.blockSignals(False)

    def print_row_label(self, row_idx):
        """🖨️ gomb: kattintás után bekéri a példányszámot (alap: 1, Enter = nyomtat), majd ZPL-lel nyomtat.
        Shift+kattintás: részletes nyomtatási ablak (PNG mód / előnézet)."""
        if self.excel_table.state() != QAbstractItemView.State.EditingState:
            self._refresh_excel_row_from_db(row_idx)
        try:
            item_a = self.excel_table.item(row_idx, 0) # A: F-Kód
            item_c = self.excel_table.item(row_idx, 2) # C: Márka
            item_d = self.excel_table.item(row_idx, 3) # D: Típus
            item_e = self.excel_table.item(row_idx, 4) # E: Cikkszám

            code_a = item_a.text().strip() if item_a else ""
            brand = item_c.text().strip() if item_c else ""
            item_type = item_d.text().strip() if item_d else ""
            part_num = item_e.text().strip() if item_e else ""

            if not code_a:
                self.trigger_toast("⚠️ Nincs F-Kód ebben a sorban!")
                return

            # Üres sor (nincs Márka, Típus, Cikkszám): valószínűleg mellékattintás - rákérdezünk
            if not (brand or item_type or part_num):
                dlg = SilentConfirmDialog(
                    "Üres sor nyomtatása?",
                    f"A(z) {code_a} sorban nincs Márka, Típus és Cikkszám sem,\n"
                    f"így a címke alja üres lesz.\n\nBiztosan ezt a sort nyomtatod?",
                    self
                )
                if dlg.exec() != QDialog.DialogCode.Accepted:
                    return

            # Shift + kattintás -> részletes nyomtatási ablak
            if QApplication.keyboardModifiers() & Qt.KeyboardModifier.ShiftModifier:
                dlg = RowPrintDialog(code_a, marka=brand, tipus=item_type, cikkszam=part_num, parent=self)
                dlg.exec()
                return

            # Sima kattintás -> példányszám bekérése (alap: 1), majd nyomtatás háttérszálon (a felület nem fagy le)
            code_u, tipus_u, cikk_u = code_a.upper(), item_type.upper(), part_num.upper()

            copies_dlg = CopiesDialog(code_u, self)
            if copies_dlg.exec() != QDialog.DialogCode.Accepted:
                return
            copies = copies_dlg.copies()

            def _job():
                try:
                    zpl = generate_zpl_code(code_u, cikkszam=cikk_u, tipus=tipus_u, copies=copies)
                    ok, msg = send_zpl_to_printer(get_zebra_printer_name(), zpl)
                except Exception as e:
                    ok, msg = False, str(e)
                self.print_result_signal.emit((f"{msg} [{code_u}]" if str(msg).startswith("📨") else f"🖨️ Címke kinyomtatva [{code_u}] - {copies} db")
                                              if ok else f"❌ Nyomtatási hiba [{code_u}]: {msg}")
                log_event(code_u, "CÍMKE NYOMTATÁS", (f"Excel Live gomb - elküldve a nyomtatós gépre ({copies} db)" if str(msg).startswith("📨")
                                                      else f"Excel Live gomb - OK ({copies} db)") if ok else f"HIBA: {msg}")

            threading.Thread(target=_job, daemon=True).start()

        except Exception as e:
            self.trigger_toast(f"❌ Nyomtatási hiba: {e}")

    def process_auto_row_sync(self):
        """4 mp-enként lefutó háttérszinkronizáció"""
        if not self.pending_dirty_rows:
            return

        # Ha az előző mentés még fut, a sorok pending-ben maradnak a következő körre
        if hasattr(self, 'row_sync_worker') and self.row_sync_worker.isRunning():
            return

        dirty_data = []
        rows_to_sync = list(self.pending_dirty_rows)
        self.pending_dirty_rows.clear()

        col_keys = ["A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P"]

        for row in rows_to_sync:
            row_dict = {}
            for col_idx, col_name in enumerate(col_keys):
                item = self.excel_table.item(row, col_idx)
                row_dict[col_name] = item.text().strip() if item else ""

            notes = self._excel_row_sources.pop(row, [])
            row_dict["_forras"] = "Excel Live - " + (" | ".join(notes) if notes else "cella szerkesztés")
            if row_dict.get("A"):
                dirty_data.append(row_dict)

        if dirty_data:
            self.row_sync_worker = AutoRowSyncWorker(dirty_data)
            # Ez a sor gondoskodik a felület azonnali frissítéséről!
            self.row_sync_worker.sync_finished_signal.connect(self.on_auto_row_sync_finished)
            self.row_sync_worker.start()

    def on_auto_row_sync_finished(self, msg):
        """Amikor a háttérszál végzett a mentéssel (SQLite + VPS + Sheets), csendben értesítjük a felhasználót ZÉRÓ UI fagyással"""
        # Toast értesítés kiírása (ez nem blokkolja a felületet)
        self.trigger_toast(msg)

        # 1. A táblázat újrarajzolása (refresh_local_db_table_view) HELYETT csak csendes frissítést hívunk!
        if hasattr(self, 'silent_update_excel_table'):
            self.silent_update_excel_table()

        # 2. A kártyákat csak akkor frissítjük, ha a Gyorskereső fülön vagyunk
        if hasattr(self, 'tabs') and self.tabs.currentWidget() is getattr(self, "tab_search", None):
            if hasattr(self, 'update_cards_only'):
                self.update_cards_only()

        # 3. Számlálók csendes frissítése
        if hasattr(self, 'refresh_counters'):
            self.refresh_counters()

    def on_outbox_conflicts(self, conflicts):
        """Amikor a flush_sync_outbox() olyan tételt talál, ami MÁS termékkel azonos R18-számot
        próbált használni: figyelmeztetjük a felhasználót, és a teljes adatok már a Naplóban
        vannak (log_event a flush_sync_outbox()-ban megtörtént) - itt csak jelezzük."""
        if not conflicts:
            return

        # Ugyanarra a kódra csak a legutolsó adat számít; ha már nyitva van az ablak, összegyűjtjük
        if not hasattr(self, "_pending_conflicts"):
            self._pending_conflicts = {}
        for c in conflicts:
            self._pending_conflicts[str(c.get("A", "")).strip().upper()] = c
        if getattr(self, "_conflict_dialog_open", False):
            return

        while self._pending_conflicts:
            batch = list(self._pending_conflicts.values())
            self._pending_conflicts = {}
            self._conflict_dialog_open = True
            try:
                dlg = ConflictDialog(batch, self)
                chosen = dlg.selected() if dlg.exec() == QDialog.DialogCode.Accepted else []
            finally:
                self._conflict_dialog_open = False

            for c in chosen:
                code = str(c.get("A", "")).strip().upper()
                new_row = {k: str(c.get(k, "")) for k in ROW_COLS}
                payload = dict(new_row)
                payload["_force"] = True          # a szerver felülírja a másik terméket
                enqueue_outbox(payload)
                log_event(code, "MÓDOSÍTÁS (R18-ütközés: felülírva az új adattal)",
                          f"Régi (szerveren volt): '{c.get('_server_existing', '')}' → új: {format_row_for_log(new_row)}",
                          None, new_row, "⚠️ Ütközés-ablak - ✅ Felülírás gomb (a felhasználó döntött)")
            skipped = [c for c in batch if c not in chosen]
            for c in skipped:
                log_event(str(c.get("A", "")).strip().upper(), "R18 ÜTKÖZÉS - NEM MENTVE (kihagyva)",
                          f"A felhasználó nem írta felül. Régi marad: '{c.get('_server_existing', '')}'",
                          None, {k: str(c.get(k, "")) for k in ROW_COLS}, "⚠️ Ütközés-ablak - Most egyiket sem")
            if chosen:
                flush_outbox_async()
                self.trigger_toast(f"✅ {len(chosen)} tétel felülírva az új adattal")
            if skipped:
                self.trigger_toast(f"⚠️ {len(skipped)} tétel nem mentődött - részletek a 📜 Naplóban")

        if hasattr(self, 'refresh_db_log_view'):
            self.refresh_db_log_view()
            
            
def ensure_maximized(win):
    """A TÉNYLEGES méret alapján ellenőriz (nem az isMaximized() jelzés alapján, ami Windowson az ablak
    újraépítése után tévesen 'igaz' maradhat): ha az ablak jóval kisebb a képernyőnél, újra teljes méretre állítja."""
    try:
        if win.isMinimized() or win.isFullScreen():
            return
        scr = (win.screen() or QApplication.primaryScreen()).availableGeometry()
        if win.width() < scr.width() * 0.9 or win.height() < scr.height() * 0.85:
            win.showNormal()
            win.showMaximized()
    except Exception as e:
        print(f"[Ablak] Teljes méretre állítási hiba: {e}")


if __name__ == "__main__":
    # A beépített böngésző (Raktártérkép) ajánlott beállítása - a QApplication létrehozása ELŐTT kell
    QApplication.setAttribute(Qt.ApplicationAttribute.AA_ShareOpenGLContexts)
    app = QApplication(sys.argv)
    window = RaktarApp()
    # MINDIG TELJES MÉRETBEN indul. Windowson az első megjelenítéskor a teljes méret néha elveszik
    # (miközben a tartalom még töltődik), ezért kétszer is ellenőrizzük és szükség esetén újra beállítjuk.
    window.showMaximized()
    for delay in (300, 1500, 4000):
        QTimer.singleShot(delay, lambda: ensure_maximized(window))
    sys.exit(app.exec())