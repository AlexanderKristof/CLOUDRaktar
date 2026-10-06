from PyQt6.QtCore import Qt, QThread, QObject, pyqtSignal, QUrl, QTimer, QPropertyAnimation, QDate, QEvent, QRectF
import sys
import io
import os
import re
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
    QStyledItemDelegate, QAbstractItemDelegate, QComboBox, QPlainTextEdit, QAbstractSpinBox, QStyle
)
from PyQt6.QtGui import QColor, QPixmap, QFont, QDesktopServices, QCloseEvent, QShortcut, QKeySequence, QIntValidator, QPainter
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

def send_to_vps_async(endpoint, data):
    """Adatok beküldése a VPS API-ra külön háttérszálon (Zéró fagyás)"""
    def _send():
        try:
            url = f"{VPS_URL}{endpoint}"
            requests.post(url, json=data, timeout=3)
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

def get_pending_outbox_codes(cursor):
    """Azok az R18-kódok, amelyeknek van még fel nem töltött módosítása az outboxban.
    Ezeket a letöltések (VPS / Sheets) NEM írhatják felül a régebbi felhős állapottal."""
    cursor.execute("SELECT payload FROM sync_outbox")
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
            res = requests.post(f"{VPS_URL}/delete_item", json={"A": code_a}, timeout=3)
        elif op == "color":
            res = requests.post(f"{VPS_URL}/update_color", json={"A": code_a, "szin": row_data.get("szin", "")}, timeout=3)
        else:
            body = {k: v for k, v in row_data.items() if not k.startswith("_")}
            if row_data.get("_force"):
                body["force"] = True
            res = requests.post(f"{VPS_URL}/update_item", json=body, timeout=3)
    except Exception:
        return "retry"

    if res.status_code == 200:
        return "ok"
    if res.status_code == 409 and op == "upsert":
        details = format_row_for_log(row_data)
        server_msg = ""
        try:
            server_msg = res.json().get("message", "")
        except Exception:
            pass
        existing = _server_conflict_existing(server_msg)
        new_row = {c: str(row_data.get(c, "")) for c in ROW_COLS}

        # 1) A szerveren csak egy FÉLIG kitöltött / üres sor van ezzel az R18-számmal (pl. a Márka már
        #    felment, a Típus / Cikkszám még nem): ez nem másik termék -> automatikusan kiegészítjük.
        if _is_fill_in_of(existing, row_data):
            try:
                body = {k: v for k, v in row_data.items() if not k.startswith("_")}
                body["force"] = True
                res2 = requests.post(f"{VPS_URL}/update_item", json=body, timeout=3)
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

    if conflicts:
        try:
            outbox_signal_bus.conflicts_detected.emit(conflicts)
        except Exception as e:
            print(f"[Outbox] Jelzési hiba: {e}")

    return conflicts

def safe_str(val):
    if val is None or pd.isna(val):
        return ""
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
            res = requests.get(f"{VPS_URL}/sync/status", timeout=3)
            if res.status_code == 200:
                self.status_signal.emit(True, "🟢 VPS Csatlakozva")
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
            db_full = {safe_str(row["A"]): {c: safe_str(row[c]) for c in ROW_COLS} for row in cursor.fetchall() if row["A"]}
            db_existing = db_full
            # A még fel nem töltött (outboxban váró) tételeket a Sheets régebbi állapota nem írhatja
            # felül - és egy helyben törölt tételt sem hozhat vissza, amíg a törlés ki nem ment.
            pending_outbox = get_pending_outbox_codes(cursor)

            update_data = []
            insert_data = []
            sheet_changes = []   # (kód, előtte, utána) a Napló fülre - a Google táblázatban közvetlenül módosítva

            for row in df_test.itertuples(index=False, name=None):
                code = safe_str(row[0]) if len(row) > 0 else ""
                if not code or not code.startswith("R18-"): 
                    continue

                vals = [safe_str(row[i]) if i < len(row) else "" for i in range(16)]
                while len(vals) < 16: 
                    vals.append("")

                cloud_c, cloud_d, cloud_e = vals[2], vals[3], vals[4]

                # ÜTKÖZÉS-VIZSGÁLAT: Ha a felhasználónál pending (szerkesztés alatti) módosítás van erre a sorra
                if code in self.pending_rows_data:
                    local_p = self.pending_rows_data[code]
                    # Ha a felhőben lévő adat eltér mind a helyitől, mind a szerkesztéstől, ütközés van!
                    if cloud_c and (cloud_c != local_p.get("C") or cloud_d != local_p.get("D")):
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

            if update_data:
                cursor.executemany("""
                    UPDATE termekek SET A=?, B=?, C=?, D=?, E=?, F=?, G=?, H=?, I=?, J=?, K=?, L=?, M=?, N=?, O=?, P=?
                    WHERE A=? OR B=?
                """, update_data)

            if insert_data:
                cursor.executemany("""
                    INSERT INTO termekek (A, B, C, D, E, F, G, H, I, J, K, L, M, N, O, P, szin)
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
                res = requests.get(f"{VPS_URL}/get_updates", params=params, timeout=2)
                if res.status_code != 200:
                    continue
                data = res.json()
                if data.get("items"):
                    self.data_updated_signal.emit(data["items"])
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
            res = requests.post(url, timeout=10)
            
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
                    hely = safe_str(r_dict.get("L")) or safe_str(r_dict.get("I")) or "Ismeretlen"
                    
                    color_val = safe_str(r_dict.get("szin")).upper()
                    c_val = safe_str(r_dict.get("C"))
                    is_sold = (color_val == "LILA") or (not c_val)

                    box = QFrame()
                    if is_sold:
                        box.setStyleSheet("QFrame { background-color: rgba(171, 71, 188, 0.25); border: 2px solid #ab47bc; border-radius: 6px; margin-bottom: 6px; padding: 8px; }")
                        info_text = f"🏷️ <b>{code}</b> | 🟣 <b style='color: #ab47bc;'>[ELADVA / TÖRÖLVE]</b> | 📍 Utolsó Hely: <b>{hely}</b>"
                    else:
                        box.setStyleSheet("QFrame { background-color: #1e1e1e; border: 1px solid #333; border-radius: 6px; margin-bottom: 6px; padding: 8px; }")
                        info_text = f"🏷️ <b>{code}</b> | 📦 Mennyiség: <b>{db} db</b> | 📍 Hely: <b>{hely}</b> | Állapot: <b>{allapot}</b> | Ár: <b>{ar} EUR</b>"

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


class ExcelCellDelegate(QStyledItemDelegate):
    """Cellaszerkesztő: a nyilak szerkesztés közben is átléptetnek a szomszéd cellába"""

    _ARROWS = (Qt.Key.Key_Left, Qt.Key.Key_Right, Qt.Key.Key_Up, Qt.Key.Key_Down)
    _MODS = (Qt.KeyboardModifier.ControlModifier | Qt.KeyboardModifier.AltModifier |
             Qt.KeyboardModifier.ShiftModifier | Qt.KeyboardModifier.MetaModifier)

    def createEditor(self, parent, option, index):
        editor = super().createEditor(parent, option, index)
        table = self.parent()
        if editor is not None:
            # gépeléssel indult a szerkesztés (nem dupla kattintással / F2-vel)?
            editor.setProperty("typing_mode", bool(getattr(table, "_typing_edit", False)))
        return editor

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

    def keyPressEvent(self, event):
        editing = self.state() == QAbstractItemView.State.EditingState

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
            items = self.selectedItems()
            if not items and self.currentItem() is not None:
                items = [self.currentItem()]

            deleted_group = []
            self._edit_source = "⌫ Delete / Backspace gomb (cella törlése)"
            try:
                for item in items:
                    if item.column() > 0 and item.text(): # A oszlop (F-Kód) védve marad
                        deleted_group.append((item.row(), item.column(), item.text(), ""))
                        item.setText("")
            finally:
                self._edit_source = None

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
def _normalize_excel_row(row):
    """Egy A-P adatbázis-sort az Excel Live táblában megjelenő szövegekké alakít (tuple, gyors összevetéshez)."""
    return tuple(str(v) if v is not None and str(v).lower() != "none" else "" for v in row[:16])


class RaktarApp(QMainWindow):
    excel_rows_signal = pyqtSignal(object)  # háttérszálon beolvasott Excel Live sorok -> fő szál
    server_status_signal = pyqtSignal(bool, str)
    counter_update_signal = pyqtSignal(str, int, list)
    process_log_signal = pyqtSignal(str)
    print_result_signal = pyqtSignal(str)   # nyomtatás eredménye (háttérszálból) -> toast

    def __init__(self):
        super().__init__()
        self.setWindowTitle("📦 Raktári Rendszer (PyQt6 Ultra-Fast)")
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
        
        # --- 1 MP-ES AUTOMATIKUS LÁTHATATLAN VPS SZINKRON INDÍTÁSA ---
        self.auto_vps_worker = AutoVPSPollWorker()
        self.auto_vps_worker.data_updated_signal.connect(self.on_auto_vps_data_received)
        self.auto_vps_worker.start()

        self.close_btn = QPushButton("🚪 Bezárás")
        self.close_btn.setStyleSheet("background-color: #ff5252; color: white;")
        self.close_btn.clicked.connect(self.close)
        header_layout.addWidget(self.close_btn)

        main_layout.addLayout(header_layout)

        # FÜLEK LÉTREHOZÁSA ÉS SORRENDJE (Excel Live Munkalap a legelső!)
        self.tabs = QTabWidget()
        main_layout.addWidget(self.tabs)

        # 1. Először létrehozzuk az Excel Live Munkalap fület legelöl
        self.setup_excel_style_tab()

        # 2. Utána hozzuk létre a többi fület
        self.tab_search = QWidget()
        self.tab_add = QWidget()
        self.tab_local_db = QWidget()
        self.tab_pda = QWidget()
        self.tab_sold = QWidget()
        self.tab_excel = QWidget()
        self.tab_single = QWidget()
        self.tab_map = QWidget()
        self.tab_log = QWidget()
        self.tab_process_log = QWidget()

        # Fülek hozzáadása a kívánt sorrendben
        self.tabs.addTab(self.tab_search, "⚡ Gyorskereső")
        self.tabs.addTab(self.tab_add, "📝 Új Termék Felvitele")
        self.tabs.addTab(self.tab_local_db, "💾 Helyi Adatbázis (Élő)")
        self.tabs.addTab(self.tab_pda, "📱 PDA / Eladás")
        self.tabs.addTab(self.tab_sold, "💰 Eladottak")
        self.tabs.addTab(self.tab_excel, "📊 Excel Kezelés")
        self.tabs.addTab(self.tab_single, "🖨️ Egyedi Címke")
        self.tabs.addTab(self.tab_map, "🗺️ Raktártérkép")
        self.tabs.addTab(self.tab_log, "📜 Napló")
        self.tabs.addTab(self.tab_process_log, "⚙️ Folyamat & Szinkron Napló")

        # Fülek tartalmának inicializálása
        self.init_search_tab()
        self.init_add_tab()
        self.init_local_db_tab()
        self.init_pda_tab()
        self.init_sold_tab()
        self.init_excel_tab()
        self.init_single_tab()
        self.init_map_tab()
        self.init_log_tab()
        self.init_process_log_tab()

        # GARANTÁLTAN AZ 1. FÜLRE (Excel Live Munkalap) FÓKUSZÁL INDÍTÁSKOR:
        self.tabs.setCurrentIndex(0)
        # Elhalasztott (nem látható fülre szóló) frissítések elvégzése fülváltáskor
        self.tabs.currentChanged.connect(self.on_tab_changed)

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
            pending_codes = get_pending_outbox_codes(cursor)

            cols = list(ROW_COLS) + ["szin"]
            changed = 0
            naplo_entries = []   # a tranzakció lezárása után írjuk a naplóba
            for item in updated_items:
                code = (safe_str(item.get("A")) or safe_str(item.get("B"))).upper()
                if not code or code in pending_codes:
                    continue

                # A másik gépen törölt (eladott) tétel
                if item.get("_torolt"):
                    elotte = get_row_snapshot(code, cursor)
                    cursor.execute("DELETE FROM termekek WHERE A=? OR B=?", (code, code))
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
                    f"SELECT rowid, {', '.join(cols)} FROM termekek WHERE A=? OR B=?", (code, code)
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
                        f"INSERT INTO termekek ({', '.join(cols)}) VALUES ({', '.join('?' * len(cols))})",
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
                SELECT rowid, * FROM termekek 
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
        if success:
            # HA ÜTKÖZÉS VAN A GOOGLE SHEETS ÉS A TE BEÍRÁSOD KÖZÖTT:
            if conflicts:
                conflict_codes = ", ".join(conflicts)
                QMessageBox.warning(
                    self,
                    "⚠️ FIGYELEM: SORKONFLIKTUS / PÁRHUZAMOS MÓDOSÍTÁS!",
                    f"Valaki más is módosította a Google Sheets-ben az alábbi sor(oka)t, miközben te gépeltél:\n\n"
                    f"📌 Érintett kódok: {conflict_codes}\n\n"
                    f"Kérlek ellenőrizd ezeket a sorokat a felülírás elkerülése érdekében!"
                )
                self.trigger_toast(f"⚠️ Sorkonfliktus detektálva: {conflict_codes}!")
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
        # Ha a felhasználó éppen aktívan gépel egy cellában, azt a sort NEM bántjuk
        active_dirty_rows = getattr(self, 'pending_dirty_rows', set())
        old = self._excel_rows_cache

        self.excel_table.blockSignals(True)
        if len(rows) != self.excel_table.rowCount():
            self.excel_table.setRowCount(len(rows))

        new_cache = list(rows)
        for r_idx, row in enumerate(rows):
            if r_idx < len(old) and old[r_idx] == row:
                continue
            if r_idx in active_dirty_rows or r_idx == current_row:
                new_cache[r_idx] = None   # következő körben újra összeveti
                continue
            for c_idx in range(16):
                val = row[c_idx]
                item = self.excel_table.item(r_idx, c_idx)
                if item:
                    if item.text() != val:
                        item.setText(val)
                else:
                    self.excel_table.setItem(r_idx, c_idx, QTableWidgetItem(val))
        self.excel_table.blockSignals(False)

        if len(new_cache) != len(old) or any(a is None or a[0] != (b[0] if b else None) for a, b in zip(new_cache, old)):
            self._excel_rows_cache = new_cache
            self._rebuild_excel_code_index()
        else:
            self._excel_rows_cache = new_cache

    def update_print_selected_label(self, row, *_):
        if not hasattr(self, "btn_print_selected"):
            return
        item = self.excel_table.item(row, 0) if row >= 0 else None
        code = item.text().strip() if item and item.text().strip() else "–"
        self.btn_print_selected.setText(f"🖨️ Nyomtatás: {code} (Ctrl+P)")

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

    def trigger_toast(self, message):
        toast = ToastNotification(message, self)
        toast.show_toast(duration=2500)

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
        hely = safe_str(row.get('L')) or pater
        
        left_box = QWidget()
        left_layout = QVBoxLayout(left_box)
        left_layout.setContentsMargins(0, 0, 0, 0)
        left_layout.setSpacing(8)

        row1 = QHBoxLayout()
        lbl_code = QLabel(f"🏷️ <b>KÓD:</b> {code}")
        lbl_code.setStyleSheet("font-size: 15px; font-weight: bold; color: #fff;")
        row1.addWidget(lbl_code)

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

        row4.addWidget(QLabel("Polc (K):"))
        inp_polc = QLineEdit(str(polc or ""))
        inp_polc.setAlignment(Qt.AlignmentFlag.AlignRight)
        inp_polc.setPlaceholderText("Polc száma")
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
                SELECT rowid, * FROM termekek 
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
        query_text = self.search_input.text().strip().replace("ü", "-").replace("Ü", "-")
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
        df = pd.read_sql_query("SELECT rowid, * FROM termekek WHERE C IS NOT NULL AND C != ''", conn)
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
            cursor.execute("SELECT rowid FROM termekek WHERE A=? OR B=?", (code_final, code_final))
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

    def record_sale(self, platform, product, stock, v):
        code = product.get("A", "").strip().upper()
        eladas_id = f"E{datetime.datetime.now().strftime('%Y%m%d%H%M%S')}-{code}-{os.urandom(2).hex()}"
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
        except Exception as e:
            self.trigger_toast(f"❌ Eladás mentési hiba: {e}")
            return

        cim = SALES_SHEETS[platform]["cim"]
        sale_details = (f"{v['darab']} db | Vevő: {v['vevo'] or '(később)'} | Ár: {v['ar'] or '(később)'} | "
                        f"Tracking: {v['tracking'] or '(később)'}")
        elotte = get_row_snapshot(code) or {c: product.get(c, "") for c in NAPLO_COLS}
        sale_src = f"PDA / Eladás fül - 💰 {cim} gomb"

        # Készlet levonása
        stock_msg = ""
        if v["reduce_stock"]:
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
        else:
            log_event(code, f"ELADÁS: {cim}", sale_details + " | Készlet nem változott (a levonás ki volt kapcsolva)",
                      elotte, elotte, sale_src)

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

        btn_sync = QPushButton("☁️ Szinkron most")
        btn_sync.setStyleSheet("background-color: #00e676; color: black; font-weight: bold;")
        btn_sync.clicked.connect(lambda: (sync_sales_async(), self.trigger_toast("☁️ Eladások feltöltése elindítva...")))
        bar.addWidget(btn_sync)

        btn_del = QPushButton("🗑️ Kijelölt eladás törlése")
        btn_del.setStyleSheet("background-color: #ff5252; color: white; font-weight: bold;")
        btn_del.clicked.connect(self.delete_selected_sale)
        bar.addWidget(btn_del)
        layout.addLayout(bar)

        self.sold_status_lbl = QLabel("")
        self.sold_status_lbl.setStyleSheet("color: #ffb74d; font-weight: bold;")
        layout.addWidget(self.sold_status_lbl)

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

    def execute_pda_search(self):
        query_text = self.pda_input.text().strip()
        if not query_text: return
        clean_query = clean_part_number(query_text)

        while self.pda_results_layout.count():
            item = self.pda_results_layout.takeAt(0)
            widget = item.widget()
            if widget:
                widget.deleteLater()

        conn = get_db_connection()
        df = pd.read_sql_query("SELECT rowid, * FROM termekek WHERE C IS NOT NULL AND C != ''", conn)
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

        try:
            self.web_view = QWebEngineView()
            self.web_view.setPage(SilentWebEnginePage(self.web_view))
            embed_url = f"https://docs.google.com/spreadsheets/d/{MAP_SHEETS_ID}/htmlembed?widget=true&headers=false"
            self.web_view.setUrl(QUrl(embed_url))
            layout.addWidget(self.web_view)
        except Exception:
            layout.addWidget(QLabel("🗺️ A beágyazott böngésző nem érhető el."))

    def search_on_map(self, *_, force_reload=False):
        """🔍 Raktártérkép kereső: a pontosan egyező cellá(ka)t színesen kiemeli és odagörget."""
        query = self.map_search_input.text().strip()
        if not query:
            self.map_search_status.setText("⚠️ Írj be egy keresett értéket!")
            return
        if not hasattr(self, "web_view"):
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
        r = self._db_log_rows[row]
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
        self.db_log_detail_lbl.setTextFormat(Qt.TextFormat.RichText)
        self.db_log_detail_lbl.setText(head)

        self.db_log_detail.setRowCount(len(NAPLO_COLS))
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
                self.db_log_detail.setItem(i, ci, it)

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

    def _tab_is_current(self, tab_attr):
        """Igaz, ha a megadott fül látszik (a fülek létrehozása előtt mindig igaz)."""
        if not hasattr(self, "tabs") or not hasattr(self, tab_attr):
            return True
        return self.tabs.currentWidget() is getattr(self, tab_attr)

    def on_tab_changed(self, index):
        """Fülváltás: a háttérben elhalasztott frissítéseket csak most, a fül megjelenése UTÁN végezzük el."""
        w = self.tabs.widget(index)

        def later():
            if w is getattr(self, "tab_search", None) and getattr(self, "_cards_dirty", False):
                self.update_cards_only()
            elif w is getattr(self, "tab_local_db", None) and getattr(self, "_local_db_dirty", False):
                self.refresh_local_db_table_view()
            elif w is getattr(self, "tab_excel_live", None) and getattr(self, "_excel_dirty", False):
                self.silent_update_excel_table()
            elif w is getattr(self, "tab_sold", None):
                self.refresh_sold_table()
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
        ctrl_layout.addWidget(self.btn_print_selected)

        ctrl_layout.addStretch()
        layout.addWidget(ctrl_box)

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

        # Az Excel Live tábla aktuális tartalma soronként (gyors összehasonlításhoz a háttérfrissítésnél)
        self._excel_rows_cache = []
        self._excel_code_index = {}
        self._excel_loading = False
        self.excel_rows_signal.connect(self._apply_excel_rows)

        self._excel_row_sources = {}   # sor -> ["⌨️ Billentyűzet - L: '573' felülírva → '574'", ...] (Napló)
        self.excel_table.itemChanged.connect(self.on_excel_cell_changed)
        layout.addWidget(self.excel_table)

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
                self.excel_table.setItem(r_idx, c_idx, QTableWidgetItem(row[c_idx]))

        self.excel_table.blockSignals(False)
        self.excel_table.setUpdatesEnabled(True)
        self._excel_rows_cache = list(rows)
        self._rebuild_excel_code_index()

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
        elif len(rows) > 0:
            # Ha a pontos szám még nem létezik az adatbázisban, az utolsó létező sorra ugrik
            last_idx = len(rows) - 1
            target_item = self.excel_table.item(last_idx, 2)
            if target_item:
                self.excel_table.setCurrentItem(target_item)
                self.excel_table.scrollToItem(target_item, self.excel_table.ScrollHint.PositionAtCenter)
            
            
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

            # Napló: honnan / hogyan változott (a régi érték az utoljára betöltött adatbázis-állapotból)
            if not self.excel_table.signalsBlocked():
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
                if old_val != new_val:
                    self.excel_table.undo_stack.append([(row, col, old_val, new_val)])
                item._previous_text = new_val

    def print_row_label(self, row_idx):
        """🖨️ gomb: kattintás után bekéri a példányszámot (alap: 1, Enter = nyomtat), majd ZPL-lel nyomtat.
        Shift+kattintás: részletes nyomtatási ablak (PNG mód / előnézet)."""
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
                self.print_result_signal.emit(f"🖨️ Címke kinyomtatva [{code_u}] - {copies} db" if ok else f"❌ Nyomtatási hiba [{code_u}]: {msg}")
                log_event(code_u, "CÍMKE NYOMTATÁS", f"Excel Live gomb - OK ({copies} db)" if ok else f"HIBA: {msg}")

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
        if hasattr(self, 'tabs') and self.tabs.currentIndex() == 1:
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
            
            
if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = RaktarApp()
    window.setWindowState(Qt.WindowState.WindowMaximized)
    window.show()
    sys.exit(app.exec())