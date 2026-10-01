#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
TÖRÖLT SOROK FIGYELŐ (VPS szolgáltatás) - /opt/raktarapp/torolt_figyelo.py

Folyamatosan (éjjel-nappal) figyeli a FIGYELT (élő) Google táblázatot. Ha onnan eltűnik egy
R18-as termék sora, vagy kiürítik a Márka / Típus / Cikkszám mezőit, a termék KORÁBBI teljes
adatai (A-P) a ToroltSor munkalap végére kerülnek (Q: törlés ideje, R: honnan).

BIZTONSÁG:
- A figyelt táblázathoz KIZÁRÓLAG olvasási joggal kapcsolódik (spreadsheets.readonly),
  így semmilyen módon nem tudja módosítani.
- A ToroltSor nem lehet a figyelt táblázatban (a program ilyenkor nem indul el).
- Hibás / üres letöltésnél, vagy ha egyszerre túl sok termék tűnik el, nem archivál.
- Ha nincs net / a Google nem elérhető, a ToroltSor-bejegyzések helyben várnak és később kimennek.

Indítás kézzel (teszt):   /opt/raktarapp/venv/bin/python /opt/raktarapp/torolt_figyelo.py --egyszer
Szolgáltatásként:         lásd torolt-figyelo.service és TELEPITES_torolt_figyelo.md
"""
import datetime
import json
import logging
import os
import sqlite3
import sys
import time
from logging.handlers import RotatingFileHandler

# -------------------------------------------------------------------
# BEÁLLÍTÁSOK - ITT KELL ÁTÍRNI
# -------------------------------------------------------------------
# A figyelt (élő) táblázat ID-ja: a docs.google.com/spreadsheets/d/<EZ>/edit részből
WATCH_SPREADSHEET_ID = ""
# A figyelt munkalap (fül) neve - üresen hagyva a táblázat ELSŐ munkalapja
WATCH_WORKSHEET = ""
# Milyen gyakran nézze meg (másodperc). A Google percenként kb. 60 olvasást enged szolgáltatásfióknak.
WATCH_INTERVAL_SEC = 30

# A ToroltSor munkalap helye (ide ÍR) - NEM lehet ugyanaz a táblázat, mint a figyelt!
DELETED_SPREADSHEET_ID = "1ZoB2yl_GvaY0kiICQN10dEKghAE7b9Z3hDlwtv9nMdc"
DELETED_WORKSHEET = "ToroltSor"

# Ha egy ellenőrzéskor ennél több kitöltött termék tűnik el egyszerre, nem archivál (valószínűleg hiba)
MISSING_SAFETY_LIMIT = 20

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
CREDENTIALS_FILE = os.path.join(BASE_DIR, "credentials.json")
# Saját kis adatbázis (pillanatkép + várakozó bejegyzések) - NEM a raktar_adatbazis.db
STATE_DB = os.path.join(BASE_DIR, "torolt_figyelo.db")
LOG_FILE = os.path.join(BASE_DIR, "naplok", "torolt_figyelo.log")
# -------------------------------------------------------------------

ROW_COLS = "ABCDEFGHIJKLMNOP"
HEADER = ["A: F-Kód", "B: Vonalkód", "C: Márka", "D: Típus", "E: Cikkszám", "F: DB",
          "G: Állapot", "H: Ár", "I: Páter", "J: Szabad", "K: Polc", "L: Raklap",
          "M: Súly", "N: eBay ID", "O: Megjegyzés", "P: Állapot részlet",
          "Q: Törlés ideje", "R: Honnan törölték"]

log = logging.getLogger("torolt_figyelo")


def setup_logging():
    os.makedirs(os.path.dirname(LOG_FILE), exist_ok=True)
    fmt = logging.Formatter("%(asctime)s [%(levelname)s] %(message)s")
    fh = RotatingFileHandler(LOG_FILE, maxBytes=2_000_000, backupCount=3, encoding="utf-8")
    fh.setFormatter(fmt)
    sh = logging.StreamHandler(sys.stdout)
    sh.setFormatter(fmt)
    log.setLevel(logging.INFO)
    log.addHandler(fh)
    log.addHandler(sh)


def db():
    conn = sqlite3.connect(STATE_DB, timeout=30)
    conn.row_factory = sqlite3.Row
    conn.execute("""CREATE TABLE IF NOT EXISTS snapshot (
                        forras TEXT NOT NULL, code TEXT NOT NULL, payload TEXT,
                        PRIMARY KEY (forras, code))""")
    conn.execute("""CREATE TABLE IF NOT EXISTS varakozo (
                        id INTEGER PRIMARY KEY AUTOINCREMENT, payload TEXT NOT NULL, created_at TEXT)""")
    return conn


def has_product_data(row):
    return any(str(row.get(c, "")).strip() for c in ("C", "D", "E"))


def _creds(readonly):
    from google.oauth2.service_account import Credentials
    scope = "https://www.googleapis.com/auth/spreadsheets" + (".readonly" if readonly else "")
    return Credentials.from_service_account_file(CREDENTIALS_FILE, scopes=[scope])


def read_watched_values():
    """Figyelt táblázat - CSAK OLVASÁS."""
    import gspread
    book = gspread.authorize(_creds(readonly=True)).open_by_key(WATCH_SPREADSHEET_ID.strip())
    ws = book.worksheet(WATCH_WORKSHEET) if WATCH_WORKSHEET.strip() else book.sheet1
    return ws.get_all_values()


def get_deleted_ws():
    import gspread
    book = gspread.authorize(_creds(readonly=False)).open_by_key(DELETED_SPREADSHEET_ID.strip())
    try:
        ws = book.worksheet(DELETED_WORKSHEET)
    except gspread.exceptions.WorksheetNotFound:
        ws = book.add_worksheet(title=DELETED_WORKSHEET, rows=1000, cols=len(HEADER))
    if not ws.row_values(1):
        ws.update(range_name="A1:R1", values=[HEADER])
    return ws


def check_once(conn):
    """Egy ellenőrzési kör: összeveti a figyelt táblázatot az előző pillanatképpel."""
    values = read_watched_values()
    current = {}
    for r in values:
        r = (list(r) + [""] * 16)[:16]
        code = str(r[0]).strip().upper()
        if code.startswith("R18-"):
            current[code] = {c: str(v).strip() for c, v in zip(ROW_COLS, r)}

    key = f"{WATCH_SPREADSHEET_ID.strip()}|{WATCH_WORKSHEET.strip()}"
    previous = {row["code"]: json.loads(row["payload"])
                for row in conn.execute("SELECT code, payload FROM snapshot WHERE forras = ?", (key,))}

    if not previous:
        log.info(f"Első ellenőrzés: {len(current)} R18-sor megjegyezve, innentől figyelem a változásokat.")
    elif len(current) < len(previous) * 0.5:
        log.warning(f"Csak {len(current)} sor jött le (előzőleg {len(previous)}) - valószínűleg hibás letöltés, kihagyva.")
        return

    gone = []
    for code, old_row in previous.items():
        if not has_product_data(old_row):
            continue
        new_row = current.get(code)
        if new_row is None:
            gone.append((old_row, "Élő tábla: sor törölve"))
        elif not has_product_data(new_row):
            gone.append((old_row, "Élő tábla: kiürítve"))

    if len(gone) > MISSING_SAFETY_LIMIT:
        log.warning(f"{len(gone)} termék tűnt el egyszerre - túl sok, automatikus archiválás kihagyva "
                    f"(ellenőrizd a táblázatot!). Kódok: {', '.join(r['A'] for r, _ in gone[:50])}")
    else:
        now = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        for old_row, forras in gone:
            vals = [old_row.get(c, "") for c in ROW_COLS] + [now, forras]
            conn.execute("INSERT INTO varakozo (payload, created_at) VALUES (?, ?)",
                         (json.dumps(vals, ensure_ascii=False), now))
            log.info(f"{forras}: {old_row.get('A')} | {old_row.get('C')} | {old_row.get('D')} | {old_row.get('E')}")

    conn.execute("DELETE FROM snapshot WHERE forras = ?", (key,))
    conn.executemany("INSERT INTO snapshot (forras, code, payload) VALUES (?, ?, ?)",
                     [(key, c, json.dumps(r, ensure_ascii=False)) for c, r in current.items()])
    conn.commit()


def flush_pending(conn):
    """A várakozó bejegyzéseket kiírja a ToroltSor munkalapra (ami már ott van, azt nem írja be újra)."""
    pending = conn.execute("SELECT id, payload FROM varakozo ORDER BY id").fetchall()
    if not pending:
        return
    ws = get_deleted_ws()
    existing = {tuple((list(r) + [""] * 16)[:16]) for r in ws.get_all_values()[1:]}
    for row in pending:
        vals = json.loads(row["payload"])
        if tuple(vals[:16]) not in existing:
            ws.append_row(vals, value_input_option="RAW")
            existing.add(tuple(vals[:16]))
        conn.execute("DELETE FROM varakozo WHERE id = ?", (row["id"],))
        conn.commit()
    log.info(f"{len(pending)} bejegyzés kiírva a '{DELETED_WORKSHEET}' munkalapra.")


def main():
    setup_logging()
    if not WATCH_SPREADSHEET_ID.strip():
        log.error("Nincs megadva a figyelt táblázat (WATCH_SPREADSHEET_ID) - a fájl elején kell beállítani.")
        sys.exit(2)
    if WATCH_SPREADSHEET_ID.strip() == DELETED_SPREADSHEET_ID.strip():
        log.error("A ToroltSor nem lehet a figyelt (élő) táblázatban! Állíts be másik DELETED_SPREADSHEET_ID-t.")
        sys.exit(2)

    once = "--egyszer" in sys.argv
    log.info(f"Indul a törölt sorok figyelése ({WATCH_INTERVAL_SEC} mp-enként), csak olvasás: {WATCH_SPREADSHEET_ID}")
    conn = db()
    while True:
        try:
            check_once(conn)
        except Exception as e:
            log.warning(f"Figyelt táblázat nem olvasható, később újra próbálom: {e}")
        try:
            flush_pending(conn)
        except Exception as e:
            log.warning(f"ToroltSor írási hiba, később újra próbálom: {e}")
        if once:
            break
        time.sleep(WATCH_INTERVAL_SEC)


if __name__ == "__main__":
    main()
