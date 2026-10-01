#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# FIGYELEM: a fajl szandekosan ekezet nelkuli (ASCII), hogy a WinSCP szerkesztoje se rontsa el.
"""
TOROLT SOROK FIGYELO (VPS szolgaltatas) - /opt/raktarapp/torolt_figyelo.py

Folyamatosan (ejjel-nappal) figyeli a FIGYELT (elo) Google tablazatot. Ha onnan eltunik egy
R18-as termek sora, vagy kiuritik a Marka / Tipus / Cikkszam mezoit, a termek KORABBI teljes
adatai (A-P) a ToroltSor munkalap vegere kerulnek (Q: torles ideje, R: honnan).

BIZTONSAG:
- A figyelt tablazathoz KIZAROLAG olvasasi joggal kapcsolodik (spreadsheets.readonly),
  igy semmilyen modon nem tudja modositani.
- A ToroltSor nem lehet a figyelt tablazatban (a program ilyenkor nem indul el).
- Hibas / ures letoltesnel, vagy ha egyszerre tul sok termek tunik el, nem archival.
- Ha nincs net / a Google nem elerheto, a ToroltSor-bejegyzesek helyben varnak es kesobb kimennek.

Inditas kezzel (teszt):   /opt/raktarapp/venv/bin/python /opt/raktarapp/torolt_figyelo.py --egyszer
Szolgaltataskent:         lasd torolt-figyelo.service es TELEPITES_torolt_figyelo.md
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
# BEALLITASOK - ITT KELL ATIRNI
# -------------------------------------------------------------------
# A figyelt (elo) tablazat ID-ja: a docs.google.com/spreadsheets/d/<EZ>/edit reszbol
WATCH_SPREADSHEET_ID = "12bvyXlrLWcndZIPxnWhoZ1EPeHKSOc_eK2F7t_fjTfg"
# A figyelt munkalap (ful) neve - uresen hagyva a tablazat ELSO munkalapja
WATCH_WORKSHEET = ""
# Milyen gyakran nezze meg (masodperc). A Google percenkent kb. 60 olvasast enged szolgaltatasfioknak.
WATCH_INTERVAL_SEC = 30

# A ToroltSor munkalap helye (ide IR) - NEM lehet ugyanaz a tablazat, mint a figyelt!
DELETED_SPREADSHEET_ID = "1ZoB2yl_GvaY0kiICQN10dEKghAE7b9Z3hDlwtv9nMdc"
DELETED_WORKSHEET = "ToroltSor"

# Ha egy ellenorzeskor ennel tobb kitoltott termek tunik el egyszerre, nem archival (valoszinuleg hiba)
MISSING_SAFETY_LIMIT = 20

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
CREDENTIALS_FILE = os.path.join(BASE_DIR, "credentials.json")
# Sajat kis adatbazis (pillanatkep + varakozo bejegyzesek) - NEM a raktar_adatbazis.db
STATE_DB = os.path.join(BASE_DIR, "torolt_figyelo.db")
LOG_FILE = os.path.join(BASE_DIR, "naplok", "torolt_figyelo.log")
# -------------------------------------------------------------------

ROW_COLS = "ABCDEFGHIJKLMNOP"
HEADER = ["A: F-K\u00f3d", "B: Vonalk\u00f3d", "C: M\u00e1rka", "D: T\u00edpus", "E: Cikksz\u00e1m", "F: DB",
          "G: \u00c1llapot", "H: \u00c1r", "I: P\u00e1ter", "J: Szabad", "K: Polc", "L: Raklap",
          "M: S\u00faly", "N: eBay ID", "O: Megjegyz\u00e9s", "P: \u00c1llapot r\u00e9szlet",
          "Q: T\u00f6rl\u00e9s ideje", "R: Honnan t\u00f6r\u00f6lt\u00e9k"]

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
    """Figyelt tablazat - CSAK OLVASAS."""
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
    """Egy ellenorzesi kor: osszeveti a figyelt tablazatot az elozo pillanatkeppel."""
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
        log.info(f"Elso ellenorzes: {len(current)} R18-sor megjegyezve, innentol figyelem a valtozasokat.")
    elif len(current) < len(previous) * 0.5:
        log.warning(f"Csak {len(current)} sor jott le (elozoleg {len(previous)}) - valoszinuleg hibas letoltes, kihagyva.")
        return

    gone = []
    for code, old_row in previous.items():
        if not has_product_data(old_row):
            continue
        new_row = current.get(code)
        if new_row is None:
            gone.append((old_row, "\u00c9l\u0151 t\u00e1bla: sor t\u00f6r\u00f6lve"))
        elif not has_product_data(new_row):
            gone.append((old_row, "\u00c9l\u0151 t\u00e1bla: ki\u00fcr\u00edtve"))

    if len(gone) > MISSING_SAFETY_LIMIT:
        log.warning(f"{len(gone)} termek tunt el egyszerre - tul sok, automatikus archivalas kihagyva "
                    f"(ellenorizd a tablazatot!). Kodok: {', '.join(r['A'] for r, _ in gone[:50])}")
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
    """A varakozo bejegyzeseket kiirja a ToroltSor munkalapra (ami mar ott van, azt nem irja be ujra)."""
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
    log.info(f"{len(pending)} bejegyzes kiirva a '{DELETED_WORKSHEET}' munkalapra.")


def main():
    setup_logging()
    if not WATCH_SPREADSHEET_ID.strip():
        log.error("Nincs megadva a figyelt tablazat (WATCH_SPREADSHEET_ID) - a fajl elejen kell beallitani.")
        sys.exit(2)
    if WATCH_SPREADSHEET_ID.strip() == DELETED_SPREADSHEET_ID.strip():
        log.error("A ToroltSor nem lehet a figyelt (elo) tablazatban! Allits be masik DELETED_SPREADSHEET_ID-t.")
        sys.exit(2)

    once = "--egyszer" in sys.argv
    log.info(f"Indul a torolt sorok figyelese ({WATCH_INTERVAL_SEC} mp-enkent), csak olvasas: {WATCH_SPREADSHEET_ID}")
    conn = db()
    while True:
        try:
            check_once(conn)
        except Exception as e:
            log.warning(f"Figyelt tablazat nem olvashato, kesobb ujra probalom: {e}")
        try:
            flush_pending(conn)
        except Exception as e:
            log.warning(f"ToroltSor irasi hiba, kesobb ujra probalom: {e}")
        if once:
            break
        time.sleep(WATCH_INTERVAL_SEC)


if __name__ == "__main__":
    main()
