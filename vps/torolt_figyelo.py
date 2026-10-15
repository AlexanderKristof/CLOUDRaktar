#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# FIGYELEM: a fajl szandekosan ekezet nelkuli (ASCII), hogy a WinSCP szerkesztoje se rontsa el.
"""
TOROLT SOROK FIGYELO (VPS szolgaltatas) - /opt/raktarapp/torolt_figyelo.py

Folyamatosan (ejjel-nappal) figyeli a FIGYELT (elo) Google tablazatot. Ha onnan eltunik egy
R18-as termek sora, vagy kiuritik a Marka / Tipus / Cikkszam mezoit, a termek KORABBI teljes
adatai (A-P) a ToroltSor munkalap vegere kerulnek (Q: torles ideje, R: honnan, S: ki torolte).

KI TOROLTE (S oszlop): a Google Drive verzioelozmenyeibol (revisions) olvassa ki, ki szerkesztette a
tablazatot a torles idejeben - szinten CSAK OLVASASI joggal (drive.metadata.readonly). A Google a
verziokat percekre osszevonja, ezert ha ugyanakkor tobben szerkesztettek, mindegyikuk email-cime bekerul.
Ehhez a Google Cloud projektben a "Google Drive API"-t engedelyezni kell (lasd TELEPITES_torolt_figyelo.md).

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
import re
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

# "Ki torolte" keresese a verzioelozmenyekben. A Google a verziot nehany perc keslessel rogziti,
# ezert legfeljebb ennyi ideig varunk vele a ToroltSor-ba irassal (masodperc). Utana "nem allapithato meg".
WHO_WAIT_MAX_SEC = 600
# A torles eszlelese elotti / utani tures a verziok idejehez (masodperc)
WHO_SLACK_BEFORE_SEC = 120
WHO_SLACK_AFTER_SEC = 180

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
          "Q: T\u00f6rl\u00e9s ideje", "R: Honnan t\u00f6r\u00f6lt\u00e9k", "S: Ki t\u00f6r\u00f6lte (email)"]

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


_CODE_RE = re.compile(r"^\s*R18\s*[-/ ]?\s*0*(\d+)\s*$", re.IGNORECASE)


def normalize_code(code):
    """R18-123 / R18/123 / R18 123 -> 'R18-123' (osszehasonlitashoz). Nem R18-kod -> None."""
    m = _CODE_RE.match(str(code or ""))
    return f"R18-{int(m.group(1))}" if m else None


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
    if len(ws.row_values(1)) < len(HEADER):
        ws.update(range_name="A1:S1", values=[HEADER])
    return ws


# -------------------------------------------------------------------
# KI TOROLTE: Google Drive verzioelozmenyek (CSAK OLVASAS)
# -------------------------------------------------------------------
_who_warned = set()


def _utc(dt_str):
    return datetime.datetime.strptime(dt_str[:19], "%Y-%m-%dT%H:%M:%S").replace(tzinfo=datetime.timezone.utc)


def list_revisions():
    """A figyelt tablazat verzioi: [(ido UTC, email vagy nev), ...]. Csak olvasas (drive.metadata.readonly)."""
    from google.oauth2.service_account import Credentials
    from google.auth.transport.requests import AuthorizedSession
    creds = Credentials.from_service_account_file(
        CREDENTIALS_FILE, scopes=["https://www.googleapis.com/auth/drive.metadata.readonly"])
    sess = AuthorizedSession(creds)
    url = f"https://www.googleapis.com/drive/v3/files/{WATCH_SPREADSHEET_ID.strip()}/revisions"
    out, token = [], None
    while True:
        params = {"pageSize": 1000, "fields": "nextPageToken,revisions(modifiedTime,lastModifyingUser(displayName,emailAddress))"}
        if token:
            params["pageToken"] = token
        res = sess.get(url, params=params, timeout=30)
        if res.status_code != 200:
            raise RuntimeError(f"HTTP {res.status_code}: {res.text[:300]}")
        js = res.json()
        for r in js.get("revisions", []):
            u = r.get("lastModifyingUser") or {}
            who = u.get("emailAddress") or u.get("displayName") or ""
            out.append((_utc(r["modifiedTime"]), who))
        token = js.get("nextPageToken")
        if not token:
            return out


def find_who(win_from, win_to, revisions):
    """A [win_from, win_to] idoablakban (UTC) szerkeszto felhasznalok, vesszovel. Ures, ha nincs ilyen verzio."""
    a = win_from - datetime.timedelta(seconds=WHO_SLACK_BEFORE_SEC)
    b = win_to + datetime.timedelta(seconds=WHO_SLACK_AFTER_SEC)
    who = []
    for t, w in revisions:
        if a <= t <= b and w and w not in who:
            who.append(w)
    return ", ".join(who)


_last_check_utc = None


def check_once(conn):
    """Egy ellenorzesi kor: osszeveti a figyelt tablazatot az elozo pillanatkeppel."""
    global _last_check_utc
    check_utc = datetime.datetime.now(datetime.timezone.utc)
    values = read_watched_values()
    current = {}
    for r in values:
        r = (list(r) + [""] * 16)[:16]
        code = normalize_code(r[0])
        if code:
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
        win_from = _last_check_utc or (check_utc - datetime.timedelta(seconds=WATCH_INTERVAL_SEC * 2))
        for old_row, forras in gone:
            vals = [old_row.get(c, "") for c in ROW_COLS] + [now, forras]
            payload = {"vals": vals, "from": win_from.isoformat(), "to": check_utc.isoformat()}
            conn.execute("INSERT INTO varakozo (payload, created_at) VALUES (?, ?)",
                         (json.dumps(payload, ensure_ascii=False), now))
            log.info(f"{forras}: {old_row.get('A')} | {old_row.get('C')} | {old_row.get('D')} | {old_row.get('E')}")

    _last_check_utc = check_utc
    conn.execute("DELETE FROM snapshot WHERE forras = ?", (key,))
    conn.executemany("INSERT INTO snapshot (forras, code, payload) VALUES (?, ?, ?)",
                     [(key, c, json.dumps(r, ensure_ascii=False)) for c, r in current.items()])
    conn.commit()


def flush_pending(conn):
    """A varakozo bejegyzeseket kiirja a ToroltSor munkalapra (ami mar ott van, azt nem irja be ujra).
    Elotte megkeresi, ki torolte (verzioelozmenyek); ha meg nincs verzio a torles idejerol, kicsit var vele."""
    pending = conn.execute("SELECT id, payload, created_at FROM varakozo ORDER BY id").fetchall()
    if not pending:
        return
    revisions, rev_error = None, ""
    try:
        revisions = list_revisions()
    except Exception as e:
        rev_error = str(e)
        if rev_error[:60] not in _who_warned:
            _who_warned.add(rev_error[:60])
            log.warning("A 'ki torolte' nem olvashato a verzioelozmenyekbol (a ToroltSor ettol meg kiirodik). "
                        "Kell: Google Drive API engedelyezve + a szolgaltatasfiok lassa a tablazatot. "
                        f"Hiba: {rev_error[:300]}")
    now_utc = datetime.datetime.now(datetime.timezone.utc)
    ws, existing, written = None, None, 0
    for row in pending:
        data = json.loads(row["payload"])
        if isinstance(data, list):                      # regi formatum (ki torolte nelkul)
            vals, who = data, ""
        else:
            vals = data["vals"]
            w_from, w_to = datetime.datetime.fromisoformat(data["from"]), datetime.datetime.fromisoformat(data["to"])
            who = find_who(w_from, w_to, revisions) if revisions is not None else ""
            age = (now_utc - w_to).total_seconds()
            if not who and revisions is not None and age < WHO_WAIT_MAX_SEC:
                continue                                # a Google meg nem rogzitette a verziot - kesobb ujra
            if not who:
                who = "(nem allapithato meg)" if not rev_error else "(verzioelozmeny nem olvashato)"
        if ws is None:
            ws = get_deleted_ws()
            existing = {tuple((list(r) + [""] * 16)[:16]) for r in ws.get_all_values()[1:]}
        vals = (list(vals) + [""] * 18)[:18] + [who]
        if tuple(vals[:16]) not in existing:
            ws.append_row(vals, value_input_option="RAW")
            existing.add(tuple(vals[:16]))
            written += 1
            log.info(f"ToroltSor: {vals[0]} | torolte: {who or '-'}")
        conn.execute("DELETE FROM varakozo WHERE id = ?", (row["id"],))
        conn.commit()
    if written:
        log.info(f"{written} bejegyzes kiirva a '{DELETED_WORKSHEET}' munkalapra.")


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
