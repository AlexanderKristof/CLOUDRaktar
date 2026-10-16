# -*- coding: utf-8 -*-
# FIGYELEM: a fajl szandekosan ekezet nelkuli (ASCII), hogy a WinSCP szerkesztoje se rontsa el.
import os
import re
import threading
import sqlite3
import datetime
import time
import requests
import io
import json
import pandas as pd
from flask import Flask, jsonify, request, send_file

app = Flask(__name__)

# ==========================================
# JELSZO (TOKEN) VEDELEM
# A token a szerveren egy fajlban van (NEM a kodban, mert a kod a GitHubon van):
#   /opt/raktarapp/api_token.txt   (egyetlen sor)
# Minden keres fejleceben ott kell lennie:  X-Raktar-Token: <token>
# Ha a fajl nem letezik / ures, a vedelem KI van kapcsolva (mint eddig).
# ==========================================
import hmac

API_TOKEN_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "api_token.txt")
# Jelszo nelkul elerheto utvonalak:
#   /pda           - a PDA app felulete (csak a kepernyo, adat nincs benne; a token a PDA-n van)
#   /pda/app.apk   - a PDA app telepitoje (titkot nem tartalmaz)
#   /add_munkalap  - a regi Flutter munkalap app (az nem tud tokent kuldeni)
OPEN_PATHS = {"/pda", "/pda/app.apk", "/add_munkalap"}


def load_api_token():
    try:
        with open(API_TOKEN_FILE, "r", encoding="utf-8") as f:
            return f.read().strip()
    except Exception:
        return ""


API_TOKEN = load_api_token()
if not API_TOKEN:
    print("[VPS] FIGYELEM: nincs api_token.txt - a szerver JELSZO NELKUL fogad mindent!")


@app.before_request
def check_token():
    if not API_TOKEN or request.path in OPEN_PATHS or request.method == "OPTIONS":
        return None
    # A VPS-en futo sajat scriptek (127.0.0.1-rol) jelszo nelkul is hivhatjak
    if request.remote_addr in ("127.0.0.1", "::1"):
        return None
    # A kepeknel (<img src=...>) a bongeszo nem tud fejlecet kuldeni, ott ?t=<token> jon
    sent = request.headers.get("X-Raktar-Token", "") or request.args.get("t", "")
    if sent and hmac.compare_digest(sent.encode("utf-8"), API_TOKEN.encode("utf-8")):
        return None
    return jsonify({"status": "error", "kod": "S-401", "message": "Hibas vagy hianyzo jelszo (token)"}), 401

DB_PATH = '/opt/raktarapp/raktar_adatbazis.db'
def init_munkalap_db():
    try:
        conn = sqlite3.connect(DB_PATH)
        cursor = conn.cursor()
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS munkalapok (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                user_email TEXT,
                dri_szam TEXT,
                ugyfelnev TEXT,
                helyszin TEXT,
                gep_tipus TEXT,
                idotartam TEXT,
                km TEXT,
                elvegzett_munka TEXT,
                hiba_leiras TEXT,
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
            )
        """)
        conn.commit()
        conn.close()
    except Exception as e:
        print(f"[VPS] Munkalap DB inicializalasi hiba: {e}")

# Minden termekek-modositas egy novekvo verzioszamot kap. A kliensek (01, 02 gep)
# a /get_updates?since=<verzio> vegponton csak az azota valtozott sorokat kerik le.
# (SQLite egyszerre csak egy irot enged, igy a MAX+1 mindig egyedi es novekvo.)
# A torolt tetelek a torolt_tetelek tablaba kerulnek (szinten verzioval), hogy a
# masik gep is ertesuljon a torlesrol - ezert a verzio a ket tabla kozos maximuma.
NEXT_VERZIO = """(SELECT MAX(v) + 1 FROM (
    SELECT COALESCE(MAX(verzio), 0) AS v FROM termekek
    UNION ALL SELECT COALESCE(MAX(verzio), 0) FROM torolt_tetelek))"""
MAX_UPDATES_PER_POLL = 500

def init_verzio_oszlop():
    try:
        conn = sqlite3.connect(DB_PATH)
        cursor = conn.cursor()
        cursor.execute("PRAGMA table_info(termekek)")
        cols = [c[1] for c in cursor.fetchall()]
        if cols and "verzio" not in cols:
            cursor.execute("ALTER TABLE termekek ADD COLUMN verzio INTEGER")
        if cols:
            cursor.execute("CREATE INDEX IF NOT EXISTS idx_termekek_verzio ON termekek(verzio)")
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS torolt_tetelek (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                kod TEXT NOT NULL,
                verzio INTEGER NOT NULL,
                torolve TIMESTAMP DEFAULT CURRENT_TIMESTAMP
            )
        """)
        cursor.execute("CREATE INDEX IF NOT EXISTS idx_torolt_verzio ON torolt_tetelek(verzio)")
        conn.commit()
        conn.close()
    except Exception as e:
        print(f"[VPS] Verzio oszlop inicializalasi hiba: {e}")

# Alkalmazas inditasakor lefut:
init_munkalap_db()
init_verzio_oszlop()
LIVE_SHEETS_ID = "1ZoB2yl_GvaY0kiICQN10dEKghAE7b9Z3hDlwtv9nMdc"

def safe_str(val):
    if val is None or pd.isna(val):
        return ""
    s = str(val).strip()
    if s.lower() in ("nan", "none", "null"):
        return ""
    if s.endswith(".0"):
        s = s[:-2]
    return s

# Masolaskor lathatatlan karakterek / "hosszu" kotojelek kerulhetnek a cellaba
_CODE_FIX = {0xfeff: None, 0x200b: None, 0x200c: None, 0x200d: None, 0x2060: None, 0xad: None,
             0x2010: "-", 0x2011: "-", 0x2012: "-", 0x2013: "-", 0x2014: "-", 0x2212: "-", 0xa0: " "}

def clean_code(code):
    return str(code or "").translate(_CODE_FIX).strip().upper()

def has_product_data(vals):
    return any(str(v).strip() for v in vals[2:5])

def run_background_sync():
    try:
        live_csv_url = f"https://docs.google.com/spreadsheets/d/{LIVE_SHEETS_ID}/export?format=csv"
        res = requests.get(live_csv_url, timeout=12)
        res.encoding = 'utf-8'

        if res.status_code == 200:
            df = pd.read_csv(io.StringIO(res.text), header=None, dtype=str, keep_default_na=False)
            conn = sqlite3.connect(DB_PATH)
            cursor = conn.cursor()
            changed = 0

            # Ha egy R18-szam tobbszor szerepel a tablazatban, az ADATOT tartalmazo sor szamit
            # (ugyanugy, mint a programban), kulonben a ket oldal mast latna.
            best_rows = {}
            for idx, row in df.iterrows():
                code = clean_code(safe_str(row.get(0)) or safe_str(row.get(1)))
                if not code.startswith("R18-"):
                    continue
                vals = [safe_str(row.get(i)) for i in range(16)]
                if clean_code(vals[0]) == code:
                    vals[0] = code
                prev = best_rows.get(code)
                if prev is not None and (has_product_data(prev) or not has_product_data(vals)):
                    continue
                best_rows[code] = vals

            for code, vals in best_rows.items():
                cursor.execute("""
                    SELECT rowid, A, B, C, D, E, F, G, H, I, J, K, L, M, N, O, P
                    FROM termekek WHERE A=? OR B=?
                """, (code, code))
                existing = cursor.fetchone()

                # Csak a tenylegesen valtozott sorok kapnak uj verziot, kulonben
                # minden teljes szinkron utan az osszes sort ujra letoltene minden kliens.
                if existing:
                    if [safe_str(v) for v in existing[1:]] == vals:
                        continue
                    cursor.execute(f"""
                        UPDATE termekek
                        SET A=?, B=?, C=?, D=?, E=?, F=?, G=?, H=?, I=?, J=?, K=?, L=?, M=?, N=?, O=?, P=?,
                            verzio={NEXT_VERZIO}
                        WHERE rowid=?
                    """, (*vals, existing[0]))
                else:
                    cursor.execute(f"""
                        INSERT INTO termekek (A, B, C, D, E, F, G, H, I, J, K, L, M, N, O, P, szin, verzio)
                        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 'FEHER', {NEXT_VERZIO})
                    """, vals)
                changed += 1

            conn.commit()
            conn.close()
            print(f"[VPS] Szinkronizacio sikeres! ({changed} sor valtozott)")
    except Exception as e:
        print(f"[VPS] Szinkronizacios hiba: {e}")

@app.route('/sync/status', methods=['GET'])
def status():
    return jsonify({"status": "ok", "message": "VPS Server is Online"}), 200

@app.route('/sync/live-to-test', methods=['POST'])
def sync_live_to_test():
    threading.Thread(target=run_background_sync, daemon=True).start()
    return jsonify({"status": "started", "message": "Szinkronizalas elindult"}), 200

@app.route('/sync/full', methods=['POST'])
def sync_full():
    threading.Thread(target=run_background_sync, daemon=True).start()
    return jsonify({"status": "started", "message": "Teljes szinkronizalas elindult"}), 200

# ==========================================
# AUTOMATIKUS SYNC VEGPONTOK (01 ES 02 GEP KOZOTTI ELO KAPCSOLAT)
# ==========================================

@app.route('/update_item', methods=['POST'])
def update_item():
    """Azonnali termekfelvitel / modositas elmentese a VPS SQLite adatbazisaba.
    UTKOZES-VEDELEM: ha az R18-szam mar letezik MAS termekkel (elteroo Marka/Tipus/Cikkszam),
    NEM irjuk felul - 409-et adunk vissza a torzsadatokkal egyutt, hogy a kliens naplozhassa."""
    try:
        data = request.json
        if not data or 'A' not in data:
            return jsonify({"status": "error", "message": "Hianyozo F-Kod (A)"}), 400

        code_a = str(data.get('A')).strip().upper()
        code_b = str(data.get('B', code_a)).strip().upper()

        conn = sqlite3.connect(DB_PATH)
        conn.row_factory = sqlite3.Row
        cursor = conn.cursor()

        cursor.execute("SELECT rowid, C, D, E FROM termekek WHERE A=? OR B=?", (code_a, code_a))
        existing = cursor.fetchone()

        if existing:
            old_c, old_d, old_e = safe_str(existing["C"]).upper(), safe_str(existing["D"]).upper(), safe_str(existing["E"]).upper()
            new_c, new_d, new_e = safe_str(data.get('C', '')).upper(), safe_str(data.get('D', '')).upper(), safe_str(data.get('E', '')).upper()

            has_old_identity = bool(old_c or old_d or old_e)
            has_new_identity = bool(new_c or new_d or new_e)
            identity_differs = has_old_identity and has_new_identity and (old_c, old_d, old_e) != (new_c, new_d, new_e)

            # 'force': szandekos javitas (kartya / Helyi DB ful szerkesztese) - nem utkozes
            if identity_differs and not data.get('force'):
                conn.close()
                return jsonify({
                    "status": "conflict",
                    "message": f"Az {code_a} R18-szam mar foglalt egy MASIK termekkel ({old_c} {old_d} {old_e}).",
                    "existing": {"C": existing["C"], "D": existing["D"], "E": existing["E"]}
                }), 409

        vals = (
            code_a, code_b, data.get('C', ''), data.get('D', ''),
            data.get('E', ''), data.get('F', 1), data.get('G', 'HASZNALT'),
            data.get('H', ''), data.get('I', ''), data.get('J', ''), data.get('K', ''),
            data.get('L', ''), data.get('M', ''), data.get('N', ''),
            data.get('O', ''), data.get('P', '')
        )

        if existing:
            cursor.execute(f"""
                UPDATE termekek
                SET A=?, B=?, C=?, D=?, E=?, F=?, G=?, H=?, I=?, J=?, K=?, L=?, M=?, N=?, O=?, P=?,
                    verzio={NEXT_VERZIO}
                WHERE rowid=?
            """, (*vals, existing["rowid"]))
        else:
            cursor.execute(f"""
                INSERT INTO termekek (A, B, C, D, E, F, G, H, I, J, K, L, M, N, O, P, szin, verzio)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 'FEHER', {NEXT_VERZIO})
            """, vals)

        conn.commit()
        conn.close()

        return jsonify({"status": "success", "message": f"{code_a} elmentve a VPS-en"}), 200

    except Exception as e:
        return jsonify({"status": "error", "message": str(e)}), 500


@app.route('/update_cell', methods=['POST'])
def update_cell():
    """Azonnali cellaszerkesztes (Pater, Polc, Hely, Suly) mentese a VPS-re"""
    try:
        data = request.json
        code = str(data.get('code')).strip().upper()
        col = str(data.get('column')).strip().upper()
        val = str(data.get('value', '')).strip()

        valid_cols = ["A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P"]
        if col not in valid_cols:
            return jsonify({"status": "error", "message": "Ervenytelen oszlop"}), 400

        conn = sqlite3.connect(DB_PATH)
        cursor = conn.cursor()
        
        query = f"UPDATE termekek SET {col} = ?, verzio = {NEXT_VERZIO} WHERE A = ? OR B = ?"
        cursor.execute(query, (val, code, code))

        conn.commit()
        conn.close()

        return jsonify({"status": "success", "message": f"{code} [{col}] frissitve a VPS-en"}), 200

    except Exception as e:
        return jsonify({"status": "error", "message": str(e)}), 500


@app.route('/update_color', methods=['POST'])
def update_color():
    """Szinjeloles (szin oszlop) mentese a VPS-re"""
    try:
        data = request.json or {}
        code = str(data.get('A', '')).strip().upper()
        szin = str(data.get('szin', '')).strip().upper()
        if not code:
            return jsonify({"status": "error", "message": "Hianyzo F-Kod (A)"}), 400

        conn = sqlite3.connect(DB_PATH)
        cursor = conn.cursor()
        cursor.execute(f"UPDATE termekek SET szin = ?, verzio = {NEXT_VERZIO} WHERE A = ? OR B = ?", (szin, code, code))
        conn.commit()
        conn.close()

        return jsonify({"status": "success", "message": f"{code} szine: {szin}"}), 200

    except Exception as e:
        return jsonify({"status": "error", "message": str(e)}), 500


@app.route('/delete_item', methods=['POST'])
def delete_item():
    """Termek torlese (eladas). A kod a torolt_tetelek tablaba kerul, hogy a masik gep is torolje."""
    try:
        data = request.json or {}
        code = str(data.get('A', '')).strip().upper()
        if not code:
            return jsonify({"status": "error", "message": "Hianyzo F-Kod (A)"}), 400

        conn = sqlite3.connect(DB_PATH)
        cursor = conn.cursor()
        # Elobb a torlesi bejegyzes (a verziot meg a torlendo sorral egyutt szamolja,
        # igy az biztosan nagyobb minden eddig kiadott verzional), utana a torles.
        cursor.execute(f"INSERT INTO torolt_tetelek (kod, verzio) VALUES (?, {NEXT_VERZIO})", (code,))
        cursor.execute("DELETE FROM termekek WHERE A = ? OR B = ?", (code, code))
        conn.commit()
        conn.close()

        return jsonify({"status": "success", "message": f"{code} torolve a VPS-rol"}), 200

    except Exception as e:
        return jsonify({"status": "error", "message": str(e)}), 500


@app.route('/get_updates', methods=['GET'])
def get_updates():
    """Az adott verzio OTA valtozott termekek (a kliensek 1 mp-enkent hivjak).
    'since' nelkul csak az aktualis verziot adja vissza (elso inditaskor ebbol indul a kliens)."""
    try:
        since = request.args.get('since', type=int)

        conn = sqlite3.connect(DB_PATH)
        conn.row_factory = sqlite3.Row
        cursor = conn.cursor()
        cursor.execute(f"SELECT {NEXT_VERZIO} - 1")
        current = cursor.fetchone()[0]

        # Elso lekeres, vagy a VPS adatbazisat visszaallitottak (a verzio visszaugrott)
        if since is None or since > current:
            conn.close()
            return jsonify({"status": "ok", "verzio": current, "items": []}), 200

        cursor.execute("""
            SELECT A, B, C, D, E, F, G, H, I, J, K, L, M, N, O, P, szin, verzio
            FROM termekek WHERE verzio > ? ORDER BY verzio LIMIT ?
        """, (since, MAX_UPDATES_PER_POLL))
        changed = []
        for row in cursor.fetchall():
            item = {k: row[k] for k in "ABCDEFGHIJKLMNOP"}
            item["szin"] = row["szin"]
            changed.append((row["verzio"], item))
        cursor.execute("""
            SELECT kod, verzio FROM torolt_tetelek WHERE verzio > ? ORDER BY verzio LIMIT ?
        """, (since, MAX_UPDATES_PER_POLL))
        deleted = [(row["verzio"], {"A": row["kod"], "_torolt": True}) for row in cursor.fetchall()]
        conn.close()

        # A modositasok es torlesek verzio szerinti sorrendben (egy torles utani ujrafelvitel is jo sorrendben jon)
        merged = sorted(changed + deleted, key=lambda x: x[0])[:MAX_UPDATES_PER_POLL]
        items = [item for _, item in merged]
        # Csak a tenylegesen visszaadott sorokig lepunk elore, igy egy epp
        # folyamatban levo (meg nem commitolt) iras sem vesz el.
        new_verzio = merged[-1][0] if merged else since

        return jsonify({"status": "ok", "verzio": new_verzio, "items": items}), 200

    except Exception as e:
        return jsonify({"status": "error", "message": str(e)}), 500


# ==========================================
# UJ: SZERVER ADATBAZIS FUL (asztali program) - CSAK OLVAS
# ==========================================
@app.route('/items', methods=['GET'])
def list_items():
    """A VPS adatbazis OSSZES termeke (a 'Szerver adatbazis' fulnek). Csak olvas, semmit nem modosit.
    Valasz: {"status": "ok", "verzio": <aktualis>, "items": [{A..P, szin, verzio}, ...],
             "torolt": [{"kod", "verzio", "torolve"}, ...]  (a legutobbi 300 torles)}"""
    try:
        conn = sqlite3.connect(DB_PATH)
        conn.row_factory = sqlite3.Row
        cursor = conn.cursor()
        cursor.execute(f"SELECT {NEXT_VERZIO} - 1")
        current = cursor.fetchone()[0]
        cursor.execute("SELECT A, B, C, D, E, F, G, H, I, J, K, L, M, N, O, P, szin, verzio FROM termekek")
        items = []
        for row in cursor.fetchall():
            item = {k: safe_str(row[k]) for k in "ABCDEFGHIJKLMNOP"}
            item["szin"] = safe_str(row["szin"])
            item["verzio"] = row["verzio"] or 0
            items.append(item)
        cursor.execute("SELECT kod, verzio, torolve FROM torolt_tetelek ORDER BY verzio DESC LIMIT 300")
        torolt = [{"kod": r["kod"], "verzio": r["verzio"], "torolve": r["torolve"]} for r in cursor.fetchall()]
        conn.close()
        return jsonify({"status": "ok", "verzio": current, "items": items, "torolt": torolt}), 200
    except Exception as e:
        return jsonify({"status": "error", "message": str(e)}), 500


@app.route('/add_munkalap', methods=['POST'])
def add_munkalap():
    """Flutter mobil appbol erkezo munkalap elmentese"""
    try:
        data = request.json
        if not data or 'dri_szam' not in data:
            return jsonify({"status": "error", "message": "Hianyzo Munkalap szam (dri_szam)"}), 400

        conn = sqlite3.connect(DB_PATH)
        cursor = conn.cursor()

        cursor.execute("""
            INSERT INTO munkalapok 
            (user_email, dri_szam, ugyfelnev, helyszin, gep_tipus, idotartam, km, elvegzett_munka, hiba_leiras)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
        """, (
            data.get('user_email', ''),
            data.get('dri_szam', ''),
            data.get('ugyfelnev', ''),
            data.get('helyszin', ''),
            data.get('gep_tipus', ''),
            data.get('idotartam', ''),
            data.get('km', ''),
            data.get('elvegzett_munka', ''),
            data.get('hiba_leiras', '')
        ))

        conn.commit()
        conn.close()

        return jsonify({"status": "success", "message": f"DRI-{data.get('dri_szam')} Munkalap elmentve a VPS-en"}), 200

    except Exception as e:
        return jsonify({"status": "error", "message": str(e)}), 500


# ==========================================
# PDA (Zebra TC27) - csippantas, kereses, eladas rogzitese
# A PDA app a /pda oldalt tolti be (pda.html, ugyanebben a mappaban).
# Az eladasokat a VPS eltarolja; az asztali program atveszi oket (/pda/sales/claim),
# o vonja le a keszletet es irja ki a Google Sheets eladasi lapjaira.
# ==========================================
PDA_HTML = os.path.join(os.path.dirname(os.path.abspath(__file__)), "pda.html")
PDA_PLATFORMS = ("EBAY_KIVULI", "EBAY", "WEBSHOP")
PDA_SALE_FIELDS = ("eladas_id", "datum", "platform", "A", "C", "D", "E", "F", "G", "L", "M", "N",
                   "vevo", "ar", "darab", "tracking", "megjegyzes", "keszlet_levonas", "eszkoz",
                   "atvette", "atvetel_ido", "eredmeny", "eredmeny_uzenet", "eredmeny_ido", "visszavonva")
PDA_CHANGE_FIELDS = ("valtozas_id", "ido", "tipus", "A", "C", "D", "E", "mezo", "regi", "uj", "megjegyzes",
                     "eszkoz", "atvette", "atvetel_ido", "eredmeny", "eredmeny_uzenet", "eredmeny_ido", "ref", "adat",
                     "visszavonva")
# A PDA-rol modosithato termek-mezok
PDA_EDITABLE = {"C": "Marka", "D": "Tipus", "E": "Cikkszam", "F": "DB", "G": "Allapot", "H": "Ar", "I": "Pater",
                "K": "Oszlop", "L": "Raklap", "M": "Suly", "N": "eBay ID", "O": "Megjegyzes", "P": "Allapot reszlet"}
# Az eladasbol a PDA-rol javithato mezok
PDA_SALE_EDITABLE = ("vevo", "ar", "tracking", "megjegyzes", "darab", "platform")


def pda_now():
    """Magyar ido (a VPS orajara UTC is lehet)."""
    try:
        from zoneinfo import ZoneInfo
        return datetime.datetime.now(ZoneInfo("Europe/Budapest")).replace(tzinfo=None)
    except Exception:
        return datetime.datetime.now()


def init_pda_db():
    try:
        conn = sqlite3.connect(DB_PATH)
        conn.execute("""
            CREATE TABLE IF NOT EXISTS pda_eladasok (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                eladas_id TEXT UNIQUE,
                datum TEXT, platform TEXT,
                A TEXT, C TEXT, D TEXT, E TEXT, F TEXT, G TEXT, L TEXT, M TEXT, N TEXT,
                vevo TEXT, ar TEXT, darab TEXT, tracking TEXT, megjegyzes TEXT,
                keszlet_levonas INTEGER NOT NULL DEFAULT 1,
                eszkoz TEXT,
                atvette TEXT,
                atvetel_ido TEXT,
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
            )
        """)
        conn.commit()
        conn.close()
    except Exception as e:
        print(f"[VPS] PDA tabla inicializalasi hiba: {e}")

init_pda_db()


def init_pda_db2():
    """Athelyezesek / modositasok sora, PDA naplo, es az eladasok eredmeny-oszlopai."""
    try:
        conn = sqlite3.connect(DB_PATH)
        cols = [c[1] for c in conn.execute("PRAGMA table_info(pda_eladasok)").fetchall()]
        for col in ("eredmeny", "eredmeny_uzenet", "eredmeny_ido", "visszavonva"):
            if col not in cols:
                conn.execute(f"ALTER TABLE pda_eladasok ADD COLUMN {col} TEXT")
        conn.execute("""
            CREATE TABLE IF NOT EXISTS pda_valtozasok (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                valtozas_id TEXT UNIQUE,
                ido TEXT, tipus TEXT, A TEXT,
                C TEXT, D TEXT, E TEXT,
                mezo TEXT, regi TEXT, uj TEXT,
                megjegyzes TEXT, eszkoz TEXT,
                atvette TEXT, atvetel_ido TEXT,
                eredmeny TEXT, eredmeny_uzenet TEXT, eredmeny_ido TEXT
            )
        """)
        vcols = [c[1] for c in conn.execute("PRAGMA table_info(pda_valtozasok)").fetchall()]
        for col in ("ref", "adat", "visszavonva"):
            if col not in vcols:
                conn.execute(f"ALTER TABLE pda_valtozasok ADD COLUMN {col} TEXT")
        # Az asztali gepek naplojanak masolata (a PDA "Elozmenyek" ablakahoz)
        conn.execute("""
            CREATE TABLE IF NOT EXISTS termek_naplo (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                gep TEXT, helyi_id INTEGER,
                idopont TEXT, r18_kod TEXT, esemeny TEXT, reszletek TEXT, forras TEXT,
                UNIQUE(gep, helyi_id)
            )
        """)
        conn.execute("CREATE INDEX IF NOT EXISTS idx_termek_naplo_kod ON termek_naplo(r18_kod)")
        conn.execute("""
            CREATE TABLE IF NOT EXISTS pda_naplo (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                ido TEXT, eszkoz TEXT, szint TEXT, esemeny TEXT, kod TEXT, reszletek TEXT
            )
        """)
        conn.commit()
        conn.close()
    except Exception as e:
        print(f"[VPS] PDA tabla (2) inicializalasi hiba: {e}")

init_pda_db2()


# ==========================================
# GEPEK ES VERZIOK: minden asztali program es PDA 20-30 mp-enkent jelentkezik
# ==========================================
def init_kliens_db():
    try:
        conn = sqlite3.connect(DB_PATH)
        conn.execute("""
            CREATE TABLE IF NOT EXISTS kliensek (
                azon TEXT PRIMARY KEY,
                tipus TEXT, nev TEXT, verzio TEXT, app_verzio TEXT,
                felhasznalo TEXT, ip TEXT, elso TEXT, utolso TEXT
            )
        """)
        conn.commit()
        conn.close()
    except Exception as e:
        print(f"[VPS] Kliens tabla inicializalasi hiba: {e}")

init_kliens_db()
ONLINE_SEC = 120


def _ver_key(v):
    """'v2026.10.16-pda2' -> osszehasonlithato kulcs (a szamokat szamkent kezeli)."""
    return [int(x) if x.isdigit() else x for x in re.split(r"(\d+)", str(v or ""))]


def _legfrissebb(cur):
    """A legujabb verzio tipusonkent (az elmult 60 napban jelentkezett gepek kozul)."""
    hatar = (pda_now() - datetime.timedelta(days=60)).strftime("%Y-%m-%d %H:%M:%S")
    out = {}
    for tipus, mezo in (("gep", "verzio"), ("pda", "verzio"), ("pda_app", "app_verzio")):
        t = "pda" if tipus.startswith("pda") else tipus
        cur.execute(f"SELECT {mezo} FROM kliensek WHERE tipus=? AND utolso>=? AND {mezo} IS NOT NULL AND {mezo}!=''", (t, hatar))
        vals = [r[0] for r in cur.fetchall()]
        out[tipus] = max(vals, key=_ver_key) if vals else ""
    return out


@app.route('/clients/heartbeat', methods=['POST'])
def clients_heartbeat():
    """{"tipus": "gep"|"pda", "nev", "verzio", "app_verzio", "felhasznalo"} - visszaadja a legfrissebb verziokat."""
    try:
        data = request.json or {}
        tipus = "pda" if str(data.get('tipus', '')).lower() == "pda" else "gep"
        nev = safe_str(data.get('nev', '')) or "ismeretlen"
        now = pda_now().strftime("%Y-%m-%d %H:%M:%S")
        conn = sqlite3.connect(DB_PATH)
        cur = conn.cursor()
        cur.execute("""INSERT INTO kliensek (azon, tipus, nev, verzio, app_verzio, felhasznalo, ip, elso, utolso)
                       VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
                       ON CONFLICT(azon) DO UPDATE SET verzio=excluded.verzio, app_verzio=excluded.app_verzio,
                           felhasznalo=excluded.felhasznalo, ip=excluded.ip, utolso=excluded.utolso""",
                    (f"{tipus}|{nev}", tipus, nev, safe_str(data.get('verzio', '')), safe_str(data.get('app_verzio', '')),
                     safe_str(data.get('felhasznalo', '')), request.remote_addr or "", now, now))
        conn.commit()
        leg = _legfrissebb(cur)
        conn.close()
        return jsonify({"status": "ok", "legfrissebb": leg}), 200
    except Exception as e:
        return jsonify({"status": "error", "message": str(e)}), 500


@app.route('/clients', methods=['GET'])
def clients_list():
    """Az osszes gep / PDA: nev, verzio, online-e, utoljara mikor jelentkezett."""
    try:
        conn = sqlite3.connect(DB_PATH)
        conn.row_factory = sqlite3.Row
        cur = conn.cursor()
        cur.execute("SELECT * FROM kliensek ORDER BY tipus, utolso DESC")
        now = pda_now()
        rows = []
        for r in cur.fetchall():
            try:
                mp = (now - datetime.datetime.strptime(r["utolso"], "%Y-%m-%d %H:%M:%S")).total_seconds()
            except Exception:
                mp = 10 ** 9
            row = {k: r[k] or "" for k in ("tipus", "nev", "verzio", "app_verzio", "felhasznalo", "ip", "elso", "utolso")}
            row.update({"online": mp <= ONLINE_SEC, "mp": int(mp)})
            rows.append(row)
        leg = _legfrissebb(cur)
        conn.close()
        return jsonify({"status": "ok", "most": now.strftime("%Y-%m-%d %H:%M:%S"), "legfrissebb": leg, "kliensek": rows}), 200
    except Exception as e:
        return jsonify({"status": "error", "message": str(e)}), 500


def pda_err(http, kod, message):
    """Hibavalasz hibakoddal (a PDA Folyamat naplojaban ez latszik)."""
    return jsonify({"status": "error", "kod": kod, "message": message}), http


def _clean_part(text):
    return re.sub(r"[\s\-_/.,]+", "", str(text or "")).upper()


def _r18_num(code):
    m = re.match(r"^\s*R18-0*(\d+)", str(code or ""), re.IGNORECASE)
    return int(m.group(1)) if m else -1


def _pda_row(row):
    return {k: safe_str(row[k]) for k in "ABCDEFGHIJKLMNOP"}


@app.route('/pda', methods=['GET'])
def pda_page():
    """A PDA app felulete (pda.html)."""
    resp = send_file(PDA_HTML, mimetype="text/html")
    resp.headers["Cache-Control"] = "no-cache"
    return resp


@app.route('/pda/app.apk', methods=['GET'])
def pda_apk():
    """A PDA app telepitoje (RaktarPDA.apk ugyanebben a mappaban) - a PDA bongeszojebol letoltheto."""
    apk = os.path.join(os.path.dirname(os.path.abspath(__file__)), "RaktarPDA.apk")
    if not os.path.exists(apk):
        return jsonify({"status": "error", "message": "RaktarPDA.apk nincs feltoltve a szerverre"}), 404
    return send_file(apk, mimetype="application/vnd.android.package-archive",
                     as_attachment=True, download_name="RaktarPDA.apk")


@app.route('/pda/search', methods=['GET'])
def pda_search():
    """Kereses a PDA-rol: pontos R18 / B egyezes elore, kulonben szoveg + cikkszam egyezes (max 15)."""
    try:
        q = clean_code(request.args.get('q', ''))
        if not q:
            return jsonify({"status": "ok", "items": []}), 200
        conn = sqlite3.connect(DB_PATH)
        conn.row_factory = sqlite3.Row
        cursor = conn.cursor()
        cursor.execute("SELECT A, B, C, D, E, F, G, H, I, J, K, L, M, N, O, P FROM termekek "
                       "WHERE C IS NOT NULL AND TRIM(C) != ''")
        rows = [_pda_row(r) for r in cursor.fetchall()]
        conn.close()

        limit = max(1, min(request.args.get('limit', default=15, type=int), 200))
        rows.sort(key=lambda r: _r18_num(r["A"]), reverse=True)      # legujabb elol
        exact = [r for r in rows if clean_code(r["A"]) == q or clean_code(r["B"]) == q]
        items = list(exact)
        if not (re.match(r"^R18-\d+$", q) and exact):
            # Tobb szo: mindegyiknek szerepelnie kell (pl. "makita 18v")
            words = [w for w in q.lower().split() if w]
            for r in rows:
                if len(items) >= limit:
                    break
                if r in exact:
                    continue
                text = " ".join(r.values())
                low, part = text.lower(), _clean_part(text)
                if all(w in low or (_clean_part(w) and _clean_part(w) in part) for w in words):
                    items.append(r)
        return jsonify({"status": "ok", "items": items[:limit]}), 200
    except Exception as e:
        return pda_err(500, "S-500", str(e))


@app.route('/pda/sale', methods=['POST'])
def pda_sale():
    """Eladas rogzitese a PDA-rol. A termek adatait a VPS adatbazisbol veszi."""
    try:
        data = request.json or {}
        platform = str(data.get('platform', '')).strip().upper()
        code = clean_code(data.get('A', ''))
        if platform not in PDA_PLATFORMS:
            return pda_err(400, "S-400", "Ismeretlen platform")
        try:
            darab = int(str(data.get('darab', '1')).strip() or 1)
        except ValueError:
            darab = 0
        if darab < 1:
            return pda_err(400, "S-400", "Hibas darabszam")

        conn = sqlite3.connect(DB_PATH)
        conn.row_factory = sqlite3.Row
        cursor = conn.cursor()
        cursor.execute("SELECT A, B, C, D, E, F, G, H, I, J, K, L, M, N, O, P FROM termekek "
                       "WHERE UPPER(TRIM(A))=? OR UPPER(TRIM(B))=?", (code, code))
        row = cursor.fetchone()
        if not row:
            conn.close()
            return pda_err(404, "S-404", f"{code} nem talalhato a raktarban")
        prod = _pda_row(row)
        now = pda_now()
        eladas_id = f"E{now.strftime('%Y%m%d%H%M%S')}-{code}-{os.urandom(2).hex()}"
        cursor.execute("""
            INSERT INTO pda_eladasok (eladas_id, datum, platform, A, C, D, E, F, G, L, M, N,
                                      vevo, ar, darab, tracking, megjegyzes, keszlet_levonas, eszkoz)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        """, (eladas_id, now.strftime("%Y-%m-%d %H:%M"), platform, code,
              prod["C"], prod["D"], prod["E"], prod["F"], prod["G"], prod["L"], prod["M"], prod["N"],
              str(data.get('vevo', '')).strip(), str(data.get('ar', '')).strip(), str(darab),
              str(data.get('tracking', '')).strip(), str(data.get('megjegyzes', '')).strip(),
              1 if data.get('keszlet_levonas', True) else 0, str(data.get('eszkoz', 'PDA')).strip()))
        conn.commit()
        conn.close()
        return jsonify({"status": "ok", "eladas_id": eladas_id}), 200
    except Exception as e:
        return pda_err(500, "S-500", str(e))


@app.route('/pda/sales', methods=['GET'])
def pda_sales():
    """A PDA-rol rogzitett eladasok (legujabb elol). ?pending=1: csak a meg at nem vett eladasok."""
    try:
        limit = min(request.args.get('limit', default=50, type=int), 500)
        conn = sqlite3.connect(DB_PATH)
        conn.row_factory = sqlite3.Row
        cursor = conn.cursor()
        if request.args.get('pending'):
            cursor.execute("SELECT * FROM pda_eladasok WHERE atvette IS NULL ORDER BY id LIMIT ?", (limit,))
        else:
            cursor.execute("SELECT * FROM pda_eladasok ORDER BY id DESC LIMIT ?", (limit,))
        sales = [{k: (r[k] if r[k] is not None else "") for k in PDA_SALE_FIELDS} for r in cursor.fetchall()]
        conn.close()
        return jsonify({"status": "ok", "sales": sales}), 200
    except Exception as e:
        return pda_err(500, "S-500", str(e))


@app.route('/pda/sales/claim', methods=['POST'])
def pda_sales_claim():
    """Az asztali program atveszi az eladast. Csak EGY gep veheti at (a keszlet ne vonodjon le ketszer)."""
    try:
        data = request.json or {}
        eladas_id = str(data.get('eladas_id', '')).strip()
        gep = str(data.get('gep', '')).strip() or "ismeretlen gep"
        conn = sqlite3.connect(DB_PATH)
        cursor = conn.cursor()
        cursor.execute("UPDATE pda_eladasok SET atvette=?, atvetel_ido=? WHERE eladas_id=? AND atvette IS NULL",
                       (gep, pda_now().strftime("%Y-%m-%d %H:%M:%S"), eladas_id))
        ok = cursor.rowcount == 1
        conn.commit()
        conn.close()
        if ok:
            return jsonify({"status": "ok"}), 200
        return jsonify({"status": "taken", "kod": "S-409", "message": "Mar atvette egy masik gep (vagy nincs ilyen)"}), 409
    except Exception as e:
        return pda_err(500, "S-500", str(e))


# ---------- Athelyezes (raklap modositas) ----------
def _norm_pallet(v):
    """Egysegesitett hely: nagybetu, egy szokoz, "KISRAKTAR80" -> "KISRAKTAR 80"."""
    v = re.sub(r"\s+", " ", safe_str(v).upper()).strip()
    return re.sub(r"([^\W\d_])\s*(\d)", r"\1 \2", v)


def _pallet_tokens(v):
    """Egy raklap-cellaban tobb raklap is lehet: "493 494" = a 493-as ES a 494-es raklapon is van belole.
    A "KISRAKTAR 80" / "ENCS" tipusu helyek EGY helynek szamitanak (nem ket raklapnak)."""
    return [_norm_pallet(t) for t in re.findall(r"[^\W\d_]+\s*\d+|[^\s,;/+]+", safe_str(v).upper())]


@app.route('/pda/change', methods=['POST'])
def pda_change():
    """Termek-modositas a PDA-rol. Az asztali program veszi at es hajtja vegre (o irja a Google Sheets-be is),
    az eredmenyt visszakuldi (/pda/result).
      - athelyezes / egy mezo:  {"A", "mezo": "L", "uj": "574", "tipus": "ATHELYEZES"}
      - tobb mezo egyszerre:    {"A", "valtozasok": {"C": "...", "F": "2"}, "tipus": "MODOSITAS"}"""
    try:
        data = request.json or {}
        code = clean_code(data.get('A', ''))
        valt = data.get('valtozasok')
        if not isinstance(valt, dict):
            valt = {str(data.get('mezo', '')).strip().upper(): safe_str(data.get('uj', ''))}
        valt = {str(k).strip().upper(): safe_str(v) for k, v in valt.items()}
        bad = [k for k in valt if k not in PDA_EDITABLE]
        if bad or not valt:
            return pda_err(400, "S-400", f"Ez a mezo nem modosithato a PDA-rol: {', '.join(bad) or '-'}")
        if "L" in valt and len(valt) == 1 and not valt["L"]:
            return pda_err(400, "S-400", "Az uj raklap ures")
        conn = sqlite3.connect(DB_PATH)
        conn.row_factory = sqlite3.Row
        cur = conn.cursor()
        cur.execute("SELECT A, B, C, D, E, F, G, H, I, J, K, L, M, N, O, P FROM termekek "
                    "WHERE UPPER(TRIM(A))=? OR UPPER(TRIM(B))=?", (code, code))
        row = cur.fetchone()
        if not row:
            conn.close()
            return pda_err(404, "S-404", f"{code} nem talalhato a raktarban")
        regi = {k: safe_str(row[k]) for k in valt}
        valt = {k: v for k, v in valt.items() if v != regi[k]}
        regi = {k: regi[k] for k in valt}
        if not valt:
            conn.close()
            return pda_err(400, "S-400", "Nincs valtozas (minden mezo ugyanaz)")
        tipus = str(data.get('tipus', 'ATHELYEZES' if list(valt) == ["L"] else 'MODOSITAS')).upper()
        egy = len(valt) == 1
        mezo = list(valt)[0] if egy else "*"
        now = pda_now()
        vid = f"V{now.strftime('%Y%m%d%H%M%S')}-{code}-{os.urandom(2).hex()}"
        cur.execute("""INSERT INTO pda_valtozasok (valtozas_id, ido, tipus, A, C, D, E, mezo, regi, uj, megjegyzes, eszkoz, adat)
                       VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)""",
                    (vid, now.strftime("%Y-%m-%d %H:%M"), tipus, code,
                     safe_str(row["C"]), safe_str(row["D"]), safe_str(row["E"]), mezo,
                     regi[mezo] if egy else json.dumps(regi, ensure_ascii=False),
                     valt[mezo] if egy else json.dumps(valt, ensure_ascii=False),
                     safe_str(data.get('megjegyzes', '')), safe_str(data.get('eszkoz', 'PDA')) or "PDA",
                     json.dumps({"valtozasok": valt, "regi": regi}, ensure_ascii=False)))
        conn.commit()
        conn.close()
        return jsonify({"status": "ok", "valtozas_id": vid, "regi": regi if not egy else regi[mezo]}), 200
    except Exception as e:
        return pda_err(500, "S-500", str(e))


# ---------- Eladas javitasa / visszavonasa ----------
def _queue_sale_change(cur, sale, tipus, adat, eszkoz, megj=""):
    now = pda_now()
    vid = f"V{now.strftime('%Y%m%d%H%M%S')}-{sale['A']}-{os.urandom(2).hex()}"
    cur.execute("""INSERT INTO pda_valtozasok (valtozas_id, ido, tipus, A, C, D, E, mezo, regi, uj, megjegyzes, eszkoz, ref, adat)
                   VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)""",
                (vid, now.strftime("%Y-%m-%d %H:%M"), tipus, sale["A"], sale["C"], sale["D"], sale["E"],
                 "*", json.dumps(adat.get("regi", {}), ensure_ascii=False),
                 json.dumps(adat.get("valtozasok", {}), ensure_ascii=False), megj, eszkoz,
                 sale["eladas_id"], json.dumps(adat, ensure_ascii=False)))
    return vid


@app.route('/pda/sale/update', methods=['POST'])
def pda_sale_update():
    """PDA-s eladas potlasa / javitasa (vevo, ar, tracking, megjegyzes, darab, platform).
    Ha a gep meg nem vette at: helyben javitjuk (az uj ertekkel veszi at). Ha mar atvette: a gep javitja
    (darabszam-valtozasnal a keszletet is korrigalja), az eredmeny a Naploban latszik."""
    try:
        data = request.json or {}
        eid = str(data.get('eladas_id', '')).strip()
        fields = {k: safe_str(v) for k, v in (data.get('mezok') or {}).items() if k in PDA_SALE_EDITABLE}
        if "platform" in fields and fields["platform"].upper() not in PDA_PLATFORMS:
            return pda_err(400, "S-400", "Ismeretlen platform")
        if "darab" in fields:
            try:
                if int(fields["darab"]) < 1:
                    raise ValueError
            except ValueError:
                return pda_err(400, "S-400", "Hibas darabszam")
        conn = sqlite3.connect(DB_PATH)
        conn.row_factory = sqlite3.Row
        cur = conn.cursor()
        cur.execute("SELECT * FROM pda_eladasok WHERE eladas_id=?", (eid,))
        sale = cur.fetchone()
        if not sale:
            conn.close()
            return pda_err(404, "S-404", "Nincs ilyen eladas")
        if sale["visszavonva"]:
            conn.close()
            return pda_err(409, "S-409", "Ez az eladas vissza lett vonva")
        regi = {k: safe_str(sale[k]) for k in fields}
        fields = {k: v for k, v in fields.items() if v != regi[k]}
        if not fields:
            conn.close()
            return pda_err(400, "S-400", "Nincs valtozas")
        sets = ", ".join(f"{k}=?" for k in fields)
        cur.execute(f"UPDATE pda_eladasok SET {sets} WHERE eladas_id=?", (*fields.values(), eid))
        mode = "helyben"
        if sale["atvette"]:
            _queue_sale_change(cur, dict(sale), "ELADAS_MODOSITAS",
                               {"valtozasok": fields, "regi": {k: regi[k] for k in fields},
                                "keszlet_levonas": sale["keszlet_levonas"]},
                               safe_str(data.get('eszkoz', 'PDA')) or "PDA")
            mode = "gepre_var"
        conn.commit()
        conn.close()
        return jsonify({"status": "ok", "mod": mode}), 200
    except Exception as e:
        return pda_err(500, "S-500", str(e))


@app.route('/pda/sale/cancel', methods=['POST'])
def pda_sale_cancel():
    """PDA-s eladas visszavonasa. {"eladas_id", "keszlet_vissza": true}"""
    try:
        data = request.json or {}
        eid = str(data.get('eladas_id', '')).strip()
        vissza = bool(data.get('keszlet_vissza', True))
        conn = sqlite3.connect(DB_PATH)
        conn.row_factory = sqlite3.Row
        cur = conn.cursor()
        cur.execute("SELECT * FROM pda_eladasok WHERE eladas_id=?", (eid,))
        sale = cur.fetchone()
        if not sale:
            conn.close()
            return pda_err(404, "S-404", "Nincs ilyen eladas")
        if sale["visszavonva"]:
            conn.close()
            return pda_err(409, "S-409", "Mar vissza lett vonva")
        now = pda_now().strftime("%Y-%m-%d %H:%M:%S")
        if not sale["atvette"]:
            # A gep meg nem vette at: egyszeruen nem adjuk at neki
            cur.execute("""UPDATE pda_eladasok SET visszavonva=?, atvette=?, atvetel_ido=?, eredmeny='ok',
                           eredmeny_uzenet=?, eredmeny_ido=? WHERE eladas_id=? AND atvette IS NULL""",
                        (now, "(visszavonva a PDA-n)", now, "Visszavonva, mielott a gep atvette - keszlet nem valtozott",
                         now, eid))
            mode = "helyben"
        else:
            cur.execute("UPDATE pda_eladasok SET visszavonva=? WHERE eladas_id=?", (now, eid))
            _queue_sale_change(cur, dict(sale), "ELADAS_TORLES",
                               {"keszlet_vissza": vissza, "keszlet_levonas": sale["keszlet_levonas"],
                                "darab": safe_str(sale["darab"])},
                               safe_str(data.get('eszkoz', 'PDA')) or "PDA")
            mode = "gepre_var"
        conn.commit()
        conn.close()
        return jsonify({"status": "ok", "mod": mode}), 200
    except Exception as e:
        return pda_err(500, "S-500", str(e))


@app.route('/pda/delete', methods=['POST'])
def pda_delete():
    """Termek TORLESE a PDA-rol. Az asztali program hajtja vegre: ToroltSor-ra archival, majd torli
    a gepen, a szerveren es a Google tablazatban. {"A", "ok": "miert", "eszkoz"}"""
    try:
        data = request.json or {}
        code = clean_code(data.get('A', ''))
        if not code.startswith("R18-"):
            return pda_err(400, "S-400", "Ervenytelen R18 kod")
        conn = sqlite3.connect(DB_PATH)
        conn.row_factory = sqlite3.Row
        cur = conn.cursor()
        cur.execute("SELECT A, B, C, D, E, F, G, H, I, J, K, L, M, N, O, P FROM termekek "
                    "WHERE UPPER(TRIM(A))=? OR UPPER(TRIM(B))=?", (code, code))
        row = cur.fetchone()
        sor = _pda_row(row) if row else {"A": code}
        now = pda_now()
        vid = f"V{now.strftime('%Y%m%d%H%M%S')}-{code}-{os.urandom(2).hex()}"
        ok = safe_str(data.get('ok', '')) or "PDA-rol torolve"
        cur.execute("""INSERT INTO pda_valtozasok (valtozas_id, ido, tipus, A, C, D, E, mezo, regi, uj, megjegyzes, eszkoz, adat)
                       VALUES (?, ?, 'TORLES', ?, ?, ?, ?, '*', 'torles', '', ?, ?, ?)""",
                    (vid, now.strftime("%Y-%m-%d %H:%M"), code, sor.get("C", ""), sor.get("D", ""), sor.get("E", ""),
                     ok, safe_str(data.get('eszkoz', 'PDA')) or "PDA", json.dumps({"sor": sor, "ok": ok}, ensure_ascii=False)))
        conn.commit()
        conn.close()
        return jsonify({"status": "ok", "valtozas_id": vid}), 200
    except Exception as e:
        return pda_err(500, "S-500", str(e))


@app.route('/pda/orphans', methods=['GET'])
def pda_orphans():
    """Az adatbazisban levo, de az elo Google tablazatban mar NEM szereplo R18-sorok."""
    try:
        res = requests.get(f"https://docs.google.com/spreadsheets/d/{LIVE_SHEETS_ID}/export?format=csv", timeout=20)
        res.encoding = 'utf-8'
        if res.status_code != 200:
            return pda_err(500, "S-500", f"A Google tablazat nem toltheto le (HTTP {res.status_code})")
        df = pd.read_csv(io.StringIO(res.text), header=None, dtype=str, keep_default_na=False)
        sheet_codes = set()
        for _, r in df.iterrows():
            for col in (0, 1):
                c = clean_code(safe_str(r.get(col)))
                if c.startswith("R18-"):
                    sheet_codes.add(c)
        conn = sqlite3.connect(DB_PATH)
        conn.row_factory = sqlite3.Row
        cur = conn.cursor()
        cur.execute("SELECT A, B, C, D, E, F, G, H, I, J, K, L, M, N, O, P FROM termekek WHERE A LIKE 'R18-%'")
        db_rows = [_pda_row(r) for r in cur.fetchall()]
        conn.close()
        # Biztonsag: ha a letoltott tablazat gyanusan kicsi, nem mondunk semmit "arvanak"
        if len(sheet_codes) < max(10, len(db_rows) * 0.5):
            return pda_err(500, "S-500", f"A letoltott tablazat gyanusan kicsi ({len(sheet_codes)} R18-sor, adatbazisban {len(db_rows)}) - kihagyva")
        arva = [r for r in db_rows if clean_code(r["A"]) not in sheet_codes and clean_code(r["B"]) not in sheet_codes]
        arva.sort(key=lambda r: _r18_num(r["A"]))
        return jsonify({"status": "ok", "tablazatban": len(sheet_codes), "adatbazisban": len(db_rows), "items": arva}), 200
    except Exception as e:
        return pda_err(500, "S-500", str(e))


@app.route('/pda/change/cancel', methods=['POST'])
def pda_change_cancel():
    """PDA-s modositas (athelyezes / szerkesztes / eladas-javitas) VISSZAVONASA - pl. veletlen nyomas utan.
    Ha a gep meg nem hajtotta vegre: egyszeruen nem adjuk at neki. Ha mar vegrehajtotta: egy ellentetes
    modositast kuldunk (a regi ertekek vissza), amit szinten a gep hajt vegre."""
    try:
        data = request.json or {}
        vid = str(data.get('valtozas_id', '')).strip()
        eszkoz = safe_str(data.get('eszkoz', 'PDA')) or "PDA"
        conn = sqlite3.connect(DB_PATH)
        conn.row_factory = sqlite3.Row
        cur = conn.cursor()
        cur.execute("SELECT * FROM pda_valtozasok WHERE valtozas_id=?", (vid,))
        ch = cur.fetchone()
        if not ch:
            conn.close()
            return pda_err(404, "S-404", "Nincs ilyen modositas")
        if ch["visszavonva"]:
            conn.close()
            return pda_err(409, "S-409", "Mar vissza lett vonva")
        if ch["tipus"] == "ELADAS_TORLES":
            conn.close()
            return pda_err(400, "S-400", "Eladas visszavonasa nem vonhato vissza - rogzitsd ujra az eladast")
        try:
            ad = json.loads(ch["adat"] or "{}")
        except Exception:
            ad = {}
        valt = ad.get("valtozasok") or ({ch["mezo"]: ch["uj"]} if ch["mezo"] and ch["mezo"] != "*" else {})
        regi = ad.get("regi") or ({ch["mezo"]: ch["regi"]} if ch["mezo"] and ch["mezo"] != "*" else {})
        now = pda_now()
        now_s = now.strftime("%Y-%m-%d %H:%M:%S")
        cur.execute("UPDATE pda_valtozasok SET visszavonva=? WHERE valtozas_id=?", (now_s, vid))
        if ch["tipus"] == "ELADAS_MODOSITAS" and ch["ref"] and regi:
            sets = ", ".join(f"{k}=?" for k in regi if k in PDA_SALE_EDITABLE)
            if sets:
                cur.execute(f"UPDATE pda_eladasok SET {sets} WHERE eladas_id=?",
                            (*[v for k, v in regi.items() if k in PDA_SALE_EDITABLE], ch["ref"]))
        if not ch["atvette"] or ch["eredmeny"] == "hiba":
            # Meg nem hajtotta vegre a gep (vagy nem sikerult neki): nincs mit visszacsinalni
            if not ch["atvette"]:
                cur.execute("""UPDATE pda_valtozasok SET atvette=?, atvetel_ido=?, eredmeny='ok', eredmeny_uzenet=?, eredmeny_ido=?
                               WHERE valtozas_id=? AND atvette IS NULL""",
                            ("(visszavonva a PDA-n)", now_s, "Visszavonva, mielott a gep vegrehajtotta - semmi nem valtozott",
                             now_s, vid))
            mode = "helyben"
        elif ch["tipus"] == "TORLES":
            # Mar torolve: visszaallitas ugyanazokkal az adatokkal (a gep hajtja vegre)
            nvid = f"V{now.strftime('%Y%m%d%H%M%S')}-{ch['A']}-{os.urandom(2).hex()}"
            cur.execute("""INSERT INTO pda_valtozasok (valtozas_id, ido, tipus, A, C, D, E, mezo, regi, uj, megjegyzes, eszkoz, adat)
                           VALUES (?, ?, 'VISSZAALLITAS', ?, ?, ?, ?, '*', '', '', ?, ?, ?)""",
                        (nvid, now.strftime("%Y-%m-%d %H:%M"), ch["A"], ch["C"], ch["D"], ch["E"],
                         f"Torles visszavonasa: {vid}", eszkoz, json.dumps({"sor": ad.get("sor") or {}, "visszavont": vid}, ensure_ascii=False)))
            mode = "gepre_var"
        else:
            nvid = f"V{now.strftime('%Y%m%d%H%M%S')}-{ch['A']}-{os.urandom(2).hex()}"
            ntipus = "ELADAS_MODOSITAS" if ch["tipus"] == "ELADAS_MODOSITAS" else "VISSZAVONAS"
            nadat = {"valtozasok": regi, "regi": valt, "visszavont": vid}
            if ch["tipus"] == "ELADAS_MODOSITAS":
                nadat["keszlet_levonas"] = ad.get("keszlet_levonas", 1)
            egy = len(regi) == 1
            mezo = list(regi)[0] if egy else "*"
            cur.execute("""INSERT INTO pda_valtozasok (valtozas_id, ido, tipus, A, C, D, E, mezo, regi, uj, megjegyzes, eszkoz, ref, adat)
                           VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)""",
                        (nvid, now.strftime("%Y-%m-%d %H:%M"), ntipus, ch["A"], ch["C"], ch["D"], ch["E"], mezo,
                         valt.get(mezo, "") if egy else json.dumps(valt, ensure_ascii=False),
                         regi.get(mezo, "") if egy else json.dumps(regi, ensure_ascii=False),
                         f"Visszavonas: {vid}", eszkoz, ch["ref"], json.dumps(nadat, ensure_ascii=False)))
            mode = "gepre_var"
        conn.commit()
        conn.close()
        return jsonify({"status": "ok", "mod": mode}), 200
    except Exception as e:
        return pda_err(500, "S-500", str(e))


@app.route('/pda/sale/mirror', methods=['POST'])
def pda_sale_mirror():
    """Az asztali programban (Eladottak ful) javitott eladas-mezo tukrozese a PDA-s eladasra, hogy a PDA
    Naplojaban is a friss adat latsszon. {"eladas_id", "mezo", "ertek"} - ha nincs ilyen PDA-s eladas, nem csinal semmit."""
    try:
        data = request.json or {}
        mezo = str(data.get('mezo', ''))
        if mezo not in PDA_SALE_EDITABLE:
            return jsonify({"status": "ok", "mod": "kihagyva"}), 200
        conn = sqlite3.connect(DB_PATH)
        conn.execute(f"UPDATE pda_eladasok SET {mezo}=? WHERE eladas_id=?",
                     (safe_str(data.get('ertek', '')), str(data.get('eladas_id', ''))))
        conn.commit()
        conn.close()
        return jsonify({"status": "ok"}), 200
    except Exception as e:
        return pda_err(500, "S-500", str(e))


# ---------- Termek elozmenyei ----------
@app.route('/pda/history/push', methods=['POST'])
def pda_history_push():
    """Az asztali program feltolti a sajat Naplojanak uj bejegyzeseit. {"gep", "sorok": [{id, idopont, r18_kod, esemeny, reszletek, forras}]}"""
    try:
        data = request.json or {}
        gep = safe_str(data.get('gep', '')) or "ismeretlen gep"
        rows = data.get('sorok') or []
        conn = sqlite3.connect(DB_PATH)
        conn.executemany("""INSERT OR IGNORE INTO termek_naplo (gep, helyi_id, idopont, r18_kod, esemeny, reszletek, forras)
                            VALUES (?, ?, ?, ?, ?, ?, ?)""",
                         [(gep, int(r.get('id') or 0), safe_str(r.get('idopont')), clean_code(r.get('r18_kod')),
                           safe_str(r.get('esemeny')), safe_str(r.get('reszletek'))[:4000], safe_str(r.get('forras')))
                          for r in rows])
        conn.commit()
        conn.close()
        return jsonify({"status": "ok", "db": len(rows)}), 200
    except Exception as e:
        return pda_err(500, "S-500", str(e))


_HIST_LBL = dict(PDA_EDITABLE, vevo="Vevo", ar="Ar", tracking="Tracking", megjegyzes="Megjegyzes",
                 darab="Darab", platform="Platform")


def _fmt_change(r):
    """Olvashato leiras egy PDA-s modositasrol: 'Marka: A -> B | DB: 1 -> 4'."""
    try:
        ad = json.loads(r["adat"] or "{}")
    except Exception:
        ad = {}
    if r["tipus"] in ("TORLES", "VISSZAALLITAS"):
        sor = ad.get("sor") or {}
        txt = ("TORLES" if r["tipus"] == "TORLES" else "VISSZAALLITAS") + ": " + " ".join(
            str(sor.get(k, "")) for k in ("C", "D", "E") if sor.get(k)) + (f" | ok: {ad.get('ok')}" if ad.get("ok") else "")
    elif r["tipus"] == "ELADAS_TORLES":
        txt = "Eladas visszavonasa, keszlet vissza: " + (f"igen (+{ad.get('darab', '')} db)" if ad.get("keszlet_vissza") else "nem")
    elif ad.get("valtozasok"):
        regi = ad.get("regi") or {}
        txt = " | ".join(f"{_HIST_LBL.get(k, k)}: {regi.get(k, '') or '-'} -> {v or '-'}" for k, v in ad["valtozasok"].items())
    else:
        txt = f"{_HIST_LBL.get(r['mezo'], r['mezo'])}: {r['regi'] or '-'} -> {r['uj'] or '-'}"
    if r["eredmeny_uzenet"]:
        txt += f"  ==> {'OK' if r['eredmeny'] == 'ok' else 'HIBA'}: {r['eredmeny_uzenet']}"
    elif not r["atvette"]:
        txt += "  ==> a gep meg nem vette at"
    return txt


@app.route('/pda/history', methods=['GET'])
def pda_history():
    """Egy termek osszes ismert esemenye (legujabb elol): a gepek naplojabol + PDA eladasok / modositasok."""
    try:
        code = clean_code(request.args.get('code', ''))
        if not code:
            return pda_err(400, "S-400", "Nincs megadva R18 kod")
        conn = sqlite3.connect(DB_PATH)
        conn.row_factory = sqlite3.Row
        cur = conn.cursor()
        out = []
        cur.execute("SELECT * FROM termek_naplo WHERE r18_kod=? ORDER BY idopont DESC LIMIT 300", (code,))
        for r in cur.fetchall():
            out.append({"forras_tipus": "gep", "ido": r["idopont"], "esemeny": r["esemeny"],
                        "reszletek": r["reszletek"], "honnan": r["forras"], "gep": r["gep"]})
        cur.execute("SELECT * FROM pda_eladasok WHERE A=? ORDER BY id DESC LIMIT 100", (code,))
        for r in cur.fetchall():
            out.append({"forras_tipus": "pda", "ido": r["datum"], "esemeny": "PDA ELADAS",
                        "reszletek": f"{r['platform']} | {r['darab']} db | vevo: {r['vevo'] or '-'} | ar: {r['ar'] or '-'}"
                                     + (" | VISSZAVONVA" if r["visszavonva"] else ""),
                        "honnan": f"Zebra PDA ({r['eszkoz']})", "gep": r["atvette"] or ""})
        cur.execute("SELECT * FROM pda_valtozasok WHERE A=? ORDER BY id DESC LIMIT 100", (code,))
        for r in cur.fetchall():
            out.append({"forras_tipus": "pda", "ido": r["ido"], "esemeny": f"PDA {r['tipus']}",
                        "reszletek": _fmt_change(r),
                        "honnan": f"Zebra PDA ({r['eszkoz']})", "gep": r["atvette"] or ""})
        conn.close()
        out.sort(key=lambda x: str(x.get("ido") or ""), reverse=True)
        return jsonify({"status": "ok", "kod": code, "esemenyek": out[:300]}), 200
    except Exception as e:
        return pda_err(500, "S-500", str(e))


@app.route('/pda/changes', methods=['GET'])
def pda_changes():
    """A PDA-rol kert modositasok. ?pending=1: csak a meg at nem vettek (az asztali programnak)."""
    try:
        limit = min(request.args.get('limit', default=50, type=int), 500)
        conn = sqlite3.connect(DB_PATH)
        conn.row_factory = sqlite3.Row
        cur = conn.cursor()
        if request.args.get('pending'):
            cur.execute("SELECT * FROM pda_valtozasok WHERE atvette IS NULL ORDER BY id LIMIT ?", (limit,))
        else:
            cur.execute("SELECT * FROM pda_valtozasok ORDER BY id DESC LIMIT ?", (limit,))
        rows = [{k: (r[k] if r[k] is not None else "") for k in PDA_CHANGE_FIELDS} for r in cur.fetchall()]
        conn.close()
        return jsonify({"status": "ok", "changes": rows}), 200
    except Exception as e:
        return pda_err(500, "S-500", str(e))


@app.route('/pda/changes/claim', methods=['POST'])
def pda_changes_claim():
    try:
        data = request.json or {}
        vid = str(data.get('valtozas_id', '')).strip()
        gep = str(data.get('gep', '')).strip() or "ismeretlen gep"
        conn = sqlite3.connect(DB_PATH)
        cur = conn.cursor()
        cur.execute("UPDATE pda_valtozasok SET atvette=?, atvetel_ido=? WHERE valtozas_id=? AND atvette IS NULL",
                    (gep, pda_now().strftime("%Y-%m-%d %H:%M:%S"), vid))
        ok = cur.rowcount == 1
        conn.commit()
        conn.close()
        if ok:
            return jsonify({"status": "ok"}), 200
        return jsonify({"status": "taken", "kod": "S-409", "message": "Mar atvette egy masik gep (vagy nincs ilyen)"}), 409
    except Exception as e:
        return pda_err(500, "S-500", str(e))


@app.route('/pda/result', methods=['POST'])
def pda_result():
    """Az asztali program visszajelzese: sikerult-e vegrehajtani az eladast / modositast.
    {"tipus": "eladas"|"valtozas", "id": ..., "ok": true/false, "uzenet": "...", "kod": "A-21"}"""
    try:
        data = request.json or {}
        tipus = str(data.get('tipus', '')).strip()
        rid = str(data.get('id', '')).strip()
        ok = bool(data.get('ok'))
        kod = safe_str(data.get('kod', ''))
        uzenet = safe_str(data.get('uzenet', ''))
        if kod and not ok:
            uzenet = f"[{kod}] {uzenet}"
        table, key = (("pda_eladasok", "eladas_id") if tipus == "eladas" else
                      ("pda_valtozasok", "valtozas_id") if tipus == "valtozas" else (None, None))
        if not table:
            return pda_err(400, "S-400", "Ismeretlen tipus")
        conn = sqlite3.connect(DB_PATH)
        cur = conn.cursor()
        cur.execute(f"UPDATE {table} SET eredmeny=?, eredmeny_uzenet=?, eredmeny_ido=? WHERE {key}=?",
                    ("ok" if ok else "hiba", uzenet, pda_now().strftime("%Y-%m-%d %H:%M:%S"), rid))
        found = cur.rowcount == 1
        if not ok:
            cur.execute("INSERT INTO pda_naplo (ido, eszkoz, szint, esemeny, kod, reszletek) VALUES (?, ?, ?, ?, ?, ?)",
                        (pda_now().strftime("%Y-%m-%d %H:%M:%S"), safe_str(data.get('gep', 'asztali gep')),
                         "HIBA", f"Asztali program: {tipus} nem sikerult", kod, f"{rid}: {uzenet}"))
        conn.commit()
        conn.close()
        if not found:
            return pda_err(404, "S-404", "Nincs ilyen tetel")
        return jsonify({"status": "ok"}), 200
    except Exception as e:
        return pda_err(500, "S-500", str(e))


# ---------- Excel Live munkalap (PDA) ----------
@app.route('/pda/sheet', methods=['GET'])
def pda_sheet():
    """Az Excel Live munkalap egy szelete R18-szam szerint sorban.
    ?from=13000&limit=100  - a 13000-es R18-tol kezdve | ?last=1 - a legutolso 'limit' sor
    ?kitoltott=1           - csak a kitoltott (Marka/Tipus/Cikkszam) sorok"""
    try:
        limit = max(10, min(request.args.get('limit', default=100, type=int), 300))
        conn = sqlite3.connect(DB_PATH)
        conn.row_factory = sqlite3.Row
        cur = conn.cursor()
        cur.execute("SELECT A, B, C, D, E, F, G, H, I, J, K, L, M, N, O, P FROM termekek WHERE A LIKE 'R18-%'")
        rows = [r for r in (_pda_row(x) for x in cur.fetchall()) if _r18_num(r["A"]) >= 0]
        conn.close()
        if request.args.get('kitoltott'):
            rows = [r for r in rows if r["C"] or r["D"] or r["E"]]
        rows.sort(key=lambda r: _r18_num(r["A"]))
        maxnum = _r18_num(rows[-1]["A"]) if rows else 0
        minnum = _r18_num(rows[0]["A"]) if rows else 0
        if request.args.get('last'):
            part = rows[-limit:]
        else:
            start = request.args.get('from', default=minnum, type=int)
            if request.args.get('before'):
                # visszafele lapozas: a 'from' ELOTTI 'limit' sor
                prev = [r for r in rows if _r18_num(r["A"]) < start]
                part = prev[-limit:]
            else:
                part = [r for r in rows if _r18_num(r["A"]) >= start][:limit]
        return jsonify({"status": "ok", "items": part, "min": minnum, "max": maxnum, "osszes": len(rows)}), 200
    except Exception as e:
        return pda_err(500, "S-500", str(e))


# ---------- Fenykepek (a NAS-on vannak; a gepi program tolti fel keresre) ----------
PDA_FOTO_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "pda_fotok")
FOTO_MAX = 8


def init_foto_db():
    try:
        os.makedirs(PDA_FOTO_DIR, exist_ok=True)
        conn = sqlite3.connect(DB_PATH)
        conn.execute("""CREATE TABLE IF NOT EXISTS foto_keresek (
                            kod TEXT PRIMARY KEY, kert TEXT, kesz TEXT, darab INTEGER, uzenet TEXT, gep TEXT)""")
        conn.commit()
        conn.close()
    except Exception as e:
        print(f"[VPS] Foto tabla inicializalasi hiba: {e}")

init_foto_db()


def _foto_dir(code):
    return os.path.join(PDA_FOTO_DIR, re.sub(r"[^A-Z0-9-]", "", code))


def _foto_list(code):
    d = _foto_dir(code)
    if not os.path.isdir(d):
        return []
    return sorted(f for f in os.listdir(d) if f.lower().endswith(".jpg"))


@app.route('/pda/photos', methods=['GET'])
def pda_photos():
    """Egy termek fenykepei. Ha meg nincsenek a szerveren, kerest tesz a gepeknek (azok toltik fel a NAS-rol).
    ?frissit=1: ujra lekeri a geptol (pl. uj fotok keszultek)."""
    try:
        code = clean_code(request.args.get('code', ''))
        if not code.startswith("R18-"):
            return pda_err(400, "S-400", "Ervenytelen R18 kod")
        conn = sqlite3.connect(DB_PATH)
        conn.row_factory = sqlite3.Row
        cur = conn.cursor()
        if request.args.get('frissit'):
            d = _foto_dir(code)
            if os.path.isdir(d):
                for f in os.listdir(d):
                    try:
                        os.remove(os.path.join(d, f))
                    except OSError:
                        pass
            cur.execute("DELETE FROM foto_keresek WHERE kod=?", (code,))
        files = _foto_list(code)
        cur.execute("SELECT * FROM foto_keresek WHERE kod=?", (code,))
        req = cur.fetchone()
        now = pda_now().strftime("%Y-%m-%d %H:%M:%S")
        if files:
            conn.close()
            return jsonify({"status": "ok", "fotok": files, "kesz": True}), 200
        if req and req["kesz"]:
            conn.close()
            return jsonify({"status": "ok", "fotok": [], "kesz": True, "uzenet": req["uzenet"] or "Nincs fenykep", "gep": req["gep"] or ""}), 200
        if not req:
            cur.execute("INSERT INTO foto_keresek (kod, kert) VALUES (?, ?)", (code, now))
            conn.commit()
        conn.close()
        return jsonify({"status": "ok", "fotok": [], "kesz": False, "kert": (req["kert"] if req else now)}), 200
    except Exception as e:
        return pda_err(500, "S-500", str(e))


@app.route('/pda/photo/<code>/<name>', methods=['GET'])
def pda_photo_file(code, name):
    code = clean_code(code)
    if not re.match(r"^[0-9]+\.jpg$", name):
        return pda_err(404, "S-404", "Nincs ilyen kep")
    path = os.path.join(_foto_dir(code), name)
    if not os.path.exists(path):
        return pda_err(404, "S-404", "Nincs ilyen kep")
    resp = send_file(path, mimetype="image/jpeg")
    resp.headers["Cache-Control"] = "max-age=86400"
    return resp


@app.route('/pda/photo_requests', methods=['GET'])
def pda_photo_requests():
    """A gepi programnak: mely termekek kepeit kerte a PDA (meg nincs kesz)."""
    try:
        conn = sqlite3.connect(DB_PATH)
        conn.row_factory = sqlite3.Row
        cur = conn.cursor()
        cur.execute("SELECT kod, kert FROM foto_keresek WHERE kesz IS NULL ORDER BY kert LIMIT 10")
        rows = [{"kod": r["kod"], "kert": r["kert"]} for r in cur.fetchall()]
        conn.close()
        return jsonify({"status": "ok", "keresek": rows}), 200
    except Exception as e:
        return pda_err(500, "S-500", str(e))


@app.route('/pda/photos/upload', methods=['POST'])
def pda_photos_upload():
    """A gepi program feltolti egy termek (mar kicsinyitett, JPG) kepeit. Ha nincs kep: uzenet a 'nincs' mezoben."""
    try:
        code = clean_code(request.form.get('kod', ''))
        if not code.startswith("R18-"):
            return pda_err(400, "S-400", "Ervenytelen R18 kod")
        d = _foto_dir(code)
        os.makedirs(d, exist_ok=True)
        for f in os.listdir(d):
            try:
                os.remove(os.path.join(d, f))
            except OSError:
                pass
        files = request.files.getlist('fotok')[:FOTO_MAX]
        for i, f in enumerate(files, start=1):
            f.save(os.path.join(d, f"{i}.jpg"))
        conn = sqlite3.connect(DB_PATH)
        conn.execute("""INSERT INTO foto_keresek (kod, kert, kesz, darab, uzenet, gep) VALUES (?, ?, ?, ?, ?, ?)
                        ON CONFLICT(kod) DO UPDATE SET kesz=excluded.kesz, darab=excluded.darab, uzenet=excluded.uzenet, gep=excluded.gep""",
                     (code, pda_now().strftime("%Y-%m-%d %H:%M:%S"), pda_now().strftime("%Y-%m-%d %H:%M:%S"), len(files),
                      safe_str(request.form.get('uzenet', '')) or ("" if files else "Nincs fenykep a NAS-on"),
                      safe_str(request.form.get('gep', ''))))
        conn.commit()
        conn.close()
        return jsonify({"status": "ok", "darab": len(files)}), 200
    except Exception as e:
        return pda_err(500, "S-500", str(e))


# ---------- Raktarterkep (ugyanaz a Google tablazat, mint a gepi programban) ----------
MAP_SHEETS_ID = "13uw_-HNPHep94j-Ra9TsJrUYXADa8qcuBYyBuFdL1Fw"
MAP_CACHE_SEC = 300
_map_cache = {"ido": 0.0, "rows": None}


def _load_map():
    import openpyxl
    res = requests.get(f"https://docs.google.com/spreadsheets/d/{MAP_SHEETS_ID}/export?format=xlsx",
                       headers={"User-Agent": "Mozilla/5.0"}, timeout=30)
    if res.status_code != 200:
        raise RuntimeError(f"A terkep tablazat nem toltheto le (HTTP {res.status_code})")
    wb = openpyxl.load_workbook(io.BytesIO(res.content), data_only=True)
    ws = wb.active
    rows = []
    for row in ws.iter_rows():
        out = []
        for cell in row:
            v = "" if cell.value is None else str(cell.value).strip()
            if v.endswith(".0"):
                v = v[:-2]
            col = ""
            try:
                rgb = cell.fill.start_color.rgb if cell.fill and cell.fill.start_color else None
                if isinstance(rgb, str) and len(rgb) >= 6 and rgb != "00000000":
                    col = "#" + rgb[-6:]
            except Exception:
                pass
            out.append([v, col])
        rows.append(out)
    # a jobb oldali / also teljesen ures reszeket levagjuk
    while rows and not any(v for v, _ in rows[-1]):
        rows.pop()
    width = max((max((i + 1 for i, (v, c) in enumerate(r) if v or c), default=0) for r in rows), default=0)
    return [r[:width] for r in rows]


@app.route('/pda/map', methods=['GET'])
def pda_map():
    """A raktarterkep cellai (ertek + hatterszin). 5 percig tarolva; ?frissit=1 ujratolti."""
    try:
        if request.args.get('frissit') or not _map_cache["rows"] or time.time() - _map_cache["ido"] > MAP_CACHE_SEC:
            _map_cache["rows"] = _load_map()
            _map_cache["ido"] = time.time()
        return jsonify({"status": "ok", "rows": _map_cache["rows"],
                        "frissitve": datetime.datetime.fromtimestamp(_map_cache["ido"]).strftime("%H:%M:%S")}), 200
    except ImportError:
        return pda_err(500, "S-500", "A szerveren nincs telepitve az openpyxl (pip install openpyxl)")
    except Exception as e:
        return pda_err(500, "S-500", str(e))


# ---------- Leltar: egy raklap termekei ----------
@app.route('/pda/pallet', methods=['GET'])
def pda_pallet():
    """A rendszer szerint az adott raklapon (L) levo termekek."""
    try:
        pallet = _norm_pallet(request.args.get('l', ''))
        if not pallet:
            return pda_err(400, "S-400", "Nincs megadva raklap")
        conn = sqlite3.connect(DB_PATH)
        conn.row_factory = sqlite3.Row
        cur = conn.cursor()
        cur.execute("SELECT A, B, C, D, E, F, G, H, I, J, K, L, M, N, O, P FROM termekek "
                    "WHERE C IS NOT NULL AND TRIM(C) != ''")
        items = [_pda_row(r) for r in cur.fetchall()]
        conn.close()
        items = [it for it in items if pallet in _pallet_tokens(it["L"]) or _norm_pallet(it["L"]) == pallet]
        items.sort(key=lambda it: it["A"])
        return jsonify({"status": "ok", "raklap": pallet, "items": items}), 200
    except Exception as e:
        return pda_err(500, "S-500", str(e))


# ---------- Naplo ----------
@app.route('/pda/log', methods=['POST'])
def pda_log_add():
    """Esemeny / hiba bejegyzese a PDA naplojaba (pl. lezart leltar, PDA oldali hiba)."""
    try:
        data = request.json or {}
        conn = sqlite3.connect(DB_PATH)
        conn.execute("INSERT INTO pda_naplo (ido, eszkoz, szint, esemeny, kod, reszletek) VALUES (?, ?, ?, ?, ?, ?)",
                     (pda_now().strftime("%Y-%m-%d %H:%M:%S"), safe_str(data.get('eszkoz', 'PDA')) or "PDA",
                      safe_str(data.get('szint', 'INFO')).upper() or "INFO", safe_str(data.get('esemeny', '')),
                      safe_str(data.get('kod', '')), safe_str(data.get('reszletek', ''))[:4000]))
        conn.commit()
        conn.close()
        return jsonify({"status": "ok"}), 200
    except Exception as e:
        return pda_err(500, "S-500", str(e))


@app.route('/pda/log', methods=['GET'])
def pda_log():
    """Osszesitett PDA naplo (legujabb elol): eladasok, athelyezesek, leltarak, hibak - az allapotukkal.
    ?hibak=1: csak a hibak (asztali program hibai + PDA hibak)."""
    try:
        limit = min(request.args.get('limit', default=100, type=int), 500)
        only_err = bool(request.args.get('hibak'))
        conn = sqlite3.connect(DB_PATH)
        conn.row_factory = sqlite3.Row
        cur = conn.cursor()
        out = []
        cur.execute("SELECT * FROM pda_eladasok ORDER BY id DESC LIMIT ?", (limit,))
        for r in cur.fetchall():
            out.append({"fajta": "eladas", "ido": r["datum"], "rendez": r["datum"] or "",
                        **{k: (r[k] if r[k] is not None else "") for k in PDA_SALE_FIELDS}})
        cur.execute("SELECT * FROM pda_valtozasok ORDER BY id DESC LIMIT ?", (limit,))
        for r in cur.fetchall():
            out.append({"fajta": "valtozas", "rendez": r["ido"] or "",
                        **{k: (r[k] if r[k] is not None else "") for k in PDA_CHANGE_FIELDS}})
        cur.execute("SELECT * FROM pda_naplo ORDER BY id DESC LIMIT ?", (limit,))
        for r in cur.fetchall():
            out.append({"fajta": "naplo", "rendez": (r["ido"] or "")[:16], "ido": r["ido"], "eszkoz": r["eszkoz"],
                        "szint": r["szint"], "esemeny": r["esemeny"], "kod": r["kod"], "reszletek": r["reszletek"]})
        conn.close()
        if only_err:
            out = [x for x in out if x.get("eredmeny") == "hiba" or x.get("szint") == "HIBA"]
        out.sort(key=lambda x: x.get("rendez", ""), reverse=True)
        return jsonify({"status": "ok", "most": pda_now().strftime("%Y-%m-%d %H:%M"), "naplo": out[:limit]}), 200
    except Exception as e:
        return pda_err(500, "S-500", str(e))


if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000)
