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
    sent = request.headers.get("X-Raktar-Token", "")
    if sent and hmac.compare_digest(sent.encode("utf-8"), API_TOKEN.encode("utf-8")):
        return None
    return jsonify({"status": "error", "message": "Hibas vagy hianyzo jelszo (token)"}), 401

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
                   "atvette", "atvetel_ido")


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


def _clean_part(text):
    return re.sub(r"[\s\-_/.,]+", "", str(text or "")).upper()


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

        exact = [r for r in rows if clean_code(r["A"]) == q or clean_code(r["B"]) == q]
        items = list(exact)
        if not (re.match(r"^R18-\d+$", q) and exact):
            q_low, q_part = q.lower(), _clean_part(q)
            for r in rows:
                if len(items) >= 15:
                    break
                if r in exact:
                    continue
                text = " ".join(r.values())
                if q_low in text.lower() or (q_part and q_part in _clean_part(text)):
                    items.append(r)
        return jsonify({"status": "ok", "items": items[:15]}), 200
    except Exception as e:
        return jsonify({"status": "error", "message": str(e)}), 500


@app.route('/pda/sale', methods=['POST'])
def pda_sale():
    """Eladas rogzitese a PDA-rol. A termek adatait a VPS adatbazisbol veszi."""
    try:
        data = request.json or {}
        platform = str(data.get('platform', '')).strip().upper()
        code = clean_code(data.get('A', ''))
        if platform not in PDA_PLATFORMS:
            return jsonify({"status": "error", "message": "Ismeretlen platform"}), 400
        try:
            darab = int(str(data.get('darab', '1')).strip() or 1)
        except ValueError:
            darab = 0
        if darab < 1:
            return jsonify({"status": "error", "message": "Hibas darabszam"}), 400

        conn = sqlite3.connect(DB_PATH)
        conn.row_factory = sqlite3.Row
        cursor = conn.cursor()
        cursor.execute("SELECT A, B, C, D, E, F, G, H, I, J, K, L, M, N, O, P FROM termekek "
                       "WHERE UPPER(TRIM(A))=? OR UPPER(TRIM(B))=?", (code, code))
        row = cursor.fetchone()
        if not row:
            conn.close()
            return jsonify({"status": "error", "message": f"{code} nem talalhato a raktarban"}), 404
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
        return jsonify({"status": "error", "message": str(e)}), 500


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
        return jsonify({"status": "error", "message": str(e)}), 500


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
        return jsonify({"status": "taken", "message": "Mar atvette egy masik gep (vagy nincs ilyen)"}), 409
    except Exception as e:
        return jsonify({"status": "error", "message": str(e)}), 500


if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000)
