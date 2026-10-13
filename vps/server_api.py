# -*- coding: utf-8 -*-
# FIGYELEM: a fajl szandekosan ekezet nelkuli (ASCII), hogy a WinSCP szerkesztoje se rontsa el.
import threading
import sqlite3
import time
import requests
import io
import pandas as pd
from flask import Flask, jsonify, request

app = Flask(__name__)

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


if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000)
