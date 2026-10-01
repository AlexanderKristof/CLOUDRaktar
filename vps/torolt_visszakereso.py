#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# FIGYELEM: a fajl szandekosan ekezet nelkuli (ASCII), hogy a WinSCP szerkesztoje se rontsa el.
"""
KORABBI TORLESEK VISSZAKERESESE - /opt/raktarapp/torolt_visszakereso.py   (egyszer futtatando)

Osszeveti az elo tablazat egy REGEBBI allapotat (Google verzioelozmeny-masolat VAGY egy regi .xlsx)
a MOSTANI elo tablazattal. Ami a regiben meg kitoltve megvolt, most viszont hianyzik vagy ki van
uritve, annak a regi, teljes adatai (A-P) a ToroltSor munkalapra kerulnek.

- Az elo tablazatot es a regi verziot CSAK OLVASSA.
- Alapbol csak KILISTAZZA, mit irna be (proba). Beirni csak az --irj kapcsoloval fog.
- Ami mar fent van a ToroltSor lapon (azonos A-P adat), azt nem irja be ujra.

Hasznalat:
  venv/bin/python torolt_visszakereso.py <regi .xlsx fajl VAGY regi tablazat ID>              (proba)
  venv/bin/python torolt_visszakereso.py <regi .xlsx fajl VAGY regi tablazat ID> --irj        (beiras)
  Opcio: --munkalap "Lap neve"   (a regi forras munkalapja; alapbol az elso)
"""
import datetime
import os
import sys

import torolt_figyelo as tf   # ugyanazok a beallitasok (elo tablazat ID, ToroltSor helye, credentials)

ROW_COLS = tf.ROW_COLS


def rows_to_dict(values):
    out = {}
    for r in values:
        r = ["" if v is None else str(v).strip() for v in (list(r) + [""] * 16)[:16]]
        code = r[0].upper()
        if code.startswith("R18-"):
            out[code] = dict(zip(ROW_COLS, r))
    return out


def read_old_source(source, worksheet):
    if source.lower().endswith((".xlsx", ".xlsm")):
        try:
            import openpyxl
        except ImportError:
            sys.exit("Hianyzik az openpyxl: venv/bin/pip install openpyxl")
        wb = openpyxl.load_workbook(source, read_only=True, data_only=True)
        ws = wb[worksheet] if worksheet else wb.worksheets[0]
        return list(ws.iter_rows(values_only=True)), os.path.basename(source)
    # Google tablazat ID - csak olvasas
    import gspread
    book = gspread.authorize(tf._creds(readonly=True)).open_by_key(source)
    ws = book.worksheet(worksheet) if worksheet else book.sheet1
    return ws.get_all_values(), book.title


def main():
    args = [a for a in sys.argv[1:]]
    if not args or args[0].startswith("--"):
        sys.exit(__doc__)
    source = args[0]
    write = "--irj" in args
    worksheet = ""
    if "--munkalap" in args:
        worksheet = args[args.index("--munkalap") + 1]

    if source.strip() == tf.DELETED_SPREADSHEET_ID.strip():
        sys.exit("A regi forras nem lehet a ToroltSor tablazata.")

    print("Regi verzio olvasasa...")
    old_values, old_name = read_old_source(source, worksheet)
    old = rows_to_dict(old_values)
    print(f"  {len(old)} R18-sor a regi verzioban ({old_name})")

    print("Mostani elo tablazat olvasasa (csak olvasas)...")
    current = rows_to_dict(tf.read_watched_values())
    print(f"  {len(current)} R18-sor most")

    found = []
    for code, old_row in old.items():
        if not tf.has_product_data(old_row):
            continue
        now_row = current.get(code)
        if now_row is None:
            found.append((old_row, "Korabbi torles: sor torolve"))
        elif not tf.has_product_data(now_row):
            found.append((old_row, "Korabbi torles: kiuritve"))

    found.sort(key=lambda x: int(x[0]["A"][4:]) if x[0]["A"][4:].isdigit() else 0)
    print(f"\n{len(found)} korabban torolt / kiuritett termek:")
    for row, forras in found:
        print(f"  {row['A']:<10} | {row['C'][:15]:<15} | {row['D'][:30]:<30} | {row['E'][:25]:<25} | {forras}")

    if not found:
        return
    if not write:
        print("\nEz csak PROBA volt, semmi nem lett beirva. Beirashoz futtasd ugyanigy, a vegere: --irj")
        return

    ws = tf.get_deleted_ws()
    existing = {tuple((list(r) + [""] * 16)[:16]) for r in ws.get_all_values()[1:]}
    stamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    rows = []
    for row, forras in found:
        vals = [row[c] for c in ROW_COLS]
        if tuple(vals) in existing:
            continue
        # Q oszlop: a pontos torlesi ido nem ismert - a visszakereses ideje + a regi verzio neve
        rows.append(vals + [f"{stamp} (visszakeresve: {old_name})", forras.replace("Korabbi torles", "Kor\u00e1bbi t\u00f6rl\u00e9s")
                            .replace("sor torolve", "sor t\u00f6r\u00f6lve").replace("kiuritve", "ki\u00fcr\u00edtve")])
        existing.add(tuple(vals))
    if rows:
        ws.append_rows(rows, value_input_option="RAW")
    print(f"\n{len(rows)} sor beirva a '{tf.DELETED_WORKSHEET}' munkalapra "
          f"({len(found) - len(rows)} mar fent volt).")


if __name__ == "__main__":
    main()
