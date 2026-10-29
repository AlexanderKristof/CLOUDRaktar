# VERZIÓK – automatikus frissítés a VPS-ről + régi verziók mentése a Google Drive-ra

## Hogyan működik

1. Az új `app_pyqt6CLOUD.py`-t **egyetlen helyre** töltöd fel: a VPS-en a `/opt/raktarapp/kliens/` mappába.
2. Minden gépen a program induláskor és utána 30 percenként megnézi, van-e újabb verzió.
   Ha van, a program tetején **villogó piros sáv** jelenik meg, és felugrik egy ablak.
   A **🆕 VERZIÓK** fülön egy gombnyomással frissíthető (a program magától újraindul).
3. Ha az új verzióhoz új Python csomag kell, a VERZIÓK fül lépésről lépésre leírja a telepítést,
   és egy gombbal automatikusan is feltelepíthető.
4. A régi verzió megmarad: a gépen a `regi_verziok` mappában, a VPS-en a `/opt/raktarapp/kliens/archivum/` mappában,
   és (ha beállítod, lásd lent) a **saját Google Drive-od** „Raktárprogram régi verziók” mappájában.

---

## 1. Az első alkalommal

1. WinSCP-vel lépj be a VPS-re, nyisd meg a `/opt/raktarapp/` mappát.
2. Hozz létre egy új mappát **`kliens`** néven (jobb klikk → Új → Könyvtár).
3. Töltsd fel a `server_api.py`-t a `/opt/raktarapp/` mappába (felülírás), és indítsd újra a szervert:
   ```
   sudo systemctl restart raktarapp
   ```
4. Ezt az egy verziót (`app_pyqt6CLOUD.py`, v2026.10.29-upd1) még **kézzel** kell minden gépre feltenni
   (ez az első, ami már tud magától frissülni). Ugyanezt a fájlt töltsd fel a `/opt/raktarapp/kliens/` mappába is.

## 2. Új verzió kiadása (később mindig így)

1. WinSCP: az új `app_pyqt6CLOUD.py` → `/opt/raktarapp/kliens/` (felülírás).
2. Kész. Szerver-újraindítás NEM kell. A gépek legkésőbb 30 percen belül jelzik
   (a VERZIÓK fülön az „🔄 Ellenőrzés most” gombbal azonnal is).

---

## 3. Régi verziók mentése a saját Google Drive-ra (egyszer kell beállítani, kb. 5 perc)

A VPS Google-fiókja (service account) a Gmail-es Drive-ra nem tud fájlt menteni, ezért egy kis
Google Apps Script végzi a mentést **a te fiókod nevében**.

### 3/a. A script létrehozása
1. Nyisd meg: **https://script.google.com** (azzal a Google fiókkal, amelyiknek a Drive-jára menteni szeretnél).
2. Bal fent: **Új projekt**.
3. A kódban lévő `function myFunction() { }` részt töröld ki, és illeszd be a teljes
   `vps/drive_mentes.gs` fájl tartalmát.
4. A kód elején az `IDE-IRJ-EGY-HOSSZU-JELSZOT` helyére írj egy hosszú, véletlenszerű jelszót
   (pl. `raktar-2026-xk93hd7q2m`). Ezt jegyezd fel, a 3/c lépésben kell.
5. Fent a projekt nevére kattintva nevezd el: `Raktar Drive mentes`. Mentés: **Ctrl+S**.

### 3/b. Közzététel webalkalmazásként
1. Jobb fent: **Telepítés (Deploy) → Új telepítés (New deployment)**.
2. A fogaskerék ikonnál válaszd: **Webalkalmazás (Web app)**.
3. Beállítások:
   - **Végrehajtás mint (Execute as):** `Én (Me)` – a saját fiókod
   - **Ki férhet hozzá (Who has access):** `Bárki (Anyone)`
     (a jelszó nélkül semmit nem lehet vele csinálni)
4. **Telepítés (Deploy)** → a Google engedélyt kér: **Hozzáférés engedélyezése** → válaszd a fiókod →
   ha „A Google nem ellenőrizte ezt az alkalmazást” üzenet jön: **Speciális → Ugrás a Raktar Drive mentes (nem biztonságos) oldalra → Engedélyezés**.
   (Ez a saját scripted, nyugodtan engedélyezheted.)
5. Másold ki a **Webalkalmazás URL**-t (`https://script.google.com/macros/s/.../exec`).
6. Ellenőrzés: nyisd meg ezt az URL-t a böngészőben → ezt kell kiírnia: `"A raktárprogram Drive mentés működik."`

### 3/c. Beállítás a VPS-en
1. A saját gépeden hozz létre egy szövegfájlt **`drive_mentes.json`** néven, ezzel a tartalommal
   (a két értéket cseréld ki a sajátodra):
   ```json
   {"url": "https://script.google.com/macros/s/IDE-A-WEBALKALMAZAS-URL/exec", "titok": "raktar-2026-xk93hd7q2m", "mappa_id": "IDE-A-DRIVE-MAPPA-AZONOSITO"}
   ```
   A **mappa_id** a Drive mappa linkjében a `/folders/` utáni rész a `?` előtt
   (pl. `https://drive.google.com/drive/folders/1AbC...xyz?usp=...` → `1AbC...xyz`).
   Ha kihagyod, a Drive-od gyökerében létrehoz egy „Raktárprogram régi verziók” mappát.
   A mappa azonosítója csak ebbe a fájlba kerül (a GitHub-ra nem).
2. WinSCP-vel töltsd fel a VPS-en a **`/opt/raktarapp/`** mappába (a `server_api.py` mellé).
3. Kész – szerver-újraindítás nem kell. A következő ellenőrzéskor (bármelyik gépen a VERZIÓK fül
   „🔄 Ellenőrzés most” gombja) a VPS feltölti az összes eddigi verziót a megadott Drive mappába. A VERZIÓK fülön látszik: ☁️ Drive: ✅ mentve.

> Ha a `drive_mentes.json` nincs a VPS-en, minden ugyanúgy működik, csak a Drive-ra mentés marad ki
> (a VERZIÓK fül kiírja: „Drive mentés nincs beállítva”).
