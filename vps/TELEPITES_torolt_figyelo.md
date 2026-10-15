# Törölt sorok figyelő – telepítés a VPS-re

> A `torolt_figyelo.py` szándékosan ékezet nélküli, így a WinSCP szerkesztőben is helyesen jelenik meg.

A `torolt_figyelo.py` éjjel-nappal figyeli az élő Google táblázatot. Ha onnan eltűnik egy R18-as sor, vagy kiürítik a Márka, Típus és Cikkszám mezőit, a termék korábbi teljes adata a **ToroltSor** munkalapra kerül.
Az élő táblázathoz **csak olvasási joggal** kapcsolódik, abba semmit nem ír.

## 1. Fájlok feltöltése (WinSCP)
Töltsd fel ezt a két fájlt a `/opt/raktarapp/` mappába:
- `torolt_figyelo.py`
- `torolt-figyelo.service`

A `credentials.json` már ott van, ugyanazt használja.

## 2. Beállítás
Nyisd meg WinSCP-ben a `/opt/raktarapp/torolt_figyelo.py` fájlt (dupla kattintás), és a fájl elején töltsd ki:
```python
WATCH_SPREADSHEET_ID = "az-élő-táblázat-ID-ja"
WATCH_WORKSHEET = ""            # munkalap neve; üresen = az első lap
DELETED_SPREADSHEET_ID = "..."  # ahol a ToroltSor van (NEM lehet az élő táblázat!)
```
Az élő táblázatot oszd meg a `credentials.json`-ban lévő `client_email` címmel. Elég a **Megtekintő** jog.

## 3. Telepítés (WinSCP: Parancsok → Konzol, vagy PuTTY)
```bash
cd /opt/raktarapp
venv/bin/pip install gspread google-auth          # ha még nincs fent
venv/bin/python torolt_figyelo.py --egyszer       # próba: egyszer lefut, kiírja, mit látott
cp torolt-figyelo.service /etc/systemd/system/
systemctl daemon-reload
systemctl enable --now torolt-figyelo
```

## 4. Ellenőrzés
```bash
systemctl status torolt-figyelo                   # fut-e ("active (running)")
tail -n 50 /opt/raktarapp/naplok/torolt_figyelo.log
```
Az első kör csak megjegyzi a táblázat állapotát („Első ellenőrzés: … R18-sor megjegyezve”), utána figyeli a változásokat.

## Hasznos parancsok
| Mit | Parancs |
|---|---|
| Újraindítás (beállítás módosítása után) | `systemctl restart torolt-figyelo` |
| Leállítás | `systemctl stop torolt-figyelo` |
| Kikapcsolás végleg | `systemctl disable --now torolt-figyelo` |
| Élő napló | `journalctl -u torolt-figyelo -f` |

A program újraindításkor és áramszünet után is magától elindul. Ha nincs internet, a ToroltSor-bejegyzések helyben várnak (`torolt_figyelo.db`), és később kimennek.

## „Ki törölte” (S oszlop) beállítása
A figyelő a Google Drive **verzióelőzményeiből** olvassa ki, ki szerkesztette a táblázatot a törlés idején, és az email-címet a ToroltSor **S** oszlopába írja. Ehhez is **csak olvasási** jogot kér (`drive.metadata.readonly`), a figyelt táblázatot továbbra sem tudja módosítani.

1. **Google Drive API bekapcsolása** (egyszer): https://console.cloud.google.com → felül válaszd ki azt a projektet, amelyikhez a `credentials.json` tartozik → *APIs & Services* → *Library* → **Google Drive API** → **Enable**.
2. **A szolgáltatásfiók lássa a verzióelőzményeket:** a Google csak a táblázat *szerkesztőinek* mutatja a verzióelőzményeket. Ha a ToroltSor S oszlopában „(verzióelőzmény nem olvasható)” áll, a figyelt táblázatot oszd meg a szolgáltatásfiókkal (`client_email` a credentials.json-ban) **Szerkesztő** joggal. A program ettől sem tud írni bele: csak olvasási engedéllyel (scope) kapcsolódik.
3. `systemctl restart torolt-figyelo`

Tudnivalók:
- A Google a verziókat néhány percre összevonja, ezért a bejegyzés legfeljebb ~10 percet vár, mire a ToroltSor-ba kerül. Ha ugyanakkor többen szerkesztettek, mindenki email-címe bekerül (vesszővel).
- Ha nem derül ki: „(nem állapítható meg)”. A régebbi, már várakozó bejegyzéseknél az S oszlop üres marad.
