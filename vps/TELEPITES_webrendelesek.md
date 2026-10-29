# 🛒 Webes rendelések (Shopify + eBay) – beállítás lépésről lépésre

## Hogyan működik
- A VPS **5 percenként** lekéri a Shopify és az eBay rendeléseit (az elmúlt 30 napból), és megkeresi,
  melyik raktári termék az (eBay ID = N oszlop, SKU = R18 kód, vagy R18 kód a termék nevében).
- A gépi programban a **🛒 Webes rendelések** fülön látszik minden rendelés (vevő, cím, ár, R18, készlet).
  A fül nevén ott a rögzítésre váró rendelések száma, új rendelésnél értesítés jön.
- **💰 Rögzítés eladásként** (vagy dupla kattintás): rákérdez, majd levonja a készletet, beírja az Eladottak közé
  és a Google táblázat eladási lapjára (eBay → Eladas_Ebay, Shopify → Eladas_Webshop).
  Egy rendelést csak egyszer lehet rögzíteni (akkor sem kétszer, ha két gépen nyomják meg).
- Ami nem párosult: **🔗 Kézi párosítás** (R18 kód beírása / csippantása). Ami nem raktári: **🚫 Kihagyás**.

A kulcsok/jelszavak csak a VPS-en vannak, egy fájlban: **`/opt/raktarapp/webshop_api.json`**
(a GitHub-ra soha nem kerül fel).

---

## 1. Shopify (kb. 10 perc)

> 2026. január 1. óta a Shopify adminban már nem lehet régi típusú „egyéni alkalmazást” (custom app) létrehozni.
> Helyette a **Dev Dashboard**-ban kell. Ha már van régi egyéni alkalmazásotok `shpat_...` tokennel, az is jó (lásd 1/e).
> A Shopify ezeket a menüpontokat időnként átnevezi – ha valami máshogy hívják, a lényeg ugyanaz.

**1/a.** Jelentkezz be a Shopify bolt adminjába tulajdonosként. Írd fel a bolt **myshopify** címét
(Beállítások → Domainek, pl. `raktar-shop.myshopify.com`).

**1/b.** Nyisd meg a **Dev Dashboard**-ot: https://dev.shopify.com/dashboard (ugyanazzal a fiókkal).
→ **Create app** (Alkalmazás létrehozása) → név: `Raktarprogram`.

**1/c.** Az alkalmazásnál a **Versions / Configuration** részen:
- **Admin API hozzáférési jogok (scopes):** `read_orders` (ha van: `read_products` is). Mást ne adj!
- Mentsd el, és adj ki egy verziót (**Release**).

**1/d.** **Install app** → válaszd ki a saját boltodat → telepítés.
Utána az alkalmazás **Settings / Credentials** oldalán másold ki:
- **Client ID**
- **Client secret**

**1/e.** (Csak ha régi egyéni alkalmazásod van `shpat_...` kezdetű Admin API tokennel: azt használhatod
a client ID / secret helyett – lásd lent a `"token"` sort.)

---

## 2. eBay (kb. 20 perc, egyszer kell)

**2/a.** Regisztrálj (ingyenes): https://developer.ebay.com → **Join** – lehet ugyanaz az e-mail, mint az eladói fiók.

**2/b.** **Application Keys** (Hi <név> → Application Keysets) → a **Production** oszlopban **Create a keyset**.
Írd fel: **App ID (Client ID)** és **Cert ID (Client Secret)**.

**2/c. ⚠️ Fiók-törlési értesítés (eBay kötelező szabálya).**
Az éles kulcs csak akkor működik, ha a fejlesztői fiókban be van állítva a **Marketplace Account Deletion** értesítés
(Alerts & Notifications). Mivel a program elmenti a vevők nevét és címét, a „nem tárolok eBay adatot” kivétel **nem**
igaz ránk – az értesítést kell beállítani. Ehhez az eBay **HTTPS** címet kér, a VPS jelenleg sima HTTP-n fut.
→ Ehhez a VPS-re HTTPS kell (domain + ingyenes tanúsítvány) – **szólj, és ebben segítek**; utána:
- Endpoint: `https://<a-domain>/ebay/fiok_torles`
- Verification token: egy 32-80 karakteres saját jelszó (betű, szám, `_`, `-`)
- Ugyanezt a kettőt be kell írni a `webshop_api.json`-ba (`torles_url`, `torles_token`), majd **Send Test Notification**.

**2/d.** **User Tokens** (Application Keys oldalon a Production keyset mellett) → **Get a Token from eBay via Your Application**
→ **Add eBay Redirect URL**:
- Display Title: `Raktarprogram`
- Privacy Policy URL: a webshop adatkezelési oldala
- **Auth Accepted URL:** `https://www.google.com` (bármilyen https oldal jó – csak a címsorból kell majd kimásolni a kódot)
- Auth Declined URL: `https://www.google.com`
- Pipáld be: **OAuth Enabled** → Save.
- Írd fel a **RuName** értéket (pl. `Kovacs_Peter-Raktarpr-PRD-abc123-def456`).

---

## 3. A beállító fájl a VPS-en

A saját gépeden hozz létre Jegyzettömbbel egy **`webshop_api.json`** fájlt (Mentés másként → Fájltípus: **Minden fájl**),
ezzel a tartalommal – a nagybetűs részeket cseréld ki (ha csak az egyik kell, a másik részt töröld ki):

```json
{
  "shopify": {"bolt": "IDE-A-BOLT.myshopify.com", "client_id": "IDE-A-CLIENT-ID", "client_secret": "IDE-A-CLIENT-SECRET"},
  "ebay": {"client_id": "IDE-AZ-APP-ID", "client_secret": "IDE-A-CERT-ID", "runame": "IDE-A-RUNAME",
           "torles_token": "", "torles_url": ""},
  "napok": 30
}
```
(Régi Shopify tokennél: `"shopify": {"bolt": "IDE-A-BOLT.myshopify.com", "token": "shpat_..."}`)

WinSCP-vel töltsd fel ide: **`/opt/raktarapp/webshop_api.json`**, majd: `sudo systemctl restart raktarapp`

---

## 4. eBay engedélyezés (egyszer, böngészőből)

1. Nyisd meg a böngészőben: `http://<VPS címe>:5000/ebay/auth?t=<az api_token.txt tartalma>`
2. Kattints: **eBay engedélyezés megnyitása** → lépj be az **eladói** eBay fiókkal → **Agree / Elfogadom**.
3. A böngésző a google.com-ra visz. A **címsorból másold ki a TELJES címet** (benne van: `code=...`),
   és **5 percen belül** illeszd be az előző oldal mezőjébe → **Mentés**.
4. „KÉSZ! Az eBay engedélyezve.” – ez kb. 18 hónapig érvényes, utána ezt a 4. lépést meg kell ismételni.

---

## 5. Ellenőrzés
A gépi programban: **🛒 Webes rendelések** fül → **🔄 Lekérés most**. Fent ezt kell látni:
`✅ Shopify · ✅ eBay · Utolsó lekérés: ...` – és megjelennek az elmúlt 30 nap rendelései.
Ha ❌ jelenik meg, a mellette lévő hibaüzenetet küldd el.
