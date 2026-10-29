/**
 * RAKTÁRPROGRAM – RÉGI VERZIÓK MENTÉSE A SAJÁT GOOGLE DRIVE-RA
 *
 * Ezt a kódot a script.google.com-on kell beilleszteni (lásd vps/TELEPITES_verziok.md).
 * A VPS ide küldi a raktárprogram (app_pyqt6CLOUD.py) minden verzióját, ez pedig elmenti
 * a Drive-od "Raktárprogram régi verziók" mappájába. Egy fájlt csak egyszer ment el.
 *
 * FONTOS: az alábbi TITOK legyen ugyanaz, mint a VPS-en a /opt/raktarapp/drive_mentes.json "titok" értéke!
 */
const TITOK = 'IDE-IRJ-EGY-HOSSZU-JELSZOT';
// Ha a VPS beállításában (drive_mentes.json) van "mappa_id", abba a mappába ment; ha nincs, ebbe a nevű mappába:
const MAPPA_NEV = 'Raktárprogram régi verziók';

function doPost(e) {
  try {
    const d = JSON.parse(e.postData.contents);
    if (d.titok !== TITOK) return valasz({ status: 'error', message: 'Rossz titok (jelszó)' });
    let mappa;
    if (d.mappa_id) {
      mappa = DriveApp.getFolderById(String(d.mappa_id));
    } else {
      const mappak = DriveApp.getFoldersByName(MAPPA_NEV);
      mappa = mappak.hasNext() ? mappak.next() : DriveApp.createFolder(MAPPA_NEV);
    }
    const nev = String(d.nev || 'app_pyqt6CLOUD.py');
    const megvan = mappa.getFilesByName(nev);
    if (megvan.hasNext()) return valasz({ status: 'ok', mar_megvolt: true, url: megvan.next().getUrl(), mappa: mappa.getUrl() });
    const blob = Utilities.newBlob(Utilities.base64Decode(d.tartalom_b64), 'text/plain', nev);
    const fajl = mappa.createFile(blob);
    if (d.leiras) fajl.setDescription(String(d.leiras));
    return valasz({ status: 'ok', url: fajl.getUrl(), mappa: mappa.getUrl() });
  } catch (err) {
    return valasz({ status: 'error', message: String(err) });
  }
}

// Böngészőből megnyitva ezt írja ki (így lehet ellenőrizni, hogy él-e)
function doGet() {
  return valasz({ status: 'ok', message: 'A raktárprogram Drive mentés működik.' });
}

function valasz(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
