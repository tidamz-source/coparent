/**
 * Coparentalité — serveur de synchronisation (Google Apps Script)
 * À coller dans Extensions › Apps Script d'un Google Sheet, puis
 * Déployer › Nouveau déploiement › Application Web
 *   Exécuter en tant que : Moi
 *   Qui a accès : Tout le monde
 */
const SECRET = 'REMPLACE-MOI-PAR-UN-CODE-SECRET';
const SHEET = 'donnees';

function doPost(e) {
  let body;
  try { body = JSON.parse(e.postData.contents); } catch (x) { return out({ ok: false, error: 'Requête invalide' }); }
  if (body.secret !== SECRET) return out({ ok: false, error: 'secret' });

  // Documents (attestation de mutuelle…) : rangés dans un dossier de ton Google Drive, jamais partagés publiquement
  if (body.action === 'upload') {
    const blob = Utilities.newBlob(Utilities.base64Decode(body.data), body.mime || 'application/pdf', body.name || 'document.pdf');
    const f = docFolder().createFile(blob);
    return out({ ok: true, id: f.getId(), name: f.getName() });
  }
  if (body.action === 'get') {
    const f = docFile(body.id);
    if (!f) return out({ ok: false, error: 'Document introuvable' });
    const b = f.getBlob();
    return out({ ok: true, data: Utilities.base64Encode(b.getBytes()), mime: b.getContentType(), name: f.getName() });
  }
  if (body.action === 'delete') {
    const f = docFile(body.id);
    if (f) f.setTrashed(true);
    return out({ ok: true });
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sh = ss.getSheetByName(SHEET);
    if (!sh) {
      sh = ss.insertSheet(SHEET);
      sh.getRange(1, 1, 1, 6).setValues([['clé', 'type', 'id', 'données', 'modifié', 'supprimé']]).setFontWeight('bold');
      sh.setFrozenRows(1);
    }
    const n = sh.getLastRow() - 1;
    const rows = n > 0 ? sh.getRange(2, 1, n, 6).getValues() : [];
    const idx = {};
    rows.forEach((r, i) => { idx[r[0]] = i; });

    let dirty = false;
    (body.changes || []).forEach(c => {
      const row = [c.k, c.c, c.id, c.del ? '' : c.d, c.u, c.del ? 1 : 0];
      if (c.k in idx) {
        const i = idx[c.k];
        if (Number(rows[i][4]) < c.u) { rows[i] = row; dirty = true; }
      } else {
        idx[c.k] = rows.length; rows.push(row); dirty = true;
      }
    });
    if (dirty && rows.length) {
      const rg = sh.getRange(2, 1, rows.length, 6);
      sh.getRange(2, 1, rows.length, 4).setNumberFormat('@');
      rg.setValues(rows);
    }
    return out({
      ok: true,
      rows: rows.map(r => ({ k: String(r[0]), c: String(r[1]), id: String(r[2]), d: String(r[3]), u: Number(r[4]), del: Number(r[5]) === 1 }))
    });
  } finally {
    lock.releaseLock();
  }
}

function doGet() {
  return ContentService.createTextOutput('Coparentalité : le script fonctionne.');
}

const FOLDER = 'Coparentalité – documents';
function docFolder() {
  const it = DriveApp.getFoldersByName(FOLDER);
  return it.hasNext() ? it.next() : DriveApp.createFolder(FOLDER);
}
// Seuls les fichiers du dossier de l'app sont accessibles
function docFile(id) {
  try {
    const f = DriveApp.getFileById(id);
    const parents = f.getParents();
    const folderId = docFolder().getId();
    while (parents.hasNext()) if (parents.next().getId() === folderId) return f;
  } catch (x) {}
  return null;
}

function out(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
