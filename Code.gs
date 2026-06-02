// ── Finance Tracker — Google Apps Script backend ─────────────────────────
// Vlož celý tento soubor do Apps Scriptu (Rozšíření → Apps Script) své
// Google tabulky. Postup viz README.md.

const SHEET_NAME = 'List 1';      // název listu s transakcemi (sloupce A–G)
const CATS_SHEET = 'Categories';  // list pro uložení kategorií (vytvoří se sám)

// ── HESLO ────────────────────────────────────────────────
// Nastav si vlastní heslo. Stejné zadáváš v aplikaci při přihlášení.
const PASSWORD = 'ZMEN_TOHLE_HESLO';

function checkAuth(e) {
  var token = (e && e.parameter && e.parameter.token) || '';
  if (!token && e && e.postData && e.postData.contents) {
    try { token = JSON.parse(e.postData.contents).token || ''; } catch (err) {}
  }
  return token === PASSWORD;
}

function doGet(e) {
  if (!checkAuth(e)) return json({ ok: false, error: 'unauthorized' });

  const action = e.parameter.action;
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME);

  try {

    // ── GET ALL TRANSACTIONS ─────────────────────────────
    if (action === 'getAll') {
      const rows = sheet.getDataRange().getValues();
      if (rows.length <= 1) return json({ ok: true, data: [] });
      // Sloupce: A=id, B=type, C=desc, D=amount, E=cat, F=date, G=created_at
      const data = rows.slice(1).map(row => ({
        id:     row[0],
        type:   row[1],
        desc:   row[2],
        amount: row[3],
        cat:    row[4],
        date:   row[5]
      }));
      return json({ ok: true, data });
    }

    // ── GET CATEGORIES ───────────────────────────────────
    if (action === 'getCategories') {
      const cs = ss.getSheetByName(CATS_SHEET);
      if (!cs || cs.getLastRow() < 2) return json({ ok: true, data: null });
      const val = cs.getRange(2, 1).getValue();
      return json({ ok: true, data: JSON.parse(val) });
    }

    // ── ADD ──────────────────────────────────────────────
    if (action === 'add') {
      const tx = JSON.parse(e.parameter.tx);
      sheet.appendRow([
        tx.id, tx.type, tx.desc, tx.amount,
        tx.cat, tx.date, new Date().toISOString()
      ]);
      return json({ ok: true });
    }

    // ── DELETE ───────────────────────────────────────────
    if (action === 'delete') {
      const id = e.parameter.id;
      const data = sheet.getDataRange().getValues();
      for (let i = 1; i < data.length; i++) {
        if (String(data[i][0]) === String(id)) {
          sheet.deleteRow(i + 1);
          break;
        }
      }
      return json({ ok: true });
    }

    // ── UPDATE CATEGORY ──────────────────────────────────
    if (action === 'updateCat') {
      const id  = e.parameter.id;
      const cat = e.parameter.cat;
      const data = sheet.getDataRange().getValues();
      for (let i = 1; i < data.length; i++) {
        if (String(data[i][0]) === String(id)) {
          sheet.getRange(i + 1, 5).setValue(cat); // sloupec E = cat
          break;
        }
      }
      return json({ ok: true });
    }

    // ── SAVE ALL via GET (small datasets) ────────────────
    if (action === 'saveAll') {
      const txs = JSON.parse(e.parameter.txs);
      return saveAllTxs(sheet, txs);
    }

  } catch(err) {
    return json({ ok: false, error: err.toString() });
  }

  return json({ ok: false, error: 'Unknown action: ' + action });
}

function doPost(e) {
  if (!checkAuth(e)) return json({ ok: false, error: 'unauthorized' });

  try {
    const body = JSON.parse(e.postData.contents);
    const action = body.action;
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_NAME);

    // ── UPDATE CATEGORY via POST ─────────────────────────
    if (action === 'updateCat') {
      const id  = String(body.id);
      const cat = body.cat;
      const data = sheet.getDataRange().getValues();
      for (let i = 1; i < data.length; i++) {
        if (String(data[i][0]) === id) {
          sheet.getRange(i + 1, 5).setValue(cat);
          return json({ ok: true });
        }
      }
      return json({ ok: false, error: 'ID not found: ' + id });
    }

    // ── UPDATE DESCRIPTION via POST ──────────────────────
    if (action === 'updateDesc') {
      const id   = String(body.id);
      const desc = body.desc;
      const data = sheet.getDataRange().getValues();
      for (let i = 1; i < data.length; i++) {
        if (String(data[i][0]) === id) {
          sheet.getRange(i + 1, 3).setValue(desc); // sloupec C = desc
          return json({ ok: true });
        }
      }
      return json({ ok: false, error: 'ID not found: ' + id });
    }

    // ── SAVE ALL via POST (large datasets) ───────────────
    if (action === 'saveAll') {
      return saveAllTxs(sheet, body.txs);
    }

    // ── ADD via POST ─────────────────────────────────────
    if (action === 'add') {
      const tx = body.tx;
      sheet.appendRow([
        tx.id, tx.type, tx.desc, tx.amount,
        tx.cat, tx.date, new Date().toISOString()
      ]);
      return json({ ok: true });
    }

    // ── SAVE CATEGORIES via POST ─────────────────────────
    if (action === 'saveCategories') {
      const cats = body.cats;
      let cs = ss.getSheetByName(CATS_SHEET);
      if (!cs) {
        cs = ss.insertSheet(CATS_SHEET);
        cs.appendRow(['categories_json']);
      }
      if (cs.getLastRow() < 2) {
        cs.appendRow([JSON.stringify(cats)]);
      } else {
        cs.getRange(2, 1).setValue(JSON.stringify(cats));
      }
      return json({ ok: true });
    }

  } catch(err) {
    return json({ ok: false, error: err.toString() });
  }

  return json({ ok: false, error: 'Unknown POST action' });
}

function saveAllTxs(sheet, txs) {
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) sheet.deleteRows(2, lastRow - 1);

  const rows = txs.map(tx => [
    tx.id, tx.type, tx.desc, tx.amount,
    tx.cat, tx.date, new Date().toISOString()
  ]);
  if (rows.length > 0) {
    sheet.getRange(2, 1, rows.length, 7).setValues(rows);
  }
  return json({ ok: true, count: rows.length });
}

function json(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
