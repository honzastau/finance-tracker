const SHEET_NAME = 'List 1';
const CATS_SHEET = 'Categories';

// ── HESLO ────────────────────────────────────────────────
// Nastav si vlastní heslo. Stejné zadáváš v aplikaci při přihlášení.
// V repozitáři je záměrně jen placeholder — skutečné heslo drž pouze
// v Apps Script editoru, repozitář je veřejný.
const PASSWORD = 'ZMEN_TOHLE_HESLO';

function checkAuth(e) {
  var token = (e && e.parameter && e.parameter.token) || '';
  if (!token && e && e.postData && e.postData.contents) {
    try { token = JSON.parse(e.postData.contents).token || ''; } catch (err) {}
  }
  return token === PASSWORD;
}

// Vrátí mapu id, která už v listu jsou. Web app občas nedoručí odpověď na
// zápis, který ve skutečnosti proběhl — klient pak pošle stejnou transakci
// znovu. Díky kontrole id se takový retry nezapíše podruhé.
function existingIdMap(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return {};
  return sheet.getRange(2, 1, lastRow - 1, 1).getValues().reduce(function(map, row) {
    map[String(row[0])] = true;
    return map;
  }, {});
}

function txRow(tx, stamp) {
  return [String(tx.id), tx.type, tx.desc, tx.amount, tx.cat, tx.date, stamp];
}

function addOne(sheet, tx) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    if (existingIdMap(sheet)[String(tx.id)]) return json({ ok: true, added: 0, skipped: 1 });
    sheet.appendRow(txRow(tx, new Date().toISOString()));
    return json({ ok: true, added: 1, skipped: 0 });
  } finally {
    lock.releaseLock();
  }
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
      // Map by fixed column positions: A=id, B=type, C=desc, D=amount, E=cat, F=date, G=created_at
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
      return addOne(sheet, JSON.parse(e.parameter.tx));
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
          sheet.getRange(i + 1, 5).setValue(cat); // column E = cat
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
          sheet.getRange(i + 1, 3).setValue(desc); // column C = desc
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
      return addOne(sheet, body.tx);
    }

    // ── ADD BATCH via POST ───────────────────────────────
    // Celý import jedním zápisem. Dřív klient posílal request na každou
    // transakci zvlášť, což u výpisu o desítkách řádků trvalo minuty a
    // v půlce naráželo na limity Apps Scriptu.
    if (action === 'addBatch') {
      const txs = body.txs || [];
      if (!txs.length) return json({ ok: true, added: 0, skipped: 0 });

      const lock = LockService.getScriptLock();
      lock.waitLock(30000);
      try {
        const seen = existingIdMap(sheet);
        const stamp = new Date().toISOString();
        const rows = [];
        txs.forEach(function(tx) {
          const id = String(tx.id);
          if (seen[id]) return;
          seen[id] = true;
          rows.push(txRow(tx, stamp));
        });
        if (rows.length) {
          sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, 7).setValues(rows);
        }
        return json({ ok: true, added: rows.length, skipped: txs.length - rows.length });
      } finally {
        lock.releaseLock();
      }
    }

    // ── SAVE CATEGORIES via POST ─────────────────────────
    if (action === 'saveCategories') {
      const cats = body.cats;
      let cs = ss.getSheetByName(CATS_SHEET);
      if (!cs) {
        cs = ss.insertSheet(CATS_SHEET);
        cs.appendRow(['categories_json']);
      }
      // Always store in row 2
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

  const rows = txs.map(tx => txRow(tx, new Date().toISOString()));
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
