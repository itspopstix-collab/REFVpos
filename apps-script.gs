/**
 * REFV POS - Google Apps Script backend.
 * Paste into the Sheet: Extensions > Apps Script, then Deploy > New deployment >
 * Web app (Execute as: Me, Who has access: Anyone). Copy the /exec URL into the POS Settings.
 *
 * Sheets used (auto-created):
 *   "Sales"    - one row per cup sold
 *   "Settings" - key/value (stores PromptPay number)
 */
const SALES = 'Sales';
const SETTINGS = 'Settings';

function sheet_(name) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  return ss.getSheetByName(name) || ss.insertSheet(name);
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function getPP_() {
  const sh = sheet_(SETTINGS), v = sh.getDataRange().getValues();
  for (const r of v) if (r[0] === 'pp') return String(r[1] || '');
  return '';
}

function setPP_(pp) {
  const sh = sheet_(SETTINGS), v = sh.getDataRange().getValues();
  for (let i = 0; i < v.length; i++) {
    if (v[i][0] === 'pp') { sh.getRange(i + 1, 2).setNumberFormat('@').setValue(pp); return; }
  }
  sh.appendRow(['pp', '']);
  sh.getRange(sh.getLastRow(), 2).setNumberFormat('@').setValue(pp);
}

function doGet(e) {
  const mode = e && e.parameter && e.parameter.mode;
  if (mode === 'config') return json_({ pp: getPP_() });
  if (mode === 'report') return json_({ sales: readSales_() });
  return ContentService.createTextOutput('REFV POS ready');
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const data = JSON.parse(e.postData.contents);
    if (data.type === 'config') {
      setPP_(String(data.pp || ''));
      return ContentService.createTextOutput('ok');
    }
    const rows = data.rows || [], head = data.head || [];
    if (!rows.length) return ContentService.createTextOutput('ok');
    const sh = sheet_(SALES);
    if (sh.getLastRow() === 0) {
      sh.appendRow(head);
      sh.setFrozenRows(1);
    }
    // Skip already-stored orders so "Re-send all" never duplicates.
    const have = {};
    if (sh.getLastRow() > 1) {
      sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues().forEach(r => have[r[0]] = true);
    }
    const fresh = rows.filter(r => !have[r[0]]);
    if (fresh.length) {
      // Force plain text for ID/date/time so Sheets does not reformat them.
      const start = sh.getLastRow() + 1;
      sh.getRange(start, 1, fresh.length, 3).setNumberFormat('@');
      sh.getRange(start, 1, fresh.length, fresh[0].length).setValues(fresh);
    }
    return ContentService.createTextOutput('ok');
  } catch (err) {
    return ContentService.createTextOutput('error ' + err);
  } finally {
    lock.releaseLock();
  }
}

/** Rebuild order objects (as the POS expects) from the Sales rows. */
function readSales_() {
  const sh = sheet_(SALES);
  if (sh.getLastRow() < 2) return [];
  const rows = sh.getRange(2, 1, sh.getLastRow() - 1, 11).getValues();
  const map = {}, order = [];
  rows.forEach(r => {
    const [id, date, time, type, cups, f1, f2, c1, c2, pay, amt] = r.map(x => x);
    if (!id) return;
    const key = String(id);
    if (!map[key]) {
      const d = new Date(String(date) + 'T' + String(time) + '+07:00');
      map[key] = {
        id: key, ts: isNaN(d) ? Date.now() : d.getTime(), items: [], total: 0,
        pay: String(pay).toUpperCase() === 'QR' ? 'qr' : 'cash', synced: 1
      };
      order.push(key);
    }
    const price = Number(amt) || 0, n = Number(cups) || 1;
    const parts = [{ name: f1, cat: c1, price: n === 2 ? price / 2 : price }];
    if (f2) parts.push({ name: f2, cat: c2, price: price / 2 });
    map[key].items.push({ type: String(type), cups: n, drink: '', parts, price });
    map[key].total += price;
  });
  return order.map(k => { map[k].recv = map[k].total; return map[k]; });
}
