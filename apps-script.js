/**
 * ================================================================
 *  Google Apps Script — ค่ายสังคมศึกษา กุดบาก
 * ================================================================
 *
 *  วิธีติดตั้ง:
 *  1. เปิด Google Sheet ที่สร้างไว้
 *  2. คลิก Extensions → Apps Script
 *  3. ลบโค้ดเดิมทิ้ง แล้ว paste โค้ดนี้ทั้งหมด
 *  4. กด Deploy → New Deployment
 *     - Type: Web App
 *     - Execute as: Me
 *     - Who has access: Anyone
 *  5. กด Deploy → คัดลอก URL ที่ได้
 *  6. วาง URL นั้นใน kudbak.html และ kudbak_score_board.html
 *     ตรงบรรทัด: const SCRIPT_URL = '...';
 * ================================================================
 */

const TX_SHEET   = 'transactions';
const CFG_SHEET  = 'config';
const TX_HEADERS = ['id', 'studentId', 'type', 'qty', 'priceAtAward', 'valueAtAward', 'timestamp'];

// ── Entry point ──────────────────────────────────────────────
function doGet(e) {
  const action = e.parameter.action || 'getAll';
  let result;
  try {
    if      (action === 'getAll')            result = getAll();
    else if (action === 'addTransaction')    result = addTransaction(JSON.parse(e.parameter.data));
    else if (action === 'deleteTransaction') result = deleteTransaction(e.parameter.id);
    else if (action === 'saveSettings')      result = saveSettings(JSON.parse(e.parameter.data));
    else                                     result = { error: 'Unknown action: ' + action };
  } catch(err) {
    result = { error: err.toString() };
  }
  return ContentService
    .createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

// ── Get all data (transactions + settings) ───────────────────
function getAll() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  return {
    transactions: readTransactions(ss),
    settings:     readSettings(ss),
  };
}

// ── Transactions ─────────────────────────────────────────────
function ensureTxSheet(ss) {
  let sh = ss.getSheetByName(TX_SHEET);
  if (!sh) {
    sh = ss.insertSheet(TX_SHEET);
    sh.appendRow(TX_HEADERS);
    sh.setFrozenRows(1);
    sh.setColumnWidth(1, 160);
    sh.setColumnWidth(7, 200);
  }
  return sh;
}

function readTransactions(ss) {
  const sh = ensureTxSheet(ss);
  const rows = sh.getDataRange().getValues();
  if (rows.length <= 1) return [];
  const headers = rows[0];
  return rows.slice(1)
    .filter(r => r[0])
    .map(r => {
      const obj = {};
      headers.forEach((h, i) => obj[h] = r[i]);
      obj.qty          = parseFloat(obj.qty)          || 0;
      obj.priceAtAward = parseFloat(obj.priceAtAward) || 0;
      obj.valueAtAward = parseFloat(obj.valueAtAward) || 0;
      return obj;
    });
}

function addTransaction(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ensureTxSheet(ss);
  sh.appendRow(TX_HEADERS.map(h => data[h] ?? ''));
  return { success: true, id: data.id };
}

function deleteTransaction(id) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ensureTxSheet(ss);
  const rows = sh.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) === String(id)) {
      sh.deleteRow(i + 1);
      return { success: true };
    }
  }
  return { success: false, error: 'Not found: ' + id };
}

// ── Settings ─────────────────────────────────────────────────
function ensureCfgSheet(ss) {
  let sh = ss.getSheetByName(CFG_SHEET);
  if (!sh) {
    sh = ss.insertSheet(CFG_SHEET);
    sh.appendRow(['key', 'value']);
    sh.appendRow(['landPrice',   100000]);
    sh.appendRow(['rewardValue', 1000]);
    sh.setFrozenRows(1);
  }
  return sh;
}

function readSettings(ss) {
  const sh = ensureCfgSheet(ss);
  const rows = sh.getDataRange().getValues();
  const cfg = { landPrice: 500000, rewardValue: 1000 };
  rows.slice(1).forEach(r => { if (r[0]) cfg[r[0]] = parseFloat(r[1]); });
  return cfg;
}

function saveSettings(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ensureCfgSheet(ss);
  const rows = sh.getDataRange().getValues();
  Object.entries(data).forEach(([key, val]) => {
    let found = false;
    for (let i = 1; i < rows.length; i++) {
      if (rows[i][0] === key) {
        sh.getRange(i + 1, 2).setValue(val);
        found = true; break;
      }
    }
    if (!found) sh.appendRow([key, val]);
  });
  return { success: true };
}
