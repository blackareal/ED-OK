/**
 * ================================================================
 *  Google Apps Script — ระบบคะแนน + ระบบภารกิจเซลฟี่คณิต SWP Math Camp12
 *  (ดัดแปลงจากระบบของค่ายกุดบาก ตัดระบบโหวตโปสเตอร์ออก
 *   ภายหลังเพิ่มระบบภารกิจเซลฟี่ (แนว Selfie Vocab Run) เข้ามาในสคริปต์เดียวกัน
 *   ใช้ Google Sheet เดิม แค่เพิ่มแท็บใหม่ชื่อ "missions")
 * ================================================================
 *
 *  วิธีติดตั้ง (ระบบคะแนน):
 *  1. สร้าง Google Sheet ใหม่ (สร้างเปล่าๆ ก็พอ ระบบจะสร้างชีตย่อยให้เอง)
 *  2. คลิก Extensions → Apps Script
 *  3. ลบโค้ดเดิมทิ้ง แล้ว paste โค้ดนี้ทั้งหมด
 *  4. กด Deploy → New Deployment
 *     - Type: Web App
 *     - Execute as: Me
 *     - Who has access: Anyone
 *  5. กด Deploy → คัดลอก URL ที่ได้
 *  6. วาง URL นั้นใน score_input.html และ score_board.html
 *     ตรงบรรทัด: const SCRIPT_URL = '...';
 *
 *  ⚠️ ถ้าเคย Deploy ไปแล้วรอบหนึ่ง (มี URL ใช้งานอยู่แล้ว) แล้วมาวางโค้ดนี้ทับ
 *  (เวอร์ชันที่เพิ่มระบบภารกิจเซลฟี่เข้ามา) ต้องกด Deploy → Manage deployments →
 *  ไอคอนดินสอ (แก้ไข) → Version: New version → Deploy ใหม่อีกครั้ง ไม่เช่นนั้น
 *  Web App จะยังรันโค้ดเวอร์ชันเก่าอยู่ (ลิงก์ไม่เปลี่ยน แต่ฟังก์ชันใหม่จะยังไม่ทำงาน)
 * ================================================================
 */

const TX_SHEET   = 'transactions';
const CFG_SHEET  = 'config';
const TX_HEADERS = ['id', 'studentId', 'type', 'qty', 'priceAtAward', 'valueAtAward', 'timestamp', 'station'];

// ── ระบบภารกิจเซลฟี่คณิต (Selfie Mission) ─────────────────────
const MISSION_SHEET       = 'missions';
const MISSION_FOLDER_NAME = 'SWP_MathCamp12_Selfie_Photos';
const TOTAL_MISSIONS      = 30;
const TEAM_COL_NAME       = 'ทีม';

// ── Entry point (GET) ────────────────────────────────────────
function doGet(e) {
  const action = e.parameter.action || 'getAll';
  let result;
  try {
    if      (action === 'getAll')            result = getAll();
    else if (action === 'addTransaction')    result = addTransaction(JSON.parse(e.parameter.data));
    else if (action === 'deleteTransaction') result = deleteTransaction(e.parameter.id);
    else if (action === 'saveSettings')      result = saveSettings(JSON.parse(e.parameter.data));
    else if (action === 'getMissions')       result = getMissions();
    else                                     result = { error: 'Unknown action: ' + action };
  } catch(err) {
    result = { error: err.toString() };
  }
  return ContentService
    .createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

// ── Entry point (POST) — ใช้เฉพาะอัปโหลดภาพภารกิจเซลฟี่ ───────
// (ระบบคะแนนใช้ doGet ล้วน ไม่ยุ่งกับ doPost นี้)
function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);
  } catch (err) {
    return jsonOutput({ status: 'error', message: 'ระบบกำลังประมวลผลคำขออื่นอยู่ ลองใหม่อีกครั้งในอีกสักครู่' });
  }
  try {
    const data = JSON.parse(e.postData.contents);
    const team = (data.team || '').toString().trim();
    const missionNo = parseInt(data.missionNo, 10);
    const imageBase64 = data.imageBase64;
    const filename = (data.filename || ('MISSION' + missionNo + '.jpg')).toString();
    const mimeType = (data.mimeType || 'image/jpeg').toString();

    if (!team) return jsonOutput({ status: 'error', message: 'กรุณาระบุทีม' });
    if (!missionNo || missionNo < 1 || missionNo > TOTAL_MISSIONS) {
      return jsonOutput({ status: 'error', message: 'หมายเลขภารกิจไม่ถูกต้อง' });
    }
    if (!imageBase64) return jsonOutput({ status: 'error', message: 'ไม่พบข้อมูลรูปภาพ' });

    const bytes = Utilities.base64Decode(imageBase64);
    const blob = Utilities.newBlob(bytes, mimeType, filename);
    const folder = getOrCreateMissionFolder();
    const file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    const fileUrl = 'https://drive.google.com/file/d/' + file.getId() + '/view';

    const sh = ensureMissionSheet();
    const teamColIndex = findExactColIndex(sh, TEAM_COL_NAME);
    const rowIndex = findOrCreateTeamRow(sh, team, teamColIndex);
    const missionColIndex = findExactColIndex(sh, 'MISSION#' + missionNo);
    sh.getRange(rowIndex, missionColIndex + 1).setValue(fileUrl);

    return jsonOutput({ status: 'ok', fileUrl: fileUrl });
  } catch (err) {
    return jsonOutput({ status: 'error', message: err.message });
  } finally {
    lock.releaseLock();
  }
}

function jsonOutput(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function getOrCreateMissionFolder() {
  const folders = DriveApp.getFoldersByName(MISSION_FOLDER_NAME);
  if (folders.hasNext()) return folders.next();
  return DriveApp.createFolder(MISSION_FOLDER_NAME);
}

function buildMissionHeaders() {
  const headers = [TEAM_COL_NAME];
  for (let i = 1; i <= TOTAL_MISSIONS; i++) headers.push('MISSION#' + i);
  return headers;
}

// สร้างแท็บ "missions" ถ้ายังไม่มี (ไม่ยุ่งกับแท็บ transactions/config เดิมเลย)
function ensureMissionSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(MISSION_SHEET);
  if (!sh) {
    sh = ss.insertSheet(MISSION_SHEET);
    sh.appendRow(buildMissionHeaders());
    sh.setFrozenRows(1);
  }
  return sh;
}

// หาคอลัมน์ตรงชื่อเป๊ะ (กัน MISSION#1 ไปจับ MISSION#10-19 ผิด) — ถ้าไม่พบ จะเพิ่ม
// คอลัมน์ใหม่ต่อท้ายให้อัตโนมัติ
function findExactColIndex(sh, name) {
  const lastCol = Math.max(sh.getLastColumn(), 1);
  const headers = sh.getRange(1, 1, 1, lastCol).getValues()[0];
  const idx = headers.findIndex(h => h.toString().trim() === name);
  if (idx >= 0) return idx;
  sh.getRange(1, headers.length + 1).setValue(name);
  return headers.length;
}

// หาแถวของทีมนี้ ถ้ายังไม่มีแถวของทีมนี้ ให้สร้างแถวใหม่
function findOrCreateTeamRow(sh, team, teamColIndex) {
  const lastRow = sh.getLastRow();
  if (lastRow >= 2) {
    const teamColValues = sh.getRange(2, teamColIndex + 1, lastRow - 1, 1).getValues();
    for (let i = 0; i < teamColValues.length; i++) {
      if ((teamColValues[i][0] || '').toString().trim() === team) {
        return i + 2;
      }
    }
  }
  const newRow = lastRow + 1;
  sh.getRange(newRow, teamColIndex + 1).setValue(team);
  return newRow;
}

// ── อ่านข้อมูลภารกิจทั้งหมด สำหรับหน้า Dashboard/หน้าเช็คสถานะของนักเรียน ──
function getMissions() {
  const sh = ensureMissionSheet();
  const rows = sh.getDataRange().getValues();
  if (rows.length < 1) return { headers: [], rows: [] };
  const headers = rows[0].map(h => h.toString().trim());
  const dataRows = rows.slice(1).filter(r => r.some(c => c !== '' && c !== null));
  return { headers, rows: dataRows };
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
    sh.appendRow(['winnerValue', 100000]);
    sh.appendRow(['loserValue',  80000]);
    sh.setFrozenRows(1);
  }
  return sh;
}

function readSettings(ss) {
  const sh = ensureCfgSheet(ss);
  const rows = sh.getDataRange().getValues();
  const cfg = { landPrice: 100000, winnerValue: 100000, loserValue: 80000 };
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
