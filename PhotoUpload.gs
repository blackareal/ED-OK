// Google Apps Script — Web App backend สำหรับอัปโหลดภาพ Selfie Vocab Run
// (แทนที่ Google Form เดิม — อัปโหลดตรงจากหน้า clues.html ไม่ต้องสลับไปกรอกฟอร์มแยก)
//
// วิธีติดตั้ง (ทำครั้งเดียว):
// 1. เปิด Google Sheet ที่จะใช้เก็บผลงาน (จะสร้างชีตใหม่ หรือใช้ของเดิมที่ผูกกับ
//    screen.html อยู่แล้วก็ได้ — ถ้าใช้ของเดิม ตรวจว่าชื่อแท็บตรงกับ SHEET_NAME
//    ด้านล่าง ไม่ตรงก็แก้ค่าคงที่ให้ตรงได้เลย)
// 2. เมนู Extensions > Apps Script แล้ววางโค้ดทั้งหมดนี้ทับของเดิม
// 3. กด Deploy > New deployment > เลือกประเภท "Web app"
//    - Execute as: Me
//    - Who has access: Anyone
//    - กด Deploy แล้วอนุญาตสิทธิ์ที่ขอ (จะขอสิทธิ์ Google Drive เพิ่มจากปกติ
//      เพราะสคริปต์นี้ต้องอัปโหลดไฟล์รูปเข้า Drive ให้)
// 4. คัดลอก "Web app URL" (ลงท้ายด้วย /exec) ไปวางแทนที่ WEBAPP_URL ในไฟล์ clues.html
// 5. Publish แท็บชีตนี้เป็น CSV ตามปกติ (เลือกแท็บนี้เจาะจง ห้ามเลือก "ทั้งเอกสาร")
//    แล้วเอาลิงก์ไปวางใน SHEET_CSV_URL ของไฟล์ screen.html — โครงสร้างคอลัมน์
//    ("ชื่อกลุ่ม" + "SVR#1"..."SVR#32") เหมือนของเดิมทุกอย่าง ไม่ต้องแก้ screen.html เลย
// 6. ทุกครั้งที่แก้โค้ดนี้ในอนาคต ต้องกด Deploy > Manage deployments > กดไอคอนดินสอ
//    > เลือก Version: New version > Deploy ใหม่ทุกครั้ง ไม่งั้น Web app จะยังรันโค้ด
//    เวอร์ชันเก่าอยู่ (ลิงก์ไม่เปลี่ยน แต่โค้ดไม่อัปเดต)
//
// รูปที่อัปโหลดจะถูกเก็บในโฟลเดอร์ Google Drive ชื่อ "Selfie Vocab Run Photos"
// (สร้างให้อัตโนมัติถ้ายังไม่มี) และตั้งค่าแชร์เป็น "ทุกคนที่มีลิงก์ - ดูได้" ให้เอง
// ทุกไฟล์ เพื่อให้ screen.html แสดงรูปได้โดยไม่ต้อง login ด้วยบัญชีเจ้าของไฟล์
//
// หมายเหตุ: อัปโหลดรูปใหม่ทับข้อเดิม (ข้อเดียวกัน กลุ่มเดียวกัน) จะ "แทนที่" ลิงก์
// รูปเก่าในชีตด้วยรูปใหม่ทันที ไม่สะสมซ้ำหลายรูปต่อข้อ — ถ่ายใหม่ได้เรื่อย ๆ ถ้าไม่พอใจ

const SHEET_NAME = "Form_Responses"; // ชื่อแท็บ — แก้ให้ตรงกับชีตจริงถ้าจำเป็น
const DRIVE_FOLDER_NAME = "Selfie Vocab Run Photos";
const TOTAL_ITEMS = 32;

function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);
  } catch (err) {
    return jsonOutput({ status: "error", message: "ระบบกำลังประมวลผลคำขออื่นอยู่ ลองใหม่อีกครั้งในอีกสักครู่" });
  }
  try {
    const data = JSON.parse(e.postData.contents);
    const group = (data.group || "").toString().trim();
    const itemNo = parseInt(data.itemNo, 10);
    const imageBase64 = data.imageBase64;
    const filename = (data.filename || ("SVR" + itemNo + ".jpg")).toString();
    const mimeType = (data.mimeType || "image/jpeg").toString();

    if (!group) return jsonOutput({ status: "error", message: "กรุณาระบุกลุ่ม" });
    if (!itemNo || itemNo < 1 || itemNo > TOTAL_ITEMS) {
      return jsonOutput({ status: "error", message: "หมายเลขข้อไม่ถูกต้อง" });
    }
    if (!imageBase64) return jsonOutput({ status: "error", message: "ไม่พบข้อมูลรูปภาพ" });

    // อัปโหลดรูปเข้า Drive + ตั้งค่าแชร์ให้ดูได้ทุกคนที่มีลิงก์
    const bytes = Utilities.base64Decode(imageBase64);
    const blob = Utilities.newBlob(bytes, mimeType, filename);
    const folder = getOrCreateFolder();
    const file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    const fileUrl = "https://drive.google.com/file/d/" + file.getId() + "/view";

    // เขียนลิงก์ลงชีต (แทนที่ของเดิมถ้ากลุ่ม+ข้อนี้เคยส่งมาก่อน)
    const sheet = getOrCreateSheet();
    const rowIndex = findOrCreateGroupRow(sheet, group);
    const colIndex = findExactColIndex(sheet, "SVR#" + itemNo);
    sheet.getRange(rowIndex, colIndex + 1).setValue(fileUrl);

    return jsonOutput({ status: "ok", fileUrl: fileUrl });
  } catch (err) {
    return jsonOutput({ status: "error", message: err.message });
  } finally {
    lock.releaseLock();
  }
}

function doGet(e) {
  return jsonOutput({ status: "ok", message: "Photo upload webhook is running." });
}

function getOrCreateFolder() {
  const folders = DriveApp.getFoldersByName(DRIVE_FOLDER_NAME);
  if (folders.hasNext()) return folders.next();
  return DriveApp.createFolder(DRIVE_FOLDER_NAME);
}

function buildHeaders() {
  const headers = ["ชื่อกลุ่ม"];
  for (let i = 1; i <= TOTAL_ITEMS; i++) headers.push("SVR#" + i);
  return headers;
}

// สร้างแท็บ + หัวตารางให้อัตโนมัติถ้ายังไม่มี (เหมือนกับระบบคอมเม้น)
function getOrCreateSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(SHEET_NAME);
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(buildHeaders());
  }
  return sheet;
}

// หาคอลัมน์ตรงชื่อเป๊ะ ๆ (กัน SVR#1 ไปจับ SVR#10-19 ผิด) — ถ้าไม่พบ (เช่นชีตเก่า
// ที่หัวตารางไม่ครบ 32 ข้อ) จะเพิ่มคอลัมน์ใหม่ต่อท้ายให้อัตโนมัติ
function findExactColIndex(sheet, name) {
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const idx = headers.findIndex(h => h.toString().trim() === name);
  if (idx >= 0) return idx;
  sheet.getRange(1, headers.length + 1).setValue(name);
  return headers.length;
}

// หาแถวของกลุ่มนี้ (ตรงเป๊ะกับคอลัมน์ A) ถ้ายังไม่มีแถวของกลุ่มนี้ ให้สร้างแถวใหม่
function findOrCreateGroupRow(sheet, group) {
  const lastRow = sheet.getLastRow();
  if (lastRow >= 2) {
    const groupCol = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
    for (let i = 0; i < groupCol.length; i++) {
      if ((groupCol[i][0] || "").toString().trim() === group) {
        return i + 2; // แถวจริงใน sheet (1-indexed, +1 เพราะแถว 1 คือหัวตาราง)
      }
    }
  }
  const newRow = lastRow + 1;
  sheet.getRange(newRow, 1).setValue(group);
  return newRow;
}

function jsonOutput(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
