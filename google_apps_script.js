/**
 * ============================================================================
 * ระบบบริหารจัดการอาชีวศึกษาระบบทวิภาคี (ปวส. 2567) - Google Apps Script (Code.gs)
 * เชื่อมโยงฐานข้อมูล Google Sheets และจัดเก็บไฟล์แนบใน Google Drive แบบ Real-time
 * ============================================================================
 * 
 * วิธีการติดตั้งและใช้งาน (Deployment Guide):
 * 1. เปิด Google Drive ของท่าน สร้างโฟลเดอร์ หรือเปิด Google Spreadsheet ใหม่
 * 2. ไปที่เมนู "ส่วนขยาย" (Extensions) > "Apps Script"
 * 3. ลบโค้ดเริ่มต้นทั้งหมดใน Code.gs แล้วคัดลอกโค้ดนี้ไปวางแทนที่
 * 4. คลิกปุ่ม "บันทึก" (Save / รูปแผ่นดิสก์)
 * 5. คลิกปุ่ม "ทำให้ใช้งานได้" (Deploy) > "การทำให้ใช้งานได้ใหม่" (New deployment)
 * 6. เลือกประเภท: "เว็บแอป" (Web app)
 *    - คำอธิบาย: DVE 2567 Realtime API v1
 *    - ปฏิบัติการในฐานะ (Execute as): ตัวฉัน (Me)
 *    - ผู้มีสิทธิ์เข้าถึง (Who has access): ทุกคน (Anyone)  <--- สำคัญมาก! เพื่อให้เว็บเชื่อมต่อได้โดยไม่ต้องติดบล็อก CORS
 * 7. คลิก "ทำให้ใช้งานได้" (Deploy) แล้วกดให้สิทธิ์ความปลอดภัย (Authorize Access)
 * 8. คัดลอก "URL ของเว็บแอป" (Web App URL) ที่ลงท้ายด้วย /exec 
 * 9. นำ URL มาวางในระบบ DVE ที่เมนู "ตั้งค่าเชื่อมต่อ Google Sheets & Drive"
 */

const FOLDER_NAME = "DVE_ทวิภาคี_เอกสารแนบและทุน";

// รายชื่อตารางในระบบ DVE 2567
const SHEET_NAMES = [
  "Students",
  "Enterprises",
  "TrainingPlans",
  "Supervisions",
  "Attendance",
  "DailyReports",
  "Portfolios",
  "Scholarships",
  "RemedialSessions",
  "Users"
];

// ----------------------------------------------------------------------------
// GET Handler - รองรับการทดสอบการเชื่อมต่อ และดึงข้อมูลทั้งหมด
// ----------------------------------------------------------------------------
function doGet(e) {
  try {
    const action = (e && e.parameter && e.parameter.action) ? e.parameter.action : "ping";
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    if (action === "ping") {
      return responseJSON({
        status: "success",
        message: "Google Apps Script DVE 2567 API พร้อมทำงานแบบ Real-time",
        spreadsheetName: ss.getName(),
        spreadsheetUrl: ss.getUrl(),
        timestamp: new Date().toISOString()
      });
    }

    if (action === "getAll") {
      initSheetsIfNeeded(ss);
      const resultData = {};
      SHEET_NAMES.forEach(sheetName => {
        resultData[sheetName] = getSheetData(ss, sheetName);
      });
      return responseJSON({
        status: "success",
        data: resultData,
        timestamp: new Date().toISOString()
      });
    }

    if (action === "getTable") {
      const targetTable = e.parameter.table;
      if (!targetTable) {
        return responseJSON({ status: "error", message: "กรุณาระบุชื่อ table" });
      }
      initSheetsIfNeeded(ss);
      const data = getSheetData(ss, targetTable);
      return responseJSON({ status: "success", table: targetTable, data: data });
    }

    return responseJSON({ status: "error", message: "Action ไม่ถูกต้อง" });
  } catch (err) {
    return responseJSON({ status: "error", message: err.toString() });
  }
}

// ----------------------------------------------------------------------------
// POST Handler - บันทึกข้อมูล Real-time และอัปโหลดไฟล์ไปยัง Google Drive
// ----------------------------------------------------------------------------
function doPost(e) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    initSheetsIfNeeded(ss);

    let body = {};
    if (e.postData && e.postData.contents) {
      body = JSON.parse(e.postData.contents);
    }

    const action = body.action || "syncAll";

    // 1. อัปโหลดไฟล์เอกสารแนบ/หลักฐานทุน/รูปภาพ ไปยัง Google Drive
    if (action === "uploadFile") {
      const folder = getOrCreateFolder(FOLDER_NAME);
      const fileName = body.fileName || ("DVE_Doc_" + Date.now());
      const mimeType = body.mimeType || "application/pdf";
      const base64Data = body.base64Data;

      if (!base64Data) {
        return responseJSON({ status: "error", message: "ไม่พบข้อมูลไฟล์ base64" });
      }

      // ทำความสะอาด Base64 string หากมี Header data:...;base64,
      const cleanBase64 = base64Data.indexOf(",") > -1 ? base64Data.split(",")[1] : base64Data;
      const decoded = Utilities.base64Decode(cleanBase64);
      const blob = Utilities.newBlob(decoded, mimeType, fileName);
      const file = folder.createFile(blob);

      // กำหนดสิทธิ์ให้ทุกคนที่มีลิงก์สามารถดูเอกสารได้
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

      return responseJSON({
        status: "success",
        fileId: file.getId(),
        fileName: file.getName(),
        driveUrl: file.getUrl(),
        downloadUrl: "https://drive.google.com/uc?export=view&id=" + file.getId()
      });
    }

    // 2. บันทึกข้อมูลเดี่ยวหรืออัปเดตเฉพาะตาราง (Real-time Table Save)
    if (action === "saveTable") {
      const tableName = body.tableName;
      const records = body.records;
      if (!tableName || !Array.isArray(records)) {
        return responseJSON({ status: "error", message: "tableName หรือ records ไม่ถูกต้อง" });
      }

      saveTableData(ss, tableName, records);
      return responseJSON({
        status: "success",
        message: "บันทึกตาราง " + tableName + " เรียบร้อยแล้ว",
        recordCount: records.length,
        timestamp: new Date().toISOString()
      });
    }

    // 3. บันทึกซิงค์ข้อมูลทั้งระบบ (Full Sync)
    if (action === "syncAll") {
      const allData = body.data || {};
      for (const key in allData) {
        if (Array.isArray(allData[key])) {
          saveTableData(ss, key, allData[key]);
        }
      }
      return responseJSON({
        status: "success",
        message: "ซิงค์ฐานข้อมูลทวิภาคีทั้งหมดเรียบร้อยแล้ว",
        timestamp: new Date().toISOString()
      });
    }

    return responseJSON({ status: "error", message: "Action POST ไม่ถูกต้อง" });
  } catch (err) {
    return responseJSON({ status: "error", message: err.toString() });
  }
}

// ----------------------------------------------------------------------------
// Helper Functions: จัดการ Sheet และ Drive
// ----------------------------------------------------------------------------

function getOrCreateFolder(folderName) {
  const folders = DriveApp.getFoldersByName(folderName);
  if (folders.hasNext()) {
    return folders.next();
  }
  return DriveApp.createFolder(folderName);
}

function initSheetsIfNeeded(ss) {
  SHEET_NAMES.forEach(name => {
    let sheet = ss.getSheetByName(name);
    if (!sheet) {
      sheet = ss.insertSheet(name);
      // ใส่ Header ตัวอย่าง
      sheet.appendRow(["_id", "data_json", "updated_at"]);
      sheet.getRange(1, 1, 1, 3).setFontWeight("bold").setBackground("#e2e8f0");
    }
  });
}

function getSheetData(ss, sheetName) {
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) return [];
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return [];

  // ตรวจสอบโครงสร้าง: หากคอลัมน์ 2 เป็น data_json ให้ deserialize
  const header = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const isJsonStore = (header[1] === "data_json");

  if (isJsonStore) {
    const rawRows = sheet.getRange(2, 1, lastRow - 1, 3).getValues();
    return rawRows.map(r => {
      try {
        return JSON.parse(r[1]);
      } catch (e) {
        return { id: r[0], raw: r[1] };
      }
    });
  } else {
    // กรณีเก็บแบบ flat columns ปกติ
    const values = sheet.getRange(2, 1, lastRow - 1, header.length).getValues();
    return values.map(row => {
      const obj = {};
      header.forEach((key, idx) => {
        obj[key] = row[idx];
      });
      return obj;
    });
  }
}

function saveTableData(ss, sheetName, records) {
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
  }

  sheet.clear();
  sheet.appendRow(["_id", "data_json", "updated_at"]);
  sheet.getRange(1, 1, 1, 3).setFontWeight("bold").setBackground("#3b82f6").setFontColor("#ffffff");

  if (records.length === 0) return;

  const now = new Date().toISOString();
  const rows = records.map(item => {
    const id = item.id || item._id || Utilities.getUuid();
    return [id, JSON.stringify(item), now];
  });

  sheet.getRange(2, 1, rows.length, 3).setValues(rows);
}

function responseJSON(data) {
  return ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}
