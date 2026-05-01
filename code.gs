// ══════════════════════════════════════════════════════════════
//  HRM System – Google Apps Script Backend
//  วิธีใช้งาน:
//  1. เปิด Google Sheets ใหม่
//  2. Extensions → Apps Script → วางโค้ดนี้ → Save
//  3. Deploy → New deployment → Web app
//  4. Execute as: Me | Who has access: Anyone
//  5. คัดลอก Web app URL ไปใส่ในหน้า hrm.html (⚙️ Google Sheets)
// ══════════════════════════════════════════════════════════════

// ── Sheet: พนักงาน ─────────────────────────────────────────────
const SHEET_EMP  = 'พนักงาน';
const EMP_HEADERS = [
  'id', 'empId', 'fullName', 'company', 'department', 'position',
  'startDate', 'address', 'email', 'supervisor',
  'sickLeaveNoDoc', 'sickLeaveWithDoc', 'personalLeave', 'annualLeave',
  'status', 'resignDate', 'createdAt', 'updatedAt'
];
const EMP_LABELS = [
  'ID (ระบบ)', 'รหัสพนักงาน', 'ชื่อ-นามสกุล', 'บริษัท', 'แผนก', 'ตำแหน่ง',
  'วันที่เริ่มงาน', 'ที่อยู่', 'อีเมล', 'ผู้บังคับบัญชา',
  'ลาป่วย(ไม่มีใบ)', 'ลาป่วย(มีใบ)', 'ลากิจ', 'ลาพักร้อน',
  'สถานะ', 'วันที่ลาออก', 'วันที่สร้าง', 'วันที่แก้ไข'
];

// ── Sheet: แผนก ────────────────────────────────────────────────
const SHEET_DEPT  = 'แผนก';
const DEPT_HEADERS = ['id', 'company', 'name', 'manager', 'assistantManager', 'createdAt', 'updatedAt'];
const DEPT_LABELS  = ['ID (ระบบ)', 'บริษัท', 'ชื่อแผนก', 'ผู้จัดการแผนก', 'ผู้ช่วยผู้จัดการ', 'วันที่สร้าง', 'วันที่แก้ไข'];

// ── Sheet helpers ──────────────────────────────────────────────
function initSheet(name, labels, widths) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    const hRange = sheet.getRange(1, 1, 1, labels.length);
    hRange.setValues([labels]);
    hRange.setFontWeight('bold');
    hRange.setBackground('#0284c7');
    hRange.setFontColor('#ffffff');
    hRange.setHorizontalAlignment('center');
    sheet.setFrozenRows(1);
    if (widths) widths.forEach((w, i) => sheet.setColumnWidth(i + 1, w));
  }
  return sheet;
}

function getEmpSheet()  { return initSheet(SHEET_EMP,  EMP_LABELS,  [140,110,160,140,120,130,110,200,160,150,100,100,80,90,120,110,160,160]); }
function getDeptSheet() { return initSheet(SHEET_DEPT, DEPT_LABELS, [140,200,160,160,160,160,160]); }

function rowToObj(headers, row) {
  const obj = {};
  headers.forEach((h, i) => { obj[h] = row[i] !== undefined ? String(row[i]) : ''; });
  return obj;
}
function objToRow(headers, obj) {
  return headers.map(h => obj[h] !== undefined ? obj[h] : '');
}

function readSheet(sheet, headers) {
  const last = sheet.getLastRow();
  if (last <= 1) return [];
  return sheet.getRange(2, 1, last - 1, headers.length).getValues()
    .filter(r => r[0] !== '')
    .map(r => rowToObj(headers, r));
}

function findRowById(sheet, id) {
  const last = sheet.getLastRow();
  if (last <= 1) return -1;
  const ids = sheet.getRange(2, 1, last - 1, 1).getValues();
  for (let i = 0; i < ids.length; i++) {
    if (String(ids[i][0]) === String(id)) return i + 2;
  }
  return -1;
}

// ── doGet ──────────────────────────────────────────────────────
function doGet(e) {
  const action = (e.parameter && e.parameter.action) || 'list';
  try {
    if (action === 'list') {
      return jsonResp({ success: true, data: readSheet(getEmpSheet(), EMP_HEADERS) });
    }
    if (action === 'listDepts') {
      return jsonResp({ success: true, data: readSheet(getDeptSheet(), DEPT_HEADERS) });
    }
    return jsonResp({ success: false, error: 'Unknown action' });
  } catch (err) {
    return jsonResp({ success: false, error: err.toString() });
  }
}

// ── doPost ─────────────────────────────────────────────────────
function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);
    const { action, data, id } = body;

    // ── Employee ──
    if (action === 'create') {
      const sheet = getEmpSheet();
      sheet.appendRow(objToRow(EMP_HEADERS, data));
      colorStatus(sheet, sheet.getLastRow(), EMP_HEADERS.indexOf('status') + 1, data.status);
      return jsonResp({ success: true });
    }
    if (action === 'update') {
      const sheet = getEmpSheet();
      const row = findRowById(sheet, data.id);
      if (row < 0) return jsonResp({ success: false, error: 'Employee not found' });
      sheet.getRange(row, 1, 1, EMP_HEADERS.length).setValues([objToRow(EMP_HEADERS, data)]);
      colorStatus(sheet, row, EMP_HEADERS.indexOf('status') + 1, data.status);
      return jsonResp({ success: true });
    }
    if (action === 'delete') {
      const sheet = getEmpSheet();
      const row = findRowById(sheet, id);
      if (row < 0) return jsonResp({ success: false, error: 'Employee not found' });
      sheet.deleteRow(row);
      return jsonResp({ success: true });
    }

    // ── Department ──
    if (action === 'createDept') {
      const sheet = getDeptSheet();
      sheet.appendRow(objToRow(DEPT_HEADERS, data));
      return jsonResp({ success: true });
    }
    if (action === 'updateDept') {
      const sheet = getDeptSheet();
      const row = findRowById(sheet, data.id);
      if (row < 0) return jsonResp({ success: false, error: 'Department not found' });
      sheet.getRange(row, 1, 1, DEPT_HEADERS.length).setValues([objToRow(DEPT_HEADERS, data)]);
      return jsonResp({ success: true });
    }
    if (action === 'deleteDept') {
      const sheet = getDeptSheet();
      const row = findRowById(sheet, id);
      if (row < 0) return jsonResp({ success: false, error: 'Department not found' });
      sheet.deleteRow(row);
      return jsonResp({ success: true });
    }

    return jsonResp({ success: false, error: 'Unknown action' });
  } catch (err) {
    return jsonResp({ success: false, error: err.toString() });
  }
}

// ── Helpers ────────────────────────────────────────────────────
function colorStatus(sheet, row, col, status) {
  const cell = sheet.getRange(row, col);
  if (status === 'active') {
    cell.setBackground('#d1fae5'); cell.setFontColor('#065f46');
  } else {
    cell.setBackground('#ffe4e6'); cell.setFontColor('#9f1239');
  }
}

function jsonResp(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
