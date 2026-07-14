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
  'id', 'empId', 'namePrefix', 'fullName', 'nickname', 'company', 'department', 'position',
  'startDate', 'birthDate', 'gender', 'bloodType', 'address', 'phone', 'education', 'email', 'supervisor',
  'reference', 'emergencyContact',
  'fullNameEn', 'nicknameEn', 'maritalStatus', 'nationality', 'ethnicity', 'religion', 'idCardNo',
  'height', 'weight', 'taxId', 'photoUrl', 'militaryStatus', 'disability', 'chronicDisease', 'hospital',
  'sickLeaveNoDoc', 'sickLeaveWithDoc', 'personalLeave', 'annualLeave',
  'salaryStartDate', 'baseSalary', 'positionAllowance', 'otherBenefits', 'housingAllowance',
  'status', 'resignDate', 'createdAt', 'updatedAt'
];
const EMP_LABELS = [
  'ID (ระบบ)', 'รหัสพนักงาน', 'คำนำหน้าชื่อ', 'ชื่อ-นามสกุล', 'ชื่อเล่น', 'บริษัท', 'แผนก', 'ตำแหน่ง',
  'วันที่เริ่มงาน', 'วันเกิด', 'เพศ', 'กรุ๊ปเลือด', 'ที่อยู่', 'เบอร์ติดต่อ', 'วุฒิการศึกษา', 'อีเมล', 'ผู้บังคับบัญชา',
  'ผู้อ้างอิง', 'เบอร์ติดต่อฉุกเฉิน',
  'ชื่อ-นามสกุล (อังกฤษ)', 'ชื่อเล่น (อังกฤษ)', 'สถานะการสมรส', 'สัญชาติ', 'เชื้อชาติ', 'ศาสนา', 'เลขบัตรประชาชน/พาสปอร์ต/ใบขับขี่',
  'ส่วนสูง (ซม.)', 'น้ำหนัก (กก.)', 'เลขผู้เสียภาษี (กรณีไม่ใช่สัญชาติไทย)', 'รูปพนักงาน (URL)', 'สถานะทางทหาร', 'ผู้พิการ', 'โรคประจำตัว/แพ้ยา-อาหาร', 'โรงพยาบาลประกันสังคม',
  'ลาป่วย(ไม่มีใบ)', 'ลาป่วย(มีใบ)', 'ลากิจ', 'ลาพักร้อน',
  'วันที่เริ่ม(เงินเดือน)', 'เงินเดือนฐาน', 'ค่าตำแหน่ง', 'สวัสดิการอื่น', 'ค่าที่พัก',
  'สถานะ', 'วันที่ลาออก', 'วันที่สร้าง', 'วันที่แก้ไข'
];

// ── Sheet: ประวัติการเปลี่ยนตำแหน่ง ───────────────────────────
const SHEET_POSITION_HISTORY  = 'ประวัติการเปลี่ยนตำแหน่ง';
const POSITION_HISTORY_HEADERS = ['id', 'employeeId', 'month', 'year', 'position', 'createdAt', 'updatedAt'];
const POSITION_HISTORY_LABELS  = ['ID', 'รหัสพนักงาน', 'เดือน', 'ปี', 'ตำแหน่ง', 'วันที่สร้าง', 'วันที่แก้ไข'];

// ── Sheet: ประวัติการเปลี่ยนเงินเดือน ─────────────────────────
const SHEET_SALARY_HISTORY  = 'ประวัติการเปลี่ยนเงินเดือน';
const SALARY_HISTORY_HEADERS = ['id', 'employeeId', 'month', 'year', 'salary', 'createdAt', 'updatedAt'];
const SALARY_HISTORY_LABELS  = ['ID', 'รหัสพนักงาน', 'เดือน', 'ปี', 'เงินเดือน', 'วันที่สร้าง', 'วันที่แก้ไข'];

// ── Sheet: แผนก ────────────────────────────────────────────────
const SHEET_DEPT  = 'แผนก';
const DEPT_HEADERS = ['id', 'company', 'name', 'manager', 'assistantManager', 'createdAt', 'updatedAt'];
const DEPT_LABELS  = ['ID (ระบบ)', 'บริษัท', 'ชื่อแผนก', 'ผู้จัดการแผนก', 'ผู้ช่วยผู้จัดการ', 'วันที่สร้าง', 'วันที่แก้ไข'];

// ── Sheet: ผู้ใช้งาน ───────────────────────────────────────────
const SHEET_USER  = 'ผู้ใช้งาน';
const USER_HEADERS = ['id', 'username', 'password', 'role', 'name', 'createdAt', 'updatedAt'];
const USER_LABELS  = ['ID', 'Username', 'Password', 'Role', 'ชื่อ-นามสกุล', 'วันที่สร้าง', 'วันที่แก้ไข'];

// ── Sheet: ใบเตือน ────────────────────────────────────────────
const SHEET_WARNING  = 'ใบเตือน';
const WARNING_HEADERS = ['id', 'warningNo', 'date', 'type', 'topic', 'employeeId', 'investigator', 'details', 'punishment', 'createdAt', 'updatedAt'];
const WARNING_LABELS  = ['ID', 'เลขที่ใบเตือน', 'วันที่ใบเตือน', 'ระดับการตักเตือน', 'หัวข้อใบเตือน', 'รหัสพนักงาน', 'ผู้สอบสวนเหตุการณ์', 'รายละเอียดเหตุการณ์', 'บทลงโทษ', 'วันที่สร้าง', 'วันที่แก้ไข'];

// ── Sheet: ประวัติอบรม ─────────────────────────────────────────
const SHEET_TRAINING  = 'ประวัติอบรม';
const TRAINING_HEADERS = ['id', 'date', 'topic', 'duration', 'employeeIds', 'cost', 'createdAt', 'updatedAt'];
const TRAINING_LABELS  = ['ID', 'วันที่อบรม', 'หัวข้อการอบรม', 'ระยะเวลา', 'รหัสพนักงาน (IDs)', 'ค่าใช้จ่าย', 'วันที่สร้าง', 'วันที่แก้ไข'];

// ── Sheet: เงินเดือน ───────────────────────────────────────────
const SHEET_PAYROLL  = 'เงินเดือน';
const PAYROLL_HEADERS = [
  'id', 'year', 'month', 'empId', 'empName', 'department',
  'baseSalary', 'ot1', 'ot15', 'ot3', 'otAmount',
  'diligence', 'perDiem', 'posAllowance', 'welfare', 'housing', 'backPay',
  'totalIncome', 'ssn', 'tax', 'gsl', 'excessLeave', 'companyDebt',
  'totalDeduct', 'netPay', 'updatedAt'
];
const PAYROLL_LABELS = [
  'ID', 'ปี', 'เดือน', 'รหัสพนักงาน', 'ชื่อพนักงาน', 'แผนก',
  'ฐานเงินเดือน', 'OT 1x (ชม)', 'OT 1.5x (ชม)', 'OT 3x (ชม)', 'รวม OT',
  'เบี้ยขยัน', 'เบี้ยเลี้ยง', 'ค่าตำแหน่ง', 'ค่าสวัสดิการ', 'ค่าที่พัก', 'ตกเบิก',
  'รายรับรวม', 'ประกันสังคม', 'ภาษี', 'กยศ', 'หักลาเกิน', 'ใช้หนี้บริษัท',
  'รายจ่ายรวม', 'คงเหลือ', 'วันที่บันทึก'
];

// ── Sheet: การลา ───────────────────────────────────────────────
const SHEET_LEAVE  = 'การลา';
const LEAVE_HEADERS = [
  'id', 'employeeId', 'leaveType', 'startDate', 'endDate', 'days', 'reason', 'attachmentUrl',
  'status', 'approver', 'approvedAt', 'rejectReason', 'createdAt', 'updatedAt'
];
const LEAVE_LABELS = [
  'ID', 'รหัสพนักงาน (FK)', 'ประเภทการลา', 'วันที่เริ่มลา', 'วันที่สิ้นสุด', 'จำนวนวัน', 'เหตุผล', 'เอกสารแนบ (URL)',
  'สถานะ', 'ผู้อนุมัติ', 'วันที่อนุมัติ/ปฏิเสธ', 'เหตุผลที่ไม่อนุมัติ', 'วันที่ยื่นคำขอ', 'วันที่แก้ไข'
];

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

// ตรวจสอบและแทรกคอลัมน์ที่หายไปโดยไม่ทำลายข้อมูลเดิม
function ensureColumns(sheet, labels) {
  const lastCol = sheet.getLastColumn();
  if (lastCol < 1) return;
  const cur = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(String);
  for (let i = 0; i < labels.length; i++) {
    if (cur[i] === labels[i]) continue;
    if (cur.indexOf(labels[i]) !== -1) continue; // มีอยู่แล้วแต่ตำแหน่งต่างกัน — ข้าม
    // แทรกคอลัมน์ใหม่ที่ตำแหน่ง i+1
    const pos = i + 1;
    if (pos <= sheet.getLastColumn()) {
      sheet.insertColumnBefore(pos);
    } else {
      sheet.insertColumnAfter(sheet.getLastColumn());
    }
    const hdr = sheet.getRange(1, pos);
    hdr.setValue(labels[i]);
    hdr.setFontWeight('bold');
    hdr.setBackground('#0284c7');
    hdr.setFontColor('#ffffff');
    hdr.setHorizontalAlignment('center');
    cur.splice(i, 0, labels[i]);
  }
}

function getEmpSheet()  {
  const s = initSheet(SHEET_EMP,  EMP_LABELS,  [140,110,90,160,90,140,120,130,110,100,80,90,200,130,150,160,150,140,140,160,110,110,100,100,100,150,80,80,140,200,140,110,220,180,100,100,80,90,110,110,110,120,110,160,160]);
  ensureColumns(s, EMP_LABELS);
  return s;
}
function getPositionHistorySheet() {
  const s = initSheet(SHEET_POSITION_HISTORY, POSITION_HISTORY_LABELS, [140,140,90,70,180,160,160]);
  ensureColumns(s, POSITION_HISTORY_LABELS);
  return s;
}
function getSalaryHistorySheet() {
  const s = initSheet(SHEET_SALARY_HISTORY, SALARY_HISTORY_LABELS, [140,140,90,70,120,160,160]);
  ensureColumns(s, SALARY_HISTORY_LABELS);
  return s;
}
function getDeptSheet() {
  const s = initSheet(SHEET_DEPT, DEPT_LABELS, [140,200,160,160,160,160,160]);
  ensureColumns(s, DEPT_LABELS);
  return s;
}
function getUserSheet() {
  const s = initSheet(SHEET_USER, USER_LABELS, [140,120,120,80,160,160,160]);
  ensureColumns(s, USER_LABELS);
  if (s.getLastRow() <= 1) {
    const now = new Date().toISOString();
    s.appendRow(objToRow(USER_HEADERS, {
      id:'U001', username:'admin', password:'admin1234',
      role:'admin', name:'Administrator', createdAt:now, updatedAt:now
    }));
  }
  return s;
}
function getWarningSheet() {
  const s = initSheet(SHEET_WARNING, WARNING_LABELS, [140,120,110,180,200,140,150,400,200,160,160]);
  ensureColumns(s, WARNING_LABELS);
  return s;
}
function getTrainingSheet() {
  const s = initSheet(SHEET_TRAINING, TRAINING_LABELS, [140,110,220,120,400,110,160,160]);
  ensureColumns(s, TRAINING_LABELS);
  return s;
}
function getPayrollSheet() {
  const s = initSheet(SHEET_PAYROLL, PAYROLL_LABELS,
    [180,60,60,120,160,120, 110,80,80,80,90, 90,90,90,90,90,90, 100,90,90,80,90,100, 100,100,160]);
  ensureColumns(s, PAYROLL_LABELS);
  return s;
}
function getLeaveSheet() {
  const s = initSheet(SHEET_LEAVE, LEAVE_LABELS, [140,140,150,110,110,80,300,220,100,140,140,300,160,160]);
  ensureColumns(s, LEAVE_LABELS);
  return s;
}

// ── Leave attachment (Google Drive) ─────────────────────────────
function getLeaveDocsFolder() {
  const folders = DriveApp.getFoldersByName('HRM_LeaveDocs');
  if (folders.hasNext()) return folders.next();
  return DriveApp.createFolder('HRM_LeaveDocs');
}
function saveLeaveAttachment(base64Data, fileName, mimeType) {
  if (!base64Data) return '';
  const bytes = Utilities.base64Decode(base64Data);
  const blob = Utilities.newBlob(bytes, mimeType || 'image/jpeg', fileName || ('leave_doc_' + Date.now()));
  const file = getLeaveDocsFolder().createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return file.getUrl();
}

// ── Employee photo (Google Drive) ───────────────────────────────
function getEmployeePhotosFolder() {
  const folders = DriveApp.getFoldersByName('HRM_EmployeePhotos');
  if (folders.hasNext()) return folders.next();
  return DriveApp.createFolder('HRM_EmployeePhotos');
}
function saveEmployeePhoto(base64Data, fileName, mimeType) {
  if (!base64Data) return '';
  const bytes = Utilities.base64Decode(base64Data);
  const blob = Utilities.newBlob(bytes, mimeType || 'image/jpeg', fileName || ('emp_photo_' + Date.now()));
  const file = getEmployeePhotosFolder().createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return file.getUrl();
}

function rowToObj(headers, row) {
  const obj = {};
  headers.forEach((h, i) => {
    const v = row[i];
    if (v instanceof Date) {
      if (isNaN(v.getTime())) {
        obj[h] = '';
      } else {
        const y = v.getFullYear();
        const mo = String(v.getMonth() + 1).padStart(2, '0');
        const d  = String(v.getDate()).padStart(2, '0');
        obj[h] = y + '-' + mo + '-' + d;
      }
    } else {
      obj[h] = (v !== undefined && v !== null) ? String(v) : '';
    }
  });
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
    if (action === 'listUsers') {
      return jsonResp({ success: true, data: readSheet(getUserSheet(), USER_HEADERS) });
    }
    if (action === 'listTrainings') {
      return jsonResp({ success: true, data: readSheet(getTrainingSheet(), TRAINING_HEADERS) });
    }
    if (action === 'listWarnings') {
      return jsonResp({ success: true, data: readSheet(getWarningSheet(), WARNING_HEADERS) });
    }
    if (action === 'listPayrolls') {
      const year  = e.parameter.year  ? String(e.parameter.year)  : '';
      const month = e.parameter.month ? String(e.parameter.month) : '';
      let rows = readSheet(getPayrollSheet(), PAYROLL_HEADERS);
      if (year)  rows = rows.filter(r => String(r.year)  === year);
      if (month) rows = rows.filter(r => String(r.month) === month);
      return jsonResp({ success: true, data: rows });
    }
    if (action === 'listLeaves') {
      return jsonResp({ success: true, data: readSheet(getLeaveSheet(), LEAVE_HEADERS) });
    }
    if (action === 'listPositionHistory') {
      return jsonResp({ success: true, data: readSheet(getPositionHistorySheet(), POSITION_HISTORY_HEADERS) });
    }
    if (action === 'listSalaryHistory') {
      return jsonResp({ success: true, data: readSheet(getSalaryHistorySheet(), SALARY_HISTORY_HEADERS) });
    }
    if (action === 'login') {
      const username = e.parameter.username || '';
      const password = e.parameter.password || '';
      const users = readSheet(getUserSheet(), USER_HEADERS);
      const user = users.find(u => u.username === username && u.password === password);
      if (user) {
        return jsonResp({ success: true, user: { id: user.id, username: user.username, name: user.name, role: user.role } });
      }
      return jsonResp({ success: false, error: 'ชื่อผู้ใช้หรือรหัสผ่านไม่ถูกต้อง' });
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
      if (data.photoData) {
        data.photoUrl = saveEmployeePhoto(data.photoData, data.photoName, data.photoType);
      }
      const sheet = getEmpSheet();
      sheet.appendRow(objToRow(EMP_HEADERS, data));
      const newRow = sheet.getLastRow();
      fixTextFields(sheet, newRow, data, ['phone', 'emergencyContact', 'idCardNo', 'taxId']);
      colorStatus(sheet, newRow, EMP_HEADERS.indexOf('status') + 1, data.status);
      return jsonResp({ success: true });
    }
    if (action === 'update') {
      const sheet = getEmpSheet();
      const row = findRowById(sheet, data.id);
      if (row < 0) return jsonResp({ success: false, error: 'Employee not found' });
      if (data.photoData) {
        data.photoUrl = saveEmployeePhoto(data.photoData, data.photoName, data.photoType);
      } else if (!data.photoUrl) {
        const existing = rowToObj(EMP_HEADERS, sheet.getRange(row, 1, 1, EMP_HEADERS.length).getValues()[0]);
        data.photoUrl = existing.photoUrl || '';
      }
      sheet.getRange(row, 1, 1, EMP_HEADERS.length).setValues([objToRow(EMP_HEADERS, data)]);
      fixTextFields(sheet, row, data, ['phone', 'emergencyContact', 'idCardNo', 'taxId']);
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

    // ── Users ──
    if (action === 'login') {
      const users = readSheet(getUserSheet(), USER_HEADERS);
      const user = users.find(u => u.username === data.username && u.password === data.password);
      if (user) {
        return jsonResp({ success: true, user: { id: user.id, username: user.username, name: user.name, role: user.role } });
      }
      return jsonResp({ success: false, error: 'ชื่อผู้ใช้หรือรหัสผ่านไม่ถูกต้อง' });
    }
    if (action === 'createUser') {
      const sheet = getUserSheet();
      const users = readSheet(sheet, USER_HEADERS);
      if (users.find(u => u.username === data.username)) {
        return jsonResp({ success: false, error: 'Username นี้ถูกใช้งานแล้ว' });
      }
      sheet.appendRow(objToRow(USER_HEADERS, data));
      return jsonResp({ success: true });
    }
    if (action === 'updateUser') {
      const sheet = getUserSheet();
      const row = findRowById(sheet, data.id);
      if (row < 0) return jsonResp({ success: false, error: 'User not found' });
      if (!data.password) {
        const existing = readSheet(sheet, USER_HEADERS).find(u => u.id === data.id);
        if (existing) data.password = existing.password;
      }
      sheet.getRange(row, 1, 1, USER_HEADERS.length).setValues([objToRow(USER_HEADERS, data)]);
      return jsonResp({ success: true });
    }
    if (action === 'deleteUser') {
      const sheet = getUserSheet();
      const row = findRowById(sheet, id);
      if (row < 0) return jsonResp({ success: false, error: 'User not found' });
      sheet.deleteRow(row);
      return jsonResp({ success: true });
    }

    // ── Training ──
    if (action === 'createTraining') {
      const sheet = getTrainingSheet();
      sheet.appendRow(objToRow(TRAINING_HEADERS, data));
      return jsonResp({ success: true });
    }
    if (action === 'updateTraining') {
      const sheet = getTrainingSheet();
      const row = findRowById(sheet, data.id);
      if (row < 0) return jsonResp({ success: false, error: 'Training not found' });
      sheet.getRange(row, 1, 1, TRAINING_HEADERS.length).setValues([objToRow(TRAINING_HEADERS, data)]);
      return jsonResp({ success: true });
    }
    if (action === 'deleteTraining') {
      const sheet = getTrainingSheet();
      const row = findRowById(sheet, id);
      if (row < 0) return jsonResp({ success: false, error: 'Training not found' });
      sheet.deleteRow(row);
      return jsonResp({ success: true });
    }

    // ── Warning ──
    if (action === 'createWarning') {
      const sheet = getWarningSheet();
      sheet.appendRow(objToRow(WARNING_HEADERS, data));
      return jsonResp({ success: true });
    }
    if (action === 'updateWarning') {
      const sheet = getWarningSheet();
      const row = findRowById(sheet, data.id);
      if (row < 0) return jsonResp({ success: false, error: 'Warning not found' });
      sheet.getRange(row, 1, 1, WARNING_HEADERS.length).setValues([objToRow(WARNING_HEADERS, data)]);
      return jsonResp({ success: true });
    }
    if (action === 'deleteWarning') {
      const sheet = getWarningSheet();
      const row = findRowById(sheet, id);
      if (row < 0) return jsonResp({ success: false, error: 'Warning not found' });
      sheet.deleteRow(row);
      return jsonResp({ success: true });
    }

    // ── Position History ──
    if (action === 'createPositionHistory') {
      const sheet = getPositionHistorySheet();
      sheet.appendRow(objToRow(POSITION_HISTORY_HEADERS, data));
      return jsonResp({ success: true });
    }
    if (action === 'updatePositionHistory') {
      const sheet = getPositionHistorySheet();
      const row = findRowById(sheet, data.id);
      if (row < 0) return jsonResp({ success: false, error: 'Position history not found' });
      sheet.getRange(row, 1, 1, POSITION_HISTORY_HEADERS.length).setValues([objToRow(POSITION_HISTORY_HEADERS, data)]);
      return jsonResp({ success: true });
    }
    if (action === 'deletePositionHistory') {
      const sheet = getPositionHistorySheet();
      const row = findRowById(sheet, id);
      if (row < 0) return jsonResp({ success: false, error: 'Position history not found' });
      sheet.deleteRow(row);
      return jsonResp({ success: true });
    }

    // ── Salary History ──
    if (action === 'createSalaryHistory') {
      const sheet = getSalaryHistorySheet();
      sheet.appendRow(objToRow(SALARY_HISTORY_HEADERS, data));
      return jsonResp({ success: true });
    }
    if (action === 'updateSalaryHistory') {
      const sheet = getSalaryHistorySheet();
      const row = findRowById(sheet, data.id);
      if (row < 0) return jsonResp({ success: false, error: 'Salary history not found' });
      sheet.getRange(row, 1, 1, SALARY_HISTORY_HEADERS.length).setValues([objToRow(SALARY_HISTORY_HEADERS, data)]);
      return jsonResp({ success: true });
    }
    if (action === 'deleteSalaryHistory') {
      const sheet = getSalaryHistorySheet();
      const row = findRowById(sheet, id);
      if (row < 0) return jsonResp({ success: false, error: 'Salary history not found' });
      sheet.deleteRow(row);
      return jsonResp({ success: true });
    }

    // ── Payroll ──
    if (action === 'savePayroll') {
      const year    = String(body.year  || '');
      const month   = String(body.month || '');
      const empData = body.data    || {};
      const empMeta = body.empMeta || [];
      const sheet   = getPayrollSheet();
      const now     = new Date().toISOString();
      const metaMap = {};
      empMeta.forEach(function(m){ metaMap[m.id] = m; });
      for (var empId in empData) {
        var pd   = empData[empId];
        var meta = metaMap[empId] || {};
        var base  = parseFloat(meta.baseSalary) || 0;
        var hRate = base / 30 / 8;
        var otAmt = Math.round(
          ((parseFloat(pd.ot1)||0)*hRate) +
          ((parseFloat(pd.ot15)||0)*hRate*1.5) +
          ((parseFloat(pd.ot3)||0)*hRate*3)
        );
        var totalIncome = Math.round(
          base + otAmt +
          (parseFloat(pd.diligence)||0) + (parseFloat(pd.perDiem)||0) +
          (parseFloat(pd.posAllowance)||0) + (parseFloat(pd.welfare)||0) +
          (parseFloat(pd.housing)||0) + (parseFloat(pd.backPay)||0)
        );
        var totalDeduct = Math.round(
          (parseFloat(pd.ssn)||0) + (parseFloat(pd.tax)||0) +
          (parseFloat(pd.gsl)||0) + (parseFloat(pd.excessLeave)||0) +
          (parseFloat(pd.companyDebt)||0)
        );
        var rec = {
          id: 'PL_' + (meta.empId||empId) + '_' + year + '_' + month,
          year: year, month: month,
          empId: meta.empId || empId,
          empName: meta.fullName || '',
          department: meta.department || '',
          baseSalary: base,
          ot1: parseFloat(pd.ot1)||0,
          ot15: parseFloat(pd.ot15)||0,
          ot3: parseFloat(pd.ot3)||0,
          otAmount: otAmt,
          diligence: parseFloat(pd.diligence)||0,
          perDiem: parseFloat(pd.perDiem)||0,
          posAllowance: parseFloat(pd.posAllowance)||0,
          welfare: parseFloat(pd.welfare)||0,
          housing: parseFloat(pd.housing)||0,
          backPay: parseFloat(pd.backPay)||0,
          totalIncome: totalIncome,
          ssn: parseFloat(pd.ssn)||0,
          tax: parseFloat(pd.tax)||0,
          gsl: parseFloat(pd.gsl)||0,
          excessLeave: parseFloat(pd.excessLeave)||0,
          companyDebt: parseFloat(pd.companyDebt)||0,
          totalDeduct: totalDeduct,
          netPay: totalIncome - totalDeduct,
          updatedAt: now
        };
        var rowIdx = findRowById(sheet, rec.id);
        if (rowIdx > 0) {
          sheet.getRange(rowIdx, 1, 1, PAYROLL_HEADERS.length).setValues([objToRow(PAYROLL_HEADERS, rec)]);
        } else {
          sheet.appendRow(objToRow(PAYROLL_HEADERS, rec));
        }
      }
      return jsonResp({ success: true });
    }

    // ── Leave ──
    if (action === 'createLeave') {
      const sheet = getLeaveSheet();
      if (data.attachmentData) {
        data.attachmentUrl = saveLeaveAttachment(data.attachmentData, data.attachmentName, data.attachmentType);
      }
      sheet.appendRow(objToRow(LEAVE_HEADERS, data));
      return jsonResp({ success: true });
    }
    if (action === 'updateLeave') {
      const sheet = getLeaveSheet();
      const row = findRowById(sheet, data.id);
      if (row < 0) return jsonResp({ success: false, error: 'Leave request not found' });
      if (data.attachmentData) {
        data.attachmentUrl = saveLeaveAttachment(data.attachmentData, data.attachmentName, data.attachmentType);
      } else if (!data.attachmentUrl) {
        const existing = rowToObj(LEAVE_HEADERS, sheet.getRange(row, 1, 1, LEAVE_HEADERS.length).getValues()[0]);
        data.attachmentUrl = existing.attachmentUrl || '';
      }
      sheet.getRange(row, 1, 1, LEAVE_HEADERS.length).setValues([objToRow(LEAVE_HEADERS, data)]);
      return jsonResp({ success: true });
    }
    if (action === 'deleteLeave') {
      const sheet = getLeaveSheet();
      const row = findRowById(sheet, id);
      if (row < 0) return jsonResp({ success: false, error: 'Leave request not found' });
      sheet.deleteRow(row);
      return jsonResp({ success: true });
    }
    if (action === 'approveLeave' || action === 'rejectLeave') {
      const sheet = getLeaveSheet();
      const row = findRowById(sheet, id);
      if (row < 0) return jsonResp({ success: false, error: 'Leave request not found' });
      const existing = rowToObj(LEAVE_HEADERS, sheet.getRange(row, 1, 1, LEAVE_HEADERS.length).getValues()[0]);
      const now = new Date().toISOString();
      existing.status = action === 'approveLeave' ? 'approved' : 'rejected';
      existing.approver = body.approver || '';
      existing.approvedAt = now;
      existing.rejectReason = action === 'rejectLeave' ? (body.rejectReason || '') : '';
      existing.updatedAt = now;
      sheet.getRange(row, 1, 1, LEAVE_HEADERS.length).setValues([objToRow(LEAVE_HEADERS, existing)]);
      return jsonResp({ success: true });
    }

    return jsonResp({ success: false, error: 'Unknown action' });
  } catch (err) {
    return jsonResp({ success: false, error: err.toString() });
  }
}

// ── Helpers ────────────────────────────────────────────────────
function fixTextFields(sheet, row, data, fields) {
  fields.forEach(function(f) {
    const col = EMP_HEADERS.indexOf(f) + 1;
    if (col > 0) {
      const cell = sheet.getRange(row, col);
      cell.setNumberFormat('@');
      cell.setValue(data[f] !== undefined ? String(data[f]) : '');
    }
  });
}

// เรียกครั้งเดียวจาก Apps Script editor (เลือกฟังก์ชันนี้แล้วกด Run)
// เพื่อตั้งฟอร์แมตคอลัมน์เบอร์ติดต่อ/เบอร์ฉุกเฉินทั้งคอลัมน์เป็น Plain text
// (ป้องกันเลข 0 นำหน้าหายในแถวที่มีอยู่แล้วซึ่งยังไม่เคยถูกแก้ไขผ่านแอป)
function fixPhoneColumnFormat() {
  const sheet = getEmpSheet();
  const rows = Math.max(sheet.getMaxRows(), 1000);
  ['phone', 'emergencyContact', 'idCardNo', 'taxId'].forEach(function(f) {
    const col = EMP_HEADERS.indexOf(f) + 1;
    if (col > 0) sheet.getRange(1, col, rows, 1).setNumberFormat('@');
  });
}

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
