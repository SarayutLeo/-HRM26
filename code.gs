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
  'id', 'empId', 'fullName', 'nickname', 'company', 'department', 'position',
  'startDate', 'birthDate', 'address', 'phone', 'education', 'email', 'supervisor',
  'sickLeaveNoDoc', 'sickLeaveWithDoc', 'personalLeave', 'annualLeave',
  'salaryStartDate', 'baseSalary', 'positionAllowance', 'otherBenefits', 'housingAllowance',
  'status', 'resignDate', 'createdAt', 'updatedAt'
];
const EMP_LABELS = [
  'ID (ระบบ)', 'รหัสพนักงาน', 'ชื่อ-นามสกุล', 'ชื่อเล่น', 'บริษัท', 'แผนก', 'ตำแหน่ง',
  'วันที่เริ่มงาน', 'วันเกิด', 'ที่อยู่', 'เบอร์ติดต่อ', 'วุฒิการศึกษา', 'อีเมล', 'ผู้บังคับบัญชา',
  'ลาป่วย(ไม่มีใบ)', 'ลาป่วย(มีใบ)', 'ลากิจ', 'ลาพักร้อน',
  'วันที่เริ่ม(เงินเดือน)', 'เงินเดือนฐาน', 'ค่าตำแหน่ง', 'สวัสดิการอื่น', 'ค่าที่พัก',
  'สถานะ', 'วันที่ลาออก', 'วันที่สร้าง', 'วันที่แก้ไข'
];

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
  const s = initSheet(SHEET_EMP,  EMP_LABELS,  [140,110,160,90,140,120,130,110,100,200,130,150,160,150,100,100,80,90,110,110,110,120,110,160,160]);
  ensureColumns(s, EMP_LABELS);
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
