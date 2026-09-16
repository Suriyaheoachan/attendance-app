const ExcelJS = require('exceljs');
const path = require('path');

const EMPLOYEE_FILE = path.join(__dirname, '../data/Data_Name.xlsx');
const REQUEST_FILE = path.join(__dirname, '../data/Attendance.xlsx');

// ---------- Employees ----------
async function getEmployees() {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(EMPLOYEE_FILE);
  const sheet = workbook.getWorksheet('Employees');
  const employees = [];
  sheet.eachRow((row, rowNumber) => {
    if (rowNumber === 1) return;
    employees.push({
      employee_id: row.getCell(1).value,
      company: row.getCell(2).value,
      department: row.getCell(3).value,
      prefix: row.getCell(4).value,
      first_name: row.getCell(5).value,
      last_name: row.getCell(6).value,
      leave_balance: {
        'ลากิจ': Number(row.getCell(7).value) || 0,
        'ลากิจจ่าย': Number(row.getCell(8).value) || 0,
        'ลาป่วย': Number(row.getCell(9).value) || 0,
        'ลาป่วยจ่าย': Number(row.getCell(10).value) || 0,
        'ลาคลอด': Number(row.getCell(11).value) || 0,
        'ลาพักร้อน': Number(row.getCell(12).value) || 0
      }
    });
  });
  return employees;
}

function fullName(p) {
  return `${p.prefix}${p.first_name} ${p.last_name}`;
}

async function getEmployeesByScope(company, department) {
  const employees = await getEmployees();
  return employees.filter(e => {
    const companyMatch = !company || e.company === company;
    const deptMatch = !department || e.department === department;
    return companyMatch && deptMatch;
  });
}

// ---------- Supervisors / Admin ----------
async function findSupervisorById(supervisorId) {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(EMPLOYEE_FILE);
  const sheet = workbook.getWorksheet('Supervisors');
  let found = null;
  sheet.eachRow((row, rowNumber) => {
    if (rowNumber === 1) return;
    if (String(row.getCell(1).value) === String(supervisorId)) {
      found = {
        supervisor_id: row.getCell(1).value,
        prefix: row.getCell(2).value,
        first_name: row.getCell(3).value,
        last_name: row.getCell(4).value,
        password: row.getCell(5).value,
        company: row.getCell(6).value,
        department: row.getCell(7).value,
        role: row.getCell(8).value, // 'admin' หรือ 'supervisor'
      };
    }
  });
  return found;
}

// ---------- ขอเพิ่มเวลา (แก้ไขเวลาลืมสแกน) ----------
const TIME_CORRECTION_REASONS = [
  'ลืมสแกนเข้างาน', 'ลืมสแกนออกงาน', 'เครื่องสแกนนิ้วขัดข้อง',
  'ระบบเครือข่ายขัดข้อง', 'ลืมนำบัตรพนักงานมา', 'อื่นๆ'
];

async function addTimeCorrectionRequest(r) {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(REQUEST_FILE);
  const sheet = workbook.getWorksheet('TimeCorrection');
  sheet.addRow([
    Date.now(), r.employee_id, r.name, r.company, r.department,
    new Date().toISOString(), r.correction_date, r.correction_time,
    r.reason_type, r.detail, r.submitted_by
  ]);
  await workbook.xlsx.writeFile(REQUEST_FILE);
}

async function getTimeCorrectionHistory(filters) {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(REQUEST_FILE);
  const sheet = workbook.getWorksheet('TimeCorrection');
  const records = [];
  sheet.eachRow((row, rowNumber) => {
    if (rowNumber === 1) return;
    records.push({
      employee_id: row.getCell(2).value,
      name: row.getCell(3).value,
      company: row.getCell(4).value,
      department: row.getCell(5).value,
      correction_date: row.getCell(7).value,
      correction_time: row.getCell(8).value,
      reason_type: row.getCell(9).value,
      detail: row.getCell(10).value,
    });
  });
  return records.filter(r => {
    if (filters.company && r.company !== filters.company) return false;
    if (filters.department && r.department !== filters.department) return false;
    if (filters.name && !String(r.name).includes(filters.name)) return false;
    if (filters.date_from && new Date(r.correction_date) < new Date(filters.date_from)) return false;
    if (filters.date_to && new Date(r.correction_date) > new Date(filters.date_to)) return false;
    return true;
  });
}

// ---------- ขอ OT (ทำงานล่วงเวลาจริง) ----------
function calculateOTHours(startStr, endStr) {
  const start = new Date(startStr);
  const end = new Date(endStr);
  if (end <= start) return 0;
  return Math.round(((end - start) / (1000 * 60 * 60)) * 100) / 100;
}

async function addOTRequest(r) {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(REQUEST_FILE);
  const sheet = workbook.getWorksheet('OT');
  sheet.addRow([
    Date.now(), r.employee_id, r.name, r.company, r.department,
    new Date().toISOString(), r.ot_date, r.ot_start_time, r.ot_end_time,
    r.ot_hours, r.detail, r.submitted_by
  ]);
  await workbook.xlsx.writeFile(REQUEST_FILE);
}

async function getOTHistory(filters) {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(REQUEST_FILE);
  const sheet = workbook.getWorksheet('OT');
  const records = [];
  sheet.eachRow((row, rowNumber) => {
    if (rowNumber === 1) return;
    records.push({
      employee_id: row.getCell(2).value,
      name: row.getCell(3).value,
      company: row.getCell(4).value,
      department: row.getCell(5).value,
      ot_date: row.getCell(7).value,
      ot_start_time: row.getCell(8).value,
      ot_end_time: row.getCell(9).value,
      ot_hours: row.getCell(10).value,
    });
  });
  return records.filter(r => {
    if (filters.company && r.company !== filters.company) return false;
    if (filters.department && r.department !== filters.department) return false;
    if (filters.name && !String(r.name).includes(filters.name)) return false;
    if (filters.date_from && new Date(r.ot_date) < new Date(filters.date_from)) return false;
    if (filters.date_to && new Date(r.ot_date) > new Date(filters.date_to)) return false;
    return true;
  });
}

// ---------- ขอลางาน ----------
async function addLeaveRequest(r) {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(REQUEST_FILE);
  const sheet = workbook.getWorksheet('Leave');
  sheet.addRow([
    Date.now(), r.employee_id, r.name, r.company, r.department,
    new Date().toISOString(), r.leave_type, r.start_datetime, r.end_datetime,
    r.detail, r.submitted_by, r.days_used   // ← คอลัมน์ L ใหม่
  ]);
  await workbook.xlsx.writeFile(REQUEST_FILE);
}

async function getLeaveHistory(filters) {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(REQUEST_FILE);
  const sheet = workbook.getWorksheet('Leave');
  const records = [];
  sheet.eachRow((row, rowNumber) => {
    if (rowNumber === 1) return;
    records.push({
      employee_id: row.getCell(2).value,
      name: row.getCell(3).value,
      company: row.getCell(4).value,
      department: row.getCell(5).value,
      leave_type: row.getCell(7).value,
      start_datetime: row.getCell(8).value,
      end_datetime: row.getCell(9).value,
      days_used: row.getCell(12).value,   // ← เพิ่มใหม่
    });
  });

  return records.filter(r => {
    if (filters.company && r.company !== filters.company) return false;
    if (filters.department && r.department !== filters.department) return false;
    if (filters.name && !String(r.name).includes(filters.name)) return false;
    if (filters.date_from && new Date(r.start_datetime) < new Date(filters.date_from)) return false;
    if (filters.date_to && new Date(r.start_datetime) > new Date(filters.date_to)) return false;
    return true;
  });
}

module.exports = {
  getEmployees, fullName, getEmployeesByScope,
  findSupervisorById,
  addLeaveRequest, getLeaveHistory
};

// ประเภทลาไหนมีโควตา ก็ผูกกับคอลัมน์ไหนใน Employees sheet
const LEAVE_BALANCE_COLUMNS = { 'ลากิจ': 7, 'ลากิจจ่าย': 8, 'ลาป่วย': 9, 'ลาป่วยจ่าย': 10, 'ลาคลอด': 11, 'ลาพักร้อน': 12 };

// เวลาทำงานมาตรฐาน — ปรับตัวเลขตรงนี้ได้ถ้าบริษัทไม่ตรงกับนี้
const WORK_START_HOUR = 8;   // เข้างาน 08:00
const WORK_END_HOUR = 17;    // เลิกงาน 17:00
const LUNCH_START_HOUR = 12; // พักเที่ยงเริ่ม 12:00
const LUNCH_END_HOUR = 13;   // พักเที่ยงจบ 13:00
const HOURS_PER_DAY = 8;     // 1 วันทำงาน = 8 ชม.

// คำนวณชั่วโมงทำงานจริงของวันเดียว โดยตัดเฉพาะช่วงในกรอบเวลาทำงาน และหักพักเที่ยง
function getDayWorkHours(dayDate, rangeStart, rangeEnd) {
  const dayWorkStart = new Date(dayDate); dayWorkStart.setHours(WORK_START_HOUR, 0, 0, 0);
  const dayWorkEnd = new Date(dayDate); dayWorkEnd.setHours(WORK_END_HOUR, 0, 0, 0);

  const start = new Date(Math.max(rangeStart, dayWorkStart));
  const end = new Date(Math.min(rangeEnd, dayWorkEnd));
  if (end <= start) return 0;

  let hours = (end - start) / (1000 * 60 * 60);

  // หักช่วงพักเที่ยงถ้าคาบเกี่ยวกับช่วงที่ขอ
  const lunchStart = new Date(dayDate); lunchStart.setHours(LUNCH_START_HOUR, 0, 0, 0);
  const lunchEnd = new Date(dayDate); lunchEnd.setHours(LUNCH_END_HOUR, 0, 0, 0);
  const overlapStart = new Date(Math.max(start, lunchStart));
  const overlapEnd = new Date(Math.min(end, lunchEnd));
  if (overlapEnd > overlapStart) {
    hours -= (overlapEnd - overlapStart) / (1000 * 60 * 60);
  }
  return Math.max(0, hours);
}

// คำนวณจำนวนวันลา จากช่วงวันที่/เวลาเริ่ม-สิ้นสุด (นับเป็นชั่วโมง แล้วแปลงเป็นวัน)
function calculateLeaveDays(startStr, endStr) {
  const start = new Date(startStr);
  const end = new Date(endStr);
  if (end <= start) return 0;

  let totalHours = 0;
  const cursor = new Date(start);
  cursor.setHours(0, 0, 0, 0);
  const lastDay = new Date(end);
  lastDay.setHours(0, 0, 0, 0);

  while (cursor <= lastDay) {
    totalHours += getDayWorkHours(cursor, start, end);
    cursor.setDate(cursor.getDate() + 1);
  }

  const rawDays = totalHours / HOURS_PER_DAY;
  return Math.round(rawDays * 2) / 2; // ปัดให้เป็นทวีคูณของ 0.5 ใกล้สุด (0, 0.5, 1, 1.5, ...)
}

// หักยอดวันลาคงเหลือใน Employees sheet หลังลาสำเร็จ
async function deductLeaveBalance(employeeId, leaveType, days) {
  const colIndex = LEAVE_BALANCE_COLUMNS[leaveType];
  if (!colIndex) return; // ประเภทที่ไม่มีโควตา (เช่น ลาคลอด, อื่นๆ) ไม่ต้องหัก

  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(EMPLOYEE_FILE);
  const sheet = workbook.getWorksheet('Employees');
  sheet.eachRow((row) => {
    if (String(row.getCell(1).value) === String(employeeId)) {
      const current = Number(row.getCell(colIndex).value) || 0;
      row.getCell(colIndex).value = Math.max(0, current - days);
    }
  });
  await workbook.xlsx.writeFile(EMPLOYEE_FILE);
}

async function addLeaveRequest(r) {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(REQUEST_FILE);
  const sheet = workbook.getWorksheet('Leave');
  sheet.addRow([
    Date.now(), r.employee_id, r.name, r.company, r.department,
    new Date().toISOString(), r.leave_type, r.start_datetime, r.end_datetime,
    r.detail, r.submitted_by, r.days_used,
    r.image_path || ''   // ← คอลัมน์ M ใหม่
  ]);
  await workbook.xlsx.writeFile(REQUEST_FILE);
}

async function getLeaveHistory(filters) {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(REQUEST_FILE);
  const sheet = workbook.getWorksheet('Leave');
  const records = [];
  sheet.eachRow((row, rowNumber) => {
    if (rowNumber === 1) return;
    records.push({
      employee_id: row.getCell(2).value,
      name: row.getCell(3).value,
      company: row.getCell(4).value,
      department: row.getCell(5).value,
      leave_type: row.getCell(7).value,
      start_datetime: row.getCell(8).value,
      end_datetime: row.getCell(9).value,
      days_used: row.getCell(12).value,
      image_path: row.getCell(13).value || '',   // ← เพิ่มใหม่
    });
  });
  return records.filter(r => {
    if (filters.company && r.company !== filters.company) return false;
    if (filters.department && r.department !== filters.department) return false;
    if (filters.name && !String(r.name).includes(filters.name)) return false;
    if (filters.date_from && new Date(r.start_datetime) < new Date(filters.date_from)) return false;
    if (filters.date_to && new Date(r.start_datetime) > new Date(filters.date_to)) return false;
    return true;
  });
}

module.exports = {
  getEmployees, fullName, getEmployeesByScope,
  findSupervisorById,
  addLeaveRequest, getLeaveHistory,
  calculateLeaveDays, deductLeaveBalance,
  TIME_CORRECTION_REASONS,
  addTimeCorrectionRequest, getTimeCorrectionHistory,
  calculateOTHours, addOTRequest, getOTHistory
};