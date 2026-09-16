const multer = require('multer');
const upload = multer({ storage: multer.memoryStorage() }); // เก็บใน RAM ชั่วคราว ไม่เขียนลง disk

router.post('/leave', requireLogin, upload.single('leave_image'), async (req, res) => {
  // ... validation เดิมทั้งหมด ...

 let image_path = '';
  if (req.file) {
    const fileName = `leave_${Date.now()}_${req.file.originalname}`;
    const { data, error } = await supabase.storage
      .from('leave-images')
      .upload(fileName, req.file.buffer, { contentType: req.file.mimetype });

    if (!error) {
      const { data: urlData } = supabase.storage.from('leave-images').getPublicUrl(fileName);
      image_path = urlData.publicUrl;
    }
  }

  await addLeaveRequest({ ...req.body, image_path, submitted_by: req.supervisor.supervisor_id });
  res.json({ message: 'นำส่งข้อมูลสำเร็จ' });
});

const express = require('express');
const router = express.Router();
const {
  getEmployeesByScope, fullName, getEmployees,
  addLeaveRequest, getLeaveHistory,
  calculateLeaveDays, deductLeaveBalance,
  TIME_CORRECTION_REASONS,
  addTimeCorrectionRequest, getTimeCorrectionHistory,
  calculateOTHours, addOTRequest, getOTHistory
} = require('../services/excelService');

const jwt = require('jsonwebtoken');
const JWT_SECRET = process.env.JWT_SECRET;

function requireLogin(req, res, next) {
  const token = req.cookies.token;
  if (!token) return res.status(401).json({ error: 'กรุณา login ก่อน' });
  try {
    req.user = jwt.verify(token, JWT_SECRET); 
    next();
  } catch {
    res.status(401).json({ error: 'session หมดอายุ กรุณา login ใหม่' });
  }
}

function requireAdmin(req, res, next) {
  if (!req.user || req.user.role !== 'admin') {
    return res.status(403).json({ error: 'เฉพาะ admin เท่านั้น' });
  }
  next();
}

router.get('/my-team', requireLogin, async (req, res) => {
  const { supervisor } = req.session;
  const company = supervisor.role === 'admin' ? (req.query.company || null) : supervisor.company;
  const department = supervisor.role === 'admin' ? (req.query.department || null) : supervisor.department;
  const employees = await getEmployeesByScope(company, department);
  res.json(employees.map(e => ({ ...e, name: fullName(e) })));
});

// รายการเหตุผลลืมสแกน (ให้ frontend ดึงไปทำ dropdown)
router.get('/time-correction/reasons', requireLogin, (req, res) => {
  res.json(TIME_CORRECTION_REASONS);
});

// ---------- ขอเพิ่มเวลา (ลืมสแกน) ----------
router.post('/time-correction', requireLogin, async (req, res) => {
  const { employee_id, name, company, department, correction_date, correction_time, reason_type, detail } = req.body;
  if (!correction_date || !correction_time || !reason_type) {
    return res.status(400).json({ error: 'กรุณากรอกวันที่ เวลา และเหตุผลให้ครบ' });
  }
  await addTimeCorrectionRequest({
    employee_id, name, company, department, correction_date, correction_time, reason_type, detail,
    submitted_by: req.supervisor.supervisor_id
  });
  res.json({ message: 'นำส่งข้อมูลสำเร็จ' });
});

router.get('/time-correction/history', requireAdmin, async (req, res) => {
  const { date_from, date_to, name, company, department } = req.query;
  const records = await getTimeCorrectionHistory({ date_from, date_to, name, company, department });
  res.json(records);
});

// ---------- ขอ OT ----------
router.post('/ot', requireLogin, async (req, res) => {
  const { employee_id, name, company, department, ot_date, ot_start_time, ot_end_time, detail } = req.body;
  if (!ot_date || !ot_start_time || !ot_end_time) {
    return res.status(400).json({ error: 'กรุณากรอกวันที่และเวลาให้ครบ' });
  }
  const startDt = `${ot_date}T${ot_start_time}`;
  const endDt = `${ot_date}T${ot_end_time}`;
  const ot_hours = calculateOTHours(startDt, endDt);
  if (ot_hours <= 0) {
    return res.status(400).json({ error: 'เวลาสิ้นสุดต้องมากกว่าเวลาเริ่ม' });
  }
  await addOTRequest({
    employee_id, name, company, department, ot_date, ot_start_time, ot_end_time, ot_hours, detail,
    submitted_by: req.supervisor.supervisor_id
  });
  res.json({ message: 'นำส่งข้อมูลสำเร็จ' });
});

router.get('/ot/history', requireAdmin, async (req, res) => {
  const { date_from, date_to, name, company, department } = req.query;
  const records = await getOTHistory({ date_from, date_to, name, company, department });
  res.json(records);
});

// ---------- ขอลางาน (เดิม ไม่เปลี่ยน) ----------
router.post('/leave', requireLogin, upload.single('leave_image'), async (req, res) => {
  const { employee_id, name, company, department, leave_type, start_datetime, end_datetime, detail } = req.body;
  if (!start_datetime || !end_datetime) {
    return res.status(400).json({ error: 'กรุณาระบุวันที่เริ่มและถึงวันที่' });
  }

  if (leave_type === 'ลาพักร้อน') {
    const now = new Date();
    const start = new Date(start_datetime);
    const noticeDays = (start - now) / (1000 * 60 * 60 * 24);
    if (noticeDays < 3) {
      return res.status(400).json({ error: 'ลาพักร้อนต้องขอล่วงหน้าอย่างน้อย 3 วัน' });
    }
  }

  const requestedDays = calculateLeaveDays(start_datetime, end_datetime);
  const employees = await getEmployees();
  const employee = employees.find(e => String(e.employee_id) === String(employee_id));

  if (employee && employee.leave_balance.hasOwnProperty(leave_type)) {
    const remaining = employee.leave_balance[leave_type];
    if (remaining <= 0) {
      return res.status(400).json({ error: `วันลาประเภท "${leave_type}" ไม่มีวันคงเหลือแล้ว` });
    }
    if (requestedDays > remaining) {
      return res.status(400).json({ error: `วันลาคงเหลือไม่พอ (เหลือ ${remaining} วัน แต่ขอ ${requestedDays} วัน)` });
    }
  }

  await addLeaveRequest({
    employee_id, name, company, department, leave_type, start_datetime, end_datetime, detail,
    submitted_by: req.supervisor.supervisor_id,
    days_used: requestedDays
  });
  await deductLeaveBalance(employee_id, leave_type, requestedDays);

  res.json({ message: 'นำส่งข้อมูลสำเร็จ' });
});

router.get('/leave/history', requireAdmin, async (req, res) => {
  const { date_from, date_to, name, company, department } = req.query;
  const records = await getLeaveHistory({ date_from, date_to, name, company, department });
  res.json(records);
});

router.post('/leave', requireLogin, upload.single('leave_image'), async (req, res) => {
  const { employee_id, name, company, department, leave_type, start_datetime, end_datetime, detail } = req.body;

  if (!start_datetime || !end_datetime) {
    return res.status(400).json({ error: 'กรุณาระบุวันที่เริ่มและถึงวันที่' });
  }

  if (leave_type === 'ลาพักร้อน') {
    const now = new Date();
    const start = new Date(start_datetime);
    const noticeDays = (start - now) / (1000 * 60 * 60 * 24);
    if (noticeDays < 3) {
      return res.status(400).json({ error: 'ลาพักร้อนต้องขอล่วงหน้าอย่างน้อย 3 วัน' });
    }
  }

  const requestedDays = calculateLeaveDays(start_datetime, end_datetime);
  const employees = await getEmployees();
  const employee = employees.find(e => String(e.employee_id) === String(employee_id));

  if (employee && employee.leave_balance.hasOwnProperty(leave_type)) {
    const remaining = employee.leave_balance[leave_type];
    if (remaining <= 0) {
      return res.status(400).json({ error: `วันลาประเภท "${leave_type}" ไม่มีวันคงเหลือแล้ว` });
    }
    if (requestedDays > remaining) {
      return res.status(400).json({ error: `วันลาคงเหลือไม่พอ (เหลือ ${remaining} วัน แต่ขอ ${requestedDays} วัน)` });
    }
  }

  const image_path = req.file ? `/uploads/${req.file.filename}` : '';

  await addLeaveRequest({
    employee_id, name, company, department, leave_type, start_datetime, end_datetime, detail,
    submitted_by: req.supervisor.supervisor_id,
    days_used: requestedDays,
    image_path
  });
  await deductLeaveBalance(employee_id, leave_type, requestedDays);

  res.json({ message: 'นำส่งข้อมูลสำเร็จ' });
});

module.exports = router;