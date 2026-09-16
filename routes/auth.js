const jwt = require('jsonwebtoken');
const JWT_SECRET = process.env.JWT_SECRET; // ตั้งใน environment variable ห้าม hardcode

router.post('/login', async (req, res) => {
  const { supervisor_id, password } = req.body;
  const supervisor = await findSupervisorById(supervisor_id);

  if (!supervisor || !(await verifyPassword(password, supervisor.password_hash))) {
    return res.status(401).json({ error: 'รหัสผู้ใช้งานหรือรหัสผ่านไม่ถูกต้อง' });
  }

  const payload = {
    supervisor_id: supervisor.supervisor_id,
    name: fullName(supervisor),
    company: supervisor.company,
    department: supervisor.department,
    role: supervisor.role
  };

  const token = jwt.sign(payload, JWT_SECRET, { expiresIn: '8h' });

  res.cookie('token', token, {
    httpOnly: true,      // JS ฝั่ง browser แตะไม่ได้ ป้องกัน XSS
    secure: true,         // ส่งผ่าน HTTPS เท่านั้น (Vercel เป็น HTTPS อยู่แล้ว)
    sameSite: 'lax',
    maxAge: 8 * 60 * 60 * 1000
  });

  res.json({ message: 'เข้าสู่ระบบสำเร็จ', supervisor: payload });
});

router.post('/logout', (req, res) => {
  res.clearCookie('token');
  res.json({ message: 'ออกจากระบบแล้ว' });
});

router.get('/me', (req, res) => {
  const token = req.cookies.token;
  if (!token) return res.status(401).json({ error: 'ยังไม่ได้ login' });
  try {
    const decoded = jwt.verify(token, JWT_SECRET);
    res.json(decoded);
  } catch {
    res.status(401).json({ error: 'session หมดอายุ กรุณา login ใหม่' });
  }
});

module.exports = router;