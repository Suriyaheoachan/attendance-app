const express = require('express');
const cookieParser = require('cookie-parser');
const path = require('path');
const authRoutes = require('./routes/auth');
const attendanceRoutes = require('./routes/attendance');

const app = express();
app.use(express.json());
app.use(express.urlencoded({ extended: true }));
app.use(cookieParser());
app.use(express.static(path.join(__dirname, 'public')));

app.use('/api/auth', authRoutes);
app.use('/api/attendance', attendanceRoutes);

// สำหรับรันในเครื่อง local
if (require.main === module) {
  app.listen(3000, () => console.log('✅ Server running: http://localhost:3000'));
}

module.exports = app; // ← export ให้ Vercel ใช้