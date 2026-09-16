module.exports = require('../server');
const express = require('express');
const cookieParser = require('cookie-parser');
const authRoutes = require('../routes/auth');
const attendanceRoutes = require('../routes/attendance');

const app = express();
app.use(express.json());
app.use(cookieParser());

app.use('/api/auth', authRoutes);
app.use('/api/attendance', attendanceRoutes);

module.exports = app;   // ⚠️ ไม่มี app.listen() เพราะ Vercel เป็นคนเรียกฟังก์ชันนี้เอง
