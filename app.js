require('dotenv').config(); // 1. โหลดค่าคอนฟิกจากไฟล์ .env
const express = require('express');
const mysql = require('mysql2');
const app = express();

// ✅ แก้ไขช่องโหว่ที่ 1: ดึงค่าจาก Environment Variables แทนการเขียนลงในโค้ดโดยตรง
const connection = mysql.createConnection({
    host: process.env.DB_HOST || 'localhost',
    user: process.env.DB_USER || 'admin',
    password: process.env.DB_PASSWORD, // ดึงมาจาก .env
    database: process.env.DB_NAME || 'user_db'
});

// ✅ แก้ไขช่องโหว่ที่ 2: ป้องกัน SQL Injection ด้วย Prepared Statements
app.get('/api/users', (req, res) => {
    const userId = req.query.id;
    
    // ใช้เครื่องหมาย ? แทนการต่อสตริงโดยตรง (Placeholder)
    const query = 'SELECT * FROM users WHERE id = ?';
    
    // ส่งค่า userId แยกไปใน Array ระบบ Database จะมองค่านี้เป็น Literal Value เสมอ ไม่ใช่คำสั่ง SQL
    connection.execute(query, [userId], (err, results) => {
        if (err) {
            // หลีกเลี่ยงการพ่น Error ละเอียดของ DB ออกไปให้ User เห็นภายนอกเพื่อความปลอดภัย
            console.error(err); 
            return res.status(500).send("Internal Server Error");
        }
        res.json(results);
    });
});

app.listen(3000, () => {
    console.log('Secure server running on port 3000');
});
