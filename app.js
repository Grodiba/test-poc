require('dotenv').config(); // 1. โหลดค่าคอนฟิกจากไฟล์ .env
const express = require('express');
const mysql = require('mysql2');
const app = express();

app.use(helmet.hsts({
    maxAge: 31536000,           // บังคับใช้เป็นเวลา 1 ปี (หน่วยเป็นวินาที)
    includeSubDomains: true,    // มีผลกับ Subdomains ทั้งหมดด้วย
    preload: true
}));

// ✅ แก้ไขช่องโหว่ที่ 1: ดึงค่าจาก Environment Variables แทนการเขียนลงในโค้ดโดยตรง
const connection = mysql.createConnection({
    host: process.env.DB_HOST || 'localhost',
    user: process.env.DB_USER || 'admin',
    password: process.env.DB_PASSWORD, // ดึงมาจาก .env
    database: process.env.DB_NAME || 'user_db'
});

// ✅ แก้ไขช่องโหว่ที่ 2: ป้องกัน SQL Injection ด้วย Prepared Statements
app.get('/myProfile', function(req, res) {
    var id = req.session.userid;

    // ตรวจสอบก่อนว่า User ทำการ Login หรือยัง (Session มีค่าไหม)
    if (!id) {
        return res.status(401).send("Unauthorized: Please login first");
    }

    // เปลี่ยนจาก SELECT * เป็นเลือกเฉพาะฟิลด์ที่จำเป็นเพื่อความปลอดภัย
    connection.query('SELECT username, email, role FROM users WHERE id=?', [id], function(err, results) {
        if (err) {
            console.error(err); // Log error ไว้หลังบ้าน
            return res.status(500).send("Internal Server Error");
        }

        if (results.length === 0) {
            return res.status(404).send("User not found");
        }

        // Handle result และส่งข้อมูลกลับไปอย่างปลอดภัย
        res.json(results[0]);
    });
}

app.listen(3000, () => {
    console.log('Secure server running on port 3000');
});
