require('dotenv').config(); 
const express = require('express');
const helmet = require('helmet'); //  แก้ไข 1: เพิ่มการ Import helmet
const mysql = require('mysql2');
const app = express();

// เปิดใช้งาน HSTS Header ป้องกัน Protocol Downgrade
app.use(helmet.hsts({
    maxAge: 31536000,           // บังคับใช้เป็นเวลา 1 ปี
    includeSubDomains: true,    // มีผลกับ Subdomains ทั้งหมด
    preload: true
}));

const connection = mysql.createConnection({
    host: process.env.DB_HOST || 'localhost',
    user: process.env.DB_USER || 'admin',
    password: process.env.DB_PASSWORD, 
    database: process.env.DB_NAME || 'user_db'
});

// ดึงข้อมูลโปรไฟล์อย่างปลอดภัย
app.get('/myProfile', function(req, res) {
    var id = req.session ? req.session.userid : null; // ป้องกันกรณี req.session เป็น undefined

    // ตรวจสอบสถานะการ Login
    if (!id) {
        return res.status(401).send("Unauthorized: Please login first");
    }

    //  แก้ไข 3: เปลี่ยนเป็น connection.execute เพื่อการทำ Prepared Statement ที่สมบูรณ์ใน mysql2
    connection.execute('SELECT username, email, role FROM users WHERE id = ?', [id], function(err, results) {
        if (err) {
            console.error(err); 
            return res.status(500).send("Internal Server Error");
        }

        if (results.length === 0) {
            return res.status(404).send("User not found");
        }

        res.json(results[0]);
    });
}); //  แก้ไข 2: ใส่ปีกกาและวงเล็บปิดหน้าบ้านของ app.get ให้ครบถ้วน

app.listen(3000, () => {
    console.log('Secure server running on port 3000');
});
