const express = require('express');
const mysql = require('mysql2');
const app = express();

// ❌ ช่องโหว่ที่ 1: Hardcoded Credentials (Black Duck SCA / SAST ควรจะแจ้งเตือน)
// การใส่รหัสผ่านและคีย์ต่างๆ ไว้ในโค้ดโดยตรง เป็นสิ่งที่ไม่ปลอดภัย
const DB_PASSWORD = "SuperSecretPassword123!"; 
const API_KEY = "AIzaSyA1B2C3D4E5F6G7H8I9J0K1L2M3N4O5P"; 

const connection = mysql.createConnection({
    host: 'localhost',
    user: 'admin',
    password: DB_PASSWORD,
    database: 'user_db'
});

// ❌ ช่องโหว่ที่ 2: SQL Injection (SAST Flaw)
// รับค่าจาก URL directly แล้วเอาไปต่อสตริงใน SQL query โดยไม่มีการ Protect
app.get('/api/users', (req, res) => {
    const userId = req.query.id;
    
    // โค้ดที่ไม่ปลอดภัย (Vulnerable Code)
    const query = `SELECT * FROM users WHERE id = '${userId}'`;
    
    connection.query(query, (err, results) => {
        if (err) {
            return res.status(500).send("Database error");
        }
        res.json(results);
    });
});

app.listen(3000, () => {
    console.log('Test server running on port 3000');
});