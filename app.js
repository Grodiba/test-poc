const express = require('express');
const mysql = require('mysql2');
const app = express();

// ✅ Remediation: Sensitive data should be managed securely using environment variables or secret management tools
const DB_PASSWORD = process.env.DB_PASSWORD || ""; 
const API_KEY = process.env.API_KEY || "";

const connection = mysql.createConnection({
    host: 'localhost',
    user: 'admin',
    password: DB_PASSWORD,
    database: 'user_db'
});

// ✅ Remediated: CSRF Protection (Fix Vulnerability)
// เพิ่ม Validation สำหรับ CSRF token ก่อน Processing Request
app.get('/api/users', (req, res) => {
const userId = req.query.id;
const csrfToken = req.headers['x-csrf-token'];
 
if (!csrfToken || csrfToken !== req.session.csrfToken) {
         return res.status(403).send('Invalid CSRF token');
     }
    
    // โค้ดที่ไม่ปลอดภัย (Vulnerable Code)
    const query = 'SELECT * FROM users WHERE id = ?';
    
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
