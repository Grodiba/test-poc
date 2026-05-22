const express = require('express');
const mysql = require('mysql2');
const app = express();

// Remediation: Remove the "X-Powered-By" header to prevent sensitive server information from being disclosed.
app.disable('x-powered-by');

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
  
 const query = `SELECT * FROM users WHERE id = ?`;
  
 connection.query(query, [userId], (err, results) => {
          if (err) {
               return res.status(500).send('Database query error');
          }
          res.status(200).json(results);
     });
 });

app.listen(3000, () => {
    console.log('Test server running on port 3000');
});
