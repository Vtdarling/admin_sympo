const express = require('express');
const mongoose = require('mongoose');
const bodyParser = require('body-parser');
const path = require('path');
const ExcelJS = require('exceljs');
const PDFDocument = require('pdfkit'); // Import PDF Library

const app = express();

// --- CONFIGURATION ---
app.set('views', path.join(__dirname, 'views'));
app.set('view engine', 'ejs');
app.use(express.static(path.join(__dirname, 'public')));
app.use(bodyParser.urlencoded({ extended: true }));

// --- DATABASE CONNECTION ---
const dbURI = "mongodb+srv://720723110803_db_user:darling%40123@cluster0.ddwhmyt.mongodb.net/symposiumDB?retryWrites=true&w=majority&appName=Cluster0";

mongoose.connect(dbURI)
    .then(() => console.log("✅ Admin Panel: Connected to Cloud Database"))
    .catch(err => console.log("❌ DB Error:", err));

// --- SCHEMA ---
const studentSchema = new mongoose.Schema({
    name: String,
    college: String,
    email: String,
    phone: String,
    events: [String],
    timestamp: { type: Date, default: Date.now }
});

const Student = mongoose.model('Student', studentSchema);

// --- ROUTES ---

// Login
app.get('/', (req, res) => res.render('login', { error: null }));

// Login Logic
app.post('/login', (req, res) => {
    const { username, password } = req.body;
    if (username === "admin" && password === "sympo2025") {
        res.redirect('/dashboard');
    } else {
        res.render('login', { error: "Invalid Credentials" });
    }
});

// Dashboard
app.get('/dashboard', async (req, res) => {
    try {
        const students = await Student.find().sort({ timestamp: -1 });
        res.render('dashboard', { students: students });
    } catch (err) {
        res.send("Error fetching data.");
    }
});

// --- EXCEL EXPORT ---
app.get('/export-excel', async (req, res) => {
    try {
        const students = await Student.find().sort({ timestamp: -1 });
        const workbook = new ExcelJS.Workbook();
        const worksheet = workbook.addWorksheet('Registrations');

        worksheet.columns = [
            { header: 'S.No', key: 's_no', width: 8 },
            { header: 'Full Name', key: 'name', width: 25 },
            { header: 'College Name', key: 'college', width: 30 },
            { header: 'Phone', key: 'phone', width: 15 },
            { header: 'Email', key: 'email', width: 30 },
            { header: 'Events', key: 'events', width: 40 }
        ];

        students.forEach((student, index) => {
            worksheet.addRow({
                s_no: index + 1,
                name: student.name,
                college: student.college,
                phone: student.phone,
                email: student.email,
                events: student.events.join(', ')
            });
        });

        worksheet.getRow(1).font = { bold: true };
        res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        res.setHeader('Content-Disposition', 'attachment; filename=Symposium_Data.xlsx');
        await workbook.xlsx.write(res);
        res.end();
    } catch (err) {
        console.log(err);
        res.send("Error generating Excel.");
    }
});

// --- PDF EXPORT (New Feature) ---
app.get('/export-pdf', async (req, res) => {
    try {
        const students = await Student.find().sort({ timestamp: -1 });
        const doc = new PDFDocument({ margin: 30, size: 'A4' });

        // Set Headers to download file
        res.setHeader('Content-Type', 'application/pdf');
        res.setHeader('Content-Disposition', 'attachment; filename=Symposium_List.pdf');

        doc.pipe(res); // Send PDF to browser

        // Title
        doc.fontSize(18).text('Symposium 2025 - Registration List', { align: 'center' });
        doc.moveDown();
        doc.fontSize(10).text(`Generated on: ${new Date().toLocaleString()}`, { align: 'center' });
        doc.moveDown();

        // Table Header
        const tableTop = 130;
        const col1 = 30;  // S.No
        const col2 = 70;  // Name
        const col3 = 200; // College
        const col4 = 350; // Events
        const col5 = 500; // Phone

        doc.fontSize(10).font('Helvetica-Bold');
        doc.text('S.No', col1, tableTop);
        doc.text('Name', col2, tableTop);
        doc.text('College', col3, tableTop);
        doc.text('Events', col4, tableTop);
        doc.text('Phone', col5, tableTop);

        doc.moveTo(30, tableTop + 15).lineTo(570, tableTop + 15).stroke(); // Line

        // Table Rows
        let y = tableTop + 25;
        doc.font('Helvetica').fontSize(9);

        students.forEach((student, i) => {
            // New page if we reach the bottom
            if (y > 750) {
                doc.addPage();
                y = 50;
            }

            doc.text(i + 1, col1, y);
            doc.text(student.name, col2, y, { width: 120 });
            doc.text(student.college, col3, y, { width: 140 });
            doc.text(student.events.join(', '), col4, y, { width: 140 });
            doc.text(student.phone, col5, y);

            y += 30; // Row height spacing
        });

        doc.end(); // Finish PDF

    } catch (err) {
        console.log(err);
        res.send("Error generating PDF");
    }
});

// --- START SERVER ---
const PORT = 4000;
app.listen(PORT, () => {
    console.log(`🛡️ Admin Panel running at: http://localhost:${PORT}`);
});