const express = require('express');
const mongoose = require('mongoose');
const bodyParser = require('body-parser');
const path = require('path');
const ExcelJS = require('exceljs');
const PDFDocument = require('pdfkit'); 

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
const userSchema = new mongoose.Schema({
    event_id: String, 
    name: String,
    email: String,
    phone: String,
    college: String,
    technical_event: String,
    non_technical_event: String,
    transaction_id: String,
    registeredAt: { type: Date, default: Date.now }
});

const User = mongoose.model('User', userSchema);

// --- ROUTES ---

// Login Page
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

// Dashboard - Fetch Users
app.get('/dashboard', async (req, res) => {
    try {
        const students = await User.find().sort({ registeredAt: -1 });
        
        // ✅ FIXED LINE: We send 'users' because dashboard.ejs uses 'users'
        res.render('dashboard', { users: students }); 
    } catch (err) {
        console.error(err);
        res.send("Error fetching data.");
    }
});

// --- EXCEL EXPORT ---
app.get('/export-excel', async (req, res) => {
    try {
        const students = await User.find().sort({ registeredAt: -1 });
        const workbook = new ExcelJS.Workbook();
        const worksheet = workbook.addWorksheet('Registrations');

        worksheet.columns = [
            { header: 'Event ID', key: 'event_id', width: 15 },
            { header: 'Full Name', key: 'name', width: 25 },
            { header: 'College', key: 'college', width: 30 },
            { header: 'Phone', key: 'phone', width: 15 },
            { header: 'Email', key: 'email', width: 30 },
            { header: 'Tech Event', key: 'tech', width: 20 },
            { header: 'Non-Tech Event', key: 'non_tech', width: 20 },
            { header: 'Trans ID', key: 'trans_id', width: 20 }
        ];

        students.forEach((student) => {
            worksheet.addRow({
                event_id: student.event_id || '-',
                name: student.name,
                college: student.college,
                phone: student.phone,
                email: student.email,
                tech: student.technical_event || 'None',
                non_tech: student.non_technical_event || 'None',
                trans_id: student.transaction_id || '-'
            });
        });

        worksheet.getRow(1).font = { bold: true };
        res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        res.setHeader('Content-Disposition', 'attachment; filename=Symposium_Data.xlsx');
        
        await workbook.xlsx.write(res);
        res.end();
    } catch (err) {
        console.log(err);
        res.status(500).send("Error generating Excel.");
    }
});

// --- PDF EXPORT ---
app.get('/export-pdf', async (req, res) => {
    try {
        const students = await User.find().sort({ registeredAt: -1 });
        const doc = new PDFDocument({ margin: 30, size: 'A4', layout: 'landscape' });

        res.setHeader('Content-Type', 'application/pdf');
        res.setHeader('Content-Disposition', 'attachment; filename=Symposium_List.pdf');

        doc.pipe(res);

        doc.fontSize(18).text('Symposium 2025 - Registration List', { align: 'center' });
        doc.moveDown();
        doc.fontSize(10).text(`Generated on: ${new Date().toLocaleString()}`, { align: 'center' });
        doc.moveDown();

        const tableTop = 130;
        const col1 = 30;  // ID
        const col2 = 90;  // Name
        const col3 = 200; // College
        const col4 = 350; // Tech Event
        const col5 = 500; // Non-Tech
        const col6 = 650; // Phone

        doc.fontSize(10).font('Helvetica-Bold');
        doc.text('ID', col1, tableTop);
        doc.text('Name', col2, tableTop);
        doc.text('College', col3, tableTop);
        doc.text('Tech Event', col4, tableTop);
        doc.text('Non-Tech', col5, tableTop);
        doc.text('Phone', col6, tableTop);

        doc.moveTo(30, tableTop + 15).lineTo(780, tableTop + 15).stroke(); 

        let y = tableTop + 25;
        doc.font('Helvetica').fontSize(9);

        students.forEach((student, i) => {
            if (y > 550) { 
                doc.addPage({ layout: 'landscape' });
                y = 50;
            }
            const tech = student.technical_event || '-';
            const nonTech = student.non_technical_event || '-';

            doc.text(student.event_id || (i+1), col1, y, { width: 50 });
            doc.text(student.name, col2, y, { width: 100 });
            doc.text(student.college, col3, y, { width: 140 });
            doc.text(tech, col4, y, { width: 140 });
            doc.text(nonTech, col5, y, { width: 140 });
            doc.text(student.phone, col6, y);

            y += 20; 
        });

        doc.end();

    } catch (err) {
        console.log(err);
        res.status(500).send("Error generating PDF");
    }
});

const PORT = 4000;
app.listen(PORT, () => {
    console.log(`🛡️ Admin Panel running at: http://localhost:${PORT}`);
});