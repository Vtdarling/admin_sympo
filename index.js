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
// Note: Ensure your IP is whitelisted in MongoDB Atlas for this connection string to work.
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

const normalizeValue = (value) => (value ? String(value).trim().toLowerCase() : '');

const sortStudentsByEventId = (students) => {
    return [...students].sort((a, b) => {
        const aId = a.event_id ? String(a.event_id).trim() : '';
        const bId = b.event_id ? String(b.event_id).trim() : '';

        const aNumMatch = aId.match(/\d+/);
        const bNumMatch = bId.match(/\d+/);
        const aNum = aNumMatch ? Number(aNumMatch[0]) : Number.NaN;
        const bNum = bNumMatch ? Number(bNumMatch[0]) : Number.NaN;

        const aHasNum = Number.isFinite(aNum);
        const bHasNum = Number.isFinite(bNum);

        if (aHasNum && bHasNum && aNum !== bNum) {
            return aNum - bNum;
        }

        if (aHasNum && !bHasNum) return -1;
        if (!aHasNum && bHasNum) return 1;

        return aId.localeCompare(bId, undefined, { numeric: true, sensitivity: 'base' });
    });
};

const isCompletedRegistration = (student) => {
    const transactionId = normalizeValue(student.transaction_id);
    const invalidTransactionValues = new Set(['', '-', 'pending', 'n/a', 'na', 'none', 'null', 'undefined']);

    const technicalEvent = normalizeValue(student.technical_event);
    const nonTechnicalEvent = normalizeValue(student.non_technical_event);
    const hasSelectedEvent =
        (technicalEvent !== '' && technicalEvent !== 'none') ||
        (nonTechnicalEvent !== '' && nonTechnicalEvent !== 'none');

    return !invalidTransactionValues.has(transactionId) && hasSelectedEvent;
};

// --- ROUTES ---

// Dashboard (Home Route - No Login Required)
app.get('/', async (req, res) => {
    try {
        const students = await User.find();
        const sortedStudents = sortStudentsByEventId(students);
        const completedStudents = sortedStudents.filter(isCompletedRegistration);
        
        // Calculate Stats for the Dashboard
        const stats = {
            total: completedStudents.length,
            technical: completedStudents.filter(s => s.technical_event && s.technical_event !== 'None').length,
            nonTechnical: completedStudents.filter(s => s.non_technical_event && s.non_technical_event !== 'None').length,
            colleges: new Set(completedStudents.map(s => s.college)).size // Count unique colleges
        };

        res.render('dashboard', { users: completedStudents, stats: stats }); 
    } catch (err) {
        console.error(err);
        res.status(500).send("Error fetching data from database.");
    }
});

// --- EXCEL EXPORT ---
app.get('/export-excel', async (req, res) => {
    try {
        const students = await User.find();
        const sortedStudents = sortStudentsByEventId(students);
        const completedStudents = sortedStudents.filter(isCompletedRegistration);
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

        completedStudents.forEach((student) => {
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
        const students = await User.find();
        const sortedStudents = sortStudentsByEventId(students);
        const completedStudents = sortedStudents.filter(isCompletedRegistration);
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

        completedStudents.forEach((student, i) => {
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