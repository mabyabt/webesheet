const express = require("express");
const multer = require("multer");
const ExcelJS = require("exceljs");
const fs = require("fs");
const { PDFDocument } = require("pdf-lib");

const app = express();
const PORT = 3000;

// Multer setup to handle file uploads
const upload = multer({ dest: "uploads/" });

app.use(express.static("public"));
app.use(express.json());

// Load Excel file and send data to frontend
app.get("/data", async (req, res) => {
    try {
        const workbook = new ExcelJS.Workbook();
        await workbook.xlsx.readFile("uploads/data.xlsx");
        const worksheet = workbook.getWorksheet(1);
        const data = [];

        worksheet.eachRow((row) => {
            data.push(row.values.slice(1)); // Remove first empty element
        });

        res.json(data);
    } catch (err) {
        res.status(500).json({ error: err.message });
    }
});

// Save edited data back to Excel
app.post("/save", async (req, res) => {
    try {
        const workbook = new ExcelJS.Workbook();
        await workbook.xlsx.readFile("uploads/data.xlsx");
        const worksheet = workbook.getWorksheet(1);

        req.body.forEach((row, rowIndex) => {
            worksheet.getRow(rowIndex + 1).values = [null, ...row];
        });

        await workbook.xlsx.writeFile("uploads/data.xlsx");
        res.json({ message: "Saved successfully" });
    } catch (err) {
        res.status(500).json({ error: err.message });
    }
});

// Generate PDF from Excel data
app.get("/export-pdf", async (req, res) => {
    try {
        const workbook = new ExcelJS.Workbook();
        await workbook.xlsx.readFile("uploads/data.xlsx");
        const worksheet = workbook.getWorksheet(1);
        const pdfDoc = await PDFDocument.create();
        const page = pdfDoc.addPage([600, 800]);
        let y = 750;

        worksheet.eachRow((row) => {
            page.drawText(row.values.slice(1).join(" | "), { x: 50, y });
            y -= 20;
        });

        const pdfBytes = await pdfDoc.save();
        fs.writeFileSync("uploads/output.pdf", pdfBytes);
        res.download("uploads/output.pdf");
    } catch (err) {
        res.status(500).json({ error: err.message });
    }
});

app.listen(PORT, () => console.log(`Server running at http://localhost:${PORT}`));
