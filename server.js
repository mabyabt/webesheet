const express = require("express");
const multer = require("multer");
const ExcelJS = require("exceljs");
const fs = require("fs");
const path = require("path");
const { PDFDocument } = require("pdf-lib");

const app = express();
const PORT = 3000;

// Create uploads directory if it doesn't exist
if (!fs.existsSync("uploads")) {
  fs.mkdirSync("uploads");
}

// Multer setup with storage configuration
const storage = multer.diskStorage({
  destination: function(req, file, cb) {
    cb(null, "uploads/");
  },
  filename: function(req, file, cb) {
    cb(null, "data.xlsx");
  }
});

const upload = multer({ storage });

// Serve static files from public directory
app.use(express.static("public"));
app.use(express.json());

// File upload endpoint
app.post("/upload", upload.single("file"), (req, res) => {
  try {
    if (!req.file) {
      return res.status(400).json({ error: "No file uploaded" });
    }
    res.json({ message: "File uploaded successfully" });
  } catch (err) {
    console.error("Upload error:", err);
    res.status(500).json({ error: err.message });
  }
});

// Load Excel file and send data to frontend
app.get("/data", async (req, res) => {
  try {
    const filePath = path.join(__dirname, "uploads", "data.xlsx");
    
    // Check if file exists
    if (!fs.existsSync(filePath)) {
      return res.status(404).json({ 
        error: "Excel file not found. Please upload a file first." 
      });
    }
    
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(filePath);
    
    // Check if worksheet exists
    const worksheet = workbook.getWorksheet(1);
    if (!worksheet) {
      return res.status(404).json({ error: "Worksheet not found in the Excel file." });
    }
    
    const data = [];
    worksheet.eachRow((row, rowNumber) => {
      // Include row values and skip the first undefined element
      data.push(row.values.slice(1));
    });
    
    res.json(data);
  } catch (err) {
    console.error("Data fetch error:", err);
    res.status(500).json({ error: err.message });
  }
});

// Save edited data back to Excel
app.post("/save", async (req, res) => {
  try {
    const filePath = path.join(__dirname, "uploads", "data.xlsx");
    
    // Check if file exists
    if (!fs.existsSync(filePath)) {
      return res.status(404).json({ error: "Excel file not found" });
    }
    
    if (!Array.isArray(req.body)) {
      return res.status(400).json({ error: "Invalid data format. Expected an array." });
    }
    
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(filePath);
    
    const worksheet = workbook.getWorksheet(1);
    if (!worksheet) {
      return res.status(404).json({ error: "Worksheet not found" });
    }
    
    // Clear existing rows
    worksheet.eachRow((row, rowNumber) => {
      if (rowNumber <= req.body.length) {
        // Keep these rows, they will be overwritten
      } else {
        // Remove extra rows
        worksheet.spliceRows(rowNumber, 1);
      }
    });
    
    // Update with new data
    req.body.forEach((row, rowIndex) => {
      const wsRow = worksheet.getRow(rowIndex + 1);
      wsRow.values = [null, ...row]; // Add null at beginning for ExcelJS 1-based indexing
      wsRow.commit();
    });
    
    await workbook.xlsx.writeFile(filePath);
    res.json({ message: "Saved successfully" });
  } catch (err) {
    console.error("Save error:", err);
    res.status(500).json({ error: err.message });
  }
});

// Generate PDF from Excel data
app.get("/export-pdf", async (req, res) => {
  try {
    const filePath = path.join(__dirname, "uploads", "data.xlsx");
    
    // Check if file exists
    if (!fs.existsSync(filePath)) {
      return res.status(404).json({ error: "Excel file not found. Please upload a file first." });
    }
    
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(filePath);
    
    // Check if worksheet exists
    const worksheet = workbook.getWorksheet(1);
    if (!worksheet) {
      return res.status(404).json({ error: "Worksheet not found in the Excel file." });
    }
    
    const pdfDoc = await PDFDocument.create();
    const page = pdfDoc.addPage([600, 800]);
    
    // Add title
    page.drawText("Excel Data Export", { 
      x: 50, 
      y: 780, 
      size: 18 
    });
    
    let y = 750;
    let rowCount = 0;
    
    // Add column headers with bold formatting (if available in first row)
    const headers = worksheet.getRow(1).values.slice(1);
    if (headers.length > 0) {
      page.drawText(headers.join(" | "), { 
        x: 50, 
        y, 
        size: 12
      });
      y -= 30; // Extra space after headers
    }
    
    // Add data rows
    worksheet.eachRow((row, rowNumber) => {
      if (rowNumber > 1) { // Skip header row which we already processed
        const rowData = row.values.slice(1);
        if (rowData.length > 0) {
          // Convert any special objects to strings
          const rowText = rowData.map(cell => {
            if (cell && typeof cell === 'object' && cell.text) {
              return cell.text;
            }
            return cell;
          }).join(" | ");
          
          page.drawText(rowText, { x: 50, y, size: 10 });
          y -= 20;
          rowCount++;
          
          // Add a new page if we're running out of space
          if (rowCount % 30 === 0 && rowCount > 0) {
            page = pdfDoc.addPage([600, 800]);
            y = 750;
          }
        }
      }
    });
    
    const pdfPath = path.join(__dirname, "uploads", "output.pdf");
    const pdfBytes = await pdfDoc.save();
    fs.writeFileSync(pdfPath, pdfBytes);
    
    res.download(pdfPath, "excel_export.pdf");
  } catch (err) {
    console.error("PDF generation error:", err);
    res.status(500).json({ error: err.message });
  }
});

// Status endpoint to check if server is running
app.get("/status", (req, res) => {
  res.json({ status: "Server running" });
});

// Start the server
app.listen(PORT, () => {
  console.log(`Server running at http://localhost:${PORT}`);
  console.log(`- Upload Excel files: POST /upload`);
  console.log(`- Get Excel data: GET /data`);
  console.log(`- Save Excel data: POST /save`);
  console.log(`- Export as PDF: GET /export-pdf`);
});
