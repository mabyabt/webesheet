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

// Debug endpoint to check what worksheets are available
app.get("/debug-excel", async (req, res) => {
  try {
    const filePath = path.join(__dirname, "uploads", "data.xlsx");
    
    if (!fs.existsSync(filePath)) {
      return res.status(404).json({ error: "Excel file not found. Please upload a file first." });
    }
    
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(filePath);
    
    const worksheetInfo = workbook.worksheets.map(sheet => ({
      id: sheet.id,
      name: sheet.name,
      rowCount: sheet.rowCount,
      columnCount: sheet.columnCount
    }));
    
    res.json({
      worksheetCount: workbook.worksheets.length,
      worksheets: worksheetInfo
    });
  } catch (err) {
    console.error("Debug error:", err);
    res.status(500).json({ error: err.message });
  }
});

// File upload endpoint
app.post("/upload", upload.single("file"), (req, res) => {
  try {
    if (!req.file) {
      return res.status(400).json({ error: "No file uploaded" });
    }
    res.json({ 
      message: "File uploaded successfully",
      file: req.file.originalname
    });
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
    
    // Try to get first worksheet by index, name or create one if needed
    let worksheet = workbook.getWorksheet(1); // Try by index (1-based)
    
    if (!worksheet && workbook.worksheets.length > 0) {
      // If no worksheet at index 1, try getting the first available worksheet
      worksheet = workbook.worksheets[0];
    }
    
    if (!worksheet) {
      // If still no worksheet, the file might be empty
      return res.status(404).json({ 
        error: "No worksheets found in the Excel file. Please check the file content." 
      });
    }
    
    console.log(`Found worksheet: ${worksheet.name} with ${worksheet.rowCount} rows and ${worksheet.columnCount} columns`);
    
    const data = [];
    
    // Only process rows that have data
    worksheet.eachRow({ includeEmpty: false }, (row, rowNumber) => {
      // Filter out undefined and null values, but keep numeric zeros
      const rowValues = row.values.slice(1).map(value => {
        if (value && typeof value === 'object') {
          // Handle Excel cell objects like Rich Text
          return value.text || value.toString() || "";
        }
        return value === undefined || value === null ? "" : value;
      });
      
      // Only add rows that have at least one non-empty value
      if (rowValues.some(val => val !== "")) {
        data.push(rowValues);
      }
    });
    
    console.log(`Processed ${data.length} data rows`);
    
    // Handle empty spreadsheet case
    if (data.length === 0) {
      // Return a default template with one empty row
      return res.json([["", "", "", ""]]);
    }
    
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
    
    if (!Array.isArray(req.body)) {
      return res.status(400).json({ error: "Invalid data format. Expected an array." });
    }
    
    // Filter out empty rows
    const dataToSave = req.body.filter(row => 
      Array.isArray(row) && row.some(cell => cell !== null && cell !== "")
    );
    
    // Create a new workbook if file doesn't exist
    const workbook = new ExcelJS.Workbook();
    
    if (fs.existsSync(filePath)) {
      await workbook.xlsx.readFile(filePath);
    }
    
    // Get the first worksheet or create a new one
    let worksheet = workbook.getWorksheet(1);
    if (!worksheet) {
      if (workbook.worksheets.length > 0) {
        worksheet = workbook.worksheets[0];
      } else {
        worksheet = workbook.addWorksheet('Sheet1');
      }
    }
    
    // Clear existing content
    worksheet.eachRow((row, rowNumber) => {
      worksheet.spliceRows(rowNumber, 1);
    });
    
    // Add new data
    dataToSave.forEach((rowData, index) => {
      const row = worksheet.getRow(index + 1);
      row.values = [null, ...rowData]; // Add null at beginning for 1-based indexing
      row.commit();
    });
    
    await workbook.xlsx.writeFile(filePath);
    res.json({ 
      message: "Saved successfully", 
      rowCount: dataToSave.length 
    });
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
    
    // Try multiple approaches to get a valid worksheet
    let worksheet = workbook.getWorksheet(1);
    
    if (!worksheet && workbook.worksheets.length > 0) {
      // If no worksheet at index 1, try getting the first available worksheet
      worksheet = workbook.worksheets[0];
    }
    
    if (!worksheet) {
      return res.status(404).json({ 
        error: "No worksheets found in the Excel file. Please check the file content."
      });
    }
    
    const pdfDoc = await PDFDocument.create();
    let page = pdfDoc.addPage([600, 800]);
    
    const titleText = `Excel Data Export - ${worksheet.name}`;
    page.drawText(titleText, { 
      x: 50, 
      y: 780, 
      size: 16 
    });
    
    page.drawText(`Generated on: ${new Date().toLocaleDateString()}`, {
      x: 50,
      y: 760,
      size: 10
    });
    
    let y = 730;
    let rowCount = 0;
    const rowsPerPage = 30;
    
    // Extract data, handling empty rows properly
    const excelData = [];
    worksheet.eachRow({ includeEmpty: false }, (row, rowNumber) => {
      // Process cell values to handle objects and formatting
      const rowValues = row.values.slice(1).map(value => {
        if (value && typeof value === 'object') {
          return value.text || value.toString() || "";
        }
        return value === undefined || value === null ? "" : String(value);
      });
      
      if (rowValues.some(val => val !== "")) {
        excelData.push(rowValues);
      }
    });
    
    if (excelData.length === 0) {
      page.drawText("No data found in spreadsheet", { x: 50, y, size: 12 });
    } else {
      // Calculate column widths based on content
      const maxColumnWidth = 100;
      const padding = 10;
      let columnWidths = [];
      
      // Initialize with header widths
      if (excelData.length > 0) {
        columnWidths = excelData[0].map(header => 
          Math.min(maxColumnWidth, (String(header).length * 7) + padding)
        );
      }
      
      // Add table headers
      let x = 50;
      excelData[0].forEach((header, colIndex) => {
        page.drawText(String(header).substring(0, 15), { 
          x,
          y,
          size: 11
        });
        x += columnWidths[colIndex] || 100;
      });
      
      y -= 20;
      
      // Horizontal line after headers
      page.drawLine({
        start: { x: 50, y: y + 5 },
        end: { x: 550, y: y + 5 },
        thickness: 1
      });
      
      y -= 10;
      
      // Draw data rows
      for (let i = 1; i < excelData.length; i++) {
        // Check if we need a new page
        if ((i - 1) % rowsPerPage === 0 && i > 1) {
          page = pdfDoc.addPage([600, 800]);
          y = 750;
          
          // Add continued header
          page.drawText(`Excel Data Export (continued) - Page ${Math.floor((i-1)/rowsPerPage) + 1}`, {
            x: 50,
            y: 780,
            size: 14
          });
          
          y = 750;
        }
        
        // Draw row
        x = 50;
        excelData[i].forEach((cell, colIndex) => {
          const cellText = String(cell).substring(0, 20);
          page.drawText(cellText, { 
            x, 
            y,
            size: 10
          });
          x += columnWidths[colIndex] || 100;
        });
        
        y -= 15;
      }
    }
    
    const pdfPath = path.join(__dirname, "uploads", "output.pdf");
    const pdfBytes = await pdfDoc.save();
    fs.writeFileSync(pdfPath, pdfBytes);
    
    res.download(pdfPath, "excel_export.pdf");
  } catch (err) {
    console.error("PDF generation error:", err.stack);
    res.status(500).json({ error: err.message });
  }
});

// Status endpoint to check if server is running
app.get("/status", (req, res) => {
  const uploadDir = path.join(__dirname, "uploads");
  const files = fs.existsSync(uploadDir) ? fs.readdirSync(uploadDir) : [];
  
  res.json({ 
    status: "Server running",
    uploadedFiles: files,
    uploadsFolderExists: fs.existsSync(uploadDir)
  });
});

// Start the server
app.listen(PORT, () => {
  console.log(`Server running at http://localhost:${PORT}`);
  console.log(`- Upload Excel files: POST /upload`);
  console.log(`- Get Excel data: GET /data`);
  console.log(`- Save Excel data: POST /save`);
  console.log(`- Export as PDF: GET /export-pdf`);
  console.log(`- Debug Excel file: GET /debug-excel`);
});
