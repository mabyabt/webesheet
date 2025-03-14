const express = require("express");
const multer = require("multer");
const ExcelJS = require("exceljs");
const fs = require("fs");
const path = require("path");
const { PDFDocument, StandardFonts } = require("pdf-lib");

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

// Helper function to extract cell value properly
function extractCellValue(cell) {
  if (!cell) return "";
  
  // Handle formula result
  if (cell.formula) {
    return cell.result || "";
  }
  
  // Handle different value types
  if (cell.value === null || cell.value === undefined) {
    return "";
  }
  
  // Handle rich text
  if (typeof cell.value === 'object' && cell.value.richText) {
    return cell.value.richText.map(rt => rt.text).join("") || "";
  }
  
  // Handle other object types
  if (typeof cell.value === 'object') {
    // Try to get a string representation
    return cell.text || 
           (cell.value.toString && cell.value.toString() !== '[object Object]' ? 
            cell.value.toString() : JSON.stringify(cell.value));
  }
  
  // Regular value
  return cell.value;
}

// Debug endpoint to check what worksheets are available
app.get("/debug-excel", async (req, res) => {
  try {
    const filePath = path.join(__dirname, "uploads", "data.xlsx");
    
    if (!fs.existsSync(filePath)) {
      return res.status(404).json({ error: "Excel file not found. Please upload a file first." });
    }
    
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(filePath);
    
    const worksheetInfo = workbook.worksheets.map(sheet => {
      // Get merged cell info
      const merges = sheet.mergeCells._merges ? 
        Object.keys(sheet._merges).map(key => {
          const range = sheet._merges[key];
          return {
            range: key,
            top: range.top,
            left: range.left,
            bottom: range.bottom,
            right: range.right
          };
        }) : [];
      
      return {
        id: sheet.id,
        name: sheet.name,
        rowCount: sheet.rowCount,
        columnCount: sheet.columnCount,
        mergedCells: merges,
        hasMergedCells: merges.length > 0
      };
    });
    
    // Sample first few cells for debugging
    const sampleCells = [];
    if (workbook.worksheets.length > 0) {
      const sheet = workbook.worksheets[0];
      sheet.eachRow({ includeEmpty: false }, (row, rowNumber) => {
        if (rowNumber <= 5) { // Sample first 5 rows
          const cells = [];
          row.eachCell({ includeEmpty: false }, (cell, colNumber) => {
            cells.push({
              address: cell.address,
              type: cell.type,
              value: cell.value,
              formula: cell.formula,
              result: cell.result,
              text: extractCellValue(cell),
              isMerged: cell.isMerged
            });
          });
          sampleCells.push(cells);
        }
      });
    }
    
    res.json({
      worksheetCount: workbook.worksheets.length,
      worksheets: worksheetInfo,
      sampleCells
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
    
    // Get merged cell information
    const mergedCells = worksheet.mergeCells._merges ? 
      Object.keys(worksheet._merges).map(key => {
        const range = worksheet._merges[key];
        return {
          range: key,
          top: range.top,
          left: range.left,
          bottom: range.bottom,
          right: range.right
        };
      }) : [];
    
    console.log(`Found ${mergedCells.length} merged cell ranges`);
    
    // Create a matrix to store cell values
    const maxRow = worksheet.rowCount;
    const maxCol = worksheet.columnCount || 10; // Fallback if columnCount is not reliable
    
    // Initialize the data matrix with empty values
    const dataMatrix = Array(maxRow).fill().map(() => Array(maxCol).fill(""));
    
    // Fill in values from the worksheet
    worksheet.eachRow({ includeEmpty: true }, (row, rowIndex) => {
      row.eachCell({ includeEmpty: true }, (cell, colIndex) => {
        dataMatrix[rowIndex-1][colIndex-1] = extractCellValue(cell);
      });
    });
    
    // Handle merged cells - copy the value to all cells in the merged range
    mergedCells.forEach(merge => {
      const topValue = dataMatrix[merge.top-1][merge.left-1];
      
      // Copy the top-left value to all cells in the merged range
      for (let r = merge.top; r <= merge.bottom; r++) {
        for (let c = merge.left; c <= merge.right; c++) {
          dataMatrix[r-1][c-1] = topValue;
        }
      }
    });
    
    // Only include rows that have at least some data
    const finalData = dataMatrix.filter(row => row.some(cell => cell !== ""));
    
    // Handle empty spreadsheet case
    if (finalData.length === 0) {
      // Return a default template with one empty row
      return res.json([["", "", "", ""]]);
    }
    
    console.log(`Processed ${finalData.length} data rows`);
    
    // Return the data along with merged cell information for the frontend
    res.json({
      data: finalData,
      mergedCells: mergedCells
    });
  } catch (err) {
    console.error("Data fetch error:", err);
    res.status(500).json({ error: err.message });
  }
});

// Save edited data back to Excel
app.post("/save", async (req, res) => {
  try {
    const filePath = path.join(__dirname, "uploads", "data.xlsx");
    
    if (!req.body || !req.body.data) {
      return res.status(400).json({ error: "Invalid data format. Expected {data: [...]}" });
    }
    
    const dataToSave = req.body.data;
    
    if (!Array.isArray(dataToSave)) {
      return res.status(400).json({ error: "Invalid data format. Expected data to be an array." });
    }
    
    // Filter out completely empty rows
    const filteredData = dataToSave.filter(row => 
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
    
    // Preserve merged cells
    const existingMerges = worksheet.mergeCells && worksheet._merges ? 
      Object.keys(worksheet._merges) : [];
    
    // Clear existing content but remember merged cells
    worksheet.eachRow((row, rowNumber) => {
      worksheet.spliceRows(rowNumber, 1);
    });
    
    // Add new data
    filteredData.forEach((rowData, rowIndex) => {
      const row = worksheet.getRow(rowIndex + 1);
      rowData.forEach((cellValue, colIndex) => {
        const cell = row.getCell(colIndex + 1);
        cell.value = cellValue;
      });
      row.commit();
    });
    
    // Restore merged cells if provided in request
    if (req.body.mergedCells && Array.isArray(req.body.mergedCells)) {
      req.body.mergedCells.forEach(merge => {
        const range = `${String.fromCharCode(64 + merge.left)}${merge.top}:${String.fromCharCode(64 + merge.right)}${merge.bottom}`;
        worksheet.mergeCells(range);
      });
    }
    
    await workbook.xlsx.writeFile(filePath);
    res.json({ 
      message: "Saved successfully", 
      rowCount: filteredData.length 
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
    
    // Embed a Unicode font that can handle special characters
    const fontBytes = fs.readFileSync(path.join(__dirname, 'node_modules', 'pdf-lib', 'assets', 'fonts', 'times-roman.ttf'));
    const customFont = await pdfDoc.embedFont(StandardFonts.TimesRoman);
    
    let page = pdfDoc.addPage([600, 800]);
    
    const titleText = `Excel Data Export - ${worksheet.name}`;
    page.drawText(titleText, { 
      x: 50, 
      y: 780, 
      size: 16,
      font: customFont
    });
    
    page.drawText(`Generated on: ${new Date().toLocaleDateString()}`, {
      x: 50,
      y: 760,
      size: 10,
      font: customFont
    });
    
    let y = 730;
    let rowCount = 0;
    const rowsPerPage = 30;
    
    // Get merged cell information
    const mergedCells = worksheet.mergeCells._merges ? 
      Object.keys(worksheet._merges).map(key => {
        const range = worksheet._merges[key];
        return {
          range: key,
          top: range.top,
          left: range.left,
          bottom: range.bottom,
          right: range.right
        };
      }) : [];
    
    // Create a matrix to store cell values
    const maxRow = worksheet.rowCount;
    const maxCol = worksheet.columnCount || 10;
    
    // Initialize the data matrix with empty values
    const dataMatrix = Array(maxRow).fill().map(() => Array(maxCol).fill(""));
    
    // Fill in values from the worksheet
    worksheet.eachRow({ includeEmpty: true }, (row, rowIndex) => {
      row.eachCell({ includeEmpty: true }, (cell, colIndex) => {
        dataMatrix[rowIndex-1][colIndex-1] = extractCellValue(cell);
      });
    });
    
    // Handle merged cells
    mergedCells.forEach(merge => {
      const topValue = dataMatrix[merge.top-1][merge.left-1];
      
      // Copy the top-left value to all cells in the merged range
      for (let r = merge.top; r <= merge.bottom; r++) {
        for (let c = merge.left; c <= merge.right; c++) {
          dataMatrix[r-1][c-1] = topValue;
        }
      }
    });
    
    // Only include rows that have at least some data
    const finalData = dataMatrix.filter(row => row.some(cell => cell !== ""));
    
    if (finalData.length === 0) {
      page.drawText("No data found in spreadsheet", { 
        x: 50, 
        y, 
        size: 12,
        font: customFont
      });
    } else {
      // Calculate column widths based on content
      const maxColumnWidth = 100;
      const padding = 10;
      let columnWidths = [];
      
      // Initialize with header widths if available
      if (finalData.length > 0) {
        const firstRow = finalData[0];
        columnWidths = firstRow.map((header, index) => {
          // Check all values in this column to determine appropriate width
          const maxContentLength = Math.max(// Check all values in this column to determine appropriate width
            ...finalData.map(row => {
              const cellContent = String(row[index] || "");
              return cellContent.length;
            })
          );
          return Math.min(maxColumnWidth, (maxContentLength * 6) + padding);
        });
      }
      
      // Draw table headers
      let x = 50;
      finalData[0].forEach((header, colIndex) => {
        // Clean and sanitize the text for PDF
        const headerText = String(header || "")
          .replace(/[^\x00-\x7F]/g, " ") // Replace non-ASCII with space
          .substring(0, 15); // Limit length
        
        if (headerText.trim()) {
          try {
            page.drawText(headerText, { 
              x,
              y,
              size: 11,
              font: customFont
            });
          } catch (err) {
            console.warn(`Could not draw text "${headerText}": ${err.message}`);
          }
        }
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
      for (let i = 1; i < finalData.length; i++) {
        // Check if we need a new page
        if ((i - 1) % rowsPerPage === 0 && i > 1) {
          page = pdfDoc.addPage([600, 800]);
          y = 750;
          
          // Add continued header
          page.drawText(`Excel Data Export (continued) - Page ${Math.floor((i-1)/rowsPerPage) + 1}`, {
            x: 50,
            y: 780,
            size: 14,
            font: customFont
          });
          
          y = 730;
        }
        
        // Draw row
        x = 50;
        finalData[i].forEach((cell, colIndex) => {
          // Clean and sanitize the text for PDF
          const cellText = String(cell || "")
            .replace(/[^\x00-\x7F]/g, " ") // Replace non-ASCII with space
            .substring(0, 20); // Limit length
          
          if (cellText.trim()) {
            try {
              page.drawText(cellText, { 
                x, 
                y,
                size: 10,
                font: customFont
              });
            } catch (err) {
              console.warn(`Could not draw text "${cellText}": ${err.message}`);
            }
          }
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
            
