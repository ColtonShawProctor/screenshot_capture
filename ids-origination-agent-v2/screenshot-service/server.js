const express = require('express');
const ExcelJS = require('exceljs');
const nodeHtmlToImage = require('node-html-to-image');
const { detectTable } = require('./tableDetector');
const ExcelVisualRenderer = require('./excelVisualRenderer');

const app = express();
const PORT = process.env.PORT || 3000;

// Initialize Excel visual renderer
const visualRenderer = new ExcelVisualRenderer();

// Middleware
app.use(express.json({ limit: '50mb' }));
app.use(express.urlencoded({ extended: true, limit: '50mb' }));

// Health check endpoint
app.get('/health', async (req, res) => {
  const rendererHealth = await visualRenderer.healthCheck();
  res.status(200).json({ 
    status: 'healthy', 
    service: 'excel-screenshot',
    visualRenderer: rendererHealth
  });
});

// Convert Excel range to PNG
app.post('/convert', async (req, res) => {
  try {
    const { excelBase64, sheetName, range, filename } = req.body;

    if (!excelBase64 || !sheetName || !range) {
      return res.status(400).json({ error: 'Missing required parameters: excelBase64, sheetName, range' });
    }

    // Validate base64 format
    const base64Pattern = /^[A-Za-z0-9+/]*={0,2}$/;
    if (!base64Pattern.test(excelBase64)) {
      return res.status(400).json({
        error: 'Invalid excelBase64 format. Must be valid base64-encoded Excel file data.',
        received: typeof excelBase64 === 'string' ? excelBase64.substring(0, 50) + '...' : typeof excelBase64
      });
    }

    // Convert base64 to buffer
    let buffer;
    try {
      buffer = Buffer.from(excelBase64, 'base64');
      
      // Basic check for Excel file signature (ZIP header)
      if (buffer.length < 4 || buffer.readUInt32LE(0) !== 0x04034b50) {
        return res.status(400).json({
          error: 'Invalid Excel file. The provided base64 data does not appear to be a valid Excel file.',
          hint: 'Make sure you are sending the actual base64-encoded content of an Excel file (.xlsx)'
        });
      }
    } catch (err) {
      return res.status(400).json({
        error: 'Failed to decode base64 data',
        details: err.message
      });
    }

    // Load workbook
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.load(buffer);

    // Find sheet (case-insensitive)
    let worksheet = null;
    workbook.eachSheet((sheet) => {
      if (sheet.name.toLowerCase() === sheetName.toLowerCase()) {
        worksheet = sheet;
      }
    });

    if (!worksheet) {
      // Try partial match if exact match fails
      workbook.eachSheet((sheet) => {
        if (sheet.name.toLowerCase().includes(sheetName.toLowerCase())) {
          worksheet = sheet;
        }
      });
    }

    if (!worksheet) {
      return res.status(400).json({ error: `Sheet "${sheetName}" not found in workbook` });
    }

    // Validate range format (e.g., "A1:H30")
    const rangeMatch = range.match(/([A-Z]+)(\d+):([A-Z]+)(\d+)/);
    if (!rangeMatch) {
      return res.status(400).json({ error: 'Invalid range format. Use format like "A1:H30"' });
    }

    console.log(`Rendering Excel range ${sheetName}!${range} using LibreOffice visual renderer...`);
    
    // Use visual renderer to preserve actual Excel formatting
    const imageBuffer = await visualRenderer.renderExcelRange(buffer, sheetName, range, filename);
    const base64Image = imageBuffer.toString('base64');
    
    res.json({
      success: true,
      filename: filename || 'screenshot.png',
      image: base64Image,
      mimeType: 'image/png',
      method: 'libreoffice-visual-rendering',
      preservedFormatting: true
    });

  } catch (error) {
    console.error('Screenshot error:', error);
    res.status(500).json({ 
      error: 'Failed to generate screenshot',
      details: error.message 
    });
  }
});

// Detect table and capture screenshot
app.post('/detect-and-capture', async (req, res) => {
  try {
    const { excelBase64, tableName, searchSheets, padding = 2, filename } = req.body;

    if (!excelBase64 || !tableName) {
      return res.status(400).json({ 
        error: 'Missing required parameters: excelBase64, tableName' 
      });
    }

    // Validate base64 format
    const base64Pattern = /^[A-Za-z0-9+/]*={0,2}$/;
    if (!base64Pattern.test(excelBase64)) {
      return res.status(400).json({
        error: 'Invalid excelBase64 format. Must be valid base64-encoded Excel file data.',
        received: typeof excelBase64 === 'string' ? excelBase64.substring(0, 50) + '...' : typeof excelBase64
      });
    }

    // Convert base64 to buffer
    let buffer;
    try {
      buffer = Buffer.from(excelBase64, 'base64');
      
      // Basic check for Excel file signature (ZIP header)
      if (buffer.length < 4 || buffer.readUInt32LE(0) !== 0x04034b50) {
        return res.status(400).json({
          error: 'Invalid Excel file. The provided base64 data does not appear to be a valid Excel file.',
          hint: 'Make sure you are sending the actual base64-encoded content of an Excel file (.xlsx)'
        });
      }
    } catch (err) {
      return res.status(400).json({
        error: 'Failed to decode base64 data',
        details: err.message
      });
    }

    // Load workbook
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.load(buffer);

    // Detect table
    const detection = await detectTable(workbook, tableName, searchSheets, padding);

    if (!detection.found) {
      return res.status(404).json({
        success: false,
        error: `Table '${tableName}' not found`,
        searchedSheets: detection.searchedSheets,
        suggestions: detection.suggestions
      });
    }

    // Generate screenshot using visual renderer (preserves Excel formatting)
    console.log(`Rendering detected table ${tableName} at ${detection.sheet}!${detection.range}...`);
    
    const outputFilename = filename || `${tableName.replace(/\s+/g, '_').toLowerCase()}.png`;
    const imageBuffer = await visualRenderer.renderExcelRange(buffer, detection.sheet, detection.range, outputFilename);
    const base64Image = imageBuffer.toString('base64');
    
    res.json({
      success: true,
      filename: outputFilename,
      image: base64Image,
      mimeType: 'image/png',
      method: 'libreoffice-visual-rendering',
      preservedFormatting: true,
      detected: {
        sheet: detection.sheet,
        range: detection.range,
        headerCell: detection.headerCell,
        confidence: detection.confidence
      }
    });

  } catch (error) {
    console.error('Table detection error:', error);
    res.status(500).json({ 
      error: 'Failed to detect and capture table',
      details: error.message 
    });
  }
});

// Helper function to convert column letter to number
function columnToNumber(column) {
  let result = 0;
  for (let i = 0; i < column.length; i++) {
    result = result * 26 + (column.charCodeAt(i) - 'A'.charCodeAt(0) + 1);
  }
  return result;
}

// Start server
app.listen(PORT, () => {
  console.log(`Excel screenshot service running on port ${PORT}`);
});

// Handle graceful shutdown
process.on('SIGTERM', () => {
  console.log('SIGTERM signal received: closing HTTP server');
  app.close(() => {
    console.log('HTTP server closed');
  });
});