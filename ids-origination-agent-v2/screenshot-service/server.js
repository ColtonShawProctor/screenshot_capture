const express = require('express');
const ExcelJS = require('exceljs');
const nodeHtmlToImage = require('node-html-to-image');
const cors = require('cors');

const app = express();
const PORT = process.env.PORT || 3000;

// Middleware
app.use(cors());
app.use(express.json({ limit: '50mb' }));
app.use(express.urlencoded({ extended: true, limit: '50mb' }));

// Health check endpoint
app.get('/health', (req, res) => {
  res.status(200).json({ status: 'healthy', service: 'excel-screenshot' });
});

// Convert Excel range to PNG
app.post('/convert', async (req, res) => {
  try {
    const { excelBase64, sheetName, range, filename } = req.body;

    if (!excelBase64 || !sheetName || !range) {
      return res.status(400).json({ error: 'Missing required parameters: excelBase64, sheetName, range' });
    }

    // Convert base64 to buffer
    const buffer = Buffer.from(excelBase64, 'base64');

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

    // Parse range (e.g., "A1:H30")
    const rangeMatch = range.match(/([A-Z]+)(\d+):([A-Z]+)(\d+)/);
    if (!rangeMatch) {
      return res.status(400).json({ error: 'Invalid range format. Use format like "A1:H30"' });
    }

    const startCol = columnToNumber(rangeMatch[1]);
    const startRow = parseInt(rangeMatch[2]);
    const endCol = columnToNumber(rangeMatch[3]);
    const endRow = parseInt(rangeMatch[4]);

    // Generate HTML table
    let html = `
      <html>
      <head>
        <style>
          body {
            font-family: Arial, sans-serif;
            margin: 0;
            padding: 20px;
            background-color: white;
          }
          table {
            border-collapse: collapse;
            width: auto;
            margin: 0 auto;
          }
          th, td {
            border: 1px solid #ddd;
            padding: 8px 12px;
            text-align: left;
            font-size: 12px;
          }
          th {
            background-color: #f5f5f5;
            font-weight: bold;
          }
          td {
            background-color: white;
          }
          .number {
            text-align: right;
          }
          .currency {
            text-align: right;
          }
          .percentage {
            text-align: right;
          }
          .bold {
            font-weight: bold;
          }
          .italic {
            font-style: italic;
          }
          tr:nth-child(even) td {
            background-color: #fafafa;
          }
        </style>
      </head>
      <body>
        <table>
    `;

    // Build table from range
    for (let rowNum = startRow; rowNum <= endRow; rowNum++) {
      const row = worksheet.getRow(rowNum);
      html += '<tr>';
      
      for (let colNum = startCol; colNum <= endCol; colNum++) {
        const cell = row.getCell(colNum);
        const value = cell.value;
        const style = cell.style || {};
        
        let cellClass = '';
        let displayValue = '';

        if (value !== null && value !== undefined) {
          // Handle different value types
          if (typeof value === 'object' && value.formula) {
            displayValue = value.result || '';
          } else {
            displayValue = value.toString();
          }

          // Apply number formatting
          if (cell.numFmt) {
            if (cell.numFmt.includes('$') || cell.numFmt.includes('¤')) {
              cellClass += ' currency';
              if (typeof value === 'number') {
                displayValue = '$' + value.toLocaleString('en-US', { minimumFractionDigits: 0, maximumFractionDigits: 0 });
              }
            } else if (cell.numFmt.includes('%')) {
              cellClass += ' percentage';
              if (typeof value === 'number') {
                displayValue = (value * 100).toFixed(1) + '%';
              }
            } else if (typeof value === 'number') {
              cellClass += ' number';
              displayValue = value.toLocaleString('en-US');
            }
          }

          // Apply text styling
          if (style.font) {
            if (style.font.bold) cellClass += ' bold';
            if (style.font.italic) cellClass += ' italic';
          }
        }

        // Determine if this is a header cell
        const isHeader = rowNum === startRow || (style.font && style.font.bold);
        const cellTag = isHeader ? 'th' : 'td';
        
        html += `<${cellTag}${cellClass ? ' class="' + cellClass.trim() + '"' : ''}>${displayValue}</${cellTag}>`;
      }
      
      html += '</tr>';
    }

    html += `
        </table>
      </body>
      </html>
    `;

    // Convert HTML to image
    const image = await nodeHtmlToImage({
      html,
      quality: 100,
      type: 'png',
      puppeteerArgs: {
        executablePath: process.env.PUPPETEER_EXECUTABLE_PATH || '/usr/bin/chromium',
        args: ['--no-sandbox', '--disable-setuid-sandbox']
      }
    });

    // Send image as base64
    const base64Image = Buffer.from(image).toString('base64');
    
    res.json({
      success: true,
      filename: filename || 'screenshot.png',
      image: base64Image,
      mimeType: 'image/png'
    });

  } catch (error) {
    console.error('Screenshot error:', error);
    res.status(500).json({ 
      error: 'Failed to generate screenshot',
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