const express = require('express');
const ExcelJS = require('exceljs');
const nodeHtmlToImage = require('node-html-to-image');

const app = express();
const PORT = process.env.PORT || 3000;

// Middleware
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

    // First pass: analyze table structure to identify section headers and totals
    const tableData = [];
    for (let rowNum = startRow; rowNum <= endRow; rowNum++) {
      const row = worksheet.getRow(rowNum);
      const rowData = [];
      
      for (let colNum = startCol; colNum <= endCol; colNum++) {
        const cell = row.getCell(colNum);
        const value = cell.value;
        let displayValue = '';
        
        if (value !== null && value !== undefined) {
          if (typeof value === 'object' && value.formula) {
            displayValue = value.result || '';
          } else {
            displayValue = value.toString();
          }
        }
        
        rowData.push({
          value: displayValue,
          cell: cell,
          isEmpty: !displayValue || displayValue.trim() === ''
        });
      }
      
      tableData.push({
        rowNum,
        data: rowData,
        row: row
      });
    }

    // Generate HTML table with Fairbridge styling
    let html = `
      <html>
      <head>
        <style>
          body {
            font-family: 'Calibri', Arial, sans-serif;
            margin: 0;
            padding: 10px;
            background-color: white;
          }
          table {
            border-collapse: collapse;
            width: auto;
            margin: 0;
            font-size: 10pt;
          }
          
          /* Header row styling */
          .header-row {
            background-color: #1F4E79 !important;
            color: white;
            font-weight: bold;
            font-size: 11pt;
          }
          
          /* Section header styling */
          .section-header {
            background-color: #4472C4 !important;
            color: white;
            font-weight: bold;
          }
          
          /* Data rows */
          .data-row-odd {
            background-color: white;
          }
          
          .data-row-even {
            background-color: #F2F2F2;
          }
          
          /* Total rows */
          .total-row {
            font-weight: bold;
            border-top: 2px solid #333 !important;
          }
          
          /* Cell styling */
          th, td {
            border: 1px solid #D0D0D0;
            padding: 6px 12px;
            text-align: left;
            vertical-align: middle;
          }
          
          /* Number alignment and formatting */
          .number, .currency, .percentage {
            text-align: right;
          }
          
          .bold {
            font-weight: bold;
          }
        </style>
      </head>
      <body>
        <table>
    `;

    // Build table from analyzed data
    let dataRowCounter = 0;
    
    for (let i = 0; i < tableData.length; i++) {
      const rowInfo = tableData[i];
      const rowNum = rowInfo.rowNum;
      const rowData = rowInfo.data;
      
      // Determine row type and styling
      let rowClass = '';
      let isHeader = false;
      let isSectionHeader = false;
      let isTotalRow = false;
      
      // First row is always header
      if (i === 0) {
        isHeader = true;
        rowClass = 'header-row';
      } else {
        // Check if this is a section header (first cell has text, rest are empty)
        const firstCellValue = rowData[0].value.trim();
        const restAreEmpty = rowData.slice(1).every(cell => cell.isEmpty);
        
        if (firstCellValue && restAreEmpty && 
            (firstCellValue.toLowerCase().includes('sources') || 
             firstCellValue.toLowerCase().includes('uses') || 
             firstCellValue.toLowerCase().includes('costs') || 
             firstCellValue.toLowerCase().includes('financing'))) {
          isSectionHeader = true;
          rowClass = 'section-header';
        }
        // Check if this is a total row
        else if (firstCellValue.toLowerCase().includes('total') || 
                 rowData.some(cell => cell.cell.style?.font?.bold)) {
          isTotalRow = true;
          rowClass = 'total-row';
        }
        // Regular data row with alternating colors
        else {
          const isEven = dataRowCounter % 2 === 1;
          rowClass = isEven ? 'data-row-even' : 'data-row-odd';
          dataRowCounter++;
        }
      }
      
      html += `<tr class="${rowClass}">`;
      
      for (let colNum = startCol; colNum <= endCol; colNum++) {
        const cellIndex = colNum - startCol;
        const cellInfo = rowData[cellIndex];
        const cell = cellInfo.cell;
        const style = cell.style || {};
        
        let cellClass = '';
        let displayValue = cellInfo.value;

        if (displayValue) {
          // Apply number formatting and alignment
          const numValue = parseFloat(displayValue.replace(/[^0-9.-]/g, ''));
          
          if (!isNaN(numValue)) {
            // Check if it's currency
            if (displayValue.includes('$') || cell.numFmt?.includes('$') || cell.numFmt?.includes('¤')) {
              cellClass += ' currency';
              if (numValue >= 1000000) {
                displayValue = '$' + (numValue / 1000000).toFixed(1) + 'M';
              } else if (numValue >= 1000) {
                displayValue = '$' + numValue.toLocaleString('en-US', { maximumFractionDigits: 0 });
              } else {
                displayValue = '$' + numValue.toLocaleString('en-US');
              }
            }
            // Check if it's percentage
            else if (displayValue.includes('%') || cell.numFmt?.includes('%')) {
              cellClass += ' percentage';
              displayValue = numValue.toFixed(1) + '%';
            }
            // Regular number
            else if (displayValue.match(/^\d+\.?\d*$/)) {
              cellClass += ' number';
              displayValue = numValue.toLocaleString('en-US');
            }
          }

          // Apply text styling
          if (style.font?.bold || isTotalRow) {
            cellClass += ' bold';
          }
        }

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

    // Convert HTML to image with tight cropping
    const image = await nodeHtmlToImage({
      html,
      quality: 100,
      type: 'png',
      puppeteerArgs: {
        executablePath: process.env.PUPPETEER_EXECUTABLE_PATH || '/usr/bin/chromium',
        args: ['--no-sandbox', '--disable-setuid-sandbox']
      },
      waitUntil: 'networkidle0',
      selector: 'table'
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