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
            displayValue = String(value.result || '');
          } else {
            displayValue = String(value);
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
            font-family: Arial, 'Calibri', sans-serif;
            margin: 0;
            padding: 10px;
            background-color: white;
          }
          
          table {
            border-collapse: collapse;
            width: auto;
            margin: 0;
            font-size: 11pt;
            border: 2px solid #999;
          }
          
          /* Header row styling - only for "Sources" and "Uses" row */
          .header-row th {
            background-color: #1F4E79 !important;
            color: white;
            font-weight: bold;
          }
          
          /* Section headers - italic and underlined */
          .section-header td {
            font-style: italic;
            text-decoration: underline;
            font-weight: normal;
            background-color: white !important;
          }
          
          /* Total rows */
          .total-row td {
            font-weight: bold;
            border-top: 2px solid #666 !important;
          }
          
          .total-row td:first-child {
            border-top: 2px solid #666 !important;
          }
          
          /* Cell styling */
          th, td {
            border: 1px solid #999;
            padding: 4px 8px;
            text-align: left;
            vertical-align: middle;
            background-color: white;
          }
          
          /* Dollar sign column */
          .dollar-sign {
            text-align: left;
            width: 20px;
            padding-right: 2px;
          }
          
          /* Amount columns */
          .amount {
            text-align: right;
            padding-left: 2px;
          }
          
          /* Percentage column */
          .percentage {
            text-align: right;
            color: #0000FF;
            font-style: italic;
          }
          
          /* Bold text */
          .bold {
            font-weight: bold;
          }
          
          /* Footnotes */
          .footnotes {
            margin-top: 10px;
            font-size: 10pt;
            font-family: Arial, sans-serif;
          }
          
          .footnotes p {
            margin: 2px 0;
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
      
      // Check for header row - "Sources and Uses" or just contains "Sources" in first cell
      const firstCellValue = String(rowData[0].value || '').trim().toLowerCase();
      
      if (firstCellValue.includes('sources') && (firstCellValue.includes('uses') || 
          rowData.some(cell => String(cell.value || '').toLowerCase().includes('uses')))) {
        isHeader = true;
        rowClass = 'header-row';
      } 
      // Check if this is a section header (Accretive Costs, Financing Costs, Closing Costs)
      else if ((firstCellValue.includes('accretive') && firstCellValue.includes('costs')) ||
               (firstCellValue.includes('financing') && firstCellValue.includes('costs')) ||
               (firstCellValue.includes('closing') && firstCellValue.includes('costs')) ||
               firstCellValue === 'costs' ||
               firstCellValue === 'accretive costs' ||
               firstCellValue === 'financing costs' ||
               firstCellValue === 'closing costs') {
        isSectionHeader = true;
        rowClass = 'section-header';
      }
      // Check if this is a total row
      else if (firstCellValue.includes('total') || 
               rowData.some(cell => cell.cell.style?.font?.bold)) {
        isTotalRow = true;
        rowClass = 'total-row';
      }
      // Regular data row - all white background
      else {
        rowClass = 'data-row';
      }
      
      html += `<tr class="${rowClass}">`;
      
      for (let colNum = startCol; colNum <= endCol; colNum++) {
        const cellIndex = colNum - startCol;
        const cellInfo = rowData[cellIndex];
        const cell = cellInfo.cell;
        const style = cell.style || {};
        
        let cellClass = '';
        let displayValue = String(cellInfo.value || '');

        if (displayValue) {
          // Check for footnote markers
          const footnoteMatch = displayValue.match(/^(.+?)(\*+)$/);
          let baseValue = displayValue;
          let footnoteMarker = '';
          
          if (footnoteMatch) {
            baseValue = footnoteMatch[1].trim();
            footnoteMarker = footnoteMatch[2];
          }
          
          // Apply number formatting and alignment
          const numValue = parseFloat(baseValue.replace(/[^0-9.-]/g, ''));
          
          if (!isNaN(numValue) && baseValue.match(/[\d,.$%]/)) {
            // Check if it's currency
            if (baseValue.includes('$') || cell.numFmt?.includes('$') || cell.numFmt?.includes('¤')) {
              // For currency, we'll split the dollar sign from the amount
              cellClass += ' amount';
              // Show full number, no abbreviation
              displayValue = '$' + numValue.toLocaleString('en-US', { minimumFractionDigits: 0, maximumFractionDigits: 0 });
            }
            // Check if it's percentage
            else if (baseValue.includes('%') || cell.numFmt?.includes('%')) {
              cellClass += ' percentage';
              displayValue = numValue.toFixed(1) + '%';
            }
            // Regular number
            else if (baseValue.match(/^\d+\.?\d*$/)) {
              cellClass += ' amount';
              displayValue = numValue.toLocaleString('en-US');
            }
            
            // Re-add footnote marker if present
            if (footnoteMarker) {
              displayValue += footnoteMarker;
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
    `;
    
    // Check if we need to add footnotes
    const footnotes = [];
    for (const row of tableData) {
      for (const cell of row.data) {
        if (cell.value && cell.value.includes('*')) {
          const match = cell.value.match(/\*+(.+?)$/);
          if (match && match[1]) {
            footnotes.push(match[1].trim());
          }
        }
      }
    }
    
    // Add common footnotes if they appear in the data
    if (tableData.some(row => row.data.some(cell => cell.value && cell.value.includes('*') && !cell.value.includes('**')))) {
      if (!footnotes.some(f => f.toLowerCase().includes('estimated'))) {
        footnotes.unshift('*Estimated');
      }
    }
    if (tableData.some(row => row.data.some(cell => cell.value && cell.value.includes('**')))) {
      if (!footnotes.some(f => f.toLowerCase().includes('interest reserve'))) {
        footnotes.push('**12 month Interest Reserve');
      }
    }
    
    // Add footnotes section if any exist
    if (footnotes.length > 0) {
      html += '<div class="footnotes">';
      for (const footnote of footnotes) {
        html += `<p>${footnote}</p>`;
      }
      html += '</div>';
    }
    
    html += `
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
      selector: 'body'
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