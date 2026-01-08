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
            font-size: 11px;
            background: white;
          }
          
          /* Remove all borders by default */
          td, th {
            border: none;
            padding: 4px 8px;
            text-align: left;
            vertical-align: middle;
            background-color: white;
          }
          
          /* Header row - navy background */
          .header-row th {
            background-color: #1F4E79;
            color: white;
            font-weight: bold;
            text-align: center;
            border-bottom: 1px solid black;
          }
          
          /* Section headers - italic and underlined */
          .section-header {
            font-style: italic;
            text-decoration: underline;
            background: white !important;
            font-weight: normal;
          }
          
          /* Total row - top border only */
          .total-row td {
            font-weight: bold;
            border-top: 1px solid black;
          }
          
          /* Dollar sign column */
          .dollar {
            text-align: right;
            width: 20px;
            padding-right: 2px;
          }
          
          /* Amount columns */
          .amount {
            text-align: right;
            padding-left: 2px;
            white-space: nowrap;
          }
          
          /* Percentage column - blue italic */
          .percent {
            text-align: right;
            color: #0000FF;
            font-style: italic;
          }
          
          /* Label columns */
          .label {
            text-align: left;
          }
          
          /* Footnotes */
          .footnotes {
            margin-top: 8px;
            font-size: 10px;
          }
          
          .footnotes p {
            margin: 2px 0;
          }
        </style>
      </head>
      <body>
        <table>
    `;

    // Build table rows (reusing logic from /convert)
    for (let i = 0; i < tableData.length; i++) {
      const rowData = tableData[i].data;
      
      // Skip empty rows
      if (rowData.every(cell => !cell.value || cell.value.trim() === '')) continue;
      
      // Get row values
      const rowValues = rowData.map(cell => String(cell.value || '').trim());
      const rowTextLower = rowValues.join(' ').toLowerCase();
      const firstCellLower = rowValues[0].toLowerCase();
      
      // Determine row type
      let isHeaderRow = firstCellLower.includes('sources') && rowTextLower.includes('uses');
      let isTotalRow = rowTextLower.includes('total');
      let isSectionHeader = false;
      
      // Check for section headers
      for (const val of rowValues) {
        const lower = val.toLowerCase();
        if ((lower.includes('accretive') && lower.includes('costs')) ||
            (lower.includes('financing') && lower.includes('costs')) ||
            (lower.includes('closing') && lower.includes('costs'))) {
          isSectionHeader = true;
          break;
        }
      }
      
      // Build row
      if (isHeaderRow) {
        // Header row
        html += '<tr class="header-row">';
        let usesIndex = -1;
        for (let j = 0; j < rowValues.length; j++) {
          if (rowValues[j].toLowerCase().includes('uses')) {
            usesIndex = j;
            break;
          }
        }
        if (usesIndex === -1) usesIndex = Math.floor(rowValues.length / 2);
        
        html += `<th colspan="${usesIndex * 3}">Sources</th>`;
        html += `<th colspan="${(rowValues.length - usesIndex) * 3}">Uses</th>`;
        html += '</tr>';
      } else {
        // Data row
        html += `<tr${isTotalRow ? ' class="total-row"' : ''}>`;
        
        // Process each cell
        for (let j = 0; j < rowValues.length; j++) {
          const cellValue = rowValues[j];
          
          if (!cellValue) {
            html += '<td></td>';
            continue;
          }
          
          // Extract footnotes
          const footnoteMatch = cellValue.match(/^(.+?)(\*+)$/);
          let baseValue = cellValue;
          let footnote = '';
          if (footnoteMatch) {
            baseValue = footnoteMatch[1].trim();
            footnote = footnoteMatch[2];
          }
          
          // Check if section header cell
          if (isSectionHeader) {
            const lower = cellValue.toLowerCase();
            if ((lower.includes('accretive') && lower.includes('costs')) ||
                (lower.includes('financing') && lower.includes('costs')) ||
                (lower.includes('closing') && lower.includes('costs'))) {
              html += `<td colspan="2" class="section-header">${cellValue}</td>`;
              continue;
            }
          }
          
          // Check cell content type
          const currencyMatch = baseValue.match(/^\$?\s*([\d,]+(?:\.\d+)?)$/);
          const percentMatch = baseValue.match(/^([\d.]+)\s*%$/);
          
          if (currencyMatch && !isSectionHeader) {
            // Currency value
            const numValue = parseFloat(currencyMatch[1].replace(/,/g, ''));
            html += '<td class="dollar">$</td>';
            html += `<td class="amount">${numValue.toLocaleString('en-US', {
              minimumFractionDigits: 0,
              maximumFractionDigits: 0
            })}${footnote}</td>`;
          } else if (percentMatch) {
            // Percentage value  
            const numValue = parseFloat(percentMatch[1]);
            html += '<td></td>'; // Empty cell for dollar sign column
            html += `<td class="percent">${numValue.toFixed(2)}%</td>`;
          } else {
            // Text label
            html += `<td class="label" colspan="2">${cellValue}</td>`;
          }
        }
        
        html += '</tr>';
      }
    }

    html += `
        </table>
    `;
    
    // Check for footnotes in the data
    let hasSingleAsterisk = false;
    let hasDoubleAsterisk = false;
    let interestReserveMonths = '';
    
    for (const row of tableData) {
      for (const cell of row.data) {
        const value = String(cell.value || '');
        if (value.includes('**') && value.toLowerCase().includes('interest reserve')) {
          hasDoubleAsterisk = true;
          // Try to extract month count
          const monthMatch = value.match(/(\d+)\s*month/i);
          if (monthMatch) {
            interestReserveMonths = monthMatch[1];
          }
        } else if (value.includes('*') && !value.includes('**')) {
          hasSingleAsterisk = true;
        }
      }
    }
    
    // Add footnotes section if asterisks were found
    if (hasSingleAsterisk || hasDoubleAsterisk) {
      html += '<div class="footnotes">';
      if (hasSingleAsterisk) {
        html += '<p>*Estimated</p>';
      }
      if (hasDoubleAsterisk) {
        const months = interestReserveMonths || '24';
        html += `<p>**${months} month Interest Reserve</p>`;
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

    // Send response
    const base64Image = Buffer.from(image).toString('base64');
    const outputFilename = filename || `${tableName.replace(/\s+/g, '_').toLowerCase()}.png`;
    
    res.json({
      success: true,
      filename: outputFilename,
      image: base64Image,
      mimeType: 'image/png',
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