const ExcelJS = require('exceljs');

// Table names and their common variations
const TABLE_MAPPINGS = {
  'Sources and Uses': ['sources and uses', 'sources & uses', 'source and use', 'source & use'],
  'Take Out Loan Sizing': ['take out loan sizing', 'takeout loan sizing', 'take-out loan sizing', 'loan sizing'],
  'Capital Stack at Closing': ['capital stack at closing', 'capital stack', 'cap stack at closing', 'cap stack'],
  'Loan to Cost': ['loan to cost', 'ltc', 'loan-to-cost', 'l2c'],
  'Loan to Value': ['loan to value', 'ltv', 'loan-to-value', 'l2v'],
  'PILOT Schedule': ['pilot schedule', 'pilot', 'payment in lieu of taxes'],
  'Occupancy': ['occupancy', 'unit mix', 'unit occupancy', 'occupancy schedule']
};

// Fuzzy string matching
function fuzzyMatch(str1, str2, threshold = 0.8) {
  const s1 = str1.toLowerCase().trim();
  const s2 = str2.toLowerCase().trim();
  
  // Exact match
  if (s1 === s2) return 1.0;
  
  // Contains match
  if (s1.includes(s2) || s2.includes(s1)) return 0.9;
  
  // Handle & vs and
  const normalized1 = s1.replace(/&/g, 'and').replace(/\s+/g, ' ');
  const normalized2 = s2.replace(/&/g, 'and').replace(/\s+/g, ' ');
  if (normalized1 === normalized2) return 0.95;
  
  // Levenshtein distance for fuzzy matching
  const distance = levenshteinDistance(normalized1, normalized2);
  const maxLength = Math.max(normalized1.length, normalized2.length);
  const similarity = 1 - (distance / maxLength);
  
  return similarity;
}

function levenshteinDistance(str1, str2) {
  const matrix = [];
  
  for (let i = 0; i <= str2.length; i++) {
    matrix[i] = [i];
  }
  
  for (let j = 0; j <= str1.length; j++) {
    matrix[0][j] = j;
  }
  
  for (let i = 1; i <= str2.length; i++) {
    for (let j = 1; j <= str1.length; j++) {
      if (str2.charAt(i - 1) === str1.charAt(j - 1)) {
        matrix[i][j] = matrix[i - 1][j - 1];
      } else {
        matrix[i][j] = Math.min(
          matrix[i - 1][j - 1] + 1,
          matrix[i][j - 1] + 1,
          matrix[i - 1][j] + 1
        );
      }
    }
  }
  
  return matrix[str2.length][str1.length];
}

// Find sheet by fuzzy matching
function findSheet(workbook, targetNames) {
  if (!Array.isArray(targetNames)) {
    targetNames = [targetNames];
  }
  
  let bestMatch = null;
  let bestScore = 0;
  
  workbook.eachSheet((sheet) => {
    for (const target of targetNames) {
      const score = fuzzyMatch(sheet.name, target);
      if (score > bestScore) {
        bestScore = score;
        bestMatch = sheet;
      }
    }
  });
  
  return bestScore > 0.6 ? bestMatch : null;
}

// Check if a cell is empty
function isCellEmpty(cell) {
  if (!cell || !cell.value) return true;
  
  const value = cell.value;
  
  // Handle formula cells
  if (typeof value === 'object' && value.formula) {
    const result = value.result;
    return !result || String(result).trim() === '';
  }
  
  // Handle regular values
  return String(value).trim() === '';
}

// Get cell display value
function getCellValue(cell) {
  if (!cell || !cell.value) return '';
  
  const value = cell.value;
  
  // Handle formula cells
  if (typeof value === 'object' && value.formula) {
    return String(value.result || '');
  }
  
  // Handle merged cells
  if (cell.isMerged && cell.master && cell.master !== cell) {
    return getCellValue(cell.master);
  }
  
  return String(value);
}

// Find table header in worksheet
function findTableHeader(worksheet, tableName, maxRows = 200, maxCols = 30) {
  const matches = [];
  
  // Get all variations of the table name
  const variations = TABLE_MAPPINGS[tableName] || [tableName.toLowerCase()];
  variations.push(tableName.toLowerCase()); // Add the original
  
  // Search for headers
  for (let row = 1; row <= Math.min(maxRows, worksheet.rowCount); row++) {
    for (let col = 1; col <= Math.min(maxCols, worksheet.columnCount); col++) {
      const cell = worksheet.getCell(row, col);
      const cellValue = getCellValue(cell);
      
      if (!cellValue) continue;
      
      // Check each variation
      for (const variation of variations) {
        const score = fuzzyMatch(cellValue, variation);
        if (score > 0.7) {
          matches.push({
            row,
            col,
            cell,
            value: cellValue,
            score,
            confidence: score >= 0.95 ? 'exact' : 'fuzzy'
          });
        }
      }
    }
  }
  
  // Sort by score and return best match
  matches.sort((a, b) => b.score - a.score);
  return matches[0] || null;
}

// Find table boundaries from header
function findTableBoundaries(worksheet, headerMatch, padding = 1) {
  const { row: headerRow, col: headerCol } = headerMatch;
  
  // Find left boundary (check if there's data to the left)
  let leftCol = headerCol;
  for (let col = headerCol - 1; col >= 1; col--) {
    const hasData = false;
    // Check if there's related data in the same row block
    for (let r = headerRow; r <= Math.min(headerRow + 5, worksheet.rowCount); r++) {
      if (!isCellEmpty(worksheet.getCell(r, col))) {
        leftCol = col;
        break;
      }
    }
    if (!hasData) break;
  }
  
  // Find right boundary
  let rightCol = headerCol;
  let emptyColCount = 0;
  for (let col = headerCol + 1; col <= worksheet.columnCount; col++) {
    let hasData = false;
    
    // Check column for data in the table area
    for (let r = headerRow; r <= Math.min(headerRow + 50, worksheet.rowCount); r++) {
      if (!isCellEmpty(worksheet.getCell(r, col))) {
        hasData = true;
        break;
      }
    }
    
    if (hasData) {
      rightCol = col;
      emptyColCount = 0;
    } else {
      emptyColCount++;
      if (emptyColCount >= 2) break;
    }
  }
  
  // Find bottom boundary
  let bottomRow = headerRow;
  let emptyRowCount = 0;
  for (let row = headerRow + 1; row <= worksheet.rowCount; row++) {
    let hasData = false;
    
    // Check row for data
    for (let col = leftCol; col <= rightCol; col++) {
      if (!isCellEmpty(worksheet.getCell(row, col))) {
        hasData = true;
        break;
      }
    }
    
    if (hasData) {
      bottomRow = row;
      emptyRowCount = 0;
      
      // Check if this might be a new section header (bold text after empty row)
      if (emptyRowCount > 0) {
        const firstCell = worksheet.getCell(row, leftCol);
        if (firstCell.font && firstCell.font.bold) {
          // This might be a new section, stop here
          bottomRow = row - emptyRowCount - 1;
          break;
        }
      }
    } else {
      emptyRowCount++;
      if (emptyRowCount >= 2) break;
    }
  }
  
  // Apply padding
  const startRow = Math.max(1, headerRow - padding);
  const endRow = Math.min(worksheet.rowCount, bottomRow + padding);
  const startCol = Math.max(1, leftCol - padding);
  const endCol = Math.min(worksheet.columnCount, rightCol + padding);
  
  // Convert to Excel range notation
  const startCell = worksheet.getCell(startRow, startCol).address;
  const endCell = worksheet.getCell(endRow, endCol).address;
  
  return {
    range: `${startCell}:${endCell}`,
    startRow,
    endRow,
    startCol,
    endCol,
    headerCell: worksheet.getCell(headerRow, headerCol).address
  };
}

// Find other potential tables as suggestions
function findSuggestions(worksheet, excludeRow, maxSuggestions = 5) {
  const suggestions = [];
  const seen = new Set();
  
  // Look for bold cells or cells with likely table names
  for (let row = 1; row <= Math.min(200, worksheet.rowCount); row++) {
    if (row === excludeRow) continue;
    
    for (let col = 1; col <= Math.min(10, worksheet.columnCount); col++) {
      const cell = worksheet.getCell(row, col);
      const value = getCellValue(cell);
      
      if (!value || seen.has(value)) continue;
      
      // Check if it might be a table header
      const valueLower = value.toLowerCase();
      const isLikelyTable = 
        (cell.font && cell.font.bold) ||
        valueLower.includes('table') ||
        valueLower.includes('schedule') ||
        valueLower.includes('summary') ||
        valueLower.includes('analysis') ||
        Object.keys(TABLE_MAPPINGS).some(key => 
          fuzzyMatch(value, key) > 0.5
        );
      
      if (isLikelyTable && value.length > 3) {
        suggestions.push(`${value} (${cell.address})`);
        seen.add(value);
        if (suggestions.length >= maxSuggestions) return suggestions;
      }
    }
  }
  
  return suggestions;
}

// Main detection function
async function detectTable(workbook, tableName, searchSheets = null, padding = 1) {
  const results = {
    found: false,
    sheet: null,
    range: null,
    headerCell: null,
    confidence: null,
    searchedSheets: [],
    suggestions: []
  };
  
  // Determine which sheets to search
  let sheetsToSearch = [];
  if (searchSheets && searchSheets.length > 0) {
    // Find specified sheets
    for (const sheetName of searchSheets) {
      const sheet = findSheet(workbook, sheetName);
      if (sheet) {
        sheetsToSearch.push(sheet);
        results.searchedSheets.push(sheet.name);
      }
    }
  } else {
    // Search all sheets
    workbook.eachSheet((sheet) => {
      sheetsToSearch.push(sheet);
      results.searchedSheets.push(sheet.name);
    });
  }
  
  // Search each sheet
  for (const sheet of sheetsToSearch) {
    const headerMatch = findTableHeader(sheet, tableName);
    
    if (headerMatch) {
      // Found the table
      const boundaries = findTableBoundaries(sheet, headerMatch, padding);
      
      results.found = true;
      results.sheet = sheet.name;
      results.range = boundaries.range;
      results.headerCell = boundaries.headerCell;
      results.confidence = headerMatch.confidence;
      
      return results;
    }
  }
  
  // Not found - generate suggestions from first searched sheet
  if (sheetsToSearch.length > 0) {
    results.suggestions = findSuggestions(sheetsToSearch[0]);
  }
  
  return results;
}

module.exports = {
  detectTable,
  fuzzyMatch,
  TABLE_MAPPINGS
};