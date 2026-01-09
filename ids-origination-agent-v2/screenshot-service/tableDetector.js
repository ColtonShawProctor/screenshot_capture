const ExcelJS = require('exceljs');

// Table names and their common variations
const TABLE_MAPPINGS = {
  'Sources and Uses': ['sources and uses', 'sources & uses', 'source and use', 'source & use', 'sources / uses', 'sources/uses'],
  'Take Out Loan Sizing': ['take out loan sizing', 'takeout loan sizing', 'take-out loan sizing', 'loan sizing', 'takeout sizing', 'take out sizing'],
  'Capital Stack at Closing': ['capital stack at closing', 'capital stack', 'cap stack at closing', 'cap stack', 'capital stack closing'],
  'Loan to Cost': ['loan to cost', 'ltc', 'loan-to-cost', 'l2c', 'loan cost', 'ltc analysis'],
  'Loan to Value': ['loan to value', 'ltv', 'loan-to-value', 'l2v', 'loan value', 'ltv analysis'],
  'PILOT Schedule': ['pilot schedule', 'pilot', 'payment in lieu of taxes', 'p.i.l.o.t', 'pilot payment'],
  'Occupancy': ['occupancy', 'unit mix', 'unit occupancy', 'occupancy schedule'],
  'LTC and LTV': ['ltc and ltv', 'ltv and ltc', 'ltc & ltv', 'ltv & ltc', 'ltc/ltv', 'ltv/ltc']
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
  
  // Common sheet name patterns
  const commonSheetPatterns = {
    'sources and uses': ['s&u', 'su', 'sources uses', 'sources & uses', 's & u'],
    'ltc': ['ltc and ltv calcs', 'ltv ltc', 'ltc ltv', 'ltc and ltv', 'ltv and ltc'],
    'pilot': ['pilot', 'pilot schedule'],
    'capital': ['capital stack', 'cap stack']
  };
  
  workbook.eachSheet((sheet) => {
    const sheetName = sheet.name.toLowerCase().trim();
    
    for (const target of targetNames) {
      const targetLower = target.toLowerCase().trim();
      
      // Direct fuzzy match
      let score = fuzzyMatch(sheetName, targetLower);
      
      // Check common patterns
      for (const [key, patterns] of Object.entries(commonSheetPatterns)) {
        if (targetLower.includes(key)) {
          for (const pattern of patterns) {
            const patternScore = fuzzyMatch(sheetName, pattern);
            if (patternScore > score) {
              score = patternScore;
            }
          }
        }
      }
      
      // Boost score for exact substring matches
      if (sheetName.includes(targetLower) || targetLower.includes(sheetName)) {
        score = Math.max(score, 0.85);
      }
      
      if (score > bestScore) {
        bestScore = score;
        bestMatch = sheet;
      }
    }
  });
  
  return bestScore > 0.5 ? bestMatch : null;
}

// Check if a cell is empty
function isCellEmpty(cell) {
  if (!cell || !cell.value) return true;
  
  const value = cell.value;
  
  // Handle formula cells
  if (typeof value === 'object' && value.formula) {
    const result = value.result;
    // Treat #NAME? errors as empty
    if (result && String(result).includes('#NAME?')) return true;
    return !result || String(result).trim() === '';
  }
  
  // Handle regular values
  return String(value).trim() === '';
}

// Get cell display value without duplication from merged cells
function getCellValue(cell, visitedMerges = new Set()) {
  if (!cell || !cell.value) return '';
  
  // Handle merged cells - only get value from master cell
  if (cell.isMerged && cell.master && cell.master !== cell) {
    // Avoid infinite recursion
    const mergeId = `${cell.master.row}_${cell.master.col}`;
    if (visitedMerges.has(mergeId)) return '';
    visitedMerges.add(mergeId);
    return getCellValue(cell.master, visitedMerges);
  }
  
  const value = cell.value;
  
  // Handle formula cells
  if (typeof value === 'object' && value.formula) {
    const result = value.result || '';
    // Handle formula errors gracefully
    if (String(result).includes('#NAME?') || String(result).includes('#REF!')) {
      return '';
    }
    return String(result);
  }
  
  return String(value);
}

// Get merged cell range if the cell is part of a merge
function getMergedRange(worksheet, cell) {
  if (!cell.isMerged) return null;
  
  // Find the merge that contains this cell
  for (const merge of worksheet._merges) {
    const [startAddr, endAddr] = merge.split(':');
    const startCell = worksheet.getCell(startAddr);
    const endCell = worksheet.getCell(endAddr);
    
    if (cell.row >= startCell.row && cell.row <= endCell.row &&
        cell.col >= startCell.col && cell.col <= endCell.col) {
      return {
        startRow: startCell.row,
        endRow: endCell.row,
        startCol: startCell.col,
        endCol: endCell.col,
        master: worksheet.getCell(startCell.row, startCell.col)
      };
    }
  }
  
  return null;
}

// Find table header in worksheet
function findTableHeader(worksheet, tableName, maxRows = 200, maxCols = 30) {
  const matches = [];
  const processedMerges = new Set();
  
  // Get all variations of the table name
  const variations = TABLE_MAPPINGS[tableName] || [tableName.toLowerCase()];
  variations.push(tableName.toLowerCase()); // Add the original
  
  // Search for headers
  for (let row = 1; row <= Math.min(maxRows, worksheet.rowCount); row++) {
    for (let col = 1; col <= Math.min(maxCols, worksheet.columnCount); col++) {
      const cell = worksheet.getCell(row, col);
      
      // Skip if this cell is part of an already processed merge
      if (cell.isMerged) {
        const mergeRange = getMergedRange(worksheet, cell);
        if (mergeRange) {
          const mergeId = `${mergeRange.startRow}_${mergeRange.startCol}`;
          if (processedMerges.has(mergeId)) continue;
          processedMerges.add(mergeId);
          // Use the master cell for merged cells
          if (cell !== mergeRange.master) continue;
        }
      }
      
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

// Check if this looks like table data (not headers or random values)
function isTableData(value, cell) {
  if (!value || value.trim() === '') return false;
  
  const trimmedValue = value.trim();
  
  // Currency/number patterns
  if (/^\$?[\d,]+\.?\d*$/.test(trimmedValue)) return true;
  if (/^-?\$?[\d,]+\.?\d*$/.test(trimmedValue)) return true; // Negative numbers
  if (/^\d+\.?\d*%?$/.test(trimmedValue)) return true; // Percentages
  if (/^[\(\)]?\$?[\d,]+\.?\d*[\(\)]?$/.test(trimmedValue)) return true; // Accounting format
  
  // Common table row labels
  const lowerValue = trimmedValue.toLowerCase();
  if (lowerValue.includes('total') || 
      lowerValue.includes('subtotal') ||
      lowerValue.includes('loan') ||
      lowerValue.includes('equity') ||
      lowerValue.includes('cost') ||
      lowerValue.includes('value')) {
    return true;
  }
  
  // Text that's likely a row label (not too short, not a cell reference)
  if (trimmedValue.length > 3 && !/^[A-Z]\d+$/.test(trimmedValue)) {
    // Check if it has table-like formatting (bold, borders, etc.)
    if (cell && (cell.font?.bold || cell.border)) {
      return true;
    }
    // Common patterns for table data
    if (/^[A-Za-z\s\-&\/]+$/.test(trimmedValue) && trimmedValue.length < 50) {
      return true;
    }
  }
  
  return false;
}

// Find table boundaries from header - IMPROVED VERSION
function findTableBoundaries(worksheet, headerMatch, padding = 1) {
  const { row: headerRow, col: headerCol } = headerMatch;
  
  // Start with header position as initial boundaries
  let leftCol = headerCol;
  let rightCol = headerCol;
  let topRow = headerRow;
  let bottomRow = headerRow;
  
  // For two-sided tables (like Sources & Uses), we need to detect gaps
  let gapCols = [];
  let hasSignificantGap = false;
  
  // Step 1: Find the actual header row extent (might be merged cells)
  const headerCell = worksheet.getCell(headerRow, headerCol);
  if (headerCell.isMerged) {
    const mergeRange = getMergedRange(worksheet, headerCell);
    if (mergeRange) {
      leftCol = Math.min(leftCol, mergeRange.startCol);
      rightCol = Math.max(rightCol, mergeRange.endCol);
    }
  }
  
  // Step 2: Find left boundary - look for start of table data
  let foundDataLeft = false;
  for (let col = headerCol; col >= 1; col--) {
    let hasData = false;
    
    // Check multiple rows below header
    for (let r = headerRow; r <= Math.min(headerRow + 10, worksheet.rowCount); r++) {
      const cell = worksheet.getCell(r, col);
      const value = getCellValue(cell);
      
      if (isTableData(value, cell)) {
        hasData = true;
        leftCol = col;
        foundDataLeft = true;
        break;
      }
    }
    
    if (!hasData && foundDataLeft) {
      // Found the edge
      break;
    }
  }
  
  // Step 3: Find right boundary - handle two-sided tables with gaps
  let consecutiveEmptyCols = 0;
  let rightmostDataCol = headerCol;
  
  for (let col = headerCol; col <= Math.min(worksheet.columnCount, headerCol + 20); col++) {
    let hasData = false;
    
    // Check if column has data
    for (let r = headerRow; r <= Math.min(headerRow + 10, worksheet.rowCount); r++) {
      const cell = worksheet.getCell(r, col);
      const value = getCellValue(cell);
      
      if (isTableData(value, cell)) {
        hasData = true;
        break;
      }
    }
    
    if (hasData) {
      rightmostDataCol = col;
      consecutiveEmptyCols = 0;
      
      // Check if we've crossed a significant gap (for two-sided tables)
      if (gapCols.length > 0 && col - gapCols[gapCols.length - 1] > 1) {
        hasSignificantGap = true;
      }
    } else {
      consecutiveEmptyCols++;
      gapCols.push(col);
      
      // For two-sided tables, continue looking after a gap
      if (consecutiveEmptyCols <= 3 && col < headerCol + 15) {
        continue;
      } else if (!hasSignificantGap) {
        // No significant gap found, so this is the end
        break;
      }
    }
  }
  
  rightCol = rightmostDataCol;
  
  // Step 4: Find top boundary - look for additional header rows
  for (let row = headerRow - 1; row >= Math.max(1, headerRow - 3); row--) {
    let hasHeaderContent = false;
    
    for (let col = leftCol; col <= rightCol; col++) {
      if (gapCols.includes(col)) continue;
      
      const cell = worksheet.getCell(row, col);
      const value = getCellValue(cell);
      
      if (value && (cell.font?.bold || cell.fill || 
          value.toLowerCase().includes('source') ||
          value.toLowerCase().includes('use') ||
          value.toLowerCase().includes('amount'))) {
        hasHeaderContent = true;
        topRow = row;
        break;
      }
    }
    
    if (!hasHeaderContent) break;
  }
  
  // Step 5: Find bottom boundary - look for end of data or total rows
  let lastDataRow = headerRow;
  let foundTotal = false;
  
  for (let row = headerRow + 1; row <= Math.min(worksheet.rowCount, headerRow + 50); row++) {
    let hasData = false;
    let rowValues = [];
    
    for (let col = leftCol; col <= rightCol; col++) {
      if (gapCols.includes(col)) continue;
      
      const cell = worksheet.getCell(row, col);
      const value = getCellValue(cell);
      
      if (value && !isCellEmpty(cell)) {
        hasData = true;
        rowValues.push(value.toLowerCase());
      }
    }
    
    if (hasData) {
      lastDataRow = row;
      
      // Check if this is a total row
      const rowText = rowValues.join(' ');
      if (rowText.includes('total') && !rowText.includes('subtotal')) {
        foundTotal = true;
        bottomRow = row;
        break;
      }
    } else {
      // Empty row - check if we should stop
      if (row - lastDataRow > 1) {
        bottomRow = lastDataRow;
        break;
      }
    }
  }
  
  if (!foundTotal) {
    bottomRow = lastDataRow;
  }
  
  // Apply minimal padding
  const finalPadding = Math.min(padding, 1);
  const startRow = Math.max(1, topRow - finalPadding);
  const endRow = Math.min(worksheet.rowCount, bottomRow + finalPadding);
  const startCol = Math.max(1, leftCol - finalPadding);
  const endCol = Math.min(worksheet.columnCount, rightCol + finalPadding);
  
  // Convert to Excel range notation
  const startCell = worksheet.getCell(startRow, startCol).address;
  const endCell = worksheet.getCell(endRow, endCol).address;
  
  return {
    range: `${startCell}:${endCell}`,
    startRow,
    endRow,
    startCol,
    endCol,
    headerCell: worksheet.getCell(headerRow, headerCol).address,
    hasGap: hasSignificantGap,
    gapColumns: gapCols,
    actualBounds: {
      topRow,
      bottomRow,
      leftCol,
      rightCol
    }
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
  
  // Determine which sheets to search with smart sheet selection
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
    // Smart sheet selection based on table name
    const tableNameLower = tableName.toLowerCase();
    
    if (tableNameLower.includes('sources') || tableNameLower.includes('uses') || 
        tableNameLower.includes('capital stack') || tableNameLower.includes('takeout') ||
        tableNameLower.includes('take out')) {
      // Look for S&U sheet first
      const suSheet = findSheet(workbook, ['s&u', 'sources and uses', 'su', 's & u']);
      if (suSheet) sheetsToSearch.push(suSheet);
    }
    
    if (tableNameLower.includes('ltc') || tableNameLower.includes('ltv') || 
        tableNameLower.includes('loan to')) {
      // Look for LTC/LTV sheet first
      const ltcSheet = findSheet(workbook, ['ltc and ltv calcs', 'ltc ltv', 'ltv ltc']);
      if (ltcSheet) sheetsToSearch.push(ltcSheet);
    }
    
    if (tableNameLower.includes('pilot')) {
      // Look for PILOT sheet
      const pilotSheet = findSheet(workbook, ['pilot', 'pilot schedule']);
      if (pilotSheet) sheetsToSearch.push(pilotSheet);
    }
    
    // If no specific sheets found or for other tables, search all sheets
    if (sheetsToSearch.length === 0) {
      workbook.eachSheet((sheet) => {
        sheetsToSearch.push(sheet);
      });
    }
    
    // Record searched sheets
    sheetsToSearch.forEach(sheet => results.searchedSheets.push(sheet.name));
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
      results.actualBounds = boundaries.actualBounds;
      
      return results;
    }
  }
  
  // Not found - generate suggestions from all searched sheets
  const allSuggestions = [];
  for (const sheet of sheetsToSearch) {
    const suggestions = findSuggestions(sheet, -1, 3);
    suggestions.forEach(suggestion => {
      allSuggestions.push(`${suggestion} [${sheet.name}]`);
    });
  }
  
  results.suggestions = [...new Set(allSuggestions)].slice(0, 8);
  
  return results;
}

module.exports = {
  detectTable,
  fuzzyMatch,
  TABLE_MAPPINGS
};