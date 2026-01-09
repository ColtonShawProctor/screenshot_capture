const fs = require('fs');
const path = require('path');
const { exec } = require('child_process');
const { promisify } = require('util');

const execAsync = promisify(exec);

/**
 * Visual Excel renderer that preserves actual Excel formatting
 * using LibreOffice headless mode for true WYSIWYG rendering
 */
class ExcelVisualRenderer {
  constructor() {
    this.tempDir = '/tmp/excel-renderer';
    this.ensureTempDir();
  }

  ensureTempDir() {
    if (!fs.existsSync(this.tempDir)) {
      fs.mkdirSync(this.tempDir, { recursive: true });
    }
  }

  /**
   * Render Excel range to PNG image preserving all formatting
   * @param {Buffer} excelBuffer - Excel file buffer
   * @param {string} sheetName - Sheet name 
   * @param {string} range - Excel range (e.g., "A1:H30")
   * @param {string} filename - Output filename
   * @returns {Promise<Buffer>} - PNG image buffer
   */
  async renderExcelRange(excelBuffer, sheetName, range, filename = 'screenshot.png') {
    const timestamp = Date.now();
    const inputFile = path.join(this.tempDir, `input_${timestamp}.xlsx`);
    const pdfFile = path.join(this.tempDir, `output_${timestamp}.pdf`);
    const pngFile = path.join(this.tempDir, `output_${timestamp}.png`);

    try {
      // Step 1: Save Excel buffer to temporary file
      fs.writeFileSync(inputFile, excelBuffer);

      // Step 2: Convert Excel to PDF using LibreOffice (preserves formatting)
      // This method maintains ALL Excel formatting: colors, borders, fonts, number formats
      await this.convertExcelToPDF(inputFile, pdfFile, sheetName, range);

      // Step 3: Convert PDF to high-quality PNG
      await this.convertPDFtoPNG(pdfFile, pngFile);

      // Step 4: Read and return PNG buffer
      const pngBuffer = fs.readFileSync(pngFile);

      // Cleanup temporary files
      this.cleanup([inputFile, pdfFile, pngFile]);

      return pngBuffer;

    } catch (error) {
      // Cleanup on error
      this.cleanup([inputFile, pdfFile, pngFile]);
      throw new Error(`Excel visual rendering failed: ${error.message}`);
    }
  }

  /**
   * Convert Excel to PDF using LibreOffice with range selection
   */
  async convertExcelToPDF(inputFile, outputFile, sheetName, range) {
    const timestamp = Date.now();
    const rangeExcelFile = path.join(this.tempDir, `range_${timestamp}.xlsx`);
    
    try {
      // Step 1: Create new Excel file with only the target range
      await this.createRangeOnlyExcel(inputFile, rangeExcelFile, sheetName, range);
      
      // Step 2: Convert the range-only Excel to PDF
      await this.convertExcelToFullPDF(rangeExcelFile, outputFile);
      
      console.log(`Successfully converted Excel range ${sheetName}!${range} to PDF`);
      
    } finally {
      // Cleanup temporary range Excel file
      try {
        if (fs.existsSync(rangeExcelFile)) {
          fs.unlinkSync(rangeExcelFile);
        }
      } catch (e) {}
    }
  }

  /**
   * Create a new Excel file containing only the specified range
   */
  async createRangeOnlyExcel(inputFile, outputFile, sheetName, range) {
    const ExcelJS = require('exceljs');
    
    // Load the original workbook
    const sourceWorkbook = new ExcelJS.Workbook();
    await sourceWorkbook.xlsx.readFile(inputFile);
    
    // Find the source sheet
    let sourceSheet = null;
    sourceWorkbook.eachSheet((sheet) => {
      if (sheet.name.toLowerCase() === sheetName.toLowerCase()) {
        sourceSheet = sheet;
      }
    });
    
    if (!sourceSheet) {
      // Try partial match
      sourceWorkbook.eachSheet((sheet) => {
        if (sheet.name.toLowerCase().includes(sheetName.toLowerCase()) || 
            sheetName.toLowerCase().includes(sheet.name.toLowerCase())) {
          sourceSheet = sheet;
        }
      });
    }
    
    if (!sourceSheet) {
      throw new Error(`Sheet "${sheetName}" not found in workbook`);
    }
    
    // Parse the range (e.g., "A4:N12")
    const rangeMatch = range.match(/([A-Z]+)(\d+):([A-Z]+)(\d+)/);
    if (!rangeMatch) {
      throw new Error('Invalid range format');
    }
    
    const startCol = this.columnToNumber(rangeMatch[1]);
    const startRow = parseInt(rangeMatch[2]);
    const endCol = this.columnToNumber(rangeMatch[3]);
    const endRow = parseInt(rangeMatch[4]);
    
    // Create new workbook with only the range data
    const targetWorkbook = new ExcelJS.Workbook();
    const targetSheet = targetWorkbook.addWorksheet('ExtractedRange');
    
    // Copy cells from source range to target sheet (starting at A1)
    let targetRowNum = 1;
    for (let sourceRowNum = startRow; sourceRowNum <= endRow; sourceRowNum++) {
      const sourceRow = sourceSheet.getRow(sourceRowNum);
      const targetRow = targetSheet.getRow(targetRowNum);
      
      let targetColNum = 1;
      for (let sourceColNum = startCol; sourceColNum <= endCol; sourceColNum++) {
        const sourceCell = sourceRow.getCell(sourceColNum);
        const targetCell = targetRow.getCell(targetColNum);
        
        // Copy cell value
        targetCell.value = sourceCell.value;
        
        // Copy cell styling (font, fill, border, alignment)
        if (sourceCell.font) targetCell.font = sourceCell.font;
        if (sourceCell.fill) targetCell.fill = sourceCell.fill;
        if (sourceCell.border) targetCell.border = sourceCell.border;
        if (sourceCell.alignment) targetCell.alignment = sourceCell.alignment;
        if (sourceCell.numFmt) targetCell.numFmt = sourceCell.numFmt;
        
        targetColNum++;
      }
      targetRowNum++;
    }
    
    // Copy column widths proportionally
    for (let colIndex = 0; colIndex < (endCol - startCol + 1); colIndex++) {
      const sourceColNum = startCol + colIndex;
      const targetColNum = colIndex + 1;
      
      const sourceColWidth = sourceSheet.getColumn(sourceColNum).width;
      if (sourceColWidth) {
        targetSheet.getColumn(targetColNum).width = sourceColWidth;
      }
    }
    
    // Copy row heights
    for (let rowIndex = 0; rowIndex < (endRow - startRow + 1); rowIndex++) {
      const sourceRowNum = startRow + rowIndex;
      const targetRowNum = rowIndex + 1;
      
      const sourceRowHeight = sourceSheet.getRow(sourceRowNum).height;
      if (sourceRowHeight) {
        targetSheet.getRow(targetRowNum).height = sourceRowHeight;
      }
    }
    
    // Save the range-only Excel file
    await targetWorkbook.xlsx.writeFile(outputFile);
    
    console.log(`Created range-only Excel file: ${range} → ${outputFile}`);
  }

  /**
   * Helper function to convert column letter to number
   */
  columnToNumber(column) {
    let result = 0;
    for (let i = 0; i < column.length; i++) {
      result = result * 26 + (column.charCodeAt(i) - 'A'.charCodeAt(0) + 1);
    }
    return result;
  }



  /**
   * Fallback: Convert entire Excel to PDF
   */
  async convertExcelToFullPDF(inputFile, outputFile) {
    const outputDir = path.dirname(outputFile);
    
    // Basic full-sheet conversion
    const convertCmd = `libreoffice --headless --convert-to pdf --outdir "${outputDir}" "${inputFile}"`;
    
    console.log('Converting Excel to PDF (full sheet):', convertCmd);
    const { stdout, stderr } = await execAsync(convertCmd);
    
    if (stderr && !stderr.includes('Warning')) {
      console.warn('LibreOffice stderr:', stderr);
    }

    // LibreOffice creates PDF with same base name as input
    const baseName = path.basename(inputFile, path.extname(inputFile));
    const generatedPdf = path.join(outputDir, `${baseName}.pdf`);
    
    // Move to expected output location
    if (fs.existsSync(generatedPdf) && generatedPdf !== outputFile) {
      fs.renameSync(generatedPdf, outputFile);
    }

    if (!fs.existsSync(outputFile)) {
      throw new Error('LibreOffice failed to generate PDF');
    }

    console.log('Successfully converted Excel to PDF (full sheet fallback)');
  }

  /**
   * Convert PDF to PNG with high quality settings
   */
  async convertPDFtoPNG(pdfFile, pngFile) {
    // Try poppler-utils first (more reliable for PDFs)
    try {
      await this.convertPDFtoPNG_Poppler(pdfFile, pngFile);
      return;
    } catch (error) {
      console.warn('Poppler failed, trying ImageMagick fallback:', error.message);
    }
    
    // Fallback to ImageMagick (may have PDF policy restrictions)
    try {
      await this.convertPDFtoPNG_ImageMagick(pdfFile, pngFile);
      return;
    } catch (error) {
      throw new Error(`Both poppler and ImageMagick failed: ${error.message}`);
    }
  }

  /**
   * Convert PDF to PNG using ImageMagick
   */
  async convertPDFtoPNG_ImageMagick(pdfFile, pngFile) {
    // Use ImageMagick to convert PDF to PNG with high quality
    // -density 300: High DPI for crisp text
    // -quality 95: High quality
    // -background white: Ensure white background
    // -alpha remove: Remove transparency
    // [0]: Take first page only
    const convertCmd = `convert -density 300 -quality 95 -background white -alpha remove "${pdfFile}[0]" "${pngFile}"`;
    
    console.log('Converting PDF to PNG with ImageMagick...');
    const { stdout, stderr } = await execAsync(convertCmd);
    
    // Check for common PDF policy error
    if (stderr && stderr.includes('not authorized') && stderr.includes('PDF')) {
      throw new Error('ImageMagick PDF policy restriction - use poppler instead');
    }
    
    if (stderr && !stderr.includes('Warning')) {
      console.warn('ImageMagick stderr:', stderr);
    }

    if (!fs.existsSync(pngFile)) {
      throw new Error('ImageMagick failed to generate PNG');
    }

    console.log('Successfully converted PDF to PNG with ImageMagick');
  }

  /**
   * Convert PDF to PNG using poppler-utils (fallback)
   */
  async convertPDFtoPNG_Poppler(pdfFile, pngFile) {
    // Use pdftoppm from poppler-utils as fallback
    // -png: Output PNG format
    // -r 300: 300 DPI resolution
    // -f 1 -l 1: First page only
    // -singlefile: Single output file
    const baseName = path.basename(pngFile, '.png');
    const outputDir = path.dirname(pngFile);
    const convertCmd = `pdftoppm -png -r 300 -f 1 -l 1 -singlefile "${pdfFile}" "${outputDir}/${baseName}"`;
    
    console.log('Converting PDF to PNG with poppler...');
    const { stdout, stderr } = await execAsync(convertCmd);
    
    if (stderr && stderr.trim()) {
      console.warn('pdftoppm stderr:', stderr);
    }

    // pdftoppm creates filename with .png extension automatically
    const popplerOutput = `${outputDir}/${baseName}.png`;
    if (!fs.existsSync(popplerOutput)) {
      throw new Error('pdftoppm failed to generate PNG');
    }

    // Move to expected location if different
    if (popplerOutput !== pngFile) {
      fs.renameSync(popplerOutput, pngFile);
    }

    console.log('Successfully converted PDF to PNG with poppler');
  }


  /**
   * Cleanup temporary files
   */
  cleanup(files) {
    files.forEach(file => {
      try {
        if (fs.existsSync(file)) {
          fs.unlinkSync(file);
        }
      } catch (error) {
        console.warn(`Failed to cleanup ${file}:`, error.message);
      }
    });
  }

  /**
   * Health check - verify LibreOffice is available
   */
  async healthCheck() {
    try {
      const { stdout } = await execAsync('libreoffice --version');
      return {
        available: true,
        version: stdout.trim()
      };
    } catch (error) {
      return {
        available: false,
        error: error.message
      };
    }
  }
}

module.exports = ExcelVisualRenderer;