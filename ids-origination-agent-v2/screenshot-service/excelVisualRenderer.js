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
    // LibreOffice command to convert Excel to PDF
    // --headless: run without GUI
    // --convert-to pdf: output format
    // --outdir: output directory
    const outputDir = path.dirname(outputFile);
    
    // Basic conversion first (LibreOffice doesn't support range selection directly)
    const convertCmd = `libreoffice --headless --convert-to pdf --outdir "${outputDir}" "${inputFile}"`;
    
    console.log('Converting Excel to PDF:', convertCmd);
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

    console.log('Successfully converted Excel to PDF');
  }

  /**
   * Convert PDF to PNG with high quality settings
   */
  async convertPDFtoPNG(pdfFile, pngFile) {
    // Try ImageMagick first (preferred for quality)
    try {
      await this.convertPDFtoPNG_ImageMagick(pdfFile, pngFile);
      return;
    } catch (error) {
      console.warn('ImageMagick failed, trying poppler fallback:', error.message);
    }
    
    // Fallback to poppler-utils (pdftoppm)
    try {
      await this.convertPDFtoPNG_Poppler(pdfFile, pngFile);
      return;
    } catch (error) {
      throw new Error(`Both ImageMagick and poppler failed: ${error.message}`);
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
    
    console.log('Converting PDF to PNG with ImageMagick:', convertCmd);
    const { stdout, stderr } = await execAsync(convertCmd);
    
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
    
    console.log('Converting PDF to PNG with poppler:', convertCmd);
    const { stdout, stderr } = await execAsync(convertCmd);
    
    if (stderr) {
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
   * Alternative method: Use LibreOffice macro for precise range selection
   * This is more complex but allows exact range extraction
   */
  async renderExcelRangeWithMacro(excelBuffer, sheetName, range, filename = 'screenshot.png') {
    const timestamp = Date.now();
    const inputFile = path.join(this.tempDir, `input_${timestamp}.xlsx`);
    const outputFile = path.join(this.tempDir, `output_${timestamp}.png`);
    const macroFile = path.join(this.tempDir, `export_range_${timestamp}.bas`);

    try {
      // Save Excel file
      fs.writeFileSync(inputFile, excelBuffer);

      // Create LibreOffice Basic macro for range export
      const macro = this.generateRangeExportMacro(sheetName, range, outputFile);
      fs.writeFileSync(macroFile, macro);

      // Run LibreOffice with macro
      const macroCmd = `libreoffice --headless --invisible --macro-execute "${macroFile}" "${inputFile}"`;
      await execAsync(macroCmd);

      if (!fs.existsSync(outputFile)) {
        // Fallback to PDF method if macro fails
        console.log('Macro method failed, falling back to PDF conversion');
        return await this.renderExcelRange(excelBuffer, sheetName, range, filename);
      }

      const pngBuffer = fs.readFileSync(outputFile);
      this.cleanup([inputFile, outputFile, macroFile]);
      return pngBuffer;

    } catch (error) {
      console.log('Macro method failed, falling back to PDF conversion:', error.message);
      this.cleanup([inputFile, outputFile, macroFile]);
      return await this.renderExcelRange(excelBuffer, sheetName, range, filename);
    }
  }

  /**
   * Generate LibreOffice Basic macro for range export
   */
  generateRangeExportMacro(sheetName, range, outputFile) {
    return `
Sub ExportRange
    Dim oDoc As Object
    Dim oSheet As Object
    Dim oRange As Object
    Dim oExportProps(0) As New com.sun.star.beans.PropertyValue
    
    ' Open document
    oDoc = ThisComponent
    
    ' Get specific sheet
    oSheet = oDoc.getSheets().getByName("${sheetName}")
    
    ' Select range
    oRange = oSheet.getCellRangeByName("${range}")
    
    ' Set up export properties
    oExportProps(0).Name = "FilterName"
    oExportProps(0).Value = "calc_png_Export"
    
    ' Export range as PNG
    oRange.storeToURL("file://${outputFile}", oExportProps())
End Sub
    `;
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