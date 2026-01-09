const fs = require('fs');
const path = require('path');
const { exec } = require('child_process');
const { promisify } = require('util');

const execAsync = promisify(exec);

/**
 * LibreOffice-based Excel renderer - much simpler and more reliable
 * Uses LibreOffice's native Excel rendering, then crops to table ranges
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
   * Render Excel range to PNG using LibreOffice native rendering + cropping
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
    const pngFile = path.join(this.tempDir, `page_${timestamp}.png`);
    const croppedFile = path.join(this.tempDir, `cropped_${timestamp}.png`);

    try {
      // Step 1: Save Excel buffer to temporary file
      fs.writeFileSync(inputFile, excelBuffer);

      // Step 2: Get sheet index to determine which PDF page to extract
      const sheetIndex = await this.getSheetIndex(excelBuffer, sheetName);
      console.log(`Target sheet '${sheetName}' is at index ${sheetIndex}`);

      // Step 3: Convert Excel to PDF using LibreOffice (preserves ALL formatting)
      await this.convertExcelToPDF(inputFile, pdfFile, sheetName);

      // Step 4: Convert specific PDF page to PNG (the target sheet)
      await this.convertPDFtoPNG(pdfFile, pngFile, timestamp, sheetIndex + 1);

      // Step 5: Crop PNG to table range
      let finalFile = pngFile;
      
      if (range && range !== 'all') {
        console.log(`Cropping image to range ${range}...`);
        await this.cropImageToRange(pngFile, croppedFile, range, sheetName);
        finalFile = croppedFile;
      }

      // Step 6: Read and return PNG buffer
      if (!fs.existsSync(finalFile)) {
        throw new Error(`Final output file not found: ${finalFile}`);
      }

      const pngBuffer = fs.readFileSync(finalFile);

      // Cleanup temporary files
      this.cleanup([inputFile, pdfFile, pngFile, croppedFile]);

      return pngBuffer;

    } catch (error) {
      // Cleanup on error
      this.cleanup([inputFile, pdfFile, pngFile, croppedFile]);
      throw new Error(`Excel visual rendering failed: ${error.message}`);
    }
  }

  /**
   * Convert Excel to PDF using LibreOffice - with sheet selection
   * LibreOffice handles all Excel complexity: merges, formatting, number formats, etc.
   */
  async convertExcelToPDF(inputFile, outputFile, sheetName = null) {
    const outputDir = path.dirname(outputFile);
    
    try {
      let convertCmd;
      
      if (sheetName) {
        // Create a macro to set the active sheet before export
        const macroContent = `
Sub SetActiveSheetAndExport
  Dim oDoc As Object
  Dim oSheets As Object
  Dim oSheet As Object
  
  oDoc = ThisComponent
  oSheets = oDoc.getSheets()
  
  ' Find and activate the target sheet
  For i = 0 To oSheets.getCount() - 1
    oSheet = oSheets.getByIndex(i)
    If oSheet.getName() = "${sheetName}" Then
      oDoc.getCurrentController().setActiveSheet(oSheet)
      Exit For
    End If
  Next i
  
  ' Set print area to the active sheet only
  oSheet = oDoc.getCurrentController().getActiveSheet()
  oDoc.getCurrentController().setActiveSheet(oSheet)
  
End Sub
`;
        
        // For now, use a simpler approach: convert all sheets but we'll extract the right page later
        convertCmd = `soffice --headless --convert-to pdf --outdir "${outputDir}" "${inputFile}"`;
        console.log(`Converting Excel to PDF (will extract sheet '${sheetName}' later):`, convertCmd);
      } else {
        // Convert entire workbook
        convertCmd = `soffice --headless --convert-to pdf --outdir "${outputDir}" "${inputFile}"`;
        console.log('Converting Excel to PDF with LibreOffice:', convertCmd);
      }
      
      const { stdout, stderr } = await execAsync(convertCmd, { timeout: 30000 });
      
      if (stderr && !stderr.includes('Warning')) {
        console.warn('LibreOffice stderr:', stderr);
      }

      // LibreOffice creates PDF with same base name as input
      const baseName = path.basename(inputFile, path.extname(inputFile));
      const generatedPdf = path.join(outputDir, `${baseName}.pdf`);
      
      // Move to expected output location if needed
      if (fs.existsSync(generatedPdf) && generatedPdf !== outputFile) {
        fs.renameSync(generatedPdf, outputFile);
      }

      if (!fs.existsSync(outputFile)) {
        throw new Error('LibreOffice failed to generate PDF');
      }

      console.log('✅ Successfully converted Excel to PDF with LibreOffice');
      
    } catch (error) {
      if (error.message.includes('soffice')) {
        throw new Error('LibreOffice (soffice) not found. Install with: apt-get install libreoffice-calc');
      }
      throw error;
    }
  }

  /**
   * Get the index of a sheet within the Excel workbook
   */
  async getSheetIndex(excelBuffer, targetSheetName) {
    const ExcelJS = require('exceljs');
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.load(excelBuffer);

    let sheetIndex = 0;
    let found = false;
    
    workbook.eachSheet((sheet, index) => {
      const sheetName = sheet.name.trim();
      const targetName = targetSheetName.trim();
      
      if (sheetName.toLowerCase() === targetName.toLowerCase()) {
        sheetIndex = index - 1; // ExcelJS is 1-based, we need 0-based
        found = true;
      }
    });

    if (!found) {
      // Try partial match
      workbook.eachSheet((sheet, index) => {
        const sheetName = sheet.name.trim();
        const targetName = targetSheetName.trim();
        
        if (sheetName.toLowerCase().includes(targetName.toLowerCase()) || 
            targetName.toLowerCase().includes(sheetName.toLowerCase())) {
          sheetIndex = index - 1;
          found = true;
        }
      });
    }

    if (!found) {
      throw new Error(`Sheet '${targetSheetName}' not found in workbook`);
    }

    return sheetIndex;
  }

  /**
   * Convert PDF to PNG using poppler-utils (more reliable than ImageMagick)
   */
  async convertPDFtoPNG(pdfFile, outputTemplate, timestamp, pageNumber = 1) {
    try {
      // Use pdftoppm to convert PDF to PNG with high quality
      // -png: PNG format
      // -r 150: 150 DPI (good quality, reasonable file size)
      // -f X -l X: Extract specific page (pageNumber)
      // -singlefile: Single output file
      const baseName = `page_${timestamp}`;
      const outputDir = path.dirname(outputTemplate);
      const convertCmd = `pdftoppm -png -r 150 -f ${pageNumber} -l ${pageNumber} -singlefile "${pdfFile}" "${outputDir}/${baseName}"`;
      
      console.log(`Converting PDF page ${pageNumber} to PNG with poppler:`, convertCmd);
      const { stdout, stderr } = await execAsync(convertCmd, { timeout: 15000 });
      
      if (stderr && stderr.trim()) {
        console.warn('pdftoppm stderr:', stderr);
      }

      // pdftoppm with -singlefile creates filename with .png (no -1 suffix)
      const popplerOutput = `${outputDir}/${baseName}.png`;
      if (!fs.existsSync(popplerOutput)) {
        throw new Error(`pdftoppm failed to generate PNG for page ${pageNumber}`);
      }

      console.log(`✅ Successfully converted PDF page ${pageNumber} to PNG with poppler`);
      return popplerOutput;
      
    } catch (error) {
      // Fallback to ImageMagick if poppler fails
      console.warn('Poppler failed, trying ImageMagick fallback:', error.message);
      await this.convertPDFtoPNG_ImageMagick(pdfFile, outputTemplate);
    }
  }

  /**
   * ImageMagick fallback for PDF to PNG conversion
   */
  async convertPDFtoPNG_ImageMagick(pdfFile, pngFile) {
    try {
      // Use ImageMagick as fallback
      const convertCmd = `convert -density 150 -quality 95 -background white -alpha remove "${pdfFile}[0]" "${pngFile}"`;
      
      console.log('Converting PDF to PNG with ImageMagick (fallback)...');
      const { stdout, stderr } = await execAsync(convertCmd, { timeout: 15000 });
      
      // Check for PDF policy error
      if (stderr && stderr.includes('not authorized') && stderr.includes('PDF')) {
        throw new Error('ImageMagick PDF policy restriction. Install poppler-utils: apt-get install poppler-utils');
      }
      
      if (!fs.existsSync(pngFile)) {
        throw new Error('ImageMagick failed to generate PNG');
      }

      console.log('✅ Successfully converted PDF to PNG with ImageMagick');
      
    } catch (error) {
      throw new Error(`Both poppler and ImageMagick failed: ${error.message}`);
    }
  }

  /**
   * Crop image to specific table range using sharp
   * Calculates pixel coordinates from Excel cell positions
   */
  async cropImageToRange(inputPng, outputPng, range, sheetName) {
    const sharp = require('sharp');
    
    try {
      // Parse range (e.g., "A6:N27")
      const rangeMatch = range.match(/([A-Z]+)(\d+):([A-Z]+)(\d+)/);
      if (!rangeMatch) {
        throw new Error(`Invalid range format: ${range}. Expected format like "A6:N27"`);
      }
      
      const [, startCol, startRow, endCol, endRow] = rangeMatch;
      const startColNum = this.columnToNumber(startCol);
      const endColNum = this.columnToNumber(endCol);
      const startRowNum = parseInt(startRow, 10);
      const endRowNum = parseInt(endRow, 10);
      
      console.log(`Cropping range ${startCol}${startRow}:${endCol}${endRow} (cols ${startColNum}-${endColNum}, rows ${startRowNum}-${endRowNum})`);
      
      // Get image dimensions
      const image = sharp(inputPng);
      const { width, height } = await image.metadata();
      
      // Estimate cell dimensions (typical LibreOffice/Excel rendering)
      // These are approximate values that work well for most Excel exports
      const averageColWidth = 64;   // pixels per column (can vary)
      const averageRowHeight = 20;  // pixels per row (more consistent)
      const headerMargin = 40;      // top margin for sheet headers
      const leftMargin = 50;        // left margin for row numbers
      
      // Calculate crop coordinates
      const left = Math.max(0, leftMargin + (startColNum - 1) * averageColWidth);
      const top = Math.max(0, headerMargin + (startRowNum - 1) * averageRowHeight);
      const right = Math.min(width, leftMargin + endColNum * averageColWidth);
      const bottom = Math.min(height, headerMargin + endRowNum * averageRowHeight);
      
      const cropWidth = right - left;
      const cropHeight = bottom - top;
      
      console.log(`Crop coordinates: left=${left}, top=${top}, width=${cropWidth}, height=${cropHeight}`);
      
      if (cropWidth <= 0 || cropHeight <= 0) {
        throw new Error('Invalid crop dimensions calculated');
      }
      
      // Perform the crop
      await image.extract({ 
        left: Math.round(left), 
        top: Math.round(top), 
        width: Math.round(cropWidth), 
        height: Math.round(cropHeight) 
      }).png().toFile(outputPng);
      
      console.log(`✅ Successfully cropped image to range ${range}`);
      
    } catch (error) {
      console.warn('Cropping failed, using full image:', error.message);
      // If cropping fails, copy the original image
      const sharp = require('sharp');
      await sharp(inputPng).png().toFile(outputPng);
    }
  }
  
  /**
   * Convert column letter to number (A=1, B=2, ..., Z=26, AA=27, etc.)
   */
  columnToNumber(column) {
    let result = 0;
    for (let i = 0; i < column.length; i++) {
      result = result * 26 + (column.charCodeAt(i) - 'A'.charCodeAt(0) + 1);
    }
    return result;
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
   * Health check - verify LibreOffice and image tools are available
   */
  async healthCheck() {
    const checks = {
      libreoffice: false,
      poppler: false,
      imagemagick: false
    };
    
    try {
      await execAsync('soffice --version');
      checks.libreoffice = true;
    } catch (error) {
      checks.libreoffice = false;
    }
    
    try {
      await execAsync('pdftoppm -v');
      checks.poppler = true;
    } catch (error) {
      checks.poppler = false;
    }
    
    try {
      await execAsync('convert -version');
      checks.imagemagick = true;
    } catch (error) {
      checks.imagemagick = false;
    }

    return {
      available: checks.libreoffice && (checks.poppler || checks.imagemagick),
      libreoffice: checks.libreoffice,
      poppler: checks.poppler,
      imagemagick: checks.imagemagick,
      recommendation: checks.libreoffice 
        ? (checks.poppler 
            ? 'All systems ready' 
            : 'Install poppler-utils for better PDF conversion')
        : 'Install LibreOffice: apt-get install libreoffice-calc'
    };
  }
}

module.exports = ExcelVisualRenderer;