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
    const pngFile = path.join(this.tempDir, `page_${timestamp}-1.png`);
    const croppedFile = path.join(this.tempDir, `cropped_${timestamp}.png`);

    try {
      // Step 1: Save Excel buffer to temporary file
      fs.writeFileSync(inputFile, excelBuffer);

      // Step 2: Convert Excel to PDF using LibreOffice (preserves ALL formatting)
      await this.convertExcelToPDF(inputFile, pdfFile);

      // Step 3: Convert PDF to PNG
      await this.convertPDFtoPNG(pdfFile, pngFile, timestamp);

      // Step 4: Crop PNG to table range (optional - if we need precise cropping)
      // For now, return the full page - LibreOffice does excellent formatting
      let finalFile = pngFile;
      
      // If cropping is needed in the future:
      // await this.cropImageToRange(pngFile, croppedFile, range, sheetName);
      // finalFile = croppedFile;

      // Step 5: Read and return PNG buffer
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
   * Convert Excel to PDF using LibreOffice - THE SIMPLE WAY
   * LibreOffice handles all Excel complexity: merges, formatting, number formats, etc.
   */
  async convertExcelToPDF(inputFile, outputFile) {
    const outputDir = path.dirname(outputFile);
    
    try {
      // Use soffice (LibreOffice) to convert Excel to PDF
      // This preserves ALL Excel formatting perfectly
      const convertCmd = `soffice --headless --convert-to pdf --outdir "${outputDir}" "${inputFile}"`;
      
      console.log('Converting Excel to PDF with LibreOffice:', convertCmd);
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
   * Convert PDF to PNG using poppler-utils (more reliable than ImageMagick)
   */
  async convertPDFtoPNG(pdfFile, outputTemplate, timestamp) {
    try {
      // Use pdftoppm to convert PDF to PNG with high quality
      // -png: PNG format
      // -r 150: 150 DPI (good quality, reasonable file size)
      // -f 1 -l 1: First page only
      // -singlefile: Single output file
      const baseName = `page_${timestamp}`;
      const outputDir = path.dirname(outputTemplate);
      const convertCmd = `pdftoppm -png -r 150 -f 1 -l 1 -singlefile "${pdfFile}" "${outputDir}/${baseName}"`;
      
      console.log('Converting PDF to PNG with poppler:', convertCmd);
      const { stdout, stderr } = await execAsync(convertCmd, { timeout: 15000 });
      
      if (stderr && stderr.trim()) {
        console.warn('pdftoppm stderr:', stderr);
      }

      // pdftoppm creates filename with -1.png suffix automatically
      const popplerOutput = `${outputDir}/${baseName}.png`;
      if (!fs.existsSync(popplerOutput)) {
        throw new Error('pdftoppm failed to generate PNG');
      }

      console.log('✅ Successfully converted PDF to PNG with poppler');
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
   * Future enhancement: Crop image to specific table range
   * This would calculate pixel coordinates from Excel cell positions
   */
  async cropImageToRange(inputPng, outputPng, range, sheetName) {
    // TODO: Implement precise cropping if needed
    // For now, LibreOffice's full-page rendering is excellent quality
    console.log(`Cropping not implemented yet. Range: ${range}, Sheet: ${sheetName}`);
    
    // For future implementation:
    // 1. Parse range (e.g., "A6:N27")
    // 2. Calculate approximate pixel coordinates based on typical Excel cell sizes
    // 3. Use sharp or jimp to crop the image
    // 4. Save cropped result
    
    // Example with sharp:
    // const sharp = require('sharp');
    // const image = sharp(inputPng);
    // const { width, height } = await image.metadata();
    // await image.extract({ left: x, top: y, width: w, height: h }).png().toFile(outputPng);
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