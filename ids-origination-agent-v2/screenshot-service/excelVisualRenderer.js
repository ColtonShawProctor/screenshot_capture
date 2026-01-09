const fs = require('fs');
const path = require('path');
const { spawn } = require('child_process');

/**
 * Python-based Excel renderer using openpyxl for perfect range extraction
 * 
 * This approach is fundamentally more reliable than PDF cropping:
 * 1. Extract only the target range from source Excel (with all formatting)
 * 2. Create new workbook containing just that range 
 * 3. Render new workbook - entire image IS the table
 * 4. No cropping needed!
 */
class ExcelVisualRenderer {
  constructor() {
    this.pythonScript = path.join(__dirname, 'extract_table.py');
    this.ensurePythonScript();
  }

  ensurePythonScript() {
    if (!fs.existsSync(this.pythonScript)) {
      throw new Error(`Python script not found: ${this.pythonScript}`);
    }
  }

  /**
   * Render Excel range to PNG using Python openpyxl extraction
   * @param {Buffer} excelBuffer - Excel file buffer
   * @param {string} sheetName - Sheet name (e.g., "S&U ")
   * @param {string} range - Excel range (e.g., "A6:N27") 
   * @param {string} filename - Output filename (unused, kept for compatibility)
   * @returns {Promise<Buffer>} - PNG image buffer
   */
  async renderExcelRange(excelBuffer, sheetName, range, filename = 'screenshot.png') {
    try {
      console.log(`Extracting and rendering ${sheetName}!${range} using Python openpyxl...`);
      
      // Convert buffer to base64 for Python script
      const excelBase64 = excelBuffer.toString('base64');
      
      // Call Python script
      const result = await this.callPythonExtractor({
        excelBase64,
        sheetName,
        cellRange: range
      });
      
      if (!result.success) {
        throw new Error(result.error);
      }
      
      console.log(`✅ Successfully extracted table from ${result.source_sheet} using Python openpyxl`);
      
      // Convert base64 image back to buffer
      return Buffer.from(result.image, 'base64');
      
    } catch (error) {
      throw new Error(`Python extraction failed: ${error.message}`);
    }
  }

  /**
   * Call Python extractor script
   * @param {object} input - Input data for Python script
   * @returns {Promise<object>} - Result from Python script
   */
  async callPythonExtractor(input) {
    return new Promise((resolve, reject) => {
      const python = spawn('python3', [this.pythonScript], {
        stdio: ['pipe', 'pipe', 'pipe']
      });
      
      let stdout = '';
      let stderr = '';
      
      // Send input as JSON
      python.stdin.write(JSON.stringify(input));
      python.stdin.end();
      
      // Collect output
      python.stdout.on('data', (data) => {
        stdout += data.toString();
      });
      
      python.stderr.on('data', (data) => {
        stderr += data.toString();
      });
      
      // Handle completion
      python.on('close', (code) => {
        if (stderr) {
          console.log('Python script stderr:', stderr);
        }
        
        if (code === 0) {
          try {
            const result = JSON.parse(stdout);
            resolve(result);
          } catch (err) {
            reject(new Error(`Failed to parse Python output: ${err.message}`));
          }
        } else {
          reject(new Error(`Python script exited with code ${code}: ${stderr}`));
        }
      });
      
      // Handle errors
      python.on('error', (err) => {
        reject(new Error(`Failed to spawn Python process: ${err.message}`));
      });
      
      // Set timeout
      const timeout = setTimeout(() => {
        python.kill();
        reject(new Error('Python script timed out after 60 seconds'));
      }, 60000);
      
      python.on('close', () => {
        clearTimeout(timeout);
      });
    });
  }

  /**
   * Health check - verify Python and required tools are available
   */
  async healthCheck() {
    const checks = {
      python: false,
      openpyxl: false,
      libreoffice: false,
      poppler: false,
      script: false
    };
    
    try {
      // Check Python
      const pythonResult = await this.runCommand('python3', ['--version']);
      checks.python = pythonResult.success;
    } catch (error) {
      checks.python = false;
    }
    
    try {
      // Check openpyxl
      const openpyxlResult = await this.runCommand('python3', ['-c', 'import openpyxl; print("OK")']);
      checks.openpyxl = openpyxlResult.success;
    } catch (error) {
      checks.openpyxl = false;
    }
    
    try {
      // Check LibreOffice
      const libreResult = await this.runCommand('soffice', ['--version']);
      checks.libreoffice = libreResult.success;
    } catch (error) {
      checks.libreoffice = false;
    }
    
    try {
      // Check poppler
      const popplerResult = await this.runCommand('pdftoppm', ['-v']);
      checks.poppler = popplerResult.success;
    } catch (error) {
      checks.poppler = false;
    }
    
    // Check script exists
    checks.script = fs.existsSync(this.pythonScript);
    
    const allReady = checks.python && checks.openpyxl && checks.libreoffice && checks.poppler && checks.script;
    
    return {
      available: allReady,
      python: checks.python,
      openpyxl: checks.openpyxl,
      libreoffice: checks.libreoffice,
      poppler: checks.poppler,
      script: checks.script,
      recommendation: !allReady ? this.getRecommendation(checks) : 'All systems ready for Python extraction'
    };
  }

  /**
   * Get installation recommendations based on missing components
   */
  getRecommendation(checks) {
    const missing = [];
    
    if (!checks.python) missing.push('python3');
    if (!checks.openpyxl) missing.push('pip3 install openpyxl');
    if (!checks.libreoffice) missing.push('apt-get install libreoffice-calc');
    if (!checks.poppler) missing.push('apt-get install poppler-utils');
    if (!checks.script) missing.push('extract_table.py script missing');
    
    return `Install missing components: ${missing.join(', ')}`;
  }

  /**
   * Helper to run shell commands
   */
  async runCommand(command, args) {
    return new Promise((resolve) => {
      const proc = spawn(command, args, { stdio: 'pipe' });
      
      proc.on('close', (code) => {
        resolve({ success: code === 0 });
      });
      
      proc.on('error', () => {
        resolve({ success: false });
      });
      
      setTimeout(() => {
        proc.kill();
        resolve({ success: false });
      }, 5000);
    });
  }
}

module.exports = ExcelVisualRenderer;