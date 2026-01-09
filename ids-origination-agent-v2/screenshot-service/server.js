const express = require('express');
const { spawn } = require('child_process');

const app = express();
app.use(express.json({ limit: '50mb' }));

app.post('/detect-and-capture', async (req, res) => {
    const { excelBase64, tableName, filename } = req.body;
    
    if (!excelBase64 || !tableName) {
        return res.status(400).json({ 
            success: false, 
            error: 'Missing excelBase64 or tableName' 
        });
    }
    
    try {
        const result = await captureTable(excelBase64, tableName);
        res.json(result);
    } catch (error) {
        res.status(500).json({ 
            success: false, 
            error: error.message 
        });
    }
});

function captureTable(excelBase64, tableName) {
    return new Promise((resolve, reject) => {
        const python = spawn('python3', ['/app/capture_table.py']);
        
        let stdout = '';
        let stderr = '';
        
        python.stdin.write(JSON.stringify({ excelBase64, tableName }));
        python.stdin.end();
        
        python.stdout.on('data', (data) => { stdout += data; });
        python.stderr.on('data', (data) => { stderr += data; });
        
        python.on('close', (code) => {
            if (code === 0 && stdout) {
                try {
                    resolve(JSON.parse(stdout));
                } catch (e) {
                    reject(new Error(`Invalid JSON: ${stdout}`));
                }
            } else {
                reject(new Error(stderr || `Exit code ${code}`));
            }
        });
        
        // Timeout after 60 seconds
        setTimeout(() => {
            python.kill();
            reject(new Error('Timeout'));
        }, 60000);
    });
}

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
    console.log(`Excel screenshot service running on port ${PORT}`);
});