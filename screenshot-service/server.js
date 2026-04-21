console.log('Starting Excel screenshot service...');

const express = require('express');
const { spawn, exec, execSync } = require('child_process');

const app = express();
app.use(express.json({ limit: '50mb' }));

// Test Python UNO import on startup
exec('python3 -c "import uno; print(\'UNO import OK\')"', (err, stdout, stderr) => {
    if (err) {
        console.error('UNO import FAILED:', stderr);
    } else {
        console.log(stdout.trim());
    }
});

// Wait for LibreOffice listener to be ready
function waitForLibreOffice(maxAttempts = 10) {
    for (let i = 0; i < maxAttempts; i++) {
        try {
            execSync('python3 -c "import uno; from com.sun.star.connection import NoConnectException; ctx = uno.getComponentContext(); resolver = ctx.ServiceManager.createInstanceWithContext(\'com.sun.star.bridge.UnoUrlResolver\', ctx); resolver.resolve(\'uno:socket,host=localhost,port=2002;urp;StarOffice.ComponentContext\')"', { timeout: 5000 });
            console.log('LibreOffice listener ready!');
            return true;
        } catch (e) {
            console.log(`Waiting for LibreOffice... attempt ${i + 1}`);
            try {
                execSync('sleep 2');
            } catch (sleepErr) {
                // Ignore sleep errors
            }
        }
    }
    console.error('LibreOffice listener failed to start');
    return false;
}

// Wait for LibreOffice and verify it's running
setTimeout(() => {
    try {
        const result = execSync('ps aux | grep soffice').toString();
        console.log('LibreOffice processes:', result);
    } catch (e) {
        console.error('Cannot check soffice process:', e.message);
    }
    
    try {
        const result = execSync('netstat -tlnp 2>/dev/null | grep 2002 || ss -tlnp | grep 2002').toString();
        console.log('Port 2002 status:', result);
    } catch (e) {
        console.log('Port 2002: not listening or cannot check');
    }
    
    // Verify LibreOffice listener is ready
    waitForLibreOffice();
}, 5000);

app.post('/detect-and-capture', async (req, res) => {
    const { excelBase64, tableName, filename, dpi } = req.body;

    if (!excelBase64 || !tableName) {
        return res.status(400).json({ success: false, error: 'Missing excelBase64 or tableName' });
    }

    try {
        const result = await captureTable(excelBase64, tableName, { dpi });
        res.json(result);
    } catch (error) {
        console.error('Capture error:', error);
        res.status(500).json({ success: false, error: error.message });
    }
});

function captureTable(excelBase64, tableName, options = {}) {
    return new Promise((resolve, reject) => {
        const python = spawn('python3', ['/app/capture_table.py']);

        let stdout = '';
        let stderr = '';

        const payload = { excelBase64, tableName };
        if (options.dpi !== undefined) payload.dpi = options.dpi;

        python.stdin.write(JSON.stringify(payload));
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

        setTimeout(() => {
            python.kill();
            reject(new Error('Timeout'));
        }, 60000);
    });
}

app.get('/health', (req, res) => {
    exec('soffice --version', (error, stdout, stderr) => {
        res.json({
            status: 'healthy',
            service: 'excel-screenshot',
            visualRenderer: {
                available: !error,
                error: error ? error.message : null,
                version: stdout ? stdout.trim() : null
            }
        });
    });
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
    console.log(`Excel screenshot service running on port ${PORT}`);
});