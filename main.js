const { app, BrowserWindow } = require('electron');
const path = require('path');
const http = require('http');
const { spawn } = require('child_process');

let mainWindow = null;
let flaskProcess = null;

// Poll the Flask backend until it responds or times out
function waitForBackend(maxAttempts = 20, delayMs = 1000) {
    return new Promise((resolve) => {
        let attempts = 0;

        function tryOnce() {
            const req = http.request(
                { host: '127.0.0.1', port: 5001, path: '/', method: 'GET' },
                (res) => {
                    if (res.statusCode === 200) return resolve(true);
                    retry();
                }
            );
            req.on('error', retry);
            req.end();
        }

        function retry() {
            attempts++;
            if (attempts >= maxAttempts) return resolve(false);
            setTimeout(tryOnce, delayMs);
        }

        tryOnce();
    });
}

function startFlaskServer() {
    // In development (`npm start`) always use the Python source so that
    // every change to web-prototype/ or scripts/ is picked up immediately.
    // Only switch to the bundled backend.exe when running as a packaged app.
    const isPackaged = app.isPackaged;
    const backendExe = path.join(__dirname, 'backend.exe');

    const cmd = isPackaged ? backendExe : 'python';
    const args = isPackaged ? [] : ['scripts/backend.py'];

    flaskProcess = spawn(cmd, args, {
        cwd: __dirname,
        env: { ...process.env, APP_ROOT: __dirname },
        // Detach=false so the child dies with parent if kill fails
        detached: false
    });

    flaskProcess.stdout.on('data', (data) => console.log(`[backend] ${data}`));
    flaskProcess.stderr.on('data', (data) => console.error(`[backend:err] ${data}`));
    flaskProcess.on('exit', (code) => console.log(`[backend] exited with code ${code}`));
}

function killFlaskServer() {
    if (!flaskProcess) return;
    try {
        // On Windows, use taskkill /T /F to kill the process tree
        if (process.platform === 'win32') {
            const { execSync } = require('child_process');
            execSync(`taskkill /PID ${flaskProcess.pid} /T /F`, { stdio: 'ignore' });
        } else {
            flaskProcess.kill('SIGTERM');
        }
    } catch (e) {
        console.error('[main] Failed to kill backend process:', e.message);
    }
    flaskProcess = null;
}

async function createWindow() {
    mainWindow = new BrowserWindow({
        width: 1200,
        height: 800,
        webPreferences: {
            nodeIntegration: false,
            contextIsolation: true,
        },
        autoHideMenuBar: true,
        icon: path.join(__dirname, 'assets', 'logo_manatech.png')
    });

    const ready = await waitForBackend();
    if (!ready) console.warn('[main] Backend did not respond in time — loading anyway.');

    mainWindow.loadURL('http://127.0.0.1:5001/');
    mainWindow.on('closed', () => { mainWindow = null; });
}

app.on('ready', () => {
    startFlaskServer();
    createWindow();
});

app.on('window-all-closed', () => {
    if (process.platform !== 'darwin') app.quit();
});

// Kill the backend process tree before quitting
app.on('before-quit', killFlaskServer);
app.on('quit', killFlaskServer);
