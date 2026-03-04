const electron = require('electron');
const app = electron.app;
const BrowserWindow = electron.BrowserWindow;
const path = require('path');
const { spawn } = require('child_process');

let mainWindow;
let flaskProcess = null;

// function to check if backend is ready
async function isBackendReady() {
    try {
        const http = require('http');
        return new Promise((resolve) => {
            const req = http.request({
                host: '127.0.0.1',
                port: 5001,
                path: '/',
                method: 'GET'
            }, (res) => {
                resolve(res.statusCode === 200);
            });
            req.on('error', () => resolve(false));
            req.end();
        });
    } catch (e) {
        return false;
    }
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

    // Retry loop until backend is responsive
    let ready = false;
    let attempts = 0;
    while (!ready && attempts < 20) {
        ready = await isBackendReady();
        if (!ready) {
            console.log("Backend not ready yet, waiting...");
            await new Promise(r => setTimeout(r, 1000));
            attempts++;
        }
    }

    mainWindow.loadURL('http://127.0.0.1:5001/');

    mainWindow.on('closed', function () {
        mainWindow = null;
    });
}

function startFlaskServer() {
    const backendExe = path.join(__dirname, 'backend.exe');
    if (require('fs').existsSync(backendExe)) {
        // Run packaged python executable
        flaskProcess = spawn(backendExe, [], {
            cwd: __dirname,
            env: { ...process.env, APP_ROOT: __dirname }
        });
    } else {
        // Fallback for development (needs python on path)
        flaskProcess = spawn('python', ['scripts/backend.py'], {
            cwd: __dirname,
            env: { ...process.env, APP_ROOT: __dirname }
        });
    }

    flaskProcess.stdout.on('data', (data) => {
        console.log(`Backend stdout: ${data}`);
    });

    flaskProcess.stderr.on('data', (data) => {
        console.error(`Backend stderr: ${data}`);
    });
}

app.on('ready', () => {
    startFlaskServer();
    createWindow();
});

app.on('window-all-closed', function () {
    if (process.platform !== 'darwin') app.quit();
});

app.on('quit', () => {
    if (flaskProcess) {
        flaskProcess.kill();
    }
});
