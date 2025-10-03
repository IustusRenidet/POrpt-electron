const { app, BrowserWindow } = require('electron');
const path = require('path');
const { init } = require('./server');

let mainWindow;

function createWindow() {
  mainWindow = new BrowserWindow({
    width: 1200, height: 800,
    webPreferences: { nodeIntegration: false, contextIsolation: true, preload: path.join(__dirname, 'preload.js') }
  });
  mainWindow.loadURL('http://localhost:3000');  // Apunta a Express
}

app.whenReady().then(async () => {
  await init();  // Inicia Express + SQLite
  createWindow();
});

app.on('window-all-closed', () => {
  require('./server').shutdown();  // Cierra DBs
  if (process.platform !== 'darwin') app.quit();
});