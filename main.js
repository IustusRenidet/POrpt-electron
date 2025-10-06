const { app, BrowserWindow, ipcMain, dialog } = require('electron');
const path = require('path');
const { init } = require('./server');

let mainWindow;

function createWindow() {
  mainWindow = new BrowserWindow({
    width: 1200,
    height: 800,
    webPreferences: {
      nodeIntegration: false,
      contextIsolation: true,
      preload: path.join(__dirname, 'preload.js')
    }
  });
  mainWindow.loadURL('http://localhost:3000');
}

ipcMain.handle('dialog:select-path', async (event, options = {}) => {
  const senderWindow = BrowserWindow.fromWebContents(event.sender);
  const type = options.type === 'directory' ? 'directory' : 'file';
  const properties = ['dontAddToRecent'];
  if (type === 'directory') {
    properties.push('openDirectory', 'createDirectory');
  } else {
    properties.push('openFile');
  }
  if (options.multiSelections) {
    properties.push('multiSelections');
  }
  const dialogOptions = {
    title: options.title || (type === 'directory' ? 'Selecciona una carpeta' : 'Selecciona un archivo'),
    defaultPath: options.defaultPath,
    filters: Array.isArray(options.filters) ? options.filters : undefined,
    properties
  };
  const result = await dialog.showOpenDialog(senderWindow || mainWindow, dialogOptions);
  if (result.canceled || !result.filePaths?.length) {
    return null;
  }
  return result.filePaths[0];
});

app.whenReady().then(async () => {
  await init();
  createWindow();
});

app.on('window-all-closed', () => {
  require('./server').shutdown();
  if (process.platform !== 'darwin') app.quit();
});

app.on('activate', () => {
  if (BrowserWindow.getAllWindows().length === 0) createWindow();
});