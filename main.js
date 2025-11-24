const { app, BrowserWindow, ipcMain, dialog } = require('electron');
const fs = require('fs');
const path = require('path');

let mainWindow;
let serverModule;

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

  if (!app.isPackaged) {
    // Abre las DevTools automáticamente en modo desarrollo para depurar login/render.
    mainWindow.webContents.openDevTools({ mode: 'detach' });
  }
}

async function fileExists(filePath) {
  try {
    await fs.promises.access(filePath, fs.constants.F_OK);
    return true;
  } catch {
    return false;
  }
}

async function prepareDatabasePath() {
  if (!app.isPackaged) {
    return path.join(__dirname, 'PERFILES.DB');
  }

  const userDataDir = app.getPath('userData');
  const destination = path.join(userDataDir, 'PERFILES.DB');
  await fs.promises.mkdir(userDataDir, { recursive: true });

  if (await fileExists(destination)) {
    return destination;
  }

  const bundledPath = path.join(process.resourcesPath, 'PERFILES.DB');
  if (!(await fileExists(bundledPath))) {
    throw new Error(`Archivo de base de datos no encontrado en ${bundledPath}`);
  }

  await fs.promises.copyFile(bundledPath, destination);
  return destination;
}

async function initializeApplication() {
  try {
    const sqliteDbPath = await prepareDatabasePath();
    process.env.SQLITE_DB = sqliteDbPath;
    serverModule = require('./server');
    await serverModule.init();
    createWindow();
  } catch (error) {
    console.error('Error inicializando la aplicación:', error);
    dialog.showErrorBox('Error crítico', `No se pudo iniciar la aplicación. ${error.message}`);
    app.quit();
  }
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

app.whenReady().then(initializeApplication);

app.on('window-all-closed', async () => {
  if (serverModule?.shutdown) {
    try {
      await serverModule.shutdown();
    } catch (error) {
      console.error('Error al cerrar el servidor:', error);
    }
  }
  if (process.platform !== 'darwin') app.quit();
});

app.on('activate', () => {
  if (BrowserWindow.getAllWindows().length === 0) createWindow();
});
