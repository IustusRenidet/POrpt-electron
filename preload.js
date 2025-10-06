const { contextBridge, ipcRenderer } = require('electron');

contextBridge.exposeInMainWorld('electronAPI', {
  selectPath: async options => ipcRenderer.invoke('dialog:select-path', options || {})
});
