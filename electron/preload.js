// /mnt/data/preload.js
const { contextBridge, ipcRenderer } = require('electron');
// IMPORTANT: require Node modules at top-level in preload (Node is available here)
const path = require('path');

contextBridge.exposeInMainWorld('electronAPI', {
  // File selection / output helpers
  selectFiles: () => ipcRenderer.invoke('select-files'),
  selectOutputFolder: () => ipcRenderer.invoke('select-output-folder'),

  // Paths
  getDownloadsPath: () => ipcRenderer.invoke('get-downloads-path'),
  openDownloadsFolder: () => ipcRenderer.invoke('open-downloads-folder'),

  // Word/WinAX conversion
  convertDocxToPdf: (options) => ipcRenderer.invoke('convert-docx-to-pdf-winax', options),
  onDocxToPdfProgress: (handler) => {
    if (typeof handler !== 'function') return () => {};
    const listener = (_event, progress) => handler(progress);
    ipcRenderer.on('docx-to-pdf-progress', listener);
    return () => ipcRenderer.removeListener('docx-to-pdf-progress', listener);
  },

  // Dialogs / Shell
  showMessageBox: (options) => ipcRenderer.invoke('show-message-box', options),
  openFolder: (folderPath) => ipcRenderer.invoke('open-folder', folderPath),

  // Minimal, safe path helpers for the renderer
  path: {
    join: (...parts) => path.join(...parts),
    basename: (p, ext) => path.basename(p, ext),
    dirname: (p) => path.dirname(p),
    extname: (p) => path.extname(p),
  },

  // Optional environment info (read-only)
  platform: process.platform,
  versions: {
    node: process.versions.node,
    chrome: process.versions.chrome,
    electron: process.versions.electron,
  }
});
