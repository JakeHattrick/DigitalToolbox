const { app, BrowserWindow, ipcMain, dialog, shell } = require("electron");
const { fork } = require('child_process');
const path = require("path");

let mainWindow;

// Register IPC handlers before app.whenReady()
ipcMain.handle('select-output-folder', async () => {
  try {
    const result = await dialog.showOpenDialog(mainWindow, {
      properties: ['openDirectory'],
      title: 'Select Output Folder'
    });
    return result;
  } catch (error) {
    return { canceled: true, error: error.message };
  }
});

ipcMain.handle('select-files', async () => {
  try {
    const result = await dialog.showOpenDialog(mainWindow, {
      properties: ['openFile', 'multiSelections'],
      filters: [
        { name: 'Word Documents', extensions: ['docx'] },
        { name: 'All Files', extensions: ['*'] }
      ],
      title: 'Select DOCX Files'
    });
    return result;
  } catch (error) {
    return { canceled: true, error: error.message };
  }
});

// Add handler to get Downloads path
ipcMain.handle('get-downloads-path', async () => {
  const os = require('os');
  return path.join(os.homedir(), 'Downloads');
});

// Add handler to open Downloads folder
ipcMain.handle('open-downloads-folder', async () => {
  try {
    const os = require('os');
    const downloadsPath = path.join(os.homedir(), 'Downloads');
    await shell.openPath(downloadsPath);
    return { success: true };
  } catch (error) {
    return { success: false, error: error.message };
  }
});

ipcMain.handle('show-message-box', async (event, options) => {
  try {
    const result = await dialog.showMessageBox(mainWindow, options);
    return result;
  } catch (error) {
    return { response: 1, error: error.message };
  }
});

ipcMain.handle('open-folder', async (event, folderPath) => {
  try {
    await shell.openPath(folderPath);
    return { success: true };
  } catch (error) {
    return { success: false, error: error.message };
  }
});

ipcMain.handle('convert-docx-to-pdf-winax', async (event, payload) => {
  const workerPath = path.join(__dirname, 'wordWorker.js'); // adjust if your worker lives elsewhere
  return new Promise((resolve) => {
    const child = fork(workerPath, [], { stdio: ['inherit', 'inherit', 'inherit', 'ipc'] });

    const timer = setTimeout(() => {
      try { child.kill('SIGKILL'); } catch {}
      resolve({ success: false, error: 'Conversion timed out' });
    }, 5 * 60 * 1000); // 5 min safety

    child.on('message', (msg) => {
      // msg can be {progress:n} or final {success:true/false,...}
      if (msg && typeof msg === 'object' && 'progress' in msg) {
        if (!event.sender.isDestroyed()) {
          event.sender.send('docx-to-pdf-progress', msg.progress);
        }
        return; // don't resolve yet
      }
      clearTimeout(timer);
      resolve(msg);
      try { child.disconnect(); child.kill(); } catch {}
    });

    child.on('exit', (code) => {
      if (code !== 0) {
        clearTimeout(timer);
        resolve({ success: false, error: `Worker exited with code ${code}` });
      }
    });

    child.send(payload);
  });
});


function createWindow() {
  // Corrected preload path for electron folder structure
  const preloadPath = path.join(__dirname, "preload.js");
  console.log('Preload path:', preloadPath);
  console.log('Preload file exists:', require('fs').existsSync(preloadPath));

  mainWindow = new BrowserWindow({
    width: 1200,
    height: 800,
    webPreferences: {
      contextIsolation: true,
      nodeIntegration: false,
      sandbox:false,
      preload: preloadPath,
    },
  });

  const startUrl = process.env.ELECTRON_START_URL;
  if (startUrl) {
    // Dev: use CRA server when ELECTRON_START_URL is provided by your "dev" script
    mainWindow.loadURL(startUrl);
    mainWindow.webContents.openDevTools({ mode: "detach" });
  } else {
    // Prod/local-prod: load the built files
    // Updated path since main.js is in electron folder
    const indexPath = path.join(__dirname, "../build/index.html");
    console.log("Loading file from:", indexPath);
    mainWindow.loadFile(indexPath);
    // Open devtools to see what errors are occurring
    mainWindow.webContents.openDevTools({ mode: "detach" });
  }

  // Enhanced diagnostics
  mainWindow.webContents.on("did-fail-load", (_e, code, desc, url) =>
    console.error("did-fail-load:", code, desc, url)
  );
  
  mainWindow.webContents.on("did-finish-load", () => {
    console.log("Window finished loading");
  });

  mainWindow.webContents.on("console-message", (event, level, message) => {
    console.log("Renderer console:", message);
  });
}

app.whenReady().then(createWindow);
app.on("window-all-closed", () => { 
  if (process.platform !== "darwin") app.quit(); 
});
app.on("activate", () => { 
  if (BrowserWindow.getAllWindows().length === 0) createWindow(); 
});