const { app, BrowserWindow } = require("electron");
const path = require("path");

function createWindow() {
  const win = new BrowserWindow({
    width: 1200,
    height: 800,
    webPreferences: {
      contextIsolation: true,
      preload: path.join(__dirname, "preload.js"),
    },
  });

  const startUrl = process.env.ELECTRON_START_URL;
  if (startUrl) {
    // Dev: use CRA server when ELECTRON_START_URL is provided by your "dev" script
    win.loadURL(startUrl);
    win.webContents.openDevTools({ mode: "detach" });
  } else {
    // Prod/local-prod: load the built files
    const indexPath = path.join(__dirname, "../build/index.html");
    console.log("Loading file from:", indexPath);
    win.loadFile(indexPath);
    // Open devtools to see what errors are occurring
    win.webContents.openDevTools({ mode: "detach" });
  }

  // Enhanced diagnostics
  win.webContents.on("did-fail-load", (_e, code, desc, url) =>
    console.error("did-fail-load:", code, desc, url)
  );
  
  win.webContents.on("did-finish-load", () => {
    console.log("Window finished loading");
  });

  win.webContents.on("console-message", (event, level, message) => {
    console.log("Renderer console:", message);
  });
}

app.whenReady().then(createWindow);
app.on("window-all-closed", () => { if (process.platform !== "darwin") app.quit(); });
app.on("activate", () => { if (BrowserWindow.getAllWindows().length === 0) createWindow(); });