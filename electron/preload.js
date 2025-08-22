// electron/preload.js
const { contextBridge } = require("electron");

contextBridge.exposeInMainWorld("electronAPI", {
  // add safe, explicit APIs here if you need them later
});
