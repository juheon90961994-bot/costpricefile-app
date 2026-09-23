const { contextBridge, ipcRenderer } = require("electron");

contextBridge.exposeInMainWorld("desktopFile", {
  saveExcel: (data, defaultName) => ipcRenderer.invoke("save-excel-file", { data, defaultName }),
});

contextBridge.exposeInMainWorld("desktopExchange", {
  getFirstUsdExchangeRate: () => ipcRenderer.invoke("get-first-usd-exchange-rate"),
});
