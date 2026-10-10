const { contextBridge, ipcRenderer } = require("electron");
contextBridge.exposeInMainWorld('desktopSales', {
  lookup: model => ipcRenderer.invoke('lookup-g2b-sale-price', model),
});

contextBridge.exposeInMainWorld("desktopFile", {
  convertLegacyExcel: data => ipcRenderer.invoke('convert-legacy-excel', data),
  saveExcel: (data, defaultName) => ipcRenderer.invoke("save-excel-file", { data, defaultName }),
});

contextBridge.exposeInMainWorld("desktopExchange", {
  getFirstUsdExchangeRate: () => ipcRenderer.invoke("get-first-usd-exchange-rate"),
});
