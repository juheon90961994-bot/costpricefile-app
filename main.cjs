const { app, BrowserWindow, dialog, ipcMain, shell } = require("electron");
const path = require("node:path");
const fs = require("node:fs/promises");
const https = require("node:https");

if (require("electron-squirrel-startup")) app.quit();

function createWindow() {
  const window = new BrowserWindow({
    width: 1280,
    height: 860,
    minWidth: 920,
    minHeight: 700,
    backgroundColor: "#f7faff",
    autoHideMenuBar: true,
    title: "Product cost Management App",
    webPreferences: {
      contextIsolation: true,
      nodeIntegration: false,
      sandbox: true,
      preload: path.join(__dirname, "preload.cjs"),
    },
  });
  window.loadFile(path.join(__dirname, "index.html"));
  window.webContents.setWindowOpenHandler(({ url }) => {
    if (/^https?:\/\//i.test(url)) shell.openExternal(url);
    return { action: "deny" };
  });
}

ipcMain.handle("save-excel-file", async (event, { data, defaultName }) => {
  const parent = BrowserWindow.fromWebContents(event.sender);
  const safeName = path.basename(defaultName || "단가표.xlsx");
  const result = await dialog.showSaveDialog(parent, {
    title: "원가 단가표 저장",
    defaultPath: path.join(app.getPath("downloads"), safeName),
    filters: [{ name: "Excel 통합 문서", extensions: ["xlsx"] }],
    properties: ["showOverwriteConfirmation", "createDirectory"],
  });
  if (result.canceled || !result.filePath) return { canceled: true };
  await fs.writeFile(result.filePath, Buffer.from(data));
  return { canceled: false, filePath: result.filePath };
});

ipcMain.handle("get-first-usd-exchange-rate", async () => {
  for (let daysAgo = 0; daysAgo < 8; daysAgo += 1) {
    const date = new Date();
    date.setDate(date.getDate() - daysAgo);
    const result = await requestFirstUsdRate(date).catch(() => null);
    if (result) return result;
  }
  throw new Error("최근 8일 이내의 하나은행 최초 공지 USD 환율을 조회하지 못했습니다.");
});

function requestFirstUsdRate(date) {
  const year = date.getFullYear();
  const month = String(date.getMonth() + 1).padStart(2, "0");
  const day = String(date.getDate()).padStart(2, "0");
  const compactDate = `${year}${month}${day}`;
  const displayDate = `${year}-${month}-${day}`;
  const body = new URLSearchParams({
    ajax: "true",
    curCd: "USD",
    tmpInqStrDt: displayDate,
    pbldDvCd: "1",
    pbldSqn: "",
    inqStrDt: compactDate,
    inqKindCd: "1",
    hid_key_data: "",
    hid_enc_data: "",
    requestTarget: "searchContentDiv",
  }).toString();

  return new Promise((resolve, reject) => {
    const request = https.request({
      hostname: "www.kebhana.com",
      path: "/cms/rate/wpfxd651_01i_01.do",
      method: "POST",
      headers: {
        "Content-Type": "application/x-www-form-urlencoded; charset=UTF-8",
        "Content-Length": Buffer.byteLength(body),
        Referer: "https://www.kebhana.com/cms/rate/index.do?contentUrl=/cms/rate/wpfxd651_01i.do",
        "User-Agent": "Mozilla/5.0",
      },
      timeout: 10000,
    }, response => {
      let html = "";
      response.setEncoding("utf8");
      response.on("data", chunk => { html += chunk; });
      response.on("end", () => {
        if (response.statusCode < 200 || response.statusCode >= 300) {
          reject(new Error(`환율 조회 실패 (${response.statusCode})`));
          return;
        }
        const usdRow = html.match(/<tr[^>]*>[\s\S]*?미국\s*USD[\s\S]*?<\/tr>/i)?.[0];
        const values = [...String(usdRow || "").matchAll(/<td[^>]*class=["'][^"']*txtAr[^"']*["'][^>]*>\s*([\d,.+-]+)\s*<\/td>/gi)]
          .map(match => Number(match[1].replace(/,/g, "")));
        const rate = values[7];
        if (!Number.isFinite(rate) || rate <= 0) {
          resolve(null);
          return;
        }
        resolve({ rate, date: compactDate, source: "하나은행 최초 공지 USD 매매기준율" });
      });
    });
    request.on("timeout", () => request.destroy(new Error("환율 조회 시간이 초과되었습니다.")));
    request.on("error", reject);
    request.end(body);
  });
}

app.whenReady().then(() => {
  createWindow();
  app.on("activate", () => {
    if (BrowserWindow.getAllWindows().length === 0) createWindow();
  });
});
app.on("window-all-closed", () => {
  if (process.platform !== "darwin") app.quit();
});
