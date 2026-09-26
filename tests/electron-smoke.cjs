// Run with: electron tests/electron-smoke.cjs [optional workbook path]
// Uses an isolated profile and an in-memory database; never writes to Supabase.
const { app, session } = require('electron');
const fs = require('node:fs');
const path = require('node:path');
const assert = require('node:assert/strict');
const root = path.resolve(__dirname, '..');
const testDir = path.join(root, 'work', 'electron-smoke');
fs.mkdirSync(testDir, { recursive: true });
app.setPath('userData', testDir);
const errors = [];
const timeout = setTimeout(() => { console.error('Electron smoke test timed out'); app.exit(1); }, 45000);
app.whenReady().then(() => {
  session.defaultSession.webRequest.onBeforeRequest({ urls: ['http://*/*', 'https://*/*'] }, (_details, callback) => callback({ cancel: true }));
});
app.on('browser-window-created', (_event, window) => {
  window.webContents.on('console-message', (event) => {
    if (event.level === 'error') errors.push(event.message);
  });
  window.webContents.on('preload-error', (_event, _file, error) => errors.push(error.message));
  window.webContents.on('render-process-gone', (_event, details) => errors.push(JSON.stringify(details)));
  window.webContents.once('did-finish-load', async () => {
    try {
      const source = process.argv[2];
      const base64 = source ? fs.readFileSync(source).toString('base64') : null;
      const result = await window.webContents.executeJavaScript(`(async () => {
        const check = (value, message) => { if (!value) throw new Error(message); };
        check(!document.querySelector('#supabaseSettings'), 'No connection settings UI');
        check(getConnectionSettings().url === 'https://lehhamzewsrayguaxpwf.supabase.co', 'Built-in connection');
        check(getConnectionSettings().accessToken === '', 'No login needed');
        check(document.title === 'Product cost Management App 4.1.5', 'Window title');
        check(document.querySelector('.app-version').textContent === 'Version 4.1.5', 'Version label');
        check(typeof window.desktopFile?.saveExcel === 'function', 'File preload bridge');
        check(typeof window.desktopExchange?.getFirstUsdExchangeRate === 'function', 'Exchange preload bridge');
        document.querySelector('#tabLookup').click();
        check(!document.querySelector('#lookupWorkspace').hidden, 'Lookup tab');
        document.querySelector('#tabGenerate').click();
        check(!document.querySelector('#costWorkspace').hidden, 'Generate tab');
        document.querySelector('#tabUpload').click();
        check(!document.querySelector('#uploadWorkspace').hidden, 'Upload tab');
        let bytes;
        if (${JSON.stringify(base64)}) {
          bytes = Uint8Array.from(atob(${JSON.stringify(base64)}), char => char.charCodeAt(0));
        } else {
          const book = XLSX.utils.book_new();
          XLSX.utils.book_append_sheet(book, XLSX.utils.aoa_to_sheet([
            ['품목','담당자','주거래처','단위','원화단가'],
            ['wbk0-00a1a','김주헌','예일산업','KRW',19740],
            ['pmetrcet1308a','이다해','미르전자공업','KRW',1100],
            ['wbk0-0092a','김주헌','예일산업','KRW',2760],
            ['test-zero','담당자','거래처','KRW',0]
          ]), 'Sheet1');
          bytes = XLSX.write(book, {type:'array',bookType:'xlsx'});
        }
        const file = new File([bytes], 'upload-test.xlsx');
        const excel = await readFirstSheet(file);
        const rows = excel.rows.map((entry, index) => ({
          id:index+1, 품목:entry.item.toUpperCase(), 품목명:'Preserved name',
          담당자:'Before', 거래처명:'Before', 단위:'USD', 원화단가:-1, 달러단가:99
        }));
        let patches = 0;
        window.fetch = async (url, options = {}) => {
          if (options.method === 'PATCH') {
            const id = Number(new URL(url).searchParams.get('id').slice(3));
            const row = rows.find(row => row.id === id);
            const payload = JSON.parse(options.body);
            check(!('달러단가' in payload), 'Omitted dollar price must be preserved');
            check(Object.values(payload).every(value => value !== null), 'No null writes');
            Object.assign(row, payload);
            patches++;
            return new Response(JSON.stringify([row]), {status:200});
          }
          check(!options.method || options.method === 'GET', 'Unexpected mutation');
          return new Response(JSON.stringify(rows), {status:200});
        };
        await selectUploadFile(file);
        check(patches === excel.rows.length, 'All input rows updated');
        check(!document.querySelector('#uploadResult').classList.contains('error'), document.querySelector('#uploadResult').textContent);
        rows.forEach((row,index) => {
          check(row.달러단가 === 99 && row.품목명 === 'Preserved name', 'Original values preserved');
          check(row.거래처명 === excel.rows[index].row.주거래처, 'Vendor alias updated');
          check(row.원화단가 === excel.rows[index].row.원화단가, 'Price updated');
        });
        return { title: document.title, rows: rows.length, patches,
          message: document.querySelector('#uploadResult').textContent };
      })()`);
      assert.deepEqual(errors, [], 'Renderer errors');
      fs.writeFileSync(path.join(testDir, 'result.json'), JSON.stringify({ ...result, errors }, null, 2));
      try { fs.writeFileSync(path.join(testDir, 'screenshot.png'), (await window.webContents.capturePage()).toPNG()); }
      catch (error) { console.warn('SCREENSHOT_UNAVAILABLE', error.message); }
      console.log('ELECTRON_SMOKE_PASS', JSON.stringify(result));
      clearTimeout(timeout);
      app.exit(0);
    } catch (error) {
      console.error('ELECTRON_SMOKE_FAIL', error, errors);
      clearTimeout(timeout);
      app.exit(1);
    }
  });
});
require('../main.cjs');
