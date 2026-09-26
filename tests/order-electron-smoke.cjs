const { app, session } = require('electron');
const fs = require('node:fs');
const path = require('node:path');
const root = path.resolve(__dirname, '..');
const dir = path.join(root, 'work', 'order-smoke');
fs.mkdirSync(dir, { recursive: true });
app.setPath('userData', dir);
const errors = [];
const timeout = setTimeout(() => { console.error('Timed out'); app.exit(1); }, 45000);
app.whenReady().then(() => session.defaultSession.webRequest.onBeforeRequest(
  { urls: ['http://*/*', 'https://*/*'] }, (_request, callback) => callback({ cancel: true })));
app.on('browser-window-created', (_event, window) => {
  window.webContents.on('console-message', event => { if (event.level === 'error') errors.push(event.message); });
  window.webContents.on('preload-error', (_event, _file, error) => errors.push(error.message));
  window.webContents.once('did-finish-load', async () => {
    try {
      const result = await window.webContents.executeJavaScript(`(async () => {
        const check = (value, message) => { if (!value) throw new Error(message); };
        check(!document.querySelector('#supabaseSettings'), 'No connection settings UI');
        check(getConnectionSettings().url === 'https://lehhamzewsrayguaxpwf.supabase.co', 'Built-in connection');
        check(getConnectionSettings().accessToken === '', 'No login needed');
        check(document.querySelector('#tabLookup').textContent === '조회', 'Main label');
        document.querySelector('#tabLookup').click();
        check(document.querySelector('#lookupWorkspace h2').textContent === '품목 조회', 'Item label');
        const rows = [
          {품목:'WBK0-00A1A',품목명:'MULTI_PORT_COVER',프로젝트명:'별내선',세부프로젝트명:'발매기',BOM:'부품 A × 2\\n부품 B × 1'},
          {품목:'WBK0-00A1A',품목명:'MULTI_PORT_COVER',프로젝트명:'다른 프로젝트',세부프로젝트명:'충전기',BOM:null},
          {품목:'PMETRCET1308A',품목명:'PCB FRONT EJECTOR LATCH',프로젝트명:'별내선',세부프로젝트명:'정산기',BOM:'PCB × 1'}
        ];
        let patchCount = 0, deleteCount = 0, deny = false;
        window.fetch = async (url, options = {}) => {
          const query = new URL(url);
          const mutation = options.method === 'PATCH' || options.method === 'DELETE';
          const filtered = query.searchParams.has('or')
            ? (Number(query.searchParams.get('offset')) ? [] : rows.filter(row => row.프로젝트명.includes('별내')))
            : rows.filter(row => ORDER_FIELDS.every(field => {
              const filter = query.searchParams.get(field);
              return filter === 'is.null' ? row[field] == null : filter === 'eq.' + row[field];
            }));
          if (mutation) {
            check(options.headers.Prefer.includes('max-affected=1'), 'Atomic mutation limit');
            if (deny) return new Response('[]', {status:200});
            check(filtered.length === 1, 'Mutation matches one record');
            if (options.method === 'PATCH') { patchCount++; Object.assign(filtered[0], JSON.parse(options.body)); }
            else { deleteCount++; rows.splice(rows.indexOf(filtered[0]),1); }
          }
          return new Response(JSON.stringify(filtered), {status:200});
        };
        document.querySelector('#orderLookupKeyword').value = '별내';
        await searchOrders();
        check(document.querySelectorAll('#orderCandidateBody tr').length === 2, 'Partial match candidates');
        document.querySelector('#orderCandidateBody button').click();
        check(!document.querySelector('#orderDetail').hidden, 'Selected detail visible');
        check(document.querySelectorAll('#orderDetailFields label').length === 5, 'All fields displayed');
        check(document.querySelector('#orderDetailFields').textContent.includes('부품 B'), 'Full multiline BOM');
        document.querySelector('#orderEdit').click();
        document.querySelector('[data-order-field="BOM"]').value = '수정된 BOM';
        await saveSelectedOrder();
        check(patchCount === 1 && rows[0].BOM === '수정된 BOM', 'Save selected record');
        check(rows[1].BOM === null, 'Other project with same item preserved');
        document.querySelector('#orderEdit').click();
        document.querySelector('[data-order-field="품목명"]').value = 'Discard me';
        document.querySelector('#orderCancel').click();
        check(rows[0].품목명 === 'MULTI_PORT_COVER', 'Cancel edit');
        window.confirm = () => false;
        await deleteSelectedOrder();
        check(deleteCount === 0, 'Cancel deletion');
        window.confirm = () => true;
        deny = true;
        await deleteSelectedOrder();
        check(document.querySelector('#orderLookupMessage').classList.contains('error'), 'Permission failure shown');
        check(!document.querySelector('#orderDetail').hidden && rows.length === 3, 'Denied deletion preserves selection');
        deny = false;
        await deleteSelectedOrder();
        check(deleteCount === 1 && rows.length === 2, 'Delete one selected record');
        check(rows.some(row => row.프로젝트명 === '다른 프로젝트'), 'Other project retained');
        await searchOrders();
        document.querySelector('#orderCandidateBody button').click();
        document.querySelector('#orderDetail').scrollIntoView({block:'end'});
        await new Promise(resolve => setTimeout(resolve, 250));
        return {partialSearch:true, selection:true, edit:true, cancel:true, deniedDelete:true, delete:true, patchCount, deleteCount};
      })()`);
      if (errors.length) throw new Error(JSON.stringify(errors));
      try { fs.writeFileSync(path.join(dir, 'screenshot.png'), (await window.webContents.capturePage()).toPNG()); }
      catch (error) { console.warn('SCREENSHOT_UNAVAILABLE', error.message); }
      fs.writeFileSync(path.join(dir, 'result.json'), JSON.stringify(result, null, 2));
      console.log('ORDER_ELECTRON_PASS', JSON.stringify(result));
      clearTimeout(timeout); app.exit(0);
    } catch (error) { console.error(error, errors); clearTimeout(timeout); app.exit(1); }
  });
});
require('../main.cjs');
