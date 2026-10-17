const {app,session,ipcMain}=require('electron');
const fs=require('node:fs'),path=require('node:path');
app.setPath('userData',path.join(__dirname,'../work/representative-batch-'+process.pid));
setTimeout(()=>app.exit(1),60000);
app.whenReady().then(()=>session.defaultSession.webRequest.onBeforeRequest({urls:['http://*/*','https://*/*']},(_,cb)=>cb({cancel:true})));
app.on('browser-window-created',(_,win)=>win.webContents.once('did-finish-load',async()=>{
  try{
    let saved;
    ipcMain.removeHandler('save-excel-file');ipcMain.handle('save-excel-file',(_,payload)=>{saved=payload;return {canceled:false};});
    const bytes=fs.readFileSync(path.join(__dirname,'../work/representative-reference.xlsx')).toString('base64');
    console.log(await win.webContents.executeJavaScript(`(async()=>{
      const check=(ok,message)=>{if(!ok)throw Error(message)};
      const bytes=Uint8Array.from(atob(${JSON.stringify(bytes)}),c=>c.charCodeAt(0));
      ExcelArchive.list=async()=>[{filename:'MODEL-A_BOM.xlsx'},{filename:'MODEL-B_BOM.xlsx'}];
      ExcelArchive.readForItem=async item=>new File([bytes],item+'_BOM.xlsx');
      let calls=0;window.costDatabase.getUnitPriceMap=async()=>{calls++;return {records:new Map()};};
      const settle=async()=>{for(let i=0;i<1000&&window.representativeModel.isBusy();i++)await new Promise(r=>setTimeout(r,20));};
      document.getElementById('tabCompetitiveness').click();
      document.getElementById('representativeSearch').click();await settle();
      const boxes=document.querySelectorAll('#representativeResults input');check(boxes.length===2,'Two archive choices');
      for(const box of boxes)box.click();
      document.getElementById('representativeForm').requestSubmit();await settle();
      check(document.getElementById('representativeMessage').textContent.includes('2개 시트 · 56개 품목 저장 완료'),'Batch saved');
      check(calls===1,'One shared price snapshot');
      document.getElementById('tabLookup').click();
      check(!document.querySelector('#representativeResults input'),'Selection resets on navigation');
      return 'PASS: archive multi-selection, one price snapshot, batch save and reset';
    })()`));
    if(!saved?.data?.length)throw Error('Missing saved workbook');
    const ExcelJS=require('../exceljs.min.js'),book=new ExcelJS.Workbook();await book.xlsx.load(Buffer.from(saved.data));
    if(book.worksheets.map(s=>s.name).join(',')!=='MODEL-A,MODEL-B')throw Error('Incorrect sheet names');
    console.log('PASS: saved workbook contains MODEL-A and MODEL-B');app.exit(0);
  }catch(error){console.error(error);app.exit(1);}
}));
require('../main.cjs');
