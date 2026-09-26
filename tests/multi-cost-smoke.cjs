const {app,session,ipcMain}=require('electron');
const path=require('node:path');
app.setPath('userData',path.join(__dirname,'../work/order-cost-smoke'));
setTimeout(()=>app.exit(1),45000);
app.whenReady().then(()=>session.defaultSession.webRequest.onBeforeRequest({urls:['https://*/*','http://*/*']},(_,cb)=>cb({cancel:true})));
app.on('browser-window-created',(_,win)=>win.webContents.once('did-finish-load',async()=>{
 try{
  let saved;
  ipcMain.removeHandler('save-excel-file');ipcMain.handle('save-excel-file',(_,data)=>{saved=data;return {canceled:false};});
  ipcMain.removeHandler('get-first-usd-exchange-rate');ipcMain.handle('get-first-usd-exchange-rate',()=>({rate:1300}));
  console.log(await win.webContents.executeJavaScript(`(async()=>{
   const check=(ok,msg)=>{if(!ok)throw Error(msg)};
   const book=XLSX.utils.book_new();
   XLSX.utils.book_append_sheet(book,XLSX.utils.aoa_to_sheet([
    ['단계','품목','품목명','규격','누적실소요량','주거래처','BOM구분','계정구분','현재고'],
    [1,'PART','부품','규격',2,'거래처','BOM','자재',0]
   ]),'Source');
   const bytes=XLSX.write(book,{type:'array',bookType:'xlsx'});
   let available=true,duplicate=false;
   window.fetch=async(url,options)=>{
    if(url.includes('/object/list/')) {
     const {offset}=JSON.parse(options.body);
     return Response.json(available&&!offset ? [{id:'one',name:ExcelArchive.key('MODEL_별내선.xlsx').slice(6)},...(duplicate?[{id:'two',name:ExcelArchive.key('MODEL_이전.xlsx').slice(6)}]:[])] : []);
    }
    return new Response(bytes);
   };
   window.costDatabase.getUnitPriceMap=async()=>({prices:new Map([['PART',100]]),records:new Map([['PART',{'단위':'KRW'}]])});
   document.querySelector('#tabLookup').click();
   selectedOrder={'품목':'MODEL','품목명':'모델','프로젝트명':'별내선','세부프로젝트명':'테스트',BOM:'있음'};
   renderOrderDetail();
   document.querySelector('#orderGenerateCost').click();
   for(let i=0;i<200 && orderMutating;i++)await new Promise(r=>setTimeout(r,20));
   check(orderMessage.textContent.includes('저장 완료'),orderMessage.textContent);
   const fromArchive=await buildProductCost(await ExcelArchive.readForItem('MODEL'));
   const direct=await buildProductCost(new File([bytes],'MODEL_별내선.xlsx'));
   const a=new ExcelJS.Workbook(),b=new ExcelJS.Workbook();
   await a.xlsx.load(fromArchive.data);await b.xlsx.load(direct.data);
   const sheet=a.worksheets[0];
   check(JSON.stringify(sheet.getSheetValues())===JSON.stringify(b.worksheets[0].getSheetValues()),'Same output contents');
   check(sheet.getCell('A1').value==='MODEL_별내선','Title excludes generated suffix and extension');
   check(sheet.getColumn(2).width===14.5 && sheet.getColumn(3).width===35.4 && sheet.getColumn(4).hidden,'Columns');
   check(sheet.getCell('A1').font.size===15 && sheet.name==='별내선','Title and sheet name');
   check(sheet.getCell('N4').value.result===200,'Cost formula result');
   const readOriginal=ExcelArchive.readForItem;
   ExcelArchive.readForItem=async(item)=>new File([bytes],item+'_별내선.xlsx');
   const multiRows=[{'품목':'MODEL','프로젝트명':'별내선','품목명':'발매기'},{'품목':'MODEL2','프로젝트명':'별내선','품목명':'개집표기'}];
   const multi=await buildMultiCost(multiRows);
   const merged=new ExcelJS.Workbook();await merged.xlsx.load(multi.data);
   check(merged.worksheets.length===2,'Selected count matches sheet count');
   check(merged.worksheets[0].name==='별내선' && merged.worksheets[1].name==='별내선 (2)','Unique sheet names');
   for(const ws of merged.worksheets){
    check(ws.getCell('N4').value.result===200 && ws.getCell('N4').value.formula==='E4*M4','Formula preserved');
    check(ws.getColumn(3).width===35.4 && ws.getColumn(4).hidden && ws.getCell('A1').font.size===15,'Formatting preserved');
    check(ws.autoFilter && ws.views[0].state==='frozen','Filter and frozen panes preserved');
   }
   check(multi.name==='별내선_멀티원가분석.xlsx','Common word filename');
   orderRows=[multiRows[0]];renderOrderCandidates();check(document.querySelector('#orderMultiToggle').hidden,'Single result hides multi');
   orderRows=multiRows;renderOrderCandidates();check(!document.querySelector('#orderMultiToggle').hidden,'Multiple results show multi');
   document.querySelector('#orderMultiToggle').click();
   for(let i=0;i<2;i++){const box=document.querySelectorAll('#orderCandidateBody input')[i];box.checked=true;box.dispatchEvent(new Event('change'));}
   check(orderMultiSelected.size===2,'Checkboxes select two');
   document.querySelector('#orderMultiGenerate').click();
   for(let i=0;i<200 && orderMutating;i++)await new Promise(r=>setTimeout(r,20));
   check(orderMessage.textContent.includes('2개 시트 저장 완료'),orderMessage.textContent);
   ExcelArchive.readForItem=readOriginal;
   available=false;let failed=false;try{await ExcelArchive.readForItem('MODEL')}catch{failed=true}check(failed,'Missing file blocked');
   available=true;duplicate=true;failed=false;try{await ExcelArchive.readForItem('MODEL')}catch{failed=true}check(failed,'Duplicate blocked');
   return 'PASS: multi selection UI, multi-sheet saved file, unique names, common filename, formats and formulas';
  })()`));
  if(!saved || saved.defaultName!=='별내선_멀티원가분석.xlsx')throw Error('Save filename mismatch');
  app.exit(0);
 }catch(error){console.error(error);app.exit(1)}
}));
require('../main.cjs');
