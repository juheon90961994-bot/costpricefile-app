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
   document.querySelector('#tabCompetitiveness').click();
   supabaseRequest=async (_settings,url)=>Number(new URL(url).searchParams.get('offset')||0)===0 ? [{'품목':'MODEL','품목명':'모델','프로젝트명':'별내선','세부프로젝트명':'테스트',BOM:'있음'},{'품목':'MODEL2','품목명':'모델2','프로젝트명':'별내선','세부프로젝트명':'테스트',BOM:'있음'}] : [];
   document.querySelector('#competitionOrderLookupKeyword').value='MODEL';
   document.querySelector('#competitionOrderLookupForm').dispatchEvent(new Event('submit',{cancelable:true}));
   for(let i=0;i<100&&!document.querySelector('#competitionOrderCandidateBody button');i++)await new Promise(r=>setTimeout(r,20));
   check(document.querySelectorAll('#competitionOrderCandidateBody tr').length===2,'Separate search results');
   document.querySelector('#competitionOrderCandidateBody button').click();
   document.querySelector('#competitionOrderGenerateCost').click();
   for(let i=0;i<200&&window.competitionLookup.isBusy();i++)await new Promise(r=>setTimeout(r,20));
   check(document.querySelector('#competitionOrderLookupMessage').textContent.includes('저장 완료'),'Competition generation button');
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
   const rules=XLSX.utils.book_new();XLSX.utils.book_append_sheet(rules,XLSX.utils.aoa_to_sheet([
   ['단계','품목','품목명','규격','누적실소요량','주거래처','BOM구분','계정구분','현재고','단위','달러단가'],
   [1,'USD','USD 자재','',2,'','사급','자재',0,'USD',10],
   [1,'USD','소모품','',3,'','사급',' 소모품 ',0,'USD',10],
   [1,'NONE','단가없는 소모품','',null,'','사급','소모품',0,'KRW',null],
   [1,'USD','도급','',1,'','도급','자재',0,'USD',10]
   ]),'BOM');
   window.costDatabase.getUnitPriceMap=async()=>({prices:new Map([['USD',500]]),records:new Map()});
   const rulesFile=new File([XLSX.write(rules,{type:'array',bookType:'xlsx'})],'MODEL_규칙.xlsx');
   for(const result of [await buildProductCost(rulesFile,{competitiveness:true}),await buildCostFiles([rulesFile,rulesFile],undefined,()=>{},{competitiveness:true})]){
     const book=new ExcelJS.Workbook();await book.xlsx.load(result.data);
     for(const ws of book.worksheets){check(ws.getCell('N4').value.result===1000&&ws.getCell('N4').value.formula==='E4*M4','USD uses KRW');for(const row of [5,6,7])check(ws.getCell('N'+row).value===0,'Excluded rows zero');check(ws.getCell('N8').value.result===1000,'Total');}
   }
   const normal=await buildProductCost(rulesFile);check(normal.workbook.worksheets[0].getCell('N4').value.result===26000,'Existing USD behavior unchanged');
   available=false;let failed=false;try{await ExcelArchive.readForItem('MODEL')}catch{failed=true}check(failed,'Missing file blocked');
   available=true;duplicate=true;failed=false;try{await ExcelArchive.readForItem('MODEL')}catch{failed=true}check(failed,'Duplicate blocked');
   return 'PASS: lookup button → archive → shared generation → save; matching outputs, formats, formulas, missing/duplicate files';
  })()`));
  if(!saved || saved.defaultName!=='MODEL_별내선_제품원가.xlsx')throw Error('Save filename mismatch');
  app.exit(0);
 }catch(error){console.error(error);app.exit(1)}
}));
require('../main.cjs');
