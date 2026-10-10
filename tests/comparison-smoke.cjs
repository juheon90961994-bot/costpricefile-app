const {app,session,ipcMain}=require('electron'),fs=require('fs'),path=require('path');
app.setPath('userData',path.join(__dirname,'../work/comparison-test'));setTimeout(()=>app.exit(1),60000);
app.whenReady().then(()=>session.defaultSession.webRequest.onBeforeRequest({urls:['http://*/*','https://*/*']},(_,cb)=>cb({cancel:true})));
app.on('browser-window-created',(_,win)=>win.webContents.once('did-finish-load',async()=>{try{
 const files=process.argv.slice(2).map(p=>fs.readFileSync(p).toString('base64'));let saved;
 ipcMain.removeHandler('save-excel-file');ipcMain.handle('save-excel-file',(_,data)=>{saved=data;return {canceled:false};});
 console.log(await win.webContents.executeJavaScript(`(async()=>{
 const check=(ok,msg)=>{if(!ok)throw Error(msg)};
 const encoded=${JSON.stringify(files)},make=(i,name)=>new File([Uint8Array.from(atob(encoded[i]),c=>c.charCodeAt(0))],name);
 const erp=make(0,'원가경쟁력.xlsx'),cost=make(1,'전자칠판_A8GC_멀티원가분석.xlsx');
 const result=await buildCostComparison(erp,cost);const out=new ExcelJS.Workbook();await out.xlsx.load(result.data);
 const sheet=out.getWorksheet('비교표');
 const original=new ExcelJS.Workbook();await original.xlsx.load(await cost.arrayBuffer());const oldComparison=original.getWorksheet('비교표');if(oldComparison)original.removeWorksheet(oldComparison.id);
 check(out.worksheets.length===original.worksheets.length+1,'Comparison appended');
 check(result.name===cost.name,'Original filename retained');
 for(const source of original.worksheets){const copy=out.getWorksheet(source.name);check(JSON.stringify(copy.getSheetValues())===JSON.stringify(source.getSheetValues()),'Existing data and formulas preserved');check(copy.getColumn(2).width===source.getColumn(2).width&&copy.getColumn(4).hidden===source.getColumn(4).hidden,'Existing columns preserved');check(JSON.stringify(copy.views)===JSON.stringify(source.views),'Existing panes preserved');}
 original.views=[];
 const withoutViews=new File([await original.xlsx.writeBuffer()],cost.name);
 const noViewResult=await buildCostComparison(erp,withoutViews);
 const noViewOutput=new ExcelJS.Workbook();await noViewOutput.xlsx.load(noViewResult.data);
 check(noViewOutput.worksheets[0].name==='비교표'&&noViewOutput.views[0].activeTab===0,'Missing workbook views handled');
 const repeated=await buildCostComparison(erp,new File([result.data],cost.name));
 const again=new ExcelJS.Workbook();await again.xlsx.load(repeated.data);
 check(again.worksheets.length===out.worksheets.length&&repeated.count===9,'Replace comparison on repeated upload');check(out.worksheets[0].name==='비교표'&&again.worksheets[0].name==='비교표','First sheet on creation and replacement');check(JSON.stringify(out.worksheets.slice(1).map(s=>s.name))===JSON.stringify(original.worksheets.map(s=>s.name)),'Original order preserved');check(result.count===9&&result.warnings===0,'All nine sheets matched');
 check(sheet.getCell('B4').value==='A8GC-1XTNT-AH000'&&sheet.getCell('C4').value==='C1 1XTNT','A1 identity');
 check(sheet.getCell('D4').value===2314306,'ERP K value');
 check(Math.abs(sheet.getCell('E4').value-2307037.8856)<0.00001,'Cost total');
 check(sheet.getCell('F4').value.formula==='D4-E4'&&Math.round(sheet.getCell('F4').value.result)===7268,'Difference formula');
 const erpBook=XLSX.read(await erp.arrayBuffer(),{type:'array'});erpBook.Sheets['변경본'].A133.v='OTHER';
 const missing=await buildCostComparison(new File([XLSX.write(erpBook,{type:'array',bookType:'xlsx'})],'erp.xlsx'),cost);
 check(missing.warnings===1,'Missing item flagged');
 document.querySelector('#tabCompetitiveness').click();document.querySelector('#comparisonOpen').click();
 const dialog=document.querySelector('dialog'),inputs=dialog.querySelectorAll('input');
 for(const [i,file]of [erp,cost].entries()){const transfer=new DataTransfer();transfer.items.add(file);inputs[i].files=transfer.files;}
 dialog.querySelector('form').requestSubmit();for(let i=0;i<250&&comparisonBusy;i++)await new Promise(r=>setTimeout(r,20));
 check(dialog.textContent.includes('9개 품목 비교표 저장 완료'),'Dialog upload to save');
 return 'PASS: nine sheets, A1 split, ERP K, cost total, difference, missing match and upload-save';
 })()`));if(!saved?.data?.length)throw Error('No save');app.exit(0);
 }catch(e){console.error(e);app.exit(1)}}));require('../main.cjs');
