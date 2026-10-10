const {app,session,ipcMain}=require('electron'),fs=require('node:fs'),path=require('node:path');
app.setPath('userData',path.join(__dirname,'../work/competitiveness-smoke'));
const timer=setTimeout(()=>app.exit(1),60000);
app.whenReady().then(()=>session.defaultSession.webRequest.onBeforeRequest({urls:['https://*/*','http://*/*']},(_,cb)=>cb({cancel:true})));
app.on('browser-window-created',(_,win)=>win.webContents.once('did-finish-load',async()=>{
 try{
  const data=fs.readFileSync(process.argv[2]).toString('base64');let saved;
  ipcMain.removeHandler('save-excel-file');ipcMain.handle('save-excel-file',(_,payload)=>{saved=payload;return {canceled:false};});
  console.log(await win.webContents.executeJavaScript(`(async()=>{
    const check=(value,message)=>{if(!value)throw Error(message)};
    const bytes=Uint8Array.from(atob(${JSON.stringify(data)}),c=>c.charCodeAt(0));
    const reference=XLSX.read(bytes,{type:'array'}),expected=XLSX.utils.sheet_to_json(reference.Sheets.Sheet2,{header:1,defval:null});
    const file=new File([bytes],'원가계산현황.xlsx');
    const result=await buildCompetitiveness(file),out=new ExcelJS.Workbook();await out.xlsx.load(result.data);
    check(out.worksheets.length===2 && out.worksheets[0].name==='원본' && out.worksheets[1].name==='변경본','Two named sheets');
    const original=out.getWorksheet('원본');
    const refBook=new ExcelJS.Workbook();await refBook.xlsx.load(bytes);
    check(JSON.stringify(original.getSheetValues())===JSON.stringify(refBook.worksheets[0].getSheetValues()),'Original values preserved');
    check(original.getCell('F5').font.name===refBook.worksheets[0].getCell('F5').font.name && original.getCell('F5').font.size===refBook.worksheets[0].getCell('F5').font.size && original.getCell('F5').alignment.vertical===refBook.worksheets[0].getCell('F5').alignment.vertical,'Original styles preserved');
    const sheet=out.getWorksheet('변경본');
    check(sheet.getCell('A4').font.name==='굴림' && sheet.getCell('A4').font.size===9,'Reference font');
    check(sheet.views[0].state==='frozen' && sheet.views[0].xSplit===2 && sheet.views[0].ySplit===3 && sheet.views[0].topLeftCell==='C4','C4 freeze panes');
    for(let r=4;r<=sheet.rowCount;r++)for(let c=10;c<=12;c++)check(sheet.getCell(r,c).numFmt.replaceAll(String.fromCharCode(92),'')===ACCOUNTING.replaceAll(String.fromCharCode(92),''),'Accounting number format '+r+':'+c+' actual='+sheet.getCell(r,c).numFmt+' expected='+ACCOUNTING);
    check(result.count===138,'Expected 138 items');
    for(let r=3;r<expected.length;r++)for(let c=0;c<12;c++){
      const actual=sheet.getCell(r+1,c+1).value;const value=actual?.formula?actual.result:actual;
      const target=expected[r][c];check(typeof target==='number'?Math.abs(value-target)<0.000001:value===target,'Mismatch '+(r+1)+':'+(c+1));
    }
    check(sheet.getCell('J4').value.formula.includes('(F4+G4+I4)/C4'),'Manufacturing formula');
    check(sheet.model.merges.length===5,'Header merges');
    const first=reference.Sheets.Sheet1;first.AD5.v=0;
    const zero=new File([XLSX.write(reference,{type:'array',bookType:'xlsx'})],'zero.xlsx');
    check((await buildCompetitiveness(zero)).warnings===1,'Zero quantity handled');
    document.querySelector('#tabCompetitiveness').click();check(!document.querySelector('#competitivenessWorkspace').hidden,'Tab visible');
    const transfer=new DataTransfer();transfer.items.add(file);const input=document.querySelector('#competitivenessInput');input.files=transfer.files;input.dispatchEvent(new Event('change'));
    for(let i=0;i<250&&competitivenessBusy;i++)await new Promise(r=>setTimeout(r,20));
    check(document.querySelector('#competitivenessMessage').textContent.includes('138개 품목 저장 완료'),'Upload converts and saves');
    document.querySelector('#tabGenerate').click();check(document.querySelector('#competitivenessWorkspace').hidden,'Other tab hides panel');
    return 'PASS: all 138 rows × 12 values match Sheet2; formulas, merges, zero quantity, tab and upload-save';
  })()`));
  if(!saved?.data?.length)throw Error('No Excel saved');clearTimeout(timer);app.exit(0);
 }catch(error){console.error(error);app.exit(1)}
}));require('../main.cjs');
