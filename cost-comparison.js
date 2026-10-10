async function buildCostComparison(erpFile,costFile,direction='erp-cost'){
  const read=async file=>XLSX.read(await file.arrayBuffer(),{type:'array'});
  const erp=await read(erpFile),cost=await read(costFile);
  const erpSheet=erp.Sheets['변경본']||erp.Sheets.Sheet2;
  if(!erpSheet)throw new Error('원가경쟁력 파일에 변경본 또는 Sheet2 시트가 없습니다.');
  const matrix=XLSX.utils.sheet_to_json(erpSheet,{header:1,defval:null});
  if(matrix[0]?.[0]!=='품목'||matrix[2]?.[10]!=='재료비')throw new Error('원가경쟁력 파일의 품목(A열)과 재료비(K열)를 확인해 주세요.');
  const numeric=value=>typeof value==='number'&&Number.isFinite(value);
  const map=new Map();for(const row of matrix.slice(3)){
    const key=String(row[0]??'').normalize('NFC').trim();if(!key)continue;
    if(map.has(key))throw new Error('ERP에 중복 품목이 있습니다: '+key);
    map.set(key,row[10]);
  }
  const workbook=new ExcelJS.Workbook();
  await workbook.xlsx.load(await costFile.arrayBuffer());
  const previous=workbook.getWorksheet('비교표');if(previous)workbook.removeWorksheet(previous.id);
  const sourceNames=cost.SheetNames.filter(name=>name!=='비교표');
  workbook.calcProperties.fullCalcOnLoad=true;
  const sheet=workbook.addWorksheet('비교표');sheet.views=structuredClone(COMPARISON_FORMAT.views);
  sheet.orderNo=Math.min(...workbook.worksheets.map(tab=>tab.orderNo))-1;
  const views=Array.isArray(workbook.views)&&workbook.views.length?workbook.views:[{}];
  workbook.views=views.map(view=>({...view,activeTab:0,firstSheet:0}));
  sheet.addRow(['비교표 작성']);sheet.addRow([]);sheet.addRow(['번호','품목','품목명','ERP','원가계산','차이','비교']);
  let warnings=0;
  for(const name of sourceNames){
    const source=cost.Sheets[name],title=String(source.A1?.v??'').trim(),split=title.indexOf('_');
    if(split<=0)throw new Error(name+' 시트의 A1에 품목_품목명이 없습니다.');
    const item=title.slice(0,split).normalize('NFC'),label=title.slice(split+1);
    const rows=XLSX.utils.sheet_to_json(source,{header:1,defval:null});
    const totals=rows.slice(3).filter(row=>row[12]==='합계');
    if(totals.length!==1||!numeric(totals[0][13]))throw new Error(name+' 시트의 합계(M열 표시/N열 값)를 확인해 주세요. Excel에서 재계산 후 저장해 주세요.');
    const calculated=totals[0][13],value=map.get(item),valid=numeric(value);
    const row=sheet.addRow([sheet.rowCount-2,item,label,valid?value:null,calculated,null,valid?null:(map.has(item)?'ERP 값 확인':'ERP 품목 없음')]);
    if(valid){const r=row.number;row.getCell(6).value={formula:direction==='cost-erp'?`E${r}-D${r}`:`D${r}-E${r}`,result:direction==='cost-erp'?calculated-value:value-calculated};}
    else warnings++;
  }
  if(sheet.rowCount===3)throw new Error('멀티 원가 분석 파일에 시트가 없습니다.');
  COMPARISON_FORMAT.columns.forEach((width,i)=>sheet.getColumn(i+1).width=width);
  for(let r=1;r<=sheet.rowCount;r++){
    const format=COMPARISON_FORMAT.rows[Math.min(r,4)-1];sheet.getRow(r).height=format.height;
    for(let c=1;c<=7;c++)sheet.getCell(r,c).style=structuredClone(format.styles[c-1]);
  }
  COMPARISON_FORMAT.merges.forEach(range=>sheet.mergeCells(range));
  return {data:await workbook.xlsx.writeBuffer(),count:sourceNames.length,warnings,name:costFile.name};
}

let comparisonBusy=false;
(() => {
 const open=document.getElementById('comparisonOpen');
 open.addEventListener('click',()=>{
  const dialog=document.createElement('dialog');dialog.style.cssText='width:520px;max-width:90vw;padding:28px;border:1px solid #ccd6ee;border-radius:16px';
  const title=document.createElement('h2');title.textContent='[3] 비교표 작성';
  const form=document.createElement('form');form.style.cssText='display:grid;gap:18px';
  const inputs=[];
  for(const name of ['원가경쟁력 파일 (변경본 포함)','멀티원가 분석 파일']){
   const label=document.createElement('label');label.textContent=name;
   const input=document.createElement('input');input.type='file';input.accept='.xlsx';input.required=true;input.style.display='block';label.append(input);form.append(label);inputs.push(input);
  }
  const note=document.createElement('p');note.textContent='멀티원가 분석 파일에 비교표 시트를 추가합니다. 기존 비교표는 교체됩니다. 차이 = ERP − 원가계산';
  const message=document.createElement('p');message.setAttribute('role','status');
  const save=document.createElement('button');save.type='submit';save.textContent='비교표 추가 후 다운로드';
  const cancel=document.createElement('button');cancel.type='button';cancel.textContent='닫기';
  const close=()=>{if(!comparisonBusy){dialog.close();dialog.remove();}};
  cancel.addEventListener('click',close);dialog.addEventListener('cancel',event=>{event.preventDefault();close();});
  form.addEventListener('submit',async event=>{
   event.preventDefault();if(comparisonBusy)return;comparisonBusy=true;save.disabled=true;cancel.disabled=true;inputs.forEach(input=>input.disabled=true);
   try{message.textContent='모든 시트의 품목과 ERP를 비교하고 있습니다…';
    const result=await buildCostComparison(inputs[0].files[0],inputs[1].files[0]);
    const saved=await window.desktopFile.saveExcel(new Uint8Array(result.data),result.name);
    message.textContent=saved.canceled?'저장을 취소했습니다.':`${result.count}개 품목 비교표 저장 완료 · ERP 확인 필요 ${result.warnings}건`;
   }catch(error){message.textContent='비교표 생성 실패: '+error.message;}
   finally{comparisonBusy=false;save.disabled=false;cancel.disabled=false;inputs.forEach(input=>input.disabled=false);}
  });
  form.append(note,message,save,cancel);dialog.append(title,form);document.body.append(dialog);dialog.showModal();
 });
})();
