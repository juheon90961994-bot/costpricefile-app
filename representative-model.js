// Representative model uses the archived BOM and a fresh cost snapshot on every request.
async function buildRepresentativeModel(file, date = new Date()) {
  const book = XLSX.read(await file.arrayBuffer(), {type:'array'});
  const source = book.Sheets['원본'] || book.Sheets[book.SheetNames[0]];
  if (!source) throw new Error('BOM 시트가 없습니다.');
  const rows = XLSX.utils.sheet_to_json(source, {header:1, defval:null, raw:true});
  const fields = ['품목','품목명','누적정미수량','주거래처','계정구분'];
  const header = rows.findIndex(row => fields.every(field => row.some(v => String(v ?? '').trim() === field)));
  if (header < 0) throw new Error('BOM 필수 열을 확인해 주세요: ' + fields.join(', '));
  const indexes = fields.map(field => rows[header].findIndex(v => String(v ?? '').trim() === field));
  const entries = rows.slice(header+1).filter(row => String(row[indexes[0]] ?? '').trim());
  if (!entries.length) throw new Error('BOM에 품목 자료가 없습니다.');
  const snapshot = await window.costDatabase.getUnitPriceMap();
  const output = new ExcelJS.Workbook();
  output.calcProperties.fullCalcOnLoad = true;
  const format = REPRESENTATIVE_FORMAT;
  const sheet = output.addWorksheet('변환본', {views:structuredClone(format.views), properties:structuredClone(format.properties), pageSetup:structuredClone(format.pageSetup)});
  for (const col of format.columns) for (let c=col.min;c<=col.max;c++) {
    sheet.getColumn(c).width=col.width;
    sheet.getColumn(c).hidden=col.hidden;
    sheet.getColumn(c).style=structuredClone(col.style);
  }
  function styleRow(r, template) {
    sheet.getRow(r).height=template.height;
    for(let c=1;c<=17;c++) sheet.getCell(r,c).style=structuredClone(template.styles[c-1]);
  }
  for(let r=1;r<=4;r++) {
    sheet.getRow(r).values=structuredClone(format.rows[r-1].values.slice(1));
    styleRow(r,format.rows[r-1]);
  }
  const month = offset => new Date(date.getFullYear(), date.getMonth()+offset, 1).getMonth()+1;
  sheet.getCell('F4').value=month(-1)+'월 예상가';
  for(const [column,offset] of [['H',-1],['K',0],['N',1],['Q',2]]) sheet.getCell(column+'4').value=month(offset)+'월';
  const numeric = value => {
    if(value===null || value===undefined || String(value).trim()==='') return null;
    const n=Number(String(value).replace(/,/g,''));
    return Number.isFinite(n)?n:null;
  };
  let warnings=0,total=0;
  entries.forEach((sourceRow,index) => {
    const r=index+5, values=indexes.map(i=>sourceRow[i]);
    const record=snapshot.records.get(normalizeItemKey(values[0]));
    const quantity=numeric(values[2]);
    if(quantity===null) throw new Error(values[0]+': 누적정미수량이 숫자가 아닙니다.');
    values[2]=quantity;
    if(record) values[3]=record['거래처명'] ?? null;
    sheet.getRow(r).values=values;
    styleRow(r,format.rows[4]);
    const won=numeric(record?.['원화단가']),usd=numeric(record?.['달러단가']);
    if(!record || won===null) warnings++;
    for(const c of [9,12,15]) sheet.getCell(r,c).value=usd;
    for(const c of [10,13,16]) sheet.getCell(r,c).value=won;
    if(won===null) sheet.getCell(r,6).value='단가없음';
    else {
      sheet.getCell(r,6).value={formula:`J${r}*C${r}`,result:won*quantity};
      total+=won*quantity;
    }
  });
  const last=entries.length+4,footer=last+1;
  styleRow(footer,format.rows[5]);
  for(const c of [5,7,10,13,16]) sheet.getCell(footer,c).value='합계';
  for(const c of ['F','H','K','N','Q']) sheet.getCell(c+footer).value={formula:`SUM(${c}5:${c}${last})`,result:c==='F'?total:0};
  return {data:await output.xlsx.writeBuffer(),count:entries.length,warnings,name:file.name.replace(/\.(xlsx|xlsm|xls)$/i,'')+'_대표모델.xlsx'};
}

(() => {
  const form=document.getElementById('representativeForm'),input=document.getElementById('representativeItem');
  const button=document.getElementById('representativeGenerate'),message=document.getElementById('representativeMessage');
  let busy=false;
  window.representativeModel={isBusy:()=>busy,reset:()=>{if(!busy){form.reset();message.textContent='';}}};
  form.addEventListener('submit',async event=>{
    event.preventDefault();
    const item=input.value.trim();if(busy || !item)return;
    busy=true;button.disabled=true;input.disabled=true;
    message.textContent='보관 BOM 및 최신 DB 단가를 조회하고 있습니다…';
    try {
      const file=await ExcelArchive.readForItem(item);
      const result=await buildRepresentativeModel(file);
      const saved=await window.desktopFile.saveExcel(new Uint8Array(result.data),result.name);
      message.textContent=saved.canceled?'저장을 취소했습니다.':`${result.count}개 품목 저장 완료`+(result.warnings?` · 단가 확인 필요 ${result.warnings}건 (단가없음 표시)`:'');
    } catch(error) {message.textContent='대표모델 생성 실패: '+error.message;}
    finally {busy=false;button.disabled=false;input.disabled=false;}
  });
})();
