// Representative model uses the archived BOM and a fresh cost snapshot on every request.
async function buildRepresentativeModel(file, date = new Date(), options = {}) {
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
  const snapshot = options.snapshot || await window.costDatabase.getUnitPriceMap();
  const output = options.workbook || new ExcelJS.Workbook();
  output.calcProperties.fullCalcOnLoad = true;
  const format = REPRESENTATIVE_FORMAT;
  const sheet = output.addWorksheet(options.sheetName || '변환본', {views:structuredClone(format.views), properties:structuredClone(format.properties), pageSetup:structuredClone(format.pageSetup)});
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
  sheet.getCell('Q3').fill=structuredClone(sheet.getCell('N3').fill);
  sheet.getCell('F4').value=month(-1)+'월 예상가';
  for(const [column,offset] of [['H',-1],['K',0],['N',1],['Q',2]]) sheet.getCell(column+'4').value=month(offset)+'월';
  for(const column of ['H','K','N','Q']) {
    sheet.getCell(column+'3').value=1360;
    sheet.getCell(column+'3').note='기본 환율 1,360원/USD (사용자 지정). 필요시 수정하세요.';
  }
  const numeric = value => {
    if(value===null || value===undefined || String(value).trim()==='') return null;
    const n=Number(String(value).replace(/,/g,''));
    return Number.isFinite(n)?n:null;
  };
  let warnings=0,total=0,monthlyTotal=0;
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
    sheet.getCell(r,7).value=won;
    const monthlyWon=normalizeItemKey(values[0])==='PACS-0002A'?850:won;
    for(const c of [10,13,16]) sheet.getCell(r,c).value=monthlyWon;
    if(['DRU0-0012A','SENM-0016A'].includes(normalizeItemKey(values[0]))) {
      for(const c of [10,13,16]) sheet.getCell(r,c).fill={type:'pattern',pattern:'solid',fgColor:{argb:'FFFFFF00'}};
    }
    sheet.getCell(r,7).numFmt=sheet.getCell(r,10).numFmt;
    if(won===null) sheet.getCell(r,8).value='단가없음';
    else {
      sheet.getCell(r,8).value={formula:`C${r}*G${r}`,result:won*quantity};
      total+=won*quantity;
    }
    for(const [amount,dollar,krw] of [['K','I','J'],['N','L','M'],['Q','O','P']]) {
      sheet.getCell(amount+r).value={
        formula:`IF(ISNUMBER(${dollar}${r}),${dollar}${r}*$${amount}$3*C${r},IF(ISNUMBER(${krw}${r}),${krw}${r}*C${r},"단가없음"))`,
        result:usd!==null?usd*1360*quantity:monthlyWon!==null?monthlyWon*quantity:'단가없음'
      };
    }
    monthlyTotal+=usd!==null?usd*1360*quantity:monthlyWon!==null?monthlyWon*quantity:0;
  });
  const last=entries.length+4,footer=last+1;
  sheet.autoFilter={from:{row:4,column:1},to:{row:last,column:17}};
  styleRow(footer,format.rows[5]);
  for(const c of [5,7,10,13,16]) sheet.getCell(footer,c).value='합계';
  for(const c of ['H','K','N','Q']) sheet.getCell(c+footer).value={formula:`SUM(${c}5:${c}${last})`,result:c==='H'?total:monthlyTotal};
  const tableStart=Math.max(38,footer+5),shift=tableStart-38;
  const shiftReferences=value=>value.replace(/(\$?[A-Z]+\$?)(\d+)/g,(_,column,row)=>column+(Number(row)+shift));
  REPRESENTATIVE_TABLE.rows.forEach((row,index)=>{
    const r=tableStart+index;
    sheet.getRow(r).height=row.height;
    row.cells.forEach((source,c)=>{
      const cell=sheet.getCell(r,c+8);
      cell.style=structuredClone(source.style);
      cell.value=structuredClone(source.value);
      if(source.value?.formula)cell.value={formula:shiftReferences(source.value.formula),result:structuredClone(source.value.result)};
    });
  });
  for(const range of REPRESENTATIVE_TABLE.merges)sheet.mergeCells(shiftReferences(range));
  // Excel enforces locked cells only when worksheet protection is enabled.
  for(let r=1;r<=tableStart+6;r++)for(let c=1;c<=17;c++)sheet.getCell(r,c).protection={locked:false};
  for(const address of ['N39','N42','Q39','Q42','K41','K44','N41','N44','Q41','Q44'])sheet.getCell(shiftReferences(address)).protection={locked:true};
  await sheet.protect('',{selectLockedCells:true,selectUnlockedCells:true,autoFilter:true,formatCells:true,formatColumns:true,formatRows:true});
  return {data:options.workbook?null:await output.xlsx.writeBuffer(),count:entries.length,warnings,name:file.name.replace(/\.(xlsx|xlsm|xls)$/i,'')+'_대표모델.xlsx'};
}

async function buildRepresentativeModels(files, date = new Date()) {
  if(!files.length)throw new Error('BOM 파일을 선택해 주세요.');
  const workbook=new ExcelJS.Workbook(),used=new Set();
  const snapshot=await window.costDatabase.getUnitPriceMap();
  let count=0,warnings=0;
  for(const file of files){
    const base=file.name.replace(/\.(xlsx|xlsm|xls)$/i,'').split('_')[0]
      .replace(/[\\/*?:\[\]]/g,' ').replace(/^'+|'+$/g,'').trim() || '대표모델';
    let sheetName=base.slice(0,31),suffix=1;
    while(used.has(sheetName.toLowerCase()) || sheetName.toLowerCase()==='history'){
      const tail=' ('+(++suffix)+')';sheetName=base.slice(0,31-tail.length)+tail;
    }
    used.add(sheetName.toLowerCase());
    try{
      const result=await buildRepresentativeModel(file,date,{workbook,snapshot,sheetName});
      count+=result.count;warnings+=result.warnings;
    }catch(error){throw new Error(file.name+': '+error.message);}
  }
  const stamp=[date.getFullYear(),String(date.getMonth()+1).padStart(2,'0'),String(date.getDate()).padStart(2,'0')].join('.');
  return {data:await workbook.xlsx.writeBuffer(),count,warnings,sheets:files.length,name:`${date.getMonth()+1}월 원가경쟁력자료 대표 모델_${stamp}.xlsx`};
}

async function findRepresentativeArchives(keyword) {
  const normalize=value=>String(value??'').normalize('NFC').toLocaleLowerCase();
  const needle=normalize(keyword.trim()),files=await ExcelArchive.list();
  if(!needle)return files;
  const fields=['품목','품목명','프로젝트명','세부프로젝트명'];
  const settings=getConnectionSettings(),matches=new Map();
  for(let offset=0;;){
    const url=new URL(settings.url+'/rest/v1/order_mobility');
    const pattern='%'+keyword.trim().replace(/\\/g,'\\\\').replace(/[%_]/g,'\\$&')+'%';
    const literal='"'+pattern.replace(/\\/g,'\\\\').replace(/"/g,'\\"')+'"';
    url.searchParams.set('select',fields.join(','));
    url.searchParams.set('or','('+fields.map(field=>field+'.ilike.'+literal).join(',')+')');
    url.searchParams.set('order',fields.map(field=>field+'.asc.nullsfirst').join(','));
    url.searchParams.set('offset',String(offset));url.searchParams.set('limit','500');
    const page=await supabaseRequest(settings,url.toString());
    if(!Array.isArray(page))throw new Error('프로젝트 검색 응답을 확인할 수 없습니다.');
    if(!page.length)break;
    for(const row of page){
      if(fields.some(field=>normalize(row[field]).includes(needle))){
        const key=normalize(row['품목']);
        if(!matches.has(key))matches.set(key,[]);
        matches.get(key).push(row);
      }
    }
    offset+=page.length;
  }
  return files.flatMap(file=>{
    const details=matches.get(normalize(ExcelArchive.identity(file.filename)))||[];
    return normalize(file.filename).includes(needle)||details.length?[{...file,details}]:[];
  });
}

(() => {
  const form=document.getElementById('representativeForm'),input=document.getElementById('representativeItem');
  const button=document.getElementById('representativeGenerate'),message=document.getElementById('representativeMessage');
  const search=document.getElementById('representativeSearch'),results=document.getElementById('representativeResults');
  const selected=new Map();
  let busy=false;
  window.representativeModel={isBusy:()=>busy,reset:()=>{if(!busy){form.reset();message.textContent='';selected.clear();results.replaceChildren();button.disabled=true;}}};
  search.addEventListener('click',async()=>{
    if(busy)return;
    busy=true;search.disabled=true;button.disabled=true;input.disabled=true;
    message.textContent='보관 BOM 목록을 조회하고 있습니다…';
    try{
      const rows=await findRepresentativeArchives(input.value);
      results.replaceChildren();
      for(const row of rows){
        const label=document.createElement('label'),checkbox=document.createElement('input');
        checkbox.type='checkbox';checkbox.checked=selected.has(row.filename);
        checkbox.addEventListener('change',()=>{
          if(checkbox.checked)selected.set(row.filename,row);else selected.delete(row.filename);
          button.disabled=!selected.size;message.textContent=selected.size+'개 BOM 선택';
        });
        label.style.cssText='display:flex;gap:8px;align-items:center;margin:8px 0';
        const details=[...new Set((row.details||[]).map(detail=>['품목명','프로젝트명','세부프로젝트명'].map(field=>detail[field]).filter(Boolean).join(' / ')).filter(Boolean))];
        label.append(checkbox,document.createTextNode(row.filename+(details.length?' — '+details.join(' · '):'')));results.append(label);
      }
      message.textContent=(rows.length?rows.length+'개 보관 파일 · 분석할 BOM을 여러 개 선택하세요.':'일치하는 보관 파일이 없습니다.')+' · 선택 '+selected.size+'개';
    }catch(error){message.textContent='보관 BOM 조회 실패: '+error.message;}
    finally{busy=false;search.disabled=false;input.disabled=false;button.disabled=!selected.size;}
  });
  input.addEventListener('keydown',event=>{if(event.key==='Enter'){event.preventDefault();search.click();}});
  form.addEventListener('submit',async event=>{
    event.preventDefault();
    if(busy || !selected.size)return;
    busy=true;button.disabled=true;input.disabled=true;search.disabled=true;
    results.querySelectorAll('input').forEach(el=>el.disabled=true);
    message.textContent='보관 BOM 및 최신 DB 단가를 조회하고 있습니다…';
    try {
      const files=[];
      for(const row of selected.values())files.push(await ExcelArchive.readForItem(ExcelArchive.identity(row.filename)));
      const result=await buildRepresentativeModels(files);
      const saved=await window.desktopFile.saveExcel(new Uint8Array(result.data),result.name);
      message.textContent=saved.canceled?'저장을 취소했습니다.':`${result.sheets}개 시트 · ${result.count}개 품목 저장 완료`+(result.warnings?` · 단가 확인 필요 ${result.warnings}건 (단가없음 표시)`:'');
    } catch(error) {message.textContent='대표모델 생성 실패: '+error.message;}
    finally {busy=false;button.disabled=!selected.size;input.disabled=false;search.disabled=false;results.querySelectorAll('input').forEach(el=>el.disabled=false);}
  });
})();
