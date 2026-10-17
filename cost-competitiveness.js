// Convert the three-row cost statement header to the supplied Sheet2 layout.
async function buildCompetitiveness(file) {
  const inputBytes=new Uint8Array(await file.arrayBuffer());
  const bytes=/\.xls$/i.test(file.name)?await window.desktopFile.convertLegacyExcel(inputBytes):inputBytes;
  const book=XLSX.read(bytes,{type:'array',cellStyles:true});
  const preserved=new ExcelJS.Workbook();await preserved.xlsx.load(bytes);
  const source=book.Sheets.Sheet1 || book.Sheets[book.SheetNames[0]];
  if(!source)throw new Error('Sheet1 자료가 없습니다.');
  const matrix=XLSX.utils.sheet_to_json(source,{header:1,defval:null,raw:true});
  // Excel stores a merged heading only in its top-left cell.
  // Expand header merges for lookup without changing source cells or data rows.
  const headers=matrix.slice(0,3).map(row=>row.slice());
  const origin=XLSX.utils.decode_range(source['!ref']||'A1').s;
  for(const merge of source['!merges']||[]){
    const value=source[XLSX.utils.encode_cell(merge.s)]?.v;
    for(let r=Math.max(merge.s.r,origin.r);r<=Math.min(merge.e.r,origin.r+2);r++){
      for(let c=Math.max(merge.s.c,origin.c);c<=merge.e.c;c++){
        if(headers[r-origin.r])headers[r-origin.r][c-origin.c]=value;
      }
    }
  }
  const text=value=>String(value??'').trim();
  const find=(top,middle,bottom)=>{
    const indexes=[];
    for(let c=0;c<(headers[0]?.length||0);c++)if(text(headers[0][c])===top&&(!middle||text(headers[1]?.[c])===middle)&&(!bottom||text(headers[2]?.[c])===bottom))indexes.push(c);
    if(indexes.length!==1)throw new Error('열을 확인해 주세요: '+[top,middle,bottom].filter(Boolean).join(' / '));
    return indexes[0];
  };
  const columns=[find('품목'),find('품목명'),find('입고','수량'),find('입고','직접재료비','금액'),find('입고','직접재료비','금액(외부)'),find('입고','노무비','금액'),find('입고','제조경비','금액'),find('입고','경상개발비','금액'),find('입고','감가상각비','금액')];
  const output=new ExcelJS.Workbook();output.calcProperties.fullCalcOnLoad=true;
  const originalSource=preserved.getWorksheet('Sheet1')||preserved.worksheets[0];
  const original=output.addWorksheet('원본');
  const originalModel=structuredClone(originalSource.model);originalModel.id=original.id;originalModel.name='원본';original.model=originalModel;
  for(const range of originalModel.merges||[])if(!original.model.merges.includes(range))original.mergeCellsWithoutStyle(range);
  const sheet=output.addWorksheet('변경본');
  sheet.addRow(['품목','품목명','입고','입고','입고','입고','입고','입고','입고']);
  sheet.addRow(['품목','품목명','수량','직접재료비','직접재료비','노무비','제조경비','경상개발비','감가상각비']);
  sheet.addRow(['품목','품목명','수량','금액','금액(외부)','금액','금액','금액','금액','제조경비','재료비','개발비']);
  let warnings=0;
  for(let index=3;index<matrix.length;index++){
    const input=matrix[index];if(!text(input[columns[0]]))continue;
    const values=columns.map((c,i)=>{
      const value=input[c];if(i<2)return value;
      if(value===null||value===undefined||value==='')return 0;
      const n=typeof value==='number'?value:Number(String(value).replace(/,/g,'').trim());
      if(!Number.isFinite(n))throw new Error((index+1)+'행의 금액/수량이 숫자가 아닙니다.');
      return n;
    });
    const row=sheet.addRow(values),r=row.number,q=values[2];
    if(q===0)warnings++;
    const formulas=[`(F${r}+G${r}+I${r})/C${r}`,`(D${r}+E${r})/C${r}`,`H${r}/C${r}`];
    const totals=[values[5]+values[6]+values[8],values[3]+values[4],values[7]];
    for(let c=0;c<3;c++)row.getCell(c+10).value={formula:`IF(C${r}=0,"수량 확인",${formulas[c]})`,result:q===0?'수량 확인':totals[c]/q};
  }
  if(sheet.rowCount===3)throw new Error('변환할 품목 자료가 없습니다.');
  const format=COMPETITION_FORMAT;
  sheet.views=[{state:'frozen',xSplit:2,ySplit:3,topLeftCell:'C4',activeCell:'C4'}];sheet.properties=structuredClone(format.properties);sheet.pageSetup=structuredClone(format.pageSetup);
  format.columns.forEach((column,i)=>{sheet.getColumn(i+1).width=column.width;sheet.getColumn(i+1).hidden=column.hidden;});
  for(let r=1;r<=sheet.rowCount;r++){
    const styleRow=format.rows[Math.min(r,4)-1];sheet.getRow(r).height=styleRow.height;
    for(let c=1;c<=12;c++)sheet.getCell(r,c).style=structuredClone(styleRow.styles[c-1]);
    if(r>=4)for(let c=10;c<=12;c++)sheet.getCell(r,c).numFmt=ACCOUNTING;
  }
  format.merges.forEach(range=>sheet.mergeCells(range));
  const materialCosts=[];
  for(let r=4;r<=sheet.rowCount;r++)materialCosts.push({item:sheet.getCell(r,1).value,price:sheet.getCell(r,11).value.result});
  return {data:await output.xlsx.writeBuffer(),materialCosts,count:sheet.rowCount-3,warnings,name:file.name.replace(/\.(xlsx|xlsm|xls)$/i,'')+'_원가경쟁력.xlsx'};
}

let competitivenessBusy=false;
(() => {
  const tab=document.getElementById('tabCompetitiveness'),panel=document.getElementById('competitivenessWorkspace');
  const input=document.getElementById('competitivenessInput'),button=document.getElementById('competitivenessUpload'),message=document.getElementById('competitivenessMessage');
  tab.addEventListener('click',()=>{
    for(const id of ['costWorkspace','uploadWorkspace','lookupWorkspace'])document.getElementById(id).hidden=true;
    for(const el of document.querySelectorAll('.mode-tab'))el.classList.remove('active');
    panel.hidden=false;tab.classList.add('active');
  });
  for(const other of [tabGenerate,tabUpload,tabLookup,tabDownload])other.addEventListener('click',()=>{
    panel.hidden=true;tab.classList.remove('active');input.value='';message.textContent='원가계산현황 포멧 변경';
  });
  button.addEventListener('click',()=>input.click());
  input.addEventListener('change',async()=>{
    const file=input.files[0];if(!file||competitivenessBusy)return;
    if(!/\.(xls|xlsx|xlsm)$/i.test(file.name)){message.textContent='Excel 파일을 선택해 주세요.';return;}
    competitivenessBusy=true;button.disabled=true;message.textContent='원가경쟁력 자료를 생성하고 있습니다…';
    try{
      const result=await buildCompetitiveness(file);
      const saved=await window.desktopFile.saveExcel(new Uint8Array(result.data),result.name);
      if(saved.canceled){message.textContent='저장을 취소했습니다.';return;}
      message.textContent=`${result.count}개 품목 저장 완료 · DB 재료비 반영 중…`;
      try{
        const sync=await window.costDatabase.updateMaterialCosts(result.materialCosts);
        message.textContent=`${result.count}개 품목 저장 완료 · DB 업데이트 ${sync.updated}건 · 동일가격 ${sync.unchanged}건 · 미등록 ${sync.unmatched}건 · 확인 필요 ${sync.review}건 · 실패 ${sync.failed}건`;
        if(sync.errors.length)message.textContent+=' · '+sync.errors.join(' / ');
      }catch(error){message.textContent='파일 저장 완료 · DB 업데이트 실패: '+error.message;}
    }catch(error){message.textContent='변환 실패: '+error.message;}
    finally{competitivenessBusy=false;button.disabled=false;input.value='';}
  });
})();
