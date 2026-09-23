const input=document.querySelector('#fileInput');
const drop=document.querySelector('#dropzone');
const button=document.querySelector('#convert');
const nameEl=document.querySelector('#fileName');
const meta=document.querySelector('#fileMeta');
const statusEl=document.querySelector('#status');
const tabGenerate=document.querySelector('#tabGenerate');
const tabUpload=document.querySelector('#tabUpload');
const tabDownload=document.querySelector('#tabDownload');
const tabLookup=document.querySelector('#tabLookup');
let selected;
let latestDownload;

const OUTPUT_HEADERS=['단계','품목','품목명','규격','소요량','담당자','주거래처','BOM구분','계정구분','현재고','단위','달러단가','원화단가','합계'];
const SOURCE_FIELDS=['단계','품목','품목명','규격','누적실소요량',null,'주거래처','BOM구분','계정구분','현재고',null,null,null,null];
const COLUMN_WIDTHS=[5.875,11.625,51.5,44.125,8.125,10.125,18.625,10.875,9.375,9.5,11.625,12.125,12.625,14];
const ACCOUNTING='_-* #,##0_-;\\-* #,##0_-;_-* "-"_-;_-@_-';
const BORDER={top:{style:'thin',color:{argb:'FF000000'}},left:{style:'thin',color:{argb:'FF000000'}},bottom:{style:'thin',color:{argb:'FF000000'}},right:{style:'thin',color:{argb:'FF000000'}}};

function setStatus(type,text){statusEl.className=`status ${type}`;statusEl.querySelector('span').textContent=type==='working'?'↻':type==='done'?'✓':type==='error'?'!':'i';statusEl.querySelector('p').textContent=text}
function choose(file){if(!file)return;if(!/\.(xls|xlsx|xlsm)$/i.test(file.name)){setStatus('error','.xls, .xlsx 또는 .xlsm 파일만 선택할 수 있어요.');return}selected=file;latestDownload=undefined;nameEl.textContent=file.name;meta.textContent=`${(file.size/1024).toFixed(file.size>102400?0:1)} KB`;button.disabled=false;setStatus('ready','원가를 생성할 준비가 되었습니다.')}

function downloadResult(){
  if(!latestDownload)return;
  const url=URL.createObjectURL(latestDownload.blob);
  const anchor=document.createElement('a');anchor.href=url;anchor.download=latestDownload.name;anchor.click();setTimeout(()=>URL.revokeObjectURL(url),1000);
}

tabGenerate.addEventListener('click',()=>{document.querySelector('#uploadWorkspace').hidden=true;document.querySelector('#lookupWorkspace').hidden=true;document.querySelector('#costWorkspace').hidden=false;tabGenerate.classList.add('active');tabUpload.classList.remove('active');tabLookup.classList.remove('active');tabDownload.classList.remove('active');document.querySelector('#costWorkspace').scrollIntoView({behavior:'smooth',block:'center'});button.focus({preventScroll:true})});
tabDownload.addEventListener('click',async()=>{
  document.querySelector('#costWorkspace').hidden=true;
  document.querySelector('#uploadWorkspace').hidden=true;
  document.querySelector('#lookupWorkspace').hidden=true;
  tabGenerate.classList.remove('active');
  tabUpload.classList.remove('active');
  tabLookup.classList.remove('active');
  tabDownload.classList.add('active');
  tabDownload.disabled=true;
  try{
    if(!window.costDatabase?.downloadCostTable)throw Error('원가 다운로드 기능을 불러오지 못했습니다. 화면을 새로고침해 주세요.');
    const result=await window.costDatabase.downloadCostTable();
    if(result.canceled){setStatus('ready','원가 다운로드를 취소했습니다.');return}
    setStatus('done',`Supabase 원가 자료 ${result.count.toLocaleString()}건을 저장했습니다.`);
  }catch(error){const message=typeof toFriendlyError==='function'?toFriendlyError(error):error.message;setStatus('error',`원가 다운로드 실패: ${message}`)}
  finally{tabDownload.disabled=false}
});

drop.addEventListener('click',()=>input.click());
input.addEventListener('change',()=>choose(input.files[0]));
drop.addEventListener('dragover',event=>{event.preventDefault();drop.classList.add('dragging')});
drop.addEventListener('dragleave',()=>drop.classList.remove('dragging'));
drop.addEventListener('drop',event=>{event.preventDefault();drop.classList.remove('dragging');choose(event.dataTransfer.files[0])});

button.addEventListener('click',async()=>{
  if(!selected)return;
  button.disabled=true;button.firstChild.textContent='제품 원가 생성 중… ';setStatus('working','제품 BOM을 확인하고 원가를 생성하고 있어요…');
  try{
    const sourceBook=XLSX.read(await selected.arrayBuffer(),{type:'array',cellDates:true,raw:true});
    const sourceSheet=sourceBook.Sheets[sourceBook.SheetNames[0]];
    if(!sourceSheet)throw Error('워크시트가 없습니다.');
    const rows=XLSX.utils.sheet_to_json(sourceSheet,{header:1,defval:null,raw:true});
    const sourceHeaders=rows[0]||[];
    const indexes=SOURCE_FIELDS.map(field=>field===null?-1:sourceHeaders.indexOf(field));
    const managerIndex=sourceHeaders.indexOf('담당자');
    const unitIndex=sourceHeaders.indexOf('단위');
    const dollarPriceIndex=sourceHeaders.indexOf('달러단가');
    const missing=SOURCE_FIELDS.filter((field,index)=>field!==null&&indexes[index]<0);
    if(missing.length)throw Error(`필수 열을 찾을 수 없습니다: ${[...new Set(missing)].join(', ')}`);

    if(!window.costDatabase)throw Error('Supabase 원가 조회 기능을 불러오지 못했습니다. 화면을 새로고침해 주세요.');
    setStatus('working','Supabase cost 테이블에서 품목별 단가를 조회하고 있어요…');
    const unitPriceResult=await window.costDatabase.getUnitPriceMap();
    const exchangeRateResult=await getFirstUsdExchangeRate();
    let matchedPriceCount=0;
    let missingPriceCount=0;

    const workbook=new ExcelJS.Workbook();
    workbook.calcProperties.fullCalcOnLoad=true;
    const sheet=workbook.addWorksheet(sourceBook.SheetNames[0]||'Sheet1',{views:[{state:'frozen',xSplit:8,ySplit:3,topLeftCell:'I4',activeCell:'I4'}]});
    sheet.addRow([]);sheet.addRow([]);sheet.addRow(OUTPUT_HEADERS);
    sheet.getCell(2,11).value='환율';
    sheet.getCell(2,12).value=exchangeRateResult.rate;
    for(const sourceRow of rows.slice(1)){
      if(sourceRow.every(value=>value===null||value===''))continue;
      const item=String(sourceRow[indexes[1]]??'').normalize('NFKC').replace(/\u00a0/g,' ').trim().toUpperCase();
      if(!item)continue;
      const outputRow=indexes.map((index,column)=>column>=10?null:sourceRow[index]);
      if(managerIndex>=0)outputRow[5]=sourceRow[managerIndex];
      if(unitIndex>=0)outputRow[10]=sourceRow[unitIndex];
      if(dollarPriceIndex>=0)outputRow[11]=sourceRow[dollarPriceIndex];
      const costRecord=unitPriceResult.records.get(item);
      if(costRecord){
        outputRow[5]=costRecord['담당자']??outputRow[5];
        outputRow[6]=costRecord['거래처명']??outputRow[6];
        outputRow[10]=costRecord['단위']??outputRow[10];
        outputRow[11]=toExcelNumberOrValue(costRecord['달러단가'],outputRow[11]);
      }
      if(item&&unitPriceResult.prices.has(item)){
        outputRow[12]=toExcelPrice(unitPriceResult.prices.get(item));
        matchedPriceCount+=1;
      }else{
        outputRow[12]='단가없음';
        missingPriceCount+=1;
      }
      sheet.addRow(outputRow);
    }
    const dataLastRow=sheet.rowCount;
    for(let row=4;row<=dataLastRow;row++){
      const quantity=toNumber(sheet.getCell(row,5).value);
      const unit=String(sheet.getCell(row,11).value??'').normalize('NFKC').trim().toUpperCase();
      const priceColumn=unit==='USD'?12:13;
      const price=toNumber(sheet.getCell(row,priceColumn).value);
      const exchangeRate=unit==='USD'?exchangeRateResult.rate:1;
      sheet.getCell(row,14).value=quantity!==null&&price!==null
        ?{formula:unit==='USD'?`L${row}*E${row}*$L$2`:`E${row}*M${row}`,result:quantity*price*exchangeRate}
        :'단가없음';
    }
    const totalRow=dataLastRow+1;
    const totalAmount=[...Array(Math.max(0,dataLastRow-3))].reduce((sum,_,index)=>{
      const value=toNumber(sheet.getCell(index+4,14).value?.result??sheet.getCell(index+4,14).value);
      return sum+(value??0);
    },0);
    sheet.getCell(totalRow,13).value='합계';
    sheet.getCell(totalRow,14).value=dataLastRow>=4
      ?{formula:`SUM(N4:N${dataLastRow})`,result:totalAmount}
      :0;
    sheet.autoFilter={
      from:{row:3,column:1},
      to:{row:dataLastRow,column:14},
    };
    applyReferenceFormat(sheet,totalRow,dataLastRow);

    const data=await workbook.xlsx.writeBuffer();
    latestDownload={blob:new Blob([data],{type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'}),name:`${selected.name.replace(/\.(xls|xlsx|xlsm)$/i,'')}_제품원가.xlsx`};
    downloadResult();
    setStatus('done',`제품 원가 생성을 완료했습니다. 단가 적용 ${matchedPriceCount}건 · 단가없음 ${missingPriceCount}건`);
  }catch(error){const message=typeof toFriendlyError==='function'?toFriendlyError(error):error.message;setStatus('error',`원가 생성 실패: ${message}`)}finally{button.disabled=false;button.firstChild.textContent='제품 원가 생성 '}
});

function toExcelPrice(value){
  if(typeof value==='number')return value;
  const text=String(value??'').trim();
  const normalized=text.replace(/,/g,'');
  return /^-?\d+(\.\d+)?$/.test(normalized)?Number(normalized):text||'단가없음';
}

function toExcelNumberOrValue(value,fallback=null){
  if(value===null||value===undefined||value==='')return fallback;
  if(typeof value==='number')return value;
  const text=String(value).trim();
  const normalized=text.replace(/,/g,'');
  return /^-?\d+(\.\d+)?$/.test(normalized)?Number(normalized):text;
}

async function getFirstUsdExchangeRate(){
  if(window.desktopExchange?.getFirstUsdExchangeRate){
    const result=await window.desktopExchange.getFirstUsdExchangeRate();
    if(Number.isFinite(Number(result?.rate))&&Number(result.rate)>0)return{...result,rate:Number(result.rate)};
    throw Error('조회 기준 최초 공지 환율을 확인할 수 없습니다.');
  }
  const input=window.prompt('브라우저 실행에서는 하나은행 환율을 자동 조회할 수 없습니다. 조회 기준 최초 공지 USD 매매기준율을 입력해 주세요.');
  const rate=toNumber(input);
  if(rate===null||rate<=0)throw Error('올바른 최초 공지 환율을 입력해 주세요.');
  return{rate,source:'사용자 입력'};
}

function toNumber(value){
  if(typeof value==='number'&&Number.isFinite(value))return value;
  const normalized=String(value??'').trim().replace(/,/g,'');
  if(!/^-?\d+(\.\d+)?$/.test(normalized))return null;
  const number=Number(normalized);
  return Number.isFinite(number)?number:null;
}

function applyReferenceFormat(sheet,lastRow,dataLastRow){
  sheet.columns.forEach((column,index)=>column.width=COLUMN_WIDTHS[index]);
  sheet.getRow(3).height=15;
  for(let row=4;row<=lastRow;row++)sheet.getRow(row).height=14.25;
  for(let row=3;row<=lastRow;row++)for(let column=1;column<=14;column++){
    const cell=sheet.getCell(row,column);cell.font={name:'맑은 고딕',size:11,color:{argb:'FF000000'}};cell.border=BORDER;cell.alignment={vertical:'middle',wrapText:false};
  }
  for(let column=1;column<=14;column++){
    const cell=sheet.getCell(3,column);cell.fill={type:'pattern',pattern:'solid',fgColor:{argb:'FFD9F2D0'}};cell.alignment={horizontal:'center',vertical:'middle',wrapText:false};
  }
  [11,12].forEach(column=>{
    const cell=sheet.getCell(2,column);cell.font={name:'맑은 고딕',size:11,bold:column===11,color:{argb:'FF000000'}};cell.border=BORDER;cell.alignment={horizontal:'center',vertical:'middle',wrapText:false};
  });
  sheet.getCell(2,11).fill={type:'pattern',pattern:'solid',fgColor:{argb:'FFD9F2D0'}};
  sheet.getCell(2,12).numFmt='#,##0.00';
  for(let row=4;row<=dataLastRow;row++){
    [1,2,5,6,8,9,10,11,12].forEach(column=>sheet.getCell(row,column).alignment={horizontal:'center',vertical:'middle',wrapText:false});
    sheet.getCell(row,5).font={name:'맑은 고딕',size:12,color:{argb:'FF000000'}};
    sheet.getCell(row,10).font={name:'맑은 고딕',size:12,color:{argb:'FF000000'}};
    sheet.getCell(row,11).font={name:'맑은 고딕',size:12,color:{argb:'FF000000'}};
    sheet.getCell(row,12).font={name:'맑은 고딕',size:12,color:{argb:'FF000000'}};
    sheet.getCell(row,5).numFmt='###,###,###,###.###';
    [10,11,12,13,14].forEach(column=>sheet.getCell(row,column).numFmt=ACCOUNTING);
  }
  [10,11,12,13,14].forEach(column=>sheet.getCell(3,column).numFmt=ACCOUNTING);
  sheet.getCell(lastRow,13).font={name:'맑은 고딕',size:11,bold:true,color:{argb:'FF000000'}};
  sheet.getCell(lastRow,13).alignment={horizontal:'center',vertical:'middle',wrapText:false};
  sheet.getCell(lastRow,14).font={name:'맑은 고딕',size:11,bold:true,color:{argb:'FF000000'}};
  sheet.getCell(lastRow,14).numFmt=ACCOUNTING;
  sheet.pageSetup={orientation:'landscape',fitToHeight:0};
}
