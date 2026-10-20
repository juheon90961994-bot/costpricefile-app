const test=require('node:test');
const assert=require('node:assert/strict');
const XLSX=require('../xlsx.full.min.js');
const ExcelJS=require('../exceljs.min.js');
const {convertBytes,TARGET_COLUMNS}=require('../ebom-mbom.js');

test('EBOM columns are rearranged into the reference MBOM layout',async()=>{
  const sourceHeaders=[
    'S','단계','표시단계','품목','품목명','규격','세부규격','내외자구분','내외자구분명','생산단위','재고단위',
    'MAKER','팬텀여부','BOM유형','BOM구분','조달구분','정미수량','로스율','실소요량','설계변경일','차수',
    '설계변경코드','누적정미수량','누적실소요량','주거래처코드','주거래처','위치번호','HS CODE','도면번호(품목)',
    '사용자정의상태코드','비고','품목군','사용자정의상태','순번','재질','첨부파일','LOCATION NO','등록일','등록자',
    '등록자명','수정일','수정자','수정자명','상위품목코드','상위품목명','규격(상위품목)','도면번호(상위품목)',
    '중량단위','중량','수입검사레벨','품목사용여부','시작일','종료일','도면번호'
  ];
  const row=sourceHeaders.map(header=>header+' 값');
  const source=XLSX.utils.aoa_to_sheet([sourceHeaders,[],row]);
  source['!cols']=sourceHeaders.map((_,index)=>({width:index+5}));
  const book=XLSX.utils.book_new();XLSX.utils.book_append_sheet(book,source,'EBOM');
  const output=await convertBytes(XLSX.write(book,{type:'array',bookType:'xlsx'}),XLSX,ExcelJS);
  const result=XLSX.read(output,{type:'array',cellStyles:true,raw:true}).Sheets['변경'];
  const value=(row,column)=>result[XLSX.utils.encode_cell({r:row,c:column})]?.v;

  assert.deepEqual(TARGET_COLUMNS.map((_,column)=>value(0,column)??null),TARGET_COLUMNS);
  assert.equal(value(2,5),'BOM구분 값');
  assert.equal(value(2,6),'규격 값');
  assert.equal(value(2,25),undefined);
  assert.equal(value(2,40),'등록자명 값');
  assert.equal(value(2,54),'도면번호 값');
  assert.equal(result.A1.s.fgColor.rgb,'969696');
  assert.equal(result.AO3.s.fgColor.rgb,'FFFF00');
  assert.equal(result.BC3.s.fgColor.rgb,'FFFF00');
});
