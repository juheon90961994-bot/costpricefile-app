const test=require('node:test');
const assert=require('node:assert/strict');
const fs=require('node:fs');
const XLSX=require('../xlsx.full.min.js');
const ExcelJS=require('../exceljs.min.js');
function setup(){
  let calls=0;
  const window={costDatabase:{getUnitPriceMap:async()=>{calls++;return {records:new Map([['PART',{원화단가:100,달러단가:2,거래처명:'최신 거래처'}]])};}}};
  const code=['representative-format.js','representative-table.js'].map(name=>fs.readFileSync(require.resolve('../'+name),'utf8')).join('\n')+'\n'+
    fs.readFileSync(require.resolve('../representative-model.js'),'utf8').split('(() => {')[0];
  const build=Function('XLSX','ExcelJS','window','normalizeItemKey',code+'\nreturn buildRepresentativeModels;')(XLSX,ExcelJS,window,value=>String(value).trim().toUpperCase());
  return {build,calls:()=>calls};
}
function file(name,quantity=2){
  const book=XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(book,XLSX.utils.aoa_to_sheet([
    ['품목','품목명','누적정미수량','주거래처','계정구분'],
    ['PART','부품',quantity,'이전 거래처','재료']
  ]),'원본');
  const bytes=XLSX.write(book,{type:'buffer',bookType:'xlsx'});
  return {name,arrayBuffer:async()=>bytes};
}
test('multiple BOMs share one price snapshot and keep separate formulas, totals and protection',async()=>{
  const {build,calls}=setup();
  const result=await build([file('MODEL-A_BOM.xlsx'),file('MODEL-B_사양_추가.xlsx',3)],new Date(2026,9,17));
  assert.equal(calls(),1);assert.equal(result.sheets,2);assert.equal(result.count,2);
  assert.equal(result.name,'10월 원가경쟁력자료 대표 모델_2026.10.17.xlsx');
  const book=new ExcelJS.Workbook();await book.xlsx.load(result.data);
  assert.deepEqual(book.worksheets.map(s=>s.name),['MODEL-A','MODEL-B']);
  for(const [i,sheet] of book.worksheets.entries()){
    assert.equal(sheet.getCell('H4').value,'9월');
    assert.equal(sheet.getCell('K4').value,'10월');
    assert.equal(sheet.getCell('N4').value,'11월');
    assert.equal(sheet.getCell('H5').value.result,100*(i+2));
    assert.equal(sheet.getCell('K5').value.result,2*1360*(i+2));
    assert.equal(sheet.getCell('H6').value.formula,'SUM(H5:H5)');
    assert.equal(sheet.getCell('D5').value,'최신 거래처');
    assert.equal(sheet.sheetProtection.sheet,true);
    assert.notEqual(sheet.getCell('N39').protection?.locked,false);
    assert.equal(sheet.getCell('G5').protection.locked,false);
    assert.ok(sheet.model.merges.length>0);
  }
});
test('lookup month headings roll over into the next year',async()=>{
  const {build}=setup();
  const result=await build([file('MODEL_BOM.xlsx')],new Date(2026,11,17));
  assert.equal(result.name,'12월 원가경쟁력자료 대표 모델_2026.12.17.xlsx');
  const book=new ExcelJS.Workbook();await book.xlsx.load(result.data);
  const sheet=book.worksheets[0];
  assert.equal(sheet.getCell('H4').value,'11월');
  assert.equal(sheet.getCell('K4').value,'12월');
  assert.equal(sheet.getCell('N4').value,'1월');
});
test('sheet names remain valid and unique after sanitizing and truncating',async()=>{
  const {build}=setup();
  const result=await build([file('A:B_one.xlsx'),file('A?B_two.xlsx'),file('A'.repeat(40)+'_one.xlsx'),file('A'.repeat(40)+'_two.xlsx')]);
  const book=new ExcelJS.Workbook();await book.xlsx.load(result.data);
  const names=book.worksheets.map(s=>s.name);
  assert.equal(new Set(names).size,4);
  assert.ok(names.every(n=>n.length<=31&&!/[\\/*?:\[\]]/.test(n)));
});
test('empty selections fail before fetching prices',async()=>{
  const {build,calls}=setup();await assert.rejects(build([]),/선택/);assert.equal(calls(),0);
});
