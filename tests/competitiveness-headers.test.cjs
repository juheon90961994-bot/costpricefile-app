const test=require('node:test');
const assert=require('node:assert/strict');
const fs=require('node:fs');
const XLSX=require('../xlsx.full.min.js');
const ExcelJS=require('../exceljs.min.js');
const build=Function('XLSX','ExcelJS','ACCOUNTING',
  fs.readFileSync(require.resolve('../competition-format.js'),'utf8')+'\n'+
  fs.readFileSync(require.resolve('../cost-competitiveness.js'),'utf8').split('let competitivenessBusy=')[0]+'\nreturn buildCompetitiveness;'
)(XLSX,ExcelJS,'#,##0');
async function convert(bytes){
  return build({name:'input.xlsx',arrayBuffer:async()=>bytes});
}
test('merged headers select incoming costs and preserve the original sheet',async()=>{
  const sheet=XLSX.utils.aoa_to_sheet([
    ['품목','품목명','기초',null,'입고'],
    [null,null,'직접재료비',null,'수량','직접재료비',null,'노무비','제조경비','경상개발비','감가상각비'],
    [null,null,'금액','금액(외부)',null,'금액','금액(외부)','금액','금액','금액','금액'],
    ['MODEL','제품',999,888,2,100,20,10,30,8,4]
  ]);
  sheet['!merges']=['A1:A3','B1:B3','C1:D1','C2:D2','E1:K1','E2:E3','F2:G2'].map(XLSX.utils.decode_range);
  const book=XLSX.utils.book_new();XLSX.utils.book_append_sheet(book,sheet,'원본');
  const result=await convert(XLSX.write(book,{type:'buffer',bookType:'xlsx'}));
  const output=new ExcelJS.Workbook();await output.xlsx.load(result.data);
  assert.equal(result.count,1);
  assert.equal(result.materialCosts[0].price,60);
  assert.deepEqual(output.getWorksheet('변경본').getRow(4).values.slice(1,10),['MODEL','제품',2,100,20,10,30,8,4]);
  assert.ok(output.getWorksheet('원본').model.merges.includes('E1:K1'));
});
if(process.env.COST_HEADER_INPUT)test('provided workbook converts without changing original cells',async()=>{
  const bytes=fs.readFileSync(process.env.COST_HEADER_INPUT);
  const result=await convert(bytes);
  const original=new ExcelJS.Workbook();await original.xlsx.load(bytes);
  const output=new ExcelJS.Workbook();await output.xlsx.load(result.data);
  assert.ok(result.count>0);
  assert.deepEqual(output.getWorksheet('원본').getSheetValues(),original.worksheets[0].getSheetValues());
  console.log('Converted items:',result.count);
});
