const test=require('node:test');
const assert=require('node:assert/strict');
const {parseProducts}=require('../g2b-sales.cjs');
test('only exact models match and the displayed sale price is used',()=>{
  const rows=[
    {itemIdnfNm:'컴퓨터, 에이텍컴퓨터, A3TC-5NBDF, i5&#40;2.5GHz&#41;',srchUprc:'1944000.0',ctrtUprc:2000000,itemIdnfNo:'25569862'},
    {itemIdnfNm:'컴퓨터, 에이텍컴퓨터, A3TC-5NBDF2',srchUprc:'1',itemIdnfNo:'25569863'},
  ];
  const products=parseProducts(rows,'A3TC-5NBDF');
  assert.equal(products.length,1);assert.equal(products[0].price,1944000);assert.match(products[0].name,/i5\(2.5GHz\)/);
});
test('missing and invalid prices cannot be reported as zero',()=>{
  for(const value of [null,undefined,'','NaN',-1])assert.throws(()=>parseProducts([{itemIdnfNm:'컴퓨터, A-B',srchUprc:value,itemIdnfNo:'12345678'}],'A-B'));
});
test('empty results stay empty and duplicate product IDs are deduplicated',()=>{
  assert.deepEqual(parseProducts([],'A-B'),[]);
  const row={itemIdnfNm:'컴퓨터, A-B',srchUprc:100,itemIdnfNo:'12345678'};
  assert.equal(parseProducts([row,row],'A-B').length,1);
});
