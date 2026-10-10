const test=require('node:test'),assert=require('node:assert/strict'),vm=require('node:vm'),fs=require('node:fs');
function load(){
  const element=()=>({addEventListener(){}});
  const context=vm.createContext({document:{querySelector:element},tabLookup:element(),ITEM_COLUMN:'품목',normalizeItemKey:value=>String(value??'').normalize('NFKC').trim().toUpperCase(),window:{}});
  vm.runInContext(fs.readFileSync(require.resolve('../cost-lookup.js'),'utf8'),context);
  return context;
}
test('model includes first hyphen and stops before second',()=>{
  const c=load();
  assert.equal(c.getSaleModel(' A3TC-5NBDF-001 '),'A3TC-5NBDF');
  assert.equal(c.getSaleModel('a3tc-5nbdf-001-extra'),'A3TC-5NBDF');
  assert.throws(()=>c.getSaleModel('A3TC-5NBDF'));
});
test('button keeps company and product eligibility rules',()=>{
  const c=load();
  assert.equal(c.canLookupSalePrice({회사:'에이텍컴퓨터',품목:'A3TC-5NBDF-001'}),true);
  assert.equal(c.canLookupSalePrice({회사:'다른회사',품목:'A3TC-5NBDF-001'}),false);
  assert.equal(c.canLookupSalePrice({회사:'에이텍컴퓨터',품목:'A3TC-5NBDF'}),false);
});
test('lookup displays price and restores button',async()=>{
  const c=load(),button={};let queried;
  c.window.desktopSales={lookup:async model=>{queried=model;return {products:[{price:1944000,itemId:'25569862'}],checkedAt:new Date().toISOString()};}};
  let message='';c.showLookupMessage=text=>{message=text;};
  await c.lookupSalePrice({회사:'에이텍컴퓨터',품목:'A3TC-5NBDF-001'},button);
  assert.equal(queried,'A3TC-5NBDF');assert.match(message,/1,944,000원/);assert.match(message,/25569862/);assert.equal(button.disabled,false);
});
