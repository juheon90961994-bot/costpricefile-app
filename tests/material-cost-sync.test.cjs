const test=require('node:test');
const assert=require('node:assert/strict');
const fs=require('node:fs');
const vm=require('node:vm');
function setup(rows){
  const element=()=>({value:'',addEventListener(){}});
  const context=vm.createContext({window:{},document:{querySelector:element},localStorage:{getItem(){}},sessionStorage:{getItem(){}},tabUpload:element(),URL});
  vm.runInContext(fs.readFileSync(require.resolve('../cost-upload.js'),'utf8'),context);
  const updates=[];
  context.getConnectionSettings=()=>({});
  context.fetchAllCostRows=async()=>rows;
  context.updateCostRow=async(_,action)=>updates.push(action);
  return {api:context.window.costDatabase,context,updates};
}
test('matches normalized items, preserves decimal/zero prices, changes only KRW price',async()=>{
  const {api,updates}=setup([{id:1,품목:'abc',원화단가:10},{id:2,품목:'ZERO',원화단가:2}]);
  const result=await api.updateMaterialCosts([{item:' ABC ',price:12.5},{item:'abc',price:12.5},{item:'ZERO',price:0}]);
  assert.equal(result.updated,2);
  assert.deepEqual(JSON.parse(JSON.stringify(updates.map(x=>x.payload))),[{원화단가:12.5},{원화단가:0}]);
  assert.equal(updates[0].databaseId,1);
});
test('skips conflicting duplicates, invalid costs, absent items and duplicate DB rows',async()=>{
  const {api,updates}=setup([{품목:'DUP',원화단가:1},{품목:'dup',원화단가:2},{품목:'SAME',원화단가:5}]);
  const result=await api.updateMaterialCosts([{item:'CONFLICT',price:1},{item:'CONFLICT',price:2},{item:'INVALID',price:'수량 확인'},{item:'MISSING',price:3},{item:'DUP',price:4},{item:'SAME',price:5}]);
  assert.equal(result.review,3);assert.equal(result.unmatched,1);assert.equal(result.unchanged,1);assert.equal(updates.length,0);
});
test('reports failures alongside successful updates',async()=>{
  const {api,context}=setup([{품목:'A',원화단가:1},{품목:'B',원화단가:1}]);
  context.updateCostRow=async(_,action)=>{if(action.item==='A')throw Error('denied');};
  const result=await api.updateMaterialCosts([{item:'A',price:2},{item:'B',price:3}]);
  assert.equal(result.failed,1);assert.equal(result.updated,1);assert.match(result.errors[0],/denied/);
});
