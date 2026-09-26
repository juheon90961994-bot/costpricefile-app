const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');

function load(request = async () => []) {
  const element = () => ({ addEventListener() {}, value: '', hidden: false });
  const context = vm.createContext({
    URL, AbortController, setTimeout, clearTimeout,
    document: { querySelector: element }, window: {},
    getConnectionSettings: () => ({url:'https://test.supabase.co'}),
    supabaseRequest: request,
  });
  vm.runInContext(fs.readFileSync(path.join(__dirname, '../order-lookup.js'), 'utf8'), context);
  return context;
}
const row = {품목:'abc-001',품목명:'Part',프로젝트명:'별내선',세부프로젝트명:'발매기',BOM:null};

test('partial matching searches all three fields case-insensitively', () => {
  const c = load();
  for (const keyword of ['ABC','별내','매기']) assert.equal(c.matchesOrderKeyword(row,keyword,''),true);
  assert.equal(c.matchesOrderKeyword(row,'Part',''),false);
  assert.equal(c.matchesOrderKeyword(row,'별내','품목'),false);
  assert.equal(c.matchesOrderKeyword(row,'매기','세부프로젝트명'),true);
});

test('search filters quote punctuation and escape SQL wildcards', () => {
  const c = load();
  const url = new URL(c.orderSearchUrl({url:'https://test.supabase.co'}, 'a,b)"%_\\', '', 500));
  assert.equal(url.pathname,'/rest/v1/order_mobility');
  assert.equal(url.searchParams.get('limit'),'500');
  assert.equal(url.searchParams.get('offset'),'500');
  assert.equal((url.searchParams.get('or').match(/\.ilike\./g)||[]).length,3);
  assert.ok(url.searchParams.get('or').includes('\\"'));
  assert.ok(url.searchParams.get('or').includes('\\\\%'));
  assert.equal(url.searchParams.get('select'),'품목,품목명,프로젝트명,세부프로젝트명,BOM');
});

test('pagination includes every match even if server caps below requested page size', async () => {
  const offsets=[];
  const c=load(async (_settings,url)=>{
    const offset=Number(new URL(url).searchParams.get('offset')); offsets.push(offset);
    return offset<3 ? [{...row,품목:'abc-'+offset}] : [];
  });
  const result=await c.fetchOrderMatches({url:'https://test.supabase.co'},'abc','');
  assert.equal(result.length,3);
  assert.deepEqual(offsets,[0,1,2,3]);
});

test('literal star does not admit unrelated wildcard results', () => {
  const c=load();
  assert.equal(c.matchesOrderKeyword(row,'abc*',''),false);
  assert.equal(c.matchesOrderKeyword({...row,품목:'abc*001'},'abc*',''),true);
});

test('mutation matches all original fields and nulls, never item alone', () => {
  const c=load();
  const url=c.orderMutationUrl({url:'https://test.supabase.co'},row);
  assert.equal([...url.searchParams].length,5);
  assert.equal(url.searchParams.get('프로젝트명'),'eq.별내선');
  assert.equal(url.searchParams.get('BOM'),'is.null');
  assert.throws(()=>c.orderMutationUrl({url:'https://test.supabase.co'},{품목:'abc'}));
});

test('duplicate match blocks PATCH and DELETE before either is sent', async () => {
  for (const method of ['PATCH','DELETE']) {
    let calls=0;
    const c=load(async (_settings,_url,options)=>{calls++; assert.equal(options,undefined); return [row,row];});
    await assert.rejects(c.mutateOrder(method,row,{품목명:'new'}),/여러 건/);
    assert.equal(calls,1);
  }
});

test('PATCH carries atomic single-row limit and exact original locator', async () => {
  const calls=[];
  const c=load(async (_settings,url,options)=>{
    calls.push({url,options}); return options ? [{...row,품목명:'new'}] : [row];
  });
  const result=await c.mutateOrder('PATCH',row,{품목명:'new'});
  assert.equal(result.품목명,'new');
  assert.equal(calls[1].options.headers.Prefer,'handling=strict,max-affected=1,return=representation');
  assert.equal(new URL(calls[1].url).searchParams.get('품목명'),'eq.Part');
  assert.deepEqual(JSON.parse(calls[1].options.body),{품목명:'new'});
});

test('zero affected rows or unsupported server protection never reports success', async () => {
  const c=load(async (_settings,_url,options)=> options ? [] : [row]);
  await assert.rejects(c.mutateOrder('DELETE',row),/처리된 자료가 없습니다/);
  assert.match(c.orderErrorMessage({code:'PGRST122'}),/중단/);
  assert.match(c.orderErrorMessage({status:403,message:'denied'}),/권한/);
});

test('editor preserves unchanged nulls, supports multiline BOM and numeric zero', () => {
  const c=load();
  const input=(field,value)=>({dataset:{orderField:field},value});
  const payload=c.orderEditPayload({...row,BOM:10},[input('품목','abc-001'),input('BOM','0')]);
  assert.equal(payload.BOM,0);
  assert.equal(Object.keys(payload).length,1);
  assert.equal(Object.keys(c.orderEditPayload(row,[input('BOM','')])).length,0);
  assert.equal(c.orderEditPayload(row,[input('BOM','line1\nline2')]).BOM,'line1\nline2');
});

test('structured BOM stays structured and rejects invalid JSON', () => {
  const c=load();
  const input={dataset:{orderField:'BOM'},value:'[{"item":"new"}]'};
  assert.equal(c.orderEditPayload({...row,BOM:[]},[input]).BOM[0].item,'new');
  input.value='bad JSON';
  assert.throws(()=>c.orderEditPayload({...row,BOM:[]},[input]),/JSON/);
});
