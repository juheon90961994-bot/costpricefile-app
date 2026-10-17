const test=require('node:test'),assert=require('node:assert/strict'),fs=require('node:fs');
function setup(pages){
  const requests=[];
  const archive={list:async()=>[{filename:'MODEL-A_BOM.xlsx'},{filename:'MODEL-B_BOM.xlsx'}],identity:name=>name.split('_')[0]};
  const code=fs.readFileSync(require.resolve('../representative-model.js'),'utf8').split('async function findRepresentativeArchives')[1].split('(() => {')[0];
  const search=Function('ExcelArchive','getConnectionSettings','supabaseRequest','return async function findRepresentativeArchives'+code)(archive,()=>({url:'https://example.test'}),async(_,url)=>{requests.push(new URL(url));return pages[requests.length-1]||[];});
  return {search,requests};
}
for(const field of ['품목','품목명','프로젝트명','세부프로젝트명'])test('partial search in '+field+' maps records to archived BOM',async()=>{
  const row={품목:'MODEL-A',[field]:field==='품목'?'MODEL-A':'분리형 데스크톱'};
  const keyword=field==='품목'?'odel-a':'분리';
  const {search,requests}=setup([[row],[{품목:'MODEL-B',품목명:'다른 제품'}],[]]);
  const result=await search(keyword);
  assert.deepEqual(result.map(r=>r.filename),['MODEL-A_BOM.xlsx']);
  assert.equal(requests.length,3);
  const filter=requests[0].searchParams.get('or');
  for(const name of ['품목','품목명','프로젝트명','세부프로젝트명'])assert.ok(filter.includes(name+'.ilike.'));
});
test('empty query lists files without querying project database',async()=>{
  const {search,requests}=setup([]);assert.equal((await search(' ')).length,2);assert.equal(requests.length,0);
});
test('filename search is retained and duplicate records do not duplicate archive choices',async()=>{
  const {search}=setup([[{품목:'MODEL-A',프로젝트명:'BOM'},{품목:'MODEL-A',프로젝트명:'BOM'}],[]]);
  assert.equal((await search('BOM')).length,2);
});
