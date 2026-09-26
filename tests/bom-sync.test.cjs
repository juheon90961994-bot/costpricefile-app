const {test}=require('node:test'),assert=require('node:assert/strict'),vm=require('node:vm'),fs=require('node:fs');
const source=fs.readFileSync(require('node:path').join(__dirname,'../excel-archive.js'),'utf8').split('\n(() => {')[0];
function setup(){
 const rows=[{품목:'ABC',BOM:'없음'},{품목:'DEF',BOM:'있음'},{품목:'XYZ',BOM:'없음'}], files=new Map(),events=[];
 const state={denyPatch:false,denyDelete:false};
 const ctx=vm.createContext({URL,TextEncoder,TextDecoder,Uint8Array,AbortSignal,Response,btoa,atob,
 ORDER_FIELDS:['품목','BOM'],getConnectionSettings:()=>({url:'https://test.invalid',anonKey:'test'}),
 supabaseRequest:async(_,url)=>{const offset=Number(new URL(url).searchParams.get('offset'));return rows.slice(offset,offset+1).map(r=>({...r}));},
 mutateOrder:async(_,row,payload)=>{if(state.denyPatch)throw Error('denied');events.push('patch:'+row.품목+':'+payload.BOM);Object.assign(rows.find(r=>r.품목===row.품목),payload);},
 fetch:async(url,options)=>{
  if(url.includes('/object/list/')){const {offset}=JSON.parse(options.body);return Response.json([...files.keys()].slice(offset,offset+1).map(path=>({id:path,name:path.slice(6)})));}
  if(options.method==='DELETE'){events.push('delete');if(!state.denyDelete)for(const p of JSON.parse(options.body).prefixes)files.delete(p);return Response.json([]);}
  if(options.method==='POST'){events.push('upload');files.set(url.split('/object/excel-archive/')[1],true);return Response.json({});}
 }});
 vm.runInContext(source+'\nthis.api=ExcelArchive;',ctx);
 return {api:ctx.api,rows,files,events,state};
}
test('full list synchronizes present/absent and skips unchanged items',async()=>{
 const {api,rows,events}=setup();assert.equal(await api.syncBom([{filename:'ABC_new.xlsx'}]),2);
 assert.deepEqual(rows.map(r=>r.BOM),['있음','없음','없음']);assert.equal(events.length,2);
});
test('upload triggers full synchronization',async()=>{
 const {api,rows}=setup();await api.upload({name:'ABC_new.xlsx',size:3});assert.deepEqual(rows.map(r=>r.BOM),['있음','없음','없음']);
});
test('delete updates BOM first and leaves unrelated items unchanged',async()=>{
 const {api,rows,files,events}=setup();rows[0].BOM='있음';const path=api.key('ABC_new.xlsx');files.set(path,true);
 await api.deleteFile(path,'ABC_new.xlsx');assert.deepEqual(events,['patch:ABC:없음','delete']);assert.equal(rows[1].BOM,'있음');assert.equal(files.size,0);
});
test('failed BOM update prevents deletion',async()=>{
 const {api,rows,files,state,events}=setup();rows[0].BOM='있음';state.denyPatch=true;const path=api.key('ABC_new.xlsx');files.set(path,true);
 await assert.rejects(api.deleteFile(path,'ABC_new.xlsx'));assert.equal(files.size,1);assert.ok(!events.includes('delete'));
});
test('denied file deletion restores BOM from remaining files',async()=>{
 const {api,rows,files,state}=setup();rows[0].BOM='있음';state.denyDelete=true;const path=api.key('ABC_new.xlsx');files.set(path,true);
 await assert.rejects(api.deleteFile(path,'ABC_new.xlsx'));assert.equal(rows[0].BOM,'있음');assert.equal(files.size,1);
});
