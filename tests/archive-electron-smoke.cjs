const {app,session,ipcMain}=require('electron');
const path=require('node:path');
app.setPath('userData',path.join(__dirname,'../work/archive-smoke'));
const timeout=setTimeout(()=>app.exit(1),45000);
app.whenReady().then(()=>session.defaultSession.webRequest.onBeforeRequest({urls:['https://*/*','http://*/*']},(_,cb)=>cb({cancel:true})));
app.on('browser-window-created',(_,win)=>win.webContents.once('did-finish-load',async()=>{
 try {
  let saved;
  ipcMain.removeHandler('save-excel-file');
  ipcMain.handle('save-excel-file',(_,data)=>{saved=data; return {canceled:false};});
  const result=await win.webContents.executeJavaScript(`(async()=>{
   const check=(ok,msg)=>{if(!ok)throw Error(msg)};
   const bomRows=[{'품목':'ABC',BOM:'없음'}];
   supabaseRequest=async (_settings,url)=>Number(new URL(url).searchParams.get('offset')||0)===0 ? bomRows.map(row=>({...row})) : [];
   mutateOrder=async (_method,row,payload)=>Object.assign(bomRows.find(item=>item['품목']===row['품목']),payload);
   const objects=new Map(); let denied=false, denyDelete=false, denyUpload=false;
   window.fetch=async(url,options={})=>{
    check(url.includes('/storage/v1/'),'Storage endpoint');
    if(denied)return new Response('row-level security',{status:403});
    if(url.includes('/object/list/')){
     const {offset}=JSON.parse(options.body);
     return Response.json([...objects].map(([key,value])=>({name:key.split('/').pop(),id:key,metadata:{size:value.length},updated_at:new Date().toISOString()})).slice(offset,offset+1));
    }
    if(options.method==='DELETE') {
     if(!denyDelete)for(const p of JSON.parse(options.body).prefixes)objects.delete(p);
     return Response.json([]);
    }
    const key=url.split('/object/excel-archive/')[1].split('?')[0];
    if(options.method==='POST'){
     if(denyUpload)return new Response('denied',{status:403});
     check(options.headers['x-upsert']==='true','Atomic overwrite');
     objects.set(key,new Uint8Array(await options.body.arrayBuffer()));return Response.json({Key:key});
    }
    return new Response(objects.get(key));
   };
   document.querySelector('#tabUpload').click();
   check(document.querySelector('#tabUpload').textContent==='업로드','Upload title');
   check(!document.querySelector('#uploadWorkspace').hidden,'Upload visible');
   check(document.querySelector('#archiveTitle').textContent==='엑셀 보관함','Archive visible');
   const name='별내선 제품원가.xlsx';
   check(ExcelArchive.filename(ExcelArchive.key(name))===name,'Korean filename roundtrip');
   check(ExcelArchive.key(name.normalize('NFD'))===ExcelArchive.key(name),'Normalized identity');
   await ExcelArchive.upload(new File(['old'],name));
   await ExcelArchive.upload(new File(['replacement'],name));
   check(objects.size===1,'Same filename replaced');
   await ExcelArchive.upload(new File(['other'],'다른 파일.xlsm'));
   check((await ExcelArchive.list()).length===2,'Capped pagination includes all');
   await ExcelArchive.download(ExcelArchive.key(name),name);
   let rejected=false;try{await ExcelArchive.upload(new File(['x'],'wrong.txt'))}catch{rejected=true}
   check(rejected,'Reject non Excel');
   denied=true; rejected=false;try{await ExcelArchive.upload(new File(['bad'],name))}catch{rejected=true}
   check(rejected && new TextDecoder().decode(objects.get(ExcelArchive.key(name)))==='replacement','Failed upload preserves file');
   denied=false;
   document.querySelector('#archiveRefresh').click();
   for(let i=0;i<50 && document.querySelector('#archiveRefresh').disabled;i++) await new Promise(r=>setTimeout(r,20));
   check(document.querySelectorAll('#archiveBody tr').length===2,'List renders');
   document.querySelector('#archiveSearch').value='별내';
   document.querySelector('#archiveSearch').dispatchEvent(new Event('input'));
   check(document.querySelectorAll('#archiveBody tr').length===1,'Partial filename search');
   check(ExcelArchive.identity('ABC_old.xlsx')===ExcelArchive.identity('ABC_new.xlsm'),'Prefix ignores suffix and extension');
   await ExcelArchive.upload(new File(['old'],'ABC_old.xlsx'));
   await ExcelArchive.upload(new File(['new'],'ABC_new.xlsm'));
   check(!objects.has(ExcelArchive.key('ABC_old.xlsx')) && objects.has(ExcelArchive.key('ABC_new.xlsm')),'Different suffix replaces old');
   check(objects.has(ExcelArchive.key(name)),'Unrelated files preserved');
   denyUpload=true;
   try{await ExcelArchive.upload(new File(['bad'],'ABC_failed.xlsx'))}catch{}
   check(objects.has(ExcelArchive.key('ABC_new.xlsm')),'Upload failure retains original');
   denyUpload=false;denyDelete=true;
   let failure='';try{await ExcelArchive.remove(ExcelArchive.key('ABC_new.xlsm'))}catch(e){failure=e.message}
   check(failure.includes('삭제되지'),'Silent RLS no-op detected');
   failure='';try{await ExcelArchive.upload(new File(['latest'],'ABC_latest.xlsx'))}catch(e){failure=e.message}
   check(failure.includes('새 파일 보관') && objects.has(ExcelArchive.key('ABC_new.xlsm')),'Partial replacement explained');
   denyDelete=false;
   await ExcelArchive.upload(new File(['latest'],'ABC_latest.xlsx'));
   check(!objects.has(ExcelArchive.key('ABC_new.xlsm')),'Retry cleans legacy file');
   document.querySelector('#archiveSearch').value='ABC';
   document.querySelector('#archiveRefresh').click();
   const idle=async()=>{for(let i=0;i<100 && document.querySelector('#archiveRefresh').disabled;i++)await new Promise(r=>setTimeout(r,20));};
   await idle();
   window.confirm=()=>false;
   document.querySelector('#archiveBody .lookup-delete').click();await idle();
   check(objects.has(ExcelArchive.key('ABC_latest.xlsx')),'Cancel preserves file');
   window.confirm=()=>true;
   document.querySelector('#archiveBody .lookup-delete').click();await idle();
   check(!objects.has(ExcelArchive.key('ABC_latest.xlsx')),'Delete button removes selected file');
   check(document.querySelectorAll('#archiveBody tr').length===0,'Deleted row removed');
   return 'PASS: prefix replacement, legacy cleanup, failed upload, denied delete, partial failure, retry, delete confirmation and download';
  })()`);
  if(saved.defaultName!=='별내선 제품원가.xlsx'||Buffer.from(saved.data).toString()!=='replacement')throw Error('Downloaded bytes mismatch');
  console.log(result);clearTimeout(timeout);app.exit(0);
 }catch(error){console.error(error);app.exit(1)}
}));
require('../main.cjs');
