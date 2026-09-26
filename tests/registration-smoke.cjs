const {app,session}=require('electron');const path=require('node:path');
app.setPath('userData',path.join(__dirname,'../work/registration-smoke'));
setTimeout(()=>app.exit(1),45000);
app.whenReady().then(()=>session.defaultSession.webRequest.onBeforeRequest({urls:['https://*/*','http://*/*']},(_,cb)=>cb({cancel:true})));
app.on('browser-window-created',(_,win)=>win.webContents.once('did-finish-load',async()=>{try{
 console.log(await win.webContents.executeJavaScript(`(async()=>{
 const check=(ok,msg)=>{if(!ok)throw Error(msg)};const records=[];let denied=false,uploads=0;
 ExcelArchive.upload=async()=>{uploads++};ExcelArchive.list=async()=>[];
 supabaseRequest=async(settings,url,options={})=>{
  if(options.method==='POST'){
   if(denied)throw Error('permission denied');
   const row=JSON.parse(options.body);records.push(row);return [row];
  }
  const item=new URL(url).searchParams.get('품목').slice(3);return records.filter(row=>row['품목']===item);
 };
 const tick=()=>new Promise(r=>setTimeout(r,20));
 async function open(name){
  const data=new DataTransfer();data.items.add(new File(['bom'],name));document.querySelector('#archiveFile').files=data.files;
  document.querySelector('#archiveUpload').click();
  for(let i=0;i<50&&!document.querySelector('dialog');i++)await tick();
 }
 document.querySelector('#tabUpload').click();await open('NEW_별내선.xlsx');
 let dialog=document.querySelector('dialog');check(dialog?.open,'New item dialog opens');
 check(dialog.querySelector('[name="품목"]').value==='NEW'&&dialog.querySelector('[name="품목"]').readOnly,'Prefix is read-only');
 for(const field of ['품목명','프로젝트명','세부프로젝트명'])dialog.querySelector('[name="'+field+'"]').value=field+' 테스트';
 denied=true;dialog.querySelector('form').requestSubmit();await tick();
 check(dialog.open&&dialog.textContent.includes('등록 실패')&&records.length===0,'Denied insert retains dialog and input');
 denied=false;dialog.querySelector('form').requestSubmit();
 for(let i=0;i<50&&document.querySelector('dialog');i++)await tick();
 check(records.length===1&&records[0].BOM==='있음'&&records[0]['품목']==='NEW','New record values');
 await open('NEW_변경.xlsx');check(!document.querySelector('dialog')&&records.length===1,'Existing item skips popup');
 await open('CANCEL_보관.xlsx');dialog=document.querySelector('dialog');dialog.querySelector('button[type="button"]').click();await tick();
 check(records.length===1&&uploads===3&&!document.querySelector('dialog'),'Cancel retains file, creates no record');
 let duplicate=false;try{await ExcelArchive.registerItem('NEW',{'품목명':'a','프로젝트명':'b','세부프로젝트명':'c'})}catch{duplicate=true}
 check(duplicate&&records.length===1,'Duplicate registration blocked');
 return 'PASS: missing item popup, prefix, insert fields, BOM, existing item, cancellation, permission retry, duplicate guard';
 })()`));app.exit(0);
 }catch(error){console.error(error);app.exit(1)}}));require('../main.cjs');
