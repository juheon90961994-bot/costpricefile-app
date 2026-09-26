// Original Excel files are stored separately from the cost database.
const ExcelArchive = (() => {
  const bucket = 'excel-archive';
  const maxSize = 20 * 1024 * 1024;
  const mime = {xlsx:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',xls:'application/vnd.ms-excel',xlsm:'application/vnd.ms-excel.sheet.macroEnabled'};
  function identity(name) {
    const base=name.normalize('NFC').replace(/\.(xlsx|xls|xlsm)$/i,'');
    const prefix=base.split('_')[0];
    if (!prefix.trim()) throw new Error('_ 앞에 파일 구분 문자열이 있어야 합니다.');
    return prefix;
  }
  async function syncBom(files, onlyKey) {
    const keys=new Set(files.map(row=>identity(row.filename)));
    const settings=getConnectionSettings();
    const records=[];
    // Complete both snapshots before modifying anything; never use the filtered UI list.
    for(let offset=0;;) {
      const url=new URL(settings.url+'/rest/v1/order_mobility');
      url.searchParams.set('select',ORDER_FIELDS.join(','));
      url.searchParams.set('order','품목.asc');
      url.searchParams.set('offset',String(offset));
      url.searchParams.set('limit','500');
      const page=await supabaseRequest(settings,url.toString());
      if(!Array.isArray(page)) throw new Error('BOM 동기화: 품목 목록 응답이 올바르지 않습니다.');
      if(!page.length) break;
      records.push(...page); offset+=page.length;
    }
    let changed=0;
    for(const row of records) {
      const item=String(row['품목']??'').normalize('NFC');
      if(onlyKey!==undefined && item!==onlyKey) continue;
      const value=keys.has(item)?'있음':'없음';
      if(row.BOM===value) continue;
      try { await mutateOrder('PATCH',row,{BOM:value}); changed++; }
      catch(error) { throw new Error('BOM 동기화 중단 ('+changed+'건 반영): '+error.message+' 목록 새로고침으로 다시 시도해 주세요.'); }
    }
    return changed;
  }
  async function deleteFile(path, name) {
    const group=identity(name);
    // Required ordering: update the matching item before deleting the physical file.
    await syncBom([],group);
    try { await remove(path); }
    catch(error) {
      try { await syncBom(await list(),group); }
      catch(recovery) { throw new Error(error.message+' BOM 복구도 실패했습니다: '+recovery.message); }
      throw error;
    }
  }
  async function remove(path) {
    if (!/^files\/[A-Za-z0-9_-]+$/.test(path)) throw new Error('올바르지 않은 보관 파일 경로입니다.');
    await request('object/'+bucket,{method:'DELETE',headers:{'Content-Type':'application/json'},body:JSON.stringify({prefixes:[path]})});
    // Storage can return success with no affected objects when RLS denies deletion.
    if ((await list()).some(row=>row.path===path)) throw new Error('파일이 삭제되지 않았습니다. 보관함 삭제 권한을 확인해 주세요.');
  }
  function key(name) {
    const bytes = new TextEncoder().encode(name.normalize('NFC'));
    if (bytes.length > 600) throw new Error('파일명이 너무 깁니다. 파일명을 줄여 주세요.');
    return 'files/' + btoa(String.fromCharCode(...bytes)).replace(/\+/g,'-').replace(/\//g,'_').replace(/=+$/,'');
  }
  function filename(value) {
    const base = value.replace(/^files\//,'').replace(/-/g,'+').replace(/_/g,'/');
    return new TextDecoder('utf-8',{fatal:true}).decode(Uint8Array.from(atob(base), c=>c.charCodeAt(0)));
  }
  async function request(route, options = {}) {
    const settings = getConnectionSettings();
    const response = await fetch(settings.url + '/storage/v1/' + route, {
      ...options, headers:{apikey:settings.anonKey,...options.headers}, signal:AbortSignal.timeout(120000)
    });
    if (!response.ok) {
      const detail = await response.text();
      if (/bucket.*not found|row.level security|permission|unauthorized/i.test(detail) || [401,403,404].includes(response.status)) {
        throw new Error('엑셀 보관함이 준비되지 않았거나 접근 권한이 없습니다. 관리자에게 보관함 설정을 요청해 주세요.');
      }
      throw new Error(`파일 보관 요청에 실패했습니다 (${response.status}). 잠시 후 다시 시도해 주세요.`);
    }
    return response;
  }
  async function list() {
    const rows = [];
    for (let offset=0;;) {
      const page = await (await request('object/list/'+bucket,{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({prefix:'files',limit:100,offset,sortBy:{column:'name',order:'asc'}})})).json();
      if (!Array.isArray(page)) throw new Error('보관 목록 응답을 확인할 수 없습니다.');
      if (!page.length) break;
      for (const row of page) {
        if (!row.id) continue;
        try { rows.push({...row,filename:filename(row.name),path:'files/'+row.name}); } catch { /* Ignore foreign objects. */ }
      }
      offset += page.length;
    }
    return rows.sort((a,b)=>String(b.updated_at).localeCompare(String(a.updated_at)));
  }
  async function upload(file) {
    const ext = file.name.split('.').pop().toLowerCase();
    if (!mime[ext]) throw new Error('xls, xlsx, xlsm 엑셀 파일만 보관할 수 있습니다.');
    if (!file.size || file.size > maxSize) throw new Error('0바이트보다 크고 20MB 이하인 파일을 선택해 주세요.');
    const group=identity(file.name);
    const path = key(file.name);
    const previous=(await list()).filter(row=>{
      try { return identity(row.filename)===group && row.path!==path; } catch { return false; }
    });
    await request('object/'+bucket+'/'+path,{method:'POST',headers:{'Content-Type':mime[ext],'x-upsert':'true','cache-control':'max-age=0'},body:file});
    // Preserve old files until the new upload has succeeded.
    let cleanupError;
    try { for (const row of previous) await remove(row.path); }
    catch(error) { cleanupError=error; }
    try { await syncBom(await list()); }
    catch(error) { throw new Error('파일 보관은 완료되었지만 '+error.message); }
    if(cleanupError) throw new Error('새 파일 보관 및 BOM 동기화는 완료되었지만 기존 파일 정리에 실패했습니다. '+cleanupError.message);
    return path;
  }
  async function readForItem(item) {
    const value=String(item??'').normalize('NFC');
    if(!value.trim()) throw new Error('선택한 자료에 품목 값이 없습니다.');
    const matches=(await list()).filter(row=>{
      try{return identity(row.filename)===value;}catch{return false;}
    });
    if(!matches.length) throw new Error(value+' 품목과 일치하는 보관 엑셀 파일이 없습니다.');
    if(matches.length>1) throw new Error(value+' 품목의 보관 파일이 여러 개입니다. 보관함에서 중복 파일을 정리한 뒤 다시 시도해 주세요.');
    const row=matches[0];
    const bytes=await (await request('object/'+bucket+'/'+row.path+'?t='+Date.now())).arrayBuffer();
    return new File([bytes],row.filename);
  }
  async function download(path, name) {
    const bytes = new Uint8Array(await (await request('object/'+bucket+'/'+path+'?t='+Date.now())).arrayBuffer());
    return window.desktopFile.saveExcel(bytes, name);
  }
  return {identity,key,filename,list,upload,download,remove,syncBom,deleteFile,readForItem};
})();

(() => {
  const input = document.getElementById('archiveFile');
  const message = document.getElementById('archiveMessage');
  const body = document.getElementById('archiveBody');
  const search = document.getElementById('archiveSearch');
  let rows = [], busy = false;
  const controls = [...document.querySelectorAll('#excelArchive input, #excelArchive button')];
  function render() {
    body.replaceChildren();
    const visible = rows.filter(row=>row.filename.toLocaleLowerCase().includes(search.value.toLocaleLowerCase()));
    for (const row of visible) {
      const tr=document.createElement('tr');
      for (const value of [row.filename,`${(Number(row.metadata?.size || 0)/1024).toLocaleString('ko-KR',{maximumFractionDigits:1})} KB`,row.updated_at ? new Date(row.updated_at).toLocaleString('ko-KR') : '']) {
        const td=document.createElement('td'); td.textContent=value; tr.append(td);
      }
      const td=document.createElement('td'), button=document.createElement('button');
      button.type='button'; button.textContent='다운로드'; button.disabled=busy;
      button.addEventListener('click',()=>run(async()=>{
        message.textContent='파일을 내려받는 중입니다…';
        const result=await ExcelArchive.download(row.path,row.filename);
        message.textContent=result.canceled ? '저장을 취소했습니다.' : `${row.filename} 파일을 저장했습니다.`;
      }));
      const deleteButton=document.createElement('button');
      deleteButton.type='button'; deleteButton.textContent='삭제'; deleteButton.className='lookup-delete'; deleteButton.disabled=busy;
      deleteButton.addEventListener('click',()=>{
        if (!window.confirm(row.filename+' 파일을 보관함에서 삭제할까요? 삭제한 파일은 복구할 수 없습니다.')) return;
        run(async()=>{
          message.textContent='파일을 삭제하는 중입니다…';
          await ExcelArchive.deleteFile(row.path,row.filename);
          rows=rows.filter(item=>item.path!==row.path);
          message.textContent=row.filename+' 파일을 삭제했습니다.';
        });
      });
      td.append(button,deleteButton); tr.append(td); body.append(tr);
    }
    document.getElementById('archiveEmpty').hidden=visible.length>0;
  }
  async function run(action) {
    if (busy) return;
    busy=true; controls.forEach(el=>el.disabled=true); render();
    try { await action(); } catch(error) { message.textContent=error.message; }
    finally { busy=false; controls.forEach(el=>el.disabled=false); render(); }
  }
  document.getElementById('archiveRefresh').addEventListener('click',()=>run(async()=>{
    message.textContent='보관 목록을 불러오는 중입니다…';
    rows=await ExcelArchive.list();
    const changed=await ExcelArchive.syncBom(rows);
    message.textContent=`보관 파일 ${rows.length}개 · BOM 동기화 ${changed}건 변경`;
  }));
  document.getElementById('archiveUpload').addEventListener('click',()=>run(async()=>{
    const file=input.files[0];
    if (!file) throw new Error('보관할 엑셀 파일을 선택해 주세요.');
    message.textContent='파일을 보관하는 중입니다…';
    try { await ExcelArchive.upload(file); }
    catch(error) { try { rows=await ExcelArchive.list(); } catch {} throw error; }
    input.value='';
    message.textContent=`${file.name} 보관 완료. _ 앞의 문자열이 같은 기존 파일은 대체되었습니다.`;
    try { rows=await ExcelArchive.list(); } catch(error) { message.textContent+=' 목록 새로고침 실패: '+error.message; }
  }));
  search.addEventListener('input',render);
  // Do not load remotely until explicitly requested, keeping cost upload independent.
})();
