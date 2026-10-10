const SHOP_URL = 'https://shop.g2b.go.kr/';
const https = require('node:https');
const SEARCH_URL = new URL('/gm/gms/gmsd/newShopUntySrchApi.do', SHOP_URL);
const agent = new https.Agent({keepAlive:true});
const pending = new Map();

function requestProducts(model, page) {
  const body=JSON.stringify({searchVO:{searchKeyword:model,selectValue:'',target:'계300001,계300002,계309999',prchsMthdSeCd:'본010001',sortCndt:'SRCH_UPRC_ASC',recordCountPerPage:'100',currentPage:page,srchSeCd:'검030006',rdoIndex:1,dgtlSrvcMallYn:'N',dscplnEvntYn:'N'}});
  return new Promise((resolve,reject)=>{
    const request=https.request(SEARCH_URL,{method:'POST',agent,headers:{'Content-Type':'application/json; charset=UTF-8','Content-Length':Buffer.byteLength(body),Referer:SHOP_URL},timeout:8000},response=>{
      let data='';
      response.setEncoding('utf8');
      response.on('data',chunk=>{data+=chunk;if(data.length>10000000)request.destroy(new Error('상품 조회 응답이 너무 큽니다.'));});
      response.on('error',reject);
      response.on('end',()=>{
        try{
          if(response.statusCode!==200)throw new Error(`나라장터 조회 실패 (${response.statusCode})`);
          const result=JSON.parse(data);
          if(Number(result.ErrorCode)!==0||!Array.isArray(result.rsltList))throw new Error('나라장터 상품 응답 형식을 확인할 수 없습니다.');
          resolve(result);
        }catch(error){reject(error);}
      });
    });
    request.on('timeout',()=>request.destroy(new Error('나라장터 데이터 조회 시간이 초과되었습니다.')));
    request.on('error',reject);
    request.end(body);
  });
}

function parseProducts(rows, model) {
  const found=new Map();
  for(const row of rows){
    const name=String(row.itemIdnfNm??'').replace(/&#(\d+);/g,(_,n)=>String.fromCodePoint(Number(n))).replace(/&amp;/g,'&');
    if(!name.split(',').some(part=>part.trim().toUpperCase()===model))continue;
    const rawPrice=row.srchUprc;
    if(rawPrice===null||rawPrice===undefined||String(rawPrice).trim()==='')throw new Error('상품 판매가격이 없습니다.');
    const price=Number(rawPrice);
    if(!Number.isFinite(price)||price<0)throw new Error('상품 판매가격을 확인할 수 없습니다.');
    const itemId=String(row.itemIdnfNo??'');
    if(!/^\d{8}$/.test(itemId))throw new Error('상품 식별번호를 확인할 수 없습니다.');
    found.set(itemId,{name,price,itemId});
  }
  return [...found.values()];
}

async function lookup(BrowserWindow, model) {
  if(typeof model!=='string'||!/^[A-Z0-9][A-Z0-9._-]{1,79}$/i.test(model))throw new Error('조회할 모델명이 올바르지 않습니다.');
  model=model.toUpperCase();
  if(pending.has(model))return pending.get(model);
  const operation=(async()=>{
    try{
      const rows=[];
      for(let page=1;page<=10;page++){
        const result=await requestProducts(model,page);
        rows.push(...result.rsltList);
        if(rows.length>=Number(result.totalSize)||result.rsltList.length<100){
          return {model,products:parseProducts(rows,model),checkedAt:new Date().toISOString(),source:SHOP_URL};
        }
      }
      throw new Error('검색 결과가 너무 많습니다.');
    }catch(error){
      // Retain the verified browser path if the site's data endpoint changes.
      return browserLookup(BrowserWindow,model);
    }
  })();
  pending.set(model,operation);
  try{return await operation;}finally{pending.delete(model);}
}

// Execute inside an isolated, sandboxed shopping-mall window.
async function searchShop(model, prepareOnly = false) {
  const deadline = Date.now() + 45000;
  let stage = '검색 화면 로딩';
  const wait = async predicate => {
    while (Date.now() < deadline) {
      const value = predicate();
      if (value) return value;
      await new Promise(resolve => setTimeout(resolve, 300));
    }
    throw new Error('나라장터 응답 시간이 초과되었습니다 (' + stage + '). 잠시 후 다시 조회해 주세요.');
  };
  const visible = element => element && element.getClientRects().length > 0;
  await wait(() => document.getElementById('mf_wfm_gnb_wfm_gnbBtm_btnMtSrch'));
  if (prepareOnly) {
    await wait(() => /다수공급자계약 2단계경쟁 정보 조회/.test(document.body.innerText));
    await new Promise(resolve => setTimeout(resolve, 4000));
  }
  await wait(() => !Array.from(document.querySelectorAll('iframe')).some(frame => visible(frame) && frame.src.includes('processMsg')));
  for (const close of document.querySelectorAll('[id$="_close"]')) if (visible(close)) close.click();
  stage = '검색 입력창';
  const input = await wait(() => {
    const field = document.getElementById('mf_wfm_gnb_wfm_gnbBtm_ibxSrchKeyword');
    if (visible(field)) return field;
    document.getElementById('mf_wfm_gnb_wfm_gnbBtm_btnMtSrch')?.click();
    return false;
  });
  // Use the site's own search handler and scoped component. DOM-only input
  // events do not reliably complete WebSquare's asynchronous navigation.
  const searchScope = await wait(() => window.$p?.main?.()?.wfm_gnb?.scope?.wfm_gnbBtm?.scope);
  searchScope.ibxSrchKeyword.setValue(model);
  await searchScope.scwin.btnSrch_onclick();
  stage = '검색 결과';
  await wait(() => {
    const text = document.body.innerText;
    return text.includes(model) && /통합검색결과/.test(text) && (/검색결과\s*\d+\s*건/.test(text) || /쇼핑몰\s*\(\s*0\s*건\s*\)/.test(text));
  });
  // Wait for the product cards as well as the result count.
  stage = '상품 가격';
  let emptySince = 0;
  const products = await wait(() => {
    const body = document.body.innerText;
    if (/검색결과\s*0\s*건/.test(body) || /쇼핑몰\s*\(\s*0\s*건\s*\)/.test(body)) {
      if (!emptySince) emptySince = Date.now();
      if (Date.now() - emptySince > 3000) return [];
      return false;
    }
    emptySince = 0;
    const list = document.querySelector('[id$="_genLstGdsList"]');
    if (!list) return false;
    const images = Array.from(list.querySelectorAll('img[alt]')).filter(image => image.alt.includes(','));
    if (!images.length) return false;
    const found = [];
    for (const image of images) {
      const name = image.alt;
      if (!name.split(',').some(part => part.trim().toUpperCase() === model)) continue;
      let card = image.parentElement;
      while (card && card !== document.body) {
        const text = card.innerText;
        const price = text.match(/(?:^|\n)\s*([\d,]+(?:\.\d+)?)\s*원\s*(?:\n|$)/);
        const id = text.match(/식별번호\s*(\d{8})/);
        if (price && id) {
          found.push({name,price:Number(price[1].replace(/,/g,'')),itemId:id[1]});
          break;
        }
        if (card.id.endsWith('_genLstGdsList')) break;
        card = card.parentElement;
      }
    }
    if (found.length) return Array.from(new Map(found.map(product => [product.itemId,product])).values());
    // Similar names must not be reported as an exact model match.
    if (images.every(image => !image.alt.split(',').some(part => part.trim().toUpperCase() === model))) return [];
    return false;
  });
  return {model, products, checkedAt:new Date().toISOString(),source:location.href};
}

async function browserLookup(BrowserWindow, model) {
  if (typeof model !== 'string' || !/^[A-Z0-9][A-Z0-9._-]{1,79}$/i.test(model)) throw new Error('조회할 모델명이 올바르지 않습니다.');
  const win = new BrowserWindow({show:false,width:1280,height:860,title:'나라장터 판매가 조회',webPreferences:{contextIsolation:true,nodeIntegration:false,sandbox:true,backgroundThrottling:false,partition:'g2b-sales'}});
  win.webContents.setUserAgent(win.webContents.getUserAgent().replace(/\sElectron\/\S+/g,'').replace(/\sstandard-cost-format\/\S+/g,''));
  win.webContents.setWindowOpenHandler(() => ({action:'deny'}));
  let timeout;
  try {
    return await Promise.race([
      (async()=>{
        await win.loadURL(SHOP_URL);
        return await win.webContents.executeJavaScript(`(${searchShop.toString()})(${JSON.stringify(model.toUpperCase())}, true)`);
      })(),
      new Promise((_,reject)=>{timeout=setTimeout(()=>reject(new Error('나라장터 조회 시간이 초과되었습니다. 다시 조회해 주세요.')),60000);}),
    ]);
  } finally {
    clearTimeout(timeout);
    if (!win.isDestroyed()) win.destroy();
  }
}
module.exports = {lookup,searchShop,parseProducts};
