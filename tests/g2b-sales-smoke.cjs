const {app,BrowserWindow}=require('electron');
const fs=require('node:fs');
const path=require('node:path');
app.on('window-all-closed',()=>{});
app.whenReady().then(async()=>{
  try{
    fs.writeFileSync(path.join(__dirname,'g2b-sales-result.json'),JSON.stringify({status:'running',at:new Date().toISOString()}));
    const started=Date.now();
    const result=await require('../g2b-sales.cjs').lookup(BrowserWindow,'A3TC-5NBDF');
    result.elapsedMs=Date.now()-started;
    if(!result.products.length||!result.products.every(p=>Number.isFinite(p.price)&&p.price>0))throw Error('판매가격을 확인하지 못했습니다.');
    fs.writeFileSync(path.join(__dirname,'g2b-sales-result.json'),JSON.stringify(result,null,2));
    setTimeout(()=>app.exit(0),1000);
  }catch(error){fs.writeFileSync(path.join(__dirname,'g2b-sales-result.json'),JSON.stringify({error:error.message,pageText:error.pageText}));setTimeout(()=>app.exit(1),1000);}
});
