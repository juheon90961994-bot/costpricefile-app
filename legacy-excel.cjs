const fs=require('node:fs/promises'),os=require('node:os'),path=require('node:path');
const {execFile}=require('node:child_process');
module.exports=async function convertLegacyExcel(data){
  const folder=await fs.mkdtemp(path.join(os.tmpdir(),'productcost-excel-'));
  const input=path.join(folder,'input.xls'),output=path.join(folder,'output.xlsx'),script=path.join(folder,'convert.ps1');
  try{
    await fs.writeFile(input,Buffer.from(data));
    await fs.copyFile(path.join(__dirname,'convert-legacy-excel.ps1'),script);
    await new Promise((resolve,reject)=>execFile('powershell.exe',['-NoProfile','-NonInteractive','-ExecutionPolicy','Bypass','-File',script,'-InputPath',input,'-OutputPath',output],{windowsHide:true,timeout:120000},error=>error?reject(new Error('XLS 서식 변환에 실패했습니다. Microsoft Excel 설치 및 파일 상태를 확인해 주세요.')):resolve()));
    return new Uint8Array(await fs.readFile(output));
  }finally{
    for(const file of [input,output,script])await fs.unlink(file).catch(()=>{});
    await fs.rmdir(folder).catch(()=>{});
  }
};
