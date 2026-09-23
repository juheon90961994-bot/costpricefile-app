"use client";
import { useMemo, useRef, useState } from "react";
type Status = "idle" | "ready" | "working" | "done" | "error";
const steps = ["불필요한 행과 열 정리", "원가표준 열 순서로 재배치", "담당자·단가·합계 항목 생성", "합계 수식과 보기 설정 적용"];
const outputName = (name:string) => `${name.replace(/\.(xlsx|xlsm)$/i, "")}_표준양식.xlsx`;

export default function Home() {
  const inputRef = useRef<HTMLInputElement>(null);
  const [file,setFile] = useState<File|null>(null);
  const [status,setStatus] = useState<Status>("idle");
  const [message,setMessage] = useState("엑셀 파일을 선택해 주세요.");
  const [dragging,setDragging] = useState(false);
  const size = useMemo(()=>file ? `${(file.size/1024).toFixed(file.size>102400?0:1)} KB` : "",[file]);
  const selectFile=(next?:File)=>{if(!next)return;if(!/\.(xlsx|xlsm)$/i.test(next.name)){setStatus("error");setMessage(".xlsx 또는 .xlsm 파일만 선택할 수 있어요.");return;}setFile(next);setStatus("ready");setMessage("변환할 준비가 되었습니다.");};
  const convert=async()=>{if(!file)return;setStatus("working");setMessage("표준 양식으로 변환하고 있어요…");try{
    const ExcelJS=(await import("exceljs")).default;const workbook=new ExcelJS.Workbook();await workbook.xlsx.load(await file.arrayBuffer());const sheet=workbook.worksheets[0];if(!sheet)throw new Error("워크시트가 없습니다.");
    sheet.spliceRows(1,0,[]);sheet.spliceRows(3,1);sheet.spliceColumns(1,1);sheet.spliceColumns(2,1);sheet.spliceColumns(6,1);sheet.spliceColumns(7,6);sheet.spliceColumns(8,9);sheet.spliceColumns(9,28);
    sheet.spliceColumns(9,0,[],[]);sheet.getColumn(7).eachCell({includeEmpty:true},(cell,row)=>{const t=sheet.getCell(row,10);t.value=cell.value;t.style={...cell.style};});
    sheet.spliceColumns(11,0,[]);sheet.getColumn(4).eachCell({includeEmpty:true},(cell,row)=>{const t=sheet.getCell(row,11);t.value=cell.value;t.style={...cell.style};});sheet.spliceColumns(4,1);
    sheet.spliceColumns(9,0,[]);sheet.getColumn(4).eachCell({includeEmpty:true},(cell,row)=>{const t=sheet.getCell(row,9);t.value=cell.value;t.style={...cell.style};});sheet.spliceColumns(4,1);
    sheet.getCell("G2").value="담당자";sheet.getCell("L2").value="단가";sheet.getCell("M2").value="합계";sheet.getCell("L3").value=null;
    for(let row=3;row<=Math.max(70,sheet.rowCount);row++)sheet.getCell(row,13).value={formula:`F${row}*L${row}`};
    sheet.getColumn(4).width=57.57;sheet.getColumn(3).width=58.29;[1,5,6,9,10].forEach(c=>sheet.getColumn(c).alignment={horizontal:"center",vertical:"middle"});
    [2,5,6].forEach(c=>{let longest=8;sheet.getColumn(c).eachCell({includeEmpty:false},cell=>{longest=Math.max(longest,String(cell.text??"").length+2);});sheet.getColumn(c).width=Math.min(longest,40);});
    sheet.views=[{state:"frozen",xSplit:8,ySplit:2,topLeftCell:"I3",activeCell:"I3"}];
    const data=await workbook.xlsx.writeBuffer();const blob=new Blob([data as BlobPart],{type:"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"});const url=URL.createObjectURL(blob);const a=document.createElement("a");a.href=url;a.download=outputName(file.name);a.click();setTimeout(()=>URL.revokeObjectURL(url),1000);setStatus("done");setMessage("변환이 완료되어 다운로드를 시작했습니다.");
  }catch(error){setStatus("error");setMessage(error instanceof Error?`변환 실패: ${error.message}`:"파일을 변환하지 못했습니다.");}};
  return <main><nav><div className="brand"><span className="mark">W</span><span>원가표준</span></div><span className="privacy">● 내 브라우저에서 안전하게 처리</span></nav>
    <section className="hero"><p className="eyebrow">COST SHEET STANDARDIZER</p><h1><em>표준엑셀원가</em><br/>양식으로 변환</h1><p className="lead">원본 엑셀을 올리면 지정된 열 정리, 재배치, 합계 수식까지 자동으로 적용한 새 파일을 만들어 드립니다.</p></section>
    <section className="workspace"><div className="upload-card"><div className="card-head"><span>01</span><div><h2>원본 파일 올리기</h2><p>첫 번째 워크시트에 변환 규칙을 적용합니다.</p></div></div>
      <button className={`dropzone ${dragging?"dragging":""}`} onClick={()=>inputRef.current?.click()} onDragOver={e=>{e.preventDefault();setDragging(true)}} onDragLeave={()=>setDragging(false)} onDrop={e=>{e.preventDefault();setDragging(false);selectFile(e.dataTransfer.files[0])}}><input ref={inputRef} type="file" accept=".xlsx,.xlsm" onChange={e=>selectFile(e.target.files?.[0])}/><span className="file-icon">X</span><strong>{file?file.name:"엑셀 파일을 이곳에 놓으세요"}</strong><small>{file?size:"또는 클릭하여 파일 선택 · XLSX, XLSM"}</small></button>
      <div className={`status ${status}`}><span>{status==="working"?"↻":status==="done"?"✓":status==="error"?"!":"i"}</span>{message}</div><button className="convert" disabled={!file||status==="working"} onClick={convert}>{status==="working"?"변환 중…":"표준 양식으로 변환"}<b>→</b></button></div>
      <aside><div className="card-head"><span>02</span><div><h2>자동 변환 항목</h2><p>제공한 매크로 규칙을 반영합니다.</p></div></div><ol>{steps.map((step,i)=><li key={step}><span>{String(i+1).padStart(2,"0")}</span><p>{step}</p><b>✓</b></li>)}</ol><div className="note"><strong>개인정보 보호</strong><p>파일은 외부 서버로 전송되지 않습니다. 모든 변환은 현재 브라우저 안에서 이루어집니다.</p></div></aside></section>
    <footer><span>원가표준 변환기</span><span>빠르고 정확한 엑셀 정리</span></footer></main>;
}
