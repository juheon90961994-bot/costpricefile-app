(function(root,factory){
  const api=factory();
  if(typeof module==='object'&&module.exports)module.exports=api;
  if(root)root.ebomMbom=api;
})(typeof window!=='undefined'?window:globalThis,function(){
  const TARGET_COLUMNS=[
    'S','단계','표시단계','품목','품목명','BOM구분','규격','세부규격','내외자구분','내외자구분명',
    '생산단위','재고단위','MAKER','팬텀여부','BOM유형','조달구분','정미수량','로스율','실소요량','설계변경일',
    '차수','설계변경코드','누적정미수량','누적실소요량','주거래처코드',null,'주거래처','위치번호','HS CODE',
    '도면번호(품목)','사용자정의상태코드','비고','품목군','사용자정의상태','순번','재질','첨부파일','LOCATION NO',
    '등록일','등록자','계정구분','수정일','수정자','수정자명','상위품목코드','상위품목명','규격(상위품목)',
    '도면번호(상위품목)','중량단위','중량','수입검사레벨','품목사용여부','시작일','종료일','현재고'
  ];
  const REQUIRED_COLUMNS=['품목','품목명','BOM구분','규격'];
  const INPUT_COLUMNS=new Set(['계정구분','현재고']);
  const SOURCE_ALIASES={'계정구분':'등록자명','현재고':'도면번호'};
  const DEFAULT_WIDTHS={spacer:12.42578125,'계정구분':12.42578125,'현재고':10};
  const HEADER_FILL={patternType:'solid',fgColor:{rgb:'969696'},bgColor:{rgb:'000000'}};
  const INPUT_FILL={patternType:'solid',fgColor:{rgb:'FFFF00'},bgColor:{rgb:'FFFFFF'}};

  function clone(value){
    if(value===undefined)return undefined;
    return typeof structuredClone==='function'?structuredClone(value):JSON.parse(JSON.stringify(value));
  }

  function normalizedHeader(value){return String(value??'').normalize('NFKC').trim();}

  function makeHeaderCell(sourceCell,label){
    const cell=clone(sourceCell)||{t:'s'};
    delete cell.f;delete cell.F;
    cell.s=clone(cell.s)||{};
    cell.s.patternType=HEADER_FILL.patternType;
    cell.s.fgColor=clone(HEADER_FILL.fgColor);
    cell.s.bgColor=clone(HEADER_FILL.bgColor);
    if(label===null){delete cell.v;delete cell.w;delete cell.z;cell.t='z';return cell;}
    cell.v=label;cell.w=label;cell.t='s';cell.z='General';
    return cell;
  }

  function makeInputCell(sourceCell){
    const cell=clone(sourceCell)||{t:'z'};
    cell.s=clone(cell.s)||{};
    cell.s.patternType=INPUT_FILL.patternType;
    cell.s.fgColor=clone(INPUT_FILL.fgColor);
    cell.s.bgColor=clone(INPUT_FILL.bgColor);
    if(cell.v===undefined)cell.t='z';
    return cell;
  }

  function transformSheet(source,XLSX){
    if(!source?.['!ref'])throw new Error('첫 번째 시트에 변환할 데이터가 없습니다.');
    const sourceRange=XLSX.utils.decode_range(source['!ref']);
    const sourceIndexes=new Map();
    for(let column=sourceRange.s.c;column<=sourceRange.e.c;column++){
      const cell=source[XLSX.utils.encode_cell({r:sourceRange.s.r,c:column})];
      const header=normalizedHeader(cell?.v);
      if(header&&!sourceIndexes.has(header))sourceIndexes.set(header,column);
    }
    const missing=REQUIRED_COLUMNS.filter(header=>!sourceIndexes.has(header));
    if(missing.length)throw new Error('EBOM 필수 열을 찾을 수 없습니다: '+missing.join(', '));

    const target={};
    for(const [key,value] of Object.entries(source)){
      if(!key.startsWith('!')||['!ref','!cols','!merges','!autofilter'].includes(key))continue;
      target[key]=clone(value);
    }
    target['!rows']=clone(source['!rows']||[]);
    target['!cols']=TARGET_COLUMNS.map(header=>{
      const sourceColumn=header===null?undefined:(sourceIndexes.get(header)??sourceIndexes.get(SOURCE_ALIASES[header]));
      if(sourceColumn!==undefined&&source['!cols']?.[sourceColumn])return clone(source['!cols'][sourceColumn]);
      return{width:DEFAULT_WIDTHS[header===null?'spacer':header]||12.42578125};
    });

    const fallbackHeader=source[XLSX.utils.encode_cell({r:sourceRange.s.r,c:sourceRange.s.c})];
    TARGET_COLUMNS.forEach((header,targetColumn)=>{
      const sourceColumn=header===null?undefined:(sourceIndexes.get(header)??sourceIndexes.get(SOURCE_ALIASES[header]));
      const sourceHeader=sourceColumn===undefined?fallbackHeader:source[XLSX.utils.encode_cell({r:sourceRange.s.r,c:sourceColumn})];
      target[XLSX.utils.encode_cell({r:0,c:targetColumn})]=makeHeaderCell(sourceHeader,header);
      for(let row=1;row<=sourceRange.e.r;row++){
        const targetAddress=XLSX.utils.encode_cell({r:row,c:targetColumn});
        const sourceCell=sourceColumn===undefined?undefined:source[XLSX.utils.encode_cell({r:row,c:sourceColumn})];
        if(INPUT_COLUMNS.has(header)&&row>=2){target[targetAddress]=makeInputCell(sourceCell);continue;}
        if(sourceCell)target[targetAddress]=clone(sourceCell);
        else if(header===null&&row>=2)target[targetAddress]={t:'z'};
      }
    });
    target['!ref']=XLSX.utils.encode_range({s:{r:0,c:0},e:{r:sourceRange.e.r,c:TARGET_COLUMNS.length-1}});
    return target;
  }

  async function convertBytes(bytes,XLSX,ExcelJS){
    if(!XLSX)throw new Error('Excel 변환 기능을 불러오지 못했습니다.');
    if(!ExcelJS)throw new Error('Excel 서식 기능을 불러오지 못했습니다.');
    const sourceBook=XLSX.read(bytes,{type:'array',raw:true,cellDates:true,cellStyles:true,cellFormula:true});
    const source=sourceBook.Sheets[sourceBook.SheetNames[0]];
    const target=transformSheet(source,XLSX);
    const styledBook=new ExcelJS.Workbook();
    const styledSheet=styledBook.addWorksheet('변경');
    const range=XLSX.utils.decode_range(target['!ref']);
    for(let column=0;column<TARGET_COLUMNS.length;column++){
      styledSheet.getColumn(column+1).width=target['!cols']?.[column]?.width||12.42578125;
    }
    for(let row=range.s.r;row<=range.e.r;row++){
      const rowInfo=target['!rows']?.[row];
      if(rowInfo?.hpt)styledSheet.getRow(row+1).height=rowInfo.hpt;
      for(let column=range.s.c;column<=range.e.c;column++){
        const sourceCell=target[XLSX.utils.encode_cell({r:row,c:column})];
        if(!sourceCell)continue;
        const cell=styledSheet.getCell(row+1,column+1);
        if(sourceCell.f)cell.value={formula:sourceCell.f,result:sourceCell.v};
        else if(sourceCell.t!=='z'&&sourceCell.v!==undefined)cell.value=sourceCell.v;
        if(sourceCell.z&&sourceCell.z!=='General')cell.numFmt=sourceCell.z;
      }
    }
    if(target['!margins'])styledSheet.pageSetup.margins=clone(target['!margins']);
    for(let column=1;column<=TARGET_COLUMNS.length;column++){
      styledSheet.getCell(1,column).fill={type:'pattern',pattern:'solid',fgColor:{argb:'FF969696'},bgColor:{argb:'FF000000'}};
    }
    for(const header of INPUT_COLUMNS){
      const column=TARGET_COLUMNS.indexOf(header)+1;
      for(let row=3;row<=styledSheet.rowCount;row++){
        styledSheet.getCell(row,column).fill={type:'pattern',pattern:'solid',fgColor:{argb:'FFFFFF00'},bgColor:{argb:'FFFFFFFF'}};
      }
    }
    return styledBook.xlsx.writeBuffer();
  }

  async function convertFile(file,XLSX,ExcelJS){
    const data=await convertBytes(await file.arrayBuffer(),XLSX,ExcelJS);
    return{data,name:file.name.replace(/\.(xls|xlsx|xlsm)$/i,'')+'_MBOM.xlsx'};
  }

  return{TARGET_COLUMNS,transformSheet,convertBytes,convertFile};
});
