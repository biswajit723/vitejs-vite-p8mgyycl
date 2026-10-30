'use strict';
const fs=require('fs');
const path=require('path');
const target=process.argv[2]||path.join('src','app.jsx');
if(!fs.existsSync(target)){console.error('ERROR: src/app.jsx not found. Run from project root.');process.exit(1);}
const original=fs.readFileSync(target,'utf8');
let code=original;
const MARKER='/* RGT_SUPPORT_CALCULATION_V1 */';
if(code.includes(MARKER)){console.log('OK: Support Validation calculation update is already applied.');process.exit(0);}
const start=code.indexOf('/* RGT_SUPPORT_VALIDATION_FINAL_V1 */');
const next=code.indexOf('function SupportToolsPage',start);
if(start<0||next<0){console.error('SAFETY STOP: current Support Validation block not found. app.jsx was not changed.');process.exit(1);}
const block=`${MARKER}
const SUPPORT_VALIDATION_FILE = '4193-SYS-1501-005_SUPPORT-SUMMARY';
const SUPPORT_LOAD_CASES = ['OPE CIVIL/STR WEIGHT IN "kN"','BLAST WEIGHT IN "kN"','TRANSIT WEIGHT IN "kN"'];
const SUPPORT_DIRECTIONS = ['PORT-STB','VERTICAL','VERTICAL (+VE)','FWD-AFT'];
function svNorm(value){return String(value??'').toUpperCase().replace(/[^A-Z0-9+]+/g,'');}
function svBookName(value){return String(value??'').toLowerCase().replace(/\\.[^.]+$/,'').replace(/[^a-z0-9]+/g,'');}
function svIsWorkbook(entry){const a=svBookName(entry?.file?.name),b=svBookName(SUPPORT_VALIDATION_FILE);return a===b||a.startsWith(b)||a.includes(b);}
function svCell(row,index){return String(row?.[index]??'').trim();}
function svNumber(value){const n=Number(value);return Number.isFinite(n)?n:null;}
function svFind(row,names){const wanted=names.map(svNorm);return (row||[]).findIndex(value=>wanted.includes(svNorm(value)));}
function svHeader(sheet){
  const rows=sheet?.rows||[];
  for(let top=0;top<Math.min(rows.length,80);top++){
    const line=svFind(rows[top],['LINE NUMBER','LINE NO','LINE']);
    const uid=svFind(rows[top],['UID','UNIQUE ID']);
    const node=svFind(rows[top],['NODE','NODE NO']);
    const supportType=svFind(rows[top],['SUPPORT TYPE']);
    if(line<0||uid<0||node<0||supportType<0)continue;
    const axial=svFind(rows[top],['AXIAL DIRECTION','AXIAL']);
    const lateral=svFind(rows[top],['LATERAL DIRECTION','LATERAL']);
    for(let directionRow=top;directionRow<=Math.min(top+3,rows.length-1);directionRow++){
      const groups={};let active='';
      for(let c=0;c<(rows[top]||[]).length;c++){
        const name=svNorm(rows[top][c]);
        if(name.includes('OPECIVIL')&&name.includes('WEIGHT'))active='OPE';
        else if(name.includes('BLASTWEIGHT')&&!name.includes('DISP'))active='BLAST';
        else if(name.includes('TRANSITWEIGHT'))active='TRANSIT';
        else if(name&&['DRYWEIGHTINKN','HYDWEIGHTINKN','SUSWEIGHTINKN','THERMALWEIGHTINKN','OCCASIONALWEIGHTINKN','BLASTDISPWEIGHTINKN','MAXDISPLACEMENTSINMM','GUIDECHECK','STIFFNESSCHECK','FRICTIONCHECK','LIFTCHECK'].some(x=>name.includes(x)))active='';
        if(!active)continue;
        const direction=svNorm(rows[directionRow]?.[c]);
        if(direction==='PORTSTB')groups[active+':PORT-STB']=c;
        else if(direction==='VERTICAL')groups[active+':VERTICAL']=c;
        else if(direction==='VERTICAL+VE')groups[active+':VERTICAL (+VE)']=c;
        else if(direction==='FWDAFT')groups[active+':FWD-AFT']=c;
      }
      if(SUPPORT_LOAD_CASES.every((_,i)=>['OPE','BLAST','TRANSIT'][i]&&SUPPORT_DIRECTIONS.every(d=>d==='VERTICAL (+VE)'&&groups[['OPE','BLAST','TRANSIT'][i]+':'+d]===undefined?true:groups[['OPE','BLAST','TRANSIT'][i]+':'+d]!==undefined))){
        if(['OPE','BLAST','TRANSIT'].every(g=>['PORT-STB','VERTICAL','FWD-AFT'].every(d=>groups[g+':'+d]!==undefined)))return{top,directionRow,line,uid,node,supportType,axial,lateral,groups};
      }
    }
  }
  return null;
}
function svComponents(text){
  return String(text??'').toUpperCase().split('+').map(part=>part.trim()).map(part=>{
    const code=part.match(/^(DS|HD|G|R)\\b/)?.[1];if(!code)return null;
    let direction=code==='R'?'VERTICAL':code==='HD'?'VERTICAL (+VE)':null;
    if(code==='G'||code==='DS'){const m=part.match(/\\((PORT\\s*-?\\s*STB|FWD\\s*-?\\s*AFT)\\)/i);if(m)direction=svNorm(m[1])==='PORTSTB'?'PORT-STB':'FWD-AFT';}
    return direction?{code,direction,label:code+(code==='G'||code==='DS'?' ('+direction+')':'')}:null;
  }).filter(Boolean);
}
function svCalculate(row,header,supportType){
  return svComponents(supportType).map(component=>{
    const values=['OPE','BLAST','TRANSIT'].map(loadCase=>{const column=header.groups[loadCase+':'+component.direction];const value=column===undefined?null:svNumber(row[column]);return{loadCase,value};});
    const available=values.filter(item=>item.value!==null);
    const governing=available.reduce((best,item)=>!best||Math.abs(item.value)>Math.abs(best.value)?item:best,null);
    const kg=governing?Math.round(Math.abs(governing.value)*102):null;
    return{...component,values,governing,kg};
  });
}
function svRows(workbooks){
  const books=(workbooks||[]).filter(svIsWorkbook);const source=books.length?books:(workbooks||[]).filter(book=>(book.sheets||[]).some(svHeader));const out=[];
  for(const book of source)for(const sheet of book.sheets||[]){const header=svHeader(sheet);if(!header)continue;const rows=sheet.rows||[];for(let i=Math.max(header.top,header.directionRow)+1;i<rows.length;i++){const row=rows[i]||[];const item={line:svCell(row,header.line),uid:svCell(row,header.uid),node:svCell(row,header.node),supportType:svCell(row,header.supportType),axial:header.axial>=0?svCell(row,header.axial):'',lateral:header.lateral>=0?svCell(row,header.lateral):'',sequence:out.length+1};if(item.line||item.uid||item.node){item.results=svCalculate(row,header,item.supportType);out.push(item);}}}return out;
}
function SupportValidationConnectedPage({workbooks,title,back}){
  const[query,setQuery]=useState('');const[selected,setSelected]=useState(null);const rows=useMemo(()=>svRows(workbooks),[workbooks]);
  const shown=useMemo(()=>{const q=query.trim().toLowerCase();return q?rows.filter(r=>[r.line,r.uid,r.node].some(v=>v.toLowerCase().includes(q))):rows;},[rows,query]);
  useEffect(()=>{if(selected&&!rows.some(r=>r.sequence===selected.sequence))setSelected(null);},[rows,selected]);
  return <main className="page"><style>{\`.sv-wrap{padding:24px;background:#f3f8fc;min-height:calc(100vh - 120px)}.sv-grid{max-width:1400px;margin:auto;display:grid;grid-template-columns:minmax(0,1fr) 390px;gap:18px}.sv-card{background:#fff;border:1px solid #d4e4ef;border-radius:18px;overflow:hidden;box-shadow:0 16px 40px rgba(16,47,74,.11)}.sv-head{padding:18px 20px;background:#dceaf5}.sv-head h2{margin:0;color:#17324d}.sv-head p{margin:6px 0 0;color:#167fac;font-weight:800}.sv-search{padding:14px 18px;display:flex;gap:12px;align-items:center}.sv-search input{flex:1;padding:12px 14px;border:1px solid #b8cfdf;border-radius:10px}.sv-tablebox{overflow:auto;max-height:650px}.sv-table{width:100%;border-collapse:collapse;min-width:700px}.sv-table th{position:sticky;top:0;background:#1785b9;color:white;padding:12px;text-align:left}.sv-table td{padding:11px 12px;border-bottom:1px solid #e3edf4}.sv-table tbody tr{cursor:pointer}.sv-table tbody tr:hover,.sv-table tbody tr.sel{background:#e6f5fc}.sv-empty{padding:55px 20px;text-align:center;color:#557086}.sv-detail{padding:18px}.sv-detail h3{margin:0 0 15px;color:#17324d}.sv-meta{display:grid;grid-template-columns:1fr 1fr;gap:10px;margin-bottom:14px}.sv-meta div,.sv-result{padding:10px;border-radius:10px;background:#f3f8fc}.sv-meta b,.sv-result small{display:block;color:#5d7689;font-size:11px}.sv-support{padding:12px;margin-bottom:14px;border-radius:10px;background:#173f61;color:#fff;font-weight:800}.sv-result{margin-bottom:10px;border-left:5px solid #1785b9}.sv-result strong{font-size:20px;color:#17324d}.sv-loads{margin-top:7px;font-size:12px;color:#536d80;line-height:1.55}.sv-note{font-size:11px;color:#71889a}@media(max-width:950px){.sv-grid{grid-template-columns:1fr}.sv-detail-card{order:-1}}@media(max-width:700px){.sv-wrap{padding:12px}.sv-search{align-items:stretch;flex-direction:column}}\`}</style><header className="top"><button onClick={back}>Back</button><div><p>SUPPORT LOAD VALIDATION</p><h1>Support Validation</h1><span>{title}</span></div></header><section className="sv-wrap"><div className="sv-grid"><div className="sv-card"><div className="sv-head"><h2>Line Number · UID · Node</h2><p>{SUPPORT_VALIDATION_FILE}</p></div>{rows.length?<><div className="sv-search"><input value={query} onChange={e=>setQuery(e.target.value)} placeholder="Search Line Number, UID or Node"/><b>{shown.length} / {rows.length}</b></div><div className="sv-tablebox"><table className="sv-table"><thead><tr><th>SL.</th><th>LINE NUMBER</th><th>UID</th><th>NODE</th></tr></thead><tbody>{shown.map(r=><tr className={selected?.sequence===r.sequence?'sel':''} onClick={()=>setSelected(r)} key={r.sequence+'|'+r.line+'|'+r.uid+'|'+r.node}><td>{r.sequence}</td><td>{r.line}</td><td>{r.uid}</td><td>{r.node}</td></tr>)}</tbody></table>{!shown.length?<div className="sv-empty">No matching data found.</div>:null}</div></>:<div className="sv-empty">Connected folder-e {SUPPORT_VALIDATION_FILE} ebong required load columns paoa jayni.</div>}</div><aside className="sv-card sv-detail-card"><div className="sv-detail">{selected?<><h3>Calculated Support Loads</h3><div className="sv-meta"><div><b>LINE NUMBER</b>{selected.line}</div><div><b>UID</b>{selected.uid}</div><div><b>NODE</b>{selected.node}</div><div><b>AXIAL / LATERAL</b>{selected.axial||'-'} / {selected.lateral||'-'}</div></div><div className="sv-support">{selected.supportType||'No Support Type'}</div>{selected.results.length?selected.results.map((result,index)=><div className="sv-result" key={result.label+'|'+index}><small>{result.label} · {result.direction}</small><strong>{result.kg===null?'No data':result.kg.toLocaleString()+' kg'}</strong><div className="sv-loads">{result.values.map(v=><div key={v.loadCase}>{v.loadCase}: {v.value===null?'-':v.value+' kN'}</div>)}{result.governing?<b>Governing: {result.governing.loadCase} · |{result.governing.value}| × 102</b>:null}</div></div>):<div className="sv-empty">Support Type-e R, G, HD ba DS component paoa jayni.</div>}<p className="sv-note">R = VERTICAL · HD = VERTICAL (+VE) · G/DS = bracket-er PORT-STB ba FWD-AFT. Blast factor apply hoyni.</p></>:<div className="sv-empty">Calculation dekhte Line Number, UID ba Node row-te click korun.</div>}</div></aside></div></section></main>;
}

`;
code=code.slice(0,start)+block+code.slice(next);
const backup=target+'.before-support-calculation.bak';
if(!fs.existsSync(backup))fs.writeFileSync(backup,original,'utf8');
fs.writeFileSync(target,code,'utf8');
console.log('OK: Support Validation component calculation and click-side-panel added.');
console.log('OK: backup saved at '+backup);
