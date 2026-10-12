const fs = require('fs');
const path = require('path');

const target = process.argv[2] || path.join(process.cwd(), 'src', 'app.jsx');
const original = fs.readFileSync(target, 'utf8');
let code = original;
const fail = (message) => { throw new Error(message + ' No file was changed.'); };

if (!code.includes('function IsoPlanningDashboardPage(') || !code.includes('function buildIsoPlanningData(')) {
  fail('ISO Planning automation code was not found. Apply the ISO Planning automation updater first.');
}
if (code.includes('const ISO_STAGE_DEFINITIONS =')) {
  console.log('ISO Console completed/work-front automation is already present. Nothing was changed.');
  process.exit(0);
}

// Add the exact nine stage definitions shown in the master Excel.
const holdAliasLine = "const ISO_HOLD_ALIASES = ['hold', 'hold status', 'iso hold', 'on hold'];";
if (!code.includes(holdAliasLine)) fail('ISO HOLD aliases were not found.');
const stageDefinitions = `${holdAliasLine}
const ISO_STAGE_DEFINITIONS = [
  { key: 'isoExtraction', label: 'ISO EXTRACTION', aliases: ['iso extraction', 'isoextraction'] },
  { key: 'layoutCheck', label: 'LAYOUT CHECK', aliases: ['layout check', 'layoutcheck'] },
  { key: 'stressCheck', label: 'STRESS CHECK', aliases: ['stress check', 'stresscheck'] },
  { key: 'supportCheck', label: 'SUPPORT CHECK', aliases: ['support check', 'supportcheck'] },
  { key: 'backDraft', label: 'BACK DRAFT', aliases: ['back draft', 'backdraft'] },
  { key: 'backCheck', label: 'BACK CHECK', aliases: ['back check', 'backcheck'] },
  { key: 'qcCheck', label: 'QC CHECK', aliases: ['qc check', 'qccheck'] },
  { key: 'moduleEngineerCheck', label: 'MODULE ENGINEER CHECK', aliases: ['module engineer check', 'moduleengineercheck'] },
  { key: 'toCp', label: 'TO CP', aliases: ['to cp', 'tocp'] },
];`;
code = code.replace(holdAliasLine, stageDefinitions);

// Add per-sheet stage-column detection.
const holdColumnDeclaration = '      let holdColumn = -1;';
if (!code.includes(holdColumnDeclaration)) fail('ISO header-column declarations were not found.');
code = code.replace(holdColumnDeclaration, holdColumnDeclaration + '\n      let stageColumns = {};');

const holdColumnDetection = '        holdColumn = findColumnByAliases(row, ISO_HOLD_ALIASES);';
if (!code.includes(holdColumnDetection)) fail('ISO HOLD column detection was not found.');
code = code.replace(holdColumnDetection, holdColumnDetection + "\n        stageColumns = Object.fromEntries(ISO_STAGE_DEFINITIONS.map((stage) => [stage.key, findColumnByAliases(row, stage.aliases)]));");

const entryPush = 'moduleRows.get(key).entries.push({ row, sheetEntry, pipeColumn, issuedColumn, revisionColumn, holdColumn });';
if (!code.includes(entryPush)) fail('ISO module entry creation was not found.');
code = code.replace(entryPush, 'moduleRows.get(key).entries.push({ row, sheetEntry, pipeColumn, issuedColumn, revisionColumn, holdColumn, stageColumns });');

// Insert counting, aggregate project-console and filtered-stage download helpers.
const helperAnchor = 'function downloadIsoMetric(moduleData, metric, label) {';
if (!code.includes(helperAnchor)) fail('ISO download helper insertion point was not found.');
const extraHelpers = `function isIsoStageTick(value) {
  if (value === true || value === 1) return true;
  const key = normalize(value);
  return ['true', 'yes', 'y', '1', 'tick', 'checked', 'complete', 'completed'].includes(key) || String(value ?? '').includes('✓') || String(value ?? '').includes('☑');
}
function isoStageRows(moduleData, stageKey, completed) {
  const entries = moduleData?.entries || [];
  return entries.filter((entry) => {
    const columnIndex = entry.stageColumns?.[stageKey] ?? -1;
    if (columnIndex < 0) return false;
    const ticked = isIsoStageTick(entry.row?.[columnIndex]);
    return completed ? ticked : !ticked;
  });
}
function isoStageSummary(moduleData) {
  return Object.fromEntries(ISO_STAGE_DEFINITIONS.map((stage) => [stage.key, {
    completed: isoStageRows(moduleData, stage.key, true).length,
    available: isoStageRows(moduleData, stage.key, false).length,
  }]));
}
function combineIsoModules(modules) {
  return { moduleName: '4193 HAMMER HEAD - ALL MODULES', entries: modules.flatMap((moduleData) => moduleData.entries || []) };
}
function downloadIsoEntries(entries, moduleName, label) {
  if (!entries.length) { window.alert('No ' + label + ' rows found for ' + moduleName + '.'); return; }
  const output = XLSX.utils.book_new();
  const groupedSheets = new Map();
  for (const entry of entries) {
    if (!groupedSheets.has(entry.sheetEntry)) groupedSheets.set(entry.sheetEntry, []);
    groupedSheets.get(entry.sheetEntry).push(entry.row);
  }
  const used = new Set();
  for (const [sheetEntry, rows] of groupedSheets) {
    const headerRowIndex = sheetEntry.header?.rowIndex ?? 0;
    const headers = sheetEntry.rows.slice(0, headerRowIndex + 1);
    const sheet = XLSX.utils.aoa_to_sheet([...headers, ...rows], { cellDates: true });
    const sourceRange = sheetEntry.sheet['!ref'] ? XLSX.utils.decode_range(sheetEntry.sheet['!ref']) : null;
    if (sourceRange) {
      for (let r = sourceRange.s.r; r <= Math.min(headerRowIndex, sourceRange.e.r); r += 1) {
        for (let c = sourceRange.s.c; c <= sourceRange.e.c; c += 1) {
          const address = XLSX.utils.encode_cell({ r, c });
          if (sheetEntry.sheet[address]?.s && sheet[address]) sheet[address].s = JSON.parse(JSON.stringify(sheetEntry.sheet[address].s));
        }
      }
    }
    if (sheetEntry.sheet['!cols']) sheet['!cols'] = sheetEntry.sheet['!cols'].map((item) => item ? { ...item } : item);
    if (sheetEntry.sheet['!rows']) sheet['!rows'] = sheetEntry.sheet['!rows'].map((item) => item ? { ...item } : item);
    if (sheetEntry.sheet['!merges']) sheet['!merges'] = sheetEntry.sheet['!merges'].filter((merge) => merge.e.r <= headerRowIndex).map((merge) => ({ s: { ...merge.s }, e: { ...merge.e } }));
    XLSX.utils.book_append_sheet(output, sheet, safeSheetName(sheetEntry.sheetName, used));
  }
  XLSX.writeFile(output, moduleName + ' ' + label + '.xlsx', { compression: true, cellStyles: true });
}
function downloadIsoStage(moduleData, stage, completed) {
  const stateLabel = completed ? 'Completed' : 'Work Front Available';
  downloadIsoEntries(isoStageRows(moduleData, stage.key, completed), moduleData.moduleName, stage.label + ' ' + stateLabel);
}

`;
code = code.replace(helperAnchor, extraHelpers + helperAnchor);

// Replace only the ISO dashboard component. Everything outside this component is untouched.
const componentStart = code.indexOf('function IsoPlanningDashboardPage(');
const componentEnd = code.indexOf('\nfunction Integrated({ open }) {', componentStart);
if (componentStart < 0 || componentEnd < 0) fail('ISO dashboard component boundaries were not found.');

const component = `function IsoPlanningDashboardPage({ workbooks, back }) {
  const [selectedModule, setSelectedModule] = useState(null);
  const [query, setQuery] = useState('');
  const [consoleMode, setConsoleMode] = useState(false);
  const modules = useMemo(() => buildIsoPlanningData(workbooks), [workbooks]);
  const visibleModules = useMemo(() => modules.filter((item) => item.moduleName.toLowerCase().includes(query.trim().toLowerCase())), [modules, query]);
  const moduleSelection = modules.find((item) => normalize(item.moduleName) === normalize(selectedModule)) || null;
  const projectConsole = useMemo(() => combineIsoModules(modules), [modules]);
  const selected = consoleMode ? projectConsole : moduleSelection;
  const total = isoMetricRows(selected, 'total').length;
  const issued = isoMetricRows(selected, 'issued').length;
  const revision = isoMetricRows(selected, 'revision').length;
  const hold = isoMetricRows(selected, 'hold').length;
  const balance = Math.max(0, total - issued - hold);
  const issuedPercent = total ? Math.round((issued / total) * 100) : 0;
  const revisionRatio = revision ? (total / revision).toFixed(2) : '0';
  const stageSummary = isoStageSummary(selected);
  const chooseModule = (moduleName) => { setSelectedModule(moduleName); setConsoleMode(false); };
  return <Layout title="ISO Planning and Management" sub="Project 4193 · HAMMER HEAD" back={back}>
    <style>{\`
      .iso-plan-layout{display:grid;grid-template-columns:300px minmax(0,1fr);gap:24px}.iso-plan-modules,.iso-console{border:1px solid #d9e7f1;background:linear-gradient(180deg,#fff,#fafdff);box-shadow:0 18px 45px rgba(16,47,74,.11)}.iso-plan-modules{padding:20px;border-radius:20px}.iso-plan-side-head{display:flex;align-items:center;justify-content:space-between;gap:10px}.iso-console-button{padding:10px 15px;border:0;border-radius:10px;background:linear-gradient(135deg,#173f61,#2498cb);color:#fff;font-weight:900;cursor:pointer}.iso-plan-modules input{box-sizing:border-box;width:100%;margin-top:12px;padding:11px 13px;border:1px solid #bfd2e1;border-radius:10px;outline:none}.iso-plan-modules section{display:grid;gap:8px;margin-top:14px;max-height:640px;overflow:auto}.iso-plan-modules section button,.iso-metric{border:0;border-radius:11px;background:linear-gradient(135deg,#2498cb,#1875aa);color:#fff;font-weight:800;cursor:pointer}.iso-plan-modules section button{padding:13px;text-align:left}.iso-plan-modules section button.sel{background:linear-gradient(135deg,#f28a3b,#df642a)}.iso-console{padding:25px;border-radius:22px;overflow:hidden}.iso-console-head{display:flex;align-items:center;justify-content:space-between;gap:28px;flex-wrap:wrap}.iso-console-eyebrow{margin:0;color:#1686bc;font-size:12px;font-weight:900;letter-spacing:.14em}.iso-console-title{margin:6px 0 0;color:#122943;font-size:30px}.iso-progress{width:160px;height:160px;border-radius:50%;display:grid;place-items:center;background:conic-gradient(#09be68 calc(var(--p)*1%),#dceaf5 0);position:relative;box-shadow:0 10px 28px rgba(13,157,91,.18)}.iso-progress:after{content:'';position:absolute;inset:16px;border-radius:50%;background:#fff}.iso-progress-content{position:relative;z-index:1;text-align:center}.iso-progress-content b{display:block;font-size:30px;color:#112944}.iso-progress-content span{font-size:12px;color:#35516b}.iso-metrics{display:grid;grid-template-columns:repeat(5,minmax(120px,1fr));gap:12px;margin-top:25px}.iso-metric{padding:15px;text-align:left;min-height:94px}.iso-metric span{display:block;font-size:12px;letter-spacing:.04em}.iso-metric b{display:block;margin-top:8px;font-size:28px}.iso-stage-card{margin-top:24px;border:1px solid #d4e3ee;border-radius:17px;overflow:auto;background:#fff}.iso-stage-heading{padding:11px 16px;background:#173f61;color:#fff;font-size:13px;font-weight:900;letter-spacing:.12em;text-align:center}.iso-stage-table{width:100%;min-width:980px;border-collapse:collapse}.iso-stage-table th,.iso-stage-table td{padding:10px 8px;border:1px solid #d4e3ee;text-align:center}.iso-stage-table th{background:#dceaf5;color:#163651;font-size:11px}.iso-stage-row-label{min-width:150px!important;background:#f3f8fc!important;font-weight:900!important;color:#173f61!important}.iso-stage-count{border:0;background:transparent;color:#126e9d;font-size:16px;font-weight:900;cursor:pointer}.iso-stage-count:hover{text-decoration:underline}.iso-empty{padding:80px 20px;text-align:center;color:#64748b}@media(max-width:1050px){.iso-plan-layout{grid-template-columns:1fr}.iso-metrics{grid-template-columns:repeat(2,1fr)}}
    \`}</style>
    <div className="iso-plan-layout">
      <aside className="iso-plan-modules"><div className="iso-plan-side-head"><h2>Module List</h2><button className="iso-console-button" onClick={() => { setConsoleMode(true); setSelectedModule(null); }}>CONSOLE</button></div><input value={query} onChange={(event) => setQuery(event.target.value)} placeholder="Search module name"/><section>{visibleModules.map((item) => <button className={!consoleMode && selectedModule === item.moduleName ? 'sel' : ''} key={item.moduleName} onClick={() => chooseModule(item.moduleName)}>{item.moduleName}</button>)}</section></aside>
      <section className="iso-console">{selected && selected.entries.length ? <><div className="iso-console-head"><div><p className="iso-console-eyebrow">{consoleMode ? 'PROJECT CONSOLE · ALL MODULES' : 'SELECTED MODULE'}</p><h2 className="iso-console-title">{selected.moduleName}</h2></div><div className="iso-progress" style={{ '--p': issuedPercent }}><div className="iso-progress-content"><b>{issuedPercent}%</b><span>Revision ratio {revisionRatio}</span></div></div></div><div className="iso-metrics"><button className="iso-metric" onClick={() => downloadIsoMetric(selected,'total','Total Isos')}><span>TOTAL ISOS</span><b>{total}</b></button><button className="iso-metric" onClick={() => downloadIsoMetric(selected,'issued','Issued')}><span>ISSUED</span><b>{issued}</b></button><button className="iso-metric" onClick={() => downloadIsoMetric(selected,'revision','Revision')}><span>REVISION</span><b>{revision}</b></button><button className="iso-metric" onClick={() => downloadIsoMetric(selected,'hold','Hold')}><span>HOLD</span><b>{hold}</b></button><button className="iso-metric" onClick={() => downloadIsoMetric(selected,'balance','Balance Hold Free')}><span>BALANCE HOLD FREE</span><b>{balance}</b></button></div><div className="iso-stage-card"><div className="iso-stage-heading">ISO PROGRESS CONTROL</div><table className="iso-stage-table"><thead><tr><th>STATUS</th>{ISO_STAGE_DEFINITIONS.map((stage) => <th key={stage.key}>{stage.label}</th>)}</tr></thead><tbody><tr><th className="iso-stage-row-label">COMPLETED</th>{ISO_STAGE_DEFINITIONS.map((stage) => <td key={stage.key}><button className="iso-stage-count" onClick={() => downloadIsoStage(selected, stage, true)}>{stageSummary[stage.key]?.completed || 0}</button></td>)}</tr><tr><th className="iso-stage-row-label">WORK FRONT AVAILABLE</th>{ISO_STAGE_DEFINITIONS.map((stage) => <td key={stage.key}><button className="iso-stage-count" onClick={() => downloadIsoStage(selected, stage, false)}>{stageSummary[stage.key]?.available || 0}</button></td>)}</tr></tbody></table></div></> : <div className="iso-empty"><h2>{modules.length ? 'Select a module or open Console' : 'ISO Excel data not found'}</h2><span>{modules.length ? 'Module view shows module-wise counts. Console shows the complete project.' : 'Connect the folder containing ' + ISO_FILE_PREFIX + ' from Reports and Status.'}</span></div>}</section>
    </div>
  </Layout>;
}
`;
code = code.slice(0, componentStart) + component + code.slice(componentEnd);

// Strict verification.
for (const stage of ['ISO EXTRACTION','LAYOUT CHECK','STRESS CHECK','SUPPORT CHECK','BACK DRAFT','BACK CHECK','QC CHECK','MODULE ENGINEER CHECK','TO CP']) {
  if (!code.includes(`label: '${stage}'`)) fail(stage + ' stage was not added.');
}
if (!code.includes('PROJECT CONSOLE · ALL MODULES')) fail('Project Console mode was not added.');
if (!code.includes('WORK FRONT AVAILABLE')) fail('Work Front Available row was not added.');
if (!code.includes('downloadIsoStage(selected, stage, true)')) fail('Completed filtered download was not added.');
if (!code.includes('downloadIsoStage(selected, stage, false)')) fail('Work Front filtered download was not added.');

const temp = target + '.tmp';
fs.writeFileSync(temp, code, 'utf8');
fs.renameSync(temp, target);

console.log('Updated successfully: ' + target);
console.log('Added module-wise automatic stage tick counts for all nine ISO control columns.');
console.log('Added premium project CONSOLE with all-module totals and the same interface.');
console.log('Added clickable Completed and Work Front Available counts with filtered Excel downloads.');
console.log('No unrelated report, UI, Excel metric or workflow was changed.');
console.log('Original bytes: ' + Buffer.byteLength(original) + ' | Updated bytes: ' + Buffer.byteLength(code));
