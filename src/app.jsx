import { useMemo, useRef, useState } from 'react';
import * as XLSX from 'xlsx';

const ACTIONS = ['Reports and Status', 'Model vs. MTO vs. PID Check', 'Equipment Status', 'ISO Planning and Management', 'Support Planning and Management'];
const REPORTS = ['MTO Report', 'Valve Report', 'Pipe Branch Report', 'Primary Support Report', 'ATTA Report', 'Elbow and Bend Report', 'Equipment Orientation and Position Report', 'Special Item Report', 'Nozzle Report'];
const GROUPS = [
  ['E3D Reports', REPORTS],
  ['Modelling Status', ['2 Inch and Below', '3 Inch and 4 Inch', '6 Inch and Above', 'Fluid Code', 'Material and Specification']],
  ['Gate Status', ['Line Status', 'Equipment Status', 'Escape Route Status', 'Safety Equipment Status', 'Material Handling Status', 'Support Status']],
  ['ISO Utility', ['Data Consistency', 'Gusset Report', 'ISO Break Check', 'Specification Mismatch', 'Pipe Aid Check', 'Pipe Insulation Check', 'Vendor Document Check', 'Weld Gap Check', 'Support Tag Check', 'Clash Report']],
];

const normalize = (value) => String(value ?? '').toLowerCase().replace(/&/g, 'and').replace(/[^a-z0-9]/g, '');
const isModuleHeader = (value) => {
  const n = normalize(value);
  return n === 'module' || n === 'modulename' || n === 'modulecode' || n === 'moduleid' || n === 'modules';
};
const isOverallHeader = (value) => {
  const n = normalize(value);
  return n === 'overall' || n.startsWith('overallcount') || n.startsWith('overallvalue') || (n.includes('overall') && (n.includes('count') || n.includes('value')));
};
const toNumber = (value) => {
  if (typeof value === 'number') return Number.isFinite(value) ? value : null;
  if (typeof value !== 'string' || !value.trim()) return null;
  let text = value.trim();
  let negative = false;
  if (/^\(.*\)$/.test(text)) { negative = true; text = text.slice(1, -1); }
  text = text.replace(/,/g, '').replace(/[^0-9.+-]/g, '');
  if (!text || !/[0-9]/.test(text)) return null;
  const number = Number(text);
  if (!Number.isFinite(number)) return null;
  return negative ? -Math.abs(number) : number;
};

function findHeader(rows) {
  const limit = Math.min(rows.length, 40);
  for (let rowIndex = 0; rowIndex < limit; rowIndex += 1) {
    const row = rows[rowIndex] ?? [];
    const moduleColumn = row.findIndex(isModuleHeader);
    if (moduleColumn < 0) continue;
    const overallColumns = [];
    row.forEach((value, index) => { if (isOverallHeader(value)) overallColumns.push(index); });
    return { rowIndex, moduleColumn, overallColumns };
  }
  return null;
}

async function parseWorkbook(file) {
  const buffer = await file.arrayBuffer();
  const workbook = XLSX.read(buffer, { type: 'array', cellDates: true, cellStyles: true });
  const sheets = workbook.SheetNames.map((sheetName) => {
    const sheet = workbook.Sheets[sheetName];
    const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, raw: true, defval: null });
    return { sheetName, sheet, rows, header: findHeader(rows) };
  });
  return { file, workbook, sheets };
}

function reportScore(entry, reportName) {
  const source = normalize(`${entry.file.name} ${entry.file.webkitRelativePath || ''} ${entry.workbook.SheetNames.join(' ')}`);
  const full = normalize(reportName);
  const tokens = reportName.toLowerCase().split(/[^a-z0-9]+/).filter((token) => token.length > 1 && token !== 'report');
  const hits = tokens.filter((token) => source.includes(normalize(token))).length;
  if (source.includes(full)) return 1000 + hits;
  if (hits >= Math.max(1, tokens.length - 1)) return 100 + hits;
  return -1;
}

function findReportWorkbook(workbooks, reportName) {
  return workbooks
    .map((entry) => ({ entry, score: reportScore(entry, reportName) }))
    .filter(({ score }) => score >= 0)
    .sort((a, b) => b.score - a.score || a.entry.file.name.localeCompare(b.entry.file.name))[0]?.entry ?? null;
}

function rowsForModule(sheetEntry, moduleName) {
  if (!sheetEntry.header) return [];
  const moduleKey = normalize(moduleName);
  return sheetEntry.rows.slice(sheetEntry.header.rowIndex + 1).filter((row) => normalize(row?.[sheetEntry.header.moduleColumn]) === moduleKey);
}

function countForModule(workbookEntry, moduleName) {
  if (!workbookEntry) return 0;
  let total = 0;
  for (const sheet of workbookEntry.sheets) {
    if (!sheet.header) continue;
    for (const row of rowsForModule(sheet, moduleName)) {
      for (const columnIndex of sheet.header.overallColumns) {
        const number = toNumber(row?.[columnIndex]);
        if (number !== null) total += number;
      }
    }
  }
  return total;
}

function safeSheetName(value, usedNames) {
  const base = String(value || 'Sheet').replace(/[\\/?*\[\]:]/g, ' ').trim().slice(0, 31) || 'Sheet';
  let name = base;
  let index = 2;
  while (usedNames.has(name.toLowerCase())) {
    const suffix = ` ${index}`;
    name = `${base.slice(0, 31 - suffix.length)}${suffix}`;
    index += 1;
  }
  usedNames.add(name.toLowerCase());
  return name;
}

function createModuleWorkbook(workbooks, moduleName) {
  const output = XLSX.utils.book_new();
  const usedNames = new Set();
  let exportedRows = 0;

  for (const workbookEntry of workbooks) {
    const workbookLabel = workbookEntry.file.name.replace(/\.[^.]+$/, '');
    for (const sheetEntry of workbookEntry.sheets) {
      if (!sheetEntry.header) continue;
      const matchingRows = rowsForModule(sheetEntry, moduleName);
      if (!matchingRows.length) continue;

      const headerRows = sheetEntry.rows.slice(0, sheetEntry.header.rowIndex + 1);
      const outputRows = [...headerRows, ...matchingRows];
      const outputSheet = XLSX.utils.aoa_to_sheet(outputRows, { cellDates: true });
      if (sheetEntry.sheet['!cols']) outputSheet['!cols'] = sheetEntry.sheet['!cols'];
      if (sheetEntry.sheet['!merges']) {
        outputSheet['!merges'] = sheetEntry.sheet['!merges'].filter((merge) => merge.e.r <= sheetEntry.header.rowIndex);
      }

      const sheetName = safeSheetName(`${workbookLabel} ${sheetEntry.sheetName}`, usedNames);
      XLSX.utils.book_append_sheet(output, outputSheet, sheetName);
      exportedRows += matchingRows.length;
    }
  }

  if (!output.SheetNames.length) return false;
  XLSX.writeFile(output, `${moduleName} All Reports.xlsx`, { compression: true });
  return exportedRows > 0;
}

function createFilteredWorkbook(workbookEntry, moduleName, reportName) {
  const output = XLSX.utils.book_new();
  for (const sheetEntry of workbookEntry.sheets) {
    if (!sheetEntry.header) continue;
    const matchingRows = rowsForModule(sheetEntry, moduleName);
    if (!matchingRows.length) continue;
    const headerRows = sheetEntry.rows.slice(0, sheetEntry.header.rowIndex + 1);
    const outputRows = [...headerRows, ...matchingRows];
    const outputSheet = XLSX.utils.aoa_to_sheet(outputRows, { cellDates: true });
    if (sheetEntry.sheet['!cols']) outputSheet['!cols'] = sheetEntry.sheet['!cols'];
    if (sheetEntry.sheet['!merges']) outputSheet['!merges'] = sheetEntry.sheet['!merges'].filter((merge) => merge.e.r <= sheetEntry.header.rowIndex);
    XLSX.utils.book_append_sheet(output, outputSheet, sheetEntry.sheetName.slice(0, 31));
  }
  if (!output.SheetNames.length) return false;
  XLSX.writeFile(output, `${moduleName} ${reportName}.xlsx`, { compression: true });
  return true;
}

export default function App() {
  const [page, setPage] = useState('home');
  const [title, setTitle] = useState('');
  const [from, setFrom] = useState('home');
  const [excelSession, setExcelSession] = useState({ folderName: '', workbooks: [], modules: [] });
  const go = (nextPage) => { setPage(nextPage); window.scrollTo(0, 0); };
  const blank = (nextTitle, previousPage) => { setTitle(nextTitle); setFrom(previousPage); go('blank'); };
  if (page === 'reports') return <Reports back={() => go('home')} open={(nextTitle) => blank(nextTitle, 'reports')} excelSession={excelSession} setExcelSession={setExcelSession} />;
  if (page === 'blank') return <Layout title={title} sub="Content Placeholder" back={() => go(from)}><div className="blank"><b>RGT</b><p>CONTENT PLACEHOLDER</p><h2>{title}</h2><span>No project data or report has been added to this page yet.</span><button onClick={() => go(from)}>Return to Previous Page</button></div></Layout>;
  return <Home reports={() => go('reports')} blank={(nextTitle) => blank(nextTitle, 'home')} />;
}

function Home({ reports, blank }) {
  const action = (code, name, actionName) => code === '4193' && actionName === 'Reports and Status' ? reports() : blank(`${code} · ${name} · ${actionName}`);
  return <><section className="hero"><div/><main><p>FPSO PROJECT PORTAL</p><h1>Reports Generate Tools</h1></main></section><section className="dashboard"><div className="projects"><Card code="4193" name="HAMMER HEAD" click={(item) => action('4193', 'HAMMER HEAD', item)}/><Card code="4173" name="Gato do Mato" click={(item) => action('4173', 'Gato do Mato', item)}/></div></section></>;
}

function Card({ code, name, click }) {
  return <article className="card"><header><b>{code}</b><div><p>PROJECT DASHBOARD</p><h2>{name}</h2></div></header><section>{ACTIONS.map((item) => <button key={item} onClick={() => click(item)}>{item}</button>)}</section></article>;
}

function Reports({ back, open, excelSession, setExcelSession }) {
  const inputRef = useRef(null);
  const [query, setQuery] = useState('');
  const [selected, setSelected] = useState(null);
  const { folderName, workbooks, modules } = excelSession;
  const [busy, setBusy] = useState(false);
  const list = useMemo(() => modules.filter((moduleName) => moduleName.toLowerCase().includes(query.trim().toLowerCase())), [modules, query]);
  const counts = useMemo(() => Object.fromEntries(REPORTS.map((reportName) => [reportName, selected ? countForModule(findReportWorkbook(workbooks, reportName), selected) : 0])), [workbooks, selected]);

  const connectFolder = () => inputRef.current?.click();
  const loadFolder = async (event) => {
    const files = Array.from(event.target.files ?? []).filter((file) => /\.(xlsx|xls|xlsm|xlsb)$/i.test(file.name) && !file.name.startsWith('~$'));
    if (!files.length) { window.alert('No Excel file found in the selected folder.'); return; }
    setBusy(true);
    try {
      const parsed = [];
      for (const file of files) parsed.push(await parseWorkbook(file));
      const moduleMap = new Map();
      for (const workbookEntry of parsed) {
        for (const sheetEntry of workbookEntry.sheets) {
          if (!sheetEntry.header) continue;
          for (const row of sheetEntry.rows.slice(sheetEntry.header.rowIndex + 1)) {
            const raw = String(row?.[sheetEntry.header.moduleColumn] ?? '').trim();
            const key = normalize(raw);
            if (key && !moduleMap.has(key)) moduleMap.set(key, raw);
          }
        }
      }
      const dynamicModules = [...moduleMap.values()].sort((a, b) => a.localeCompare(b, undefined, { numeric: true, sensitivity: 'base' }));
      const firstPath = files[0].webkitRelativePath || files[0].name;
      const connectedFolderName = firstPath.includes('/') ? firstPath.split('/')[0] : 'Selected Excel files';
      setExcelSession({ folderName: connectedFolderName, workbooks: parsed, modules: dynamicModules });
      setSelected((current) => dynamicModules.some((item) => normalize(item) === normalize(current)) ? current : null);
      if (!dynamicModules.length) window.alert('Excel files connected, but no Module/Module Name column was found.');
    } catch (error) {
      console.error(error);
      window.alert('Excel folder could not be read. Please check the workbook format.');
    } finally {
      setBusy(false);
      event.target.value = '';
    }
  };

  const downloadModuleExcel = () => {
    if (!selected) return;
    if (!createModuleWorkbook(workbooks, selected)) {
      window.alert(`No ${selected} rows found in the connected Excel files.`);
    }
  };

  const downloadReport = (reportName) => {
    if (!selected) return;
    const workbookEntry = findReportWorkbook(workbooks, reportName);
    if (!workbookEntry) { window.alert(`No master Excel found for ${reportName}.`); return; }
    if (!createFilteredWorkbook(workbookEntry, selected, reportName)) window.alert(`No ${selected} row found for ${reportName}.`);
  };

  return <Layout title="Reports and Status" sub="Project 4193 · HAMMER HEAD" back={back}><div className="work"><aside><header><p>PROJECT MODULES</p><h2>Module List</h2></header><div className="folder-connect"><button onClick={connectFolder}>Connect Excel Folder</button><span>{folderName ? `Connected: ${folderName}` : 'No folder connected'}</span><input ref={inputRef} className="hidden-folder-input" type="file" accept=".xlsx,.xls,.xlsm,.xlsb" multiple webkitdirectory="" directory="" onChange={loadFolder}/></div><label>⌕<input value={query} onChange={(event) => setQuery(event.target.value)} placeholder="Search module name"/></label><section>{list.map((moduleName) => <button className={moduleName === selected ? 'sel' : ''} key={moduleName} onClick={() => setSelected((current) => current === moduleName ? null : moduleName)}>{moduleName}</button>)}</section></aside><article className="reports">{selected ? <><header><div><p>SELECTED MODULE</p><h2>{selected}</h2></div><button onClick={() => setSelected(null)}>Close</button></header><section><button className="module-excel-download" onClick={downloadModuleExcel}><span>Download {selected} Excel</span><b>↓</b></button>{REPORTS.map((reportName) => <button key={reportName} onClick={() => downloadReport(reportName)}><span>{reportName}</span><b>{busy ? '…' : counts[reportName] ?? 0}</b></button>)}</section></> : <div className="empty"><b>RGT</b><p>MODULE SELECTION</p><h2>{busy ? 'Reading Excel files…' : modules.length ? 'Select a module' : 'Connect the master Excel folder'}</h2><span>{modules.length ? 'Select a module on the left to open reports here.' : 'The module list will be created automatically from the Module column.'}</span></div>}</article></div><Integrated open={open}/></Layout>;
}

function Integrated({ open }) {
  return <section className="integrated"><header><p>INTEGRATED TOOL</p><h2>Overall Project and Module-Wise Status</h2></header><div>{GROUPS.map(([groupName, items], groupIndex) => <article key={groupName}><h3>{groupName}</h3><section>{items.map((item, itemIndex) => <button className={`c${(groupIndex + itemIndex) % 5}`} key={item} onClick={() => open(item)}>{item}</button>)}</section></article>)}</div></section>;
}

function Layout({ title, sub, back, children }) {
  return <main className="page"><header className="top"><button onClick={back}>Back</button><div><p>REPORTS GENERATE TOOLS</p><h1>{title}</h1><span>{sub}</span></div></header><section className="content">{children}</section></main>;
}
