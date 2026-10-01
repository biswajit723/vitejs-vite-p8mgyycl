import { useMemo, useRef, useState } from 'react';
import * as XLSX from 'xlsx';

const ACTIONS = ['Reports and Status', 'Model vs. MTO vs. PID Check', 'Equipment Status', 'ISO Planning and Management', 'Support Planning and Management'];
const REPORTS = ['MTO Report', 'Valve Report', 'Pipe Branch Report', 'Primary Support Report', 'ATTA Report', 'Elbow and Bend Report', 'Equipment Orientation and Position Report', 'Special Item Report', 'Nozzle Report'];
const GROUPS = [
  ['E3D Reports', REPORTS],
  ['Modelling Status', ['Size Status', 'Fluid Code', 'Material and Specification']],
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


const isSizeHeader = (value) => {
  const n = normalize(value);
  return ['size', 'pipesize', 'nominalsize', 'nominaldiameter', 'diameter', 'nps', 'sizegroup', 'sizerange'].includes(n) || n.includes('pipesize');
};
const isModelStatusHeader = (value) => {
  const n = normalize(value);
  return ['status', 'modelstatus', 'modellingstatus', 'modelingstatus', 'modelledstatus', 'modeledstatus'].includes(n) || (n.includes('model') && n.includes('status'));
};
const sizeBucket = (value) => {
  const text = String(value ?? '').trim();
  const n = normalize(text);
  if (!n) return null;
  if (n.includes('lessthan2') || n.includes('below2') || n === '2andbelow') return '<2"';
  if (n.includes('2to4') || n.includes('24')) return '2"-4"';
  if (n.includes('6to8') || n.includes('68')) return '6"-8"';
  if (n.includes('10to14') || n.includes('1014')) return '10"-14"';
  if (n.includes('16to20') || n.includes('1620')) return '16"-20"';
  const firstNumber = Number((text.match(/\d+(?:\.\d+)?/) || [])[0]);
  if (!Number.isFinite(firstNumber)) return null;
  if (firstNumber <= 2) return '<2"';
  if (firstNumber <= 4) return '2"-4"';
  if (firstNumber <= 8) return '6"-8"';
  if (firstNumber <= 14) return '10"-14"';
  if (firstNumber <= 20) return '16"-20"';
  return null;
};
const modelStatus = (value) => {
  const n = normalize(value);
  if (n.includes('notmodelled') || n.includes('notmodeled') || n.includes('unmodelled') || n.includes('unmodeled')) return 'NOT MODELLED';
  if (n.includes('modelled') || n.includes('modeled')) return 'MODELLED';
  return null;
};
const SIZE_COLUMNS = ['<2"', '2"-4"', '6"-8"', '10"-14"', '16"-20"'];
function createEmptySizeSummary() {
  return {
    MODELLED: Object.fromEntries(SIZE_COLUMNS.map((item) => [item, 0])),
    'NOT MODELLED': Object.fromEntries(SIZE_COLUMNS.map((item) => [item, 0])),
  };
}
function getSizeStatusSummary(workbooks, moduleName) {
  const summary = createEmptySizeSummary();
  for (const workbookEntry of workbooks) {
    for (const sheetEntry of workbookEntry.sheets) {
      if (!sheetEntry.header) continue;
      const headerRow = sheetEntry.rows[sheetEntry.header.rowIndex] ?? [];
      const sizeColumn = headerRow.findIndex(isSizeHeader);
      const statusColumn = headerRow.findIndex(isModelStatusHeader);
      if (sizeColumn < 0 || statusColumn < 0) continue;
      const overallColumns = [];
      headerRow.forEach((value, columnIndex) => { if (isOverallHeader(value)) overallColumns.push(columnIndex); });
      for (const row of rowsForModule(sheetEntry, moduleName)) {
        const bucket = sizeBucket(row?.[sizeColumn]);
        const status = modelStatus(row?.[statusColumn]);
        if (!bucket || !status) continue;
        let weight = 1;
        if (overallColumns.length) {
          const values = overallColumns.map((columnIndex) => toNumber(row?.[columnIndex])).filter((value) => value !== null);
          if (values.length) weight = values.reduce((total, value) => total + value, 0);
        }
        summary[status][bucket] += weight;
      }
    }
  }
  return summary;
}
function sizeSummaryTotal(summary) {
  return Object.values(summary).reduce((grandTotal, row) => grandTotal + Object.values(row).reduce((total, value) => total + value, 0), 0);
}
function downloadSizeStatusExcel(moduleName, summary) {
  const modelled = SIZE_COLUMNS.map((column) => summary.MODELLED[column]);
  const notModelled = SIZE_COLUMNS.map((column) => summary['NOT MODELLED'][column]);
  const totals = SIZE_COLUMNS.map((column, index) => modelled[index] + notModelled[index]);
  const rows = [
    [`${moduleName}-Status`, ...SIZE_COLUMNS, 'Grand Total'],
    ['MODELLED', ...modelled, modelled.reduce((total, value) => total + value, 0)],
    ['NOT MODELLED', ...notModelled, notModelled.reduce((total, value) => total + value, 0)],
    [],
    ['GRAND TOTAL', ...totals, totals.reduce((total, value) => total + value, 0)],
  ];
  const workbook = XLSX.utils.book_new();
  const worksheet = XLSX.utils.aoa_to_sheet(rows);
  worksheet['!cols'] = [{ wch: 20 }, ...SIZE_COLUMNS.map(() => ({ wch: 12 })), { wch: 14 }];
  XLSX.utils.book_append_sheet(workbook, worksheet, `${moduleName}-Status`.slice(0, 31));
  XLSX.writeFile(workbook, `${moduleName} Size Status.xlsx`, { compression: true });
}
function SizeStatusView({ moduleName, workbooks, onBack }) {
  const summary = useMemo(() => getSizeStatusSummary(workbooks, moduleName), [workbooks, moduleName]);
  const modelledTotal = SIZE_COLUMNS.reduce((total, column) => total + summary.MODELLED[column], 0);
  const notModelledTotal = SIZE_COLUMNS.reduce((total, column) => total + summary['NOT MODELLED'][column], 0);
  return <section className="size-status-page"><header><div><p>MODELLING STATUS</p><h2>{moduleName} Size Status</h2></div><div className="size-status-actions"><button onClick={() => downloadSizeStatusExcel(moduleName, summary)}>Download Excel</button><button onClick={onBack}>Back</button></div></header><div className="size-status-table-wrap"><table><thead><tr><th>{moduleName}-Status</th>{SIZE_COLUMNS.map((column) => <th key={column}>{column}</th>)}<th>Grand Total</th></tr></thead><tbody><tr><th>MODELLED</th>{SIZE_COLUMNS.map((column) => <td key={column}>{summary.MODELLED[column]}</td>)}<td><strong>{modelledTotal}</strong></td></tr><tr><th>NOT MODELLED</th>{SIZE_COLUMNS.map((column) => <td key={column}>{summary['NOT MODELLED'][column]}</td>)}<td><strong>{notModelledTotal}</strong></td></tr><tr className="grand-total-row"><th>GRAND TOTAL</th>{SIZE_COLUMNS.map((column) => <td key={column}><strong>{summary.MODELLED[column] + summary['NOT MODELLED'][column]}</strong></td>)}<td><strong>{modelledTotal + notModelledTotal}</strong></td></tr></tbody></table></div></section>;
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
  const [showSizeStatus, setShowSizeStatus] = useState(false);
  const list = useMemo(() => modules.filter((moduleName) => moduleName.toLowerCase().includes(query.trim().toLowerCase())), [modules, query]);
  const counts = useMemo(() => Object.fromEntries(GROUPS.flatMap(([, items]) => items).map((item) => [item, selected ? (item === 'Size Status' ? sizeSummaryTotal(getSizeStatusSummary(workbooks, selected)) : countForModule(findReportWorkbook(workbooks, item), selected)) : 0])), [workbooks, selected]);

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

  return <Layout title="Reports and Status" sub="Project 4193 · HAMMER HEAD" back={back}><div className="work"><aside><header><p>PROJECT MODULES</p><h2>Module List</h2></header><div className="folder-connect"><button onClick={connectFolder}>Connect Excel Folder</button><span>{folderName ? `Connected: ${folderName}` : 'No folder connected'}</span><input ref={inputRef} className="hidden-folder-input" type="file" accept=".xlsx,.xls,.xlsm,.xlsb" multiple webkitdirectory="" directory="" onChange={loadFolder}/></div><label>⌕<input value={query} onChange={(event) => setQuery(event.target.value)} placeholder="Search module name"/></label><section>{list.map((moduleName) => <button className={moduleName === selected ? 'sel' : ''} key={moduleName} onClick={() => setSelected((current) => current === moduleName ? null : moduleName)}>{moduleName}</button>)}</section></aside><article className="reports">{selected ? (showSizeStatus ? <SizeStatusView moduleName={selected} workbooks={workbooks} onBack={() => setShowSizeStatus(false)} /> : <><header><div><p>SELECTED MODULE</p><h2>{selected}</h2></div><div className="selected-module-actions"><button className="header-download-button" onClick={downloadModuleExcel}>Download {selected} Excel</button><button onClick={() => setSelected(null)}>Close</button></div></header><section className="selected-module-columns">{GROUPS.map(([groupName, items], groupIndex) => <article className="selected-module-column" key={groupName}><h3>{groupName}</h3><div>{items.map((item, itemIndex) => <button className={`c${(groupIndex + itemIndex) % 5}`} key={item} onClick={() => item === 'Size Status' ? setShowSizeStatus(true) : downloadReport(item)}><span>{item}</span><b>{busy ? '…' : counts[item] ?? 0}</b></button>)}</div></article>)}</section></>) : <div className="empty"><b>RGT</b><p>MODULE SELECTION</p><h2>{busy ? 'Reading Excel files…' : modules.length ? 'Select a module' : 'Connect the master Excel folder'}</h2><span>{modules.length ? 'Select a module on the left to open reports here.' : 'The module list will be created automatically from the Module column.'}</span></div>}</article></div><Integrated open={open}/></Layout>;
}

function Integrated({ open }) {
  return <section className="integrated"><header><p>INTEGRATED TOOL</p><h2>Overall Project and Module-Wise Status</h2></header><div>{GROUPS.map(([groupName, items], groupIndex) => <article key={groupName}><h3>{groupName}</h3><section>{items.map((item, itemIndex) => <button className={`c${(groupIndex + itemIndex) % 5}`} key={item} onClick={() => open(item)}>{item}</button>)}</section></article>)}</div></section>;
}

function Layout({ title, sub, back, children }) {
  return <main className="page"><header className="top"><button onClick={back}>Back</button><div><p>REPORTS GENERATE TOOLS</p><h1>{title}</h1><span>{sub}</span></div></header><section className="content">{children}</section></main>;
}
