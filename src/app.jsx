import { useEffect, useMemo, useRef, useState } from 'react';
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

const SIZE_BUCKETS = ['<2"', '2"-4"', '6"-8"', '10"-14"', '16"-20"'];
const SIZE_HEADER_ALIASES = ['line size', 'nps', 'nps size', 'nominal size', 'nominal diameter', 'pipe size', 'size', 'diameter', 'dia'];
const MODEL_STATUS_HEADER_ALIASES = ['model status', 'modelled status', 'modeled status', 'modelling status', 'modeling status', 'modelled', 'modeled'];

function findColumnByAliases(row, aliases) {
  if (!Array.isArray(row)) return -1;
  const normalizedAliases = aliases.map(normalize);
  return row.findIndex((value) => normalizedAliases.includes(normalize(value)));
}

function sizeBucket(value) {
  const text = String(value ?? '').trim().toLowerCase();
  if (!text) return null;
  const numberMatch = text.match(/\d+(?:\.\d+)?/);
  if (!numberMatch) return null;
  const size = Number(numberMatch[0]);
  if (!Number.isFinite(size)) return null;
  if (size < 2) return '<2"';
  if (size <= 4) return '2"-4"';
  if (size >= 6 && size <= 8) return '6"-8"';
  if (size >= 10 && size <= 14) return '10"-14"';
  if (size >= 16 && size <= 20) return '16"-20"';
  return null;
}

function modellingState(value) {
  const n = normalize(value);
  if (!n) return null;
  if (n.includes('notmodel') || n === 'notmodeled' || n === 'notmodelled') return 'NOT MODELLED';
  if (n.includes('model') || ['yes', 'true', 'done', 'complete', 'completed', '1'].includes(n)) return 'MODELLED';
  if (['no', 'false', '0'].includes(n)) return 'NOT MODELLED';
  return null;
}

function getSizeStatus(workbooks, moduleName) {
  const result = {
    'MODELLED': Object.fromEntries(SIZE_BUCKETS.map((bucket) => [bucket, 0])),
    'NOT MODELLED': Object.fromEntries(SIZE_BUCKETS.map((bucket) => [bucket, 0])),
  };

  if (!moduleName) return result;

  for (const workbookEntry of workbooks) {
    for (const sheetEntry of workbookEntry.sheets) {
      if (!sheetEntry.header) continue;
      const headerRow = sheetEntry.rows[sheetEntry.header.rowIndex] ?? [];
      const sizeColumn = findColumnByAliases(headerRow, SIZE_HEADER_ALIASES);
      const modelColumn = findColumnByAliases(headerRow, MODEL_STATUS_HEADER_ALIASES);
      if (sizeColumn < 0 || modelColumn < 0) continue;

      for (const row of rowsForModule(sheetEntry, moduleName)) {
        const bucket = sizeBucket(row?.[sizeColumn]);
        const state = modellingState(row?.[modelColumn]);
        if (bucket && state) result[state][bucket] += 1;
      }
    }
  }

  return result;
}

function createSizeStatusWorkbook(moduleName, status) {
  const rows = [
    ['1PA-Status', ...SIZE_BUCKETS, 'Grand Total'],
    ['MODELLED', ...SIZE_BUCKETS.map((bucket) => status['MODELLED'][bucket]), Object.values(status['MODELLED']).reduce((sum, value) => sum + value, 0)],
    ['NOT MODELLED', ...SIZE_BUCKETS.map((bucket) => status['NOT MODELLED'][bucket]), Object.values(status['NOT MODELLED']).reduce((sum, value) => sum + value, 0)],
    ['GRAND TOTAL', ...SIZE_BUCKETS.map((bucket) => status['MODELLED'][bucket] + status['NOT MODELLED'][bucket]), SIZE_BUCKETS.reduce((sum, bucket) => sum + status['MODELLED'][bucket] + status['NOT MODELLED'][bucket], 0)],
  ];
  const output = XLSX.utils.book_new();
  const sheet = XLSX.utils.aoa_to_sheet(rows);
  XLSX.utils.book_append_sheet(output, sheet, 'Size Status');
  XLSX.writeFile(output, `${moduleName} Size Status.xlsx`, { compression: true });
}

const MODULE_CACHE_KEY = 'rgt:module-list:v1';
const ADMIN_TOKEN_KEY = 'rgt:admin-sync-token:v1';

function readCachedModules() {
  try {
    const raw = localStorage.getItem(MODULE_CACHE_KEY);
    const parsed = JSON.parse(raw || '[]');
    return Array.isArray(parsed) ? parsed.filter((value) => typeof value === 'string' && value.trim()) : [];
  } catch {
    return [];
  }
}

function cacheModules(modules) {
  try {
    localStorage.setItem(MODULE_CACHE_KEY, JSON.stringify(modules));
  } catch {
    // Local cache is only a fallback; central storage remains the source of truth.
  }
}

async function fetchCentralModules() {
  const response = await fetch('/api/modules', { headers: { Accept: 'application/json' } });
  if (!response.ok) throw new Error(`Central module service returned ${response.status}`);
  const data = await response.json();
  const modules = Array.isArray(data?.modules) ? data.modules.filter((value) => typeof value === 'string' && value.trim()) : [];
  return { modules, source: data?.source || 'central' };
}

async function publishCentralModules(modules, token) {
  const response = await fetch('/api/modules', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json', Accept: 'application/json', 'x-admin-token': token },
    body: JSON.stringify({ modules }),
  });
  let data = null;
  try { data = await response.json(); } catch { /* ignore malformed error bodies */ }
  if (!response.ok) {
    throw new Error(data?.error || `Module publish failed (${response.status})`);
  }
  return data;
}

export default function App() {
  const [page, setPage] = useState('home');
  const [title, setTitle] = useState('');
  const [from, setFrom] = useState('home');
  const [excelSession, setExcelSession] = useState(() => ({ folderName: '', workbooks: [], modules: readCachedModules() }));
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
  const [statusView, setStatusView] = useState(null);
  const [centralStatus, setCentralStatus] = useState('loading');
  const { folderName, workbooks, modules } = excelSession;
  const [busy, setBusy] = useState(false);

  useEffect(() => {
    let cancelled = false;
    (async () => {
      try {
        const result = await fetchCentralModules();
        if (cancelled) return;
        if (result.modules.length) {
          cacheModules(result.modules);
          setExcelSession((current) => ({ ...current, modules: result.modules }));
          setCentralStatus('connected');
        } else {
          setCentralStatus('empty');
        }
      } catch (error) {
        console.warn('Central module list unavailable; using cached modules.', error);
        if (!cancelled) setCentralStatus('offline');
      }
    })();
    return () => { cancelled = true; };
  }, [setExcelSession]);
  const list = useMemo(() => modules.filter((moduleName) => moduleName.toLowerCase().includes(query.trim().toLowerCase())), [modules, query]);
  const COUNT_ITEMS = useMemo(() => GROUPS.flatMap(([, items]) => items).filter((item) => item !== 'Size Status'), []);
  const counts = useMemo(() => Object.fromEntries(COUNT_ITEMS.map((item) => [item, selected ? countForModule(findReportWorkbook(workbooks, item), selected) : 0])), [COUNT_ITEMS, workbooks, selected]);

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
      cacheModules(dynamicModules);
      setSelected((current) => dynamicModules.some((item) => normalize(item) === normalize(current)) ? current : null);
      if (!dynamicModules.length) {
        window.alert('Excel files connected, but no Module/Module Name column was found.');
        return;
      }

      let adminToken = '';
      try { adminToken = sessionStorage.getItem(ADMIN_TOKEN_KEY) || ''; } catch { /* ignore */ }
      if (!adminToken) {
        adminToken = window.prompt('ADMIN SYNC TOKEN\nEnter the admin token to publish this module list for all website users.\nCancel keeps the list local only.') || '';
        if (adminToken) {
          try { sessionStorage.setItem(ADMIN_TOKEN_KEY, adminToken); } catch { /* ignore */ }
        }
      }
      if (adminToken) {
        try {
          await publishCentralModules(dynamicModules, adminToken);
          setCentralStatus('connected');
          window.alert('Master module list published successfully. All users will see this module list.');
        } catch (error) {
          console.error(error);
          try { sessionStorage.removeItem(ADMIN_TOKEN_KEY); } catch { /* ignore */ }
          setCentralStatus('offline');
          window.alert(`Module list was loaded locally, but central publish failed.\n${error.message}`);
        }
      }
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

  const sizeStatus = useMemo(() => getSizeStatus(workbooks, selected), [workbooks, selected]);

  if (statusView === 'Size Status' && selected) {
    return <SizeStatusPage moduleName={selected} status={sizeStatus} back={() => setStatusView(null)} />;
  }

  const openIntegrated = (item) => {
    if (item === 'Size Status' && selected) {
      setStatusView('Size Status');
      return;
    }
    open(item);
  };

  const moduleSourceText = modules.length
    ? (centralStatus === 'connected' ? 'Connected: MASTER MODULE LIST' : folderName ? `Connected: ${folderName}` : 'Module list loaded')
    : 'No master module list connected';

  return <Layout title="Reports and Status" sub="Project 4193 · HAMMER HEAD" back={back}><div className="work"><aside><header><p>PROJECT MODULES</p><h2>Module List</h2></header><div className="folder-connect"><button onClick={connectFolder}>Connect Excel Folder</button><span>{moduleSourceText}</span><input ref={inputRef} className="hidden-folder-input" type="file" accept=".xlsx,.xls,.xlsm,.xlsb" multiple webkitdirectory="" directory="" onChange={loadFolder}/></div><label>⌕<input value={query} onChange={(event) => setQuery(event.target.value)} placeholder="Search module name"/></label><section>{list.map((moduleName) => <button className={moduleName === selected ? 'sel' : ''} key={moduleName} onClick={() => setSelected((current) => current === moduleName ? null : moduleName)}>{moduleName}</button>)}</section></aside><article className="reports">{selected ? <><header><div><p>SELECTED MODULE</p><h2>{selected}</h2></div><div className="selected-module-actions"><button className="header-download-button" onClick={downloadModuleExcel}>Download {selected} Excel</button><button onClick={() => setSelected(null)}>Close</button></div></header><section className="selected-module-columns">{GROUPS.map(([groupName, items], groupIndex) => <div className="selected-module-column" key={groupName}><h3>{groupName}</h3><div>{items.map((item, itemIndex) => { const isReport = REPORTS.includes(item); const isSizeStatus = item === 'Size Status'; const count = isSizeStatus ? null : (busy ? '…' : counts[item] ?? 0); return <button className={`c${(groupIndex + itemIndex) % 5}`} key={item} onClick={() => isReport ? downloadReport(item) : openIntegrated(item)}><span>{item}</span>{count !== null ? <b>{count}</b> : null}</button>; })}</div></div>)}</section></> : <div className="empty"><b>RGT</b><p>MODULE SELECTION</p><h2>{busy ? 'Reading Excel files…' : modules.length ? 'Select a module' : 'Connect the master Excel folder'}</h2><span>{modules.length ? 'Select a module on the left to open reports here.' : 'The module list will be created automatically from the Module column.'}</span></div>}</article></div><Integrated open={openIntegrated}/></Layout>;
}

function SizeStatusPage({ moduleName, status, back }) {
  const modelledTotal = SIZE_BUCKETS.reduce((sum, bucket) => sum + status['MODELLED'][bucket], 0);
  const notModelledTotal = SIZE_BUCKETS.reduce((sum, bucket) => sum + status['NOT MODELLED'][bucket], 0);
  const grandTotal = modelledTotal + notModelledTotal;
  return <main className="page size-status-page">
    <header className="top size-status-top">
      <div className="size-status-heading"><p>MODELLING STATUS</p><h1>{moduleName} Size Status</h1></div>
      <div className="size-status-actions"><button onClick={() => createSizeStatusWorkbook(moduleName, status)}>Download Excel</button><button onClick={back}>Back</button></div>
    </header>
    <section className="content size-status-content">
      <div className="size-status-table-wrap">
        <table className="size-status-table">
          <thead><tr><th>{moduleName}-Status</th>{SIZE_BUCKETS.map((bucket) => <th key={bucket}>{bucket}</th>)}<th>Grand Total</th></tr></thead>
          <tbody>
            <tr><th>MODELLED</th>{SIZE_BUCKETS.map((bucket) => <td key={bucket}>{status['MODELLED'][bucket]}</td>)}<td>{modelledTotal}</td></tr>
            <tr><th>NOT MODELLED</th>{SIZE_BUCKETS.map((bucket) => <td key={bucket}>{status['NOT MODELLED'][bucket]}</td>)}<td>{notModelledTotal}</td></tr>
            <tr className="size-status-grand"><th>GRAND TOTAL</th>{SIZE_BUCKETS.map((bucket) => <td key={bucket}>{status['MODELLED'][bucket] + status['NOT MODELLED'][bucket]}</td>)}<td>{grandTotal}</td></tr>
          </tbody>
        </table>
      </div>
    </section>
  </main>;
}

function Integrated({ open }) {
  return <section className="integrated"><header><p>INTEGRATED TOOL</p><h2>Overall Project and Module-Wise Status</h2></header><div>{GROUPS.map(([groupName, items], groupIndex) => <article key={groupName}><h3>{groupName}</h3><section>{items.map((item, itemIndex) => <button className={`c${(groupIndex + itemIndex) % 5}`} key={item} onClick={() => open(item)}>{item}</button>)}</section></article>)}</div></section>;
}

function Layout({ title, sub, back, children }) {
  return <main className="page"><header className="top"><button onClick={back}>Back</button><div><p>REPORTS GENERATE TOOLS</p><h1>{title}</h1><span>{sub}</span></div></header><section className="content">{children}</section></main>;
}
