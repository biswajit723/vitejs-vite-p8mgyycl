const fs = require('fs');
const path = require('path');

const target = process.argv[2] || path.join(process.cwd(), 'src', 'app.jsx');
let code = fs.readFileSync(target, 'utf8');
const original = code;

function replaceOnce(search, replacement, label) {
  if (!code.includes(search)) throw new Error(`Update stopped: ${label} block was not found. No file was changed.`);
  code = code.replace(search, replacement);
}

replaceOnce(
`async function parseWorkbook(file) {
  const buffer = await file.arrayBuffer();
  const workbook = XLSX.read(buffer, { type: 'array', cellDates: true, cellStyles: true });
  const sheets = workbook.SheetNames.map((sheetName) => {
    const sheet = workbook.Sheets[sheetName];
    const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, raw: true, defval: null });
    return { sheetName, sheet, rows, header: findHeader(rows) };
  });
  return { file, workbook, sheets };
}`,
`async function parseWorkbook(file) {
  const buffer = await file.arrayBuffer();
  const workbook = XLSX.read(buffer, { type: 'array', cellDates: true, cellStyles: true, dense: true });
  const sheets = workbook.SheetNames.map((sheetName) => {
    const sheet = workbook.Sheets[sheetName];
    const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, raw: true, defval: null, blankrows: false });
    const header = findHeader(rows);
    const moduleRows = new Map();
    if (header) {
      for (let rowIndex = header.rowIndex + 1; rowIndex < rows.length; rowIndex += 1) {
        const row = rows[rowIndex];
        const key = normalize(row?.[header.moduleColumn]);
        if (!key) continue;
        if (!moduleRows.has(key)) moduleRows.set(key, []);
        moduleRows.get(key).push(row);
      }
    }
    return { sheetName, sheet, rows, header, moduleRows };
  });
  return { file, workbook, sheets };
}

async function parseWorkbooksFast(files, onProgress) {
  const results = new Array(files.length);
  let nextIndex = 0;
  let completed = 0;
  const workerCount = Math.min(files.length, Math.max(2, Math.min(4, navigator.hardwareConcurrency || 4)));
  const worker = async () => {
    while (nextIndex < files.length) {
      const index = nextIndex;
      nextIndex += 1;
      results[index] = await parseWorkbook(files[index]);
      completed += 1;
      onProgress?.(completed, files.length, files[index].name);
      await new Promise((resolve) => window.setTimeout(resolve, 0));
    }
  };
  await Promise.all(Array.from({ length: workerCount }, worker));
  return results;
}`,
'parseWorkbook');

replaceOnce(
`function rowsForModule(sheetEntry, moduleName) {
  if (!sheetEntry.header) return [];
  const moduleKey = normalize(moduleName);
  return sheetEntry.rows.slice(sheetEntry.header.rowIndex + 1).filter((row) => normalize(row?.[sheetEntry.header.moduleColumn]) === moduleKey);
}`,
`function rowsForModule(sheetEntry, moduleName) {
  if (!sheetEntry.header) return [];
  const moduleKey = normalize(moduleName);
  if (sheetEntry.moduleRows) return sheetEntry.moduleRows.get(moduleKey) || [];
  return sheetEntry.rows.slice(sheetEntry.header.rowIndex + 1).filter((row) => normalize(row?.[sheetEntry.header.moduleColumn]) === moduleKey);
}`,
'rowsForModule');

replaceOnce(
`  const [busy, setBusy] = useState(false);`,
`  const [busy, setBusy] = useState(false);
  const [notice, setNotice] = useState(null);
  const [tokenModal, setTokenModal] = useState(null);
  const noticeTimerRef = useRef(null);
  const showNotice = (type, title, message, persistent = false) => {
    window.clearTimeout(noticeTimerRef.current);
    setNotice({ type, title, message, persistent });
    if (!persistent) noticeTimerRef.current = window.setTimeout(() => setNotice(null), 4200);
  };
  const requestAdminToken = () => new Promise((resolve) => setTokenModal({ value: '', resolve }));
  const closeTokenModal = (value = '') => {
    setTokenModal((current) => {
      current?.resolve(value.trim());
      return null;
    });
  };
  useEffect(() => () => window.clearTimeout(noticeTimerRef.current), []);`,
'Reports state');

replaceOnce(
`    if (!files.length) { window.alert('No Excel file found in the selected folder.'); return; }
    setBusy(true);
    try {
      const parsed = [];
      for (const file of files) parsed.push(await parseWorkbook(file));`,
`    if (!files.length) { showNotice('warning', 'No Excel file found', 'Select a folder containing .xlsx, .xls, .xlsm or .xlsb files.'); return; }
    setBusy(true);
    showNotice('loading', 'Connecting Excel folder', 'Preparing ' + files.length + ' Excel file' + (files.length === 1 ? '' : 's') + '...', true);
    try {
      const parsed = await parseWorkbooksFast(files, (completed, total, fileName) => {
        showNotice('loading', 'Reading Excel files', completed + ' of ' + total + ' processed · ' + fileName, true);
      });`,
'load start');

replaceOnce(
`        window.alert('Excel files connected, but no Module/Module Name column was found.');
        return;`,
`        showNotice('warning', 'Folder connected', 'No Module or Module Name column was found in the connected Excel files.');
        return;`,
'no module notice');

replaceOnce(
`        adminToken = window.prompt('ADMIN SYNC TOKEN\\nEnter the admin token to publish this module list for all website users.\\nCancel keeps the list local only.') || '';`,
`        adminToken = await requestAdminToken();`,
'admin prompt');

replaceOnce(
`            window.alert('Master module list published successfully. All users will see this module list.');`,
`            showNotice('success', 'Excel folder connected', files.length + ' files processed. Master module list published for all users.');`,
'publish success');

replaceOnce(
`          window.alert(\`Module list was loaded locally, but central publish failed.\\n\${error.message}\`);`,
`          showNotice('warning', 'Connected locally', 'Excel data is ready, but central publish failed: ' + error.message);`,
'publish failure');

replaceOnce(
`      }
    } catch (error) {
      console.error(error);
      window.alert('Excel folder could not be read. Please check the workbook format.');`,
`      } else {
        showNotice('success', 'Excel folder connected', files.length + ' files processed successfully. ' + dynamicModules.length + ' modules are ready.');
      }
    } catch (error) {
      console.error(error);
      showNotice('error', 'Excel connection failed', 'The folder could not be read. Check the workbook format and try again.');`,
'load finish');

const returnNeedle = `  return <Layout title="Reports and Status" sub="Project 4193 · HAMMER HEAD" back={back}>`;
replaceOnce(returnNeedle,
`  return <>
    <style>{\`
      .rgt-notice-stack{position:fixed;top:24px;right:24px;z-index:10000;width:min(390px,calc(100vw - 32px));pointer-events:none}.rgt-notice{pointer-events:auto;display:grid;grid-template-columns:42px 1fr 30px;gap:12px;align-items:center;padding:14px 14px 14px 16px;border:1px solid rgba(148,163,184,.28);border-radius:16px;background:rgba(15,23,42,.96);box-shadow:0 24px 70px rgba(15,23,42,.32);color:#fff;backdrop-filter:blur(16px);animation:rgtNoticeIn .28s cubic-bezier(.2,.8,.2,1)}.rgt-notice-icon{width:40px;height:40px;display:grid;place-items:center;border-radius:12px;background:#0ea5e9;font-weight:900}.rgt-notice.success .rgt-notice-icon{background:#10b981}.rgt-notice.warning .rgt-notice-icon{background:#f59e0b}.rgt-notice.error .rgt-notice-icon{background:#ef4444}.rgt-notice.loading .rgt-notice-icon:after{content:'';width:18px;height:18px;border:3px solid rgba(255,255,255,.42);border-top-color:#fff;border-radius:50%;animation:rgtSpin .8s linear infinite}.rgt-notice strong{display:block;font-size:14px;line-height:1.25}.rgt-notice p{margin:4px 0 0;color:#cbd5e1;font-size:12px;line-height:1.45;word-break:break-word}.rgt-notice-close{border:0;background:transparent;color:#94a3b8;font-size:22px;cursor:pointer}.rgt-token-backdrop{position:fixed;inset:0;z-index:9999;display:grid;place-items:center;padding:20px;background:rgba(15,23,42,.55);backdrop-filter:blur(6px);animation:rgtFade .2s ease}.rgt-token-modal{width:min(430px,100%);padding:24px;border-radius:20px;background:#fff;box-shadow:0 30px 90px rgba(15,23,42,.38)}.rgt-token-modal p{margin:0;color:#0284c7;font-size:12px;font-weight:800;letter-spacing:.12em}.rgt-token-modal h3{margin:7px 0 8px;color:#0f172a}.rgt-token-modal span{display:block;color:#64748b;font-size:13px;line-height:1.5}.rgt-token-modal input{box-sizing:border-box;width:100%;margin:18px 0;padding:12px 14px;border:1px solid #cbd5e1;border-radius:11px;outline:none}.rgt-token-modal input:focus{border-color:#0ea5e9;box-shadow:0 0 0 3px rgba(14,165,233,.14)}.rgt-token-actions{display:flex;justify-content:flex-end;gap:10px}.rgt-token-actions button{padding:10px 15px;border:0;border-radius:10px;cursor:pointer;font-weight:700}.rgt-token-actions button:first-child{background:#e2e8f0;color:#334155}.rgt-token-actions button:last-child{background:#0284c7;color:#fff}@keyframes rgtSpin{to{transform:rotate(360deg)}}@keyframes rgtNoticeIn{from{opacity:0;transform:translate3d(24px,-8px,0)}to{opacity:1;transform:none}}@keyframes rgtFade{from{opacity:0}to{opacity:1}}@media(max-width:640px){.rgt-notice-stack{top:14px;right:16px}}
    \`}</style>
    {notice ? <div className="rgt-notice-stack" role="status" aria-live="polite"><div className={'rgt-notice ' + notice.type}><div className="rgt-notice-icon">{notice.type === 'success' ? '✓' : notice.type === 'warning' ? '!' : notice.type === 'error' ? '×' : ''}</div><div><strong>{notice.title}</strong><p>{notice.message}</p></div>{notice.persistent ? null : <button className="rgt-notice-close" onClick={() => setNotice(null)} aria-label="Close notification">×</button>}</div></div> : null}
    {tokenModal ? <div className="rgt-token-backdrop" role="dialog" aria-modal="true" aria-labelledby="rgt-token-title"><form className="rgt-token-modal" onSubmit={(event) => { event.preventDefault(); closeTokenModal(tokenModal.value); }}><p>SECURE ADMIN SYNC</p><h3 id="rgt-token-title">Publish module list</h3><span>Enter the admin token to publish this module list for all website users. You can continue locally without publishing.</span><input autoFocus type="password" value={tokenModal.value} onChange={(event) => setTokenModal((current) => ({ ...current, value: event.target.value }))} placeholder="Admin sync token" autoComplete="off"/><div className="rgt-token-actions"><button type="button" onClick={() => closeTokenModal('')}>Keep Local</button><button type="submit" disabled={!tokenModal.value.trim()}>Publish</button></div></form></div> : null}
    <Layout title="Reports and Status" sub="Project 4193 · HAMMER HEAD" back={back}>`,
'Reports return');

replaceOnce(`</Layout>;
}`, `</Layout>
  </>;
}`, 'Reports closing fragment');

const temp = `${target}.tmp`;
fs.writeFileSync(temp, code);
fs.renameSync(temp, target);
console.log(`Updated successfully: ${target}`);
console.log(`Original bytes: ${Buffer.byteLength(original)} | Updated bytes: ${Buffer.byteLength(code)}`);
