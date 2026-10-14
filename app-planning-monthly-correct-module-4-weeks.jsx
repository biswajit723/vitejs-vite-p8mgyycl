import { useEffect, useMemo, useRef, useState } from 'react';
import * as XLSX from 'xlsx-js-style';

const ACTIONS = ['Reports and Status', 'Model vs. MTO vs. PID Check', 'Equipment Status', 'ISO Planning and Management', 'Support Planning and Management', 'SUPPORT TOOLS'];
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
const isRemarksHeader = (value) => normalize(value).startsWith('remarks');
const hasMeaningfulRemark = (value) => {
  const text = String(value ?? '').trim();
  if (!text) return false;
  return !['unset', 'null', 'nil', 'na', 'none', 'noremark', 'noremarks'].includes(normalize(text));
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
    const remarksColumns = [];
    row.forEach((value, index) => {
      if (isOverallHeader(value)) overallColumns.push(index);
      if (isRemarksHeader(value)) remarksColumns.push(index);
    });
    return { rowIndex, moduleColumn, overallColumns, remarksColumns };
  }
  return null;
}

async function parseWorkbook(file) {
  const buffer = await file.arrayBuffer();
  const workbook = XLSX.read(buffer, { type: 'array', cellDates: true, cellStyles: true, dense: true });
  const sheets = workbook.SheetNames.map((sheetName) => {
    const sheet = workbook.Sheets[sheetName];
    const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, raw: true, defval: null, blankrows: false });
    // Real worksheet row number of every parsed row (blank rows are skipped above, so index != sheet row).
    const rowsWithBlanks = XLSX.utils.sheet_to_json(sheet, { header: 1, raw: true, defval: null, blankrows: true });
    const sheetStartRow = XLSX.utils.decode_range(sheet?.['!ref'] || 'A1').s.r;
    const sourceRowIndexes = [];
    rowsWithBlanks.forEach((row, fullIndex) => {
      if ((row ?? []).some((value) => value !== null && value !== undefined)) sourceRowIndexes.push(sheetStartRow + fullIndex);
    });
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
    return { sheetName, sheet, rows, header, moduleRows, sourceRowIndexes };
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
}

function isIsoPlanningWorkbook(entry) {
  const fileNameWithoutExtension = entry.file.name.replace(/\.[^.]+$/, '');
  const normalizedName = normalize(fileNameWithoutExtension);
  const expectedName = normalize('4193_HH_Isometrics Control Sheet-Piping_OFSI');
  // Allow a trailing numeric suffix such as "...OFSI 2" or "...OFSI (2)".
  return normalizedName === expectedName || (
    normalizedName.startsWith(expectedName) && /^\d+$/.test(normalizedName.slice(expectedName.length))
  );
}

function getIsoPlanningModules(workbooks) {
  const modules = new Map();
  for (const workbookEntry of workbooks) {
    if (!isIsoPlanningWorkbook(workbookEntry)) continue;
    for (const sheetEntry of workbookEntry.sheets) {
      let headerRowIndex = -1;
      let moduleColumn = -1;
      for (let rowIndex = 0; rowIndex < Math.min(sheetEntry.rows.length, 40); rowIndex += 1) {
        const row = sheetEntry.rows[rowIndex] ?? [];
        const foundModuleColumn = row.findIndex(isModuleHeader);
        if (foundModuleColumn >= 0) {
          headerRowIndex = rowIndex;
          moduleColumn = foundModuleColumn;
          break;
        }
      }
      if (headerRowIndex < 0) continue;
      for (const row of sheetEntry.rows.slice(headerRowIndex + 1)) {
        const rawModule = String(row?.[moduleColumn] ?? '').trim();
        const key = normalize(rawModule);
        if (key && !modules.has(key)) modules.set(key, rawModule);
      }
    }
  }
  return [...modules.values()];
}

const ISO_FILE_PREFIX = '4193_HH_Isometrics Control Sheet-Piping_OFSI';
const ISO_PIPE_ALIASES = ['pipe name', 'pipename'];
const ISO_REVISION_ALIASES = ['latest revision', 'latestrevision'];
const ISO_DOCUMENTUM_ALIASES = ['documentum cp status', 'documentumcpstatus'];
const ISO_HOLD_ALIASES = ['hold', 'hold status', 'iso hold', 'on hold'];
const ISO_FORECAST_ALIASES = ['forecast issue date', 'forecastissuedate'];
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
];

function isoModuleFromPipeName(value) {
  const text = String(value ?? '').trim().replace(/^\/+/, '');
  if (!text) return '';
  return text.split('-').map((part) => part.trim()).filter(Boolean)[0] || '';
}
function isExcelDateValue(value) {
  if (value instanceof Date) return !Number.isNaN(value.getTime());
  if (typeof value === 'number') return value > 1 && value < 100000;
  const text = String(value ?? '').trim();
  if (!text || normalize(text).includes('linedeleted')) return false;
  if (!/[\/\-.]/.test(text) || !/\d/.test(text)) return false;
  return !Number.isNaN(Date.parse(text));
}
function isNonZeroRevision(value) {
  if (value === null || value === undefined || String(value).trim() === '') return false;
  const number = Number(String(value).trim().replace(/[^0-9.-]/g, ''));
  return Number.isFinite(number) && number !== 0;
}
function isHoldValue(value) {
  return normalize(value) === 'yes';
}
const isIsoForecastHeader = (value) => {
  const n = normalize(value);
  return ISO_FORECAST_ALIASES.map(normalize).includes(n) || (n.includes('forecast') && (n.includes('issue') || n.includes('date')));
};
function findIsoForecastColumn(rows, headerRowIndex) {
  const exact = findColumnByAliases(rows[headerRowIndex] ?? [], ISO_FORECAST_ALIASES);
  if (exact >= 0) return exact;
  // Header text may differ slightly (extra words / line breaks) or sit a row above/below "PIPE NAME".
  for (const offset of [0, -1, 1, -2, 2]) {
    const row = rows[headerRowIndex + offset];
    if (!Array.isArray(row)) continue;
    const found = row.findIndex(isIsoForecastHeader);
    if (found >= 0) return found;
  }
  return -1;
}
function buildIsoPlanningData(workbooks) {
  const moduleRows = new Map();
  for (const workbookEntry of workbooks) {
    if (!isIsoPlanningWorkbook(workbookEntry)) continue;
    for (const sheetEntry of workbookEntry.sheets) {
      let headerRowIndex = -1;
      let pipeColumn = -1;
      let moduleColumn = -1;
      let issuedColumn = -1;
      let revisionColumn = -1;
      let holdColumn = -1;
      let forecastColumn = -1;
      let stageColumns = {};
      for (let rowIndex = 0; rowIndex < Math.min(sheetEntry.rows.length, 60); rowIndex += 1) {
        const row = sheetEntry.rows[rowIndex] ?? [];
        const foundPipe = findColumnByAliases(row, ISO_PIPE_ALIASES);
        if (foundPipe < 0) continue;
        headerRowIndex = rowIndex;
        pipeColumn = foundPipe;
        moduleColumn = row.findIndex(isModuleHeader);
        issuedColumn = findColumnByAliases(row, ISO_DOCUMENTUM_ALIASES);
        revisionColumn = findColumnByAliases(row, ISO_REVISION_ALIASES);
        holdColumn = findColumnByAliases(row, ISO_HOLD_ALIASES);
        forecastColumn = findIsoForecastColumn(sheetEntry.rows, rowIndex);
        stageColumns = Object.fromEntries(ISO_STAGE_DEFINITIONS.map((stage) => [stage.key, findColumnByAliases(row, stage.aliases)]));
        break;
      }
      if (headerRowIndex < 0) continue;
      for (let rowIndex = headerRowIndex + 1; rowIndex < sheetEntry.rows.length; rowIndex += 1) {
        const row = sheetEntry.rows[rowIndex] ?? [];
        const pipeName = String(row?.[pipeColumn] ?? '').trim();
        if (!pipeName || normalize(pipeName).includes('grandtotal')) continue;
        const moduleName = String(moduleColumn >= 0 ? row?.[moduleColumn] ?? '' : '').trim() || isoModuleFromPipeName(pipeName);
        if (!moduleName) continue;
        const key = normalize(moduleName);
        if (!moduleRows.has(key)) moduleRows.set(key, { moduleName, entries: [] });
        moduleRows.get(key).entries.push({ row, rowIndex, sheetEntry, pipeColumn, issuedColumn, revisionColumn, holdColumn, stageColumns, forecastColumn });
      }
    }
  }
  return [...moduleRows.values()].sort((a, b) => a.moduleName.localeCompare(b.moduleName, undefined, { numeric: true, sensitivity: 'base' }));
}
function parseIsoForecastDate(value) {
  if (value instanceof Date && !Number.isNaN(value.getTime())) return new Date(value.getFullYear(), value.getMonth(), value.getDate());
  if (typeof value === 'number' && value > 1 && value < 100000) {
    const parsed = XLSX.SSF.parse_date_code(value);
    return parsed ? new Date(parsed.y, parsed.m - 1, parsed.d) : null;
  }
  const text = String(value ?? '').trim();
  if (!text || normalize(text).includes('linedeleted')) return null;

  const monthNames = {
    jan: 0, january: 0, feb: 1, february: 1, mar: 2, march: 2,
    apr: 3, april: 3, may: 4, jun: 5, june: 5, jul: 6, july: 6,
    aug: 7, august: 7, sep: 8, sept: 8, september: 8,
    oct: 9, october: 9, nov: 10, november: 10, dec: 11, december: 11,
  };
  const namedMonthMatch = text.match(/^(\d{1,2})[\s\/.-]+([A-Za-z]{3,9})[\s\/.-]+(\d{2,4})$/);
  if (namedMonthMatch) {
    const month = monthNames[namedMonthMatch[2].toLowerCase()];
    const yearNumber = Number(namedMonthMatch[3]);
    const year = yearNumber < 100 ? 2000 + yearNumber : yearNumber;
    const day = Number(namedMonthMatch[1]);
    if (month !== undefined) {
      const parsed = new Date(year, month, day);
      if (parsed.getFullYear() === year && parsed.getMonth() === month && parsed.getDate() === day) return parsed;
    }
  }

  const numericMatch = text.match(/^(\d{1,2})[\/.-](\d{1,2})[\/.-](\d{2,4})$/);
  if (numericMatch) {
    const first = Number(numericMatch[1]);
    const second = Number(numericMatch[2]);
    const yearNumber = Number(numericMatch[3]);
    const year = yearNumber < 100 ? 2000 + yearNumber : yearNumber;
    const month = first > 12 ? second - 1 : first - 1;
    const day = first > 12 ? first : second;
    const parsed = new Date(year, month, day);
    if (parsed.getFullYear() === year && parsed.getMonth() === month && parsed.getDate() === day) return parsed;
  }

  const direct = new Date(text);
  return Number.isNaN(direct.getTime()) ? null : new Date(direct.getFullYear(), direct.getMonth(), direct.getDate());
}

function isoMonday(date) {
  const value = new Date(date.getFullYear(), date.getMonth(), date.getDate());
  const day = value.getDay();
  value.setDate(value.getDate() - (day === 0 ? 6 : day - 1));
  return value;
}
function isoDateKey(date) {
  return [date.getFullYear(), String(date.getMonth() + 1).padStart(2, '0'), String(date.getDate()).padStart(2, '0')].join('-');
}
function isoShortDate(date) {
  return date.toLocaleDateString(undefined, { day: '2-digit', month: 'short' });
}
function isIsoPlanningValidEntry(entry) {
  const pipeName = String(entry.row?.[entry.pipeColumn] ?? '').trim();
  if (!pipeName) return false;

  // Excel may store the delete marker in FORECAST ISSUE DATE or another cell in the row.
  // Exclude all common word orders, not only the exact "Line Deleted" spelling.
  const deleteMarkers = ['linedeleted', 'linedelete', 'deletedline', 'deleteline'];
  const rowHasDeleteMarker = (entry.row || []).some((value) => {
    const cellText = normalize(value);
    return deleteMarkers.some((marker) => cellText.includes(marker));
  });
  return !rowHasDeleteMarker;
}
function buildIsoFourWeekPlan(moduleData) {
  const validEntries = (moduleData?.entries || [])
    .filter(isIsoPlanningValidEntry)
    .map((entry) => ({
      entry,
      date: entry.forecastColumn >= 0 ? parseIsoForecastDate(entry.row?.[entry.forecastColumn]) : null,
    }))
    .filter((item) => item.date)
    .sort((a, b) => a.date - b.date);

  const availableDateMap = new Map();
  for (const item of validEntries) availableDateMap.set(isoDateKey(item.date), item.date);
  const availableDates = [...availableDateMap.values()].sort((a, b) => a - b);
  const selectedDates = availableDates.slice(-4);
  while (selectedDates.length < 4) selectedDates.unshift(null);

  const weeks = selectedDates.map((selectedDate, weekIndex) => {
    const dateKey = selectedDate ? isoDateKey(selectedDate) : '';
    const rows = selectedDate ? validEntries.filter((item) => isoDateKey(item.date) === dateKey) : [];
    return {
      weekIndex,
      label: 'W' + (weekIndex + 1),
      range: selectedDate ? isoShortDate(selectedDate) : 'No Date',
      rows,
      total: rows.length,
      holds: rows.filter((item) => item.entry.holdColumn >= 0 && normalize(item.entry.row?.[item.entry.holdColumn]) === 'yes').length,
      days: [],
    };
  });

  return {
    weeks,
    maxValue: Math.max(1, ...weeks.map((week) => week.total)),
  };
}

function isIsoForecastCellStruckThrough(entry) {
  if (!entry || entry.forecastColumn < 0 || entry.rowIndex == null) return false;
  const sheet = entry.sheetEntry?.sheet;
  const realRow = entry.sheetEntry?.sourceRowIndexes?.[entry.rowIndex] ?? entry.rowIndex;
  const cell = sheet?.['!data']?.[realRow]?.[entry.forecastColumn]
    || sheet?.[XLSX.utils.encode_cell({ r: realRow, c: entry.forecastColumn })];
  const font = cell?.s?.font || cell?.s?.Font || {};
  const strikeValue = font.strike ?? font.strikeout ?? font.strikeThrough;
  return strikeValue === true || strikeValue === 1 || normalize(strikeValue) === 'true';
}

const ISO_FORECAST_SHEET_NAME = 'HH_TOPSIDE_OVERALL_CS_OFSI';
function isIsoCellStruck(cell, workbook) {
  const style = cell?.s;
  if (!style) return false;
  const font = style.font || style.Font || workbook?.Styles?.Fonts?.[style.fontId] || {};
  const strike = font.strike ?? font.strikeout ?? font.strikeThrough;
  return strike === true || strike === 1 || strike === '1' || normalize(strike) === 'true';
}
// Reads ONLY the HH_TOPSIDE_OVERALL_CS_OFSI sheet and groups dated FORECAST ISSUE DATE rows by module.
function buildIsoForecastMap(workbooks) {
  const map = new Map();
  const targetName = normalize(ISO_FORECAST_SHEET_NAME);
  const exactFileName = normalize(ISO_FILE_PREFIX);
  const availableWorkbooks = (workbooks || []).filter((entry) => entry?.file?.name);
  const exactWorkbooks = availableWorkbooks.filter((entry) => normalize(entry.file.name.replace(/\.[^.]+$/, '')) === exactFileName);
  const fallbackWorkbooks = availableWorkbooks.filter((entry) => normalize(entry.file.name.replace(/\.[^.]+$/, '')).startsWith(exactFileName));
  const sourceWorkbook = (exactWorkbooks.length ? exactWorkbooks : fallbackWorkbooks).slice().sort((a, b) => Number(b.file?.lastModified || 0) - Number(a.file?.lastModified || 0))[0];
  if (!sourceWorkbook) return map;
  const sheets = sourceWorkbook.sheets || [];
  const sourceSheet = sheets.find((sheetEntry) => normalize(sheetEntry.sheetName) === targetName) || sheets.find((sheetEntry) => normalize(sheetEntry.sheetName).includes(targetName));
  if (!sourceSheet) return map;
  const rows = sourceSheet.rows || [];
  let headerRowIndex = -1;
  let forecastColumn = -1;
  for (let i = 0; i < Math.min(rows.length, 60); i += 1) {
    const found = (rows[i] ?? []).findIndex(isIsoForecastHeader);
    if (found >= 0) { headerRowIndex = i; forecastColumn = found; break; }
  }
  if (headerRowIndex < 0) return map;
  const headerRow = rows[headerRowIndex] ?? [];
  const moduleColumn = headerRow.findIndex(isModuleHeader);
  const pipeColumn = findColumnByAliases(headerRow, ISO_PIPE_ALIASES);
  const holdColumn = findColumnByAliases(headerRow, ISO_HOLD_ALIASES);
  const dataCell = (rowIndex, column) => {
    if (column < 0) return null;
    const realRow = sourceSheet.sourceRowIndexes?.[rowIndex] ?? rowIndex;
    return sourceSheet.sheet?.['!data']?.[realRow]?.[column] || sourceSheet.sheet?.[XLSX.utils.encode_cell({ r: realRow, c: column })] || null;
  };
  for (let rowIndex = headerRowIndex + 1; rowIndex < rows.length; rowIndex += 1) {
    const row = rows[rowIndex] ?? [];
    const moduleName = String(moduleColumn >= 0 ? row[moduleColumn] ?? '' : '').trim() || (pipeColumn >= 0 ? isoModuleFromPipeName(row[pipeColumn]) : '');
    if (!moduleName) continue;
    const deleteMarkers = ['linedeleted', 'linedelete', 'deletedline', 'deleteline'];
    if (row.some((value) => { const text = normalize(value); return text && deleteMarkers.some((marker) => text.includes(marker)); })) continue;
    if ([forecastColumn, pipeColumn, moduleColumn].some((column) => isIsoCellStruck(dataCell(rowIndex, column), sourceWorkbook.workbook))) continue;
    const date = parseIsoForecastDate(row[forecastColumn]);
    if (!date) continue;
    const key = normalize(moduleName);
    if (!map.has(key)) map.set(key, []);
    map.get(key).push({ date, hold: holdColumn >= 0 && isHoldValue(row[holdColumn]) });
  }
  return map;
}

function buildIsoLatestFourMonthPlan(moduleData, forecastMap) {
  const items = (forecastMap?.get(normalize(moduleData?.moduleName)) || []).slice().sort((a, b) => a.date - b.date);
  const monthGroups = new Map();
  for (const item of items) {
    const key = `${item.date.getFullYear()}-${String(item.date.getMonth() + 1).padStart(2, '0')}`;
    if (!monthGroups.has(key)) monthGroups.set(key, { year: item.date.getFullYear(), month: item.date.getMonth(), rows: [] });
    monthGroups.get(key).rows.push(item);
  }
  const selectedMonths = [...monthGroups.values()]
    .sort((a, b) => (a.year - b.year) || (a.month - b.month))
    .slice(-4);
  while (selectedMonths.length < 4) selectedMonths.unshift(null);

  const months = selectedMonths.map((group, index) => ({
    key: group ? `${group.year}-${String(group.month + 1).padStart(2, '0')}` : `empty-${index}`,
    slot: `M${index + 1}`,
    label: group ? new Date(group.year, group.month, 1).toLocaleDateString(undefined, { month: 'short', year: 'numeric' }) : 'No Month',
    total: group?.rows.length || 0,
    holds: group?.rows.filter((item) => item.hold).length || 0,
  }));
  return {
    moduleName: moduleData?.moduleName || '',
    months,
    total: months.reduce((sum, month) => sum + month.total, 0),
    holds: months.reduce((sum, month) => sum + month.holds, 0),
  };
}

function isoMetricRows(moduleData, metric) {
  const entries = moduleData?.entries || [];
  if (metric === 'total') return entries;
  if (metric === 'issued') return entries.filter((entry) => entry.issuedColumn >= 0 && isExcelDateValue(entry.row?.[entry.issuedColumn]));
  if (metric === 'revision') return entries.filter((entry) => entry.revisionColumn >= 0 && isNonZeroRevision(entry.row?.[entry.revisionColumn]));
  if (metric === 'hold') return entries.filter((entry) => entry.holdColumn >= 0 && isHoldValue(entry.row?.[entry.holdColumn]));
  if (metric === 'balance') return entries.filter((entry) => !(entry.issuedColumn >= 0 && isExcelDateValue(entry.row?.[entry.issuedColumn])) && !(entry.holdColumn >= 0 && isHoldValue(entry.row?.[entry.holdColumn])));
  return [];
}
function isIsoStageTick(value) {
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

function downloadIsoMetric(moduleData, metric, label) {
  const selected = isoMetricRows(moduleData, metric);
  if (!selected.length) { window.alert('No ' + label + ' rows found for ' + moduleData.moduleName + '.'); return; }
  const output = XLSX.utils.book_new();
  const sheets = new Map();
  for (const entry of selected) {
    if (!sheets.has(entry.sheetEntry)) sheets.set(entry.sheetEntry, []);
    sheets.get(entry.sheetEntry).push(entry.row);
  }
  const used = new Set();
  for (const [sheetEntry, rows] of sheets) {
    const headerRowIndex = sheetEntry.header?.rowIndex ?? 0;
    const headers = sheetEntry.rows.slice(0, headerRowIndex + 1);
    const sheet = XLSX.utils.aoa_to_sheet([...headers, ...rows], { cellDates: true });
    const sourceRange = sheetEntry.sheet['!ref'] ? XLSX.utils.decode_range(sheetEntry.sheet['!ref']) : null;
    if (sourceRange) {
      for (let rowIndex = sourceRange.s.r; rowIndex <= Math.min(headerRowIndex, sourceRange.e.r); rowIndex += 1) {
        for (let columnIndex = sourceRange.s.c; columnIndex <= sourceRange.e.c; columnIndex += 1) {
          const address = XLSX.utils.encode_cell({ r: rowIndex, c: columnIndex });
          if (sheetEntry.sheet[address]?.s && sheet[address]) {
            sheet[address].s = JSON.parse(JSON.stringify(sheetEntry.sheet[address].s));
          }
        }
      }
    }
    if (sheetEntry.sheet['!cols']) sheet['!cols'] = sheetEntry.sheet['!cols'].map((column) => column ? { ...column } : column);
    if (sheetEntry.sheet['!rows']) sheet['!rows'] = sheetEntry.sheet['!rows'].map((row) => row ? { ...row } : row);
    if (sheetEntry.sheet['!merges']) {
      sheet['!merges'] = sheetEntry.sheet['!merges']
        .filter((merge) => merge.e.r <= headerRowIndex)
        .map((merge) => ({ s: { ...merge.s }, e: { ...merge.e } }));
    }
    XLSX.utils.book_append_sheet(output, sheet, safeSheetName(sheetEntry.sheetName, used));
  }
  XLSX.writeFile(output, moduleData.moduleName + ' ' + label + '.xlsx', { compression: true, cellStyles: true });
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
  if (sheetEntry.moduleRows) return sheetEntry.moduleRows.get(moduleKey) || [];
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

function rowHasRemarks(sheetEntry, row) {
  return (sheetEntry.header?.remarksColumns ?? []).some((columnIndex) => hasMeaningfulRemark(row?.[columnIndex]));
}
function isNamedReportWorkbook(entry, expectedName) {
  const fileNameWithoutExtension = entry.file.name.replace(/\.[^.]+$/, '');
  const nameWithoutTrailingNumbers = fileNameWithoutExtension.replace(/[_\s-]*\d+$/, '');
  return normalize(nameWithoutTrailingNumbers) === normalize(expectedName);
}
function isPipeMtoModuleWorkbook(entry) {
  const fileNameWithoutExtension = entry.file.name.replace(/\.[^.]+$/, '');
  return normalize(fileNameWithoutExtension).startsWith(normalize('EHH_PIPE_MTO_REPORT'));
}
function mtoWorkbookCandidates(workbooks) {
  const mtoOnlyWorkbooks = workbooks.filter((entry) =>
    !isNamedReportWorkbook(entry, 'EHH_VALVE_REPORT') &&
    !isNamedReportWorkbook(entry, 'EHH_BRAN_REPORT') &&
    !isNamedReportWorkbook(entry, 'EHH_ATTA_REPORT') &&
    !isNamedReportWorkbook(entry, 'EHH_SUPPORT_REPORT')
  );
  const matched = findReportWorkbook(mtoOnlyWorkbooks, 'MTO Report');
  return matched ? [matched] : mtoOnlyWorkbooks.filter((entry) => entry.sheets.some((sheet) => (sheet.header?.remarksColumns?.length ?? 0) > 0));
}
function countMtoRemarks(workbooks, moduleName) {
  return mtoWorkbookCandidates(workbooks).reduce((total, entry) => total + entry.sheets.reduce((sheetTotal, sheet) => sheetTotal + (sheet.header?.remarksColumns?.length ? rowsForModule(sheet, moduleName).filter((row) => rowHasRemarks(sheet, row)).length : 0), 0), 0);
}
function headerFillRgb(cell) {
  const fill = cell?.s?.fill ?? cell?.s;
  const color = fill?.fgColor ?? fill?.bgColor;
  const rgb = String(color?.rgb ?? '').replace(/^FF/i, '').toUpperCase();
  return /^[0-9A-F]{6}$/.test(rgb) ? rgb : '';
}
function isRedHeaderCell(cell) {
  const rgb = headerFillRgb(cell);
  if (!rgb) return false;
  const red = parseInt(rgb.slice(0, 2), 16);
  const green = parseInt(rgb.slice(2, 4), 16);
  const blue = parseInt(rgb.slice(4, 6), 16);
  return red >= 140 && red >= green * 1.25 && red >= blue * 1.25;
}
function reportRemarksColumns(sheetEntry) {
  if (!sheetEntry.header) return [];
  const headerRow = sheetEntry.rows[sheetEntry.header.rowIndex] ?? [];
  const columns = new Set(sheetEntry.header.remarksColumns ?? []);
  const range = sheetEntry.sheet['!ref'] ? XLSX.utils.decode_range(sheetEntry.sheet['!ref']) : null;
  if (!range) return [...columns];
  for (let columnIndex = range.s.c; columnIndex <= range.e.c; columnIndex += 1) {
    const address = XLSX.utils.encode_cell({ r: sheetEntry.header.rowIndex, c: columnIndex });
    if (isRemarksHeader(headerRow[columnIndex]) || isRedHeaderCell(sheetEntry.sheet[address])) columns.add(columnIndex);
  }
  return [...columns];
}
function valveRowHasRemarks(sheetEntry, row) {
  return reportRemarksColumns(sheetEntry).some((columnIndex) => hasMeaningfulRemark(row?.[columnIndex]));
}
function findValveReportWorkbook(workbooks) {
  return workbooks.find((entry) => isNamedReportWorkbook(entry, 'EHH_VALVE_REPORT')) ?? null;
}
function findPipeBranchReportWorkbook(workbooks) {
  return workbooks.find((entry) => isNamedReportWorkbook(entry, 'EHH_BRAN_REPORT')) ?? null;
}
function findAttaReportWorkbook(workbooks) {
  return workbooks.find((entry) => isNamedReportWorkbook(entry, 'EHH_ATTA_REPORT')) ?? null;
}
function findPrimarySupportReportWorkbook(workbooks) {
  return workbooks.find((entry) => isNamedReportWorkbook(entry, 'EHH_SUPPORT_REPORT')) ?? null;
}
function findElbowBendReportWorkbook(workbooks) {
  return workbooks.find((entry) => {
    const fileNameWithoutExtension = entry.file.name.replace(/\.[^.]+$/, '');
    return normalize(fileNameWithoutExtension).startsWith(normalize('EHH_ELBOW_BEND_ANG_REPORT'));
  }) ?? null;
}
function findEquipmentReportWorkbook(workbooks) {
  return workbooks.find((entry) => {
    const fileNameWithoutExtension = entry.file.name.replace(/\.[^.]+$/, '');
    return normalize(fileNameWithoutExtension).startsWith(normalize('EHH_EQUIPMENT_REPORT'));
  }) ?? null;
}
function countValveRemarks(workbooks, moduleName) {
  const workbookEntry = findValveReportWorkbook(workbooks);
  if (!workbookEntry) return 0;
  return workbookEntry.sheets.reduce((total, sheetEntry) => total + (reportRemarksColumns(sheetEntry).length ? rowsForModule(sheetEntry, moduleName).filter((row) => valveRowHasRemarks(sheetEntry, row)).length : 0), 0);
}
function countReportRemarks(workbookEntry, moduleName) {
  if (!workbookEntry) return 0;
  return workbookEntry.sheets.reduce((total, sheetEntry) => total + (reportRemarksColumns(sheetEntry).length ? rowsForModule(sheetEntry, moduleName).filter((row) => valveRowHasRemarks(sheetEntry, row)).length : 0), 0);
}
function countPipeBranchRemarks(workbooks, moduleName) {
  return countReportRemarks(findPipeBranchReportWorkbook(workbooks), moduleName);
}
function countAttaRemarks(workbooks, moduleName) {
  return countReportRemarks(findAttaReportWorkbook(workbooks), moduleName);
}
function countPrimarySupportRemarks(workbooks, moduleName) {
  return countReportRemarks(findPrimarySupportReportWorkbook(workbooks), moduleName);
}
function countElbowBendRemarks(workbooks, moduleName) {
  return countReportRemarks(findElbowBendReportWorkbook(workbooks), moduleName);
}
function countEquipmentOrientationRemarks(workbooks, moduleName) {
  return countReportRemarks(findEquipmentReportWorkbook(workbooks), moduleName);
}
function copyExactHeaderFill(sourceSheet, outputSheet, headerRowIndex) {
  const range = sourceSheet['!ref'] ? XLSX.utils.decode_range(sourceSheet['!ref']) : null;
  if (!range) return;
  for (let r = range.s.r; r <= Math.min(headerRowIndex, range.e.r); r += 1) {
    for (let c = range.s.c; c <= range.e.c; c += 1) {
      const address = XLSX.utils.encode_cell({ r, c });
      if (sourceSheet[address]?.s && outputSheet[address]) outputSheet[address].s = { fill: JSON.parse(JSON.stringify(sourceSheet[address].s)) };
    }
  }
  if (sourceSheet['!rows']) outputSheet['!rows'] = sourceSheet['!rows'].slice(0, headerRowIndex + 1).map((row) => row ? { ...row } : row);
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

function createMtoRemarksWorkbook(workbooks, moduleName) {
  const output = XLSX.utils.book_new();
  const usedNames = new Set();
  let exportedRows = 0;
  for (const entry of mtoWorkbookCandidates(workbooks)) {
    for (const sheetEntry of entry.sheets) {
      if (!sheetEntry.header?.remarksColumns?.length) continue;
      const matchingRows = rowsForModule(sheetEntry, moduleName).filter((row) => rowHasRemarks(sheetEntry, row));
      if (!matchingRows.length) continue;
      const headerRows = sheetEntry.rows.slice(0, sheetEntry.header.rowIndex + 1);
      const outputSheet = XLSX.utils.aoa_to_sheet([...headerRows, ...matchingRows], { cellDates: true });
      copyExactHeaderFill(sheetEntry.sheet, outputSheet, sheetEntry.header.rowIndex);
      if (sheetEntry.sheet['!cols']) outputSheet['!cols'] = sheetEntry.sheet['!cols'];
      if (sheetEntry.sheet['!merges']) outputSheet['!merges'] = sheetEntry.sheet['!merges'].filter((merge) => merge.e.r <= sheetEntry.header.rowIndex);
      XLSX.utils.book_append_sheet(output, outputSheet, safeSheetName(sheetEntry.sheetName, usedNames));
      exportedRows += matchingRows.length;
    }
  }
  if (!exportedRows) return false;
  XLSX.writeFile(output, `${moduleName} MTO Report Remarks.xlsx`, { compression: true, cellStyles: true });
  return true;
}

function createValveRemarksWorkbook(workbooks, moduleName) {
  const workbookEntry = findValveReportWorkbook(workbooks);
  if (!workbookEntry) return false;
  const output = XLSX.utils.book_new();
  const usedNames = new Set();
  let exportedRows = 0;
  for (const sheetEntry of workbookEntry.sheets) {
    if (!reportRemarksColumns(sheetEntry).length) continue;
    const matchingRows = rowsForModule(sheetEntry, moduleName).filter((row) => valveRowHasRemarks(sheetEntry, row));
    if (!matchingRows.length) continue;
    const headerRows = sheetEntry.rows.slice(0, sheetEntry.header.rowIndex + 1);
    const outputSheet = XLSX.utils.aoa_to_sheet([...headerRows, ...matchingRows], { cellDates: true });
    copyExactHeaderFill(sheetEntry.sheet, outputSheet, sheetEntry.header.rowIndex);
    if (sheetEntry.sheet['!cols']) outputSheet['!cols'] = sheetEntry.sheet['!cols'];
    if (sheetEntry.sheet['!merges']) outputSheet['!merges'] = sheetEntry.sheet['!merges'].filter((merge) => merge.e.r <= sheetEntry.header.rowIndex);
    XLSX.utils.book_append_sheet(output, outputSheet, safeSheetName(sheetEntry.sheetName, usedNames));
    exportedRows += matchingRows.length;
  }
  if (!exportedRows) return false;
  XLSX.writeFile(output, `${moduleName} Valve Report Remarks.xlsx`, { compression: true, cellStyles: true });
  return true;
}

function createNamedRemarksWorkbook(workbookEntry, moduleName, reportName) {
  if (!workbookEntry) return false;
  const output = XLSX.utils.book_new();
  const usedNames = new Set();
  let exportedRows = 0;
  for (const sheetEntry of workbookEntry.sheets) {
    if (!reportRemarksColumns(sheetEntry).length) continue;
    const matchingRows = rowsForModule(sheetEntry, moduleName).filter((row) => valveRowHasRemarks(sheetEntry, row));
    if (!matchingRows.length) continue;
    const headerRows = sheetEntry.rows.slice(0, sheetEntry.header.rowIndex + 1);
    const outputSheet = XLSX.utils.aoa_to_sheet([...headerRows, ...matchingRows], { cellDates: true });
    copyExactHeaderFill(sheetEntry.sheet, outputSheet, sheetEntry.header.rowIndex);
    if (sheetEntry.sheet['!cols']) outputSheet['!cols'] = sheetEntry.sheet['!cols'];
    if (sheetEntry.sheet['!merges']) outputSheet['!merges'] = sheetEntry.sheet['!merges'].filter((merge) => merge.e.r <= sheetEntry.header.rowIndex);
    XLSX.utils.book_append_sheet(output, outputSheet, safeSheetName(sheetEntry.sheetName, usedNames));
    exportedRows += matchingRows.length;
  }
  if (!exportedRows) return false;
  XLSX.writeFile(output, `${moduleName} ${reportName} Remarks.xlsx`, { compression: true, cellStyles: true });
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

function isPipeE3dStatusWorkbook(workbookEntry) {
  const fileNameWithoutExtension = workbookEntry.file.name.replace(/\.[^.]+$/, '');
  return normalize(fileNameWithoutExtension).startsWith('pipee3dstatusehh');
}

function pipeModuleAndSize(value) {
  const parts = String(value ?? '').trim().replace(/^\/+/, '').split('-').map((part) => part.trim()).filter(Boolean);
  if (parts.length < 3) return null;
  return { moduleName: parts[0], bucket: sizeBucket(parts[2]) };
}

const FLUID_HEADER_ALIASES = ['fluid code', 'fluidcode', 'fluid', 'service code', 'servicecode', 'service'];

function pipeLineParts(value) {
  const parts = String(value ?? '').trim().replace(/^\/+/, '').split('-').map((part) => part.trim()).filter(Boolean);
  if (parts.length < 3) return null;
  const rawSize = parts[2];
  const numericSize = Number(String(rawSize).match(/\d+(?:\.\d+)?/)?.[0]);
  if (!Number.isFinite(numericSize)) return null;
  return { moduleName: parts[0], rawSize, numericSize, fallbackFluidCode: parts[3] || '' };
}

function formatPipeSize(value) {
  return `${value}\"`;
}

function getFluidCodeStatus(workbooks, moduleName) {
  const counts = new Map();
  const sizes = new Set();
  if (!moduleName) return { sizes: [], fluidCodes: [], counts: {} };

  for (const workbookEntry of workbooks) {
    if (!isPipeE3dStatusWorkbook(workbookEntry)) continue;
    for (const sheetEntry of workbookEntry.sheets) {
      let headerRowIndex = -1;
      let pipeColumn = -1;
      let fluidColumn = -1;
      for (let rowIndex = 0; rowIndex < Math.min(sheetEntry.rows.length, 40); rowIndex += 1) {
        const row = sheetEntry.rows[rowIndex] ?? [];
        const foundPipeColumn = findColumnByAliases(row, ['pipe']);
        if (foundPipeColumn >= 0) {
          headerRowIndex = rowIndex;
          pipeColumn = foundPipeColumn;
          fluidColumn = findColumnByAliases(row, FLUID_HEADER_ALIASES);
          break;
        }
      }
      if (headerRowIndex < 0) continue;
      for (const row of sheetEntry.rows.slice(headerRowIndex + 1)) {
        const parsed = pipeLineParts(row?.[pipeColumn]);
        if (!parsed || normalize(parsed.moduleName) !== normalize(moduleName)) continue;
        const fluidCode = String(fluidColumn >= 0 ? row?.[fluidColumn] : parsed.fallbackFluidCode).trim().toUpperCase();
        if (!fluidCode) continue;
        sizes.add(parsed.numericSize);
        if (!counts.has(fluidCode)) counts.set(fluidCode, new Map());
        const fluidCounts = counts.get(fluidCode);
        fluidCounts.set(parsed.numericSize, (fluidCounts.get(parsed.numericSize) || 0) + 1);
      }
    }
  }

  const sortedSizes = [...sizes].sort((a, b) => a - b);
  const fluidCodes = [...counts.keys()].sort((a, b) => a.localeCompare(b));
  return {
    sizes: sortedSizes,
    fluidCodes,
    counts: Object.fromEntries(fluidCodes.map((fluidCode) => [fluidCode, Object.fromEntries(sortedSizes.map((size) => [size, counts.get(fluidCode)?.get(size) || 0]))])),
  };
}

function createFluidCodeWorkbook(moduleName, data) {
  const rows = [
    ['Fluid Code', ...data.sizes.map(formatPipeSize), 'Grand Total'],
    ...data.fluidCodes.map((fluidCode) => [fluidCode, ...data.sizes.map((size) => data.counts[fluidCode]?.[size] || 0), data.sizes.reduce((sum, size) => sum + (data.counts[fluidCode]?.[size] || 0), 0)]),
    ['GRAND TOTAL', ...data.sizes.map((size) => data.fluidCodes.reduce((sum, fluidCode) => sum + (data.counts[fluidCode]?.[size] || 0), 0)), data.fluidCodes.reduce((total, fluidCode) => total + data.sizes.reduce((sum, size) => sum + (data.counts[fluidCode]?.[size] || 0), 0), 0)],
  ];
  const output = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(output, XLSX.utils.aoa_to_sheet(rows), 'Fluid Code');
  XLSX.writeFile(output, `${moduleName} Fluid Code.xlsx`, { compression: true });
}


const STATUS_CODE_HEADER_ALIASES = ['status code', 'statuscode'];
function getLineStatus(workbooks, moduleName) {
  const nestedCounts = new Map();
  const sizes = new Set();
  const statusCodes = new Set();
  if (!moduleName) return { sizes: [], statusCodes: [], counts: {}, grandTotal: 0 };
  for (const workbookEntry of workbooks) {
    if (!isPipeE3dStatusWorkbook(workbookEntry)) continue;
    for (const sheetEntry of workbookEntry.sheets) {
      let headerRowIndex = -1;
      let pipeColumn = -1;
      let statusCodeColumn = -1;
      for (let rowIndex = 0; rowIndex < Math.min(sheetEntry.rows.length, 40); rowIndex += 1) {
        const row = sheetEntry.rows[rowIndex] ?? [];
        const foundPipeColumn = findColumnByAliases(row, ['pipe']);
        const foundStatusCodeColumn = findColumnByAliases(row, STATUS_CODE_HEADER_ALIASES);
        if (foundPipeColumn >= 0 && foundStatusCodeColumn >= 0) {
          headerRowIndex = rowIndex;
          pipeColumn = foundPipeColumn;
          statusCodeColumn = foundStatusCodeColumn;
          break;
        }
      }
      if (headerRowIndex < 0) continue;
      for (const row of sheetEntry.rows.slice(headerRowIndex + 1)) {
        const parsed = pipeLineParts(row?.[pipeColumn]);
        if (!parsed || normalize(parsed.moduleName) !== normalize(moduleName)) continue;
        const statusCode = String(row?.[statusCodeColumn] ?? '').trim().replace(/^\/+/, '').toUpperCase();
        if (!statusCode) continue;
        const size = parsed.numericSize;
        sizes.add(size);
        statusCodes.add(statusCode);
        if (!nestedCounts.has(size)) nestedCounts.set(size, new Map());
        nestedCounts.get(size).set(statusCode, (nestedCounts.get(size).get(statusCode) || 0) + 1);
      }
    }
  }
  const sortedSizes = [...sizes].sort((a, b) => a - b);
  const sortedStatusCodes = [...statusCodes].sort((a, b) => a.localeCompare(b, undefined, { numeric: true }));
  const counts = Object.fromEntries(sortedSizes.map((size) => [size, Object.fromEntries(sortedStatusCodes.map((statusCode) => [statusCode, nestedCounts.get(size)?.get(statusCode) || 0]))]));
  const grandTotal = sortedSizes.reduce((total, size) => total + sortedStatusCodes.reduce((sum, statusCode) => sum + (counts[size]?.[statusCode] || 0), 0), 0);
  return { sizes: sortedSizes, statusCodes: sortedStatusCodes, counts, grandTotal };
}
function createLineStatusWorkbook(moduleName, data) {
  const rows = [
    ['SIZE', ...data.statusCodes, 'Grand Total'],
    ...data.sizes.map((size) => [formatPipeSize(size), ...data.statusCodes.map((statusCode) => data.counts[size]?.[statusCode] || 0), data.statusCodes.reduce((sum, statusCode) => sum + (data.counts[size]?.[statusCode] || 0), 0)]),
    ['GRAND TOTAL', ...data.statusCodes.map((statusCode) => data.sizes.reduce((sum, size) => sum + (data.counts[size]?.[statusCode] || 0), 0)), data.grandTotal],
  ];
  const output = XLSX.utils.book_new();
  const sheet = XLSX.utils.aoa_to_sheet(rows);
  const range = XLSX.utils.decode_range(sheet['!ref']);
  const border = { top: { style: 'thin', color: { rgb: 'FFFFFF' } }, bottom: { style: 'thin', color: { rgb: 'FFFFFF' } }, left: { style: 'thin', color: { rgb: 'FFFFFF' } }, right: { style: 'thin', color: { rgb: 'FFFFFF' } } };
  for (let r = range.s.r; r <= range.e.r; r += 1) for (let c = range.s.c; c <= range.e.c; c += 1) {
    const address = XLSX.utils.encode_cell({ r, c });
    const isHeader = r === 0;
    const isGrandTotal = r === range.e.r;
    sheet[address].s = {
      font: { bold: isHeader || isGrandTotal || c === 0, color: { rgb: isHeader ? 'FFFFFF' : '17324D' } },
      fill: { patternType: 'solid', fgColor: { rgb: isHeader ? '2498CB' : isGrandTotal ? 'BFD7EA' : 'DCEAF5' } },
      alignment: { horizontal: 'center', vertical: 'center', wrapText: true }, border,
    };
  }
  sheet['!cols'] = rows[0].map(() => ({ wch: 16 }));
  sheet['!rows'] = rows.map(() => ({ hpt: 28 }));
  XLSX.utils.book_append_sheet(output, sheet, 'Line Status');
  XLSX.writeFile(output, `${moduleName} Line Status.xlsx`, { compression: true, cellStyles: true });
}

function getSizeStatus(workbooks, moduleName) {
  const result = {
    'MODELLED': Object.fromEntries(SIZE_BUCKETS.map((bucket) => [bucket, 0])),
    'NOT MODELLED': Object.fromEntries(SIZE_BUCKETS.map((bucket) => [bucket, 0])),
  };

  if (!moduleName) return result;

  for (const workbookEntry of workbooks) {
    for (const sheetEntry of workbookEntry.sheets) {
      if (isPipeE3dStatusWorkbook(workbookEntry)) {
        let pipeHeaderRowIndex = -1;
        let pipeColumn = -1;
        for (let rowIndex = 0; rowIndex < Math.min(sheetEntry.rows.length, 40); rowIndex += 1) {
          const foundColumn = findColumnByAliases(sheetEntry.rows[rowIndex] ?? [], ['pipe']);
          if (foundColumn >= 0) {
            pipeHeaderRowIndex = rowIndex;
            pipeColumn = foundColumn;
            break;
          }
        }
        if (pipeHeaderRowIndex < 0) continue;
        for (const row of sheetEntry.rows.slice(pipeHeaderRowIndex + 1)) {
          const parsed = pipeModuleAndSize(row?.[pipeColumn]);
          if (parsed && normalize(parsed.moduleName) === normalize(moduleName) && parsed.bucket) {
            result['MODELLED'][parsed.bucket] += 1;
          }
        }
        continue;
      }

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

function createSizeStatusWorkbook(moduleName, status, planned) {
  const calculatedNotModelled = Object.fromEntries(SIZE_BUCKETS.map((bucket) => [bucket, Math.max(0, planned[bucket] - status['MODELLED'][bucket])]));
  const rows = [
    [`${moduleName}-Status`, ...SIZE_BUCKETS, 'Grand Total'],
    ['MODELLED', ...SIZE_BUCKETS.map((bucket) => status['MODELLED'][bucket]), Object.values(status['MODELLED']).reduce((sum, value) => sum + value, 0)],
    ['NOT MODELLED', ...SIZE_BUCKETS.map((bucket) => calculatedNotModelled[bucket]), SIZE_BUCKETS.reduce((sum, bucket) => sum + calculatedNotModelled[bucket], 0)],
    ['PLANNED', ...SIZE_BUCKETS.map((bucket) => planned[bucket]), SIZE_BUCKETS.reduce((sum, bucket) => sum + planned[bucket], 0)],
    ['GRAND TOTAL', ...SIZE_BUCKETS.map((bucket) => status['MODELLED'][bucket] + calculatedNotModelled[bucket]), SIZE_BUCKETS.reduce((sum, bucket) => sum + status['MODELLED'][bucket] + calculatedNotModelled[bucket], 0)],
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
  // Primary read path: Vercel server API.
  try {
    const response = await fetch('/api/modules', { headers: { Accept: 'application/json' } });
    if (response.ok) {
      const data = await response.json();
      const modules = Array.isArray(data?.modules) ? data.modules.filter((value) => typeof value === 'string' && value.trim()) : [];
      if (modules.length) return { modules, source: data?.source || 'central-api' };
    }
  } catch (error) {
    console.warn('Central API module read failed; trying Supabase public read.', error);
  }

  // Fallback read path: directly read the public active module rows from Supabase.
  // This keeps the module list visible to every browser even if the Vercel function is unavailable.
  const supabaseUrl = import.meta.env.VITE_SUPABASE_URL;
  const supabaseKey = import.meta.env.VITE_SUPABASE_PUBLISHABLE_KEY || import.meta.env.VITE_SUPABASE_ANON_KEY;
  if (!supabaseUrl || !supabaseKey) throw new Error('Supabase public read configuration is missing.');

  const response = await fetch(`${supabaseUrl.replace(/\/$/, '')}/rest/v1/modules?select=module_name,sort_order&active=eq.true&order=sort_order.asc`, {
    headers: {
      apikey: supabaseKey,
      Authorization: `Bearer ${supabaseKey}`,
      Accept: 'application/json',
    },
  });
  if (!response.ok) {
    const body = await response.text().catch(() => '');
    throw new Error(`Supabase module read returned ${response.status}${body ? `: ${body}` : ''}`);
  }
  const rows = await response.json();
  const modules = Array.isArray(rows) ? rows.map((row) => row?.module_name).filter((value) => typeof value === 'string' && value.trim()) : [];
  return { modules, source: 'supabase-direct' };
}

async function publishCentralModules(modules, token) {
  const response = await fetch('/api/modules', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json', Accept: 'application/json', 'x-admin-token': token },
    body: JSON.stringify({ modules }),
  });
  let data = null;
  try { data = await response.json(); } catch { /* ignore malformed error bodies */ }
  if (response.status === 404) {
    return { ok: false, skipped: true, reason: 'module-api-unavailable' };
  }
  if (!response.ok) {
    throw new Error(data?.error || `Module publish failed (${response.status})`);
  }
  return data;
}

export default function App() {
  const [page, setPage] = useState('home');
  const [title, setTitle] = useState('');
  const [from, setFrom] = useState('home');
  const [supportProject, setSupportProject] = useState({ code: '4193', name: 'HAMMER HEAD' });
  const [excelSession, setExcelSession] = useState(() => ({ folderName: '', workbooks: [], modules: readCachedModules() }));
  const go = (nextPage) => { setPage(nextPage); window.scrollTo(0, 0); };
  const blank = (nextTitle, previousPage) => { setTitle(nextTitle); setFrom(previousPage); go('blank'); };
  if (page === 'reports') return <Reports back={() => go('home')} open={(nextTitle) => blank(nextTitle, 'reports')} excelSession={excelSession} setExcelSession={setExcelSession} />;
  if (page === 'iso-planning-dashboard') return <IsoPlanningDashboardPage workbooks={excelSession.workbooks} back={() => go('home')} />;
  if (page === 'support-tools') return <SupportToolsPage projectLabel={supportProject.code + ' · ' + supportProject.name} open={(nextTitle) => blank(nextTitle, 'support-tools')} back={() => go('home')} />;
  if (page === 'iso-planning') return <IsoPlanningManagementPage workbooks={excelSession.workbooks} back={() => go('home')} open={(nextTitle) => blank(nextTitle, 'iso-planning')} />;
  if (page === 'blank') return <Layout title={title} sub="Content Placeholder" back={() => go(from)}><div className="blank"><b>RGT</b><p>CONTENT PLACEHOLDER</p><h2>{title}</h2><span>No project data or report has been added to this page yet.</span><button onClick={() => go(from)}>Return to Previous Page</button></div></Layout>;
  return <Home isoPlanning={() => go('iso-planning-dashboard')} reports={() => go('reports')} supportTools={(code, name) => { setSupportProject({ code, name }); go('support-tools'); }} blank={(nextTitle) => blank(nextTitle, 'home')} />;
}

function Home({ reports, isoPlanning, supportTools, blank }) {
  const action = (code, name, actionName) => {
    if (code === '4193' && actionName === 'Reports and Status') return reports();
    if (code === '4193' && actionName === 'ISO Planning and Management') return isoPlanning();
    if (actionName === 'SUPPORT TOOLS') return supportTools(code, name);
    return blank(`${code} · ${name} · ${actionName}`);
  };
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
  useEffect(() => () => window.clearTimeout(noticeTimerRef.current), []);

  useEffect(() => {
    let cancelled = false;

    const refreshCentralModules = async () => {
      try {
        const result = await fetchCentralModules();
        if (cancelled) return;
        if (result.modules.length) {
          cacheModules(result.modules);
          setExcelSession((current) => ({ ...current, modules: result.modules }));
          setSelected((current) => result.modules.some((item) => normalize(item) === normalize(current)) ? current : null);
          setCentralStatus('connected');
        } else {
          setCentralStatus('empty');
        }
      } catch (error) {
        console.warn('Central module list unavailable; using cached modules.', error);
        if (!cancelled) setCentralStatus('offline');
      }
    };

    refreshCentralModules();
    const intervalId = window.setInterval(refreshCentralModules, 30000);

    return () => {
      cancelled = true;
      window.clearInterval(intervalId);
    };
  }, [setExcelSession]);
  const list = useMemo(() => modules
    .filter((moduleName) => normalize(moduleName) !== 'grandtotal' && moduleName.toLowerCase().includes(query.trim().toLowerCase()))
    .sort((a, b) => a.localeCompare(b, undefined, { numeric: true, sensitivity: 'base' })), [modules, query]);
  const COUNT_ITEMS = useMemo(() => GROUPS.flatMap(([, items]) => items).filter((item) => item !== 'Size Status'), []);
  const fluidCodeStatus = useMemo(() => getFluidCodeStatus(workbooks, selected), [workbooks, selected]);
  const fluidCodeGrandTotal = useMemo(() => fluidCodeStatus.fluidCodes.reduce((total, fluidCode) => total + fluidCodeStatus.sizes.reduce((sum, size) => sum + (fluidCodeStatus.counts[fluidCode]?.[size] || 0), 0), 0), [fluidCodeStatus]);
  const lineStatus = useMemo(() => getLineStatus(workbooks, selected), [workbooks, selected]);
  const counts = useMemo(() => Object.fromEntries(COUNT_ITEMS.map((item) => [item, selected ? (item === 'MTO Report' ? countMtoRemarks(workbooks, selected) : item === 'Valve Report' ? countValveRemarks(workbooks, selected) : item === 'Pipe Branch Report' ? countPipeBranchRemarks(workbooks, selected) : item === 'ATTA Report' ? countAttaRemarks(workbooks, selected) : item === 'Primary Support Report' ? countPrimarySupportRemarks(workbooks, selected) : item === 'Elbow and Bend Report' ? countElbowBendRemarks(workbooks, selected) : item === 'Equipment Orientation and Position Report' ? countEquipmentOrientationRemarks(workbooks, selected) : item === 'Fluid Code' ? fluidCodeGrandTotal : item === 'Line Status' ? lineStatus.grandTotal : countForModule(findReportWorkbook(workbooks, item), selected)) : 0])), [COUNT_ITEMS, workbooks, selected, fluidCodeGrandTotal, lineStatus]);

  const connectFolder = () => inputRef.current?.click();
  const loadFolder = async (event) => {
    const files = Array.from(event.target.files ?? []).filter((file) => /\.(xlsx|xls|xlsm|xlsb)$/i.test(file.name) && !file.name.startsWith('~$'));
    if (!files.length) { showNotice('warning', 'No Excel file found', 'Select a folder containing .xlsx, .xls, .xlsm or .xlsb files.'); return; }
    setBusy(true);
    showNotice('loading', 'Connecting Excel folder', 'Preparing ' + files.length + ' Excel file' + (files.length === 1 ? '' : 's') + '...', true);
    try {
      const parsed = await parseWorkbooksFast(files, (completed, total, fileName) => {
        showNotice('loading', 'Reading Excel files', completed + ' of ' + total + ' processed · ' + fileName, true);
      });
      const moduleMap = new Map();
      for (const workbookEntry of parsed) {
        if (!isPipeMtoModuleWorkbook(workbookEntry)) continue;
        for (const sheetEntry of workbookEntry.sheets) {
          if (!sheetEntry.header) continue;
          for (const row of sheetEntry.rows.slice(sheetEntry.header.rowIndex + 1)) {
            const raw = String(row?.[sheetEntry.header.moduleColumn] ?? '').trim();
            const key = normalize(raw);
            if (key && !moduleMap.has(key)) moduleMap.set(key, raw);
          }
        }
      }
      // Preserve the module order from the Master Excel.
      const dynamicModules = [...moduleMap.values()];
      const firstPath = files[0].webkitRelativePath || files[0].name;
      const connectedFolderName = firstPath.includes('/') ? firstPath.split('/')[0] : 'Selected Excel files';
      setExcelSession({ folderName: connectedFolderName, workbooks: parsed, modules: dynamicModules });
      cacheModules(dynamicModules);
      setSelected((current) => dynamicModules.some((item) => normalize(item) === normalize(current)) ? current : null);
      if (!dynamicModules.length) {
        showNotice('warning', 'Folder connected', 'No Module or Module Name column was found in EHH_PIPE_MTO_REPORT.');
        return;
      }

      let adminToken = '';
      try { adminToken = sessionStorage.getItem(ADMIN_TOKEN_KEY) || ''; } catch { /* ignore */ }
      if (!adminToken) {
        adminToken = await requestAdminToken();
        if (adminToken) {
          try { sessionStorage.setItem(ADMIN_TOKEN_KEY, adminToken); } catch { /* ignore */ }
        }
      }
      if (adminToken) {
        try {
          const publishResult = await publishCentralModules(dynamicModules, adminToken);
          if (publishResult?.skipped) {
            setCentralStatus('offline');
          } else {
            setCentralStatus('connected');
            showNotice('success', 'Excel folder connected', files.length + ' files processed. Master module list published for all users.');
          }
        } catch (error) {
          console.error(error);
          try { sessionStorage.removeItem(ADMIN_TOKEN_KEY); } catch { /* ignore */ }
          setCentralStatus('offline');
          showNotice('warning', 'Connected locally', 'Excel data is ready, but central publish failed: ' + error.message);
        }
      } else {
        showNotice('success', 'Excel folder connected', files.length + ' files processed successfully. ' + dynamicModules.length + ' modules are ready.');
      }
    } catch (error) {
      console.error(error);
      showNotice('error', 'Excel connection failed', 'The folder could not be read. Check the workbook format and try again.');
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
    if (reportName === 'MTO Report') {
      if (!createMtoRemarksWorkbook(workbooks, selected)) window.alert(`No ${selected} row with remarks was found for MTO Report.`);
      return;
    }
    if (reportName === 'Valve Report') {
      if (!findValveReportWorkbook(workbooks)) {
        window.alert('No EHH_VALVE_REPORT Excel file was found in the connected folder.');
        return;
      }
      if (!createValveRemarksWorkbook(workbooks, selected)) window.alert(`No ${selected} row with remarks was found under red header columns in Valve Report.`);
      return;
    }
    if (reportName === 'Pipe Branch Report') {
      const workbookEntry = findPipeBranchReportWorkbook(workbooks);
      if (!workbookEntry) { window.alert('No EHH_BRAN_REPORT Excel file was found in the connected folder.'); return; }
      if (!createNamedRemarksWorkbook(workbookEntry, selected, 'Pipe Branch Report')) window.alert(`No ${selected} row with remarks was found for Pipe Branch Report.`);
      return;
    }
    if (reportName === 'ATTA Report') {
      const workbookEntry = findAttaReportWorkbook(workbooks);
      if (!workbookEntry) { window.alert('No EHH_ATTA_REPORT Excel file was found in the connected folder.'); return; }
      if (!createNamedRemarksWorkbook(workbookEntry, selected, 'ATTA Report')) window.alert(`No ${selected} row with remarks was found for ATTA Report.`);
      return;
    }
    if (reportName === 'Primary Support Report') {
      const workbookEntry = findPrimarySupportReportWorkbook(workbooks);
      if (!workbookEntry) { window.alert('No EHH_SUPPORT_REPORT Excel file was found in the connected folder.'); return; }
      if (!createNamedRemarksWorkbook(workbookEntry, selected, 'Primary Support Report')) window.alert(`No ${selected} row with remarks was found for Primary Support Report.`);
      return;
    }
    if (reportName === 'Elbow and Bend Report') {
      const workbookEntry = findElbowBendReportWorkbook(workbooks);
      if (!workbookEntry) { window.alert('No EHH_ELBOW_BEND_ANG_REPORT Excel file was found in the connected folder.'); return; }
      if (!createNamedRemarksWorkbook(workbookEntry, selected, 'Elbow and Bend Report')) window.alert('No ' + selected + ' row with remarks was found for Elbow and Bend Report.');
      return;
    }
    if (reportName === 'Equipment Orientation and Position Report') {
      const workbookEntry = findEquipmentReportWorkbook(workbooks);
      if (!workbookEntry) { window.alert('No EHH_EQUIPMENT_REPORT Excel file was found in the connected folder.'); return; }
      if (!createNamedRemarksWorkbook(workbookEntry, selected, 'Equipment Orientation and Position Report')) window.alert('No ' + selected + ' row with remarks was found for Equipment Orientation and Position Report.');
      return;
    }
    const workbookEntry = findReportWorkbook(workbooks, reportName);
    if (!workbookEntry) { window.alert(`No master Excel found for ${reportName}.`); return; }
    if (!createFilteredWorkbook(workbookEntry, selected, reportName)) window.alert(`No ${selected} row found for ${reportName}.`);
  };

  const sizeStatus = useMemo(() => getSizeStatus(workbooks, selected), [workbooks, selected]);
  if (statusView === 'Size Status' && selected) {
    return <SizeStatusPage moduleName={selected} status={sizeStatus} back={() => setStatusView(null)} />;
  }
  if (statusView === 'Fluid Code' && selected) {
    return <FluidCodePage moduleName={selected} data={fluidCodeStatus} back={() => setStatusView(null)} />;
  }
  if (statusView === 'Line Status' && selected) {
    return <LineStatusPage moduleName={selected} data={lineStatus} back={() => setStatusView(null)} />;
  }
  if (statusView === 'SUPPORT TOOLS' && selected) {
    return <SupportToolsPage moduleName={selected} open={open} back={() => setStatusView(null)} />;
  }

  const openIntegrated = (item) => {
    if ((item === 'Size Status' || item === 'Fluid Code' || item === 'Line Status' || item === 'SUPPORT TOOLS') && selected) {
      setStatusView(item);
      return;
    }
    open(item);
  };

  const moduleSourceText = modules.length
    ? (centralStatus === 'connected' ? 'Connected: MASTER MODULE LIST' : folderName ? `Connected: ${folderName}` : 'Module list loaded')
    : 'No master module list connected';

  return <>
    <style>{`
      .rgt-notice-stack{position:absolute;top:12px;right:12px;z-index:20;width:min(320px,calc(100vw - 24px));pointer-events:none}.rgt-notice{pointer-events:auto;min-height:48px;display:grid;grid-template-columns:34px minmax(0,1fr) 24px;gap:9px;align-items:center;padding:8px 9px 8px 10px;border:1px solid rgba(148,163,184,.3);border-radius:11px;background:rgba(8,47,73,.97);box-shadow:0 10px 28px rgba(15,23,42,.24);color:#fff;backdrop-filter:blur(14px);animation:rgtNoticeRight .26s cubic-bezier(.2,.8,.2,1)}.rgt-notice-icon{width:34px;height:34px;display:grid;place-items:center;border-radius:12px;background:#0ea5e9;font-weight:900}.rgt-notice.success .rgt-notice-icon{background:#10b981}.rgt-notice.warning .rgt-notice-icon{background:#f59e0b}.rgt-notice.error .rgt-notice-icon{background:#ef4444}.rgt-notice.loading .rgt-notice-icon:after{content:'';width:18px;height:18px;border:3px solid rgba(255,255,255,.42);border-top-color:#fff;border-radius:50%;animation:rgtSpin .8s linear infinite}.rgt-notice strong{display:block;font-size:12px;line-height:1.25}.rgt-notice p{margin:2px 0 0;color:#cbd5e1;font-size:10px;line-height:1.45;word-break:break-word}.rgt-notice-close{border:0;background:transparent;color:#94a3b8;font-size:18px;cursor:pointer}.rgt-token-backdrop{position:fixed;inset:0;z-index:9999;display:grid;place-items:center;padding:20px;background:rgba(15,23,42,.55);backdrop-filter:blur(6px);animation:rgtFade .2s ease}.rgt-token-modal{width:min(430px,100%);padding:24px;border-radius:20px;background:#fff;box-shadow:0 30px 90px rgba(15,23,42,.38)}.rgt-token-modal p{margin:0;color:#0284c7;font-size:12px;font-weight:800;letter-spacing:.12em}.rgt-token-modal h3{margin:7px 0 8px;color:#0f172a}.rgt-token-modal span{display:block;color:#64748b;font-size:13px;line-height:1.5}.rgt-token-modal input{box-sizing:border-box;width:100%;margin:18px 0;padding:12px 14px;border:1px solid #cbd5e1;border-radius:11px;outline:none}.rgt-token-modal input:focus{border-color:#0ea5e9;box-shadow:0 0 0 3px rgba(14,165,233,.14)}.rgt-token-actions{display:flex;justify-content:flex-end;gap:10px}.rgt-token-actions button{padding:10px 15px;border:0;border-radius:10px;cursor:pointer;font-weight:700}.rgt-token-actions button:first-child{background:#e2e8f0;color:#334155}.rgt-token-actions button:last-child{background:#0284c7;color:#fff}@keyframes rgtSpin{to{transform:rotate(360deg)}}@keyframes rgtNoticeRight{from{opacity:0;transform:translate3d(18px,0,0)}to{opacity:1;transform:none}}@keyframes rgtNoticeLeft{from{opacity:0;transform:translate3d(-18px,0,0)}to{opacity:1;transform:none}}@keyframes rgtFade{from{opacity:0}to{opacity:1}}@media(max-width:640px){.rgt-notice-stack{top:8px;left:8px;width:min(300px,calc(100vw - 16px))}.rgt-notice{min-height:46px;padding:7px 8px;grid-template-columns:32px minmax(0,1fr) 22px}.rgt-notice-icon{width:32px;height:32px}.rgt-notice p{white-space:nowrap;overflow:hidden;text-overflow:ellipsis}}
    `}</style>
    {notice ? <div className="rgt-notice-stack" role="status" aria-live="polite"><div className={'rgt-notice ' + notice.type}><div className="rgt-notice-icon">{notice.type === 'success' ? '✓' : notice.type === 'warning' ? '!' : notice.type === 'error' ? '×' : ''}</div><div><strong>{notice.title}</strong><p>{notice.message}</p></div>{notice.persistent ? null : <button className="rgt-notice-close" onClick={() => setNotice(null)} aria-label="Close notification">×</button>}</div></div> : null}
    {tokenModal ? <div className="rgt-token-backdrop" role="dialog" aria-modal="true" aria-labelledby="rgt-token-title"><form className="rgt-token-modal" onSubmit={(event) => { event.preventDefault(); closeTokenModal(tokenModal.value); }}><p>SECURE ADMIN SYNC</p><h3 id="rgt-token-title">Publish module list</h3><span>Enter the admin token to publish this module list for all website users. You can continue locally without publishing.</span><input autoFocus type="password" value={tokenModal.value} onChange={(event) => setTokenModal((current) => ({ ...current, value: event.target.value }))} placeholder="Admin sync token" autoComplete="off"/><div className="rgt-token-actions"><button type="button" onClick={() => closeTokenModal('')}>Keep Local</button><button type="submit" disabled={!tokenModal.value.trim()}>Publish</button></div></form></div> : null}
    <Layout title="Reports and Status" sub="Project 4193 · HAMMER HEAD" back={back}><div className="work"><aside><header><p>PROJECT MODULES</p><h2>Module List</h2></header><div className="folder-connect"><button onClick={connectFolder}>Connect Excel Folder</button><span>{moduleSourceText}</span><input ref={inputRef} className="hidden-folder-input" type="file" accept=".xlsx,.xls,.xlsm,.xlsb" multiple webkitdirectory="" directory="" onChange={loadFolder}/></div><label>⌕<input value={query} onChange={(event) => setQuery(event.target.value)} placeholder="Search module name"/></label><section>{list.map((moduleName) => <button className={moduleName === selected ? 'sel' : ''} key={moduleName} onClick={() => setSelected((current) => current === moduleName ? null : moduleName)}>{moduleName}</button>)}</section></aside><article className="reports">{selected ? <><header><div><p>SELECTED MODULE</p><h2>{selected}</h2></div><div className="selected-module-actions"><button className="header-download-button" onClick={downloadModuleExcel}>Download {selected} Excel</button><button onClick={() => setSelected(null)}>Close</button></div></header><section className="selected-module-columns">{GROUPS.map(([groupName, items], groupIndex) => <div className="selected-module-column" key={groupName}><h3>{groupName}</h3><div>{items.map((item, itemIndex) => { const isReport = REPORTS.includes(item); const isSizeStatus = item === 'Size Status'; const count = isSizeStatus ? null : (busy ? '…' : counts[item] ?? 0); return <button className={`c${(groupIndex + itemIndex) % 5}`} key={item} onClick={() => item === 'SUPPORT TOOLS' ? openIntegrated(item) : isReport ? downloadReport(item) : openIntegrated(item)}><span>{item}</span>{count !== null ? <b>{count}</b> : null}</button>; })}</div></div>)}</section></> : <div className="empty"><b>RGT</b><p>MODULE SELECTION</p><h2>{busy ? 'Reading Excel files…' : modules.length ? 'Select a module' : 'Connect the master Excel folder'}</h2><span>{modules.length ? 'Select a module on the left to open reports here.' : 'The module list will be created automatically from the Module column.'}</span></div>}</article></div><Integrated open={openIntegrated}/></Layout>
  </>;
}

function SizeStatusPage({ moduleName, status, back }) {
  const plannedStorageKey = `rgt:planned-size-status:${normalize(moduleName)}`;
  const [planned, setPlanned] = useState(() => {
    const values = Object.fromEntries(SIZE_BUCKETS.map((bucket) => [bucket, 0]));
    try {
      const saved = JSON.parse(localStorage.getItem(plannedStorageKey) || '{}');
      SIZE_BUCKETS.forEach((bucket) => { values[bucket] = Math.max(0, Number(saved[bucket]) || 0); });
    } catch { /* keep zero values */ }
    return values;
  });
  const updatePlanned = (bucket, value) => {
    const nextValue = value === '' ? '' : Math.max(0, Math.floor(Number(value) || 0));
    setPlanned((current) => {
      const next = { ...current, [bucket]: nextValue };
      try { localStorage.setItem(plannedStorageKey, JSON.stringify(next)); } catch { /* ignore storage errors */ }
      return next;
    });
  };
  const plannedNumber = (bucket) => Math.max(0, Number(planned[bucket]) || 0);
  const plannedForExport = Object.fromEntries(SIZE_BUCKETS.map((bucket) => [bucket, plannedNumber(bucket)]));
  const calculatedNotModelled = Object.fromEntries(SIZE_BUCKETS.map((bucket) => [bucket, Math.max(0, plannedForExport[bucket] - status['MODELLED'][bucket])]));
  const modelledTotal = SIZE_BUCKETS.reduce((sum, bucket) => sum + status['MODELLED'][bucket], 0);
  const notModelledTotal = SIZE_BUCKETS.reduce((sum, bucket) => sum + calculatedNotModelled[bucket], 0);
  const plannedTotal = SIZE_BUCKETS.reduce((sum, bucket) => sum + plannedForExport[bucket], 0);
  const grandTotal = modelledTotal + notModelledTotal;
  return <main className="page size-status-page">
    <header className="top size-status-top">
      <div className="size-status-heading"><p>MODELLING STATUS</p><h1>{moduleName} Size Status</h1></div>
      <div className="size-status-actions"><button onClick={() => createSizeStatusWorkbook(moduleName, status, plannedForExport)}>Download Excel</button><button onClick={back}>Back</button></div>
    </header>
    <section className="content size-status-content"><div className="size-status-table-wrap"><table className="size-status-table">
      <thead><tr><th>{moduleName}-Status</th>{SIZE_BUCKETS.map((bucket) => <th key={bucket}>{bucket}</th>)}<th>Grand Total</th></tr></thead>
      <tbody>
        <tr><th>MODELLED</th>{SIZE_BUCKETS.map((bucket) => <td key={bucket}>{status['MODELLED'][bucket]}</td>)}<td>{modelledTotal}</td></tr>
        <tr><th>NOT MODELLED</th>{SIZE_BUCKETS.map((bucket) => <td key={bucket}>{calculatedNotModelled[bucket]}</td>)}<td>{notModelledTotal}</td></tr>
        <tr className="size-status-planned"><th>PLANNED</th>{SIZE_BUCKETS.map((bucket) => <td key={bucket}><input type="number" min="0" step="1" inputMode="numeric" value={planned[bucket]} onChange={(event) => updatePlanned(bucket, event.target.value)} onBlur={() => updatePlanned(bucket, plannedNumber(bucket))} aria-label={`Planned ${bucket}`} /></td>)}<td>{plannedTotal}</td></tr>
        <tr className="size-status-grand"><th>GRAND TOTAL</th>{SIZE_BUCKETS.map((bucket) => <td key={bucket}>{status['MODELLED'][bucket] + calculatedNotModelled[bucket]}</td>)}<td>{grandTotal}</td></tr>
      </tbody>
    </table></div></section>
  </main>;
}

function FluidCodePage({ moduleName, data, back }) {
  const grandTotal = data.fluidCodes.reduce((total, fluidCode) => total + data.sizes.reduce((sum, size) => sum + (data.counts[fluidCode]?.[size] || 0), 0), 0);
  return <main className="page size-status-page">
    <header className="top size-status-top">
      <div className="size-status-heading"><p>MODELLING STATUS</p><h1>{moduleName} Fluid Code</h1></div>
      <div className="size-status-actions"><button onClick={() => createFluidCodeWorkbook(moduleName, data)}>Download Excel</button><button onClick={back}>Back</button></div>
    </header>
    <section className="content size-status-content">
      <div className="size-status-table-wrap">
        <table className="size-status-table">
          <thead><tr><th>Fluid Code</th>{data.sizes.map((size) => <th key={size}>{formatPipeSize(size)}</th>)}<th>Grand Total</th></tr></thead>
          <tbody>
            {data.fluidCodes.map((fluidCode) => <tr key={fluidCode}><th>{fluidCode}</th>{data.sizes.map((size) => <td key={size}>{data.counts[fluidCode]?.[size] || 0}</td>)}<td>{data.sizes.reduce((sum, size) => sum + (data.counts[fluidCode]?.[size] || 0), 0)}</td></tr>)}
            <tr className="size-status-grand"><th>GRAND TOTAL</th>{data.sizes.map((size) => <td key={size}>{data.fluidCodes.reduce((sum, fluidCode) => sum + (data.counts[fluidCode]?.[size] || 0), 0)}</td>)}<td>{grandTotal}</td></tr>
          </tbody>
        </table>
      </div>
    </section>
  </main>;
}

function LineStatusPage({ moduleName, data, back }) {
  return <main className="page size-status-page">
    <header className="top size-status-top">
      <div className="size-status-heading"><p>GATE STATUS</p><h1>{moduleName} Line Status</h1></div>
      <div className="size-status-actions"><button onClick={() => createLineStatusWorkbook(moduleName, data)}>Download Excel</button><button onClick={back}>Back</button></div>
    </header>
    <section className="content size-status-content"><div className="size-status-table-wrap">
      <table className="size-status-table">
        <thead><tr><th>SIZE</th>{data.statusCodes.map((statusCode) => <th key={statusCode}>{statusCode}</th>)}<th>Grand Total</th></tr></thead>
        <tbody>
          {data.sizes.map((size) => <tr key={size}><th>{formatPipeSize(size)}</th>{data.statusCodes.map((statusCode) => <td key={statusCode}>{data.counts[size]?.[statusCode] || 0}</td>)}<td>{data.statusCodes.reduce((sum, statusCode) => sum + (data.counts[size]?.[statusCode] || 0), 0)}</td></tr>)}
          <tr className="size-status-grand"><th>GRAND TOTAL</th>{data.statusCodes.map((statusCode) => <td key={statusCode}>{data.sizes.reduce((sum, size) => sum + (data.counts[size]?.[statusCode] || 0), 0)}</td>)}<td>{data.grandTotal}</td></tr>
        </tbody>
      </table>
      {!data.sizes.length ? <div className="empty"><p>LINE STATUS</p><h2>No size-wise Status Code data found</h2><span>Connect an Excel file starting with PIPE_E3D_STATUS-EHH_ that contains PIPE and STATUS CODE columns.</span></div> : null}
    </div></section>
  </main>;
}

function SupportToolsPage({ moduleName, open, back }) {
  const spanTools = ['≤2"', '3"-4"', '≥6"'];
  const loadTools = [
    'SUPPORT VALIDATION',
    'SUPPORT TYPE',
    'LOCAL STRESS/LOAD FOR PRIMARY SUPPORT',
    'PRIMARY SUPPORT SELECTION',
    'SECONDARY SUPPORT VALIDATION AND SELECTION',
  ];
  const openBlank = (toolName) => open('SUPPORT TOOLS · ' + moduleName + ' · ' + toolName);
  return <main className="page support-tools-page">
    <style>{`
      .support-tools-top{background:linear-gradient(135deg,#102f4a,#194f73)}
      .support-tools-content{padding:32px!important;background:#f3f8fc}
      .support-tools-shell{max-width:1100px;margin:0 auto}
      .support-tools-intro{margin-bottom:22px}.support-tools-intro p{margin:0;color:#1686bc;font-size:12px;font-weight:800;letter-spacing:.14em}.support-tools-intro h2{margin:6px 0 0;color:#17324d;font-size:26px}
      .support-tools-grid{display:grid;grid-template-columns:1fr 1.25fr;gap:22px}
      .support-tools-panel{padding:24px;border:1px solid #d7e5f0;border-radius:24px;background:#dceaf5;box-shadow:0 14px 34px rgba(23,50,77,.1)}
      .support-tools-panel h3{margin:0 0 20px;text-align:center;color:#172b3c;font-size:23px;font-weight:700}
      .support-tools-buttons{display:grid;gap:10px}.support-tools-buttons button{min-height:58px;padding:12px 16px;border:2px solid rgba(255,255,255,.92);border-radius:9px;color:#fff;font-size:14px;font-weight:800;box-shadow:0 5px 12px rgba(23,50,77,.14);cursor:pointer;transition:transform .18s ease,box-shadow .18s ease,filter .18s ease}.support-tools-buttons button:hover{transform:translateY(-2px);box-shadow:0 9px 18px rgba(23,50,77,.2);filter:brightness(1.04)}.support-tools-buttons button:active{transform:translateY(0)}
      .support-span-buttons button{min-height:88px}.support-span-buttons button:nth-child(1){background:#2498cb}.support-span-buttons button:nth-child(2){background:#704b9c}.support-span-buttons button:nth-child(3){background:#ef615d}
      .support-load-buttons button:nth-child(1){background:#167fb4}.support-load-buttons button:nth-child(2){background:#126f9f}.support-load-buttons button:nth-child(3){background:#1b8fc4}.support-load-buttons button:nth-child(4){background:#197caf}.support-load-buttons button:nth-child(5){background:#155f8e}
      @media(max-width:760px){.support-tools-content{padding:18px!important}.support-tools-grid{grid-template-columns:1fr}.support-tools-panel{padding:18px;border-radius:18px}.support-tools-panel h3{font-size:20px}}
    `}</style>
    <header className="top support-tools-top">
      <button onClick={back}>Back</button>
      <div><p>REPORTS GENERATE TOOLS</p><h1>Support Tools</h1><span>{moduleName}</span></div>
    </header>
    <section className="content support-tools-content">
      <div className="support-tools-shell">
        <div className="support-tools-intro"><p>SUPPORT ENGINEERING</p><h2>Select a Support Tool</h2></div>
        <div className="support-tools-grid">
          <article className="support-tools-panel"><h3>Support Span Check</h3><div className="support-tools-buttons support-span-buttons">{spanTools.map((tool) => <button key={tool} onClick={() => openBlank(tool)}>{tool}</button>)}</div></article>
          <article className="support-tools-panel"><h3>Support Load Validation</h3><div className="support-tools-buttons support-load-buttons">{loadTools.map((tool) => <button key={tool} onClick={() => openBlank(tool)}>{tool}</button>)}</div></article>
        </div>
      </div>
    </section>
  </main>;
}

function IsoPlanningManagementPage({ workbooks, back, open }) {
  const [query, setQuery] = useState('');
  const modules = useMemo(() => getIsoPlanningModules(workbooks)
    .sort((a, b) => a.localeCompare(b, undefined, { numeric: true, sensitivity: 'base' })), [workbooks]);
  const filteredModules = useMemo(() => modules.filter((moduleName) => moduleName.toLowerCase().includes(query.trim().toLowerCase())), [modules, query]);
  return <Layout title="ISO Planning and Management" sub="Project 4193 · HAMMER HEAD" back={back}>
    <div className="work">
      <aside>
        <header><p>PROJECT MODULES</p><h2>Module List</h2></header>
        <label>⌕<input value={query} onChange={(event) => setQuery(event.target.value)} placeholder="Search module name"/></label>
        <section>{filteredModules.map((moduleName) => <button key={moduleName} onClick={() => window.scrollTo(0, 0)}>{moduleName}</button>)}</section>
      </aside>
      <article className="reports">
        <div className="empty"><b>RGT</b><p>ISO PLANNING AND MANAGEMENT</p><h2>{modules.length ? 'Select a module' : 'ISO module data is not loaded'}</h2><span>{modules.length ? 'Select a module on the left to continue.' : 'Connect the Excel folder once from Reports and Status. ISO modules are read only from 4193_HH_Isometrics Control Sheet-Piping_OFSI.'}</span></div>
      </article>
    </div>
  </Layout>;
}

function IsoPlanningDashboardPage({ workbooks, back }) {
  const [selectedModule, setSelectedModule] = useState(null);
  const [query, setQuery] = useState('');
  const [consoleMode, setConsoleMode] = useState(false);
  const [planningMode, setPlanningMode] = useState(false);
  const [blankToolPage, setBlankToolPage] = useState(null);
  const metricsRef = useRef(null);
  const progressRef = useRef(null);
  const modules = useMemo(() => buildIsoPlanningData(workbooks), [workbooks]);
  const visibleModules = useMemo(() => modules.filter((item) => item.moduleName.toLowerCase().includes(query.trim().toLowerCase())), [modules, query]);
  const moduleSelection = modules.find((item) => normalize(item.moduleName) === normalize(selectedModule)) || null;
  const projectConsole = useMemo(() => combineIsoModules(modules), [modules]);
  const selected = planningMode ? projectConsole : (consoleMode ? projectConsole : moduleSelection);
  const total = isoMetricRows(selected, 'total').length;
  const issued = isoMetricRows(selected, 'issued').length;
  const revision = isoMetricRows(selected, 'revision').length;
  const hold = isoMetricRows(selected, 'hold').length;
  const balance = Math.max(0, total - issued - hold);
  const issuedPercent = total ? Math.round((issued / total) * 100) : 0;
  const revisionRatio = revision ? (total / revision).toFixed(2) : '0';
  const stageSummary = isoStageSummary(selected);
  const planSelection = planningMode ? projectConsole : moduleSelection;
  const forecastMap = useMemo(() => buildIsoForecastMap(workbooks), [workbooks]);
  const fourWeekPlan = useMemo(() => buildIsoFourWeekPlan(planSelection), [planSelection]);
  const projectMonthPlan = useMemo(() => {
    const modulePlans = modules.map((moduleData) => buildIsoLatestFourMonthPlan(moduleData, forecastMap));
    return {
      modules: modulePlans,
      maxValue: Math.max(1, ...modulePlans.flatMap((item) => item.months.map((month) => month.total))),
      grandTotal: modulePlans.reduce((sum, item) => sum + item.total, 0),
      grandHolds: modulePlans.reduce((sum, item) => sum + item.holds, 0),
    };
  }, [modules, forecastMap]);
  const chooseModule = (moduleName) => { setSelectedModule(moduleName); setConsoleMode(false);  setPlanningMode(false); };
  const openPlanning = () => {
    setPlanningMode(true);
    setConsoleMode(false);
    setSelectedModule(null);
    setBlankToolPage(null);
    window.setTimeout(() => metricsRef.current?.scrollIntoView({ behavior: 'smooth', block: 'start' }), 0);
  };
  const openConsole = () => {
    setPlanningMode(false);
    setBlankToolPage(null);
    setConsoleMode(true);
    setSelectedModule(null);
    window.setTimeout(() => metricsRef.current?.scrollIntoView({ behavior: 'smooth', block: 'start' }), 0);
  };
  const openAnalytics = () => setBlankToolPage('ANALYTICS');
  const openRevision = () => setBlankToolPage('REVISION');
  if (blankToolPage) {
    return <Layout title={blankToolPage} sub="ISO Planning and Management · Project 4193" back={() => setBlankToolPage(null)}>
      <style>{`
        .iso-tool-blank{min-height:560px;display:grid;place-items:center;padding:32px;background:linear-gradient(180deg,#f5faff,#ebf3f9)}.iso-tool-blank-card{width:min(720px,100%);padding:70px 30px;border:1px solid #d8e6f0;border-radius:24px;background:#fff;box-shadow:0 20px 55px rgba(16,47,74,.12);text-align:center}.iso-tool-blank-mark{width:72px;height:72px;margin:0 auto 20px;display:grid;place-items:center;border-radius:20px;background:linear-gradient(135deg,#173f61,#2498cb);color:#fff;font-size:22px;font-weight:900;letter-spacing:.08em}.iso-tool-blank-card p{margin:0;color:#1686bc;font-size:12px;font-weight:900;letter-spacing:.16em}.iso-tool-blank-card h2{margin:10px 0 8px;color:#122943;font-size:32px}.iso-tool-blank-card span{color:#70869a;font-size:14px}
      `}</style>
      <section className="iso-tool-blank"><div className="iso-tool-blank-card"><div className="iso-tool-blank-mark">RGT</div><p>ISO PLANNING AND MANAGEMENT</p><h2>{blankToolPage}</h2><span>Content will be added later.</span></div></section>
    </Layout>;
  }
  return <Layout title="ISO Planning and Management" sub="Project 4193 · HAMMER HEAD" back={back}>
    <style>{`
      .iso-plan-layout{display:grid;grid-template-columns:300px minmax(0,1fr);gap:24px}.iso-plan-modules,.iso-console{border:1px solid #d9e7f1;background:linear-gradient(180deg,#fff,#fafdff);box-shadow:0 18px 45px rgba(16,47,74,.11)}.iso-plan-modules{padding:20px;border-radius:20px}.iso-plan-side-head{display:flex;align-items:center;justify-content:space-between;gap:10px}.iso-top-actions{display:flex;justify-content:flex-end;align-items:center;gap:14px;margin:0 0 34px;flex-wrap:wrap}.iso-top-action{min-width:142px;padding:14px 20px;border:1px solid rgba(255,255,255,.8);border-radius:13px;background:linear-gradient(135deg,#2498cb,#176f9f);color:#fff;font-size:13px;font-weight:900;letter-spacing:.06em;cursor:pointer;box-shadow:0 9px 22px rgba(23,117,170,.22);transition:transform .18s ease,box-shadow .18s ease,filter .18s ease}.iso-top-action:hover{transform:translateY(-2px);box-shadow:0 13px 27px rgba(23,117,170,.3);filter:brightness(1.04)}.iso-top-action.active{background:linear-gradient(135deg,#173f61,#0f2f49);box-shadow:0 11px 25px rgba(15,47,73,.28)}.iso-top-action:disabled{opacity:.45;cursor:not-allowed;transform:none;box-shadow:none}.iso-module-plan-below-progress{margin-top:42px}.iso-module-plan-below-progress:before{content:'MODULE-WISE PLANNING';display:block;margin:0 0 14px;color:#7dd3fc;font-size:10px;font-weight:900;letter-spacing:.16em}.iso-plan-chart-card{padding:26px;border:1px solid #d3e3ef;border-radius:20px;background:linear-gradient(145deg,#102f4a,#173f61 58%,#1b5478);box-shadow:0 20px 46px rgba(15,47,73,.24);color:#fff}.iso-plan-chart-head{display:flex;justify-content:space-between;align-items:flex-end;gap:20px;margin-bottom:24px}.iso-plan-chart-head p{margin:0;color:#7dd3fc;font-size:11px;font-weight:900;letter-spacing:.15em}.iso-plan-chart-head h2{margin:5px 0 0;font-size:28px}.iso-plan-chart-source{padding:9px 13px;border:1px solid rgba(255,255,255,.18);border-radius:10px;background:rgba(255,255,255,.08);font-size:11px;font-weight:900}.iso-plan-week-grid{display:grid;grid-template-columns:repeat(4,minmax(180px,1fr));gap:15px}.iso-plan-week{padding:17px;border:1px solid rgba(255,255,255,.14);border-radius:16px;background:rgba(255,255,255,.075);backdrop-filter:blur(7px)}.iso-plan-week header{display:flex;justify-content:space-between;gap:10px;align-items:flex-start}.iso-plan-week header b{display:block;font-size:18px}.iso-plan-week header span{display:block;margin-top:3px;color:#b9d3e5;font-size:10px}.iso-plan-week header strong{color:#67e8f9;font-size:12px}.iso-plan-bars{height:170px;display:grid;grid-template-columns:repeat(5,1fr);gap:6px;align-items:end;margin-top:18px;padding-top:22px;border-bottom:1px solid rgba(255,255,255,.2)}.iso-plan-day{height:100%;display:flex;flex-direction:column;justify-content:flex-end;align-items:center}.iso-plan-bar-track{width:100%;height:135px;display:flex;align-items:flex-end;justify-content:center}.iso-plan-bar{position:relative;width:70%;min-height:2px;border-radius:7px 7px 2px 2px;background:linear-gradient(180deg,#22d3ee,#0284c7);box-shadow:0 4px 14px rgba(34,211,238,.28);transition:height .35s ease}.iso-plan-bar span{position:absolute;top:-19px;left:50%;transform:translateX(-50%);font-size:10px;font-weight:900}.iso-plan-day small{margin:7px 0 5px;color:#d9ebf5;font-size:9px;font-weight:800;text-align:center;line-height:1.25}.iso-plan-day small em{display:block;margin-top:3px;color:#7dd3fc;font-size:8px;font-style:normal;font-weight:800;white-space:nowrap}.iso-plan-week footer{display:flex;justify-content:space-between;gap:8px;margin-top:13px;padding-top:12px;border-top:1px solid rgba(255,255,255,.12);font-size:10px}.iso-plan-week footer b{font-size:15px}.iso-plan-hold{color:#fda4af}.iso-plan-chart-note{display:flex;justify-content:space-between;gap:15px;margin-top:18px;color:#b9d3e5;font-size:10px}@media(max-width:1250px){.iso-plan-week-grid{grid-template-columns:repeat(2,minmax(180px,1fr))}}@media(max-width:650px){.iso-plan-week-grid{grid-template-columns:1fr}.iso-plan-chart-head,.iso-plan-chart-note{align-items:flex-start;flex-direction:column}.iso-plan-chart-card{padding:18px}}.iso-plan-modules input{box-sizing:border-box;width:100%;margin-top:12px;padding:11px 13px;border:1px solid #bfd2e1;border-radius:10px;outline:none}.iso-plan-modules section{display:grid;gap:8px;margin-top:14px;max-height:640px;overflow:auto}.iso-plan-modules section button,.iso-metric{border:0;border-radius:11px;background:linear-gradient(135deg,#2498cb,#1875aa);color:#fff;font-weight:800;cursor:pointer}.iso-plan-modules section button{padding:13px;text-align:left}.iso-plan-modules section button.sel{background:linear-gradient(135deg,#f28a3b,#df642a)}.iso-console{padding:25px;border-radius:22px;overflow:hidden}.iso-console-head{display:flex;align-items:center;justify-content:space-between;gap:28px;flex-wrap:wrap}.iso-console-eyebrow{margin:0;color:#1686bc;font-size:12px;font-weight:900;letter-spacing:.14em}.iso-console-title{margin:6px 0 0;color:#122943;font-size:30px}.iso-progress{width:160px;height:160px;border-radius:50%;display:grid;place-items:center;background:conic-gradient(#09be68 calc(var(--p)*1%),#dceaf5 0);position:relative;box-shadow:0 10px 28px rgba(13,157,91,.18)}.iso-progress:after{content:'';position:absolute;inset:16px;border-radius:50%;background:#fff}.iso-progress-content{position:relative;z-index:1;text-align:center}.iso-progress-content b{display:block;font-size:30px;color:#112944}.iso-progress-content span{font-size:12px;color:#35516b}.iso-metrics{display:grid;grid-template-columns:repeat(5,minmax(120px,1fr));gap:12px;margin-top:25px}.iso-metric{padding:15px;text-align:left;min-height:94px}.iso-metric span{display:block;font-size:12px;letter-spacing:.04em}.iso-metric b{display:block;margin-top:8px;font-size:28px}.iso-stage-card{margin-top:52px;border:1px solid #d4e3ee;border-radius:17px;overflow:auto;background:#fff}.iso-stage-heading{padding:11px 16px;background:#173f61;color:#fff;font-size:13px;font-weight:900;letter-spacing:.12em;text-align:center}.iso-stage-table{width:100%;min-width:980px;border-collapse:collapse}.iso-stage-table tbody tr+tr th,.iso-stage-table tbody tr+tr td{border-top:18px solid #111827}.iso-stage-table tbody tr+tr th:first-child{border-top-left-radius:8px}.iso-stage-table tbody tr+tr td:last-child{border-top-right-radius:8px}.iso-stage-table th,.iso-stage-table td{padding:10px 8px;border:1px solid #d4e3ee;text-align:center}.iso-stage-table th{background:#dceaf5;color:#163651;font-size:11px}.iso-stage-row-label{min-width:150px!important;background:#f3f8fc!important;font-weight:900!important;color:#173f61!important}.iso-stage-count{border:0;background:transparent;color:#126e9d;font-size:16px;font-weight:900;cursor:pointer}.iso-stage-count:hover{text-decoration:underline}.iso-empty{padding:80px 20px;text-align:center;color:#64748b}.iso-week-overview{margin-top:0}.iso-week-summary{display:flex;gap:12px;flex-wrap:wrap;justify-content:flex-end}.iso-week-summary span{padding:9px 13px;border:1px solid rgba(255,255,255,.16);border-radius:10px;background:rgba(255,255,255,.08);font-size:11px;font-weight:900}.iso-module-week-grid{display:grid;grid-template-columns:repeat(auto-fit,minmax(250px,1fr));gap:18px}.iso-module-week-card{padding:18px;border:1px solid rgba(255,255,255,.13);border-radius:16px;background:linear-gradient(180deg,rgba(255,255,255,.10),rgba(255,255,255,.045));box-shadow:inset 0 1px rgba(255,255,255,.08),0 12px 25px rgba(0,0,0,.13)}.iso-module-week-head{display:flex;align-items:center;justify-content:space-between;gap:14px;margin-bottom:16px}.iso-module-week-head b{font-size:19px;letter-spacing:.03em}.iso-module-week-head strong{color:#f8cf58;font-size:16px}.iso-module-week-bars{height:190px;display:grid;grid-template-columns:repeat(5,minmax(30px,1fr));align-items:end;gap:10px;padding:18px 10px 8px;border-radius:13px;background:linear-gradient(180deg,rgba(4,22,37,.34),rgba(4,22,37,.64));border:1px solid rgba(255,255,255,.08)}.iso-module-week-bar{height:100%;display:flex;flex-direction:column;align-items:center;justify-content:flex-end;gap:7px}.iso-module-week-bar i{position:relative;width:100%;max-width:34px;min-height:3px;border-radius:7px 7px 2px 2px;background:linear-gradient(180deg,#43b4e8,#1486bd);box-shadow:0 0 15px rgba(67,180,232,.32);transition:height .3s ease}.iso-module-week-bar.total i{background:linear-gradient(180deg,#ffd85d,#dca817);box-shadow:0 0 16px rgba(255,216,93,.38)}.iso-module-week-bar i em{position:absolute;left:50%;top:-22px;transform:translateX(-50%);font-style:normal;color:#fff;font-size:12px;font-weight:900}.iso-module-week-bar span{color:#cae1ef;font-size:11px;font-weight:900}.iso-module-week-foot{display:flex;justify-content:space-between;gap:12px;margin-top:13px;color:#c7dceb;font-size:11px}.iso-single-week-grid{display:grid;grid-template-columns:repeat(4,minmax(120px,1fr));gap:16px}.iso-single-week{padding:18px;border:1px solid rgba(255,255,255,.13);border-radius:15px;background:rgba(255,255,255,.07);text-align:center}.iso-single-week b{display:block;color:#82d9ff;font-size:14px}.iso-single-week strong{display:block;margin:10px 0 4px;font-size:34px}.iso-single-week span{color:#bfd6e5;font-size:11px}.iso-single-total{margin-top:18px;display:flex;justify-content:space-between;align-items:center;padding:16px 18px;border-radius:14px;background:linear-gradient(90deg,rgba(255,211,76,.18),rgba(255,255,255,.06));border:1px solid rgba(255,211,76,.28)}.iso-single-total strong{color:#ffd85d;font-size:25px}.iso-premium-module-plan{position:relative;overflow:hidden;background:radial-gradient(circle at 88% 8%,rgba(39,174,237,.20),transparent 30%),linear-gradient(145deg,#0b263d,#123f61 58%,#17658b);border-color:rgba(104,207,255,.24);box-shadow:0 28px 70px rgba(7,33,53,.30)}.iso-premium-module-plan:after{content:'';position:absolute;inset:0;pointer-events:none;background:linear-gradient(120deg,rgba(255,255,255,.06),transparent 28%)}.iso-premium-kpis{position:relative;z-index:1;display:grid;grid-template-columns:repeat(2,minmax(0,1fr));gap:14px;margin-bottom:22px}.iso-premium-kpis article{padding:15px 18px;border:1px solid rgba(255,255,255,.13);border-radius:15px;background:rgba(255,255,255,.075);box-shadow:inset 0 1px rgba(255,255,255,.08)}.iso-premium-kpis span{display:block;color:#92dffd;font-size:10px;font-weight:900;letter-spacing:.14em}.iso-premium-kpis strong{display:block;margin:5px 0 2px;color:#fff;font-size:30px;line-height:1}.iso-premium-kpis small{color:#b8d0e0;font-size:10px}.iso-premium-chart{position:relative;z-index:1;height:360px;padding:30px 24px 28px;border:1px solid rgba(255,255,255,.12);border-radius:18px;background:linear-gradient(180deg,rgba(1,15,27,.28),rgba(1,15,27,.58));overflow:hidden}.iso-premium-grid-lines{position:absolute;inset:30px 24px 52px;display:flex;flex-direction:column;justify-content:space-between}.iso-premium-grid-lines i{display:block;border-top:1px dashed rgba(185,219,239,.15)}.iso-premium-bars{position:relative;z-index:1;height:100%;display:grid;grid-template-columns:repeat(4,minmax(0,1fr));gap:28px;align-items:end;padding-bottom:34px}.iso-premium-week{position:relative;height:100%;display:flex;flex-direction:column;align-items:center;justify-content:flex-end;padding-bottom:68px}.iso-premium-value{margin-bottom:8px;color:#eaf8ff;font-size:12px;font-weight:900}.iso-premium-track{position:relative;width:min(72px,68%);height:220px;display:flex;align-items:flex-end;border-radius:12px 12px 4px 4px;background:rgba(255,255,255,.045);box-shadow:inset 0 0 0 1px rgba(255,255,255,.07)}.iso-premium-fill{position:relative;width:100%;min-height:4px;border-radius:12px 12px 4px 4px;transition:height .35s ease;box-shadow:0 0 24px rgba(38,169,224,.34)}.iso-premium-fill:after{content:'';position:absolute;inset:0;border-radius:inherit;background:linear-gradient(90deg,rgba(255,255,255,.28),transparent 45%)}.iso-premium-fill.w1{background:linear-gradient(180deg,#61d5ff,#168bc1)}.iso-premium-fill.w2{background:linear-gradient(180deg,#9c8cff,#5748bd)}.iso-premium-fill.w3{background:linear-gradient(180deg,#ff8c73,#db4f3f)}.iso-premium-fill.w4{background:linear-gradient(180deg,#31d5ba,#0b8f80)}.iso-premium-fill span{position:absolute;left:50%;top:10px;transform:translateX(-50%);color:#fff;font-size:14px;font-weight:900}.iso-premium-label{position:absolute;left:0;right:0;bottom:-8px;min-height:48px;text-align:center}.iso-premium-label b{display:block;color:#fff;font-size:15px;line-height:1.1}.iso-premium-label em{display:block;margin-top:3px;color:#b7d7e9;font-size:9px;font-style:normal;line-height:1.1;white-space:nowrap}.iso-premium-label small{display:block;margin-top:4px;color:#89b7d2;font-size:9px;line-height:1.1;letter-spacing:.09em;white-space:nowrap}.iso-premium-table{position:relative;z-index:1;margin-top:18px;border:1px solid rgba(255,255,255,.13);border-radius:13px;overflow:hidden}.iso-premium-table>div{display:grid;grid-template-columns:1.5fr repeat(4,1fr) 1.3fr;align-items:center}.iso-premium-table .head{background:rgba(255,255,255,.09)}.iso-premium-table .body{background:rgba(0,0,0,.13)}.iso-premium-table span,.iso-premium-table b,.iso-premium-table strong{padding:11px 10px;border-right:1px solid rgba(255,255,255,.09);text-align:center;font-size:11px}.iso-premium-table span{text-align:left;color:#a8d9f3;font-weight:900}.iso-premium-table strong{color:#ffd85d}.iso-premium-table b{color:#fff}.iso-month-plan{position:relative;overflow:hidden;background:radial-gradient(circle at 90% 5%,rgba(54,184,238,.20),transparent 32%),linear-gradient(145deg,#0b263d,#123f61 58%,#17658b);border-color:rgba(104,207,255,.24);box-shadow:0 28px 70px rgba(7,33,53,.30)}.iso-month-plan-summary{display:flex;gap:12px;flex-wrap:wrap;justify-content:flex-end}.iso-month-plan-summary span{padding:9px 13px;border:1px solid rgba(255,255,255,.16);border-radius:10px;background:rgba(255,255,255,.08);font-size:11px;font-weight:900}.iso-month-legend{display:flex;gap:16px;flex-wrap:wrap;margin:-5px 0 22px;color:#c4deed;font-size:10px;font-weight:900}.iso-month-legend span{display:flex;align-items:center;gap:6px}.iso-month-legend i{width:10px;height:10px;border-radius:3px}.iso-month-legend .m1{background:#4dc8f2}.iso-month-legend .m2{background:#8b78ea}.iso-month-legend .m3{background:#f27665}.iso-month-legend .m4{background:#25bca5}.iso-month-chart-scroll{overflow-x:auto;padding-bottom:8px}.iso-month-chart{position:relative;min-width:max(920px,calc(var(--module-count) * 150px));height:420px;padding:30px 24px 70px;border:1px solid rgba(255,255,255,.12);border-radius:18px;background:linear-gradient(180deg,rgba(1,15,27,.26),rgba(1,15,27,.60))}.iso-month-grid-lines{position:absolute;inset:30px 24px 92px;display:flex;flex-direction:column;justify-content:space-between}.iso-month-grid-lines i{border-top:1px dashed rgba(185,219,239,.15)}.iso-month-module-groups{position:relative;z-index:1;height:100%;display:grid;grid-template-columns:repeat(var(--module-count),minmax(120px,1fr));gap:22px;align-items:end}.iso-month-module{position:relative;height:100%;display:flex;flex-direction:column;justify-content:flex-end;align-items:center;padding-bottom:58px}.iso-month-bars{width:100%;height:270px;display:grid;grid-template-columns:repeat(4,minmax(18px,1fr));align-items:end;gap:8px}.iso-month-bar{position:relative;height:100%;display:flex;align-items:flex-end}.iso-month-fill{position:relative;width:100%;min-height:4px;border-radius:8px 8px 3px 3px;box-shadow:0 0 20px rgba(38,169,224,.28)}.iso-month-fill.m1{background:linear-gradient(180deg,#70dcff,#168bc1)}.iso-month-fill.m2{background:linear-gradient(180deg,#a99aff,#5748bd)}.iso-month-fill.m3{background:linear-gradient(180deg,#ff967f,#db4f3f)}.iso-month-fill.m4{background:linear-gradient(180deg,#49dbc3,#0b8f80)}.iso-month-fill b{position:absolute;left:50%;top:-22px;transform:translateX(-50%);color:#fff;font-size:12px}.iso-month-module-label{position:absolute;left:0;right:0;bottom:0;text-align:center}.iso-month-module-label strong{display:block;color:#fff;font-size:14px}.iso-month-module-label span{display:block;margin-top:4px;color:#ffd85d;font-size:11px;font-weight:900}.iso-month-matrix{margin-top:20px;border:1px solid rgba(255,255,255,.13);border-radius:14px;overflow:auto}.iso-month-matrix>div{display:grid;grid-template-columns:140px repeat(var(--module-count),minmax(94px,1fr)) 110px;min-width:max(900px,calc(var(--module-count) * 105px + 250px))}.iso-month-matrix span,.iso-month-matrix b,.iso-month-matrix strong{padding:11px 9px;border-right:1px solid rgba(255,255,255,.09);border-bottom:1px solid rgba(255,255,255,.09);text-align:center;font-size:10px}.iso-month-matrix .head{background:rgba(255,255,255,.10)}.iso-month-matrix .row{background:rgba(0,0,0,.12)}.iso-month-matrix span{text-align:left;color:#9fd8f5;font-weight:900}.iso-month-matrix strong{color:#ffd85d}.iso-month-cell small{display:block;margin-top:3px;color:#91b7cc;font-size:8px;font-weight:700}.iso-month-empty{padding:44px 20px;text-align:center;color:#c4deed}@media(max-width:1050px){.iso-plan-layout{grid-template-columns:1fr}.iso-metrics{grid-template-columns:repeat(2,1fr)}}@media(max-width:1050px){.iso-top-actions{justify-content:flex-start}.iso-top-action{min-width:124px}}@media(max-width:640px){.iso-top-actions{display:grid;grid-template-columns:repeat(2,minmax(0,1fr));gap:10px}.iso-top-action{width:100%;min-width:0;padding:12px 10px}.iso-single-week-grid{grid-template-columns:repeat(2,minmax(0,1fr))}.iso-module-week-grid{grid-template-columns:1fr}.iso-plan-chart-head{align-items:flex-start;flex-direction:column}.iso-week-summary{justify-content:flex-start}.iso-premium-kpis{grid-template-columns:1fr 1fr}.iso-premium-chart{height:315px;padding:24px 10px 24px}.iso-premium-bars{gap:9px}.iso-premium-track{width:70%;height:180px}.iso-premium-table{overflow-x:auto}.iso-premium-table>div{min-width:620px}}
    `}</style>
    <div className="iso-plan-layout">
      <aside className="iso-plan-modules"><div className="iso-plan-side-head"><h2>Module List</h2></div><input value={query} onChange={(event) => setQuery(event.target.value)} placeholder="Search module name"/><section>{visibleModules.map((item) => <button className={!consoleMode && selectedModule === item.moduleName ? 'sel' : ''} key={item.moduleName} onClick={() => chooseModule(item.moduleName)}>{item.moduleName}</button>)}</section></aside>
      <section className="iso-console" ref={metricsRef}>
      <div className="iso-top-actions">
        <button className={'iso-top-action ' + (planningMode ? 'active' : '')} onClick={openPlanning}>PLANNING</button>
        <button className={'iso-top-action ' + (!planningMode && consoleMode ? 'active' : '')} onClick={openConsole}>CONSOLE</button>
        <button className="iso-top-action" onClick={openAnalytics}>ANALYTICS</button>
        <button className="iso-top-action" onClick={openRevision}>REVISION</button>
      </div>{selected && selected.entries.length ? <>{!planningMode ? <><div className="iso-console-head"><div><p className="iso-console-eyebrow">{consoleMode ? 'PROJECT CONSOLE · ALL MODULES' : 'SELECTED MODULE'}</p><h2 className="iso-console-title">{selected.moduleName}</h2></div><div className="iso-progress" style={{ '--p': issuedPercent }}><div className="iso-progress-content"><b>{issuedPercent}%</b><span>Revision ratio {revisionRatio}</span></div></div></div><div className="iso-metrics"><button className="iso-metric" onClick={() => downloadIsoMetric(selected,'total','Total Isos')}><span>TOTAL ISOS</span><b>{total}</b></button><button className="iso-metric" onClick={() => downloadIsoMetric(selected,'issued','Issued')}><span>ISSUED</span><b>{issued}</b></button><button className="iso-metric" onClick={() => downloadIsoMetric(selected,'revision','Revision')}><span>REVISION</span><b>{revision}</b></button><button className="iso-metric" onClick={() => downloadIsoMetric(selected,'hold','Hold')}><span>HOLD</span><b>{hold}</b></button><button className="iso-metric" onClick={() => downloadIsoMetric(selected,'balance','Balance Hold Free')}><span>BALANCE HOLD FREE</span><b>{balance}</b></button></div><div className="iso-stage-card" ref={progressRef}><div className="iso-stage-heading">ISO PROGRESS CONTROL</div><table className="iso-stage-table"><thead><tr><th>STATUS</th>{ISO_STAGE_DEFINITIONS.map((stage) => <th key={stage.key}>{stage.label}</th>)}</tr></thead><tbody><tr><th className="iso-stage-row-label">COMPLETED</th>{ISO_STAGE_DEFINITIONS.map((stage) => <td key={stage.key}><button className="iso-stage-count" onClick={() => downloadIsoStage(selected, stage, true)}>{stageSummary[stage.key]?.completed || 0}</button></td>)}</tr><tr><th className="iso-stage-row-label">WORK FRONT AVAILABLE</th>{ISO_STAGE_DEFINITIONS.map((stage) => <td key={stage.key}><button className="iso-stage-count" onClick={() => downloadIsoStage(selected, stage, false)}>{stageSummary[stage.key]?.available || 0}</button></td>)}</tr></tbody></table></div></> : null}{planningMode ? <section className="iso-plan-chart-card iso-month-plan" style={{'--module-count':Math.max(1,projectMonthPlan.modules.length)}}>
        <div className="iso-plan-chart-head"><div><p>ALL MODULES · LATEST 4 MONTHS AUTOMATIC FORECAST</p><h2>4 MONTHS PLAN</h2></div><div className="iso-month-plan-summary"><span>NO. OF ISOS <b>{projectMonthPlan.grandTotal}</b></span><span>HOLDS <b>{projectMonthPlan.grandHolds}</b></span></div></div>
        <div className="iso-month-legend"><span><i className="m1"/>MONTH 1</span><span><i className="m2"/>MONTH 2</span><span><i className="m3"/>MONTH 3</span><span><i className="m4"/>MONTH 4</span></div>
        {projectMonthPlan.modules.length ? <><div className="iso-month-chart-scroll"><div className="iso-month-chart"><div className="iso-month-grid-lines"><i/><i/><i/><i/><i/></div><div className="iso-month-module-groups">{projectMonthPlan.modules.map((item)=><article className="iso-month-module" key={item.moduleName}><div className="iso-month-bars">{item.months.map((month,index)=><div className="iso-month-bar" key={month.key}><div className={'iso-month-fill m'+(index+1)} style={{height:Math.max(month.total?10:3,(month.total/projectMonthPlan.maxValue)*100)+'%'}}><b>{month.total}</b></div></div>)}</div><div className="iso-month-module-label"><strong>{item.moduleName}</strong><span>{item.total} ISOs</span></div></article>)}</div></div></div>
        <div className="iso-month-matrix"><div className="head"><span>MONTH / MODULE</span>{projectMonthPlan.modules.map((item)=><b key={item.moduleName}>{item.moduleName}</b>)}<strong>TOTAL</strong></div>{[0,1,2,3].map((monthIndex)=><div className="row" key={monthIndex}><span>MONTH {monthIndex+1}</span>{projectMonthPlan.modules.map((item)=><b className="iso-month-cell" key={item.moduleName}>{item.months[monthIndex].total}<small>{item.months[monthIndex].label}</small></b>)}<strong>{projectMonthPlan.modules.reduce((sum,item)=>sum+item.months[monthIndex].total,0)}</strong></div>)}</div></> : <div className="iso-month-empty">No module forecast data found.</div>}
        <div className="iso-plan-chart-note"><span>Source sheet: HH_TOPSIDE_OVERALL_CS_OFSI · each module uses its own latest four distinct months from FORECAST ISSUE DATE.</span><span>Blank rows, LINE DELETED and struck-through rows are excluded.</span></div>
      </section> : planSelection && planSelection.entries.length ? <section className="iso-plan-chart-card iso-module-plan-below-progress iso-premium-module-plan">
        <div className="iso-plan-chart-head"><div><p>MODULE-WISE · LATEST 4 WEEKS AUTOMATIC FORECAST</p><h2>4 WEEKS PLAN</h2></div><div className="iso-plan-chart-source">FORECAST ISSUE DATE</div></div>
        <div className="iso-premium-kpis"><article><span>NO. OF ISOs</span><strong>{fourWeekPlan.weeks.reduce((sum,week)=>sum+week.total,0)}</strong><small>Latest four forecast weeks</small></article><article><span>HOLDS</span><strong>{fourWeekPlan.weeks.reduce((sum,week)=>sum+week.holds,0)}</strong><small>Latest four forecast weeks</small></article></div>
        <div className="iso-premium-chart"><div className="iso-premium-grid-lines"><i/><i/><i/><i/><i/></div><div className="iso-premium-bars">{fourWeekPlan.weeks.map((week,index)=>{const maxWeek=Math.max(1,...fourWeekPlan.weeks.map((item)=>item.total));return <article className="iso-premium-week" key={week.label}><div className="iso-premium-value">{week.total}</div><div className="iso-premium-track"><div className={'iso-premium-fill w'+(index+1)} style={{height:Math.max(week.total?12:3,(week.total/maxWeek)*100)+'%'}}><span>{week.total}</span></div></div><div className="iso-premium-label"><b>{week.label}</b><em>{week.range}</em><small>NO. OF ISOs</small></div></article>})}</div></div>
        <div className="iso-premium-table"><div className="head"><span>PLAN SUMMARY</span>{fourWeekPlan.weeks.map((week)=><b key={week.label}>{week.label}</b>)}<strong>NO. OF ISOs</strong></div><div className="body"><span>FORECAST COUNT</span>{fourWeekPlan.weeks.map((week)=><b key={week.label}>{week.total}</b>)}<strong>{fourWeekPlan.weeks.reduce((sum,week)=>sum+week.total,0)}</strong></div></div>
        <div className="iso-plan-chart-note"><span>Counts dated FORECAST ISSUE DATE rows from the latest four available weeks.</span><span>Blank rows, delete markers, and struck-through forecast dates are excluded. HOLD counts only explicit YES values.</span></div>
      </section> : null}</> : <div className="iso-empty"><h2>{modules.length ? 'Select a module or open Console' : 'ISO Excel data not found'}</h2><span>{modules.length ? 'Module view shows module-wise counts. Console shows the complete project.' : 'Connect the folder containing ' + ISO_FILE_PREFIX + ' from Reports and Status.'}</span></div>}</section>
    </div>
  </Layout>;
}

function Integrated({ open }) {
  return <section className="integrated"><header><p>INTEGRATED TOOL</p><h2>Overall Project and Module-Wise Status</h2></header><div>{GROUPS.map(([groupName, items], groupIndex) => <article key={groupName}><h3>{groupName}</h3><section>{items.map((item, itemIndex) => <button className={`c${(groupIndex + itemIndex) % 5}`} key={item} onClick={() => open(item)}>{item}</button>)}</section></article>)}</div></section>;
}

function Layout({ title, sub, back, children }) {
  return <main className="page"><header className="top"><button onClick={back}>Back</button><div><p>REPORTS GENERATE TOOLS</p><h1>{title}</h1><span>{sub}</span></div></header><section className="content">{children}</section></main>;
}
