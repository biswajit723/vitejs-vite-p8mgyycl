import { useEffect, useMemo, useRef, useState } from 'react';
import * as XLSX from 'xlsx-js-style';

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

function rowHasRemarks(sheetEntry, row) {
  return (sheetEntry.header?.remarksColumns ?? []).some((columnIndex) => hasMeaningfulRemark(row?.[columnIndex]));
}
function isNamedReportWorkbook(entry, expectedName) {
  const fileNameWithoutExtension = entry.file.name.replace(/\.[^.]+$/, '');
  const nameWithoutTrailingNumbers = fileNameWithoutExtension.replace(/[_\s-]*\d+$/, '');
  return normalize(nameWithoutTrailingNumbers) === normalize(expectedName);
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
  const rows = [
    [`${moduleName}-Status`, ...SIZE_BUCKETS, 'Grand Total'],
    ['MODELLED', ...SIZE_BUCKETS.map((bucket) => status['MODELLED'][bucket]), Object.values(status['MODELLED']).reduce((sum, value) => sum + value, 0)],
    ['NOT MODELLED', ...SIZE_BUCKETS.map((bucket) => status['NOT MODELLED'][bucket]), Object.values(status['NOT MODELLED']).reduce((sum, value) => sum + value, 0)],
    ['PLANNED', ...SIZE_BUCKETS.map((bucket) => planned[bucket]), SIZE_BUCKETS.reduce((sum, bucket) => sum + planned[bucket], 0)],
    ['GRAND TOTAL', ...SIZE_BUCKETS.map((bucket) => status['MODELLED'][bucket] + status['NOT MODELLED'][bucket] + planned[bucket]), SIZE_BUCKETS.reduce((sum, bucket) => sum + status['MODELLED'][bucket] + status['NOT MODELLED'][bucket] + planned[bucket], 0)],
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
  const list = useMemo(() => modules.filter((moduleName) => moduleName.toLowerCase().includes(query.trim().toLowerCase())), [modules, query]);
  const COUNT_ITEMS = useMemo(() => GROUPS.flatMap(([, items]) => items).filter((item) => item !== 'Size Status'), []);
  const fluidCodeStatus = useMemo(() => getFluidCodeStatus(workbooks, selected), [workbooks, selected]);
  const fluidCodeGrandTotal = useMemo(() => fluidCodeStatus.fluidCodes.reduce((total, fluidCode) => total + fluidCodeStatus.sizes.reduce((sum, size) => sum + (fluidCodeStatus.counts[fluidCode]?.[size] || 0), 0), 0), [fluidCodeStatus]);
  const lineStatus = useMemo(() => getLineStatus(workbooks, selected), [workbooks, selected]);
  const counts = useMemo(() => Object.fromEntries(COUNT_ITEMS.map((item) => [item, selected ? (item === 'MTO Report' ? countMtoRemarks(workbooks, selected) : item === 'Valve Report' ? countValveRemarks(workbooks, selected) : item === 'Pipe Branch Report' ? countPipeBranchRemarks(workbooks, selected) : item === 'ATTA Report' ? countAttaRemarks(workbooks, selected) : item === 'Primary Support Report' ? countPrimarySupportRemarks(workbooks, selected) : item === 'Fluid Code' ? fluidCodeGrandTotal : item === 'Line Status' ? lineStatus.grandTotal : countForModule(findReportWorkbook(workbooks, item), selected)) : 0])), [COUNT_ITEMS, workbooks, selected, fluidCodeGrandTotal, lineStatus]);

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
      // Preserve the module order from the Master Excel.
      const dynamicModules = [...moduleMap.values()];
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
          const publishResult = await publishCentralModules(dynamicModules, adminToken);
          if (publishResult?.skipped) {
            setCentralStatus('offline');
          } else {
            setCentralStatus('connected');
            window.alert('Master module list published successfully. All users will see this module list.');
          }
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

  const openIntegrated = (item) => {
    if ((item === 'Size Status' || item === 'Fluid Code' || item === 'Line Status') && selected) {
      setStatusView(item);
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
  const modelledTotal = SIZE_BUCKETS.reduce((sum, bucket) => sum + status['MODELLED'][bucket], 0);
  const notModelledTotal = SIZE_BUCKETS.reduce((sum, bucket) => sum + status['NOT MODELLED'][bucket], 0);
  const plannedTotal = SIZE_BUCKETS.reduce((sum, bucket) => sum + plannedNumber(bucket), 0);
  const grandTotal = modelledTotal + notModelledTotal + plannedTotal;
  const plannedForExport = Object.fromEntries(SIZE_BUCKETS.map((bucket) => [bucket, plannedNumber(bucket)]));
  return <main className="page size-status-page">
    <header className="top size-status-top">
      <div className="size-status-heading"><p>MODELLING STATUS</p><h1>{moduleName} Size Status</h1></div>
      <div className="size-status-actions"><button onClick={() => createSizeStatusWorkbook(moduleName, status, plannedForExport)}>Download Excel</button><button onClick={back}>Back</button></div>
    </header>
    <section className="content size-status-content"><div className="size-status-table-wrap"><table className="size-status-table">
      <thead><tr><th>{moduleName}-Status</th>{SIZE_BUCKETS.map((bucket) => <th key={bucket}>{bucket}</th>)}<th>Grand Total</th></tr></thead>
      <tbody>
        <tr><th>MODELLED</th>{SIZE_BUCKETS.map((bucket) => <td key={bucket}>{status['MODELLED'][bucket]}</td>)}<td>{modelledTotal}</td></tr>
        <tr><th>NOT MODELLED</th>{SIZE_BUCKETS.map((bucket) => <td key={bucket}>{status['NOT MODELLED'][bucket]}</td>)}<td>{notModelledTotal}</td></tr>
        <tr className="size-status-planned"><th>PLANNED</th>{SIZE_BUCKETS.map((bucket) => <td key={bucket}><input type="number" min="0" step="1" inputMode="numeric" value={planned[bucket]} onChange={(event) => updatePlanned(bucket, event.target.value)} onBlur={() => updatePlanned(bucket, plannedNumber(bucket))} aria-label={`Planned ${bucket}`} /></td>)}<td>{plannedTotal}</td></tr>
        <tr className="size-status-grand"><th>GRAND TOTAL</th>{SIZE_BUCKETS.map((bucket) => <td key={bucket}>{status['MODELLED'][bucket] + status['NOT MODELLED'][bucket] + plannedNumber(bucket)}</td>)}<td>{grandTotal}</td></tr>
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

function Integrated({ open }) {
  return <section className="integrated"><header><p>INTEGRATED TOOL</p><h2>Overall Project and Module-Wise Status</h2></header><div>{GROUPS.map(([groupName, items], groupIndex) => <article key={groupName}><h3>{groupName}</h3><section>{items.map((item, itemIndex) => <button className={`c${(groupIndex + itemIndex) % 5}`} key={item} onClick={() => open(item)}>{item}</button>)}</section></article>)}</div></section>;
}

function Layout({ title, sub, back, children }) {
  return <main className="page"><header className="top"><button onClick={back}>Back</button><div><p>REPORTS GENERATE TOOLS</p><h1>{title}</h1><span>{sub}</span></div></header><section className="content">{children}</section></main>;
}
