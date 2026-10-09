const fs = require('fs');
const path = require('path');

const target = process.argv[2] || path.join(process.cwd(), 'src', 'app.jsx');
const original = fs.readFileSync(target, 'utf8');
let code = original;

function fail(message) {
  throw new Error(message + ' No file was changed.');
}

function insertAfterOnce(anchor, addition, label) {
  if (!code.includes(anchor)) fail('Update stopped: ' + label + ' anchor was not found.');
  code = code.replace(anchor, anchor + addition);
}

if (code.includes('function findEquipmentReportWorkbook(') || code.includes('function countEquipmentOrientationRemarks(')) {
  console.log('Equipment Orientation and Position Report automation is already present. Nothing was changed.');
  process.exit(0);
}

const finderAnchorElbow = [
  'function findElbowBendReportWorkbook(workbooks) {',
  '  return workbooks.find((entry) => {',
  "    const fileNameWithoutExtension = entry.file.name.replace(/\\.[^.]+$/, '');",
  "    return normalize(fileNameWithoutExtension).startsWith(normalize('EHH_ELBOW_BEND_ANG_REPORT'));",
  '  }) ?? null;',
  '}'
].join('\n');

const finderAnchorPrimary = [
  'function findPrimarySupportReportWorkbook(workbooks) {',
  "  return workbooks.find((entry) => isNamedReportWorkbook(entry, 'EHH_SUPPORT_REPORT')) ?? null;",
  '}'
].join('\n');

const finderAddition = [
  '',
  'function findEquipmentReportWorkbook(workbooks) {',
  '  return workbooks.find((entry) => {',
  "    const fileNameWithoutExtension = entry.file.name.replace(/\\.[^.]+$/, '');",
  "    return normalize(fileNameWithoutExtension).startsWith(normalize('EHH_EQUIPMENT_REPORT'));",
  '  }) ?? null;',
  '}'
].join('\n');

if (code.includes(finderAnchorElbow)) insertAfterOnce(finderAnchorElbow, finderAddition, 'equipment workbook finder');
else insertAfterOnce(finderAnchorPrimary, finderAddition, 'equipment workbook finder');

const countAnchorElbow = [
  'function countElbowBendRemarks(workbooks, moduleName) {',
  '  return countReportRemarks(findElbowBendReportWorkbook(workbooks), moduleName);',
  '}'
].join('\n');

const countAnchorPrimary = [
  'function countPrimarySupportRemarks(workbooks, moduleName) {',
  '  return countReportRemarks(findPrimarySupportReportWorkbook(workbooks), moduleName);',
  '}'
].join('\n');

const countAddition = [
  '',
  'function countEquipmentOrientationRemarks(workbooks, moduleName) {',
  '  return countReportRemarks(findEquipmentReportWorkbook(workbooks), moduleName);',
  '}'
].join('\n');

if (code.includes(countAnchorElbow)) insertAfterOnce(countAnchorElbow, countAddition, 'equipment remarks counter');
else insertAfterOnce(countAnchorPrimary, countAddition, 'equipment remarks counter');

const countSearchElbow = "item === 'Elbow and Bend Report' ? countElbowBendRemarks(workbooks, selected) : item === 'Fluid Code'";
const countReplaceElbow = "item === 'Elbow and Bend Report' ? countElbowBendRemarks(workbooks, selected) : item === 'Equipment Orientation and Position Report' ? countEquipmentOrientationRemarks(workbooks, selected) : item === 'Fluid Code'";
const countSearchPrimary = "item === 'Primary Support Report' ? countPrimarySupportRemarks(workbooks, selected) : item === 'Fluid Code'";
const countReplacePrimary = "item === 'Primary Support Report' ? countPrimarySupportRemarks(workbooks, selected) : item === 'Equipment Orientation and Position Report' ? countEquipmentOrientationRemarks(workbooks, selected) : item === 'Fluid Code'";

if (code.includes(countSearchElbow)) code = code.replace(countSearchElbow, countReplaceElbow);
else if (code.includes(countSearchPrimary)) code = code.replace(countSearchPrimary, countReplacePrimary);
else fail('Update stopped: dashboard count anchor was not found.');

const downloadAnchorElbow = [
  "    if (reportName === 'Elbow and Bend Report') {",
  '      const workbookEntry = findElbowBendReportWorkbook(workbooks);',
  "      if (!workbookEntry) { window.alert('No EHH_ELBOW_BEND_ANG_REPORT Excel file was found in the connected folder.'); return; }",
  "      if (!createNamedRemarksWorkbook(workbookEntry, selected, 'Elbow and Bend Report')) window.alert('No ' + selected + ' row with remarks was found for Elbow and Bend Report.');",
  '      return;',
  '    }'
].join('\n');

const downloadAnchorPrimaryTemplate = [
  "    if (reportName === 'Primary Support Report') {",
  '      const workbookEntry = findPrimarySupportReportWorkbook(workbooks);',
  "      if (!workbookEntry) { window.alert('No EHH_SUPPORT_REPORT Excel file was found in the connected folder.'); return; }",
  "      if (!createNamedRemarksWorkbook(workbookEntry, selected, 'Primary Support Report')) window.alert(`No ${selected} row with remarks was found for Primary Support Report.`);",
  '      return;',
  '    }'
].join('\n');

const downloadAnchorPrimaryConcat = [
  "    if (reportName === 'Primary Support Report') {",
  '      const workbookEntry = findPrimarySupportReportWorkbook(workbooks);',
  "      if (!workbookEntry) { window.alert('No EHH_SUPPORT_REPORT Excel file was found in the connected folder.'); return; }",
  "      if (!createNamedRemarksWorkbook(workbookEntry, selected, 'Primary Support Report')) window.alert('No ' + selected + ' row with remarks was found for Primary Support Report.');",
  '      return;',
  '    }'
].join('\n');

const downloadAddition = [
  '',
  "    if (reportName === 'Equipment Orientation and Position Report') {",
  '      const workbookEntry = findEquipmentReportWorkbook(workbooks);',
  "      if (!workbookEntry) { window.alert('No EHH_EQUIPMENT_REPORT Excel file was found in the connected folder.'); return; }",
  "      if (!createNamedRemarksWorkbook(workbookEntry, selected, 'Equipment Orientation and Position Report')) window.alert('No ' + selected + ' row with remarks was found for Equipment Orientation and Position Report.');",
  '      return;',
  '    }'
].join('\n');

if (code.includes(downloadAnchorElbow)) insertAfterOnce(downloadAnchorElbow, downloadAddition, 'equipment download logic');
else if (code.includes(downloadAnchorPrimaryTemplate)) insertAfterOnce(downloadAnchorPrimaryTemplate, downloadAddition, 'equipment download logic');
else if (code.includes(downloadAnchorPrimaryConcat)) insertAfterOnce(downloadAnchorPrimaryConcat, downloadAddition, 'equipment download logic');
else fail('Update stopped: download logic anchor was not found.');

const temp = target + '.tmp';
fs.writeFileSync(temp, code, 'utf8');
fs.renameSync(temp, target);

console.log('Updated successfully: ' + target);
console.log('Added only Equipment Orientation and Position Report count and remarks-download automation.');
console.log('Original bytes: ' + Buffer.byteLength(original) + ' | Updated bytes: ' + Buffer.byteLength(code));
