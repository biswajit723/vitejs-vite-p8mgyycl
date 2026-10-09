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

if (code.includes('function findElbowBendReportWorkbook(') || code.includes('function countElbowBendRemarks(')) {
  console.log('Elbow and Bend Report automation is already present. Nothing was changed.');
  process.exit(0);
}

const finderAnchor = [
  "function findPrimarySupportReportWorkbook(workbooks) {",
  "  return workbooks.find((entry) => isNamedReportWorkbook(entry, 'EHH_SUPPORT_REPORT')) ?? null;",
  "}"
].join('\n');

const finderAddition = [
  '',
  'function findElbowBendReportWorkbook(workbooks) {',
  '  return workbooks.find((entry) => {',
  "    const fileNameWithoutExtension = entry.file.name.replace(/\\.[^.]+$/, '');",
  "    return normalize(fileNameWithoutExtension).startsWith(normalize('EHH_ELBOW_BEND_ANG_REPORT'));",
  '  }) ?? null;',
  '}'
].join('\n');
insertAfterOnce(finderAnchor, finderAddition, 'workbook finder');

const countAnchor = [
  'function countPrimarySupportRemarks(workbooks, moduleName) {',
  '  return countReportRemarks(findPrimarySupportReportWorkbook(workbooks), moduleName);',
  '}'
].join('\n');

const countAddition = [
  '',
  'function countElbowBendRemarks(workbooks, moduleName) {',
  '  return countReportRemarks(findElbowBendReportWorkbook(workbooks), moduleName);',
  '}'
].join('\n');
insertAfterOnce(countAnchor, countAddition, 'remarks counter');

const countSearch = "item === 'Primary Support Report' ? countPrimarySupportRemarks(workbooks, selected) : item === 'Fluid Code'";
const countReplacement = "item === 'Primary Support Report' ? countPrimarySupportRemarks(workbooks, selected) : item === 'Elbow and Bend Report' ? countElbowBendRemarks(workbooks, selected) : item === 'Fluid Code'";
if (!code.includes(countSearch)) fail('Update stopped: dashboard count anchor was not found.');
code = code.replace(countSearch, countReplacement);

const downloadAnchor = [
  "    if (reportName === 'Primary Support Report') {",
  '      const workbookEntry = findPrimarySupportReportWorkbook(workbooks);',
  "      if (!workbookEntry) { window.alert('No EHH_SUPPORT_REPORT Excel file was found in the connected folder.'); return; }",
  "      if (!createNamedRemarksWorkbook(workbookEntry, selected, 'Primary Support Report')) window.alert(`No ${selected} row with remarks was found for Primary Support Report.`);",
  '      return;',
  '    }'
].join('\n');

const downloadAddition = [
  '',
  "    if (reportName === 'Elbow and Bend Report') {",
  '      const workbookEntry = findElbowBendReportWorkbook(workbooks);',
  "      if (!workbookEntry) { window.alert('No EHH_ELBOW_BEND_ANG_REPORT Excel file was found in the connected folder.'); return; }",
  "      if (!createNamedRemarksWorkbook(workbookEntry, selected, 'Elbow and Bend Report')) window.alert('No ' + selected + ' row with remarks was found for Elbow and Bend Report.');",
  '      return;',
  '    }'
].join('\n');
insertAfterOnce(downloadAnchor, downloadAddition, 'download logic');

const temp = target + '.tmp';
fs.writeFileSync(temp, code, 'utf8');
fs.renameSync(temp, target);

console.log('Updated successfully: ' + target);
console.log('Added only Elbow and Bend Report count and remarks-download automation.');
console.log('Original bytes: ' + Buffer.byteLength(original) + ' | Updated bytes: ' + Buffer.byteLength(code));
