const fs = require('fs');
const path = require('path');

const target = process.argv[2] || path.join(process.cwd(), 'src', 'app.jsx');
const original = fs.readFileSync(target, 'utf8');
let code = original;

const fail = (message) => {
  throw new Error(message + ' No file was changed.');
};

const componentStart = code.indexOf('function IsoPlanningDashboardPage(');
if (componentStart < 0) fail('ISO Planning Dashboard was not found.');
const componentEnd = code.indexOf('\nfunction Integrated({ open }) {', componentStart);
if (componentEnd < 0) fail('ISO Planning Dashboard end boundary was not found.');

const before = code.slice(0, componentStart);
let component = code.slice(componentStart, componentEnd);
const after = code.slice(componentEnd);

if (!component.includes('ISO PROGRESS CONTROL')) fail('ISO Progress Control was not found.');
if (!component.includes('WORK FRONT AVAILABLE')) fail('Work Front Available row was not found.');

const oldLightGap = '.iso-stage-table tbody tr+tr th,.iso-stage-table tbody tr+tr td{border-top:18px solid #f4f8fb}.iso-stage-table tbody tr+tr th:first-child{border-top-left-radius:8px}.iso-stage-table tbody tr+tr td:last-child{border-top-right-radius:8px}';
const blackGap = '.iso-stage-table tbody tr+tr th,.iso-stage-table tbody tr+tr td{border-top:18px solid #111827}.iso-stage-table tbody tr+tr th:first-child{border-top-left-radius:8px}.iso-stage-table tbody tr+tr td:last-child{border-top-right-radius:8px}';

if (component.includes(blackGap)) {
  console.log('Black bordered gap is already present. Nothing was changed.');
  process.exit(0);
}

if (component.includes(oldLightGap)) {
  component = component.replace(oldLightGap, blackGap);
} else {
  const cssAnchor = '.iso-stage-table th,.iso-stage-table td{';
  if (!component.includes(cssAnchor)) fail('ISO Progress table cell CSS was not found.');
  component = component.replace(cssAnchor, blackGap + cssAnchor);
}

if (!component.includes('border-top:18px solid #111827')) fail('Black bordered gap was not applied.');

code = before + component + after;
const temp = target + '.tmp';
fs.writeFileSync(temp, code, 'utf8');
fs.renameSync(temp, target);

console.log('Updated successfully: ' + target);
console.log('Added only a visible 18px black bordered gap between Completed and Work Front Available.');
console.log('All counts, rows, buttons, downloads, formulas, toolbar and Excel logic remain unchanged.');
console.log('Original bytes: ' + Buffer.byteLength(original) + ' | Updated bytes: ' + Buffer.byteLength(code));
