const fs = require('fs');
const path = require('path');

const target = path.resolve(process.argv[2] || 'src/app.jsx');
const fail = (message) => {
  console.error('Error: ' + message + ' No file was changed.');
  process.exit(1);
};

if (!fs.existsSync(target)) fail('Target file not found: ' + target);
const before = fs.readFileSync(target, 'utf8');
let after = before;

const exactPrefix = "const ISO_FILE_PREFIX = '4193_HH_Isometrics Control Sheet-Piping_OFSI-1';";
const connectedPrefix = "const ISO_FILE_PREFIX = '4193_HH_Isometrics Control Sheet-Piping_OFSI';";
if (after.includes(exactPrefix)) after = after.replace(exactPrefix, connectedPrefix);
else if (!after.includes(connectedPrefix)) fail('ISO_FILE_PREFIX line was not found.');

const exactOnly = `  const sourceWorkbook = exactWorkbooks
    .slice()
    .sort((a, b) => Number(b.file?.lastModified || 0) - Number(a.file?.lastModified || 0))[0];`;

const connectedSelection = `  const fallbackWorkbooks = availableWorkbooks.filter((entry) =>
    normalize(entry.file.name.replace(/\\.[^.]+$/, '')).startsWith(exactFileName)
  );
  const sourceWorkbook = (exactWorkbooks.length ? exactWorkbooks : fallbackWorkbooks)
    .slice()
    .sort((a, b) => Number(b.file?.lastModified || 0) - Number(a.file?.lastModified || 0))[0];`;

if (after.includes(exactOnly)) after = after.replace(exactOnly, connectedSelection);
else if (!after.includes(connectedSelection)) fail('Excel source-selection block was not found.');

if (after === before) {
  console.log('No update needed: connected Excel source selection is already restored.');
  console.log('No file was changed.');
  process.exit(0);
}

fs.writeFileSync(target, after, 'utf8');
console.log('Updated successfully: ' + target);
console.log('Planning Excel source selection was restored.');
console.log('No counting, UI, CSS, button, or other file was changed.');
