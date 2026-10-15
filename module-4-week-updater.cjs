#!/usr/bin/env node
/*
 * module-4-week-updater.cjs
 *
 * Patches src/app.jsx IN PLACE (your file is NOT replaced; a backup is made first).
 *
 * Changes ONLY the module-wise plan that opens when you select a module (outside the Planning page):
 *   - it becomes a "4 WEEKS PLAN"  (W1 .. W4)
 *   - W1..W4 = the latest four distinct dates of that module in column FORECAST ISSUE DATE
 *     of sheet HH_TOPSIDE_OVERALL_CS_OFSI (oldest to newest); every bar counts the rows of that date
 *   - blank rows, "Line Deleted" rows and struck-through rows are not counted (same rule as Planning)
 * The Planning page (4 MONTHS PLAN) and every other part of the app are not touched.
 *
 * Run from the project root:   node module-4-week-updater.cjs
 * Optional custom path:        node module-4-week-updater.cjs path/to/app.jsx
 */
const fs = require('fs');
const path = require('path');

const target = path.resolve(process.argv[2] || path.join('src', 'app.jsx'));
if (!fs.existsSync(target)) { console.error('File not found: ' + target); process.exit(1); }

const NEW_FUNCTION = [
  "function buildIsoFourWeekPlan(moduleData, forecastMap) {",
  "  // Module-wise 4 WEEK plan: reads ONLY sheet HH_TOPSIDE_OVERALL_CS_OFSI (column FORECAST ISSUE DATE) through forecastMap,",
  "  // so blank rows, Line Deleted rows and struck-through rows are already excluded (same rule as the Planning page).",
  "  // The latest four distinct forecast dates of the module become W1..W4 (oldest to newest) and the rows of each date are counted.",
  "  const items = (forecastMap?.get(normalize(moduleData?.moduleName)) || []).slice().sort((a, b) => a.date - b.date);",
  "  const dateMap = new Map();",
  "  for (const item of items) dateMap.set(isoDateKey(item.date), item.date);",
  "  const selectedDates = [...dateMap.values()].sort((a, b) => a - b).slice(-4);",
  "  while (selectedDates.length < 4) selectedDates.unshift(null);",
  "  const weeks = selectedDates.map((selectedDate, weekIndex) => {",
  "    const dateKey = selectedDate ? isoDateKey(selectedDate) : '';",
  "    const rows = selectedDate ? items.filter((item) => isoDateKey(item.date) === dateKey) : [];",
  "    return {",
  "      weekIndex,",
  "      label: 'W' + (weekIndex + 1),",
  "      range: selectedDate ? isoShortDate(selectedDate) : 'No Date',",
  "      rows,",
  "      total: rows.length,",
  "      holds: rows.filter((item) => item.hold).length,",
  "      days: [],",
  "    };",
  "  });",
  "  return {",
  "    weeks,",
  "    total: weeks.reduce((sum, week) => sum + week.total, 0),",
  "    holds: weeks.reduce((sum, week) => sum + week.holds, 0),",
  "    maxValue: Math.max(1, ...weeks.map((week) => week.total)),",
  "  };",
  "}",
  ""
].join('\n');

const NEW_JSX = [
  "<div className=\"iso-plan-chart-head\"><div><p>MODULE-WISE \u00b7 LATEST 4 WEEKS AUTOMATIC FORECAST</p><h2>4 WEEKS PLAN</h2></div><div className=\"iso-plan-chart-source\">FORECAST ISSUE DATE</div></div>",
  "        <div className=\"iso-premium-kpis\"><article><span>NO. OF ISOs</span><strong>{moduleWeekPlan.total}</strong><small>Latest four forecast dates</small></article><article><span>HOLDS</span><strong>{moduleWeekPlan.holds}</strong><small>Latest four forecast dates</small></article></div>",
  "        <div className=\"iso-premium-chart\"><div className=\"iso-premium-grid-lines\"><i/><i/><i/><i/><i/></div><div className=\"iso-premium-bars\">{moduleWeekPlan.weeks.map((week,index)=><article className=\"iso-premium-week\" key={'week-'+index}><div className=\"iso-premium-value\">{week.total}</div><div className=\"iso-premium-track\"><div className={'iso-premium-fill w'+(index+1)} style={{height:Math.max(week.total?12:3,(week.total/moduleWeekPlan.maxValue)*100)+'%'}}><span>{week.total}</span></div></div><div className=\"iso-premium-label\"><b>{week.label}</b><em>{week.range}</em><small>NO. OF ISOs</small></div></article>)}</div></div>",
  "        <div className=\"iso-premium-table\"><div className=\"head\"><span>PLAN SUMMARY</span>{moduleWeekPlan.weeks.map((week)=><b key={'head-'+week.label}>{week.label}</b>)}<strong>NO. OF ISOs</strong></div><div className=\"body\"><span>FORECAST COUNT</span>{moduleWeekPlan.weeks.map((week)=><b key={'count-'+week.label}>{week.total}</b>)}<strong>{moduleWeekPlan.total}</strong></div></div>",
  "        <div className=\"iso-plan-chart-note\"><span>Counts FORECAST ISSUE DATE rows of the latest four distinct dates (W1 to W4) from sheet HH_TOPSIDE_OVERALL_CS_OFSI.</span><span>Blank rows, Line Deleted rows and struck-through rows are excluded. HOLD counts only explicit YES values.</span></div>"
].join('\n');

let s = fs.readFileSync(target, 'utf8');
const original = s;
const log = [];
const fail = (message) => { console.error('STOPPED (nothing changed): ' + message); process.exit(1); };

if (!s.includes('buildIsoForecastMap(workbooks)')) fail('the sheet-based forecast map (buildIsoForecastMap) is not in this app.jsx. Run forecast-fix-updater.cjs first.');

// 1) buildIsoFourWeekPlan -> reads the same forecastMap as the Planning page
const fnStart = s.indexOf('function buildIsoFourWeekPlan(');
if (fnStart < 0) fail('could not find: function buildIsoFourWeekPlan(');
const fnEnd = s.indexOf('\nfunction ', fnStart + 20);
if (fnEnd < 0) fail('could not find the end of buildIsoFourWeekPlan');
s = s.slice(0, fnStart) + NEW_FUNCTION + s.slice(fnEnd);
log.push('done   - buildIsoFourWeekPlan now uses sheet HH_TOPSIDE_OVERALL_CS_OFSI (latest 4 dates)');

// 2) dashboard: moduleWeekPlan
if (!s.includes('const moduleWeekPlan = useMemo(')) {
  const anchor = '  const moduleMonthPlan = useMemo(() => buildIsoLatestFourMonthPlan(planSelection, forecastMap), [planSelection, forecastMap]);';
  const at = s.indexOf(anchor);
  if (at < 0) fail('could not find: moduleMonthPlan line in IsoPlanningDashboardPage');
  s = s.slice(0, at + anchor.length) + '\n  const moduleWeekPlan = useMemo(() => buildIsoFourWeekPlan(planSelection, forecastMap), [planSelection, forecastMap]);' + s.slice(at + anchor.length);
  log.push('done   - dashboard: moduleWeekPlan added');
} else log.push('skip   - moduleWeekPlan already present');

// 3) module view JSX: months -> weeks
const jsxStartMarker = '<div className="iso-plan-chart-head"><div><p>MODULE-WISE ·';
const jsxStart = s.indexOf(jsxStartMarker);
if (jsxStart < 0) fail('could not find the MODULE-WISE plan block');
const jsxEndMarker = '\n      </section> : null}';
const jsxEnd = s.indexOf(jsxEndMarker, jsxStart);
if (jsxEnd < 0) fail('could not find the end of the MODULE-WISE plan block');
s = s.slice(0, jsxStart) + NEW_JSX + s.slice(jsxEnd);
log.push('done   - module view now shows "4 WEEKS PLAN" (W1 to W4)');

// safety checks before writing
const open = (s.match(/\{/g) || []).length; const close = (s.match(/\}/g) || []).length;
if (open !== close) fail('brace mismatch after patch (' + open + ' vs ' + close + ')');
if ((s.match(/function buildIsoFourWeekPlan\(/g) || []).length !== 1) fail('duplicate buildIsoFourWeekPlan');
if ((s.match(/const moduleWeekPlan = useMemo\(/g) || []).length !== 1) fail('duplicate moduleWeekPlan');
if (!s.includes('LATEST 4 MONTHS AUTOMATIC FORECAST</p><h2>4 MONTHS PLAN</h2>')) fail('the Planning page block would have been changed (safety stop)');

let backup = path.join(path.dirname(target), 'app.before-module-week-plan.jsx');
if (fs.existsSync(backup)) backup = path.join(path.dirname(target), 'app.before-module-week-plan-' + Date.now() + '.jsx');
fs.writeFileSync(backup, original, 'utf8');
fs.writeFileSync(target, s, 'utf8');

console.log(log.join('\n'));
console.log('\nBackup : ' + backup);
console.log('Updated: ' + target);
console.log('\nNext: restart "npm run dev" and select a module (for example M1C).');
