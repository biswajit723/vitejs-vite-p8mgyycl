const fs = require('fs');
const path = require('path');

const target = process.argv[2] || path.join(process.cwd(), 'src', 'app.jsx');
const original = fs.readFileSync(target, 'utf8');
let code = original;

const fail = (message) => {
  throw new Error(message + ' No file was changed.');
};

if (!code.includes('function buildIsoFourWeekPlan(')) {
  fail('4 Weeks Plan calculation was not found. Apply the 4 Weeks Planning Graph updater first.');
}

const componentStart = code.indexOf('function IsoPlanningDashboardPage(');
if (componentStart < 0) fail('ISO Planning Dashboard was not found.');
const componentEnd = code.indexOf('\nfunction Integrated({ open }) {', componentStart);
if (componentEnd < 0) fail('ISO Planning Dashboard end boundary was not found.');

let before = code.slice(0, componentStart);
let component = code.slice(componentStart, componentEnd);
let after = code.slice(componentEnd);

if (!component.includes('iso-module-plan-below-progress')) {
  fail('Module-wise graph below ISO Progress Control was not found. Apply the below-table updater first.');
}

// Replace only the weekly grouping function.
const weeklyFunctionRegex = /function buildIsoFourWeekPlan\(moduleData\) \{[\s\S]*?\n\}/;
const weeklyMatch = before.match(weeklyFunctionRegex);
if (!weeklyMatch) fail('4 Weeks Plan function block was not found safely.');

const weeklyFunction = `function buildIsoFourWeekPlan(moduleData) {
  const validEntries = (moduleData?.entries || [])
    .filter(isIsoPlanningValidEntry)
    .map((entry) => ({
      entry,
      date: entry.forecastColumn >= 0 ? parseIsoForecastDate(entry.row?.[entry.forecastColumn]) : null,
    }))
    .filter((item) => item.date && item.date.getDay() >= 1 && item.date.getDay() <= 5)
    .sort((a, b) => a.date - b.date);

  const weekGroups = new Map();
  for (const item of validEntries) {
    const monday = isoMonday(item.date);
    const key = isoDateKey(monday);
    if (!weekGroups.has(key)) weekGroups.set(key, { monday, rows: [] });
    weekGroups.get(key).rows.push(item);
  }

  const availableWeeks = [...weekGroups.values()].sort((a, b) => a.monday - b.monday);
  const selectedWeeks = availableWeeks.slice(0, 4);
  const fallbackStart = availableWeeks[0]?.monday || isoMonday(new Date());
  while (selectedWeeks.length < 4) {
    const monday = new Date(fallbackStart);
    monday.setDate(fallbackStart.getDate() + selectedWeeks.length * 7);
    selectedWeeks.push({ monday, rows: [] });
  }

  const dayLabels = ['MON', 'TUE', 'WED', 'THU', 'FRI'];
  const weeks = selectedWeeks.map((group, weekIndex) => {
    const monday = new Date(group.monday);
    const friday = new Date(monday);
    friday.setDate(monday.getDate() + 4);
    const days = dayLabels.map((label, dayIndex) => {
      const date = new Date(monday);
      date.setDate(monday.getDate() + dayIndex);
      const dateRows = group.rows.filter((item) => isoDateKey(item.date) === isoDateKey(date));
      return { label, date, dateLabel: isoShortDate(date), count: dateRows.length };
    });
    return {
      weekIndex,
      label: 'W' + (weekIndex + 1),
      range: isoShortDate(monday) + ' - ' + isoShortDate(friday),
      rows: group.rows,
      total: group.rows.length,
      holds: group.rows.filter((item) => item.entry.holdColumn >= 0 && normalize(item.entry.row?.[item.entry.holdColumn]) === 'yes').length,
      days,
    };
  });

  return {
    weeks,
    maxValue: Math.max(1, ...weeks.flatMap((week) => week.days.map((day) => day.count))),
  };
}`;
before = before.replace(weeklyMatch[0], weeklyFunction);

// Graph must appear automatically after module click, not depend on PLANNING button.
component = component.replace(
  /\{planningMode && moduleSelection && moduleSelection\.entries\.length \? <section className="iso-plan-chart-card iso-module-plan-below-progress">/,
  '{moduleSelection && moduleSelection.entries.length ? <section className="iso-plan-chart-card iso-module-plan-below-progress">'
);

if (!component.includes('{moduleSelection && moduleSelection.entries.length ? <section className="iso-plan-chart-card iso-module-plan-below-progress">')) {
  fail('Automatic module graph condition could not be applied.');
}

// Planning button no longer controls graph visibility. It keeps the existing module dashboard view only.
component = component.replace(
  /const openPlanning = \(\) => \{[\s\S]*?\n  \};/,
  `const openPlanning = () => {
    setPlanningMode(false);
    if (typeof setBlankToolPage === 'function') setBlankToolPage(null);
    window.setTimeout(() => metricsRef.current?.scrollIntoView({ behavior: 'smooth', block: 'start' }), 0);
  };`
);
if (!component.includes('const [blankToolPage, setBlankToolPage]')) {
  component = component.replace("    if (typeof setBlankToolPage === 'function') setBlankToolPage(null);\n", '');
}

// Display both weekday and actual date under each daily bar.
const oldDayLabel = '<small>{day.label}</small>';
const newDayLabel = '<small>{day.label}<em>{day.dateLabel}</em></small>';
if (!component.includes(oldDayLabel)) fail('Daily graph label was not found.');
component = component.replace(oldDayLabel, newDayLabel);

// Add date-label styling only.
const oldSmallCss = '.iso-plan-day small{margin:7px 0 5px;color:#d9ebf5;font-size:9px;font-weight:800}';
const newSmallCss = '.iso-plan-day small{margin:7px 0 5px;color:#d9ebf5;font-size:9px;font-weight:800;text-align:center;line-height:1.25}.iso-plan-day small em{display:block;margin-top:3px;color:#7dd3fc;font-size:8px;font-style:normal;font-weight:800;white-space:nowrap}';
if (!component.includes(oldSmallCss)) fail('Daily graph-label CSS was not found.');
component = component.replace(oldSmallCss, newSmallCss);

// Planning button active styling remains visually normal; module click is the automatic trigger.
component = component.replace("(planningMode ? 'active' : '')", "(!consoleMode && moduleSelection ? 'active' : '')");

// Strict verification.
if (component.includes('planningMode && moduleSelection && moduleSelection.entries.length')) {
  fail('Graph is still connected to Planning mode.');
}
if (!component.includes('day.dateLabel')) fail('Actual date labels were not added.');
if (!before.includes('const weekGroups = new Map();')) fail('Week-wise date grouping was not added.');
if (!before.includes('availableWeeks.slice(0, 4)')) fail('Four populated Excel weeks filter was not added.');
if (!component.includes('moduleSelection && moduleSelection.entries.length')) fail('Automatic module-click graph condition is missing.');

code = before + component + after;
const temp = target + '.tmp';
fs.writeFileSync(temp, code, 'utf8');
fs.renameSync(temp, target);

console.log('Updated successfully: ' + target);
console.log('The graph now appears automatically below ISO Progress Control when a module is clicked.');
console.log('Forecast dates are grouped into the first four populated Excel weeks and every Monday-Friday bar shows its actual date.');
console.log('Planning button no longer controls graph visibility.');
console.log('All existing counts, HOLD rules, downloads, tables, Console, Analytics and Revision remain unchanged.');
console.log('Original bytes: ' + Buffer.byteLength(original) + ' | Updated bytes: ' + Buffer.byteLength(code));
