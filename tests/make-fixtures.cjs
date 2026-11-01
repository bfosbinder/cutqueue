// Builds made-up workbooks that move the schedule table around, to exercise
// the Excel loader. Not real data.
//   npm install xlsx@0.18.5   (anywhere; then NODE_PATH=<that node_modules>)
//   node tests/make-fixtures.cjs
const XLSX = require('xlsx');
const path = require('path');
const out = name => path.join(__dirname, 'fixtures', name);
const d = (m, day) => new Date(2026, m - 1, day, 12);   // noon avoids tz edge cases

const HEAD = ['ORDER TO RUN', 'Job', 'Part', 'Op #', 'Cell', 'Op Start Date', 'Op Cpt Date', 'Job Cpt Date',
  'Run Hours', 'Setup', 'Total', 'Hrs Left', 'OTTB %', 'OTTP%', 'NOTES', 'Schedule Status', 'LOADED ON MEMEX'];
const row = (o, job, part, op, cell, s, e, jc, run, setup, left) =>
  [o, job, part, op, cell, s, e, jc, run, setup, run + setup, left, '', '', '', 'Scheduled', 'Y'];

function sheet(aoa, atRow, atCol) {            // place a table at an offset
  const grid = [];
  aoa.forEach((r, i) => { grid[atRow + i] = []; r.forEach((v, j) => { grid[atRow + i][atCol + j] = v; }); });
  for (let i = 0; i < grid.length; i++) if (!grid[i]) grid[i] = [];
  return XLSX.utils.aoa_to_sheet(grid, { cellDates: true });
}
const write = (name, sheets) => {
  const wb = XLSX.utils.book_new();
  sheets.forEach(([n, ws]) => XLSX.utils.book_append_sheet(wb, ws, n));
  XLSX.writeFile(wb, out(name), { cellDates: true });
};

// 1. Header at C6 on one tab, junk above/below, a bad date, a totals row, a second tab with an old table at A1
const live = [
  HEAD,
  row(1, '41007', 'BRACKET-A114', '20', 'HAAS VF-2', d(11, 2), d(11, 4), d(11, 6), 14, 2, 16),
  row(2, '41012', 'SHAFT-220', '10', 'HAAS VF-2', d(11, 4), d(11, 6), d(11, 9), 12, 3, 15),
  row(3, '41020', 'MANIFOLD-07', '30', 'HAAS VF-2', d(11, 6), 'TBD', d(11, 12), 18, 4, 22),
  row(1, '41005', 'FLANGE-88', '10', 'MAZAK QT-200', d(11, 2), d(11, 3), d(11, 5), 7, 1, 8),
  row(2, '41015', 'PIN-045', '10', 'MAZAK QT-200', d(11, 3), d(11, 5), d(11, 6), 13, 2, 15),
  ['', '', '', '', 'TOTAL', '', '', '', '', '', '', 54],
];
const liveWs = sheet([['Weekly cut schedule'], ['Printed 11/1/2026'], [], [], ...live], 0, 2);
const archive = [HEAD, row(1, '39001', 'OLD-PART', '10', 'HAAS VF-2', d(10, 1), d(10, 3), d(10, 4), 10, 1, 0)];
write('offset-tabs.xlsx', [
  ['Notes', XLSX.utils.aoa_to_sheet([['Read me'], ['Schedule is on the next tab']])],
  ['Schedule', liveWs],
  ['Archive', sheet(archive, 0, 0)],
]);

// 2. Columns reordered, renamed headers, table starts at F2
const alt = [
  ['Machine', 'Job #', 'Op Start', 'Op Complete Date', 'Hours Left', 'Part Number', 'Op'],
  ['CNC-1', '50001', d(11, 3), d(11, 5), 10, 'GEAR-9', '10'],
  ['CNC-1', '50002', d(11, 5), d(11, 6), 6, 'GEAR-10', '20'],
  ['CNC-2', '50010', d(11, 2), d(11, 9), 30, 'CASE-3', '10'],
];
write('reordered.xlsx', [['Sheet1', sheet(alt, 1, 5)]]);

// 3. Nothing usable: Op Cpt Date column missing
write('missing-column.xlsx', [['Sheet1', sheet([['Job', 'Op Start Date', 'Hrs Left'], ['1', d(11, 2), 5]], 3, 1)]]);
console.log('fixtures written');
