const XLSX = require('xlsx');
const Core = require('./core.js');
const fs = require('fs');
const U = (process.argv[2] || './fixtures').replace(/\/?$/, '/');

function aoa(file) {
  const wb = XLSX.readFile(file, { cellDates: false });
  return XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { header: 1, raw: true, defval: null });
}
const iso = t => {
  const m = String(t).match(/(\d{1,2})\.(\d{1,2})\.(\d{4})/);
  return m ? `${m[3]}-${m[2].padStart(2, '0')}-${m[1].padStart(2, '0')}` : null;
};

let fails = 0;
const ok = (cond, msg) => { console.log((cond ? 'PASS' : 'FAIL') + ' - ' + msg); if (!cond) fails++; };

// ---------- Test 1: F1 (June 2026) + F4 ----------
const f1 = aoa(U + 'F1_clean.xls'), f4 = aoa(U + 'F4_clean.xls');
const r1 = Core.process(f1, f4, {});
ok(r1.log.status !== 'ERROR', 'F1+F4 processes: ' + r1.log.status);
ok(r1.result.summary.length === 5, `F1: 5 employees (got ${r1.result.summary.length})`);
ok(r1.result.accounting.length > 90, `F4: employees parsed (got ${r1.result.accounting.length})`);
const w = re => r1.log.warnings.filter(x => re.test(x));
ok(w(/Дневна сума несъответствие/).length === 0, 'F1: recomputed day earnings match report');
ok(w(/Месечна сума несъответствие/).length === 0, 'F1: monthly totals match report');
ok(w(/% несъответствие/).length === 0, 'F1: recomputed % matches report');
console.log('  F4 control-sum notes (anonymization noise expected):');
w(/F4 контролна/).forEach(x => console.log('   ', x));
console.log('  Skipped-row warnings:');
r1.log.warnings.filter(x => /Пропуснат ред/.test(x)).forEach(x => console.log('   ', x));

// ---------- Test 2: F2 (Dec 2024) - replicate the MANUAL bonus formulas exactly ----------
const f2 = aoa(U + 'F2_clean.xls');
const expDay = new Map(), expMonth = new Map();
let block = -1;
for (let i = 0; i < f2.length; i++) {
  const a = String((f2[i] || [])[0] || '').trim();
  if (/,\s*БРИГАДА/.test(a)) { block++; continue; }
  const dm = a.match(/^ЗА\s+ДАТА\s*\/\s*([\d.]+)\s*\/\s*:?$/);
  if (dm && typeof f2[i][11] === 'number') expDay.set(block + '|' + iso(dm[1]), f2[i][11]);
  if (/^СУМА\s+ЗАРАБОТКА/.test(a) && typeof f2[i][11] === 'number') expMonth.set(block, f2[i][11]);
}
// manual formula uses REPORT day earnings (col K) -> compare in 'report' mode
const r2 = Core.process(f2, null, { roundDecimals: -1, dayEarningsSource: 'report' });
ok(r2.log.status !== 'ERROR', 'F2 processes: ' + r2.log.status);
ok(r2.result.summary.length === 10, `F2: 10 employee blocks (got ${r2.result.summary.length})`);

// days array is built employee-by-employee; slice by each employee's daysWorked count
const gotDay = new Map();
{
  let i = 0;
  r2.result.summary.forEach((emp, ei) => {
    for (let k = 0; k < emp['Отработени дни (F1)']; k++, i++) {
      const d = r2.result.days[i];
      gotDay.set(ei + '|' + d['Дата'], d['Бонус (лв)']);
    }
  });
}
let dayDiffs = [], dayMax = 0, compared = 0;
for (const [k, exp] of expDay) {
  const got = gotDay.get(k);
  if (got == null) { dayDiffs.push({ k, exp, got: 'MISSING' }); continue; }
  compared++;
  const diff = Math.abs(exp - got); if (diff > dayMax) dayMax = diff;
  if (diff > 0.005) dayDiffs.push({ k, exp, got });
}
ok(dayDiffs.length === 0 && compared > 0, `F2: bonus per day == manual formulas (${compared} compared, maxDiff=${dayMax.toExponential(2)})`);
if (dayDiffs.length) console.log(dayDiffs.slice(0, 8));

let mDiffs = [], mCompared = 0;
for (const [b, exp] of expMonth) {
  const got = r2.result.summary[b] && r2.result.summary[b]['Бонус общо (лв)'];
  mCompared++;
  if (got == null || Math.abs(exp - got) > 0.005) mDiffs.push({ b, exp, got });
}
ok(mDiffs.length === 0 && mCompared > 0, `F2: monthly bonus totals == manual results (${mCompared} compared)`);
if (mDiffs.length) console.log(mDiffs);

// ---------- Test 3: F3 - 'both' mode replicates the internal tool's totals ----------
const f3 = aoa(U + 'F3_clean.xlsx');
const r3 = Core.process(f3, null, { roundDecimals: -1, bonusLevel: 'both', bonusThresholdPct: 100, dayEarningsSource: 'report' });
ok(r3.log.status !== 'ERROR', 'F3 processes: ' + r3.log.status);
const emp1 = r3.result.summary[0];
ok(Math.abs(emp1['Бонус общо (лв)'] - 31.2523) < 0.02,
   `F3 'both' mode reproduces internal tool total 31.2523 (got ${emp1['Бонус общо (лв)'].toFixed(4)})`);
const r3day = Core.process(f3, null, { roundDecimals: -1, bonusLevel: 'day', dayEarningsSource: 'report' });
console.log('  F3 with day-level rule instead:', r3day.result.summary.map(e => e['Бонус общо (лв)'].toFixed(3)).join(', '));

// ---------- sample outputs from F1+F4 ----------
const out = XLSX.utils.book_new();
XLSX.utils.book_append_sheet(out, XLSX.utils.json_to_sheet(r1.result.summary), 'Обобщение');
XLSX.utils.book_append_sheet(out, XLSX.utils.json_to_sheet(r1.result.days), 'По дни');
XLSX.utils.book_append_sheet(out, XLSX.utils.json_to_sheet(r1.result.ops), 'Операции');
XLSX.utils.book_append_sheet(out, XLSX.utils.json_to_sheet(r1.result.accounting), 'Счетоводство');
XLSX.writeFile(out, 'sample_result.xlsx');
fs.writeFileSync('sample_status.txt', Core.statusText(r1.log));
console.log('\nSample outputs written. Status:', r1.log.status, '| warnings:', r1.log.warnings.length);
console.log(r1.log.info.join('\n'));
console.log('\n' + (fails ? fails + ' FAILURES' : 'ALL TESTS PASSED'));
process.exitCode = fails ? 1 : 0;
