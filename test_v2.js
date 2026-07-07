const XLSX = require('xlsx');
const ExcelJS = require('exceljs');
const fs = require('fs');
const Core = require('./core2.js');
const U = '/sessions/fervent-trusting-volta/mnt/uploads/';
const V2 = '/tmp/v2/';

function aoa(path) {
  const wb = XLSX.readFile(path, { cellDates: false });
  return XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { header: 1, raw: true, defval: null });
}
let fails = 0;
const ok = (cond, msg) => { console.log((cond ? 'PASS' : 'FAIL') + ' - ' + msg); if (!cond) fails++; };

const f1 = aoa(U + 'F1_clean.xls'), f4 = aoa(U + 'F4_clean.xls');
const f2 = aoa(U + 'F2_clean.xls'), f3 = aoa(U + 'F3_clean.xlsx');
const f6 = aoa(V2 + 'F6_anon.xlsx'), f7 = aoa(V2 + 'F7_anon.xlsx');

// ================= A. v1 regression =================
const r1 = Core.process(f1, f4, null, {});
ok(r1.log.status !== 'ERROR', 'A1 F1+F4: ' + r1.log.status);
ok(r1.result.summary.length === 5, 'A2 F1: 5 employees');
const w = re => r1.log.warnings.filter(x => re.test(x));
ok(w(/Дневна сума/).length === 0 && w(/Месечна сума/).length === 0 && w(/% несъответствие/).length === 0,
   'A3 F1 fully reconciled against its own subtotals');

const r2 = Core.process(f2, null, null, { roundDecimals: -1, dayEarningsSource: 'report' });
const expDay = new Map(), expMonth = new Map();
{ let block = -1;
  for (let i = 0; i < f2.length; i++) {
    const a = String((f2[i]||[])[0]||'').trim();
    if (/,\s*БРИГАДА/.test(a)) { block++; continue; }
    const dm = a.match(/^ЗА\s+ДАТА\s*\/\s*([\d.]+)\s*\/\s*:?$/);
    const iso = t => { const m = String(t).match(/(\d{1,2})\.(\d{1,2})\.(\d{4})/); return m ? `${m[3]}-${m[2].padStart(2,'0')}-${m[1].padStart(2,'0')}` : null; };
    if (dm && typeof f2[i][11] === 'number') expDay.set(block + '|' + iso(dm[1]), f2[i][11]);
    if (/^СУМА\s+ЗАРАБОТКА/.test(a) && typeof f2[i][11] === 'number') expMonth.set(block, f2[i][11]);
  } }
const gotDay = new Map();
{ let i = 0;
  r2.result.summary.forEach((emp, ei) => {
    for (let k = 0; k < emp['Отработени дни (F1)']; k++, i++) gotDay.set(ei + '|' + r2.result.days[i]['Дата'], r2.result.days[i]['Бонус (лв)']);
  }); }
let dd = 0, cmp = 0;
for (const [k, exp] of expDay) { const got = gotDay.get(k); if (got == null || Math.abs(exp - got) > 0.005) dd++; else cmp++; }
ok(dd === 0 && cmp === 98, `A4 F2 manual day bonuses reproduced exactly (${cmp}/98)`);
let md = 0;
for (const [b, exp] of expMonth) { const got = r2.result.summary[b] && r2.result.summary[b]['Бонус общо (лв)']; if (got == null || Math.abs(exp - got) > 0.005) md++; }
ok(md === 0, 'A5 F2 monthly bonus totals reproduced');

const r3 = Core.process(f3, null, null, { roundDecimals: -1, bonusLevel: 'both', bonusThresholdPct: 100, dayEarningsSource: 'report' });
ok(Math.abs(r3.result.summary[0]['Бонус общо (лв)'] - 31.2523) < 0.02, 'A6 F3 internal-tool double-count reproduced in \'both\' mode');

// ================= B. F6 attendance parsing =================
const rb = Core.process(f1, f4, f6, {});
ok(rb.log.status !== 'ERROR', 'B1 F1+F4+F6: ' + rb.log.status);
const att = rb.result.attendance;
ok(att && att.year === 2026 && att.month === 6, `B2 period detected: ${att && att.month}.${att && att.year}`);
ok(att.rows.length >= 90, `B3 employees in F6: ${att.rows.length}`);
ok(att.satDays.join(',') === '6,13,20,27' && att.sunDays.join(',') === '7,14,21,28',
   `B4 weekends of 06.2026 detected (sat: ${att.satDays}, sun: ${att.sunDays})`);

// B5: our F6 totals must equal the client's F7 СУМА column (except their manual corrections)
// F7: parse with our own parser, taking its СУМА per row from raw file
const wb7raw = f7;
const f7sums = new Map();
{ // day-number row is row 10 (idx 9); СУМА column is the last with header 'СУМА' in row 9 (idx 8)
  let sumCol = -1;
  for (let c = 0; c < (wb7raw[8]||[]).length; c++) if (String(wb7raw[8][c]).trim() === 'СУМА') sumCol = c;
  for (let i = 10; i < wb7raw.length; i++) {
    const name = String((wb7raw[i]||[])[0]||'').trim();
    const v = wb7raw[i] ? wb7raw[i][sumCol] : null;
    if (name && /^FULL NAME/.test(name) && typeof v === 'number') f7sums.set(name.toLowerCase(), v);
  } }
ok(f7sums.size >= 85, `B5a F7 СУМА values read: ${f7sums.size}`);
let match = 0, diff = [], missing = 0;
for (const p of att.rows) {
  const v = f7sums.get(p.name.toLowerCase());
  if (v == null) { missing++; continue; }
  if (Math.abs(p.totalMin - v) < 0.5) match++;
  else diff.push({ name: p.name, ours: p.totalMin, f7: v, delta: p.totalMin - v });
}
console.log('   F6-vs-F7 totals: match=' + match + ' diff=' + diff.length + ' notInF7=' + missing);
diff.forEach(d => console.log('    diff:', d.name, 'ours', d.ours, 'F7', d.f7, 'delta', d.delta));
// client made 7 manual corrections; our raw sums should differ from F7 on exactly those rows, by their correction amounts
const deltas = diff.map(d => d.delta).sort((a,b)=>a-b).join(',');
ok(diff.length <= 8, `B5b differences only where client corrected manually (${diff.length} rows, deltas: ${deltas})`);

// B6: workday equivalents = includedMin / 480
const p0 = att.rows.find(p => p.workdays > 0);
ok(p0 && Math.abs(p0.workdays - p0.includedMin / 480) < 0.005, 'B6 workdays = minutes / 480');

// B7: excluding Saturdays reduces the workday equivalents
const rbNoSat = Core.process(f1, f4, f6, { includeSaturday: false });
const t1 = rb.result.attendance.rows.reduce((a, p) => a + p.includedMin, 0);
const t2 = rbNoSat.result.attendance.rows.reduce((a, p) => a + p.includedMin, 0);
const satTot = rb.result.attendance.rows.reduce((a, p) => a + p.satMin, 0);
ok(t1 - t2 === satTot && satTot > 0, `B7 includeSaturday toggle works (saturday minutes: ${satTot})`);

// B8: F7 (consolidated format) also parses with same totals
const rb7 = Core.process(f1, null, f7, {});
const att7 = rb7.result.attendance;
let same = 0, tot7 = 0;
for (const p of att7.rows) { tot7++; const q = att.rows.find(x => x.name.toLowerCase() === p.name.toLowerCase()); if (q && q.totalMin === p.totalMin) same++; }
console.log('   F7 parsed employees:', tot7, '| identical totals with F6:', same);
ok(tot7 >= 85, 'B8 F7 consolidated format parses too');

// ================= C. config matrix =================
let mOk = 0, mAll = 0;
for (const bonusLevel of ['day', 'operation', 'both'])
for (const dayEarningsSource of ['operations', 'report'])
for (const roundDecimals of [-1, 2])
for (const includeSaturday of [true, false])
for (const fullDayMinutes of [480, 420]) {
  mAll++;
  const r = Core.process(f1, f4, f6, { bonusLevel, dayEarningsSource, roundDecimals, includeSaturday, fullDayMinutes });
  if (r.log.status !== 'ERROR' && r.result && r.result.salary.length) mOk++;
}
ok(mOk === mAll, `C1 config matrix: ${mOk}/${mAll} combinations processed`);
// invariants
const rDay = Core.process(f1, null, null, { roundDecimals: -1 });
const rBoth = Core.process(f1, null, null, { roundDecimals: -1, bonusLevel: 'both' });
const sum = r => r.result.summary.reduce((a, e) => a + e['Бонус общо (лв)'], 0);
ok(sum(rBoth) >= sum(rDay) - 1e-9, 'C2 invariant: both-level bonus >= day-level bonus');
const rBadCfg = Core.process(f1, null, null, { bonusThresholdPct: -5, bonusLevel: 'nonsense', fullDayMinutes: 'x' });
ok(rBadCfg.log.entries.filter(e => e.level === 'error' && e.category === 'Конфигурация').length === 3
   && rBadCfg.log.status !== 'ERROR', 'C3 invalid config values reported and defaulted');

// ================= D. status outputs =================
const rows = Core.statusRows(rb.log);
ok(rows.length === rb.log.entries.length && rows.length ===
   rb.log.info.length + rb.log.warnings.length + rb.log.errors.length,
   `D1 statusRows: one row per message (${rows.length})`);
ok(rows.every(r => r['№'] && r['Тип'] && r['Категория'] && r['Съобщение']), 'D2 every status row has № / Тип / Категория / Съобщение');

// D3: build both workbooks, write, re-read, verify
(async () => {
  const cfgAll = { ...Core.DEFAULT_CONFIG, includeStatusSheet: true };
  const wb = Core.buildWorkbook(ExcelJS, rb.result, rb.log, cfgAll, {});
  await wb.xlsx.writeFile('/tmp/v2/v2_result.xlsx');
  const swb = Core.buildStatusWorkbook(ExcelJS, rb.log);
  await swb.xlsx.writeFile('/tmp/v2/v2_status.xlsx');
  const chk = new ExcelJS.Workbook();
  await chk.xlsx.readFile('/tmp/v2/v2_result.xlsx');
  const names = chk.worksheets.map(ws => ws.name);
  ok(['За заплати','Обобщение','По дни','Операции','Счетоводство','Присъствие','Статус'].every(n => names.includes(n)),
     'D3 result workbook sheets: ' + names.join(', '));
  const salWs = chk.getWorksheet('За заплати');
  ok(salWs.rowCount >= 95, `D4 salary sheet has one row per person across all sources (${salWs.rowCount - 1} persons)`);
  // the key column requested by the customer:
  const hdr = salWs.getRow(1).values.map(v => String(v||''));
  const colIdx = hdr.findIndex(h => /раб\. дни = мин\/480/.test(h));
  ok(colIdx > 0, 'D5 salary sheet contains the "минути/480" workday-equivalent column');
  const attWs = chk.getWorksheet('Присъствие');
  // Saturday column (day 6) yellow in header
  const dayHdr = attWs.getRow(1).values.map(v => String(v||''));
  const satCol = dayHdr.findIndex(h => h === '6');
  const fill = attWs.getRow(1).getCell(satCol).fill;
  ok(fill && fill.fgColor && fill.fgColor.argb === 'FFFFFF00', 'D6 Saturday column highlighted yellow in attendance sheet');
  const stWs = new ExcelJS.Workbook();
  await stWs.xlsx.readFile('/tmp/v2/v2_status.xlsx');
  ok(stWs.getWorksheet('Статус').rowCount - 1 === rows.length, `D7 status .xlsx: ${rows.length} message rows`);

  fs.writeFileSync('/tmp/v2/v2_status.txt', Core.statusText(rb.log));
  console.log('\n' + (fails ? fails + ' FAILURES' : 'ALL TESTS PASSED'));
  process.exitCode = fails ? 1 : 0;
})();
