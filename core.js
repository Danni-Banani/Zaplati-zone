/* ===== Payroll-prep core v2 (pure functions, no DOM) =====
 * Works in browser and Node. Inputs: arrays-of-arrays from SheetJS.
 * v2: attendance (F6), structured status entries, config validation,
 *     ExcelJS workbook builder, salary sheet.
 */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory();
  else root.PayrollCore = factory();
})(typeof self !== 'undefined' ? self : this, function () {
  'use strict';

  const VERSION = '2.0.0';

  const DEFAULT_CONFIG = {
    // --- bonus rule ---
    bonusThresholdPct: 101,          // strictly greater than
    bonusRatePct: 10,
    bonusLevel: 'day',               // 'day' | 'operation' | 'both'
    dayEarningsSource: 'operations', // 'operations' | 'report'
    // --- output precision ---
    roundDecimals: -1,               // -1 = max precision
    attendanceDecimals: 2,           // decimals for workday-equivalent
    // --- reconciliation ---
    mismatchTolerancePct: 0.05,
    earningsTolerance: 0.02,
    scaleOverride: 'auto',           // 'auto' | '1' | '1000'
    // --- attendance (F6) ---
    fullDayMinutes: 480,             // divisor: minutes -> workday equivalents
    includeSaturday: true,           // include Saturday minutes in the total
    includeSunday: true,
    capDayMinutes: 0,                // 0 = no cap, else cap each day at this value
    maxDayMinutes: 720,              // warn above this per-day value (0 = off)
    checkAttendanceVsProduction: true,
    checkAttendanceVsAccounting: true,
    attendanceDayTolerance: 1,       // allowed |F4 worked days - attendance days|
    // --- workbook styling ---
    highlightWeekend: true,
    brigadeColors: true,
    includeStatusSheet: false
  };

  // limits used by validateConfig: [min, max]
  const CONFIG_LIMITS = {
    bonusThresholdPct: [0, 1000], bonusRatePct: [0, 100],
    roundDecimals: [-1, 6], attendanceDecimals: [0, 6],
    mismatchTolerancePct: [0, 100], earningsTolerance: [0, 100],
    fullDayMinutes: [1, 1440], capDayMinutes: [0, 1440], maxDayMinutes: [0, 1440],
    attendanceDayTolerance: [0, 31]
  };

  // ---------- helpers ----------
  const isNum = v => typeof v === 'number' && isFinite(v);
  const s = v => (v == null ? '' : String(v)).trim();

  function parseBgNumber(v) {
    if (isNum(v)) return v;
    let t = s(v).replace(/\s/g, '');
    if (!t) return null;
    if (/^-?\d{1,3}(\.\d{3})+,\d+$/.test(t)) t = t.replace(/\./g, '').replace(',', '.');
    else if (/,/.test(t) && !/\./.test(t)) t = t.replace(',', '.');
    if (!/^-?\d+(\.\d+)?$/.test(t)) return null;
    const n = Number(t);
    return isFinite(n) ? n : null;
  }

  function parsePct(v) {
    if (isNum(v)) return { value: v, hadSign: false };
    let t = s(v).replace(/\s/g, '');
    if (!t) return null;
    const hadSign = /%$/.test(t);
    t = t.replace(/%$/, '');
    const n = parseBgNumber(t);
    return n == null ? null : { value: n, hadSign };
  }

  const DATE_RE = /^(\d{1,2})\.(\d{1,2})\.(\d{4})(\s*г\.?)?$/;
  function parseBgDate(v) {
    const m = s(v).match(DATE_RE);
    if (!m) return null;
    const [, d, mo, y] = m;
    return `${y}-${mo.padStart(2, '0')}-${d.padStart(2, '0')}`;
  }

  function normName(name) { return s(name).replace(/\s+/g, ' ').toLowerCase(); }

  function round(v, dec) {
    if (dec == null || dec < 0 || v == null || !isNum(v)) return v;
    const f = Math.pow(10, dec);
    return Math.round((v + Number.EPSILON) * f) / f;
  }

  // ---------- structured logging ----------
  function newLog(cfg) {
    return { entries: [], info: [], warnings: [], errors: [], config: cfg,
             startedAt: new Date().toISOString(), version: VERSION };
  }
  function add(log, level, category, message) {
    log.entries.push({ level, category, message });
    const line = '[' + category + '] ' + message;
    if (level === 'error') log.errors.push(line);
    else if (level === 'warning') log.warnings.push(line);
    else log.info.push(line);
  }

  // ---------- config validation ----------
  function validateConfig(userCfg, log) {
    const cfg = { ...DEFAULT_CONFIG };
    for (const [k, v] of Object.entries(userCfg || {})) {
      if (!(k in DEFAULT_CONFIG)) { add(log, 'warning', 'Конфигурация', `Непозната настройка „${k}“ — игнорирана.`); continue; }
      cfg[k] = v;
    }
    for (const [k, lim] of Object.entries(CONFIG_LIMITS)) {
      let v = cfg[k];
      if (typeof v === 'string') v = parseBgNumber(v);
      if (!isNum(v)) { add(log, 'error', 'Конфигурация', `„${k}“ трябва да е число (получено: ${JSON.stringify(cfg[k])}).`); cfg[k] = DEFAULT_CONFIG[k]; continue; }
      if (v < lim[0] || v > lim[1]) { add(log, 'error', 'Конфигурация', `„${k}“=${v} е извън допустимия диапазон [${lim[0]}..${lim[1]}] — върната е стойността по подразбиране.`); v = DEFAULT_CONFIG[k]; }
      cfg[k] = v;
    }
    if (!['day', 'operation', 'both'].includes(cfg.bonusLevel)) { add(log, 'error', 'Конфигурация', `Невалидно ниво на бонус „${cfg.bonusLevel}“ — използвано е 'day'.`); cfg.bonusLevel = 'day'; }
    if (!['operations', 'report'].includes(cfg.dayEarningsSource)) { add(log, 'error', 'Конфигурация', `Невалиден източник „${cfg.dayEarningsSource}“ — използвано е 'operations'.`); cfg.dayEarningsSource = 'operations'; }
    if (!['auto', '1', '1000', 1, 1000].includes(cfg.scaleOverride)) { add(log, 'error', 'Конфигурация', `Невалиден мащаб „${cfg.scaleOverride}“ — използвано е 'auto'.`); cfg.scaleOverride = 'auto'; }
    if (cfg.bonusLevel === 'both') add(log, 'warning', 'Конфигурация', 'Ниво „и двете“ брои едно преизпълнение два пъти — използвайте само за сравнение с вътрешния инструмент.');
    if (cfg.capDayMinutes > 0 && cfg.capDayMinutes < cfg.fullDayMinutes) add(log, 'warning', 'Конфигурация', `Ограничението на ден (${cfg.capDayMinutes} мин) е под пълния работен ден (${cfg.fullDayMinutes} мин).`);
    return cfg;
  }

  // ---------- F1 / F3 (productivity report) parser ----------
  const EMP_RE = /^(.*),\s*(.*?),\s*БРИГАДА\s*:?\s*(\S+)\s*$/i;
  const DAYTOTAL_RE = /^ЗА\s+ДАТА\s*\/\s*(\d{1,2}\.\d{1,2}\.\d{4})\s*\/\s*:?\s*$/i;
  const MONTHTOTAL_RE = /^СУМА\s+ЗАРАБОТКА/i;
  const TECHCARD_RE = /^\S+\/\S+$/;

  function rowIsJunk(row) {
    for (const c of row) if (/^Стр\.\s*\d+$/.test(s(c))) return true;
    if (DATE_RE.test(s(row[0])) && /^\d{1,2}:\d{2}(:\d{2})?$/.test(s(row[1]))) return true;
    return false;
  }
  function rowIsEmpty(row) { return row.every(c => s(c) === ''); }

  function parseProductivity(aoa, log) {
    const meta = { title: null, period: null };
    const employees = [];
    const skipped = [];
    let cur = null, curDate = null, lastWasMonthTotal = true;
    const note = (i, reason, row) =>
      skipped.push({ row: i + 1, reason, preview: row.slice(0, 4).map(s).join(' | ').slice(0, 80) });

    for (let i = 0; i < aoa.length; i++) {
      const row = aoa[i] || [];
      const a = s(row[0]);
      if (rowIsEmpty(row)) continue;
      if (rowIsJunk(row)) { note(i, 'page-junk', row); continue; }
      if (/^АНАЛИЗ/i.test(a)) { meta.title = a; continue; }
      if (/^От\s*:/i.test(a)) { meta.period = a; continue; }
      if (/^ЗА\s+БРИГАДА/i.test(a)) continue;
      if (/^ТЕХНОЛ/i.test(a)) continue;

      const empM = a.match(EMP_RE);
      if (empM) {
        const [, name, code, brigade] = empM.map(x => s(x));
        const prev = cur;
        const samePerson = prev && !lastWasMonthTotal &&
          normName(prev.name) === normName(name) && s(prev.code) === code;
        if (!samePerson) {
          cur = { name, code, brigade, ops: [], dayTotals: {}, monthTotalReported: null, order: employees.length };
          employees.push(cur);
        }
        lastWasMonthTotal = false; curDate = null; continue;
      }

      const iso = parseBgDate(a);
      if (iso && row.slice(1).every(c => s(c) === '')) { curDate = iso; continue; }

      const dayM = a.match(DAYTOTAL_RE);
      if (dayM && cur) {
        const iso2 = parseBgDate(dayM[1]) || curDate;
        cur.dayTotals[iso2] = { reportPctRaw: row[9] != null ? row[9] : null, reportEarn: parseBgNumber(row[10]) };
        continue;
      }
      if (MONTHTOTAL_RE.test(a) && cur) {
        cur.monthTotalReported = parseBgNumber(row[10]);
        lastWasMonthTotal = true; continue;
      }
      if (TECHCARD_RE.test(a)) {
        if (!cur) { note(i, 'operation-row-outside-employee-block', row); continue; }
        if (!curDate) { note(i, 'operation-row-without-date', row); continue; }
        const qty = parseBgNumber(row[6]);
        const norm = parseBgNumber(row[8]);
        const pct = parsePct(row[9]);
        const earnRaw = parseBgNumber(row[10]);
        const earnScalable = isNum(row[10]);
        if (qty == null || norm == null) { note(i, 'unparseable-qty-or-norm', row); continue; }
        cur.ops.push({ date: curDate, techCard: a, opNo: s(row[2]), operation: s(row[3]),
          qty, norm, reportPct: pct, earnRaw, earnScalable, srcRow: i + 1 });
        continue;
      }
      note(i, 'unrecognized-row', row);
    }
    add(log, 'info', 'Производителност', `Разпознати ${employees.length} блока служители, ` +
      `${employees.reduce((n, e) => n + e.ops.length, 0)} операционни реда; пропуснати ${skipped.length} реда.`);
    if (meta.period) add(log, 'info', 'Производителност', `Период: ${meta.period}`);
    return { meta, employees, skipped };
  }

  function detectEarnScale(employees, cfg, log) {
    if (cfg.scaleOverride !== 'auto') {
      const sc = Number(cfg.scaleOverride);
      add(log, 'info', 'Мащаб', `Мащаб на заработката: зададен ръчно ×${sc}.`);
      return sc;
    }
    const ratios = [];
    for (const e of employees) {
      const byDate = {};
      for (const op of e.ops) {
        if (op.earnRaw == null) continue;
        (byDate[op.date] = byDate[op.date] || []).push(op);
      }
      for (const [d, list] of Object.entries(byDate)) {
        const rep = e.dayTotals[d] && e.dayTotals[d].reportEarn;
        const sumScal = list.filter(o => o.earnScalable).reduce((a, o) => a + o.earnRaw, 0);
        const sumFixed = list.filter(o => !o.earnScalable).reduce((a, o) => a + o.earnRaw, 0);
        const denom = rep != null ? rep - sumFixed : null;
        if (denom && denom > 0 && sumScal) ratios.push(sumScal / denom);
      }
    }
    if (!ratios.length) { add(log, 'warning', 'Мащаб', 'Няма данни за автоматично определяне на мащаба — приет ×1.'); return 1; }
    ratios.sort((a, b) => a - b);
    const med = ratios[Math.floor(ratios.length / 2)];
    const scale = med > 100 ? 1000 : 1;
    add(log, 'info', 'Мащаб', `Мащаб на заработката по операции: ${scale === 1000 ? '×1000 (стойностите се делят на 1000)' : 'нормален ×1'} (медианно отношение ≈ ${med.toFixed(2)}).`);
    return scale;
  }

  // ---------- productivity computation ----------
  function compute(parsed, cfg, log) {
    const scale = detectEarnScale(parsed.employees, cfg, log);
    const thr = cfg.bonusThresholdPct, rate = cfg.bonusRatePct / 100;
    const out = { employees: [], days: [], ops: [], scale };

    for (const e of parsed.employees) {
      const byDate = {};
      for (const op of e.ops) (byDate[op.date] = byDate[op.date] || []).push(op);
      let totEarn = 0, totBonus = 0;
      const dates = Object.keys(byDate).sort();

      for (const d of dates) {
        const ops = byDate[d];
        let dayPct = 0, dayEarnOps = 0, dayBonusOpLevel = 0;
        for (const op of ops) {
          const pct = (op.qty / op.norm) * 100;
          const earn = op.earnRaw != null ? (op.earnScalable ? op.earnRaw / scale : op.earnRaw) : null;
          dayPct += pct;
          if (earn != null) dayEarnOps += earn;
          if (op.reportPct) {
            const rv = op.reportPct.value;
            const candidates = op.reportPct.hadSign ? [rv] : [rv, rv * 100];
            const best = candidates.reduce((b, c) => Math.abs(c - pct) < Math.abs(b - pct) ? c : b);
            if (Math.abs(best - pct) > cfg.mismatchTolerancePct) {
              add(log, 'warning', 'Сверка', `% несъответствие: ${e.name} (${e.code}) ${d} карта ${op.techCard} оп.${op.opNo}: отчет ${best.toFixed(2)}% ≠ преизчислено ${pct.toFixed(2)}% (ред ${op.srcRow}).`);
            }
          }
          const opBonus = (earn != null && pct > thr) ? earn * rate : 0;
          dayBonusOpLevel += opBonus;
          out.ops.push({ name: e.name, code: e.code, brigade: e.brigade, date: d,
            techCard: op.techCard, opNo: op.opNo, operation: op.operation,
            qty: op.qty, norm: op.norm, pct, earn, opBonus });
        }
        const rep = e.dayTotals[d] || {};
        const dayEarn = (cfg.dayEarningsSource === 'report' && rep.reportEarn != null) ? rep.reportEarn : dayEarnOps;
        if (rep.reportEarn != null && Math.abs(dayEarnOps - rep.reportEarn) > cfg.earningsTolerance) {
          add(log, 'warning', 'Сверка', `Дневна сума несъответствие: ${e.name} (${e.code}) ${d}: операции ${dayEarnOps.toFixed(3)} лв ≠ отчет ${rep.reportEarn.toFixed(2)} лв.`);
        }
        const dayBonusDayLevel = dayPct > thr ? dayEarn * rate : 0;
        const bonus = cfg.bonusLevel === 'day' ? dayBonusDayLevel :
          cfg.bonusLevel === 'operation' ? dayBonusOpLevel : dayBonusDayLevel + dayBonusOpLevel;
        totEarn += dayEarn; totBonus += bonus;
        out.days.push({ name: e.name, code: e.code, brigade: e.brigade, date: d,
          dayPct, dayEarn, overNorm: dayPct > thr, bonus, reportEarn: rep.reportEarn != null ? rep.reportEarn : null });
      }
      if (e.monthTotalReported != null && Math.abs(totEarn - e.monthTotalReported) > cfg.earningsTolerance * Math.max(1, dates.length)) {
        add(log, 'warning', 'Сверка', `Месечна сума несъответствие: ${e.name} (${e.code}): изчислено ${totEarn.toFixed(2)} лв ≠ отчет ${e.monthTotalReported.toFixed(2)} лв.`);
      }
      out.employees.push({ name: e.name, code: e.code, brigade: e.brigade,
        daysWorked: dates.length, totalEarn: totEarn, totalBonus: totBonus,
        monthTotalReported: e.monthTotalReported, prodDates: dates });
    }
    return out;
  }

  // ---------- F4 (accounting, Crystal Reports) parser ----------
  function parseAccounting(aoa, log) {
    let headerIdx = -1, firstCol = -1;
    for (let i = 0; i < Math.min(aoa.length, 30) && headerIdx < 0; i++) {
      const row = aoa[i] || [];
      for (let c = 0; c < row.length; c++) if (s(row[c]) === 'Име') { headerIdx = i; firstCol = c; break; }
    }
    if (headerIdx < 0) throw new Error('Счетоводен файл: не е намерен заглавен ред с колона „Име“.');
    const header = aoa[headerIdx].map(s);
    let lastCol = firstCol;
    for (let c = firstCol; c < header.length; c++) if (header[c]) lastCol = c;
    const rows = [], totals = [];
    for (let i = headerIdx + 1; i < aoa.length; i++) {
      const row = aoa[i] || [];
      const name = s(row[firstCol]);
      const hasNums = row.slice(firstCol + 1, lastCol + 1).some(v => parseBgNumber(v) != null);
      if (!name && !hasNums) continue;
      const rec = {};
      for (let c = firstCol + 1; c <= lastCol; c++) {
        if (!header[c]) continue;
        const n = parseBgNumber(row[c]);
        rec[header[c]] = n != null ? n : (s(row[c]) || null);
      }
      if (name) rows.push({ name, ...rec }); else totals.push(rec);
    }
    const numCols = Object.keys(rows[0] || {}).filter(k => k !== 'name' && rows.some(r => isNum(r[k])));
    if (totals.length) {
      for (const k of numCols) {
        const grand = totals.reduce((a, t) => a + (isNum(t[k]) ? t[k] : 0), 0);
        const ours = rows.reduce((a, r) => a + (isNum(r[k]) ? r[k] : 0), 0);
        if (grand !== 0 && Math.abs(ours - grand) > 0.05) {
          add(log, 'warning', 'Счетоводство', `Контролна сума „${k}“: изчислено ${ours.toFixed(2)} ≠ отчет ${grand.toFixed(2)}.`);
        }
      }
    }
    add(log, 'info', 'Счетоводство', `${rows.length} служители; ${totals.length} тотални реда за сверка.`);
    return { header: header.slice(firstCol, lastCol + 1), rows, totals };
  }

  // ---------- F6 (attendance) parser ----------
  function parseAttendance(aoa, log) {
    const people = new Map();
    let year = null, month = null, brigadeLabel = null;
    let colDay = null, collecting = false, expectingDays = false, sections = 0;
    const MARKER = /НОРМА МЕНИДЖЪР|ПРИСЪСТВЕНА ФОРМА|^Печат$/;

    for (let i = 0; i < aoa.length; i++) {
      const row = aoa[i] || [];
      const a = s(row[0]);
      const joined = row.map(s).filter(Boolean).join(' ');
      if (!joined) continue;

      if (/^Период/.test(a)) {
        const m = joined.match(/(\d{1,2})\.(\d{1,2})\.(\d{4})/);
        if (m) { month = +m[2]; year = +m[3]; }
        continue;
      }
      if (/^Бригада/.test(a)) { brigadeLabel = row.slice(1).map(s).filter(Boolean)[0] || brigadeLabel; continue; }
      if (row.some(c => /ДЕН В МЕСЕЦА/.test(s(c)))) { expectingDays = true; collecting = false; continue; }

      if (expectingDays && !a) {
        const days = {}; let cnt = 0, last = 0, ascending = true;
        row.forEach((c, ci) => {
          const n = parseBgNumber(c);
          if (n != null && Number.isInteger(n) && n >= 1 && n <= 31) {
            if (n <= last) ascending = false;
            days[ci] = n; last = n; cnt++;
          }
        });
        if (cnt >= 2 && ascending) { colDay = days; collecting = true; expectingDays = false; sections++; continue; }
        continue;
      }
      if (MARKER.test(a)) { collecting = false; continue; }
      if (/^СУМА/i.test(a)) { collecting = false; continue; }

      if (collecting && a && colDay) {
        const key = normName(a);
        let rec = people.get(key);
        if (!rec) { rec = { name: a, brigade: brigadeLabel, days: {} }; people.set(key, rec); }
        for (const [ci, d] of Object.entries(colDay)) {
          const v = parseBgNumber(row[ci]);
          if (v == null) continue;
          if (rec.days[d] != null && rec.days[d] !== v) {
            add(log, 'warning', 'Присъствие', `Дублирани данни: ${rec.name}, ден ${d}: ${rec.days[d]} мин и ${v} мин — използвана е първата стойност.`);
          } else if (rec.days[d] == null) rec.days[d] = v;
        }
      }
    }
    if (!people.size) throw new Error('Файлът с присъствия не съдържа разпознаваеми данни (търси се секция „ДЕН В МЕСЕЦА“ с ред от номера на дни).');
    if (!year || !month) add(log, 'warning', 'Присъствие', 'Периодът (месец/година) не бе разпознат — почивните дни няма да бъдат разграничени.');
    add(log, 'info', 'Присъствие', `Разпознати ${people.size} служители в ${sections} секции${year ? `, период ${String(month).padStart(2, '0')}.${year}` : ''}.`);
    return { year, month, people: [...people.values()] };
  }

  function computeAttendance(att, cfg, log) {
    const { year, month } = att;
    const dow = d => (year && month) ? new Date(year, month - 1, d).getDay() : -1;
    const rows = att.people.map(p => {
      let wk = 0, sat = 0, sun = 0, daysPresent = 0, satDays = 0, capped = 0;
      for (const [dStr, vRaw] of Object.entries(p.days)) {
        const d = +dStr; let v = vRaw;
        if (cfg.maxDayMinutes > 0 && v > cfg.maxDayMinutes) {
          add(log, 'warning', 'Присъствие', `${p.name}: ден ${d} има ${v} мин (над прага ${cfg.maxDayMinutes} мин) — проверете стойността.`);
        }
        if (cfg.capDayMinutes > 0 && v > cfg.capDayMinutes) { v = cfg.capDayMinutes; capped++; }
        const w = dow(d);
        if (w === 6) { sat += v; if (v > 0) satDays++; }
        else if (w === 0) sun += v;
        else wk += v;
        if (v > 0) daysPresent++;
      }
      if (capped) add(log, 'info', 'Присъствие', `${p.name}: ${capped} дни са ограничени до ${cfg.capDayMinutes} мин.`);
      const includedMin = wk + (cfg.includeSaturday ? sat : 0) + (cfg.includeSunday ? sun : 0);
      return { ...p, totalMin: wk + sat + sun, includedMin, satMin: sat, sunMin: sun,
               daysPresent, satDays, workdays: includedMin / cfg.fullDayMinutes };
    });
    return { year, month, rows };
  }

  // ---------- join + result assembly ----------
  function buildResult(prod, acc, attC, cfg, log) {
    const r = cfg.roundDecimals, ad = cfg.attendanceDecimals;
    const accByName = new Map();
    for (const a of acc.rows) {
      const k = normName(a.name);
      if (accByName.has(k)) add(log, 'warning', 'Съпоставяне', `Счетоводство: дублирано име „${a.name}“ — съпоставянето по име е ненадеждно за него.`);
      accByName.set(k, a);
    }
    const attByName = new Map();
    if (attC) for (const p of attC.rows) {
      const k = normName(p.name);
      if (attByName.has(k)) add(log, 'warning', 'Съпоставяне', `Присъствие: дублирано име „${p.name}“.`);
      attByName.set(k, p);
    }

    const matchedAcc = new Set(), matchedAtt = new Set();

    // per-day attendance vs production checks (only for matched employees)
    if (attC && cfg.checkAttendanceVsProduction) {
      for (const e of prod.employees) {
        const p = attByName.get(normName(e.name));
        if (!p) continue;
        for (const dIso of e.prodDates) {
          const dayNum = +dIso.slice(8, 10);
          const dayRec = prod.days.find(x => normName(x.name) === normName(e.name) && x.date === dIso);
          const min = p.days[dayNum];
          if (dayRec && dayRec.dayEarn > 0 && (!min || min <= 0)) {
            add(log, 'warning', 'Присъствие', `${e.name}: има заработка на ${dIso} (${dayRec.dayEarn.toFixed(2)} лв), но няма присъствие за ден ${dayNum}.`);
          }
        }
      }
    }

    const summary = prod.employees.map(e => {
      const k = normName(e.name);
      const a = accByName.get(k) || null;
      const p = attByName.get(k) || null;
      if (a) matchedAcc.add(k);
      else if (acc.rows.length) add(log, 'warning', 'Съпоставяне', `„${e.name}“ (${e.code}) има производителност, но липсва в счетоводния файл.`);
      if (p) matchedAtt.add(k);
      else if (attByName.size) add(log, 'warning', 'Съпоставяне', `„${e.name}“ (${e.code}) има производителност, но липсва във файла с присъствия.`);
      const base = {
        'Име': e.name, 'Код': e.code, 'Бригада': e.brigade,
        'Отработени дни (F1)': e.daysWorked,
        'Заработка общо (лв)': round(e.totalEarn, r),
        'Бонус общо (лв)': round(e.totalBonus, r)
      };
      if (p) { base['Присъствие (мин)'] = p.includedMin; base['Присъствие (раб. дни)'] = round(p.workdays, ad); }
      if (a) for (const [kk, v] of Object.entries(a)) { if (kk !== 'name') base[kk] = isNum(v) ? round(v, r) : v; }
      else if (acc.rows.length) base['F4'] = 'НЕНАМЕРЕН';
      return base;
    });

    for (const a of acc.rows) if (!matchedAcc.has(normName(a.name)) && prod.employees.length)
      add(log, 'warning', 'Съпоставяне', `„${a.name}“ е в счетоводния файл, но няма данни за производителност.`);

    // accounting worked-days vs attendance days
    if (attC && cfg.checkAttendanceVsAccounting && acc.rows.length) {
      const wdKey = Object.keys(acc.rows[0] || {}).find(kk => /Отработени дни/i.test(kk));
      if (wdKey) {
        for (const a of acc.rows) {
          const p = attByName.get(normName(a.name));
          if (!p || !isNum(a[wdKey])) continue;
          if (Math.abs(p.daysPresent - a[wdKey]) > cfg.attendanceDayTolerance) {
            add(log, 'warning', 'Присъствие', `${a.name}: присъствени дни ${p.daysPresent} ≠ отработени дни по ведомост ${a[wdKey]} (толеранс ${cfg.attendanceDayTolerance}).`);
          }
        }
      } else add(log, 'info', 'Присъствие', 'Счетоводният файл няма колона „Отработени дни“ — проверката дни/присъствие е пропусната.');
    }

    // ---- salary sheet: union of all sources, one row per person ----
    const keys = new Map(); // normName -> display name
    for (const e of prod.employees) keys.set(normName(e.name), e.name);
    for (const a of acc.rows) if (!keys.has(normName(a.name))) keys.set(normName(a.name), a.name);
    if (attC) for (const p of attC.rows) if (!keys.has(normName(p.name))) keys.set(normName(p.name), p.name);

    const prodByName = new Map(prod.employees.map(e => [normName(e.name), e]));
    const salary = [];
    for (const [k, disp] of keys) {
      const e = prodByName.get(k), a = accByName.get(k), p = attByName.get(k);
      const checks = [];
      if (!e && prod.employees.length) checks.push('няма производителност');
      if (!a && acc.rows.length) checks.push('няма ведомост');
      if (!p && attByName.size) checks.push('няма присъствие');
      const row = {
        'Име': disp,
        'Код': e ? e.code : null,
        'Бригада': (e && e.brigade) || (p && p.brigade) || null,
        'Заработка (лв)': e ? round(e.totalEarn, r) : null,
        'Бонус (лв)': e ? round(e.totalBonus, r) : null,
        'Заработка + бонус (лв)': e ? round(e.totalEarn + e.totalBonus, r) : null,
        'Присъствие общо (мин)': p ? p.includedMin : null,
        ['Присъствие (раб. дни = мин/' + cfg.fullDayMinutes + ')']: p ? round(p.workdays, ad) : null,
        'Присъствени дни (брой)': p ? p.daysPresent : null,
        'Събота (мин)': p ? p.satMin : null,
        'Неделя (мин)': p ? p.sunMin : null
      };
      if (a) for (const [kk, v] of Object.entries(a)) { if (kk !== 'name') row['F4: ' + kk] = isNum(v) ? round(v, r) : v; }
      row['Проверки'] = checks.join('; ') || 'ОК';
      salary.push(row);
    }

    const days = prod.days.map(d => ({
      'Име': d.name, 'Код': d.code, 'Бригада': d.brigade, 'Дата': d.date,
      '% от норма (ден)': round(d.dayPct, Math.max(r, 2)),
      'Заработка (лв)': round(d.dayEarn, r),
      'Над прага': d.overNorm ? 'ДА' : '',
      'Бонус (лв)': round(d.bonus, r),
      'Заработка по отчет (лв)': d.reportEarn
    }));
    const ops = prod.ops.map(o => ({
      'Име': o.name, 'Код': o.code, 'Дата': o.date, 'Технол. карта': o.techCard,
      'Оп. №': o.opNo, 'Операция': o.operation, 'Брой': o.qty, 'Норма': o.norm,
      '% от норма': round(o.pct, Math.max(r, 2)), 'Заработка (лв)': round(o.earn, Math.max(r, 3)),
      'Бонус на ниво операция (лв)': cfg.bonusLevel !== 'day' ? round(o.opBonus, r) : undefined
    }));
    const accounting = acc.rows.map(a => {
      const o = { 'Име': a.name };
      for (const [kk, v] of Object.entries(a)) if (kk !== 'name') o[kk] = v;
      return o;
    });

    // attendance matrix (for the styled sheet)
    let attendance = null;
    if (attC) {
      const dim = (attC.year && attC.month) ? new Date(attC.year, attC.month, 0).getDate() : 31;
      const daysArr = []; const satCols = [], sunCols = [];
      for (let d = 1; d <= dim; d++) {
        daysArr.push(d);
        if (attC.year && attC.month) {
          const w = new Date(attC.year, attC.month - 1, d).getDay();
          if (w === 6) satCols.push(d); else if (w === 0) sunCols.push(d);
        }
      }
      attendance = {
        year: attC.year, month: attC.month, days: daysArr, satDays: satCols, sunDays: sunCols,
        rows: attC.rows.map(p => ({ name: p.name, brigade: p.brigade, days: p.days,
          totalMin: p.totalMin, includedMin: p.includedMin, satMin: p.satMin, sunMin: p.sunMin,
          daysPresent: p.daysPresent, workdays: round(p.workdays, ad) }))
      };
    }
    return { summary, salary, days, ops, accounting, attendance };
  }

  // ---------- status renderers ----------
  function statusRows(log) {
    const lvl = { error: 'ГРЕШКА', warning: 'ПРЕДУПРЕЖДЕНИЕ', info: 'ИНФОРМАЦИЯ' };
    return log.entries.map((e, i) => ({ '№': i + 1, 'Тип': lvl[e.level] || e.level, 'Категория': e.category, 'Съобщение': e.message }));
  }
  function statusText(log) {
    const L = [];
    L.push(`=== СТАТУС: ${log.status} === (v${log.version})`);
    L.push(`Начало: ${log.startedAt}  Край: ${log.finishedAt}`);
    L.push('', '--- Конфигурация ---');
    for (const [k, v] of Object.entries(log.config)) L.push(`${k}: ${v}`);
    L.push('', `--- Информация (${log.info.length}) ---`, ...log.info);
    L.push('', `--- Предупреждения (${log.warnings.length}) ---`, ...(log.warnings.length ? log.warnings : ['няма']));
    L.push('', `--- Грешки (${log.errors.length}) ---`, ...(log.errors.length ? log.errors : ['няма']));
    return L.join('\n');
  }

  // ---------- ExcelJS workbook builder (shared by browser UI and Node tests) ----------
  const BRIGADE_PALETTE = ['FFC00000', 'FF0066CC', 'FF008080', 'FF7030A0', 'FF006600', 'FF806000', 'FFCC00CC', 'FF444444'];

  function addTableSheet(wb, title, rows) {
    const ws = wb.addWorksheet(title);
    if (!rows || !rows.length) { ws.getCell('A1').value = '(няма данни)'; return ws; }
    const cols = [];
    for (const row of rows) for (const k of Object.keys(row)) if (!cols.includes(k) && row[k] !== undefined) cols.push(k);
    ws.addRow(cols).font = { bold: true };
    ws.getRow(1).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF1F5F9' } };
    for (const row of rows) ws.addRow(cols.map(c => row[c] === undefined ? null : row[c]));
    ws.columns.forEach((col, i) => {
      const header = String(cols[i] || '');
      let w = Math.min(Math.max(header.length + 2, 10), 44);
      col.width = w;
    });
    ws.views = [{ state: 'frozen', ySplit: 1 }];
    return ws;
  }

  function buildWorkbook(ExcelJS, data, log, cfg, sheets) {
    sheets = sheets || {};
    const on = k => sheets[k] !== false;
    const wb = new ExcelJS.Workbook();
    wb.creator = 'Zaplati-zone v' + VERSION;

    if (on('salary')) addTableSheet(wb, 'За заплати', data.salary);
    if (on('summary')) addTableSheet(wb, 'Обобщение', data.summary);
    if (on('days')) addTableSheet(wb, 'По дни', data.days);
    if (on('ops')) addTableSheet(wb, 'Операции', data.ops);
    if (on('accounting') && data.accounting.length) addTableSheet(wb, 'Счетоводство', data.accounting);

    if (on('attendance') && data.attendance) {
      const A = data.attendance;
      const ws = wb.addWorksheet('Присъствие');
      const header = ['Име', 'Бригада', ...A.days.map(d => String(d)), 'Общо (мин)', 'Раб. дни'];
      ws.addRow(header).font = { bold: true };
      ws.getRow(1).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF1F5F9' } };
      const brigades = [...new Set(A.rows.map(x => x.brigade).filter(Boolean))];
      for (const p of A.rows) {
        const row = ws.addRow([p.name, p.brigade, ...A.days.map(d => p.days[d] != null ? p.days[d] : null), p.includedMin, p.workdays]);
        if (cfg.brigadeColors && p.brigade) {
          const bi = brigades.indexOf(p.brigade);
          if (bi >= 0) row.font = { color: { argb: BRIGADE_PALETTE[bi % BRIGADE_PALETTE.length] } };
        }
      }
      // totals row
      const totals = ['ОБЩО', null, ...A.days.map(d => A.rows.reduce((a, p) => a + (p.days[d] || 0), 0) || null),
        A.rows.reduce((a, p) => a + p.includedMin, 0),
        Math.round(A.rows.reduce((a, p) => a + (p.workdays || 0), 0) * 100) / 100];
      ws.addRow(totals).font = { bold: true };
      if (cfg.highlightWeekend) {
        A.days.forEach((d, i) => {
          const colIdx = 3 + i;
          const isSat = A.satDays.includes(d), isSun = A.sunDays.includes(d);
          if (!isSat && !isSun) return;
          const argb = isSat ? 'FFFFFF00' : 'FFD9D9D9';
          for (let rIdx = 1; rIdx <= ws.rowCount; rIdx++) {
            ws.getRow(rIdx).getCell(colIdx).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb } };
          }
        });
      }
      ws.getColumn(1).width = 28; ws.getColumn(2).width = 14;
      for (let i = 0; i < A.days.length; i++) ws.getColumn(3 + i).width = 6;
      ws.getColumn(3 + A.days.length).width = 12; ws.getColumn(4 + A.days.length).width = 10;
      ws.views = [{ state: 'frozen', xSplit: 1, ySplit: 1 }];
    }

    if (cfg.includeStatusSheet || sheets.status) addTableSheet(wb, 'Статус', statusRows(log));
    return wb;
  }

  function buildStatusWorkbook(ExcelJS, log) {
    const wb = new ExcelJS.Workbook();
    wb.creator = 'Zaplati-zone v' + VERSION;
    addTableSheet(wb, 'Статус', statusRows(log));
    addTableSheet(wb, 'Конфигурация', Object.entries(log.config).map(([k, v]) => ({ 'Настройка': k, 'Стойност': String(v) })));
    return wb;
  }

  // ---------- top-level ----------
  function process(prodAoa, accAoa, attAoa, userCfg) {
    const log = newLog(null);
    const cfg = validateConfig(userCfg, log);
    log.config = cfg;
    let result = null;
    try {
      const parsedProd = parseProductivity(prodAoa, log);
      if (!parsedProd.employees.length) throw new Error('Файлът с производителност не съдържа разпознаваеми блокове служители.');
      for (const sk of parsedProd.skipped) {
        if (sk.reason !== 'page-junk') add(log, 'warning', 'Производителност', `Пропуснат ред ${sk.row} (${sk.reason}): ${sk.preview}`);
      }
      add(log, 'info', 'Производителност', `Премахнати технически редове от страниране: ${parsedProd.skipped.filter(x => x.reason === 'page-junk').length}.`);

      const computed = compute(parsedProd, cfg, log);
      const acc = accAoa ? parseAccounting(accAoa, log) : { header: [], rows: [], totals: [] };
      const attC = attAoa ? computeAttendance(parseAttendance(attAoa, log), cfg, log) : null;
      result = buildResult(computed, acc, attC, cfg, log);
      log.status = log.errors.length ? 'SUCCESS_WITH_ERRORS' :
                   log.warnings.length ? 'SUCCESS_WITH_WARNINGS' : 'SUCCESS';
    } catch (err) {
      add(log, 'error', 'Система', String(err && err.message || err));
      log.status = 'ERROR';
    }
    log.finishedAt = new Date().toISOString();
    return { result, log };
  }

  return { VERSION, DEFAULT_CONFIG, process, statusText, statusRows,
           buildWorkbook, buildStatusWorkbook, parseBgNumber, parsePct, parseBgDate };
});
