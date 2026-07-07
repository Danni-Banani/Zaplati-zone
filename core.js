/* ===== Payroll-prep core (pure functions, no DOM) =====
 * Works in browser and Node. Input: arrays-of-arrays from SheetJS.
 * All row/col indices are 0-based.
 */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory();
  else root.PayrollCore = factory();
})(typeof self !== 'undefined' ? self : this, function () {
  'use strict';

  const DEFAULT_CONFIG = {
    bonusThresholdPct: 101,      // strictly greater than
    bonusRatePct: 10,
    bonusLevel: 'day',           // 'day' | 'operation' | 'both'
    roundDecimals: 2,            // -1 = no rounding (max precision)
    dayEarningsSource: 'operations', // 'operations' (precise) | 'report'
    mismatchTolerancePct: 0.05,  // tolerance when comparing recomputed vs report %
    earningsTolerance: 0.02      // лв tolerance for earnings reconciliation
  };

  // ---------- helpers ----------
  const isNum = v => typeof v === 'number' && isFinite(v);
  const s = v => (v == null ? '' : String(v)).trim();

  function parseBgNumber(v) {
    // "41,49" | "772.78" | 41.49 -> Number, null if not numeric
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
    // "96,55%" -> 96.55 ; numeric left as-is (scale resolved by caller)
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
    const [ , d, mo, y ] = m;
    return `${y}-${mo.padStart(2, '0')}-${d.padStart(2, '0')}`; // ISO for sorting
  }

  function normName(name) {
    return s(name).replace(/\s+/g, ' ').toLowerCase();
  }

  function round(v, dec) {
    if (dec == null || dec < 0 || v == null) return v;
    const f = Math.pow(10, dec);
    return Math.round((v + Number.EPSILON) * f) / f;
  }

  // ---------- F1 / F3 (productivity report) parser ----------
  const EMP_RE = /^(.*),\s*(.*?),\s*БРИГАДА\s*:?\s*(\S+)\s*$/i;
  const DAYTOTAL_RE = /^ЗА\s+ДАТА\s*\/\s*(\d{1,2}\.\d{1,2}\.\d{4})\s*\/\s*:?\s*$/i;
  const MONTHTOTAL_RE = /^СУМА\s+ЗАРАБОТКА/i;
  const TECHCARD_RE = /^\S+\/\S+$/;   // e.g. 00000866/0526, 1ОБЩИ/ОБЩИ (hourly/general work)

  function rowIsJunk(row) {
    // page footers/headers: any cell containing "Стр." or a lone timestamp pair
    for (const c of row) if (/^Стр\.\s*\d+$/.test(s(c))) return true;
    if (DATE_RE.test(s(row[0])) && /^\d{1,2}:\d{2}(:\d{2})?$/.test(s(row[1]))) return true;
    return false;
  }
  function rowIsEmpty(row) { return row.every(c => s(c) === ''); }

  function parseProductivity(aoa, log) {
    const meta = { title: null, period: null };
    const employees = [];           // {name, code, brigade, blocks:[...]}
    const skipped = [];             // {row, reason, preview}
    let cur = null;                 // current employee
    let curDate = null;
    let lastWasMonthTotal = true;   // header after month-total => new employee

    const note = (i, reason, row) =>
      skipped.push({ row: i + 1, reason, preview: row.slice(0, 4).map(s).join(' | ').slice(0, 80) });

    for (let i = 0; i < aoa.length; i++) {
      const row = aoa[i] || [];
      const a = s(row[0]);

      if (rowIsEmpty(row)) continue;
      if (rowIsJunk(row)) { note(i, 'page-junk', row); continue; }
      if (/^АНАЛИЗ/i.test(a)) { meta.title = a; continue; }
      if (/^От\s*:/i.test(a)) { meta.period = a; continue; }
      if (/^ЗА\s+БРИГАДА/i.test(a)) continue;                      // repeated page header
      if (/^ТЕХНОЛ/i.test(a)) continue;                            // column header

      const empM = a.match(EMP_RE);
      if (empM) {
        const [ , name, code, brigade ] = empM.map(x => s(x));
        const prev = cur;
        const samePerson = prev && !lastWasMonthTotal &&
          normName(prev.name) === normName(name) && s(prev.code) === code;
        if (samePerson) { /* continuation after page break */ }
        else {
          cur = { name, code, brigade, ops: [], dayTotals: {}, monthTotalReported: null, order: employees.length };
          employees.push(cur);
        }
        lastWasMonthTotal = false;
        curDate = null;
        continue;
      }

      const iso = parseBgDate(a);
      if (iso && row.slice(1).every(c => s(c) === '')) {           // pure date row
        curDate = iso; continue;
      }

      const dayM = a.match(DAYTOTAL_RE);
      if (dayM && cur) {
        const iso2 = parseBgDate(dayM[1]) || curDate;
        cur.dayTotals[iso2] = {
          reportPctRaw: row[9] != null ? row[9] : null,
          reportEarn: parseBgNumber(row[10])
        };
        continue;
      }

      if (MONTHTOTAL_RE.test(a) && cur) {
        cur.monthTotalReported = parseBgNumber(row[10]);
        lastWasMonthTotal = true;
        continue;
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
        cur.ops.push({
          date: curDate, techCard: a, opNo: s(row[2]), operation: s(row[3]),
          qty, norm, reportPct: pct, earnRaw, earnScalable, srcRow: i + 1
        });
        continue;
      }

      note(i, 'unrecognized-row', row);
    }

    log.info.push(`Продуктивност: разпознати ${employees.length} блока служители, ` +
      `${employees.reduce((n, e) => n + e.ops.length, 0)} операционни реда; пропуснати ${skipped.length} реда.`);
    return { meta, employees, skipped };
  }

  // Detect earnings scale: report may store op earnings ×1000 ("33079" = 33.079 лв)
  function detectEarnScale(employees, log) {
    const ratios = [];
    for (const e of employees) {
      const byDate = {};
      for (const op of e.ops) {
        if (op.earnRaw == null) continue;
        (byDate[op.date] = byDate[op.date] || []).push(op);
      }
      for (const [d, list] of Object.entries(byDate)) {
        const rep = e.dayTotals[d] && e.dayTotals[d].reportEarn;
        const sumScal = list.filter(o=>o.earnScalable).reduce((a,o)=>a+o.earnRaw,0);
        const sumFixed = list.filter(o=>!o.earnScalable).reduce((a,o)=>a+o.earnRaw,0);
        const denom = rep != null ? rep - sumFixed : null;
        if (denom && denom > 0 && sumScal) ratios.push(sumScal / denom);
      }
    }
    if (!ratios.length) { log.warnings.push('Мащаб на заработката: няма данни за автоматично определяне — приет 1.'); return 1; }
    ratios.sort((a, b) => a - b);
    const med = ratios[Math.floor(ratios.length / 2)];
    const scale = med > 100 ? 1000 : 1;
    log.info.push(`Мащаб на заработката по операции: ${scale === 1000 ? '×1000 (стойностите се делят на 1000)' : 'нормален'} ` +
      `(медианно отношение операции/дневна сума ≈ ${med.toFixed(2)}).`);
    return scale;
  }

  // ---------- computation ----------
  function compute(parsed, cfg, log) {
    const scale = detectEarnScale(parsed.employees, log);
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

          // cross-check report's op % (resolve scale text/fraction/percent)
          if (op.reportPct) {
            const rv = op.reportPct.value;
            const candidates = op.reportPct.hadSign ? [rv] : [rv, rv * 100];
            const best = candidates.reduce((b, c) => Math.abs(c - pct) < Math.abs(b - pct) ? c : b);
            if (Math.abs(best - pct) > cfg.mismatchTolerancePct) {
              log.warnings.push(`% несъответствие: ${e.name} (${e.code}) ${d} карта ${op.techCard} оп.${op.opNo}: ` +
                `отчет ${best.toFixed(2)}% ≠ преизчислено ${pct.toFixed(2)}% (ред ${op.srcRow}).`);
            }
          }
          const opBonus = (earn != null && pct > thr) ? earn * rate : 0;
          dayBonusOpLevel += opBonus;
          out.ops.push({
            name: e.name, code: e.code, brigade: e.brigade, date: d,
            techCard: op.techCard, opNo: op.opNo, operation: op.operation,
            qty: op.qty, norm: op.norm, pct, earn, opBonus
          });
        }

        const rep = e.dayTotals[d] || {};
        const dayEarn = (cfg.dayEarningsSource === 'report' && rep.reportEarn != null)
          ? rep.reportEarn : dayEarnOps;

        if (rep.reportEarn != null && Math.abs(dayEarnOps - rep.reportEarn) > cfg.earningsTolerance) {
          log.warnings.push(`Дневна сума несъответствие: ${e.name} (${e.code}) ${d}: ` +
            `операции ${dayEarnOps.toFixed(3)} лв ≠ отчет ${rep.reportEarn.toFixed(2)} лв.`);
        }

        const dayBonusDayLevel = dayPct > thr ? dayEarn * rate : 0;
        const bonus =
          cfg.bonusLevel === 'day' ? dayBonusDayLevel :
          cfg.bonusLevel === 'operation' ? dayBonusOpLevel :
          dayBonusDayLevel + dayBonusOpLevel;   // 'both' – replicates internal tool F3 (double pay!)

        totEarn += dayEarn; totBonus += bonus;
        out.days.push({
          name: e.name, code: e.code, brigade: e.brigade, date: d,
          dayPct, dayEarn, overNorm: dayPct > thr, bonus,
          reportEarn: rep.reportEarn != null ? rep.reportEarn : null
        });
      }

      if (e.monthTotalReported != null && Math.abs(totEarn - e.monthTotalReported) > cfg.earningsTolerance * Math.max(1, dates.length)) {
        log.warnings.push(`Месечна сума несъответствие: ${e.name} (${e.code}): ` +
          `изчислено ${totEarn.toFixed(2)} лв ≠ отчет ${e.monthTotalReported.toFixed(2)} лв.`);
      }

      out.employees.push({
        name: e.name, code: e.code, brigade: e.brigade,
        daysWorked: dates.length, totalEarn: totEarn, totalBonus: totBonus,
        monthTotalReported: e.monthTotalReported
      });
    }
    return out;
  }

  // ---------- F4 (accounting, Crystal Reports) parser ----------
  function parseAccounting(aoa, log) {
    let headerIdx = -1, firstCol = -1;
    for (let i = 0; i < Math.min(aoa.length, 30) && headerIdx < 0; i++) {
      const row = aoa[i] || [];
      for (let c = 0; c < row.length; c++) {
        if (s(row[c]) === 'Име') { headerIdx = i; firstCol = c; break; }
      }
    }
    if (headerIdx < 0) throw new Error('F4: не е намерен заглавен ред с колона „Име“.');

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
      if (name) rows.push({ name, ...rec });
      else totals.push(rec);                     // report's own total rows -> reconciliation
    }

    // reconcile our column sums vs report totals
    const numCols = Object.keys(rows[0] || {}).filter(k => k !== 'name' && rows.some(r => isNum(r[k])));
    if (totals.length) {
      const grand = {};
      for (const k of numCols) grand[k] = totals.reduce((a, t) => a + (isNum(t[k]) ? t[k] : 0), 0);
      for (const k of numCols) {
        const ours = rows.reduce((a, r) => a + (isNum(r[k]) ? r[k] : 0), 0);
        if (isNum(grand[k]) && grand[k] !== 0 && Math.abs(ours - grand[k]) > 0.05) {
          log.warnings.push(`F4 контролна сума „${k}“: изчислено ${ours.toFixed(2)} ≠ отчет ${grand[k].toFixed(2)}.`);
        }
      }
      log.info.push(`F4: ${rows.length} служители; контролните суми са сравнени с ${totals.length} тотални реда на отчета.`);
    } else log.info.push(`F4: ${rows.length} служители; отчетът няма тотални редове за сверка.`);

    return { header: header.slice(firstCol, lastCol + 1), rows, totals };
  }

  // ---------- join + workbook assembly ----------
  function buildResult(prod, acc, cfg, log) {
    const r = cfg.roundDecimals;
    const accByName = new Map();
    for (const a of acc.rows) {
      const k = normName(a.name);
      if (accByName.has(k)) log.warnings.push(`F4: дублирано име „${a.name}“ — съпоставянето по име е ненадеждно за него.`);
      accByName.set(k, a);
    }

    const matchedAcc = new Set();
    const summary = prod.employees.map(e => {
      const a = accByName.get(normName(e.name)) || null;
      if (a) matchedAcc.add(normName(e.name));
      else log.warnings.push(`Несъпоставен: „${e.name}“ (${e.code}) има производителност, но липсва в счетоводния файл.`);
      return {
        'Име': e.name, 'Код': e.code, 'Бригада': e.brigade,
        'Отработени дни (F1)': e.daysWorked,
        'Заработка общо (лв)': round(e.totalEarn, r),
        'Бонус общо (лв)': round(e.totalBonus, r),
        ...(a ? Object.fromEntries(Object.entries(a).filter(([k]) => k !== 'name')
              .map(([k, v]) => [k, isNum(v) ? round(v, r) : v])) : { 'F4': 'НЕНАМЕРЕН' })
      };
    });
    for (const a of acc.rows) {
      if (!matchedAcc.has(normName(a.name))) {
        log.warnings.push(`Несъпоставен: „${a.name}“ е в счетоводния файл, но няма данни за производителност.`);
      }
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
      for (const [k, v] of Object.entries(a)) if (k !== 'name') o[k] = v;
      return o;
    });

    return { summary, days, ops, accounting };
  }

  // ---------- top-level ----------
  function process(prodAoa, accAoa, userCfg) {
    const cfg = { ...DEFAULT_CONFIG, ...(userCfg || {}) };
    const log = { info: [], warnings: [], errors: [], config: cfg, startedAt: new Date().toISOString() };
    let result = null;
    try {
      const parsedProd = parseProductivity(prodAoa, log);
      if (!parsedProd.employees.length) throw new Error('Файлът с производителност не съдържа разпознаваеми блокове служители.');
      for (const sk of parsedProd.skipped) {
        if (sk.reason !== 'page-junk') log.warnings.push(`Пропуснат ред ${sk.row} (${sk.reason}): ${sk.preview}`);
      }
      log.info.push(`Премахнати технически редове от страниране: ${parsedProd.skipped.filter(x => x.reason === 'page-junk').length}.`);
      if (parsedProd.meta.period) log.info.push(`Период: ${parsedProd.meta.period}`);

      const computed = compute(parsedProd, cfg, log);
      const acc = accAoa ? parseAccounting(accAoa, log) : { header: [], rows: [], totals: [] };
      result = buildResult(computed, acc, cfg, log);
      log.status = log.warnings.length ? 'SUCCESS_WITH_WARNINGS' : 'SUCCESS';
    } catch (err) {
      log.errors.push(String(err && err.message || err));
      log.status = 'ERROR';
    }
    log.finishedAt = new Date().toISOString();
    return { result, log };
  }

  function statusText(log) {
    const L = [];
    L.push(`=== СТАТУС: ${log.status} ===`);
    L.push(`Начало: ${log.startedAt}  Край: ${log.finishedAt}`);
    L.push('', '--- Конфигурация ---');
    for (const [k, v] of Object.entries(log.config)) L.push(`${k}: ${v}`);
    L.push('', `--- Информация (${log.info.length}) ---`, ...log.info);
    L.push('', `--- Предупреждения (${log.warnings.length}) ---`, ...(log.warnings.length ? log.warnings : ['няма']));
    L.push('', `--- Грешки (${log.errors.length}) ---`, ...(log.errors.length ? log.errors : ['няма']));
    return L.join('\n');
  }

  return { DEFAULT_CONFIG, process, statusText, parseBgNumber, parsePct, parseBgDate };
});
