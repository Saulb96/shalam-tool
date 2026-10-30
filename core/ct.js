// Container Tracking Sheet: Booking and Loaded (PO blocks under round sections) and Open and Complete (one row per
// container). Ported from the Late & Short engine (parse_rounds, parse_po_sheet, parse_container_sheet,
// parse_tracker), with columns found by header name. Pure: no network, no DOM; today (a day number) is passed in.
// Loads in the page (window.InvCT, after core/common.js) and in Node (module.exports).
(function (root, factory) {
  const C = typeof module === 'object' && module.exports ? require('./common.js') : root.InvCommon;
  if (!C) throw new Error('core/common.js must load before core/ct.js');
  const api = factory(C);
  if (typeof module === 'object' && module.exports) module.exports = api;
  root.InvCT = api;
})(typeof globalThis !== 'undefined' ? globalThis : this, function (C) {
  'use strict';
  const SOURCE = 'Container Tracking';
  const S = C.S;

  // Booking and Loaded share one layout. Booking's order date and ship by columns have no header, so they are absent.
  const PO_SPEC = {
    category: { names: ['Category'], required: true },
    style: { names: ['Style', 'Style Number', 'Style No'], required: true },
    qty: { names: ['Quantity', 'Qty', 'Units'], required: true },
    po: { names: ['PO Number', 'PO', 'PO No'], required: true },
    upc: { names: ['Barcode', 'UPC'] },
    master: { names: ['Master', 'Master Pack', 'Master Carton'] },
    factory: { names: ['Company', 'Factory', 'Supplier'] },
    name: { names: ['Product name', 'Description'] },
    price: { names: ['Price'] },
    amount: { names: ['PO Amount'] },
    orderDate: { names: ['Order Date'] },
    shipBy: { names: ['Ship by', 'Ship Date'] },
    readyDate: { names: ['Ready Date', 'Cargo Ready Date', 'CRD'] },
    needBy: { names: ['Need by', 'Need By Date'] },
  };
  const CONT_SPEC = {
    id: { names: ['ID', 'ID Number'] },
    container: { names: ['Container', 'Container Number', 'Container No', 'Containers'], required: true },
    round: { names: ['Round', 'Round Number', 'Round No'], required: true },
    etd: { names: ['ETD'], required: true },
    eta: { names: ['ETA'], required: true },
    factory: { names: ['Factory', 'Company', 'Supplier'] },
    po: { names: ['PO', 'PO Number'] },
    freightReleased: { names: ['Freight released', 'Freight Release'] },
    customsCleared: { names: ['Customs Cleared', 'Customs Released', 'Customs'] },
    fda: { names: ['FDA Released', 'FDA'] },
    doSent: { names: ['DO Sent to trucker', 'DO Sent'] },
    plSentToWh: { names: ['Packing list sent to WH', 'Packing List Sent', 'PL sent to WH'] },
    notes: { names: ['Notes'] },
  };
  const LABELS = {
    orderDate: 'Order Date', shipBy: 'Ship by', readyDate: 'Ready Date', etd: 'ETD', eta: 'ETA',
    freightReleased: 'Freight released', customsCleared: 'Customs Cleared', doSent: 'DO Sent to trucker',
    plSentToWh: 'Packing list sent to WH', name: 'Product name', upc: 'Barcode', price: 'Price', po: 'PO Number',
    factory: 'Company', amount: 'PO Amount',
  };
  // Cells a section header row doesn't have: one filled in means the row may be a product line read as a header.
  const LINE_CELLS = ['name', 'upc', 'price', 'po', 'factory', 'amount'];
  const TABS = ['Booking', 'Loaded', 'Open', 'Complete'];
  const CONT_RE = /[A-Z]{4}\d{7}/g;
  // Words people type in a date column meaning "nothing to report"; read as blank, not listed.
  const NOT_APPLICABLE = /^(n\/?a|none|-+)$/i;
  const FLOOR = C.ymd(2015, 1, 1);

  // 'Round 302-306' -> [302..306]; 'Round 298 & 299' -> [298, 299]; 'Round 358 incomplete' -> [358]; else [].
  function parseRounds(text) {
    const m = S(text).match(/^\s*ROUND\s*(.*)$/i);
    if (!m) return [];
    const out = new Set();
    const RANGE = /(\d+)\s*-\s*(\d+)/g;
    for (const [, a, b] of m[1].matchAll(RANGE)) {
      const x = +a, y = +b;
      if (y - x > 0 && y - x < 20) for (let n = x; n <= y; n++) out.add(n);
    }
    for (const n of m[1].replace(RANGE, '').match(/\d+/g) || []) out.add(+n);
    return [...out].sort((p, q) => p - q);
  }
  // parseRounds, plus odd: true when the text has other numbers after the round ('Round 371 - sent 11/2',
  // 'Round 371 (2 containers)'). Those are dates or counts, not rounds, so only the first number or range is kept.
  function readRounds(text) {
    const all = parseRounds(text);
    const m = S(text).match(/^\s*ROUND\s*(\d+(?:\s*-\s*\d+)?)(.*)$/i);
    if (!m || all.length < 2) return { rounds: all, odd: false };
    const lead = +m[1].match(/\d+/)[0];
    if (!all.some(n => Math.abs(n - lead) > 20) && !/[a-z][^]*\d/i.test(m[2])) return { rounds: all, odd: false };
    const first = parseRounds('Round ' + m[1]);
    return { rounds: first.length ? first : [lead], odd: true };
  }

  const blank = v => S(v) === '';
  function addYears(day, k) {
    const o = new Date(day * C.DAY), y = o.getUTCFullYear() + k, m = o.getUTCMonth() + 1, d = o.getUTCDate();
    return C.ymd(y, m, d) ?? C.ymd(y, m, d - 1); // 2/29 -> 2/28
  }
  // A date cell: { day, yearless, text, bad }. Blank and "N/A" are { day: null }.
  function readDate(v, refYear) {
    const text = S(v);
    if (text === '' || NOT_APPLICABLE.test(text)) return { day: null, yearless: false, text: '' };
    const r = C.toDay(v, refYear);
    return r ? { day: r.day, yearless: r.yearless, text } : { day: null, yearless: false, text, bad: true };
  }

  function parseContainerTracking(XLSX, wb, today) {
    if (typeof today !== 'number' || !Number.isFinite(today)) throw new TypeError('parseContainerTracking: today must be a day number');
    const todayYear = C.dayYear(today);
    // A date that reads cleanly but is over a year ahead or before 2015 is most likely a typo in the year.
    const far = d => d != null && (d > today + 366 || d < FLOOR);
    const fail = (error, errorDetail) => ({ ok: false, error, errorDetail, blocks: [], rounds: { open: {}, complete: {} },
      issues: [], counts: {} });

    // Every tab and every required column is checked before anything is read, so one message names all of them.
    const tabs = {}, problems = [];
    for (const tab of TABS) {
      const hit = C.findSheet(wb, tab);
      if (!hit) { problems.push({ sheet: tab, missingTab: true, seen: (wb.SheetNames || []).slice() }); continue; }
      const rows = C.sheetRows(XLSX, hit.ws);
      const spec = tab === 'Booking' || tab === 'Loaded' ? PO_SPEC : CONT_SPEC;
      const f = C.findColumns(rows[0] || [], spec);
      if (f.missing.length) problems.push({ sheet: tab, missing: f.missing, seen: f.seen });
      tabs[tab] = { rows, cols: f.cols };
    }
    if (problems.length) {
      const msg = problems.map(p => p.missingTab
        ? `there is no "${p.sheet}" tab (tabs seen: ${p.seen.join(', ') || 'none'})`
        : `the ${p.sheet} tab has no ${p.missing.map(n => '"' + n + '"').join(', ')} column (headers seen in row 1: ${p.seen.join(', ') || 'none'})`);
      return fail(SOURCE + ': ' + msg.join('; ') + '.', problems);
    }

    const issues = [];
    const issue = (code, sheet, row, message, extra) => issues.push(Object.assign({ source: SOURCE, code, sheet, row, message }, extra));
    const counts = {};
    const blocks = [];
    // Rows read stop at the last row with anything in it; the tabs carry thousands of empty formatted rows below.
    const lastRow = rows => { let n = rows.length; while (n > 1 && rows[n - 1].every(blank)) n--; return n; };

    // ---- Booking and Loaded: section header rows (Category filled, Style and Quantity empty) name the round(s);
    // a row with a PO number starts a PO block; style rows under it are its lines.
    const bookingPoSection = new Map(); // po -> { section, row } first seen on Booking
    for (const sheet of ['Booking', 'Loaded']) {
      const { rows, cols } = tabs[sheet];
      const at = (r, f) => (cols[f] == null ? null : r[cols[f]] ?? null);
      const n = { rowsRead: 0, blankRows: 0, sectionRows: 0, styleRows: 0, counted: 0, notCounted: 0, otherRows: 0, blocks: 0, units: 0 };
      let section = null, rounds = [], cur = null;
      const lineCats = new Set(), sections = [];   // Category values on product lines; plain section header rows
      const end = lastRow(rows);
      for (let i = 1; i < end; i++) {
        const r = rows[i] || [], row = i + 1;
        n.rowsRead++;
        if (r.every(blank)) { n.blankRows++; continue; }
        const styleV = at(r, 'style'), qtyV = at(r, 'qty'), poV = at(r, 'po'), cat = S(at(r, 'category'));
        if (cat && blank(styleV) && blank(qtyV)) {
          n.sectionRows++;
          section = cat;
          const rr = readRounds(section);
          rounds = rr.rounds;
          if (/^\s*round/i.test(section) && !rounds.length)
            issue('round-no-number', sheet, row, `${sheet} row ${row}: the section "${section}" looks like a round, but no round number could be read.`);
          if (rr.odd)
            issue('round-odd', sheet, row, `${sheet} row ${row}: the section "${section}" has other numbers after the round, so it is read as round ${rounds.join(', ')} only. Check the header.`, { round: rounds[0] });
          const filled = LINE_CELLS.filter(f => !blank(at(r, f)));
          if (filled.length) {
            issue('odd-section', sheet, row, `${sheet} row ${row}: "${section}" is read as a section header (it has no style and no quantity), but it also has ${filled.map(f => LABELS[f]).join(', ')}. If it is a product line, add its style and quantity; the POs below it are filed under "${section}".`);
          } else sections.push({ row, section });
          cur = null;
          continue;
        }
        if (/^\s*round\s*\d/i.test(cat))
          issue('odd-section', sheet, row, `${sheet} row ${row}: "${cat}" looks like a round header, but the row has a style or quantity, so it isn't read as one; the POs below it stay under "${section ?? '(no section)'}".`);
        if (!blank(poV)) {
          const po = C.normPo(poV);
          const dates = {};
          for (const f of ['orderDate', 'shipBy', 'readyDate']) {
            if (cols[f] == null) { dates[f] = null; continue; }
            const d = readDate(at(r, f), todayYear);
            // A Ready Date typed without a year (Booking and Loaded are goods still to ship) takes the year that puts it
            // nearest today: "1/20" typed in November is next January's, "12/20" typed in January is last December's.
            if (f === 'readyDate' && d.yearless && d.day != null) {
              if (d.day < today - 182) d.day = addYears(d.day, 1);
              else if (d.day > today + 182) d.day = addYears(d.day, -1);
            }
            dates[f] = d.day;
            if (d.bad) issue('not-a-date', sheet, row, `${sheet} row ${row}, PO ${po}: ${LABELS[f]} "${d.text.slice(0, 30)}" is not a date.`, { po });
            else if (d.yearless) issue('yearless-date', sheet, row, `${sheet} row ${row}, PO ${po}: ${LABELS[f]} "${d.text}" has no year; read as ${C.fmtDay(d.day)}.`, { po });
            if (far(d.day)) issue('date-far', sheet, row, `${sheet} row ${row}, PO ${po}: ${LABELS[f]} reads as ${C.fmtDay(d.day)}, more than a year from today or before 2015. Check the year.`, { po });
          }
          if (sheet === 'Booking') {
            const first = bookingPoSection.get(po);
            if (!first) bookingPoSection.set(po, { section, row });
            else if (first.section !== section)
              issue('po-two-sections', sheet, row, `Booking: PO ${po} is under "${first.section ?? '(no section)'}" (row ${first.row}) and under "${section ?? '(no section)'}" (row ${row}).`, { po });
          }
          cur = {
            sheet, row, section, rounds, po, poRaw: S(poV), factory: S(at(r, 'factory')),
            orderDate: dates.orderDate, shipBy: dates.shipBy, readyDate: dates.readyDate, needBy: S(at(r, 'needBy')),
            lines: [],
          };
          blocks.push(cur);
          n.blocks++;
        }
        if (blank(styleV)) {
          n.otherRows++;
          if (!blank(qtyV))
            issue('qty-no-style', sheet, row, `${sheet} row ${row}: a quantity ("${S(qtyV).slice(0, 20)}") with no style, so it is not counted.`);
          else if (['name', 'upc', 'price', 'amount'].some(f => !blank(at(r, f))))
            issue('no-style', sheet, row, `${sheet} row ${row}${cur ? ', PO ' + cur.po : ''}: a product line with no style and no quantity, so it is not counted.`, cur ? { po: cur.po } : {});
          continue;
        }
        n.styleRows++;
        const style = C.normStyle(styleV);
        if (cat) lineCats.add(cat.toLowerCase());
        if (!cur) {
          n.notCounted++;
          issue('no-po', sheet, row, `${sheet} row ${row}, style ${style}: there is no PO number above it in its section, so ${blank(qtyV) ? 'it is' : 'its quantity (' + S(qtyV) + ') is'} not counted.`, { style });
          continue;
        }
        const qty = C.num(qtyV);
        if (qty == null || Number.isNaN(qty)) {
          n.notCounted++;
          issue('qty-not-number', sheet, row, `${sheet} row ${row}, PO ${cur.po}, style ${style}: the quantity ${qty == null ? 'is blank' : '"' + S(qtyV).slice(0, 20) + '" is not a number'}, so it is not counted.`, { po: cur.po, style });
          continue;
        }
        const master = C.num(at(r, 'master'));
        const upcV = at(r, 'upc');
        if (!blank(upcV) && !NOT_APPLICABLE.test(S(upcV)) && !C.normUpc(upcV))
          issue('bad-upc', sheet, row, `${sheet} row ${row}, PO ${cur.po}, style ${style}: the barcode "${S(upcV).slice(0, 20)}" isn't 12 or 13 digits, so it can't be matched by UPC. The line is still counted.`, { po: cur.po, style });
        cur.lines.push({ row, style, styleRaw: S(styleV), upc: C.normUpc(upcV), upcRaw: S(upcV), master: Number.isNaN(master) ? null : master, qty });
        n.counted++;
        n.units += qty;
      }
      // A section named like a product Category ("Lip") is more likely a product line missing its style and quantity.
      for (const x of sections) {
        if (!readRounds(x.section).rounds.length && lineCats.has(x.section.toLowerCase()))
          issue('odd-section', sheet, x.row, `${sheet} row ${x.row}: "${x.section}" is read as a section header, but product lines on this tab use "${x.section}" as their Category. If it is a product line, add its style and quantity; the POs below it are filed under "${x.section}".`);
      }
      counts[sheet] = n;
    }

    // ---- Open and Complete: every row with a round number is kept, under each round it names.
    const rounds = { open: {}, complete: {} };
    for (const sheet of ['Open', 'Complete']) {
      const { rows, cols } = tabs[sheet];
      const at = (r, f) => (cols[f] == null ? null : r[cols[f]] ?? null);
      const n = { rowsRead: 0, blankRows: 0, labelRows: 0, roundRows: 0, noRoundRows: 0 };
      const into = rounds[sheet.toLowerCase()];
      const end = lastRow(rows);
      for (let i = 1; i < end; i++) {
        const r = rows[i] || [], row = i + 1;
        n.rowsRead++;
        if (r.every(blank)) { n.blankRows++; continue; }
        const roundText = S(at(r, 'round')), containerText = S(at(r, 'container'));
        if (!roundText) {
          // A title row (one text cell in column A or ID#, such as "Complete Shipments") is not a container row; any
          // other row without a round is listed.
          const cells = r.map((v, j) => (blank(v) ? -1 : j)).filter(j => j >= 0);
          const title = cells.length === 1 && (cells[0] === 0 || cells[0] === cols.id) && Number.isNaN(Number(S(r[cells[0]])));
          if (!title) {
            n.noRoundRows++;
            issue('no-round', sheet, row, `${sheet} row ${row}: ${containerText ? 'container ' + containerText.replace(/\s+/g, ' ').slice(0, 40) : 'this row'} has no round number, so it is not tied to any round.`);
          } else n.labelRows++;
          continue;
        }
        const rr = readRounds('Round ' + roundText), nums = rr.rounds;
        if (!nums.length) {
          n.noRoundRows++;
          issue('round-not-number', sheet, row, `${sheet} row ${row}: the round "${roundText.slice(0, 30)}" is not a round number, so this row is not tied to any round.`);
          continue;
        }
        n.roundRows++;
        const where = `${sheet} row ${row}, round ${roundText}`;
        if (rr.odd) issue('round-odd', sheet, row, `${where}: the Round # has other numbers after the round, so it is read as round ${nums.join(', ')} only. Check the cell.`, { round: nums[0] });
        const listDate = (f, d) => {
          if (d.bad) issue('not-a-date', sheet, row, `${where}: ${LABELS[f]} "${d.text.slice(0, 30)}" is not a date.`, { round: nums[0] });
          else if (d.yearless) issue('yearless-date', sheet, row, `${where}: ${LABELS[f]} "${d.text}" has no year; read as ${C.fmtDay(d.day)}.`, { round: nums[0] });
          if (far(d.day)) issue('date-far', sheet, row, `${where}: ${LABELS[f]} reads as ${C.fmtDay(d.day)}, more than a year from today or before 2015. Check the year.`, { round: nums[0] });
        };
        // Complete rows are history: a date typed without a year can't be after today (nor, for the ETA, before the ETD).
        const past = (d, floor) => (sheet === 'Complete' && d.yearless && d.day > today && (floor == null || addYears(d.day, -1) >= floor) ? addYears(d.day, -1) : d.day);
        // Year fixes for dates typed without a year (Late & Short's, plus the two added by the plan review).
        const etdR = readDate(at(r, 'etd'), todayYear);
        if (etdR.yearless && etdR.day > today + 60) etdR.day = addYears(etdR.day, -1);
        etdR.day = past(etdR);
        const etd = etdR.day;
        const etaR = readDate(at(r, 'eta'), etd != null ? C.dayYear(etd) : todayYear);
        if (etaR.yearless) {
          if (etd != null && etaR.day < etd) etaR.day = addYears(etaR.day, 1);
          else if (etaR.day > today + 120) etaR.day = addYears(etaR.day, -1);
          // Only on Open: a shipment still in transit can't have arrived months ago.
          else if (etd == null && sheet === 'Open' && etaR.day < today - 60) etaR.day = addYears(etaR.day, 1);
          etaR.day = past(etaR, etd);
        }
        const eta = etaR.day;
        listDate('etd', etdR);
        listDate('eta', etaR);
        if (sheet === 'Open' && eta == null && !etaR.bad)
          issue('no-eta', sheet, row, `${where}: no ETA yet${etd != null ? ' (ETD ' + C.fmtDay(etd) + ')' : ''}.`, { round: nums[0] });
        const otherYear = C.dayYear(eta ?? etd ?? today);
        const other = {};
        for (const f of ['freightReleased', 'customsCleared', 'doSent', 'plSentToWh']) {
          if (cols[f] == null) { other[f] = null; continue; }
          const d = readDate(at(r, f), otherYear);
          d.day = past(d);
          other[f] = d.day;
          listDate(f, d);
        }
        const containers = [...new Set(containerText.toUpperCase().match(CONT_RE) || [])];
        if (sheet === 'Open' && containerText && !containers.length)
          issue('bad-container', sheet, row, `${where}: "${containerText.replace(/\s+/g, ' ').slice(0, 40)}" has no container number (4 letters and 7 digits), so the forwarder and drayage reports can't be matched to it.`, { round: nums[0] });
        const rec = {
          sheet, row, id: S(at(r, 'id')), containerText, containers,
          roundText, rounds: nums, etd, eta, etaYearless: etaR.yearless, etdYearless: etdR.yearless,
          factoryText: S(at(r, 'factory')), poText: S(at(r, 'po')),
          freightReleased: other.freightReleased, customsCleared: other.customsCleared, fdaText: S(at(r, 'fda')),
          doSent: other.doSent, plSentToWh: other.plSentToWh, notes: S(at(r, 'notes')),
        };
        for (const k of nums) (into[k] = into[k] || []).push(rec);
      }
      n.rounds = Object.keys(into).length;
      counts[sheet] = n;
    }

    // ---- Checks across tabs.
    const shipped = k => rounds.open[k] || rounds.complete[k];
    const flaggedRound = new Set();
    for (const b of blocks) {
      for (const k of b.rounds) {
        const key = b.sheet + '|' + k;
        if (shipped(k) || flaggedRound.has(key)) continue;
        flaggedRound.add(key);
        issue('round-no-open-complete', b.sheet, b.row, `${b.sheet} row ${b.row} (section "${b.section}"): round ${k} has no row on Open or Complete.`, { round: k });
      }
    }
    const onLoaded = new Map(); // po|style|round -> Loaded row
    for (const b of blocks) if (b.sheet === 'Loaded')
      for (const l of b.lines) for (const k of b.rounds) { const key = b.po + '|' + l.style + '|' + k; if (!onLoaded.has(key)) onLoaded.set(key, l.row); }
    const flaggedBoth = new Set();
    for (const b of blocks) if (b.sheet === 'Booking')
      for (const l of b.lines) for (const k of b.rounds) {
        const key = b.po + '|' + l.style + '|' + k;
        if (!onLoaded.has(key) || flaggedBoth.has(key)) continue;
        flaggedBoth.add(key);
        issue('on-both-tabs', 'Booking', l.row, `PO ${b.po}, style ${l.style}, round ${k} is on both Booking (row ${l.row}) and Loaded (row ${onLoaded.get(key)}); the stage numbers will count it once, from Loaded.`,
          { po: b.po, style: l.style, round: k, loadedRow: onLoaded.get(key) });
      }
    // On both tabs under different sections where one side isn't in a round: maybe a Booking row that is out of date.
    // A PO + style split across two real rounds is a split shipment and isn't listed.
    const loadedAt = new Map(); // po|style -> [{ row, section, rounds }]
    for (const b of blocks) if (b.sheet === 'Loaded')
      for (const l of b.lines) { const key = b.po + '|' + l.style; (loadedAt.get(key) || loadedAt.set(key, []).get(key)).push({ row: l.row, section: b.section, rounds: b.rounds }); }
    const flaggedOther = new Set();
    for (const b of blocks) if (b.sheet === 'Booking')
      for (const l of b.lines) {
        const key = b.po + '|' + l.style, L = loadedAt.get(key);
        if (!L || flaggedOther.has(key)) continue;
        if (L.some(x => (b.rounds.length ? x.rounds.some(k => b.rounds.includes(k)) : x.section === b.section))) continue;
        if (b.rounds.length && L.every(x => x.rounds.length)) continue;
        flaggedOther.add(key);
        issue('on-both-tabs-other-round', 'Booking', l.row, `PO ${b.po}, style ${l.style} is on Booking (row ${l.row}, under "${b.section ?? '(no section)'}") and on Loaded (row ${L[0].row}, under "${L[0].section ?? '(no section)'}"). If the Booking row is out of date it would be counted twice: check which one is current.`,
          { po: b.po, style: l.style, loadedRow: L[0].row });
      }

    // A file that opens but gives nothing is a failure, never "nothing in transit".
    const empty = [];
    if (counts.Booking.counted + counts.Loaded.counted === 0) empty.push('no lines were read from Booking or Loaded');
    if (counts.Open.roundRows + counts.Complete.roundRows === 0) empty.push('no container rows were read from Open or Complete');
    if (empty.length) return { ok: false, error: SOURCE + ': ' + empty.join('; ') + '.', errorDetail: [{ sheet: null, empty }], blocks, rounds, issues, counts };

    return { ok: true, error: null, blocks, rounds, issues, counts };
  }

  return { parseRounds, readRounds, parseContainerTracking, PO_SPEC, CONT_SPEC };
});
