// The sales sheet (Saul, Nov 2): what a salesperson may see of Available to sell, for the Excel file saved into the one
// library they can open. It holds only what's listed here, per style: style, UPCs, description, category and sub
// category (Saul, Nov 3: a tab per category; the ATS report's, else UPC_MASTER's, the pair from one place), case pack, units and full cases available now, from China, arriving and
// in total, the first dates, a photo; and a schedule by date. Never
// a factory, customer, PO, PO type, section, container, round, person, file name, SharePoint link, ERP warehouse, In
// WH, customer orders, sold, oversold, flag or problem: every field is copied by name (an allow-list), never by
// spreading an object of Available to sell.
// - salesRows(sell, { today }): the rows and the schedule from InvSell.buildSell's result; none (ok false, with the
//   reason in plain words) when "Available now" would be wrong: no usable ATS report, Available to sell not worked out,
//   or its units don't add up.
// - buildWorkbook(ExcelJS, sheet, photos, { print }): the Excel file (ExcelJS 4.4.0, vendor/exceljs): Available to
//   sell (every style, a photo per style fitted in its cell the way Quote Tracker does it), then a tab per category
//   (categoryTabs: A to Z, named as the ERP has them, never joined), then By date; the page's fingerprint of it in its
//   keywords. Each picture is stored once in the file, however many tabs show it.
// Pure: no network, no DOM; today (a day number) is passed in. Loads in the page (window.InvSalesSheet, after
// core/common.js) and in Node (module.exports).
(function (root, factory) {
  const node = typeof module === 'object' && module.exports;
  const api = factory(node ? require('./common.js') : root.InvCommon);
  if (node) module.exports = api;
  root.InvSalesSheet = api;
})(typeof globalThis !== 'undefined' ? globalThis : this, function (C) {
  'use strict';
  if (!C) throw new Error('core/common.js must load before core/salesSheet.js');
  const { S, DAY } = C;
  const fmt = C.fmtDay;

  const TITLE = 'M1S · Available to sell';
  const TABS = { sell: 'Available to sell', byDate: 'By date' };
  // Where a schedule row's goods are, in the order they're listed on the same date.
  const WHERE = ['In stock', 'From China', 'Arriving'];
  const isDay = v => typeof v === 'number' && Number.isFinite(v);
  const units = v => (isDay(v) && v > 0 ? v : 0);
  // Full cases (core/sell.js's casesOf): a part case is never offered as a case; null when the Master isn't known.
  const casesOf = (n, master) => (isDay(master) && master > 0 ? Math.max(0, Math.floor(n / master + 1e-9)) : null);
  // Quote Tracker's photo key: the style uppercased, only A-Z, 0-9, _ and - kept. Exact key only (no fallback from
  // "X-F" to "X").
  const photoKey = s => String(s ?? '').toUpperCase().replace(/[^A-Z0-9_-]/g, '');
  const byStyle = (a, b) => (a < b ? -1 : a > b ? 1 : 0);
  // Categories and sub categories A to Z (Cosmetics1 before Cosmetics2, case aside), a blank one last.
  const byName = (a, b) => (!a) - (!b) || a.localeCompare(b, 'en', { sensitivity: 'base', numeric: true }) || byStyle(a, b);
  // A schedule row's date: now first, then by day, no date last.
  const whenRank = w => (w === 'now' ? -Infinity : isDay(w) ? w : Infinity);

  // Why there can be no sheet ('' when there can).
  function refusal(sell) {
    if (!sell || sell.ok !== true || !Array.isArray(sell.styles)) {
      const nl = sell && Array.isArray(sell.notLoaded) && sell.notLoaded.length ? ` (${sell.notLoaded.join(' and ')} didn't load)` : '';
      return `Available to sell couldn't be worked out${nl}, so there's no sales sheet.`;
    }
    if (sell.atsMissing) {
      return sell.atsError
        ? `The ATS report couldn't be used (${S(sell.atsError)}), so stock available now isn't known and there's no sales sheet.`
        : "There's no ATS report, so stock available now isn't known and there's no sales sheet.";
    }
    if ((sell.issues || []).some(i => i && i.code === 'balance')) return "Available to sell doesn't add up (a bug in Shalam Tool), so there's no sales sheet until it's fixed.";
    return '';
  }

  // sell: InvSell.buildSell's result. opts.today: the day the sheet is made (a day number; defaults to sell's own).
  // Answers { ok, reason, asOf, reportDay, rows, byDate, counts }:
  // rows, one per style with something available, by style: { style, upc, desc, casePack (a number, 'mixed' or null),
  //   now, nowCases, fromChina, fromChinaCases, firstFromChina: { date, now, estimate }, arriving, arrivingCases,
  //   firstArrival, total, totalCases, photoKey }. Units are numbers (0 for none); cases are full cases, null when the
  //   case pack is unknown or mixed (never guessed). firstFromChina.now: its pull-from date is before today;
  //   .estimate: that date is about 75 days after the order date.
  // byDate, one per style per date and where (In stock, From China, Arriving), by style then date (now first, no date
  //   last): { style, desc, when ('now', a day number, or null: no date yet), estimate, where, units, cases }. Units are
  //   summed over the style's lots with the same date and place; cases per the lots' Master, null when they differ or
  //   one isn't known.
  // counts: { styles (on the sheet), leftOff (nothing available), byDateRows, noCasePack, mixedCasePack }.
  function salesRows(sell, opts) {
    const today = opts && opts.today != null ? opts.today : sell && sell.today;
    if (!isDay(today)) throw new TypeError('InvSalesSheet.salesRows: today must be a day number');
    const counts = { styles: 0, leftOff: 0, byDateRows: 0, noCasePack: 0, mixedCasePack: 0 };
    const reportDay = sell && isDay(sell.reportDay) ? sell.reportDay : null;
    const why = refusal(sell);
    if (why) return { ok: false, reason: why, asOf: today, reportDay, rows: [], byDate: [], counts };

    const rows = [], byDate = [];
    // A style's category and sub category: the ATS report's (Saul, Nov 3), else the page's (UPC_MASTER, then Master
    // Sheet Live); both from the same place, never one from each. By category, sub category, then style: each
    // category's tab and the By date tab keep the same order.
    const catOf = s => (S(s.atsCategory) ? [S(s.atsCategory), S(s.atsSub)] : [S(s.category), S(s.sub)]);
    const order = (a, b) => { const x = catOf(a), y = catOf(b); return byName(x[0], y[0]) || byName(x[1], y[1]) || byStyle(a.style, b.style); };
    for (const s of sell.styles.slice().sort(order)) {
      const now = units(s.now), fromChina = units(s.fromChina), arriving = units(s.arriving);
      const total = now + fromChina + arriving;
      if (!(total > 0)) { counts.leftOff++; continue; }
      const style = S(s.styleRaw || s.style), desc = S(s.desc), [category, sub] = catOf(s);
      const casePack = isDay(s.casePack) && s.casePack > 0 ? s.casePack : s.casePack === 'mixed' ? 'mixed' : null;
      const pack = typeof casePack === 'number' ? casePack : null;
      // A column's cases: core/sell.js's own count, when the style's case pack is one number and every unit in the
      // column has it (no units without a Master).
      const cases = k => (pack != null && s.cases && s.noCasePackUnits && s.noCasePackUnits[k] === 0 && isDay(s.cases[k]) ? s.cases[k] : null);
      // A lot's Master: our own lots carry theirs; the ERP's shipments take the style's, only when it's one number.
      const masterOf = l => (l.source === 'erp' ? pack : isDay(l.master) && l.master > 0 ? l.master : null);

      // The schedule: stock now, then each open lot by its date and place.
      const groups = new Map();
      const put = (where, when, n, master, cs, estimate) => {
        const key = where + '|' + (when == null ? '' : when);
        const g = groups.get(key) || groups.set(key, { where, when, units: 0, masters: new Set(), cases: 0, estimate: false }).get(key);
        g.units += n; g.masters.add(master); g.cases += cs == null ? 0 : cs;
        if (estimate) g.estimate = true;
      };
      if (now > 0) put('In stock', 'now', now, pack, cases('now'), false);
      let first = null, firstArrival = null;
      for (const l of Array.isArray(s.lots) ? s.lots : []) {
        const n = units(l.available);
        if (!n) continue;
        const m = masterOf(l);
        if (l.inChina) {
          const day = isDay(l.pullFrom) ? l.pullFrom : null;
          const isNow = day != null && day < today;
          put('From China', isNow ? 'now' : day, n, m, casesOf(n, m), !isNow && !!l.estimate);
          if (day != null && (!first || day < first.date || (day === first.date && !l.estimate))) first = { date: day, now: isNow, estimate: !isNow && !!l.estimate };
        } else {
          const day = isDay(l.arrival) ? l.arrival : null;
          put('Arriving', day, n, m, casesOf(n, m), false);
          if (day != null && (firstArrival == null || day < firstArrival)) firstArrival = day;
        }
      }
      const sched = [...groups.values()].sort((a, b) => whenRank(a.when) - whenRank(b.when) || WHERE.indexOf(a.where) - WHERE.indexOf(b.where));
      for (const g of sched) {
        const one = g.masters.size === 1 && [...g.masters][0] != null;
        byDate.push({ style, desc, category, when: g.when, estimate: g.estimate, where: g.where, units: g.units, cases: one ? g.cases : null });
      }

      const totalCases = cases('total');
      rows.push({
        style, upc: (Array.isArray(s.upcs) ? s.upcs : []).map(S).filter(Boolean).join(', '), desc, category, sub, casePack,
        now, nowCases: cases('now'), fromChina, fromChinaCases: cases('fromChina'),
        firstFromChina: first ? { date: first.date, now: first.now, estimate: first.estimate } : { date: null, now: false, estimate: false },
        arriving, arrivingCases: cases('arriving'), firstArrival,
        total, totalCases, photoKey: photoKey(style),
      });
      if (casePack == null) counts.noCasePack++;
      if (casePack === 'mixed') counts.mixedCasePack++;
    }
    counts.styles = rows.length; counts.byDateRows = byDate.length;
    return { ok: true, reason: '', asOf: today, reportDay, rows, byDate, counts };
  }

  // ---------- the Excel file
  // The Available to sell tab's columns: header, width (characters), the row's field, and how it's shown.
  const COLS = [
    { h: 'Photo', w: 15, photo: true },
    { h: 'Style', w: 16, k: 'style' },
    { h: 'UPC', w: 16, k: 'upc', wrap: true },
    { h: 'Description', w: 34, k: 'desc', wrap: true },
    { h: 'Category', w: 14, k: 'category', wrap: true },
    { h: 'Sub category', w: 16, k: 'sub', wrap: true },
    { h: 'Case pack', w: 9, k: 'casePack', n: true },
    { h: 'Available now (units)', w: 11, k: 'now', n: true },
    { h: 'Available now (cases)', w: 11, k: 'nowCases', n: true },
    { h: 'Available from China (units)', w: 12, k: 'fromChina', n: true },
    { h: 'Available from China (cases)', w: 12, k: 'fromChinaCases', n: true },
    { h: 'First date from China', w: 14, k: 'firstFromChina', first: true },
    { h: 'Arriving (units)', w: 11, k: 'arriving', n: true },
    { h: 'Arriving (cases)', w: 11, k: 'arrivingCases', n: true },
    { h: 'First arrival', w: 12, k: 'firstArrival', date: true },
    { h: 'Total available (units)', w: 12, k: 'total', n: true },
    { h: 'Total available (cases)', w: 12, k: 'totalCases', n: true },
  ];
  const DATE_COLS = [
    { h: 'Style', w: 16 }, { h: 'Description', w: 34 }, { h: 'Category', w: 14 }, { h: 'Available from', w: 16 }, { h: 'Where', w: 12 },
    { h: 'Units', w: 11 }, { h: 'Cases', w: 11 },
  ];
  const NUM = '#,##0', DATE = 'm/d/yyyy';
  const FONT = { name: 'Calibri', size: 11 };
  const HEAD_FILL = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE7EEF5' } };
  const THIN = { style: 'thin', color: { argb: 'FFBFC8D2' } };
  const BOX = { top: THIN, left: THIN, bottom: THIN, right: THIN };
  // Photo placement, ported from Quote Tracker (index.html, fitPhoto): the photo scaled to PHOTO_INSET of the cell's
  // box and centred in it on both axes with fractional anchors, pinned to its one cell; every row with photos keeps one
  // fixed height (rows of different heights read as broken).
  const PHOTO_INSET = 0.86, PHOTO_ROW_PT = 82.5;
  const colPx = ch => Math.round(ch * 7 + 5), ptPx = pt => Math.round(pt * 96 / 72);
  function fitPhoto(colIdx, colChars, rowPt, rowIdx, d) {
    const cellW = colPx(colChars), cellH = ptPx(rowPt);
    const sc = Math.min(cellW * PHOTO_INSET / d.w, cellH * PHOTO_INSET / d.h);
    const w = Math.max(1, Math.round(d.w * sc)), h = Math.max(1, Math.round(d.h * sc));
    return { tl: { col: colIdx + ((cellW - w) / 2) / cellW, row: rowIdx + ((cellH - h) / 2) / cellH }, ext: { width: w, height: h }, editAs: 'oneCell' };
  }
  // A day number as a date cell (UTC midnight: ExcelJS turns it into that day's serial, whatever the time zone).
  const dateOf = d => new Date(d * DAY);
  const blankIfNone = v => (isDay(v) && v !== 0 ? v : null);
  // A date as the sheet shows it: now, about m/d/yyyy (an estimate), or the date itself.
  const whenCell = (when, estimate, any) => (when === 'now' ? 'now' : isDay(when) ? (estimate ? `about ${fmt(when)}` : dateOf(when)) : any ? 'no date yet' : null);
  // A row's photo: from a list in the rows' order (the page's: each style's own picture, the ERP report's or Quote
  // Images'), else by its photo key.
  const photoOf = (photos, key, n) => (!photos ? null : Array.isArray(photos) ? photos[n] || null : !key ? null : photos instanceof Map ? photos.get(key)
    : Object.prototype.hasOwnProperty.call(photos, key) ? photos[key] : null);

  // The category tabs, A to Z with no category last: { name (the tab's name), category (as the ERP has it; '' for
  // none), rows (their places in sheet.rows, in its order) }. A name is the category with what Excel won't take in a
  // tab's name ([ ] : * ? / \) made a space, at most 31 characters, never "History" (Excel keeps that one) and never
  // the same as another tab's (case aside: a second one gets " (2)"). Categories are never joined: Cosmetics1 and
  // cosmetics1 are two tabs.
  const NO_CAT = 'No category';
  function categoryTabs(rows) {
    const groups = new Map();
    (rows || []).forEach((r, n) => {
      const c = S(r && r.category);
      (groups.get(c) || groups.set(c, []).get(c)).push(n);
    });
    const used = new Set(Object.values(TABS).map(t => t.toLowerCase()));
    return [...groups.keys()].sort(byName).map(category => {
      let base = category.replace(/[[\]:*?/\\]/g, ' ').replace(/\s+/g, ' ').trim().replace(/^'+|'+$/g, '').trim() || NO_CAT;
      if (base.toLowerCase() === 'history') base += ' (category)';
      base = base.slice(0, 31).trim();
      let name = base;
      for (let k = 2; used.has(name.toLowerCase()); k++) { const tag = ` (${k})`; name = base.slice(0, 31 - tag.length).trim() + tag; }
      used.add(name.toLowerCase());
      return { name, category, rows: groups.get(category) };
    });
  }

  // ExcelJS: the library (window.ExcelJS, or require('vendor/exceljs/exceljs.min.js') in Node). sheet: salesRows's
  // answer (ok, with rows). photos: a Map (or object) of photoKey -> { data (the image's bytes: ArrayBuffer or
  // Uint8Array), width, height (pixels), extension ('jpeg' when not given) }, or a list of those in the rows' order (null
  // for none); a style with none gets an empty cell.
  // opts.print: the page's fingerprint of what the file holds (a hash: letters and digits only), kept in the file's
  // keywords, so a later look at the file in the library can tell it holds the same sheet whoever saved it.
  // Answers the ExcelJS Workbook (workbookBytes writes it).
  function buildWorkbook(ExcelJS, sheet, photos, opts) {
    if (!ExcelJS || typeof ExcelJS.Workbook !== 'function') throw new Error("The Excel library didn't load, so the sales sheet can't be made.");
    if (!sheet || sheet.ok !== true) throw new Error((sheet && sheet.reason) || "There's no sales sheet to write.");
    if (!Array.isArray(sheet.rows) || !sheet.rows.length) throw new Error('Nothing is available to sell, so the sales sheet would be empty.');
    const print = S(opts && opts.print);
    if (print && !/^[0-9a-z]{1,128}$/i.test(print)) throw new Error("The sales sheet's fingerprint isn't a hash, so it isn't written into the file.");
    const wb = new ExcelJS.Workbook();
    wb.creator = 'Shalam Tool'; wb.lastModifiedBy = 'Shalam Tool';
    if (print) wb.keywords = print;

    // A tab of styles: the title row, the header row, a row per style (rowIdx: their places in sheet.rows). Each
    // picture goes into the file once (ids), however many tabs show it.
    const withPhotos = !!photos && (Array.isArray(photos) ? photos.some(Boolean) : photos instanceof Map ? photos.size > 0 : Object.keys(photos).length > 0);
    const picCol = COLS.findIndex(c => c.photo);
    const ids = new Map();
    const imageOf = n => {
      if (ids.has(n)) return ids.get(n);
      const p = photoOf(photos, sheet.rows[n].photoKey, n);
      const img = !p || !p.data ? null : { id: wb.addImage({ buffer: p.data, extension: p.extension || 'jpeg' }),
        d: isDay(p.width) && p.width > 0 && isDay(p.height) && p.height > 0 ? { w: p.width, h: p.height } : { w: 160, h: 160 } };
      ids.set(n, img);
      return img;
    };
    const styleTab = (name, heading, rowIdx) => {
      const ws = wb.addWorksheet(name);
      ws.columns = COLS.map(c => ({ width: c.w }));
      const title = ws.getRow(1);
      title.getCell(1).value = heading; title.getCell(1).font = { ...FONT, size: 14, bold: true };
      title.getCell(4).value = `as of ${fmt(sheet.asOf)} · stock from the ERP report of ${isDay(sheet.reportDay) ? fmt(sheet.reportDay) : '(date not known)'}`;
      title.getCell(4).font = { ...FONT, color: { argb: 'FF55606E' } };
      title.height = 24;
      const head = ws.getRow(2);
      COLS.forEach((c, i) => {
        const cell = head.getCell(i + 1);
        cell.value = c.h; cell.font = { ...FONT, bold: true }; cell.fill = HEAD_FILL; cell.border = BOX;
        cell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
      });
      head.height = 45;
      rowIdx.forEach((at, n) => {
        const r = sheet.rows[at], row = ws.getRow(n + 3);
        COLS.forEach((c, i) => {
          const cell = row.getCell(i + 1);
          cell.font = FONT; cell.border = BOX;
          cell.alignment = { vertical: 'middle', wrapText: !!c.wrap, ...(c.n ? { horizontal: 'right' } : c.first || c.date ? { horizontal: 'center' } : {}) };
          if (c.photo) return;
          if (c.first) { const f = r.firstFromChina || {}; cell.value = whenCell(f.now ? 'now' : f.date, f.estimate, r.fromChina > 0); }
          else if (c.date) cell.value = whenCell(r[c.k], false, r.arriving > 0);
          else if (c.n) cell.value = c.k === 'casePack' && r.casePack === 'mixed' ? 'mixed' : blankIfNone(r[c.k]);
          else cell.value = S(r[c.k]) || null;
          if (c.n) cell.numFmt = NUM;
          if ((c.first || c.date) && cell.value instanceof Date) cell.numFmt = DATE;
        });
        if (withPhotos) row.height = PHOTO_ROW_PT;
        const img = imageOf(at);
        if (img) ws.addImage(img.id, fitPhoto(picCol, COLS[picCol].w, PHOTO_ROW_PT, n + 2, img.d));
      });
      ws.views = [{ state: 'frozen', xSplit: 2, ySplit: 2, topLeftCell: 'C3' }];
      ws.autoFilter = { from: { row: 2, column: 1 }, to: { row: 2, column: COLS.length } };
      ws.pageSetup = { orientation: 'landscape', fitToPage: true, fitToWidth: 1, fitToHeight: 0, printTitlesRow: '2:2' };
    };

    // 1. Available to sell: every style. 2. A tab per category.
    styleTab(TABS.sell, TITLE, sheet.rows.map((r, n) => n));
    for (const t of categoryTabs(sheet.rows)) styleTab(t.name, `${TITLE} · ${t.category || NO_CAT}`, t.rows);

    // 3. By date: the schedule.
    const wd = wb.addWorksheet(TABS.byDate);
    wd.columns = DATE_COLS.map(c => ({ width: c.w }));
    const dh = wd.getRow(1);
    DATE_COLS.forEach((c, i) => {
      const cell = dh.getCell(i + 1);
      cell.value = c.h; cell.font = { ...FONT, bold: true }; cell.fill = HEAD_FILL; cell.border = BOX;
      cell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
    });
    (sheet.byDate || []).forEach((r, n) => {
      const row = wd.getRow(n + 2);
      const vals = [S(r.style) || null, S(r.desc) || null, S(r.category) || null, whenCell(r.when, r.estimate, true), S(r.where) || null, blankIfNone(r.units), blankIfNone(r.cases)];
      vals.forEach((v, i) => {
        const cell = row.getCell(i + 1);
        cell.value = v; cell.font = FONT; cell.border = BOX;
        cell.alignment = { vertical: 'middle', wrapText: i === 1 || i === 2, ...(i >= 5 ? { horizontal: 'right' } : i === 3 ? { horizontal: 'center' } : {}) };
        if (i >= 5) cell.numFmt = NUM;
        if (v instanceof Date) cell.numFmt = DATE;
      });
    });
    wd.views = [{ state: 'frozen', xSplit: 1, ySplit: 1, topLeftCell: 'B2' }];
    wd.autoFilter = { from: { row: 1, column: 1 }, to: { row: 1, column: DATE_COLS.length } };
    return wb;
  }

  // The workbook as .xlsx bytes (a Uint8Array).
  async function workbookBytes(wb) {
    const b = await wb.xlsx.writeBuffer();
    return b instanceof Uint8Array ? new Uint8Array(b.buffer, b.byteOffset, b.byteLength) : new Uint8Array(b);
  }

  return { salesRows, categoryTabs, buildWorkbook, workbookBytes, photoKey, fitPhoto, TITLE, TABS, NO_CAT, WHERE, COLS, DATE_COLS, PHOTO_ROW_PT, PHOTO_INSET };
});
