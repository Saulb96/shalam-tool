// The sales sheets (Saul, Nov 2; remade Nov 3; a line per pack Nov 4): what a salesperson may see, as two Excel files
// saved into the one library they can open.
// - On hand ATS: the ATS report's own numbers, style by style (Saul, Nov 3: "that is the gold standard of our ATS").
//   Available on hand is its "ATS from in WH" (0 when below 0: oversold). Available in transit is its In Trans
//   shipments, every one, with what the style is oversold taken off its earliest ETA first; each is available ETA + 7
//   days ("now" once that day has come). Total = on hand + in transit. Only styles on the report; our shipping sheets
//   add nothing to it.
// - Pull from China: our goods still in China that aren't sold, one row per style, pull-from date and pack
//   (core/sell.js: 10 days before the ETD; no ETD, 5 days before the ready date; no ready date, about 75 days after the
//   order date). What a style is oversold beyond its In Trans comes off them, earliest pull-from date first (the orders
//   that ship from China with goods sold there aside: core/sell.js). A pull-from date before today is too late: left off
//   (Saul, Nov 3).
// A style in more than one pack (Master/Inner) has a line per pack (Saul, Nov 4: "split by PO"). The ATS report has no
// pack, so each In Trans shipment takes the pack of our goods core/sell.js matched it to, and the stock goes to the
// packs of our goods landed, the latest first (older goods ship first: a best guess); what neither places is in the
// style's own pack (its Master Sheet Live Master). Goods on a PO ending in a suffix named in config.js's reserved (held
// for a customer) are their own line, marked with its note (Saul, Nov 4: "-BCF", "put as reserved").
// Per row only: style, UPCs, description, brand, category and sub category (the ATS report's, else UPC_MASTER's, the
// pair from one place), Master/Inner, units (never cases), the date, the note and a photo. Never a factory, customer,
// PO, PO type, section, container, round, person, file name, SharePoint link, ERP warehouse, In WH, customer orders,
// sold, oversold, flag or problem: every field is copied by name (an allow-list), never by spreading an object of
// Available to sell.
// - salesRows(sell, { today, reserved }) / chinaRows(sell, { today, reserved }): the rows (and On hand ATS's schedule by
//   date) from InvSell.buildSell's result; none (ok false, with the reason in plain words) when there's no usable ATS
//   report, Available to sell wasn't worked out, or its units don't add up. reserved: config.js's { suffix: note }.
// - buildWorkbook(ExcelJS, sheet, photos, { print }): the Excel file (ExcelJS 4.4.0, vendor/exceljs): every row (a photo
//   per row fitted in its cell the way Quote Tracker does it), then a tab per category (categoryTabs: A to Z, named as
//   the ERP has them, never joined), then By date (On hand ATS only); the page's fingerprint of it in its keywords.
//   Each picture is stored once in the file, however many tabs show it.
// - readWorkbook(wb): one of these files read back (the sales page, sales.js): its kind, rows, schedule by date and the
//   rows' photos, as buildWorkbook was given them.
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

  const TITLE = 'M1S · On hand ATS';
  const TABS = { sell: 'On hand ATS', byDate: 'By date', china: 'Pull from China' };
  // In transit is available this many days after its ETA (Saul, Nov 3: "ETA date + 7 days to be clean").
  const TRANSIT_DAYS = 7;
  // Where a schedule row's goods are, in the order they're listed on the same date.
  const WHERE = ['On hand', 'In transit'];
  const isDay = v => typeof v === 'number' && Number.isFinite(v);
  const units = v => (isDay(v) && v > 0 ? v : 0);
  // Quote Tracker's photo key: the style uppercased, only A-Z, 0-9, _ and - kept. Exact key only (no fallback from
  // "X-F" to "X").
  const photoKey = s => String(s ?? '').toUpperCase().replace(/[^A-Z0-9_-]/g, '');
  const byStyle = (a, b) => (a < b ? -1 : a > b ? 1 : 0);
  // Categories and sub categories A to Z (Cosmetics1 before Cosmetics2, case aside), a blank one last.
  const byName = (a, b) => (!a) - (!b) || a.localeCompare(b, 'en', { sensitivity: 'base', numeric: true }) || byStyle(a, b);
  // A date: now first, then by day, no date last.
  const whenRank = w => (w === 'now' ? -Infinity : isDay(w) ? w : Infinity);
  // Master/Inner as the sheets show it: "48/12", "48/–" with no Inner, each pack of a style "48/12 · 24/6".
  const packText = packs => (Array.isArray(packs) ? packs : []).filter(p => p && isDay(p.master) && p.master > 0)
    .map(p => `${p.master}/${isDay(p.inner) && p.inner > 0 ? p.inner : '–'}`).join(' · ');

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

  // A style's category and sub category: the ATS report's (Saul, Nov 3), else the page's (UPC_MASTER, then Master
  // Sheet Live); both from the same place, never one from each. Rows go by category, sub category, then style: each
  // category's tab and the By date tab keep the same order.
  const catOf = s => (S(s.atsCategory) ? [S(s.atsCategory), S(s.atsSub)] : [S(s.category), S(s.sub)]);
  const order = (a, b) => { const x = catOf(a), y = catOf(b); return byName(x[0], y[0]) || byName(x[1], y[1]) || byStyle(a.style, b.style); };
  // What every row of a style carries, copied by name.
  const base = s => {
    const style = S(s.styleRaw || s.style), [category, sub] = catOf(s);
    return { style, upc: (Array.isArray(s.upcs) ? s.upcs : []).map(S).filter(Boolean).join(', '), desc: S(s.desc), brand: S(s.brand),
      category, sub, photoKey: photoKey(style) };
  };
  // A pack as { master, inner } (null for none), and the key of a line: its pack and note.
  const packOf = p => (p && isDay(p.master) && p.master > 0 ? { master: p.master, inner: isDay(p.inner) && p.inner > 0 ? p.inner : null } : null);
  const lineKey = (pack, note) => `${pack ? pack.master : ''}/${pack ? pack.inner ?? '' : ''}|${note}`;
  const lineOrder = (a, b) => (!a.pack) - (!b.pack) || (a.pack ? a.pack.master - b.pack.master || (a.pack.inner ?? 0) - (b.pack.inner ?? 0) : 0) ||
    (!!a.note) - (!!b.note) || byStyle(a.note, b.note);
  // config.js's reserved ({ "-SUFFIX": "its note" }): a PO's note ('' for none), matched at the PO's end, case aside.
  const reservedOf = opts => Object.entries((opts && opts.reserved && typeof opts.reserved === 'object') ? opts.reserved : {})
    .map(([k, v]) => [S(k).toUpperCase(), S(v)]).filter(([k, v]) => k && v);
  const noteOf = (list, po) => { const p = S(po).toUpperCase(); const hit = p && list.find(([k]) => p.endsWith(k)); return hit ? hit[1] : ''; };

  // A style's stock by the ATS report: { known, onHand, transit: [{ units, when, parts }], transitKnown, left }. Its "ATS
  // from in WH" (else In WH − Total Customer Orders, when the report has no such figure for it); what's below 0
  // (oversold) comes off its In Trans shipments, earliest ETA first (no ETA last), off each one's parts in turn; left:
  // what's still oversold after them. A shipment's parts: our goods core/sell.js matched it to ({ units, master, inner,
  // po }), then the rest of it (master null: the style's own pack).
  function erpStock(s, today) {
    const ats = isDay(s.erpAts) ? s.erpAts : isDay(s.inWh) && isDay(s.orders) ? s.inWh - s.orders : null;
    if (ats == null) return { known: false, onHand: 0, transit: [], transitKnown: false, left: 0 };
    let left = Math.max(0, -ats);
    const ships = (Array.isArray(s.erpTransit) ? s.erpTransit : []).filter(x => x && units(x.units) > 0)
      .map(x => {
        const parts = [];
        let rest = x.units;
        for (const p of Array.isArray(x.parts) ? x.parts : []) {
          const n = Math.min(units(p && p.units), rest);
          if (n > 0) { parts.push({ units: n, master: p.master, inner: p.inner, po: S(p.po) }); rest -= n; }
        }
        if (rest > 0) parts.push({ units: rest, master: null, inner: null, po: '' });
        return { units: x.units, eta: isDay(x.eta) ? x.eta : null, parts };
      })
      .sort((a, b) => (a.eta == null) - (b.eta == null) || (a.eta ?? 0) - (b.eta ?? 0));
    const transit = [];
    for (const x of ships) {
      const parts = [];
      for (const p of x.parts) {
        const off = Math.min(left, p.units);
        left -= off;
        if (p.units - off > 0) parts.push({ ...p, units: p.units - off });
      }
      const n = parts.reduce((t, p) => t + p.units, 0);
      if (n > 0) {
        const day = x.eta == null ? null : x.eta + TRANSIT_DAYS;
        transit.push({ units: n, when: day == null ? null : day <= today ? 'now' : day, parts });
      }
    }
    return { known: true, onHand: Math.max(0, ats), transit, transitKnown: Array.isArray(s.erpTransit), left };
  }
  const todayOf = (sell, opts) => {
    const today = opts && opts.today != null ? opts.today : sell && sell.today;
    if (!isDay(today)) throw new TypeError('InvSalesSheet: today must be a day number');
    return today;
  };

  // sell: InvSell.buildSell's result. opts.today: the day the sheet is made (a day number; defaults to sell's own);
  // opts.reserved: config.js's { suffix: note }.
  // Answers { ok, reason, kind: 'sales', asOf, reportDay, rows, byDate, totals, counts }:
  // rows, one per style on the ATS report with something available and pack (and note: above), by category, sub
  //   category, style, then pack: { style, upc, desc, brand, category, sub, pack (Master/Inner text, '' when not known),
  //   note, onHand, inTransit, transitFrom ('now', a day number, or null: no ETA), total, photoKey }.
  // byDate, one per row per date and where (On hand, In transit), in the rows' order then date (now first, no date
  //   last): { style, desc, category, pack, note, when ('now', a day number, or null), where, units }.
  // totals: { onHand, inTransit, total } (units). counts: { styles, rows, split (styles on more than one row), reserved
  //   (rows with a note), leftOff (nothing available), stockUnknown (no ATS figure: left off), transitUnknown (on hand
  //   only: its In Trans can't be read), notOnReport, byDateRows, noPack (rows) }.
  function salesRows(sell, opts) {
    const today = todayOf(sell, opts), reserved = reservedOf(opts);
    const counts = { styles: 0, rows: 0, split: 0, reserved: 0, leftOff: 0, stockUnknown: 0, transitUnknown: 0, notOnReport: 0, byDateRows: 0, noPack: 0 };
    const totals = { onHand: 0, inTransit: 0, total: 0 };
    const reportDay = sell && isDay(sell.reportDay) ? sell.reportDay : null;
    const why = refusal(sell);
    if (why) return { ok: false, reason: why, kind: 'sales', asOf: today, reportDay, rows: [], byDate: [], totals, counts };

    const rows = [], byDate = [];
    for (const s of sell.styles.slice().sort(order)) {
      if (s.onReport !== true) { counts.notOnReport++; continue; }
      const k = erpStock(s, today);
      if (!k.known) { counts.stockUnknown++; continue; }
      if (!k.transitKnown) counts.transitUnknown++;
      const inTransit = k.transit.reduce((t, x) => t + x.units, 0), total = k.onHand + inTransit;
      if (!(total > 0)) { counts.leftOff++; continue; }
      const b = base(s), own = packOf(s.stockPack);
      // The style's lines: pack and note -> { pack, note, onHand, dates (when -> units) }.
      const lines = new Map();
      const lineOf = (p, po) => {
        const pack = packOf(p) || own, note = noteOf(reserved, po), key = lineKey(pack, note);
        return lines.get(key) || lines.set(key, { pack, note, onHand: 0, dates: new Map() }).get(key);
      };
      // The stock: our Landed goods, the latest first, each up to its units; the rest in the style's own pack.
      let stock = k.onHand;
      for (const l of Array.isArray(s.landedPacks) ? s.landedPacks : []) {
        if (stock <= 0) break;
        const n = Math.min(units(l && l.units), stock);
        if (n > 0) { lineOf(l, l.po).onHand += n; stock -= n; }
      }
      if (stock > 0) lineOf(null, '').onHand += stock;
      for (const x of k.transit) for (const p of x.parts) { const L = lineOf(p, p.po); L.dates.set(x.when, (L.dates.get(x.when) || 0) + p.units); }
      const list = [...lines.values()].sort(lineOrder);
      for (const L of list) {
        const pack = packText(L.pack ? [L.pack] : []);
        const dates = [...L.dates.keys()].sort((p, q) => whenRank(p) - whenRank(q));
        const lt = dates.reduce((t, w) => t + L.dates.get(w), 0);
        if (L.onHand > 0) byDate.push({ style: b.style, desc: b.desc, category: b.category, pack, note: L.note, when: 'now', where: WHERE[0], units: L.onHand });
        for (const w of dates) byDate.push({ style: b.style, desc: b.desc, category: b.category, pack, note: L.note, when: w, where: WHERE[1], units: L.dates.get(w) });
        rows.push({ style: b.style, upc: b.upc, desc: b.desc, brand: b.brand, category: b.category, sub: b.sub, pack, note: L.note,
          onHand: L.onHand, inTransit: lt, transitFrom: dates.length ? dates[0] : null, total: L.onHand + lt, photoKey: b.photoKey });
        if (!pack) counts.noPack++;
        if (L.note) counts.reserved++;
      }
      counts.styles++;
      if (list.length > 1) counts.split++;
      totals.onHand += k.onHand; totals.inTransit += inTransit; totals.total += total;
    }
    counts.rows = rows.length; counts.byDateRows = byDate.length;
    return { ok: true, reason: '', kind: 'sales', asOf: today, reportDay, rows, byDate, totals, counts };
  }

  // sell, opts: as salesRows. Answers { ok, reason, kind: 'china', asOf, reportDay, rows, totals, counts }:
  // rows, one per style, pull-from date and pack (and note), by category, sub category, style, then date (no date last),
  //   then pack: { style, upc, desc, brand, category, sub, pack (Master/Inner of those goods), note, units, pullFrom (a
  //   day number, or null: no ETD, ready or order date), estimate (that date is about 75 days after the order date),
  //   photoKey }.
  // totals: { units }. counts: { rows, styles, reserved (rows with a note), ordersOff (units the oversold took), past
  //   (units whose pull-from date is before today: left off), noDate (units with no date), stockUnknown (styles on the
  //   report with no ATS figure: left off, their orders not known) }.
  function chinaRows(sell, opts) {
    const today = todayOf(sell, opts), reserved = reservedOf(opts);
    const counts = { rows: 0, styles: 0, reserved: 0, ordersOff: 0, past: 0, noDate: 0, stockUnknown: 0 };
    const totals = { units: 0 };
    const reportDay = sell && isDay(sell.reportDay) ? sell.reportDay : null;
    const why = refusal(sell);
    if (why) return { ok: false, reason: why, kind: 'china', asOf: today, reportDay, rows: [], totals, counts };

    const rows = [], styles = new Set();
    for (const s of sell.styles.slice().sort(order)) {
      // Our lots still in China, as they stood before customer orders (core/sell.js took orders off them by its own
      // rule): units = available + taken.
      const lots = (Array.isArray(s.lots) ? s.lots : []).filter(l => l && l.inChina && l.source !== 'erp')
        .map(l => ({ units: units(l.available) + units(l.taken), pullFrom: isDay(l.pullFrom) ? l.pullFrom : null, estimate: !!l.estimate,
          master: l.master, inner: l.inner, po: S(l.po) }))
        .filter(l => l.units > 0)
        .sort((a, b) => (a.pullFrom == null) - (b.pullFrom == null) || (a.pullFrom ?? 0) - (b.pullFrom ?? 0));
      if (!lots.length) continue;
      // What the style is oversold beyond its In Trans, less the orders that ship from China with goods sold there.
      let left = 0;
      if (s.onReport === true) {
        const k = erpStock(s, today);
        if (!k.known) { counts.stockUnknown++; continue; }
        left = Math.max(0, k.left - units(s.ordersShipChina));
      }
      const own = packOf(s.stockPack);
      // Date, then pack and note -> { day, pack, note, units, estimate }.
      const groups = new Map();
      for (const l of lots) {
        const off = Math.min(left, l.units);
        left -= off; counts.ordersOff += off;
        const n = l.units - off;
        if (!(n > 0)) continue;
        if (l.pullFrom != null && l.pullFrom < today) { counts.past += n; continue; }
        if (l.pullFrom == null) counts.noDate += n;
        const pack = packOf(l) || own, note = noteOf(reserved, l.po), key = `${l.pullFrom}|${lineKey(pack, note)}`;
        const g = groups.get(key) || groups.set(key, { day: l.pullFrom, pack, note, units: 0, estimate: false }).get(key);
        g.units += n; if (l.estimate) g.estimate = true;
      }
      if (!groups.size) continue;
      const b = base(s);
      for (const g of [...groups.values()].sort((p, q) => whenRank(p.day) - whenRank(q.day) || lineOrder(p, q))) {
        rows.push({ style: b.style, upc: b.upc, desc: b.desc, brand: b.brand, category: b.category, sub: b.sub, pack: packText(g.pack ? [g.pack] : []),
          note: g.note, units: g.units, pullFrom: g.day, estimate: g.estimate, photoKey: b.photoKey });
        totals.units += g.units;
        if (g.note) counts.reserved++;
      }
      styles.add(b.style);
    }
    counts.rows = rows.length; counts.styles = styles.size;
    return { ok: true, reason: '', kind: 'china', asOf: today, reportDay, rows, totals, counts };
  }

  // ---------- the Excel files
  // A day number as a date cell (UTC midnight: ExcelJS turns it into that day's serial, whatever the time zone).
  const dateOf = d => new Date(d * DAY);
  const blankIfNone = v => (isDay(v) && v !== 0 ? v : null);
  // A date as the sheet shows it: now, about m/d/yyyy (an estimate), the date itself, or (none) "date to come".
  const whenCell = (when, estimate) => (when === 'now' ? 'now' : isDay(when) ? (estimate ? `about ${fmt(when)}` : dateOf(when)) : 'date to come');
  // A tab's columns: header, width (characters), the row's field (k), and how it's shown: n (a number, right), when (a
  // date, centred: get gives [when, estimate], or null for none), center.
  const COLS = [
    { h: 'Photo', w: 15, photo: true },
    { h: 'Style', w: 16, k: 'style' },
    { h: 'UPC', w: 16, k: 'upc', wrap: true },
    { h: 'Description', w: 34, k: 'desc', wrap: true },
    { h: 'Brand', w: 14, k: 'brand', wrap: true },
    { h: 'Category', w: 14, k: 'category', wrap: true },
    { h: 'Sub category', w: 16, k: 'sub', wrap: true },
    { h: 'Master/Inner', w: 11, k: 'pack', center: true, wrap: true },
    { h: 'Available on hand (units)', w: 12, k: 'onHand', n: true },
    { h: 'Available in transit (units)', w: 12, k: 'inTransit', n: true },
    { h: 'In transit available from', w: 13, k: 'transitFrom', when: r => (r.inTransit > 0 ? [r.transitFrom, false] : null) },
    { h: 'Total available (units)', w: 12, k: 'total', n: true },
    { h: 'Notes', w: 16, k: 'note', wrap: true },
  ];
  const CHINA_COLS = [
    ...COLS.slice(0, 8),
    { h: 'Units', w: 12, k: 'units', n: true },
    { h: 'Pull from date', w: 16, k: 'pullFrom', when: r => [r.pullFrom, r.estimate] },
    COLS[COLS.length - 1],
  ];
  const DATE_COLS = [
    { h: 'Style', w: 16, k: 'style' }, { h: 'Description', w: 34, k: 'desc', wrap: true }, { h: 'Category', w: 14, k: 'category', wrap: true },
    { h: 'Master/Inner', w: 11, k: 'pack', center: true }, { h: 'Available from', w: 16, k: 'when', when: true }, { h: 'Where', w: 12, k: 'where' },
    { h: 'Units', w: 11, k: 'units', n: true }, { h: 'Notes', w: 16, k: 'note', wrap: true },
  ];
  // Each file: its title, its first tab, its columns, and whether it has a By date tab.
  const KINDS = {
    sales: { title: TITLE, tab: TABS.sell, cols: COLS, byDate: true },
    china: { title: 'M1S · Pull from China', tab: TABS.china, cols: CHINA_COLS, byDate: false },
  };
  // A sheet's file: its kind's, Available to sell's when it says none.
  const kindOf = sheet => KINDS[(sheet && sheet.kind) || 'sales'] || KINDS.sales;
  // The tabs a sheet's file has, in order.
  const tabsOf = sheet => { const k = kindOf(sheet); return [k.tab, ...categoryTabs(sheet.rows).map(t => t.name), ...(k.byDate ? [TABS.byDate] : [])]; };
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

  // ExcelJS: the library (window.ExcelJS, or require('vendor/exceljs/exceljs.min.js') in Node). sheet: salesRows's or
  // chinaRows's answer (ok, with rows). photos: a Map (or object) of photoKey -> { data (the image's bytes: ArrayBuffer
  // or Uint8Array), width, height (pixels), extension ('jpeg' when not given) }, or a list of those in the rows' order
  // (null for none); a row with none gets an empty cell.
  // opts.print: the page's fingerprint of what the file holds (a hash: letters and digits only), kept in the file's
  // keywords, so a later look at the file in the library can tell it holds the same sheet whoever saved it.
  // Answers the ExcelJS Workbook (workbookBytes writes it).
  function buildWorkbook(ExcelJS, sheet, photos, opts) {
    if (!ExcelJS || typeof ExcelJS.Workbook !== 'function') throw new Error("The Excel library didn't load, so the sales sheet can't be made.");
    if (!sheet || sheet.ok !== true) throw new Error((sheet && sheet.reason) || "There's no sales sheet to write.");
    if (!Array.isArray(sheet.rows) || !sheet.rows.length) throw new Error('Nothing is available, so the sheet would be empty.');
    const print = S(opts && opts.print);
    if (print && !/^[0-9a-z]{1,128}$/i.test(print)) throw new Error("The sales sheet's fingerprint isn't a hash, so it isn't written into the file.");
    const kind = kindOf(sheet), cols = kind.cols;
    const wb = new ExcelJS.Workbook();
    wb.creator = 'Shalam Tool'; wb.lastModifiedBy = 'Shalam Tool';
    if (print) wb.keywords = print;

    // A tab of rows: the title row, the header row, the rows (rowIdx: their places in sheet.rows). Each picture goes
    // into the file once (by its photo key: a style on several rows shows the same one), however many tabs show it.
    const withPhotos = !!photos && (Array.isArray(photos) ? photos.some(Boolean) : photos instanceof Map ? photos.size > 0 : Object.keys(photos).length > 0);
    const picCol = cols.findIndex(c => c.photo);
    const ids = new Map();
    const imageOf = n => {
      const key = sheet.rows[n].photoKey || `#${n}`;
      if (ids.has(key)) return ids.get(key);
      const p = photoOf(photos, sheet.rows[n].photoKey, n);
      const img = !p || !p.data ? null : { id: wb.addImage({ buffer: p.data, extension: p.extension || 'jpeg' }),
        d: isDay(p.width) && p.width > 0 && isDay(p.height) && p.height > 0 ? { w: p.width, h: p.height } : { w: 160, h: 160 } };
      ids.set(key, img);
      return img;
    };
    const rowsTab = (name, heading, rowIdx) => {
      const ws = wb.addWorksheet(name);
      ws.columns = cols.map(c => ({ width: c.w }));
      const title = ws.getRow(1);
      title.getCell(1).value = heading; title.getCell(1).font = { ...FONT, size: 14, bold: true };
      title.height = 24;
      const head = ws.getRow(2);
      cols.forEach((c, i) => {
        const cell = head.getCell(i + 1);
        cell.value = c.h; cell.font = { ...FONT, bold: true }; cell.fill = HEAD_FILL; cell.border = BOX;
        cell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
      });
      head.height = 45;
      rowIdx.forEach((at, n) => {
        const r = sheet.rows[at], row = ws.getRow(n + 3);
        cols.forEach((c, i) => {
          const cell = row.getCell(i + 1);
          cell.font = FONT; cell.border = BOX;
          cell.alignment = { vertical: 'middle', wrapText: !!c.wrap, ...(c.n ? { horizontal: 'right' } : c.when || c.center ? { horizontal: 'center' } : {}) };
          if (c.photo) return;
          if (c.when) { const w = c.when(r); cell.value = w ? whenCell(w[0], w[1]) : null; }
          else if (c.n) cell.value = blankIfNone(r[c.k]);
          else cell.value = S(r[c.k]) || null;
          if (c.n) cell.numFmt = NUM;
          if (cell.value instanceof Date) cell.numFmt = DATE;
        });
        if (withPhotos) row.height = PHOTO_ROW_PT;
        const img = imageOf(at);
        if (img) ws.addImage(img.id, fitPhoto(picCol, cols[picCol].w, PHOTO_ROW_PT, n + 2, img.d));
      });
      ws.views = [{ state: 'frozen', xSplit: 2, ySplit: 2, topLeftCell: 'C3' }];
      ws.autoFilter = { from: { row: 2, column: 1 }, to: { row: 2, column: cols.length } };
      ws.pageSetup = { orientation: 'landscape', fitToPage: true, fitToWidth: 1, fitToHeight: 0, printTitlesRow: '2:2' };
    };

    // 1. Every row. 2. A tab per category.
    rowsTab(kind.tab, kind.title, sheet.rows.map((r, n) => n));
    for (const t of categoryTabs(sheet.rows)) rowsTab(t.name, `${kind.title} · ${t.category || NO_CAT}`, t.rows);
    if (!kind.byDate) return wb;

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
      DATE_COLS.forEach((c, i) => {
        const cell = row.getCell(i + 1);
        cell.value = c.when ? whenCell(r.when, false) : c.n ? blankIfNone(r[c.k]) : S(r[c.k]) || null;
        cell.font = FONT; cell.border = BOX;
        cell.alignment = { vertical: 'middle', wrapText: !!c.wrap, ...(c.n ? { horizontal: 'right' } : c.when || c.center ? { horizontal: 'center' } : {}) };
        if (c.n) cell.numFmt = NUM;
        if (cell.value instanceof Date) cell.numFmt = DATE;
      });
    });
    wd.views = [{ state: 'frozen', xSplit: 1, ySplit: 1, topLeftCell: 'B2' }];
    wd.autoFilter = { from: { row: 1, column: 1 }, to: { row: 1, column: DATE_COLS.length } };
    return wb;
  }

  // ---------- one of these files read back (the sales page)
  // A cell's value as buildWorkbook wrote it: text, a number (blank: 0), or a date cell's day ('now', 'about m/d/yyyy'
  // is that day as an estimate, 'date to come' and blank are none).
  const textOf = v => (v == null ? '' : v instanceof Date ? '' : typeof v === 'object' ? S(v.text ?? (Array.isArray(v.richText) ? v.richText.map(x => x.text).join('') : v.result ?? '')) : S(v));
  const numOf = v => (typeof v === 'number' && Number.isFinite(v) ? v : typeof v === 'object' && v && typeof v.result === 'number' ? v.result : 0);
  const dayOfDate = d => Math.round(d.getTime() / DAY);
  function whenOf(v) {
    if (v instanceof Date && !Number.isNaN(v.getTime())) return { when: dayOfDate(v), estimate: false };
    const t = textOf(v);
    if (t === 'now') return { when: 'now', estimate: false };
    const m = /^about (\d{1,2})\/(\d{1,2})\/(\d{4})$/.exec(t);
    if (m) return { when: C.ymd(+m[3], +m[1], +m[2]), estimate: true };
    return { when: null, estimate: false };
  }
  // The fields of a row (and of a schedule row) in the order salesRows and chinaRows give them.
  const ROW_KEYS = {
    sales: ['style', 'upc', 'desc', 'brand', 'category', 'sub', 'pack', 'note', 'onHand', 'inTransit', 'transitFrom', 'total', 'photoKey'],
    china: ['style', 'upc', 'desc', 'brand', 'category', 'sub', 'pack', 'note', 'units', 'pullFrom', 'estimate', 'photoKey'],
  };
  const DATE_KEYS = ['style', 'desc', 'category', 'pack', 'note', 'when', 'where', 'units'];
  // wb: an ExcelJS Workbook loaded from one of these files. Answers { kind ('sales' | 'china'), rows, byDate (On hand
  // ATS's; [] for Pull from China), photos (in the rows' order: { data, extension, width, height } or null) }, the rows
  // as salesRows / chinaRows gave them. Throws (in plain words) when it isn't one of them.
  function readWorkbook(wb) {
    const ws = wb && wb.worksheets && wb.worksheets[0];
    const kind = ws && Object.keys(KINDS).find(k => KINDS[k].tab === ws.name);
    const K = kind && KINDS[kind];
    const heads = ws ? K && K.cols.map((c, i) => textOf(ws.getRow(2).getCell(i + 1).value)) : null;
    if (!K || heads.join('|') !== K.cols.map(c => c.h).join('|')) throw new Error("This file isn't one of the sales sheets Shalam Tool makes (or it was changed), so it can't be shown.");
    const rows = [];
    for (let n = 3; n <= ws.rowCount; n++) {
      const row = ws.getRow(n), r = {};
      if (!textOf(row.getCell(K.cols.findIndex(c => c.k === 'style') + 1).value)) break;
      K.cols.forEach((c, i) => {
        if (c.photo) return;
        const v = row.getCell(i + 1).value;
        if (c.when) {
          const w = whenOf(v);
          if (c.k === 'pullFrom') { r.pullFrom = w.when === 'now' ? null : w.when; r.estimate = w.estimate; } else r[c.k] = w.when;
        } else r[c.k] = c.n ? numOf(v) : textOf(v);
      });
      if (kind === 'sales' && !(r.inTransit > 0)) r.transitFrom = null;
      r.photoKey = photoKey(r.style);
      rows.push(Object.fromEntries(ROW_KEYS[kind].map(k => [k, r[k]])));
    }
    // Each row's photo: the picture anchored in its Photo cell.
    const photos = rows.map(() => null), pic = K.cols.findIndex(c => c.photo);
    for (const im of typeof ws.getImages === 'function' ? ws.getImages() : []) {
      const tl = im && im.range && im.range.tl;
      const at = tl ? Math.floor(tl.nativeRow ?? tl.row) - 2 : -1, col = tl ? Math.floor(tl.nativeCol ?? tl.col) : -1;
      if (col !== pic || at < 0 || at >= rows.length || photos[at]) continue;
      const img = wb.getImage(Number(im.imageId));
      const data = img && (img.buffer || img.base64);
      if (!data) continue;
      const ext = im.range.ext;
      photos[at] = { data: img.buffer ? new Uint8Array(img.buffer.buffer || img.buffer, img.buffer.byteOffset || 0, img.buffer.byteLength) : data, extension: img.extension || 'jpeg',
        width: ext && ext.width > 0 ? ext.width : 160, height: ext && ext.height > 0 ? ext.height : 160 };
    }
    const byDate = [];
    const wd = kind === 'sales' ? wb.getWorksheet(TABS.byDate) : null;
    if (wd) {
      for (let n = 2; n <= wd.rowCount; n++) {
        const row = wd.getRow(n), r = {};
        if (!textOf(row.getCell(1).value)) break;
        DATE_COLS.forEach((c, i) => { const v = row.getCell(i + 1).value; r[c.k] = c.when ? whenOf(v).when : c.n ? numOf(v) : textOf(v); });
        byDate.push(Object.fromEntries(DATE_KEYS.map(k => [k, r[k]])));
      }
    }
    return { kind, rows, byDate, photos };
  }

  // The workbook as .xlsx bytes (a Uint8Array).
  async function workbookBytes(wb) {
    const b = await wb.xlsx.writeBuffer();
    return b instanceof Uint8Array ? new Uint8Array(b.buffer, b.byteOffset, b.byteLength) : new Uint8Array(b);
  }

  return { salesRows, chinaRows, categoryTabs, tabsOf, kindOf, buildWorkbook, readWorkbook, workbookBytes, photoKey, packText, fitPhoto,
    TITLE, TABS, KINDS, NO_CAT, WHERE, COLS, CHINA_COLS, DATE_COLS, TRANSIT_DAYS, PHOTO_ROW_PT, PHOTO_INSET };
});
