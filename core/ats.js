// The ERP's weekly ATS report ("Inventory Report - On Hand_In Transit_In Production <m-d-yyyy> ….xlsx", sheet "Report"):
// per style, the units in the warehouse (In WH) and the customer orders against them (Total Customer Orders), summed
// over the style's warehouse rows. Its "Grand Total" row is the check: the rows' In WH must add up to it. The tool works
// out stock after orders itself (In WH − Total Customer Orders, Saul, Oct 29); the ERP's own "ATS from in WH" is only
// compared with it. In Trans Qty (a list, one quantity per shipment) and ETA Date (a list lined up with it) are what
// core/sell.js counts as Arriving (Saul, Oct 30): each row keeps both part for part (inTransList, etaParts; an ETA that
// can't be read stays in its place as null). An In Trans quantity that isn't a number (a comma only as a thousands
// separator), could be one quantity or two ("720,816" with an ETA for each part), or is below zero, is listed. The
// Import PO Backlog list is read for cross-checks, never counted. A negative In WH or Total Customer Orders is listed
// and used as it is.
// Pure: no network, no DOM. Loads in the page (window.InvAts, after core/common.js) and in Node (module.exports).
(function (root, factory) {
  const node = typeof module === 'object' && module.exports;
  const api = factory(node ? require('./common.js') : root.InvCommon);
  if (node) module.exports = api;
  root.InvAts = api;
})(typeof globalThis !== 'undefined' ? globalThis : this, function (C) {
  'use strict';
  if (!C) throw new Error('core/common.js must load before core/ats.js');
  const { S } = C;
  const SOURCE = 'ATS report';
  const SHEET = 'Report';
  const HEADER_ROWS = 10;   // the header is looked for by name in the first 10 rows

  const SPEC = {
    style: { names: ['Style', 'Style Number', 'Style #'], required: true },
    inWh: { names: ['In WH', 'In Warehouse'], required: true },
    orders: { names: ['Total Customer Orders', 'Customer Orders'], required: true },
    warehouse: { names: ['Warehouse', 'WH'] },
    erpAts: { names: ['ATS from in WH', 'ATS from WH'] },
    desc: { names: ['Style Description', 'Description'] },
    color: { names: ['Color Code', 'Color'] },
    category: { names: ['Category'] },
    sub: { names: ['Sub Category', 'SubCategory'] },
    inTrans: { names: ['In Trans Qty', 'In Transit Qty'] },
    backlog: { names: ['Import PO Backlog Qty', 'PO Backlog Qty'] },
    eta: { names: ['ETA Date', 'ETA'] },
    etd: { names: ['ETD Date', 'ETD'] },
  };
  const LABEL = { inWh: 'In WH', orders: 'Total Customer Orders', erpAts: 'ATS from in WH', inTrans: 'In Trans Qty', backlog: 'Import PO Backlog Qty' };

  const fmtN = n => Number(n).toLocaleString('en-US');
  const unitsText = n => `${fmtN(n)} unit${n === 1 ? '' : 's'}`;
  const blank = v => S(v) === '';
  // The total row: any cell (the style's too) says "Grand Total" or just "Total", any case. Never read as a style.
  const isGrand = r => r.some(v => /^(grand\s+)?total$/i.test(S(v)));

  // The report's day from its file name: m-d-yyyy (the time after it, if any, is ignored). null when there is none.
  function dayFromName(fileName) {
    const m = S(fileName).match(/(?:^|[^\d])(\d{1,2})-(\d{1,2})-(\d{4})(?!\d)/);
    return m ? C.ymd(+m[3], +m[1], +m[2]) : null;
  }

  // A list cell ("1200, 3400") as numbers; { sum, list, bad } (bad: the parts that aren't numbers).
  function numList(v) {
    if (v == null || blank(v)) return { sum: null, list: [], bad: [] };
    if (typeof v === 'number') return Number.isFinite(v) ? { sum: v, list: [v], bad: [] } : { sum: null, list: [], bad: [String(v)] };
    const list = [], bad = [];
    // Parts are split at ", " (or ; or a line break), so a thousands comma inside "1,200" stays in its number. A comma
    // anywhere else ("1000,1400": two quantities with no space) isn't read as one number: the part is listed.
    for (const p of S(v).split(/,\s+|;|\n/).map(S).filter(Boolean)) {
      const n = p.includes(',') && !THOUSANDS.test(p) ? NaN : C.num(p);
      if (typeof n === 'number' && Number.isFinite(n)) list.push(n); else bad.push(p);
    }
    return { sum: list.length ? list.reduce((s, n) => s + n, 0) : null, list, bad };
  }
  const THOUSANDS = /^-?\d{1,3}(,\d{3})+(\.\d+)?$/;
  // A list of dates ("11/14/2026, 12/1/2026") as day numbers: list (unreadable parts dropped, counted in bad) and parts
  // (every part in its place, null where it can't be read, so it lines up with another list on the same row).
  function dayList(v, refYear) {
    if (v == null || blank(v)) return { list: [], parts: [], bad: 0 };
    if (typeof v === 'number') { const d = C.dayOf(v); return d == null ? { list: [], parts: [null], bad: 1 } : { list: [d], parts: [d], bad: 0 }; }
    const list = [], parts = []; let bad = 0;
    for (const p of S(v).split(/[;,\n]/).map(S).filter(Boolean)) { const d = C.dayOf(p, refYear); parts.push(d ?? null); if (d == null) bad++; else list.push(d); }
    return { list, parts, bad };
  }

  // XLSX: the spreadsheet library; wb: the workbook; fileName: the file's name (its day); today: a day number (dates
  // without a year take its year).
  function parseAts(XLSX, wb, fileName, today) {
    const reportDay = dayFromName(fileName);
    const base = { reportDay, reportDayFrom: reportDay != null ? 'file name' : 'none', headerRow: null, sheet: '',
      styles: new Map(), issues: [], counts: {} };
    const fail = (error, extra) => Object.assign(base, { ok: false, error }, extra || {});
    if (!wb || !Array.isArray(wb.SheetNames) || !wb.SheetNames.length) return fail(`${SOURCE}: the file has no sheets.`);

    const hit = C.findSheet(wb, SHEET) || { name: wb.SheetNames[0], ws: wb.Sheets[wb.SheetNames[0]] };
    base.sheet = hit.name;
    const rows = C.sheetRows(XLSX, hit.ws);
    // The header: the first of the first 10 rows that has every required column; otherwise the one with the most.
    let at = -1, best = null;
    for (let i = 0; i < Math.min(HEADER_ROWS, rows.length); i++) {
      const f = C.findColumns(rows[i] || [], SPEC);
      if (!f.missing.length) { at = i; best = f; break; }
      const found = Object.keys(f.cols).length;
      if (!best || found > Object.keys(best.cols).length) best = Object.assign(f, { at: i });
    }
    if (at < 0) {
      const where = best && Object.keys(best.cols).length ? `row ${best.at + 1} (headers seen: ${best.seen.join(', ') || 'none'})` : `the first ${HEADER_ROWS} rows`;
      const missing = best ? best.missing : SPEC_REQUIRED;
      return fail(`${SOURCE}: no ${missing.map(n => `"${n}"`).join(', ')} column in ${where} of sheet "${hit.name}".`, { missing });
    }
    const cols = best.cols;
    base.headerRow = at + 1;
    base.columns = Object.keys(cols);   // the fields found (e.g. 'inTrans'), so a blank cell can be told from a missing column
    const cell = (r, f) => (cols[f] == null ? null : r[cols[f]] ?? null);
    const refYear = typeof today === 'number' && Number.isFinite(today) ? C.dayYear(today) : (reportDay != null ? C.dayYear(reportDay) : undefined);

    const issues = base.issues;
    const issue = (code, style, row, message, extra) => issues.push(Object.assign({ source: SOURCE, code, po: '', style: style || '', row: row ?? '', message }, extra));
    const counts = base.counts = { rowsRead: 0, blankRows: 0, styleRows: 0, grandTotalRows: 0, noStyleRows: 0,
      styles: 0, multiWarehouseStyles: 0, notANumber: 0, blankNumbers: 0, unknownStyles: 0, inWh: 0, orders: 0,
      negativeInWh: 0, negativeOrders: 0, erpAtsDiffers: 0, inTransUnreadable: 0 };
    const styles = base.styles;
    let grand = null;
    const badNum = [];       // rows whose In WH isn't a number (the Grand Total check can't hold)
    const blankCells = [];   // style rows with a blank In WH or Total Customer Orders: { row, style, fields }
    const num = v => C.num(typeof v === 'string' ? v.replace(/^\((.*)\)$/, '-$1') : v);   // "(1,234)" is -1234

    for (let i = at + 1; i < rows.length; i++) {
      const r = rows[i] || [], row = i + 1;
      counts.rowsRead++;
      if (r.every(blank)) { counts.blankRows++; continue; }
      if (isGrand(r)) {
        counts.grandTotalRows++;
        if (grand == null) grand = { row, inWh: num(cell(r, 'inWh')), orders: num(cell(r, 'orders')) };
        else issue('grand-total', '', row, `Row ${row} is another "Grand Total" row; only the first (row ${grand.row}) is the check, and this one isn't counted.`);
        continue;
      }
      const styleRaw = S(cell(r, 'style')), style = C.normStyle(styleRaw);
      // Numbers may be stored as text ("-1234"). Text that isn't a number is listed, never read as 0; so is a blank.
      const nums = {}, bad = [], blanks = [];
      let badStock = false;   // In WH or Total Customer Orders isn't a number (not just the ERP's ATS)
      for (const f of ['inWh', 'orders', 'erpAts']) {
        if (cols[f] == null) { nums[f] = null; continue; }
        const v = cell(r, f), n = num(v);
        if (Number.isNaN(n)) { bad.push(`${LABEL[f]} "${S(v).slice(0, 20)}"`); nums[f] = NaN; if (f !== 'erpAts') badStock = true; }
        else { nums[f] = n; if (n == null && f !== 'erpAts') blanks.push(LABEL[f]); }
      }
      if (!style) {
        counts.noStyleRows++;
        // A row with units but no style: listed (they can't go to any style). An empty-looking one is only counted.
        const said = [nums.inWh ? `${unitsText(nums.inWh)} in the warehouse` : '', nums.orders ? `${unitsText(nums.orders)} on customer orders` : ''].filter(Boolean);
        if (said.length || bad.length) issue('no-style', '', row, `Row ${row} has no style${said.length ? ` but ${said.join(' and ')}` : ''}, so it isn't counted for any style.`);
        if (Number.isNaN(nums.inWh)) badNum.push(row);
        else if (nums.inWh != null) counts.inWh += nums.inWh;
        continue;
      }
      counts.styleRows++;
      if (bad.length) {
        counts.notANumber++;
        issue('not-a-number', style, row, `Style ${styleRaw} (row ${row}): ${bad.join(' and ')} ${bad.length > 1 ? "aren't numbers" : "isn't a number"}, so ${badStock ? "the style's stock after orders isn't known (never read as 0)" : "the ERP's ATS isn't compared for it"}.`);
      }
      if (blanks.length) { counts.blankNumbers++; blankCells.push({ row, style: styleRaw, fields: blanks }); }
      const inTrans = numList(cell(r, 'inTrans')), backlog = numList(cell(r, 'backlog'));
      const eta = dayList(cell(r, 'eta'), refYear), etd = dayList(cell(r, 'etd'), refYear);
      // "720,816" reads as one number with a thousands comma, but when the ETA Date list has one date per comma part
      // it's two quantities with no space: neither reading is safe, so the part is listed, never counted.
      const transText = typeof cell(r, 'inTrans') === 'string' ? S(cell(r, 'inTrans')) : '';
      const ambiguous = [];
      if (/,\S/.test(transText) && eta.parts.length > 1 && inTrans.list.length + inTrans.bad.length < eta.parts.length &&
        transText.split(/[,;\n]/).map(S).filter(Boolean).length === eta.parts.length) {
        const parts = transText.split(/,\s+|;|\n/).map(S).filter(Boolean);
        ambiguous.push(...parts.filter(p => /,\S/.test(p) && !inTrans.bad.includes(p)));
        inTrans.bad.push(...ambiguous);
        inTrans.list = parts.filter(p => !/,\S/.test(p)).map(p => C.num(p)).filter(n => typeof n === 'number' && Number.isFinite(n));
        inTrans.sum = inTrans.list.length ? inTrans.list.reduce((s, n) => s + n, 0) : null;
      }
      const wh = S(cell(r, 'warehouse'));
      const rec = { row, warehouse: wh, inWh: nums.inWh, orders: nums.orders, erpAts: nums.erpAts,
        inTrans: inTrans.sum, inTransList: inTrans.list, inTransBad: inTrans.bad, backlog: backlog.sum, backlogList: backlog.list,
        eta: eta.list, etaParts: eta.parts, etd: etd.list };
      // In Trans is counted as Arriving (core/sell.js): a part that isn't a number, or is below zero, is never read as 0.
      const negT = inTrans.list.filter(n => n < 0);
      if (inTrans.bad.length || negT.length) counts.inTransUnreadable++;
      const unknownTail = "so the ERP's In Trans for the style isn't known: its Arriving comes from our shipping data (on the water, at port) instead.";
      if (inTrans.bad.length) {
        const what = inTrans.bad.map(b => (ambiguous.includes(b) ? `"${b.slice(0, 20)}" could be one quantity or two (the ETA Date list has a date for each comma part)` : `"${b.slice(0, 20)}" isn't a number`)).join(' and ');
        issue('not-a-number', style, row, `Style ${styleRaw} (row ${row}): In Trans Qty ${what}, ${unknownTail}`);
      }
      if (negT.length) issue('negative-in-trans', style, row, `Style ${styleRaw} (row ${row}): In Trans Qty ${negT.map(n => fmtN(n)).join(' and ')} ${negT.length > 1 ? 'are' : 'is'} below zero, ${unknownTail}`);
      let s = styles.get(style);
      if (!s) styles.set(style, s = { style, styleRaw, desc: '', color: '', category: '', sub: '', inWh: 0, orders: 0, erpAts: null,
        inTrans: null, backlog: null, warehouses: [], rows: [], unknown: false });
      s.rows.push(rec);
      if (!s.desc) s.desc = S(cell(r, 'desc'));
      if (!s.color) s.color = S(cell(r, 'color'));
      if (!s.category) s.category = S(cell(r, 'category'));
      if (!s.sub) s.sub = S(cell(r, 'sub'));
      if (wh && !s.warehouses.includes(wh)) s.warehouses.push(wh);
      // A blank In WH or Total Customer Orders (the report leaves none on style rows today), or one that isn't a number,
      // makes that figure of the style unknown: never 0.
      for (const f of ['inWh', 'orders']) {
        if (nums[f] == null || Number.isNaN(nums[f])) { s[f] = null; s.unknown = true; }
        else if (s[f] != null) s[f] += nums[f];
      }
      if (Number.isNaN(nums.erpAts)) s.erpAtsBad = true;
      else if (nums.erpAts != null) s.erpAts = (s.erpAts ?? 0) + nums.erpAts;
      if (inTrans.sum != null) s.inTrans = (s.inTrans ?? 0) + inTrans.sum;
      if (backlog.sum != null) s.backlog = (s.backlog ?? 0) + backlog.sum;
      if (Number.isNaN(nums.inWh)) badNum.push(row);
      else if (nums.inWh != null) counts.inWh += nums.inWh;
      if (typeof nums.orders === 'number' && !Number.isNaN(nums.orders)) counts.orders += nums.orders;
    }
    counts.styles = styles.size;
    if (blankCells.length) {
      const ex = blankCells.slice(0, 5).map(b => `row ${b.row} (style ${b.style}: ${b.fields.join(', ')})`).join('; ');
      issue('not-a-number', blankCells.length === 1 ? C.normStyle(blankCells[0].style) : '', blankCells.length === 1 ? blankCells[0].row : '',
        `${fmtN(blankCells.length)} style row${blankCells.length > 1 ? 's have' : ' has'} a blank In WH or Total Customer Orders, so ${blankCells.length > 1 ? 'those styles\'' : 'that style\'s'} stock after orders isn't known (a blank is never read as 0): ${ex}${blankCells.length > 5 ? '…' : ''}.`);
    }

    // Per style: physical stock below 0, and the ERP's ATS from in WH against In WH − Total Customer Orders.
    for (const s of styles.values()) {
      if (s.unknown) counts.unknownStyles++;
      if (s.warehouses.length > 1 || s.rows.length > 1) counts.multiWarehouseStyles++;
      const rowsText = s.rows.map(x => x.row).join(', ');
      const neg = s.rows.filter(x => typeof x.inWh === 'number' && x.inWh < 0);
      if (neg.length || (s.inWh != null && s.inWh < 0)) {
        counts.negativeInWh++;
        issue('negative-in-wh', s.style, s.rows[0].row, `Style ${s.styleRaw}: In WH is ${neg.map(x => fmtN(x.inWh)).join(' and ') || fmtN(s.inWh)} (row${neg.length > 1 ? 's' : ''} ${(neg.length ? neg : s.rows).map(x => x.row).join(', ')}), physical stock below zero. It's used as it is, so customer orders come off lots from China sooner; check it in the ERP.`);
      }
      const negO = s.rows.filter(x => typeof x.orders === 'number' && x.orders < 0);
      if (negO.length || (s.orders != null && s.orders < 0)) {
        counts.negativeOrders++;
        issue('negative-orders', s.style, s.rows[0].row, `Style ${s.styleRaw}: Total Customer Orders is ${negO.map(x => fmtN(x.orders)).join(' and ') || fmtN(s.orders)} (row${negO.length > 1 ? 's' : ''} ${(negO.length ? negO : s.rows).map(x => x.row).join(', ')}), below zero (a credit or return?). It's used as it is, so Available now is more than In WH; check it in the ERP.`);
      }
      if (cols.erpAts != null && s.inWh != null && s.orders != null && s.erpAts != null && !s.erpAtsBad) {
        const mine = s.inWh - s.orders, erp = s.erpAts;
        if (Math.abs(mine - erp) > 1e-9) {
          counts.erpAtsDiffers++;
          issue('erp-ats-differs', s.style, s.rows[0].row, `Style ${s.styleRaw} (row${s.rows.length > 1 ? 's' : ''} ${rowsText}): the ERP's "ATS from in WH" is ${fmtN(erp)}, but In WH ${fmtN(s.inWh)} − Total Customer Orders ${fmtN(s.orders)} = ${fmtN(mine)}. The tool uses ${fmtN(mine)}.`);
        }
      }
      delete s.erpAtsBad;
    }

    if (!styles.size) return fail(`${SOURCE}: no style rows below the header (row ${base.headerRow}) of sheet "${hit.name}".`);
    // The Grand Total row is the check: the style rows' In WH must add up to it.
    if (!grand) {
      issue('grand-total', '', '', `The report has no "Grand Total" row, so its In WH (${unitsText(counts.inWh)} over ${fmtN(counts.styleRows)} style rows) can't be checked against the ERP's own total.`);
    } else if (grand.inWh == null || Number.isNaN(grand.inWh) || Math.abs(grand.inWh - counts.inWh) > 1e-6 || badNum.length) {
      const why = badNum.length ? ` (${fmtN(badNum.length)} row${badNum.length > 1 ? 's have' : ' has'} an In WH that isn't a number: row${badNum.length > 1 ? 's' : ''} ${badNum.slice(0, 5).join(', ')}${badNum.length > 5 ? '…' : ''})` : '';
      const gt = grand.inWh == null || Number.isNaN(grand.inWh) ? 'blank or not a number' : fmtN(grand.inWh);
      return fail(`${SOURCE}: the report doesn't add up: rows ${fmtN(counts.inWh)}, Grand Total ${gt}${why}.`, { grandTotal: grand });
    }
    return Object.assign(base, { ok: true, error: '', grandTotal: grand });
  }
  const SPEC_REQUIRED = Object.entries(SPEC).filter(([, f]) => f.required).map(([, f]) => f.names[0]);

  return { parseAts, dayFromName, SOURCE, SHEET };
});
