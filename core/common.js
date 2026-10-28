// Shared helpers for the parsers: text, numbers, dates, header lookup and sheet reading. Pure: no network, no DOM,
// and "today" is always passed in. Loads in the page (window.InvCommon) and in Node (module.exports).
(function (root, factory) {
  const api = factory();
  if (typeof module === 'object' && module.exports) module.exports = api;
  root.InvCommon = api;
})(typeof globalThis !== 'undefined' ? globalThis : this, function () {
  'use strict';
  const DAY = 864e5;

  // Text with the ends trimmed, byte-order marks and non-breaking spaces included.
  const S = v => String(v ?? '').replace(/^[\s﻿ ]+|[\s﻿ ]+$/g, '');

  // null for a blank cell, NaN for text that isn't a number (the caller lists it; it never becomes 0).
  function num(v) {
    if (v == null) return null;
    if (typeof v === 'number') return Number.isFinite(v) ? v : NaN;
    if (typeof v === 'boolean') return NaN;
    const s = S(v).replace(/,/g, '');
    if (s === '') return null;
    const n = Number(s);
    return Number.isFinite(n) ? n : NaN;
  }

  // Styles keep inner spaces and -F endings: "RM 2201" and "RM2201" are different items (Saul, Oct 26).
  const normStyle = v => S(v).toUpperCase();
  // PO numbers never hold spaces; suffixes such as -BCF or -1 are kept and only ever match themselves.
  const normPo = v => S(v).toUpperCase().replace(/\s+/g, '');
  // A UPC is text of 12 or 13 digits. Anything else is null (the caller lists it); it is never reshaped.
  function normUpc(v) {
    const s = typeof v === 'number' && Number.isInteger(v) ? String(v) : S(v);
    return /^\d{12,13}$/.test(s) ? s : null;
  }

  // Header lookup by name. Names are compared on letters and digits only, ignoring case, so "Container #",
  // "Container#" and "container" match. spec: { field: { names: [...], required: true|false } }.
  // Returns { cols: { field: index }, missing: [first name of each missing required field], seen: [header texts] }.
  const headerKey = h => S(h).toLowerCase().replace(/[^a-z0-9]/g, '');
  function findColumns(headerRow, spec) {
    const keys = (headerRow || []).map(headerKey);
    const cols = {}, missing = [];
    for (const [field, f] of Object.entries(spec)) {
      let at = -1;
      for (const name of f.names) {
        at = keys.indexOf(headerKey(name));
        if (at >= 0) break;
      }
      if (at >= 0) cols[field] = at;
      else if (f.required) missing.push(f.names[0]);
    }
    return { cols, missing, seen: (headerRow || []).map(S).filter(Boolean) };
  }

  // Dates are day numbers (days since 1970-01-01, UTC), or null.
  function ymd(y, m, d) {
    if (!(y >= 1990 && y <= 2200 && m >= 1 && m <= 12 && d >= 1 && d <= 31)) return null;
    const t = Date.UTC(y, m - 1, d);
    return new Date(t).getUTCDate() === d ? t / DAY : null;
  }
  const dayYear = d => new Date(d * DAY).getUTCFullYear();

  // A cell as a date: { day, yearless } or null. Accepts Excel serials, Date objects, m/d/yyyy, m/d/yy, m-d-yy,
  // m.d.yy, yyyy-mm-dd (a time after it is ignored), and m/d with no year, which takes refYear and is marked
  // yearless so the caller can correct and list it. Anything else is null.
  function toDay(v, refYear) {
    if (v == null || v === '') return null;
    if (typeof v === 'number') return v >= 20000 && v <= 80000 ? { day: Math.floor(v) - 25569, yearless: false } : null;
    if (v instanceof Date) {
      if (isNaN(v)) return null;
      const d = ymd(v.getFullYear(), v.getMonth() + 1, v.getDate());
      return d == null ? null : { day: d, yearless: false };
    }
    const s = S(v);
    let m = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})(?:[T\s].*)?$/);
    if (m) { const d = ymd(+m[1], +m[2], +m[3]); return d == null ? null : { day: d, yearless: false }; }
    m = s.match(/^(\d{1,2})[/.-](\d{1,2})[/.-](\d{2}|\d{4})$/);
    if (m) {
      let y = +m[3]; if (y < 100) y += 2000;
      const d = ymd(y, +m[1], +m[2]);
      return d == null ? null : { day: d, yearless: false };
    }
    m = s.match(/^(\d{1,2})[/.](\d{1,2})$/);
    if (m && refYear) { const d = ymd(refYear, +m[1], +m[2]); return d == null ? null : { day: d, yearless: true }; }
    return null;
  }
  const dayOf = (v, refYear) => { const r = toDay(v, refYear); return r ? r.day : null; };

  const fmtDay = d => { if (d == null) return ''; const o = new Date(d * DAY); return (o.getUTCMonth() + 1) + '/' + o.getUTCDate() + '/' + o.getUTCFullYear(); };
  const isoDay = d => (d == null ? '' : new Date(d * DAY).toISOString().slice(0, 10));
  const dayFromIso = s => { const m = S(s).match(/^(\d{4})-(\d{2})-(\d{2})/); return m ? ymd(+m[1], +m[2], +m[3]) : null; };

  // Today's date in New York, as a day number. nowMs is passed in so tests can fix the date.
  function todayNY(nowMs) {
    const p = new Intl.DateTimeFormat('en-US', { timeZone: 'America/New_York', year: 'numeric', month: 'numeric', day: 'numeric' })
      .formatToParts(new Date(nowMs)).reduce((o, x) => (o[x.type] = x.value, o), {});
    return ymd(+p.year, +p.month, +p.day);
  }

  // Working days (Monday to Friday) after day a, up to and including day b. 0 when b <= a.
  function businessDaysBetween(a, b) {
    let n = 0;
    for (let d = a + 1; d <= b; d++) { const w = new Date(d * DAY).getUTCDay(); if (w !== 0 && w !== 6) n++; }
    return n;
  }

  // A sheet by name, ignoring case and spaces at the ends, or null.
  function findSheet(wb, name) {
    const want = S(name).toLowerCase();
    const hit = (wb.SheetNames || []).find(n => S(n).toLowerCase() === want);
    return hit == null ? null : { name: hit, ws: wb.Sheets[hit] };
  }

  // Every row of a sheet as an array of raw cell values, always counted from column A and row 1 (SheetJS would
  // otherwise start at the sheet's recorded range, shifting every column when it starts at B).
  function sheetRows(XLSX, ws) {
    if (!ws || !ws['!ref']) return [];
    const range = XLSX.utils.decode_range(ws['!ref']);
    range.s.r = 0; range.s.c = 0;
    return XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: null, blankrows: true, range });
  }

  // Excel column letter for a 0-based index (for messages that name a cell).
  function colLetter(i) { let s = ''; i += 1; while (i > 0) { const r = (i - 1) % 26; s = String.fromCharCode(65 + r) + s; i = Math.floor((i - 1) / 26); } return s; }

  return { DAY, S, num, normStyle, normPo, normUpc, headerKey, findColumns, ymd, dayYear, toDay, dayOf, fmtDay, isoDay,
    dayFromIso, todayNY, businessDaysBetween, findSheet, sheetRows, colLetter };
});
