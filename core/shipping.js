// Forwarder ("Active Shipment" .xlsx) and drayage (.csv) reports, PO text expansion, and the newest-usable-file
// rule. Ported from the Late & Short Function (shipping.py, function_app.py). Pure: no network, no DOM, and "today"
// (a day number) is passed in. Loads in the page (window.InvShipping) and in Node (module.exports).
(function (root, factory) {
  const node = typeof module === 'object' && module.exports;
  const api = factory(node ? require('./common.js') : root.InvCommon);
  if (node) module.exports = api;
  root.InvShipping = api;
})(typeof globalThis !== 'undefined' ? globalThis : this, function (C) {
  'use strict';
  if (!C) throw new Error('core/common.js must be loaded before core/shipping.js');

  const KNOWN_STATUSES = ['Pending Departure', 'Sailed', 'Soon To Arrive', 'Arrived'];
  const KNOWN_AVAILABILITY = ['DELIVERED', 'AVAILABLE', 'NOT AVAILABLE'];
  const MAX_REPORT_DAYS = 7;       // older files are not used
  const TRY_FILES = 3;             // newest files tried before giving up
  const AMBER_BUSINESS_DAYS = 2;   // a Friday report is not "old" on Monday
  const REUSE_DAYS = 14;           // one container on two shipments this far apart = the number was reused

  const CONT_RE = /[A-Z]{4}\d{7}/g;
  // Text as shipping.py's cleanstr: non-breaking spaces become spaces, ends trimmed.
  const T = v => C.S(v == null ? '' : String(v).replace(/ /g, ' '));
  const containersIn = text => T(text).toUpperCase().match(CONT_RE) || [];

  // ------------------------------------------------------------ PO text

  const TOKEN = /(\d{6,7})([A-Z]*)((?:-+[A-Z0-9]+)*)/g;
  const SIBLING = /(\d{6})([A-Z]+)((?:\/[A-Z](?![A-Z0-9]))+)/g;   // 900107A/B -> 900107A/900107B

  function letterRange(a, b) {
    if (a.length !== 1 || b.length !== 1 || a > b) return [];
    const out = [];
    for (let c = a.charCodeAt(0); c <= b.charCodeAt(0); c++) out.push(String.fromCharCode(c));
    return out;
  }

  // The import POs written in a free-text PO cell ("900529A-B-C", "900420A---900420K", "900502ABCDEG", truncated
  // text then " - " full text, 9004360 typed for 900436O), as normalized keys, sorted. Keys in `known` (a Set or
  // array of tracker POs) are kept whole, so a real two-letter suffix (900436AB) or "-BCF" is never split.
  function poKeys(text, known) {
    const has = known && typeof known.has === 'function' ? k => known.has(k)
      : Array.isArray(known) ? k => known.includes(k) : () => false;
    let s = T(text).toUpperCase();
    if (!s) return [];
    const parts = s.split(/\s+-\s+/).filter(p => p.trim());
    if (parts.length) s = parts.reduce((a, b) => (b.length > a.length ? b : a));
    s = s.replace(SIBLING, (m, d, l, rest) => [d + l].concat(rest.split('/').slice(1).map(x => d + x)).join('/'));
    const out = new Set();
    for (const m of s.matchAll(TOKEN)) {
      let digits = m[1], letters = m[2];
      const tail = m[3];
      if (digits.length === 7) {
        if (!digits.endsWith('0')) continue;
        digits = digits.slice(0, 6); letters = 'O' + letters;   // zero typed for the letter O
      }
      const base = digits + letters;
      if (tail) {
        const full = C.normPo(base + tail);
        if (has(full)) { out.add(full); continue; }
        const r = tail.match(/^-+(\d{6})([A-Z])$/);
        if (r && r[1] === digits && letters) {                     // 900420A---900420K
          const span = letterRange(letters.slice(-1), r[2]);
          (span.length ? span : [letters]).forEach(c => out.add(digits + c));
          continue;
        }
        const singles = Array.from(tail.matchAll(/-([A-Z])(?=-|$)/g), x => x[1]);
        if (singles.length && letters && singles.map(c => '-' + c).join('') === tail) {
          const span = singles.length === 1 && letters.length === 1 ? letterRange(letters, singles[0]) : [];
          (span.length ? span : [letters].concat(singles)).forEach(c => out.add(digits + c));   // 900052A-J, 900529A-B-C
          continue;
        }
        out.add(full);                                             // 900447-BCF, 900508-1
        continue;
      }
      if (has(base) || letters.length <= 1) out.add(base);
      else for (const c of letters) out.add(digits + c);          // 900502ABCDEG -> A, B, C, D, E, G
    }
    return Array.from(out).sort();
  }

  // ------------------------------------------------------------ shared cell readers

  // A date cell as a day number, or null. Text that isn't a date is listed, never guessed. A date over a year from
  // today, or before 2015, is kept but listed: most likely a typo in the year.
  const FLOOR = C.ymd(2015, 1, 1);
  function dateCell(v, refYear, note, what, today) {
    if (v == null) return null;
    const x = typeof v === 'string' ? T(v) : v;
    if (x === '') return null;
    const r = C.toDay(x, refYear);
    if (!r) { note('bad-date', `${what} "${T(v)}" isn't a date`); return null; }
    if (r.yearless) note('yearless-date', `${what} "${T(v)}" has no year; read as ${C.fmtDay(r.day)}`);
    if (r.day < FLOOR || (today != null && r.day > today + 366)) note('date-far', `${what} reads as ${C.fmtDay(r.day)}, more than a year from today or before 2015; check the year`);
    return r.day;
  }

  // One line of a multi-line cell (port of _per_container): the raw cell is split before trimming, because a blank
  // line is a container with no value yet. A single-line cell is the value for every container.
  function lineOf(v, i) {
    if (v == null || typeof v === 'number' || v instanceof Date) return v;
    const raw = String(v).replace(/ /g, ' ').replace(/\r/g, '');
    if (!raw.includes('\n')) return raw.trim();
    const lines = raw.split('\n');
    return i < lines.length ? lines[i].trim() : '';
  }

  // Yes when every line says Yes (port of _yes). Blank is null (not known), not "no".
  function yesCell(v, note, what) {
    const vals = T(v).split('\n').map(x => x.trim().toLowerCase()).filter(Boolean);
    if (!vals.length) return null;
    const odd = vals.filter(x => x !== 'yes' && x !== 'no');
    if (odd.length) note('bad-yes-no', `${what} reads "${odd.join(', ')}", not Yes or No; read as not released`);
    return vals.every(x => x === 'yes');
  }
  const and3 = (a, b) => (a === false || b === false ? false : a == null || b == null ? null : true);

  const multiText = v => T(v == null ? '' : String(v).replace(/\r/g, ''));

  // ------------------------------------------------------------ forwarder

  const FWD_SPEC = {
    id: { names: ['ID'], required: true },
    status: { names: ['Status'], required: true },
    etd: { names: ['ETD'], required: true },
    containers: { names: ['Container Numbers'], required: true },
    po: { names: ['PO Numbers'], required: true },
    finalEta: { names: ['Final ETA'] },
    eta: { names: ['ETA'] },
    mode: { names: ['Mode'] },
    supplier: { names: ['Supplier Name'] },
    customs: { names: ['Customs Released'] },
    freight: { names: ['Freight Released'] },
    cargoReady: { names: ['Cargo Ready Date'] },
    cartons: { names: ['Cartons'] },
    totalCartons: { names: ['Total Cartons'] },
    lfd: { names: ['LFD'] },
    discharged: { names: ['Discharged'] },
    pickedUp: { names: ['Picked Up Destination'] },
  };

  // The forwarder's Active Shipment report: one record per shipment row, plus byContainer (one entry per container,
  // as shipping.py keys it: a container number reused more than 14 days apart keeps the later trip; one container on
  // two shipments closer than that is read as one; a shipment with no container yet is "#<ID>").
  function parseForwarder(XLSX, wb, today) {
    const issues = [];
    const counts = { rowsRead: 0, records: 0, containers: 0, blank: 0, noId: 0, noContainer: 0 };
    const out = { ok: false, error: null, headerRow: null, headers: [], records: [], byContainer: new Map(), issues,
      counts, unknownStatuses: [] };
    const fail = (error, extra) => Object.assign(out, { ok: false, error }, extra);

    const sheetName = wb && wb.SheetNames && wb.SheetNames[0];
    if (!sheetName) return fail('The forwarder report has no sheets.');
    const rows = C.sheetRows(XLSX, wb.Sheets[sheetName]);
    if (!rows.length) return fail('The forwarder report is empty.');

    const hi = rows.slice(0, 10).findIndex(r => (r || []).some(c => C.headerKey(c) === 'containernumbers'));
    if (hi < 0) {
      const seen = rows.slice(0, 10).map(r => (r || []).map(T).filter(Boolean)).reduce((a, b) => (b.length > a.length ? b : a), []);
      return fail(`The forwarder report has no header row with a "Container Numbers" column in its first 10 rows. ` +
        `Headers seen: ${seen.join(', ') || 'none'}.`, { missing: ['Container Numbers'], seen });
    }
    const f = C.findColumns(rows[hi], FWD_SPEC);
    out.headerRow = hi + 1;
    out.headers = f.seen;
    const missing = f.missing.slice();
    if (f.cols.finalEta == null && f.cols.eta == null) missing.push('Final ETA or ETA');
    if (missing.length) {
      return fail(`The forwarder report is missing the column${missing.length > 1 ? 's' : ''} ${missing.map(m => `"${m}"`).join(', ')}. ` +
        `Headers seen on row ${hi + 1}: ${f.seen.join(', ')}.`, { missing, seen: f.seen });
    }
    const col = f.cols;
    const refYear = today != null ? C.dayYear(today) : undefined;
    const get = (r, k) => (col[k] == null ? null : r[col[k]]);

    const unknown = new Set();
    const seenIds = new Map();
    const entriesOf = new Map();   // container -> the perContainer entries that byContainer's entry came from

    for (let i = hi + 1; i < rows.length; i++) {
      const r = rows[i] || [];
      const row = i + 1;
      counts.rowsRead++;
      const note = (code, message) => issues.push({ source: 'Forwarder', code, row, message });
      if (r.every(v => T(v) === '')) { counts.blank++; continue; }
      const id = T(get(r, 'id'));
      if (!id) { counts.noId++; note('no-id', 'Row has data but no shipment ID; not read'); continue; }
      const who = `Shipment ${id}`;
      if (seenIds.has(id)) note('duplicate-id', `${who} is also on row ${seenIds.get(id)}`);
      seenIds.set(id, row);

      const status = T(get(r, 'status'));
      if (!status) note('no-status', `${who} has no status`);
      else if (!KNOWN_STATUSES.includes(status)) {
        unknown.add(status);
        note('unknown-status', `${who} has the status "${status}", which this tool doesn't know`);
      }
      const etd = dateCell(get(r, 'etd'), refYear, note, `${who} ETD`, today);
      const eta = dateCell(get(r, 'eta'), refYear, note, `${who} ETA`, today);
      const finalEta = dateCell(get(r, 'finalEta'), refYear, note, `${who} Final ETA`, today);
      const cargoReady = dateCell(get(r, 'cargoReady'), refYear, note, `${who} Cargo Ready Date`, today);
      const tripEta = finalEta != null ? finalEta : eta;

      const cartons = [];
      const cv = get(r, 'cartons');
      if (typeof cv === 'number') cartons.push(cv);
      else if (multiText(cv)) {
        for (const line of String(cv).replace(/\r/g, '').split('\n')) {
          const n = C.num(line);
          if (Number.isNaN(n)) note('bad-number', `${who} Cartons "${T(line)}" isn't a number`);
          cartons.push(Number.isNaN(n) ? null : n);
        }
      }
      let totalCartons = C.num(get(r, 'totalCartons'));
      if (Number.isNaN(totalCartons)) { note('bad-number', `${who} Total Cartons "${T(get(r, 'totalCartons'))}" isn't a number`); totalCartons = null; }

      // Containers, each with the line of the date columns that belongs to it.
      const rawC = get(r, 'containers');
      let conts = [];
      if (rawC != null) {
        const lines = String(rawC).replace(/ /g, ' ').toUpperCase().split('\n');
        const perLine = lines.map(x => x.match(CONT_RE) || []);
        if (perLine.every(x => x.length <= 1)) conts = perLine.flatMap((x, j) => (x.length ? [[x[0], j]] : []));
        else conts = containersIn(rawC).map((c, j) => [c, j]);
        lines.forEach((x, j) => {
          if (!perLine[j].length && T(x)) note('bad-container', `${who} Container Numbers line ${j + 1} reads "${T(x)}", which isn't a container number`);
        });
      }

      const rec = {
        row, id, status, mode: T(get(r, 'mode')), etd, eta, finalEta, cargoReady,
        containers: [], pseudo: false, perContainer: [],
        poText: multiText(get(r, 'po')),
        customsReleased: yesCell(get(r, 'customs'), note, `${who} Customs Released`),
        freightReleased: yesCell(get(r, 'freight'), note, `${who} Freight Released`),
        cartons, totalCartons, supplier: multiText(get(r, 'supplier')),
      };
      out.records.push(rec);
      const entry = (container, j) => ({
        container,
        lfd: dateCell(lineOf(get(r, 'lfd'), j), refYear, note, `${who} ${container} LFD`, today),
        discharged: dateCell(lineOf(get(r, 'discharged'), j), refYear, note, `${who} ${container} Discharged`, today),
        pickedUp: dateCell(lineOf(get(r, 'pickedUp'), j), refYear, note, `${who} ${container} Picked Up Destination`, today),
        olderTrip: false,
      });
      const view = e => ({
        container: e.container, ids: [id], rows: [row], status, mode: rec.mode, etd, eta: tripEta,
        lfd: e.lfd, discharged: e.discharged, pickedUp: e.pickedUp, poTexts: [rec.poText],
        customsReleased: rec.customsReleased, freightReleased: rec.freightReleased, pseudo: rec.pseudo,
      });

      if (!conts.length) {
        // Not loaded yet, or LCL / air: still counts for its POs' ETA and status.
        counts.noContainer++;
        if (status !== 'Pending Departure' && status !== '') note('no-container', `${who} (${status}) has no container number`);
        const key = '#' + id;
        rec.pseudo = true;
        rec.containers.push(key);
        const e = entry(key, 0);
        rec.perContainer.push(e);
        out.byContainer.set(key, view(e));
        continue;
      }
      for (const [c, j] of conts) {
        const e = entry(c, j);
        rec.containers.push(c);
        rec.perContainer.push(e);
        const v = view(e);
        const prev = out.byContainer.get(c);
        if (prev && prev.eta != null && v.eta != null && Math.abs(prev.eta - v.eta) > REUSE_DAYS) {
          const [older, newer] = prev.eta > v.eta ? [id, prev.ids.join(', ')] : [prev.ids.join(', '), id];
          note('container-reused', `Container ${c} is on shipments ${older} and ${newer}, more than ${REUSE_DAYS} days apart; ` +
            `read as a reused number, the later trip (${newer}) is kept`);
          if (prev.eta > v.eta) { e.olderTrip = true; continue; }
          entriesOf.get(c).forEach(x => { x.olderTrip = true; });
          entriesOf.set(c, []);
        } else if (prev) {
          // One container on two shipments (consolidated house bills): read as one.
          if (prev.ids.includes(id)) note('container-listed-twice', `${who} lists container ${c} twice`);
          else note('container-on-two-shipments', `Container ${c} is on shipments ${prev.ids.join(', ')} and ${id}; read as one container`);
          v.ids = prev.ids.concat(v.ids.filter(x => !prev.ids.includes(x)));
          v.rows = prev.rows.concat(v.rows);
          v.poTexts = prev.poTexts.concat(v.poTexts);
          for (const k of ['lfd', 'discharged', 'pickedUp']) if (v[k] == null) v[k] = prev[k];
          v.customsReleased = and3(v.customsReleased, prev.customsReleased);
          v.freightReleased = and3(v.freightReleased, prev.freightReleased);
        }
        if (!entriesOf.has(c)) entriesOf.set(c, []);
        entriesOf.get(c).push(e);
        out.byContainer.set(c, v);
      }
    }

    counts.records = out.records.length;
    counts.containers = Array.from(out.byContainer.keys()).filter(k => !k.startsWith('#')).length;
    out.unknownStatuses = Array.from(unknown).sort();
    if (!counts.records) return fail(`The forwarder report has no shipments (${counts.rowsRead} rows read below the header on row ${hi + 1}).`);
    if (!counts.containers) return fail(`The forwarder report has ${counts.records} shipments but no container numbers.`);
    out.ok = true;
    return out;
  }

  // ------------------------------------------------------------ CSV

  // RFC 4180: quoted fields with doubled quotes, commas and line breaks inside quotes, CRLF or LF; a leading BOM is
  // dropped. A quote that doesn't open a field is kept as written (="0123" stays ="0123"). Every line is a row, a
  // blank line included ([""]), so row numbers match the file opened in Excel. A file ending inside quotes throws.
  function parseCsv(text) {
    let s = String(text == null ? '' : text);
    if (s.charCodeAt(0) === 0xfeff) s = s.slice(1);
    const rows = [];
    let row = [], field = '', quoted = false, fresh = true;
    for (let i = 0; i < s.length; i++) {
      const ch = s[i];
      if (quoted) {
        if (ch === '"') {
          if (s[i + 1] === '"') { field += '"'; i++; } else quoted = false;
        } else field += ch;
        continue;
      }
      if (ch === '"' && fresh) { quoted = true; fresh = false; continue; }
      if (ch === ',') { row.push(field); field = ''; fresh = true; continue; }
      if (ch === '\r' || ch === '\n') {
        if (ch === '\r' && s[i + 1] === '\n') i++;
        row.push(field); rows.push(row);
        row = []; field = ''; fresh = true;
        continue;
      }
      field += ch; fresh = false;
    }
    if (quoted) throw new Error('The file ends inside a quoted value (a quote mark that is never closed).');
    if (!fresh || row.length || field) { row.push(field); rows.push(row); }
    return rows;
  }

  // ------------------------------------------------------------ drayage

  const DRAY_SPEC = {
    container: { names: ['Container#', 'Container Number'], required: true },
    eta: { names: ['ETA'], required: true },
    availability: { names: ['Availability'], required: true },
    actualPickup: { names: ['Actual Pick Up Date'], required: true },
    actualDelivery: { names: ['Actual Delivery Date'], required: true },
    ref: { names: ['Ref#'] },
    bl: { names: ['BL Number'] },
    lfd: { names: ['LFD'] },
    terminal: { names: ['Terminal'] },
    scheduledPickup: { names: ['Scheduled Pick Up Date'] },
    proposedDelivery: { names: ['Proposed Deliver Date', 'Proposed Delivery Date'] },
    emptyReturn: { names: ['Empty Return Date'] },
  };

  // Excel's way of keeping text as text: ="0123" -> 0123.
  const unwrap = v => { const s = T(v); const m = s.match(/^="([\s\S]*)"$/); return m ? m[1] : s; };

  // The drayage event report: one record per container. A container on several rows keeps the latest ETA (numbers
  // get reused); on equal ETAs the later row wins, as in shipping.py.
  function parseDrayage(text, today) {
    const issues = [];
    const counts = { rowsRead: 0, records: 0, blank: 0, noContainer: 0, duplicates: 0 };
    const out = { ok: false, error: null, headerRow: null, headers: [], records: [], issues, counts, unknownAvailability: [] };
    const fail = (error, extra) => Object.assign(out, { ok: false, error }, extra);

    let rows;
    try { rows = parseCsv(text); } catch (e) { return fail('The drayage report can\'t be read: ' + e.message); }
    const hi = rows.findIndex(r => r.some(v => T(v) !== ''));
    if (hi < 0) return fail('The drayage report is empty.');
    const f = C.findColumns(rows[hi], DRAY_SPEC);
    out.headerRow = hi + 1;
    out.headers = f.seen;
    if (f.missing.length) {
      return fail(`The drayage report is missing the column${f.missing.length > 1 ? 's' : ''} ${f.missing.map(m => `"${m}"`).join(', ')}. ` +
        `Headers seen on row ${hi + 1}: ${f.seen.join(', ')}.`, { missing: f.missing, seen: f.seen });
    }
    const col = f.cols;
    const refYear = today != null ? C.dayYear(today) : undefined;
    const unknown = new Set();
    const byContainer = new Map();

    for (let i = hi + 1; i < rows.length; i++) {
      const cells = rows[i].map(unwrap);
      const row = i + 1;
      counts.rowsRead++;
      const note = (code, message) => issues.push({ source: 'Drayage', code, row, message });
      if (cells.every(v => v === '')) { counts.blank++; continue; }
      if (cells.length !== rows[hi].length) {
        note('field-count', `Row has ${cells.length} values where the header has ${rows[hi].length}; its columns may be shifted`);
      }
      const get = k => (col[k] == null ? '' : cells[col[k]] == null ? '' : cells[col[k]]);
      const found = containersIn(get('container'));
      if (!found.length) {
        counts.noContainer++;
        note('no-container', get('container') ? `Container# "${get('container')}" isn't a container number; row not read` : 'Row has no container number; not read');
        continue;
      }
      const c = found[0];
      if (found.length > 1) note('several-containers', `Container# lists ${found.join(', ')}; only ${c} is read`);
      const d = (k, what) => dateCell(get(k), refYear, note, `${c} ${what}`, today);
      const availability = get('availability').toUpperCase();
      if (!availability) note('no-availability', `${c} has no availability`);
      else if (!KNOWN_AVAILABILITY.includes(availability)) {
        unknown.add(availability);
        note('unknown-availability', `${c} has the availability "${availability}", which this tool doesn't know`);
      }
      const rec = {
        row, container: c, ref: get('ref'), eta: d('eta', 'ETA'), lfd: d('lfd', 'LFD'), availability,
        terminal: get('terminal'), scheduledPickup: d('scheduledPickup', 'Scheduled Pick Up Date'),
        actualPickup: d('actualPickup', 'Actual Pick Up Date'), proposedDelivery: d('proposedDelivery', 'Proposed Deliver Date'),
        actualDelivery: d('actualDelivery', 'Actual Delivery Date'), emptyReturn: d('emptyReturn', 'Empty Return Date'),
        bl: get('bl'),
      };
      const prev = byContainer.get(c);
      if (prev) {
        counts.duplicates++;
        const keep = (rec.eta == null ? -Infinity : rec.eta) >= (prev.eta == null ? -Infinity : prev.eta) ? rec : prev;
        note('duplicate-container', `Container ${c} is also on row ${prev.row}; row ${keep.row} is kept (latest ETA)`);
        if (keep === prev) continue;
      }
      byContainer.set(c, rec);
    }

    out.records = Array.from(byContainer.values());
    counts.records = out.records.length;
    out.unknownAvailability = Array.from(unknown).sort();
    if (!counts.records) return fail(`The drayage report has no container rows (${counts.rowsRead} rows read below the header).`);
    out.ok = true;
    return out;
  }

  // ------------------------------------------------------------ newest usable file

  // Report files in a folder (Graph driveItems: name, lastModifiedDateTime), newest first, skipping Office lock files
  // (~$...), folders and other extensions. The caller tries up to TRY_FILES of them in order: it stops at the first
  // one that is tooOld (everything after it is older), skips one that doesn't parse or has no records, and shows
  // amber when the one it uses is AMBER_BUSINESS_DAYS or more business days old.
  function pickNewest(files, ext, today) {
    const want = String(ext || '').toLowerCase();
    return (files || [])
      .filter(x => x && !x.folder && typeof x.name === 'string' && !x.name.startsWith('~$') && x.name.toLowerCase().endsWith(want))
      .map(x => {
        const ms = Date.parse(x.lastModifiedDateTime);
        const day = Number.isFinite(ms) ? C.todayNY(ms) : null;
        const ageDays = day == null || today == null ? null : today - day;
        const businessDaysOld = day == null || today == null ? null : C.businessDaysBetween(day, today);
        return { id: x.id, name: x.name, lastModifiedDateTime: x.lastModifiedDateTime, item: x, ms, day, ageDays, businessDaysOld,
          tooOld: ageDays == null || ageDays > MAX_REPORT_DAYS, amber: businessDaysOld != null && businessDaysOld >= AMBER_BUSINESS_DAYS };
      })
      .sort((a, b) => (Number.isFinite(b.ms) ? b.ms : -Infinity) - (Number.isFinite(a.ms) ? a.ms : -Infinity))
      .map(({ ms, ...rest }) => rest);
  }

  return { KNOWN_STATUSES, KNOWN_AVAILABILITY, MAX_REPORT_DAYS, TRY_FILES, AMBER_BUSINESS_DAYS, REUSE_DAYS,
    containersIn, poKeys, parseForwarder, parseCsv, parseDrayage, pickNewest };
});
