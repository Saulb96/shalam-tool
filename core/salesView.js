// The sales page's filters (sales.html, Saul, Nov 4): one of the two sales sheets as InvSalesSheet.readWorkbook read it
// back, cut down to what a salesperson picked. Brand, category and sub category (each one value, or every one), a
// search (every word in the style, UPC or description), on hand or in transit (On hand ATS only), and available by a
// date: on hand counts at once, in transit from its date (ETA + 7) on or before it, Pull from China from its
// pull-from date on or before it. Each line keeps its own pack and note; a line with nothing left is left off. The
// download is made from what this gives (InvSalesSheet.buildWorkbook), so it holds exactly what the page shows.
// Pure: no network, no DOM. Loads in the page (window.InvSalesView, after core/common.js) and in Node.
(function (root, factory) {
  const node = typeof module === 'object' && module.exports;
  const api = factory(node ? require('./common.js') : root.InvCommon);
  if (node) module.exports = api;
  root.InvSalesView = api;
})(typeof globalThis !== 'undefined' ? globalThis : this, function (C) {
  'use strict';
  if (!C) throw new Error('core/common.js must load before core/salesView.js');
  const { S } = C;
  const isDay = v => typeof v === 'number' && Number.isFinite(v);
  const units = v => (isDay(v) && v > 0 ? v : 0);
  const ALL = '*';
  const WHERE = { all: 'all', onHand: 'onHand', transit: 'transit' };
  const byName = (a, b) => a.localeCompare(b, 'en', { sensitivity: 'base', numeric: true }) || (a < b ? -1 : a > b ? 1 : 0);
  const lineKey = r => `${S(r.style)}|${S(r.pack)}|${S(r.note)}`;

  // The choices each filter offers: every brand and category on the sheet A to Z ('' for none, last), and each
  // category's sub categories (ALL: every one's).
  function choices(sheet) {
    const rows = (sheet && sheet.rows) || [];
    const list = vals => { const u = [...new Set(vals.map(S))]; return [...u.filter(Boolean).sort(byName), ...(u.includes('') ? [''] : [])]; };
    const subs = { [ALL]: list(rows.map(r => r.sub)) };
    for (const c of list(rows.map(r => r.category))) subs[c] = list(rows.filter(r => S(r.category) === c).map(r => r.sub));
    return { brands: list(rows.map(r => r.brand)), categories: list(rows.map(r => r.category)), subs };
  }

  // f: { q, brand, category, sub (ALL or a value, '' for none), where (WHERE, On hand ATS only), by (a day number, or
  // null for any date) }. Answers { rows, byDate, photos, at (each row's place in sheet.rows), totals, counts: { rows,
  // styles } }: totals { onHand, inTransit, total } for On hand ATS, { units } for Pull from China.
  function view(sheet, f) {
    f = f || {};
    const kind = sheet && sheet.kind === 'china' ? 'china' : 'sales';
    const words = S(f.q).toLowerCase().split(/\s+/).filter(Boolean);
    const pick = (v, want) => want == null || want === ALL || S(v) === S(want);
    const by = isDay(f.by) ? f.by : null, where = kind === 'sales' && WHERE[f.where] ? f.where : WHERE.all;
    const inDate = w => by == null || w === 'now' || (isDay(w) && w <= by);
    // On hand ATS's schedule, by line.
    const sched = new Map();
    for (const b of (sheet && sheet.byDate) || []) (sched.get(lineKey(b)) || sched.set(lineKey(b), []).get(lineKey(b))).push(b);
    const rows = [], byDate = [], photos = [], at = [], styles = new Set();
    const totals = kind === 'sales' ? { onHand: 0, inTransit: 0, total: 0 } : { units: 0 };
    ((sheet && sheet.rows) || []).forEach((r, i) => {
      if (!pick(r.brand, f.brand) || !pick(r.category, f.category) || !pick(r.sub, f.sub)) return;
      if (words.length) {
        const hay = [r.style, r.upc, r.desc].map(S).join(' ').toLowerCase();
        if (!words.every(w => hay.includes(w))) return;
      }
      let out;
      if (kind === 'china') {
        if (by != null && !(isDay(r.pullFrom) && r.pullFrom <= by)) return;
        out = { ...r };
        totals.units += units(r.units);
      } else {
        const onHand = where === WHERE.transit ? 0 : units(r.onHand);
        const list = sched.get(lineKey(r));
        // The line's in-transit dates: its schedule (each date's units), else the row's own (one date for all of it).
        const ships = list ? list.filter(b => b.where === 'In transit').map(b => ({ when: b.when, units: units(b.units) }))
          : units(r.inTransit) ? [{ when: r.transitFrom ?? null, units: units(r.inTransit) }] : [];
        const kept = where === WHERE.onHand ? [] : ships.filter(x => x.units > 0 && inDate(x.when));
        const inTransit = kept.reduce((t, x) => t + x.units, 0), total = onHand + inTransit;
        if (!(total > 0)) return;
        out = { ...r, onHand, inTransit, transitFrom: kept.length ? kept[0].when : null, total };
        if (list) for (const b of list) {
          if (b.where === 'On hand' ? onHand > 0 : kept.some(x => x.when === b.when)) byDate.push({ ...b });
        }
        totals.onHand += onHand; totals.inTransit += inTransit; totals.total += total;
      }
      rows.push(out); at.push(i); photos.push(((sheet && sheet.photos) || [])[i] || null); styles.add(S(r.style));
    });
    return { kind, rows, byDate, photos, at, totals, counts: { rows: rows.length, styles: styles.size } };
  }

  // Whether any filter is set (the download is then "filtered").
  const filtered = f => !!f && (!!S(f.q) || [f.brand, f.category, f.sub].some(v => v != null && v !== ALL) || (WHERE[f.where] && f.where !== WHERE.all) || isDay(f.by));

  return { choices, view, filtered, ALL, WHERE };
});
