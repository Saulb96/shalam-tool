import test from 'node:test';
import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import { createHash } from 'node:crypto';
import { core, root, loadXLSX, loadExcelJS, xlsxParts } from './load.mjs';

const C = core('common');
const SELL = core('sell');
const SS = core('salesSheet');
const TODAY = C.ymd(2026, 11, 2);
const D = (m, d, y = 2026) => C.ymd(y, m, d);
const serial = d => d + 25569;   // a day number as Excel's date serial

// ---------- invented data (POs 9xxxxx, styles SS…, Factory A)

// An allocation as core/stages.js gives it (only the fields the selling view reads).
let nextId = 1;
const NO_DATES = { ctReady: null, ctRow: null, poBook: null, factoryCrd: null, crdExpected: null, shipBy: null, orderDate: null, original: null };
function alloc(o = {}) {
  const units = o.units ?? 1000, master = o.master === undefined ? 24 : o.master;
  return { id: 'a' + nextId++, bucket: 'inProd', units, cases: master ? Math.ceil(units / master) : null, master, style: 'SS100', styleRaw: 'SS100',
    upc: '000111222333', po: '900101', poShown: '', ptype: 'LA', customer: '', factory: 'Factory A', crd: null, od: null,
    etd: null, eta: null, stage: 'In production', flags: [], sheetRow: 10, round: '', containers: [], section: '',
    ...o, dates: { ...NO_DATES, ...(o.dates || {}) } };
}
function stagesOf(allocs, extra = {}) {
  const styles = new Map();
  for (const a of allocs) {
    if (!styles.has(a.style)) styles.set(a.style, { style: a.style, styleRaw: a.styleRaw, desc: 'Item ' + a.styleRaw, category: extra.category || 'Lip',
      factory: a.factory, factories: [a.factory], upcs: [], allocs: [] });
    const s = styles.get(a.style);
    if (a.upc && !s.upcs.some(u => u.upc === a.upc)) s.upcs.push({ upc: a.upc });
    s.allocs.push(a);
  }
  return { ok: true, allocs, styles: [...styles.values()] };
}
// A parseAts result: { style: [inWh, orders, erpAts (its "ATS from in WH"; null: none), shipments ([units, eta] each)] }.
// Its Import PO Backlog keeps our goods in China (none of its In Trans comes off them).
function atsOf(map, reportDay = D(10, 31)) {
  const styles = new Map(Object.entries(map).map(([style, v]) => {
    const ships = v[3] || [];
    return [style, { style, styleRaw: style, desc: 'ERP ' + style, category: 'Lip', inWh: v[0], orders: v[1], erpAts: v[2] ?? null, warehouses: ['WH 1'],
      rows: [{ row: 3, warehouse: 'WH 1', inWh: v[0], orders: v[1], erpAts: v[2] ?? null, inTrans: ships.length ? ships.reduce((t, s) => t + s[0], 0) : null,
        inTransList: ships.map(s => s[0]), inTransBad: [], etaParts: ships.map(s => s[1]), eta: ships.map(s => s[1]).filter(d => d != null),
        backlog: 100000, backlogList: [100000], backlogBad: [], poParts: [] }],
      unknown: v[0] == null || v[1] == null }];
  }));
  return { ok: true, reportDay, reportDayFrom: 'file name', styles, columns: ['style', 'inWh', 'orders', 'erpAts', 'inTrans', 'eta', 'backlog'], issues: [], counts: {} };
}
const sellOf = (allocs, ats) => SELL.buildSell({ stages: stagesOf(allocs), ats, today: TODAY });

// Today 11/2/2026, the ATS report of 10/31.
// SS050: not on the ATS report; one lot in China dated from its order date (an estimate).
// SS100: ATS from in WH 460; In Trans 300 (ETA 11/20: ours on the water, the same shipment) and 120 (ETA 10/20: due
//   10/27, so now); in China a Booked lot whose pull-from date has passed (10/10), 480 on 12/1 and 100 about 12/15.
// SS200: oversold 50: its 30 In Trans and 20 of its 70 in China go to the orders; no case pack.
// SS300: ATS from in WH 10; lots in two packs; 60 of ours on the water the ERP doesn't have (never on the sheet).
// SS400: nothing available (left off). SS500: its "ATS from in WH" (80) isn't In WH − orders (90): the report's wins;
//   an In Trans shipment with no ETA.
function mainSell() {
  const as = [
    alloc({ style: 'SS050', styleRaw: 'SS050', upc: '000000000500', units: 96, inner: 6, dates: { orderDate: D(10, 1) } }),
    alloc({ po: '900101', bucket: 'booked', units: 240, inner: 6, etd: D(10, 20), stage: 'Booked' }),
    alloc({ po: '900102', units: 480, inner: 6, dates: { factoryCrd: D(12, 6) } }),
    alloc({ po: '900104', units: 100, inner: 6, dates: { orderDate: D(10, 1) } }),
    alloc({ po: '900105', bucket: 'water', units: 300, inner: 6, eta: D(11, 20), stage: 'On the water' }),
    alloc({ style: 'SS200', styleRaw: 'SS200', upc: '000000000200', master: null, units: 70, dates: { factoryCrd: D(12, 6) } }),
    alloc({ style: 'SS300', styleRaw: 'SS300', upc: '000000000300', master: 24, inner: 6, units: 48, dates: { factoryCrd: D(12, 6) } }),
    alloc({ style: 'SS300', styleRaw: 'SS300', upc: '000000000301', master: 12, inner: 3, units: 36, po: '900106', dates: { factoryCrd: D(12, 6) } }),
    alloc({ style: 'SS300', styleRaw: 'SS300', upc: '000000000300', master: 12, inner: 3, units: 30, po: '900107', dates: { factoryCrd: D(12, 25) } }),
    alloc({ style: 'SS300', styleRaw: 'SS300', upc: '000000000300', master: 12, units: 60, po: '900108', bucket: 'water', eta: D(11, 15), stage: 'On the water' }),
  ];
  return sellOf(as, atsOf({ SS100: [500, 40, 460, [[300, D(11, 20)], [120, D(10, 20)]]], SS200: [0, 50, -50, [[30, D(11, 10)]]], SS300: [10, 0, 10],
    SS400: [0, 0, 0], SS500: [100, 10, 80, [[24, null]]] }));
}
const ROW_KEYS = ['style', 'upc', 'desc', 'category', 'sub', 'pack', 'onHand', 'inTransit', 'transitFrom', 'total', 'photoKey'];
const DATE_KEYS = ['style', 'desc', 'category', 'when', 'where', 'units'];
const CHINA_KEYS = ['style', 'upc', 'desc', 'category', 'sub', 'pack', 'units', 'pullFrom', 'estimate', 'photoKey'];

// ---------- refusals

test('no sheet when Available to sell is not worked out, the ATS report is missing or unusable, or the units don\'t add up', () => {
  for (const make of ['salesRows', 'chinaRows']) {
    const nl = SS[make](SELL.buildSell({ stages: { ok: false, notLoaded: ['Container Tracking'] }, ats: null, today: TODAY }), { today: TODAY });
    assert.equal(nl.ok, false);
    assert.equal(nl.reason, "Available to sell couldn't be worked out (Container Tracking didn't load), so there's no sales sheet.");
    assert.deepEqual(nl.rows, []);
    assert.equal(SS[make](null, { today: TODAY }).ok, false);

    const as = [alloc({ dates: { factoryCrd: D(12, 6) } })];
    const none = SS[make](sellOf(as, null), { today: TODAY });
    assert.equal(none.ok, false);
    assert.equal(none.reason, "There's no ATS report, so stock available now isn't known and there's no sales sheet.");
    const bad = SS[make](sellOf(as, { ok: false, error: 'the In WH column is missing' }), { today: TODAY });
    assert.equal(bad.ok, false);
    assert.equal(bad.reason, "The ATS report couldn't be used (the In WH column is missing), so stock available now isn't known and there's no sales sheet.");

    const r = sellOf(as, atsOf({ SS100: [10, 0] }));
    assert.equal(SS[make](r, { today: TODAY }).ok, true);
    r.issues.push({ code: 'balance', message: 'invented' });
    const off = SS[make](r, { today: TODAY });
    assert.equal(off.ok, false);
    assert.match(off.reason, /doesn't add up \(a bug in Shalam Tool\)/);
    assert.equal(off.rows.length, 0);

    assert.throws(() => SS[make](r, { today: '11/2/2026' }), /today must be a day number/);
    assert.equal(SS[make](sellOf(as, atsOf({ SS100: [10, 0] }))).asOf, TODAY, "today defaults to Available to sell's own");
  }
});

// ---------- Available to sell: the ATS report's numbers (Saul, Nov 3)

test('Available to sell: one row per style on the ATS report with something available; only the listed fields', () => {
  const sh = SS.salesRows(mainSell(), { today: TODAY });
  assert.equal(sh.ok, true);
  assert.equal(sh.reason, '');
  assert.equal(sh.kind, 'sales');
  assert.equal(sh.asOf, TODAY);
  assert.equal(sh.reportDay, D(10, 31));
  assert.deepEqual(sh.rows.map(r => r.style), ['SS100', 'SS300', 'SS500']);
  for (const r of sh.rows) {
    assert.deepEqual(Object.keys(r), ROW_KEYS, r.style);
    assert.equal(r.total, r.onHand + r.inTransit, r.style);
  }
  for (const b of sh.byDate) assert.deepEqual(Object.keys(b), DATE_KEYS);
  // SS050 isn't on the report; SS200 (oversold beyond its In Trans) and SS400 have nothing.
  assert.deepEqual(sh.counts, { styles: 3, leftOff: 2, stockUnknown: 0, transitUnknown: 0, notOnReport: 1, byDateRows: sh.byDate.length, noPack: 1 });
  assert.deepEqual(sh.totals, { onHand: 550, inTransit: 444, total: 994 });
});

test("on hand is the ATS report's own ATS from in WH (0 when below 0); our shipping sheets add nothing", () => {
  const sh = SS.salesRows(mainSell(), { today: TODAY }), row = st => sh.rows.find(r => r.style === st);
  assert.deepEqual([row('SS100').onHand, row('SS300').onHand, row('SS500').onHand], [460, 10, 80]);
  // SS500: In WH 100 − orders 10 would be 90; the report says 80, and the report wins.
  // SS300: our 60 on the water aren't on the ERP's In Trans, so they aren't offered.
  assert.equal(row('SS300').inTransit, 0);
  // With no "ATS from in WH" for a style, In WH − orders.
  const fb = SS.salesRows(sellOf([alloc({ dates: { factoryCrd: D(12, 6) } })], atsOf({ SS100: [70, 30] })), { today: TODAY });
  assert.deepEqual([fb.rows[0].onHand, fb.rows[0].total], [40, 40]);
  // Oversold: 0 on hand, the rest of it off the In Trans.
  const ov = SS.salesRows(sellOf([alloc({ dates: { factoryCrd: D(12, 6) } })], atsOf({ SS100: [0, 25, -25, [[100, D(11, 20)]]] })), { today: TODAY });
  assert.deepEqual([ov.rows[0].onHand, ov.rows[0].inTransit], [0, 75]);
  assert.equal(SS.photoKey(' ab 12/3-f_x '), 'AB123-F_X');
  assert.equal(SS.photoKey('AB100-F'), 'AB100-F', 'exact key only: -F is kept');
  assert.equal(row('SS100').upc, '000111222333');
  assert.equal(row('SS300').upc, '000000000300, 000000000301');
  assert.equal(row('SS100').photoKey, 'SS100');
});

test('in transit: every In Trans shipment, the oversold off the earliest first, available ETA + 7 days (now once that day has come)', () => {
  const sh = SS.salesRows(mainSell(), { today: TODAY }), row = st => sh.rows.find(r => r.style === st);
  assert.equal(SS.TRANSIT_DAYS, 7);
  // SS100: 300 due 11/27 and 120 due 10/27 (before today: now). Ours on the water is the same 300, counted once.
  assert.deepEqual([row('SS100').inTransit, row('SS100').transitFrom, row('SS100').total], [420, 'now', 880]);
  // SS500: a shipment with no ETA has no date.
  assert.deepEqual([row('SS500').inTransit, row('SS500').transitFrom], [24, null]);
  assert.equal(row('SS300').transitFrom, null);
  // The oversold comes off the earliest ETA first; a shipment with no ETA last.
  const ats = atsOf({ SS100: [0, 130, -130, [[24, null], [100, D(11, 25)], [60, D(11, 12)]]] });
  const o = SS.salesRows(sellOf([alloc({ dates: { factoryCrd: D(12, 6) } })], ats), { today: TODAY });
  assert.deepEqual(o.byDate.map(b => [b.when, b.where, b.units]), [[D(12, 2), 'In transit', 30], [null, 'In transit', 24]]);
});

test('Master/Inner: every pack of the style once, "–" with no Inner, blank when not known', () => {
  assert.equal(SS.packText([{ master: 48, inner: 12 }]), '48/12');
  assert.equal(SS.packText([{ master: 24, inner: null }, { master: 48, inner: 12 }]), '24/– · 48/12');
  assert.equal(SS.packText([]), '');
  assert.equal(SS.packText(null), '');
  const sh = SS.salesRows(mainSell(), { today: TODAY }), row = st => sh.rows.find(r => r.style === st);
  assert.equal(row('SS100').pack, '24/6');
  assert.equal(row('SS300').pack, '12/– · 12/3 · 24/6');
  assert.equal(row('SS500').pack, '', 'only on the ATS report: no pack known');
  // As the selling engine gives them: our lots' pairs; none, the latest Master Sheet line's.
  const s = mainSell().styles.find(x => x.style === 'SS300');
  assert.deepEqual(s.packs, [{ master: 12, inner: null }, { master: 12, inner: 3 }, { master: 24, inner: 6 }]);
});

test('by date: on hand now, then each in-transit date, units only; in the rows\' order', () => {
  const sh = SS.salesRows(mainSell(), { today: TODAY });
  const of = st => sh.byDate.filter(b => b.style === st).map(b => [b.when, b.where, b.units]);
  assert.deepEqual(of('SS100'), [['now', 'On hand', 460], ['now', 'In transit', 120], [D(11, 27), 'In transit', 300]]);
  assert.deepEqual(of('SS300'), [['now', 'On hand', 10]]);
  assert.deepEqual(of('SS500'), [['now', 'On hand', 80], [null, 'In transit', 24]]);
  assert.ok(sh.byDate.every(b => b.units > 0));
  for (const r of sh.rows) assert.equal(sh.byDate.filter(b => b.style === r.style).reduce((t, b) => t + b.units, 0), r.total, r.style);
  assert.deepEqual([...new Set(sh.byDate.map(b => b.style))], sh.rows.map(r => r.style));
});

// ---------- Pull from China

test('Pull from China: our goods in China not sold, one row per style and pull-from date; a date before today left off', () => {
  const sh = SS.chinaRows(mainSell(), { today: TODAY });
  assert.equal(sh.ok, true);
  assert.equal(sh.kind, 'china');
  for (const r of sh.rows) assert.deepEqual(Object.keys(r), CHINA_KEYS, r.style);
  assert.deepEqual(sh.rows.map(r => [r.style, r.pullFrom, r.estimate, r.units, r.pack]), [
    ['SS050', D(10, 1) + 75, true, 96, '24/6'],
    ['SS100', D(12, 1), false, 480, '24/6'],
    ['SS100', D(10, 1) + 75, true, 100, '24/6'],
    // SS200: oversold 50, 30 of it off its In Trans, 20 off its goods in China.
    ['SS200', D(12, 1), false, 50, ''],
    ['SS300', D(12, 1), false, 84, '12/3 · 24/6'],
    ['SS300', D(12, 20), false, 30, '12/3'],
  ]);
  // SS100's Booked 240 had 10/10 to pull from: too late, left off.
  assert.deepEqual(sh.counts, { rows: 6, styles: 4, ordersOff: 20, past: 240, noDate: 0, stockUnknown: 0 });
  assert.deepEqual(sh.totals, { units: 840 });
  // What a style is oversold beyond its In Trans comes off its earliest goods in China first (a past one too).
  const o = SS.chinaRows(sellOf([alloc({ units: 100, etd: D(10, 20), bucket: 'booked', stage: 'Booked' }), alloc({ units: 200, po: '900109', dates: { factoryCrd: D(12, 6) } })],
    atsOf({ SS100: [0, 150, -150] })), { today: TODAY });
  assert.deepEqual(o.rows.map(r => [r.pullFrom, r.units]), [[D(12, 1), 150]]);
  assert.deepEqual([o.counts.ordersOff, o.counts.past], [150, 0]);
});

// ---------- nothing that must stay inside reaches the sheets

// Distinctive marker strings in every field the salesperson must never see: factory, customer, PO, PO type, section,
// container, round, people's names, file names, links, ERP warehouse, and the numbers for In WH, customer orders, sold
// and oversold. Category and sub category are shown (Saul, Nov 3: a tab per category).
const MARK = { factory: 'Factory QXLEAKFAC', customer: 'Customer QXLEAKCUST', po: '987650-QXLEAKPO', poShown: 'PO QXLEAKSHOWN', ptype: 'QXLEAKPTYPE',
  section: 'Section QXLEAKSECT', round: 'Round QXLEAKROUND', container: 'QXLC1234567', person: 'Pat QXLEAKPERSON', file: 'QXLEAKFILE.xlsx',
  link: 'https://contoso.sharepoint.example/sites/QXLEAKSITE/QXLEAKLINK', warehouse: 'WH QXLEAKWH', erpPo: '987651-QXLEAKERPPO',
  inWh: '77791', orders: '4321', sold: '31415', oversold: '27182', issue: 'QXLEAKISSUE' };
function leakySell() {
  const tag = { factory: MARK.factory, customer: MARK.customer, po: MARK.po, poShown: MARK.poShown, ptype: MARK.ptype, section: MARK.section,
    round: MARK.round, containers: [MARK.container], savedBy: MARK.person, fileName: MARK.file, webUrl: MARK.link };
  const as = [
    alloc({ ...tag, units: 480, dates: { factoryCrd: D(12, 6), orderDate: D(9, 1) } }),
    alloc({ ...tag, bucket: 'water', units: 240, eta: D(11, 20), stage: 'On the water', flags: ['msl-behind-ct'] }),
    alloc({ ...tag, bucket: 'pullFromChina', units: 31415, ptype: 'China Pickup' }),
    alloc({ ...tag, style: 'SS700', styleRaw: 'SS700', units: 120, master: null, dates: { orderDate: D(10, 1) } }),
    alloc({ ...tag, style: 'SS701', styleRaw: 'SS701', units: 48, master: 12, dates: { factoryCrd: D(12, 6) } }),
    alloc({ ...tag, style: 'SS701', styleRaw: 'SS701', units: 48, master: 24, dates: { factoryCrd: D(10, 1) } }),
  ];
  const ats = atsOf({ SS100: [77791, 4321, 73470, [[240, D(11, 20)]]], SS700: [10, 0, 10, [[60, D(11, 28)]]], SS701: [0, 0, 0], SS799: [0, 27182, -27182] });
  for (const [k, e] of ats.styles) {
    e.warehouses = [MARK.warehouse];
    e.rows[0].warehouse = MARK.warehouse;
    if (k === 'SS700') e.rows[0].poParts = [MARK.erpPo];
  }
  Object.assign(ats, { columns: [...ats.columns, 'po'], fileName: MARK.file, savedBy: MARK.person });
  const sell = SELL.buildSell({ stages: stagesOf(as), ats, today: TODAY });
  // Whatever the page or a later engine might hang on the result: none of it may be copied.
  Object.assign(sell, { atsFile: MARK.file, savedBy: MARK.person, webUrl: MARK.link });
  for (const x of [...sell.styles, ...sell.lots]) Object.assign(x, { savedBy: MARK.person, fileName: MARK.file, webUrl: MARK.link, warehouse: MARK.warehouse, note: MARK.issue });
  sell.issues.push({ source: 'Available to sell', code: 'invented', po: MARK.po, style: 'SS100', message: MARK.issue });
  return sell;
}
// Every string and number anywhere in a value, as text.
const allText = v => JSON.stringify(v, (k, x) => (x instanceof Map ? [...x] : x instanceof Set ? [...x] : x));
function noMarks(text, where) {
  for (const [k, m] of Object.entries(MARK)) assert.ok(!text.includes(m), `${where}: the ${k} marker "${m}" is in it`);
  for (const m of ['QXLEAK', 'QXLC', '987650', '987651']) assert.ok(!text.includes(m), `${where}: "${m}" is in it`);
}

test('NO LEAK: the rows and the schedule hold none of the names, numbers, flags or problems that must stay inside', () => {
  const sell = leakySell();
  // The markers are really in what the sheets are made from.
  const src = allText(sell);
  for (const m of [MARK.factory, MARK.customer, MARK.po, MARK.section, MARK.round, MARK.container, MARK.person, MARK.file, MARK.link, MARK.warehouse, MARK.erpPo, MARK.issue]) assert.ok(src.includes(m), m);
  const s1 = sell.styles.find(s => s.style === 'SS100');
  assert.deepEqual([s1.inWh, s1.orders, s1.soldChina], [77791, 4321, 31415]);
  assert.equal(sell.styles.find(s => s.style === 'SS799').oversold, 27182);
  const flags = new Set([...sell.styles, ...sell.lots].flatMap(x => x.flags || []).filter(f => f.includes('-')));
  assert.ok(flags.size >= 3, [...flags].join(', '));

  const sh = SS.salesRows(sell, { today: TODAY }), ch = SS.chinaRows(sell, { today: TODAY });
  assert.equal(sh.ok, true);
  assert.equal(ch.ok, true);
  assert.deepEqual(sh.rows.map(r => r.style), ['SS100', 'SS700']);
  assert.deepEqual(ch.rows.map(r => r.style), ['SS100', 'SS700', 'SS701']);
  for (const [out, where] of [[allText(sh), 'the sales rows'], [allText(ch), 'the Pull from China rows']]) {
    noMarks(out, where);
    for (const f of flags) assert.ok(!out.includes(f), `the flag ${f} is in ${where}`);
    for (const i of sell.issues) assert.ok(!out.includes(i.message), `a problem is in ${where}: ${i.message.slice(0, 80)}`);
  }
  // Only the allow-listed fields, at every level.
  for (const r of sh.rows) assert.deepEqual(Object.keys(r), ROW_KEYS);
  for (const r of ch.rows) assert.deepEqual(Object.keys(r), CHINA_KEYS);
  assert.deepEqual(Object.keys(sh).sort(), ['asOf', 'byDate', 'counts', 'kind', 'ok', 'reason', 'reportDay', 'rows', 'totals']);
  assert.deepEqual(Object.keys(ch).sort(), ['asOf', 'counts', 'kind', 'ok', 'reason', 'reportDay', 'rows', 'totals']);
});

// ---------- the Excel file (ExcelJS, vendored)

// A 1×1 PNG and a 2×1 one (invented pictures).
const PNG1 = Uint8Array.from(Buffer.from('iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==', 'base64'));
const PNG2 = Uint8Array.from(Buffer.from('iVBORw0KGgoAAAANSUhEUgAAAAIAAAABCAYAAAD0In+KAAAAEUlEQVR42mP8z8DAwMDAwAAAD/4CAVv6FX4AAAAASUVORK5CYII=', 'base64'));
const readBack = bytes => loadXLSX().read(bytes, { type: 'array', cellNF: true });
const text = (parts, filter = () => true) => [...parts].filter(([p]) => filter(p)).map(([p, b]) => `--- ${p}\n` + Buffer.from(b).toString('latin1')).join('\n');

test('the vendored ExcelJS is 4.4.0 as published (the hash in its VERSION.txt), with its MIT license, and loads in Node', () => {
  const file = readFileSync(root('vendor/exceljs/exceljs.min.js'));
  const version = readFileSync(root('vendor/exceljs/VERSION.txt'), 'utf8');
  assert.match(version, /^ExcelJS 4\.4\.0 /);
  assert.equal(createHash('sha256').update(file).digest('hex'), /sha256: ([0-9a-f]{64})/.exec(version)[1]);
  assert.equal(file.length, 947702);
  assert.ok(file.subarray(0, 40).toString('latin1').startsWith('/*! ExcelJS 19-10-2023 */'));
  assert.match(readFileSync(root('vendor/exceljs/LICENSE'), 'utf8'), /^The MIT License \(MIT\)\s+Copyright \(c\) 2014-2019 Guyon Roche/);
  const E = loadExcelJS();
  assert.equal(typeof E.Workbook, 'function');
});

test('the workbook: every style, a tab per category, By date; the title, frozen photo and style, a filter, numbers and dates as Excel keeps them', async () => {
  const sh = SS.salesRows(mainSell(), { today: TODAY });
  const photos = new Map([['SS100', { data: PNG1, width: 1, height: 1, extension: 'png' }], ['SS300', { data: PNG2.buffer.slice(0), width: 2, height: 1, extension: 'png' }]]);
  const wb = SS.buildWorkbook(loadExcelJS(), sh, photos);
  const bytes = await SS.workbookBytes(wb);
  assert.ok(bytes instanceof Uint8Array && bytes.length > 5000, `${bytes.length} bytes`);
  const book = readBack(bytes);
  assert.deepEqual(book.SheetNames, ['Available to sell', 'Lip', 'By date']);
  assert.deepEqual(SS.tabsOf(sh), book.SheetNames);
  const ws = book.Sheets['Available to sell'], v = a => (ws[a] ? ws[a].v : undefined);
  assert.equal(v('A1'), 'M1S · Available to sell');
  assert.equal(v('D1'), undefined, 'no "as of" line (Saul, Nov 3)');
  assert.deepEqual('ABCDEFGHIJK'.split('').map(c => v(c + '2')), SS.COLS.map(c => c.h));
  assert.deepEqual(SS.COLS.map(c => c.h), ['Photo', 'Style', 'UPC', 'Description', 'Category', 'Sub category', 'Master/Inner', 'Available on hand (units)',
    'Available in transit (units)', 'In transit available from', 'Total available (units)']);
  assert.deepEqual('BCDEFGHIJK'.split('').map(c => v(c + '3')), ['SS100', '000111222333', 'Item SS100', 'Lip', undefined, '24/6', 460, 420, 'now', 880]);
  assert.deepEqual('BGHIJK'.split('').map(c => v(c + '4')), ['SS300', '12/– · 12/3 · 24/6', 10, undefined, undefined, 10]);
  assert.deepEqual('BDGHIJK'.split('').map(c => v(c + '5')), ['SS500', 'ERP SS500', undefined, 80, 24, 'date to come', 104]);
  assert.equal(v('B6'), undefined, 'three styles, no more');
  assert.equal(ws.C3.t, 's', 'UPCs stay text');
  assert.equal(ws.H3.z, '#,##0');
  assert.equal(ws.G3.t, 's', 'Master/Inner is text');
  // A date cell: a style whose first shipment is due after today.
  const later = SS.salesRows(sellOf([alloc({ dates: { factoryCrd: D(12, 6) } })], atsOf({ SS100: [0, 0, 0, [[48, D(11, 20)]]] })), { today: TODAY });
  const lb = readBack(await SS.workbookBytes(SS.buildWorkbook(loadExcelJS(), later, null))).Sheets['Available to sell'];
  assert.deepEqual([lb.J3.v, lb.J3.z, lb.K3.v], [serial(D(11, 27)), 'm/d/yyyy', 48]);

  const wd = book.Sheets['By date'], w = a => (wd[a] ? wd[a].v : undefined);
  assert.deepEqual(['A1', 'B1', 'C1', 'D1', 'E1', 'F1', 'G1'].map(w), ['Style', 'Description', 'Category', 'Available from', 'Where', 'Units', undefined]);
  const sched = [];
  for (let r = 2; w('A' + r) !== undefined; r++) sched.push(['A', 'C', 'D', 'E', 'F'].map(c => w(c + r)));
  assert.deepEqual(sched, [['SS100', 'Lip', 'now', 'On hand', 460], ['SS100', 'Lip', 'now', 'In transit', 120], ['SS100', 'Lip', serial(D(11, 27)), 'In transit', 300],
    ['SS300', 'Lip', 'now', 'On hand', 10], ['SS500', 'Lip', 'now', 'On hand', 80], ['SS500', 'Lip', 'date to come', 'In transit', 24]]);

  const parts = xlsxParts(bytes);
  const sheet1 = Buffer.from(parts.get('xl/worksheets/sheet1.xml')).toString('utf8'), sheet3 = Buffer.from(parts.get('xl/worksheets/sheet3.xml')).toString('utf8');
  assert.match(sheet1, /<pane xSplit="2" ySplit="2" topLeftCell="C3" activePane="bottomRight" state="frozen"\/>/);
  assert.match(sheet1, /<autoFilter ref="A2:K2"\/>/);
  assert.match(sheet3, /<pane xSplit="1" ySplit="1" topLeftCell="B2" activePane="bottomRight" state="frozen"\/>/);
  assert.match(sheet3, /<autoFilter ref="A1:F1"\/>/);
  // Two photos, in the Photo column on their styles' rows (SS100 on row 3, SS300 on row 4: 0-based 2 and 3), fitted
  // inside the cell; every style row has the photo row height. The Lip tab shows them too, from the same two pictures.
  assert.equal([...parts.keys()].filter(p => /^xl\/media\//.test(p)).length, 2);
  assert.equal(Buffer.from(parts.get('xl/drawings/drawing2.xml')).toString('utf8').match(/<xdr:pic>/g).length, 2);
  const drawing = Buffer.from(parts.get('xl/drawings/drawing1.xml')).toString('utf8');
  const anchors = [...drawing.matchAll(/<xdr:from><xdr:col>(\d+)<\/xdr:col><xdr:colOff>(\d+)<\/xdr:colOff><xdr:row>(\d+)<\/xdr:row><xdr:rowOff>(\d+)<\/xdr:rowOff><\/xdr:from><xdr:ext cx="(\d+)" cy="(\d+)"\/>/g)]
    .map(m => m.slice(1).map(Number));
  assert.deepEqual(anchors.map(a => [a[0], a[2]]), [[0, 2], [0, 3]]);
  const EMU = 9525, cellW = Math.round(15 * 7 + 5), cellH = Math.round(SS.PHOTO_ROW_PT * 96 / 72);
  for (const [, colOff, , rowOff, cx, cy] of anchors) {
    assert.ok(colOff / EMU + cx / EMU <= cellW + 1 && rowOff / EMU + cy / EMU <= cellH + 1, `photo inside its cell: ${[colOff, rowOff, cx, cy].map(n => Math.round(n / EMU))}`);
    assert.ok(colOff > 0 && rowOff > 0, 'centred, not hugging the corner');
  }
  for (let r = 3; r <= 5; r++) assert.match(sheet1, new RegExp(`<row r="${r}"[^>]* ht="${SS.PHOTO_ROW_PT}"`));
});

test('the Pull from China workbook: every row, a tab per category, no By date; units and the pull-from date; a style on two rows stores its picture once', async () => {
  const sh = SS.chinaRows(mainSell(), { today: TODAY });
  const photos = new Map([['SS100', { data: PNG1, width: 1, height: 1, extension: 'png' }], ['SS300', { data: PNG2, width: 2, height: 1, extension: 'png' }]]);
  const bytes = await SS.workbookBytes(SS.buildWorkbook(loadExcelJS(), sh, photos));
  const book = readBack(bytes);
  assert.deepEqual(book.SheetNames, ['Pull from China', 'Lip']);
  assert.deepEqual(SS.tabsOf(sh), book.SheetNames);
  const ws = book.Sheets['Pull from China'], v = a => (ws[a] ? ws[a].v : undefined);
  assert.equal(v('A1'), 'M1S · Pull from China');
  assert.equal(book.Sheets.Lip.A1.v, 'M1S · Pull from China · Lip');
  assert.deepEqual('ABCDEFGHI'.split('').map(c => v(c + '2')), ['Photo', 'Style', 'UPC', 'Description', 'Category', 'Sub category', 'Master/Inner', 'Units', 'Pull from date']);
  assert.deepEqual(SS.CHINA_COLS.map(c => c.h), 'ABCDEFGHI'.split('').map(c => v(c + '2')));
  const about = `about ${C.fmtDay(D(10, 1) + 75)}`;
  const got = [];
  for (let r = 3; v('B' + r) !== undefined; r++) got.push(['B', 'G', 'H', 'I'].map(c => v(c + r)));
  assert.deepEqual(got, [['SS050', '24/6', 96, about], ['SS100', '24/6', 480, serial(D(12, 1))], ['SS100', '24/6', 100, about], ['SS200', undefined, 50, serial(D(12, 1))],
    ['SS300', '12/3 · 24/6', 84, serial(D(12, 1))], ['SS300', '12/3', 30, serial(D(12, 20))]]);
  assert.equal(ws.I4.z, 'm/d/yyyy');
  assert.match(Buffer.from(xlsxParts(bytes).get('xl/worksheets/sheet1.xml')).toString('utf8'), /<autoFilter ref="A2:I2"\/>/);
  // SS100 on rows 4 and 5 and SS300 on rows 7 and 8, on both tabs: two pictures in the file, shown eight times.
  const parts = xlsxParts(bytes);
  assert.equal([...parts.keys()].filter(p => /^xl\/media\//.test(p)).length, 2);
  const d1 = Buffer.from(parts.get('xl/drawings/drawing1.xml')).toString('utf8');
  assert.deepEqual([...d1.matchAll(/<xdr:from><xdr:col>\d+<\/xdr:col><xdr:colOff>\d+<\/xdr:colOff><xdr:row>(\d+)<\/xdr:row>/g)].map(m => +m[1]), [3, 4, 6, 7]);
  assert.equal(Buffer.from(parts.get('xl/drawings/drawing2.xml')).toString('utf8').match(/<xdr:pic>/g).length, 4);
});

test('a tab per category: A to Z, no category last, rows by sub category then style, the same numbers; each picture stored once', async () => {
  const sell = mainSell();
  const cat = { SS100: ['Cosmetics2', 'Lip Balm'], SS300: ['Bath', 'Gloves'], SS500: ['', ''] };
  // The page's categories only (none from the ATS report).
  for (const s of sell.styles) if (cat[s.style]) { [s.category, s.sub] = cat[s.style]; s.atsCategory = s.atsSub = ''; }
  const sh = SS.salesRows(sell, { today: TODAY });
  assert.deepEqual(sh.rows.map(r => [r.style, r.category, r.sub]), [['SS300', 'Bath', 'Gloves'], ['SS100', 'Cosmetics2', 'Lip Balm'], ['SS500', '', '']]);
  // The schedule in the same order, with each style's category.
  assert.deepEqual([...new Set(sh.byDate.map(b => b.style))], ['SS300', 'SS100', 'SS500']);
  for (const b of sh.byDate) assert.equal(b.category, cat[b.style][0], b.style);
  assert.deepEqual(SS.categoryTabs(sh.rows).map(t => [t.name, t.category, t.rows]), [['Bath', 'Bath', [0]], ['Cosmetics2', 'Cosmetics2', [1]], ['No category', '', [2]]]);

  const photos = new Map([['SS300', { data: PNG1, width: 1, height: 1, extension: 'png' }], ['SS100', { data: PNG2, width: 2, height: 1, extension: 'png' }]]);
  const bytes = await SS.workbookBytes(SS.buildWorkbook(loadExcelJS(), sh, photos));
  const book = readBack(bytes);
  assert.deepEqual(book.SheetNames, ['Available to sell', 'Bath', 'Cosmetics2', 'No category', 'By date']);
  const cell = (tab, a) => (book.Sheets[tab][a] ? book.Sheets[tab][a].v : undefined);
  assert.equal(cell('Bath', 'A1'), 'M1S · Available to sell · Bath');
  assert.equal(cell('No category', 'A1'), 'M1S · Available to sell · No category');
  assert.deepEqual('ABCDEFGHIJK'.split('').map(c => cell('Cosmetics2', c + '2')), SS.COLS.map(c => c.h));
  assert.deepEqual(['B3', 'E3', 'F3', 'B4'].map(a => cell('Bath', a)), ['SS300', 'Bath', 'Gloves', undefined]);
  assert.deepEqual(['B3', 'E3', 'F3', 'B4'].map(a => cell('No category', a)), ['SS500', undefined, undefined, undefined]);
  // Every category tab's cells are the all-styles tab's, row for row, and together they hold every style once.
  const tabs = SS.categoryTabs(sh.rows);
  assert.deepEqual(tabs.flatMap(t => t.rows).sort(), sh.rows.map((r, n) => n));
  for (const t of tabs) {
    t.rows.forEach((at, n) => {
      for (const c of 'BCDEFGHIJK') assert.deepEqual(cell(t.name, c + (n + 3)), cell('Available to sell', c + (at + 3)), `${t.name} ${c}${n + 3}`);
    });
    assert.equal(cell(t.name, 'B' + (t.rows.length + 3)), undefined, `${t.name}: no more rows`);
  }
  // Two pictures in the file, each shown twice: on the all-styles tab and on its category's.
  const parts = xlsxParts(bytes);
  assert.equal([...parts.keys()].filter(p => /^xl\/media\//.test(p)).length, 2);
  const shown = [];
  for (const [p, b] of parts) {
    const m = /^xl\/drawings\/(drawing\d+)\.xml$/.exec(p);
    if (!m) continue;
    const rels = Buffer.from(parts.get(`xl/drawings/_rels/${m[1]}.xml.rels`)).toString('utf8');
    const target = Object.fromEntries([...rels.matchAll(/<Relationship [^>]*>/g)].map(r => [/Id="([^"]+)"/.exec(r[0])[1], /Target="\.\.\/media\/([^"]+)"/.exec(r[0])[1]]));
    for (const a of Buffer.from(b).toString('utf8').matchAll(/<xdr:row>(\d+)<\/xdr:row>[^]*?r:embed="([^"]+)"/g)) shown.push(`${m[1]} row ${a[1]} ${target[a[2]]}`);
  }
  assert.equal(shown.length, 4, shown.join('; '));
  assert.equal(new Set(shown.map(s => s.split(' ').pop())).size, 2, 'the same two pictures on every tab');
});

test("a style's category and sub category: the ATS report's, else the page's (UPC_MASTER), the pair from one place", () => {
  const sell = mainSell();
  const set = { SS050: ['', '', 'Nails', 'Polish'], SS100: ['Cosmetics1', 'Lip Balm', 'Cosmetics', 'Balm'], SS300: ['Bath', '', 'Body', 'Gloves'], SS500: ['', 'Stray', '', ''] };
  for (const s of sell.styles) if (set[s.style]) [s.atsCategory, s.atsSub, s.category, s.sub] = set[s.style];
  const sh = SS.salesRows(sell, { today: TODAY });
  assert.deepEqual(sh.rows.map(r => [r.style, r.category, r.sub]), [['SS300', 'Bath', ''], ['SS100', 'Cosmetics1', 'Lip Balm'], ['SS500', '', '']]);
  // The Pull from China sheet the same way (SS050 is only on it: the page's pair).
  assert.deepEqual([...new Set(SS.chinaRows(sell, { today: TODAY }).rows.map(r => [r.style, r.category, r.sub].join('|')))],
    ['SS300|Bath|', 'SS100|Cosmetics1|Lip Balm', 'SS200|Lip|', 'SS050|Nails|Polish']);
  // As the selling engine gives them: the ATS report's category with its own sub category (never UPC_MASTER's).
  const s1 = sellOf([alloc({ dates: { factoryCrd: D(12, 6) } })], atsOf({ SS100: [10, 0] }));
  assert.deepEqual([s1.styles[0].category, s1.styles[0].atsCategory, s1.styles[0].atsSub], ['Lip', 'Lip', '']);
  const ats = atsOf({ SS100: [10, 0] });
  ats.styles.get('SS100').sub = 'Gloss';
  const s2 = sellOf([alloc({ dates: { factoryCrd: D(12, 6) } })], ats);
  assert.deepEqual([s2.styles[0].atsCategory, s2.styles[0].atsSub], ['Lip', 'Gloss']);
  ats.styles.get('SS100').category = '';
  const s3 = sellOf([alloc({ dates: { factoryCrd: D(12, 6) } })], ats);
  assert.deepEqual([s3.styles[0].atsCategory, s3.styles[0].atsSub], ['', ''], 'a sub category with no category is not the pair from the report');
  assert.deepEqual(SS.salesRows(s3, { today: TODAY }).rows.map(r => [r.category, r.sub]), [['Lip', '']]);
});

test("category tab names: what Excel won't take made a space, at most 31 characters, never History or another tab's name; never joined", () => {
  const cats = ['Hair/Nails', 'Cosmetics1', 'cosmetics1', 'Cosmetics10', 'Cosmetics2', 'By Date', 'history', 'A category name far longer than Excel allows', '[Lip]: *?', '  ', "'Quoted'"];
  const tabs = SS.categoryTabs(cats.map(category => ({ category: category.trim() })));
  const name = Object.fromEntries(tabs.map(t => [t.category, t.name]));
  assert.deepEqual(name, { 'Hair/Nails': 'Hair Nails', Cosmetics1: 'Cosmetics1', cosmetics1: 'cosmetics1 (2)', Cosmetics10: 'Cosmetics10', Cosmetics2: 'Cosmetics2',
    'By Date': 'By Date (2)', history: 'history (category)', 'A category name far longer than Excel allows': 'A category name far longer than',
    '[Lip]: *?': 'Lip', '': 'No category', "'Quoted'": 'Quoted' });
  for (const t of tabs) {
    assert.ok(t.name.length >= 1 && t.name.length <= 31, t.name);
    assert.doesNotMatch(t.name, /[[\]:*?/\\]|^'|'$/, t.name);
  }
  assert.equal(new Set(tabs.map(t => t.name.toLowerCase()).concat(['available to sell', 'by date'])).size, tabs.length + 2, 'every tab name differs');
  const order = tabs.map(t => t.category);
  assert.ok(order.indexOf('Cosmetics1') < order.indexOf('Cosmetics2') && order.indexOf('Cosmetics2') < order.indexOf('Cosmetics10'), order.join(' | '));
  assert.equal(order[order.length - 1], '', 'no category last');
  // A long name that's taken keeps its tag inside 31 characters.
  const long = SS.categoryTabs(['x'.repeat(40), 'X'.repeat(40)].map(category => ({ category })));
  assert.deepEqual(long.map(t => t.name), ['X'.repeat(31), 'x'.repeat(27) + ' (2)']);
  assert.deepEqual(SS.categoryTabs([]), []);
});

test('the workbook without photos: no pictures, ordinary rows; never an empty sheet', async () => {
  const sh = SS.salesRows(mainSell(), { today: TODAY });
  const parts = xlsxParts(await SS.workbookBytes(SS.buildWorkbook(loadExcelJS(), sh, null)));
  assert.equal([...parts.keys()].filter(p => /^xl\/(media|drawings)\//.test(p)).length, 0);
  assert.doesNotMatch(Buffer.from(parts.get('xl/worksheets/sheet1.xml')).toString('utf8'), /<row r="3"[^>]* ht="82.5"/);
  // A photo for a style that isn't on the sheet is never placed.
  const one = xlsxParts(await SS.workbookBytes(SS.buildWorkbook(loadExcelJS(), sh, { SS999: { data: PNG1, width: 1, height: 1, extension: 'png' } })));
  assert.equal([...one.keys()].filter(p => /^xl\/media\//.test(p)).length, 0);

  // Photos as a list in the rows' order (the page's: each style's own picture, whatever its photo key): placed on those
  // rows only; null for none; an empty list is no photos.
  const list = sh.rows.map((r, i) => (i === 0 ? { data: PNG1, width: 1, height: 1, extension: 'png' } : i === 2 ? { data: PNG2, width: 2, height: 1, extension: 'png' } : null));
  const lp = xlsxParts(await SS.workbookBytes(SS.buildWorkbook(loadExcelJS(), sh, list)));
  assert.equal([...lp.keys()].filter(p => /^xl\/media\//.test(p)).length, 2);
  const ld = Buffer.from(lp.get('xl/drawings/drawing1.xml')).toString('utf8');
  assert.deepEqual([...ld.matchAll(/<xdr:from><xdr:col>(\d+)<\/xdr:col><xdr:colOff>\d+<\/xdr:colOff><xdr:row>(\d+)<\/xdr:row>/g)].map(m => [+m[1], +m[2]]), [[0, 2], [0, 4]]);
  assert.match(Buffer.from(lp.get('xl/worksheets/sheet1.xml')).toString('utf8'), new RegExp(`<row r="3"[^>]* ht="${SS.PHOTO_ROW_PT}"`));
  const none = xlsxParts(await SS.workbookBytes(SS.buildWorkbook(loadExcelJS(), sh, sh.rows.map(() => null))));
  assert.equal([...none.keys()].filter(p => /^xl\/(media|drawings)\//.test(p)).length, 0);
  assert.doesNotMatch(Buffer.from(none.get('xl/worksheets/sheet1.xml')).toString('utf8'), /<row r="3"[^>]* ht="82.5"/);

  const E = loadExcelJS();
  assert.throws(() => SS.buildWorkbook(E, SS.salesRows(sellOf([alloc()], null), { today: TODAY }), null), /There's no ATS report/);
  assert.throws(() => SS.buildWorkbook(E, SS.salesRows(sellOf([], atsOf({ SS100: [0, 0] })), { today: TODAY }), null), /would be empty/);
  assert.throws(() => SS.buildWorkbook(null, sh, null), /Excel library didn't load/);
});

test("the file keeps the page's fingerprint of it in its keywords (a hash only), so a later look can tell it holds the same sheet", async () => {
  const sh = SS.salesRows(mainSell(), { today: TODAY }), E = loadExcelJS(), print = 'c0ffee'.repeat(10) + 'beef';
  const core = async o => Buffer.from(xlsxParts(await SS.workbookBytes(SS.buildWorkbook(E, sh, null, o))).get('docProps/core.xml')).toString('utf8');
  assert.match(await core({ print }), new RegExp(`<cp:keywords>${print}</cp:keywords>`));
  // As the page reads it back (SheetJS, the book's properties only).
  assert.equal(loadXLSX().read(await SS.workbookBytes(SS.buildWorkbook(E, sh, null, { print })), { type: 'array', bookProps: true }).Props.Keywords, print);
  assert.match(await core(undefined), /<cp:keywords><\/cp:keywords>/);
  assert.match(await core({ print: '  ' }), /<cp:keywords><\/cp:keywords>/);
  // Only a hash: never any other text.
  for (const bad of ['Factory A', 'a/b', 'x'.repeat(129), '<cp:x>']) assert.throws(() => SS.buildWorkbook(E, sh, null, { print: bad }), /fingerprint isn't a hash/, bad);
});

test('NO LEAK: no part of the Excel file holds a name, number, flag or problem that must stay inside', async () => {
  const sell = leakySell();
  for (const sh of [SS.salesRows(sell, { today: TODAY }), SS.chinaRows(sell, { today: TODAY })]) {
  const photos = new Map(sh.rows.map(r => [r.photoKey, { data: PNG1, width: 1, height: 1, extension: 'png' }]));
  const parts = xlsxParts(await SS.workbookBytes(SS.buildWorkbook(loadExcelJS(), sh, photos, { print: 'c0ffee'.repeat(10) + 'beef' })));
  assert.ok(parts.has('xl/worksheets/sheet1.xml') && parts.has('xl/sharedStrings.xml') && parts.has('docProps/core.xml'), [...parts.keys()].join(', '));
  const all = text(parts);
  assert.ok(all.includes('SS700') && all.includes('Item SS100'), 'the scan reads the sheet text');
  noMarks(all, 'the Excel file');
  for (const f of new Set([...sell.styles, ...sell.lots].flatMap(x => x.flags || []).filter(f => f.includes('-')))) assert.ok(!all.includes(f), `the flag ${f} is in the Excel file`);
  // Its author is the tool, never a person.
  assert.match(Buffer.from(parts.get('docProps/core.xml')).toString('utf8'), /<dc:creator>Shalam Tool<\/dc:creator>[^]*<cp:lastModifiedBy>Shalam Tool<\/cp:lastModifiedBy>/);
  }
});
