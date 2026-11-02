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
// A parseAts result: { style: [inWh, orders] }.
function atsOf(map, reportDay = D(10, 31)) {
  const styles = new Map(Object.entries(map).map(([style, v]) => [style, { style, styleRaw: style, desc: 'ERP ' + style,
    category: 'Lip', inWh: v[0], orders: v[1], erpAts: null, warehouses: ['WH 1'], rows: [{ row: 3 }], unknown: v[0] == null || v[1] == null }]));
  return { ok: true, reportDay, reportDayFrom: 'file name', styles, issues: [], counts: {} };
}
const sellOf = (allocs, ats) => SELL.buildSell({ stages: stagesOf(allocs), ats, today: TODAY });

// SS050: not on the ATS report, one lot dated from its order date (an estimate). SS100: stock, a lot whose pull-from
// date has passed, two lots on the same pull-from date, an estimate, one on the water. SS200: no case pack. SS300: lots
// in two case packs. SS400: nothing available (left off).
function mainSell() {
  const as = [
    alloc({ style: 'SS050', styleRaw: 'SS050', upc: '000000000500', units: 96, dates: { orderDate: D(10, 1) } }),
    alloc({ po: '900101', bucket: 'booked', units: 240, etd: D(10, 20), stage: 'Booked' }),
    alloc({ po: '900102', units: 480, dates: { factoryCrd: D(12, 6) } }),
    alloc({ po: '900103', units: 120, dates: { factoryCrd: D(12, 6) } }),
    alloc({ po: '900104', units: 100, dates: { orderDate: D(10, 1) } }),
    alloc({ po: '900105', bucket: 'water', units: 300, eta: D(11, 20), stage: 'On the water' }),
    alloc({ style: 'SS200', styleRaw: 'SS200', upc: '000000000200', master: null, units: 70, dates: { factoryCrd: D(12, 6) } }),
    alloc({ style: 'SS300', styleRaw: 'SS300', upc: '000000000300', master: 24, units: 48, dates: { factoryCrd: D(12, 6) } }),
    alloc({ style: 'SS300', styleRaw: 'SS300', upc: '000000000301', master: 12, units: 36, po: '900106', dates: { factoryCrd: D(12, 6) } }),
    alloc({ style: 'SS300', styleRaw: 'SS300', upc: '000000000300', master: 12, units: 30, po: '900107', dates: { factoryCrd: D(12, 25) } }),
  ];
  return sellOf(as, atsOf({ SS100: [500, 20], SS200: [50, 0], SS300: [10, 10], SS400: [0, 0] }));
}
const ROW_KEYS = ['style', 'upc', 'desc', 'casePack', 'now', 'nowCases', 'fromChina', 'fromChinaCases', 'firstFromChina', 'arriving',
  'arrivingCases', 'firstArrival', 'total', 'totalCases', 'photoKey'];

// ---------- refusals

test('no sheet when Available to sell is not worked out, the ATS report is missing or unusable, or the units don\'t add up', () => {
  const nl = SS.salesRows(SELL.buildSell({ stages: { ok: false, notLoaded: ['Container Tracking'] }, ats: null, today: TODAY }), { today: TODAY });
  assert.equal(nl.ok, false);
  assert.equal(nl.reason, "Available to sell couldn't be worked out (Container Tracking didn't load), so there's no sales sheet.");
  assert.deepEqual([nl.rows, nl.byDate, nl.counts.styles], [[], [], 0]);
  assert.equal(SS.salesRows(null, { today: TODAY }).ok, false);

  const as = [alloc({ dates: { factoryCrd: D(12, 6) } })];
  const none = SS.salesRows(sellOf(as, null), { today: TODAY });
  assert.equal(none.ok, false);
  assert.equal(none.reason, "There's no ATS report, so stock available now isn't known and there's no sales sheet.");
  const bad = SS.salesRows(sellOf(as, { ok: false, error: 'the In WH column is missing' }), { today: TODAY });
  assert.equal(bad.ok, false);
  assert.equal(bad.reason, "The ATS report couldn't be used (the In WH column is missing), so stock available now isn't known and there's no sales sheet.");

  const r = sellOf(as, atsOf({ SS100: [10, 0] }));
  assert.equal(SS.salesRows(r, { today: TODAY }).ok, true);
  r.issues.push({ code: 'balance', message: 'invented' });
  const off = SS.salesRows(r, { today: TODAY });
  assert.equal(off.ok, false);
  assert.match(off.reason, /doesn't add up \(a bug in Shalam Tool\)/);
  assert.equal(off.rows.length, 0);

  assert.throws(() => SS.salesRows(r, { today: '11/2/2026' }), /today must be a day number/);
  assert.equal(SS.salesRows(sellOf(as, atsOf({ SS100: [10, 0] }))).asOf, TODAY, "today defaults to Available to sell's own");
});

// ---------- the rows

test('one row per style with something available, by style; only the listed fields', () => {
  const sell = mainSell(), sh = SS.salesRows(sell, { today: TODAY });
  assert.equal(sh.ok, true);
  assert.equal(sh.reason, '');
  assert.equal(sh.asOf, TODAY);
  assert.equal(sh.reportDay, D(10, 31));
  assert.deepEqual(sh.rows.map(r => r.style), ['SS050', 'SS100', 'SS200', 'SS300']);
  for (const r of sh.rows) {
    assert.deepEqual(Object.keys(r), ROW_KEYS, r.style);
    assert.deepEqual(Object.keys(r.firstFromChina), ['date', 'now', 'estimate']);
    assert.equal(r.total, r.now + r.fromChina + r.arriving, r.style);
    assert.ok(r.total > 0);
  }
  for (const b of sh.byDate) assert.deepEqual(Object.keys(b), ['style', 'desc', 'when', 'estimate', 'where', 'units', 'cases']);
  assert.deepEqual(sh.counts, { styles: 4, leftOff: 1, byDateRows: sh.byDate.length, noCasePack: 1, mixedCasePack: 1 });
  // The same units as the page's Available to sell.
  for (const r of sh.rows) {
    const s = sell.styles.find(x => x.styleRaw === r.style);
    assert.deepEqual([r.now, r.fromChina, r.arriving, r.total], [s.now ?? 0, s.fromChina, s.arriving, s.total], r.style);
  }
});

test('units and full cases: stock now, from China, arriving, total; a style\'s UPCs and photo key', () => {
  const sh = SS.salesRows(mainSell(), { today: TODAY }), r = sh.rows.find(x => x.style === 'SS100');
  assert.equal(r.upc, '000111222333');
  assert.equal(r.desc, 'Item SS100');
  assert.equal(r.casePack, 24);
  assert.deepEqual([r.now, r.nowCases], [480, 20]);
  // 240 (10 cases) + 480 (20) + 120 (5) + 100 (4 full cases, 4 loose): cases are whole cases lot by lot.
  assert.deepEqual([r.fromChina, r.fromChinaCases], [940, 39]);
  assert.deepEqual([r.arriving, r.arrivingCases], [300, 12]);
  assert.deepEqual([r.total, r.totalCases], [1720, 71]);
  assert.equal(r.photoKey, 'SS100');
  assert.equal(SS.photoKey(' ab 12/3-f_x '), 'AB123-F_X');
  assert.equal(SS.photoKey('AB100-F'), 'AB100-F', 'exact key only: -F is kept');
  // SS300 has two UPCs.
  assert.equal(sh.rows.find(x => x.style === 'SS300').upc, '000000000300, 000000000301');
});

test('cases are null, never guessed, when the case pack is unknown or mixed', () => {
  const sh = SS.salesRows(mainSell(), { today: TODAY });
  const none = sh.rows.find(x => x.style === 'SS200'), mixed = sh.rows.find(x => x.style === 'SS300');
  assert.equal(none.casePack, null);
  assert.deepEqual([none.now, none.nowCases, none.fromChina, none.fromChinaCases, none.totalCases], [50, null, 70, null, null]);
  assert.equal(mixed.casePack, 'mixed');
  assert.deepEqual([mixed.now, mixed.nowCases, mixed.fromChina, mixed.fromChinaCases, mixed.arrivingCases, mixed.totalCases], [0, null, 114, null, null, null]);
});

test('first dates: now when the pull-from date has passed; about (an estimate) when it is 75 days after the order date', () => {
  const sh = SS.salesRows(mainSell(), { today: TODAY });
  const r = sh.rows.find(x => x.style === 'SS100');
  assert.deepEqual(r.firstFromChina, { date: D(10, 10), now: true, estimate: false });
  assert.equal(r.firstArrival, D(11, 20));
  const est = sh.rows.find(x => x.style === 'SS050');
  assert.deepEqual(est.firstFromChina, { date: D(10, 1) + 75, now: false, estimate: true });
  assert.equal(est.firstArrival, null);
  assert.equal(est.now, 0, 'not on the ATS report: no stock now');
  assert.deepEqual(sh.rows.find(x => x.style === 'SS200').firstFromChina, { date: D(12, 1), now: false, estimate: false });
  // A firm date beats an estimate on the same day.
  const both = SS.salesRows(sellOf([alloc({ units: 24, dates: { orderDate: D(9, 22) } }), alloc({ units: 48, dates: { factoryCrd: D(12, 11) } })],
    atsOf({ SS100: [0, 0] })), { today: TODAY });
  assert.deepEqual(both.rows[0].firstFromChina, { date: D(12, 6), now: false, estimate: false });
});

// ---------- the schedule

test('by date: stock now, then each date and place, units summed over lots with the same date and place', () => {
  const sh = SS.salesRows(mainSell(), { today: TODAY });
  const of = st => sh.byDate.filter(b => b.style === st).map(b => [b.when, b.where, b.units, b.cases, b.estimate]);
  assert.deepEqual(of('SS100'), [
    ['now', 'In stock', 480, 20, false],
    ['now', 'From China', 240, 10, false],
    [D(11, 20), 'Arriving', 300, 12, false],
    [D(12, 1), 'From China', 600, 25, false],
    [D(10, 1) + 75, 'From China', 100, 4, true],
  ]);
  // Cases per the lots' Master: two packs on one date can't be counted in cases.
  assert.deepEqual(of('SS300'), [[D(12, 1), 'From China', 84, null, false], [D(12, 20), 'From China', 30, 2, false]]);
  assert.deepEqual(of('SS200'), [['now', 'In stock', 50, null, false], [D(12, 1), 'From China', 70, null, false]]);
  assert.ok(sh.byDate.every(b => b.units > 0));
  assert.ok(!sh.byDate.some(b => b.style === 'SS400'), 'a style with nothing available has no schedule');
  // Every unit on a row is on its schedule.
  for (const r of sh.rows) {
    const mine = sh.byDate.filter(b => b.style === r.style);
    assert.equal(mine.reduce((t, b) => t + b.units, 0), r.total, r.style);
    assert.ok(mine.every(b => b.desc === r.desc));
  }
  // By style, then date (now first).
  const order = sh.byDate.map(b => [b.style, b.when === 'now' ? -1 : b.when]);
  assert.deepEqual(order, order.slice().sort((a, b) => (a[0] < b[0] ? -1 : a[0] > b[0] ? 1 : a[1] - b[1])));
});

test('by date: a lot with no date to sell from is listed last with no date; customer orders beyond stock come off the lots', () => {
  const as = [alloc({ units: 100, master: 10, dates: { factoryCrd: D(12, 6) } }), alloc({ units: 50, master: 10, po: '900109' }),
    alloc({ bucket: 'port', units: 40, master: 10, eta: D(10, 25), stage: 'At port' })];
  const sh = SS.salesRows(sellOf(as, atsOf({ SS100: [20, 50] })), { today: TODAY });
  // Orders 50 against stock 20: 30 come off the earliest lot (at port, ETA 10/25), which keeps 10.
  assert.deepEqual(sh.byDate.map(b => [b.when, b.where, b.units, b.cases]), [
    [D(10, 25), 'Arriving', 10, 1], [D(12, 1), 'From China', 100, 10], [null, 'From China', 50, 5]]);
  const r = sh.rows[0];
  assert.deepEqual([r.now, r.nowCases, r.arriving, r.fromChina, r.total], [0, 0, 10, 150, 160]);
  assert.deepEqual(r.firstFromChina, { date: D(12, 1), now: false, estimate: false });
});

test('by date: the ATS report\'s In Trans shipments are Arriving on their ETA, in the style\'s case pack', () => {
  const ats = atsOf({ SS100: [0, 0] });
  const s = ats.styles.get('SS100');
  // Its Import PO Backlog still has our 24 in China, so its In Trans doesn't come off them.
  s.rows = [{ row: 3, warehouse: 'WH 1', inWh: 0, orders: 0, inTrans: 72, inTransList: [48, 24], inTransBad: [], etaParts: [D(11, 15), D(11, 29)],
    eta: [D(11, 15), D(11, 29)], backlog: 24, backlogList: [24], poParts: [] }];
  ats.columns = ['style', 'inWh', 'orders', 'inTrans', 'eta', 'backlog'];
  const sh = SS.salesRows(sellOf([alloc({ units: 24, master: 12, dates: { factoryCrd: D(12, 6) } })], ats), { today: TODAY });
  assert.deepEqual(sh.byDate.map(b => [b.when, b.where, b.units, b.cases]), [
    [D(11, 15), 'Arriving', 48, 4], [D(11, 29), 'Arriving', 24, 2], [D(12, 1), 'From China', 24, 2]]);
  assert.deepEqual([sh.rows[0].arriving, sh.rows[0].arrivingCases, sh.rows[0].firstArrival], [72, 6, D(11, 15)]);
});

// ---------- nothing that must stay inside reaches the sheet

// Distinctive marker strings in every field the salesperson must never see: factory, customer, PO, PO type, section,
// container, round, people's names, file names, links, ERP warehouse, and the numbers for In WH, customer orders, sold
// and oversold.
const MARK = { factory: 'Factory QXLEAKFAC', customer: 'Customer QXLEAKCUST', po: '987650-QXLEAKPO', poShown: 'PO QXLEAKSHOWN', ptype: 'QXLEAKPTYPE',
  section: 'Section QXLEAKSECT', round: 'Round QXLEAKROUND', container: 'QXLC1234567', person: 'Pat QXLEAKPERSON', file: 'QXLEAKFILE.xlsx',
  link: 'https://contoso.sharepoint.example/sites/QXLEAKSITE/QXLEAKLINK', warehouse: 'WH QXLEAKWH', erpPo: '987651-QXLEAKERPPO', category: 'QXLEAKCAT',
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
  const ats = atsOf({ SS100: [77791, 4321], SS700: [10, 0], SS701: [0, 0], SS799: [0, 27182] });
  for (const [k, e] of ats.styles) {
    e.warehouses = [MARK.warehouse];
    e.rows = [{ row: 3, warehouse: MARK.warehouse, inWh: e.inWh, orders: e.orders, inTrans: k === 'SS700' ? 60 : null, inTransList: k === 'SS700' ? [60] : [],
      inTransBad: [], etaParts: k === 'SS700' ? [D(11, 28)] : [], eta: [], backlog: null, backlogList: [], poParts: k === 'SS700' ? [MARK.erpPo] : [] }];
  }
  Object.assign(ats, { columns: ['style', 'inWh', 'orders', 'inTrans', 'eta', 'po'], fileName: MARK.file, savedBy: MARK.person });
  const sell = SELL.buildSell({ stages: stagesOf(as, { category: MARK.category }), ats, today: TODAY });
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
  // The markers are really in what the sheet is made from.
  const src = allText(sell);
  for (const m of [MARK.factory, MARK.customer, MARK.po, MARK.section, MARK.round, MARK.container, MARK.person, MARK.file, MARK.link, MARK.warehouse, MARK.erpPo, MARK.category, MARK.issue]) assert.ok(src.includes(m), m);
  const s1 = sell.styles.find(s => s.style === 'SS100');
  assert.deepEqual([s1.inWh, s1.orders, s1.soldChina], [77791, 4321, 31415]);
  assert.equal(sell.styles.find(s => s.style === 'SS799').oversold, 27182);
  const flags = new Set([...sell.styles, ...sell.lots].flatMap(x => x.flags || []).filter(f => f.includes('-')));
  assert.ok(flags.size >= 3, [...flags].join(', '));

  const sh = SS.salesRows(sell, { today: TODAY });
  assert.equal(sh.ok, true);
  assert.deepEqual(sh.rows.map(r => r.style), ['SS100', 'SS700', 'SS701']);
  const out = allText(sh);
  noMarks(out, 'the sales rows');
  for (const f of flags) assert.ok(!out.includes(f), `the flag ${f} is in the sales rows`);
  for (const i of sell.issues) assert.ok(!out.includes(i.message), `a problem is in the sales rows: ${i.message.slice(0, 80)}`);
  // Only the allow-listed fields, at every level.
  for (const r of sh.rows) assert.deepEqual(Object.keys(r), ROW_KEYS);
  assert.deepEqual(Object.keys(sh).sort(), ['asOf', 'byDate', 'counts', 'ok', 'reason', 'reportDay', 'rows']);
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

test('the workbook: two tabs, the title with its dates, frozen photo and style, a filter, numbers and dates as Excel keeps them', async () => {
  const sh = SS.salesRows(mainSell(), { today: TODAY });
  const photos = new Map([['SS100', { data: PNG1, width: 1, height: 1, extension: 'png' }], ['SS300', { data: PNG2.buffer.slice(0), width: 2, height: 1, extension: 'png' }]]);
  const wb = SS.buildWorkbook(loadExcelJS(), sh, photos);
  const bytes = await SS.workbookBytes(wb);
  assert.ok(bytes instanceof Uint8Array && bytes.length > 5000, `${bytes.length} bytes`);
  const book = readBack(bytes);
  assert.deepEqual(book.SheetNames, ['Available to sell', 'By date']);
  const ws = book.Sheets['Available to sell'], v = a => (ws[a] ? ws[a].v : undefined);
  assert.equal(v('A1'), 'M1S · Available to sell');
  assert.equal(v('D1'), 'as of 11/2/2026 · stock from the ERP report of 10/31/2026');
  assert.deepEqual('ABCDEFGHIJKLMNO'.split('').map(c => v(c + '2')), SS.COLS.map(c => c.h));
  assert.deepEqual(SS.COLS.map(c => c.h), ['Photo', 'Style', 'UPC', 'Description', 'Case pack', 'Available now (units)', 'Available now (cases)',
    'Available from China (units)', 'Available from China (cases)', 'First date from China', 'Arriving (units)', 'Arriving (cases)', 'First arrival',
    'Total available (units)', 'Total available (cases)']);
  // Row 3 SS050 (an estimate, nothing now), row 4 SS100.
  assert.deepEqual(['B3', 'E3', 'F3', 'G3', 'H3', 'I3', 'J3', 'K3', 'M3', 'N3', 'O3'].map(v), ['SS050', 24, undefined, undefined, 96, 4, `about ${C.fmtDay(D(10, 1) + 75)}`, undefined, undefined, 96, 4]);
  assert.deepEqual(['B4', 'C4', 'D4', 'E4', 'F4', 'G4', 'H4', 'I4', 'J4', 'K4', 'L4', 'M4', 'N4', 'O4'].map(v),
    ['SS100', '000111222333', 'Item SS100', 24, 480, 20, 940, 39, 'now', 300, 12, serial(D(11, 20)), 1720, 71]);
  assert.equal(ws.C4.t, 's', 'UPCs stay text');
  assert.equal(ws.F4.z, '#,##0');
  assert.equal(ws.M4.z, 'm/d/yyyy');
  assert.deepEqual(['B5', 'E5', 'G5', 'J5', 'O5'].map(v), ['SS200', undefined, undefined, serial(D(12, 1)), undefined]);
  assert.deepEqual(['B6', 'E6', 'I6'].map(v), ['SS300', 'mixed', undefined]);
  assert.equal(v('B7'), undefined, 'four styles, no more');
  const wd = book.Sheets['By date'], w = a => (wd[a] ? wd[a].v : undefined);
  assert.deepEqual(['A1', 'B1', 'C1', 'D1', 'E1', 'F1'].map(w), ['Style', 'Description', 'Available from', 'Where', 'Units', 'Cases']);
  const sched = [];
  for (let r = 2; w('A' + r) !== undefined; r++) sched.push(['A', 'C', 'D', 'E', 'F'].map(c => w(c + r)));
  assert.equal(sched.length, sh.byDate.length);
  assert.deepEqual(sched.filter(x => x[0] === 'SS100'), [['SS100', 'now', 'In stock', 480, 20], ['SS100', 'now', 'From China', 240, 10],
    ['SS100', serial(D(11, 20)), 'Arriving', 300, 12], ['SS100', serial(D(12, 1)), 'From China', 600, 25],
    ['SS100', `about ${C.fmtDay(D(10, 1) + 75)}`, 'From China', 100, 4]]);
  assert.deepEqual(sched.filter(x => x[0] === 'SS300'), [['SS300', serial(D(12, 1)), 'From China', 84, undefined], ['SS300', serial(D(12, 20)), 'From China', 30, 2]]);

  const parts = xlsxParts(bytes);
  const sheet1 = Buffer.from(parts.get('xl/worksheets/sheet1.xml')).toString('utf8'), sheet2 = Buffer.from(parts.get('xl/worksheets/sheet2.xml')).toString('utf8');
  assert.match(sheet1, /<pane xSplit="2" ySplit="2" topLeftCell="C3" activePane="bottomRight" state="frozen"\/>/);
  assert.match(sheet1, /<autoFilter ref="A2:O2"\/>/);
  assert.match(sheet2, /<pane xSplit="1" ySplit="1" topLeftCell="B2" activePane="bottomRight" state="frozen"\/>/);
  assert.match(sheet2, /<autoFilter ref="A1:F1"\/>/);
  // Two photos, in the Photo column on their styles' rows (SS100 on row 4, SS300 on row 6: 0-based 3 and 5), fitted
  // inside the cell; every style row has the photo row height.
  assert.equal([...parts.keys()].filter(p => /^xl\/media\//.test(p)).length, 2);
  const drawing = Buffer.from(parts.get('xl/drawings/drawing1.xml')).toString('utf8');
  const anchors = [...drawing.matchAll(/<xdr:from><xdr:col>(\d+)<\/xdr:col><xdr:colOff>(\d+)<\/xdr:colOff><xdr:row>(\d+)<\/xdr:row><xdr:rowOff>(\d+)<\/xdr:rowOff><\/xdr:from><xdr:ext cx="(\d+)" cy="(\d+)"\/>/g)]
    .map(m => m.slice(1).map(Number));
  assert.deepEqual(anchors.map(a => [a[0], a[2]]), [[0, 3], [0, 5]]);
  const EMU = 9525, cellW = Math.round(15 * 7 + 5), cellH = Math.round(SS.PHOTO_ROW_PT * 96 / 72);
  for (const [, colOff, , rowOff, cx, cy] of anchors) {
    assert.ok(colOff / EMU + cx / EMU <= cellW + 1 && rowOff / EMU + cy / EMU <= cellH + 1, `photo inside its cell: ${[colOff, rowOff, cx, cy].map(n => Math.round(n / EMU))}`);
    assert.ok(colOff > 0 && rowOff > 0, 'centred, not hugging the corner');
  }
  for (let r = 3; r <= 6; r++) assert.match(sheet1, new RegExp(`<row r="${r}"[^>]* ht="${SS.PHOTO_ROW_PT}"`));
});

test('the workbook without photos: no pictures, ordinary rows; never an empty sheet', async () => {
  const sh = SS.salesRows(mainSell(), { today: TODAY });
  const parts = xlsxParts(await SS.workbookBytes(SS.buildWorkbook(loadExcelJS(), sh, null)));
  assert.equal([...parts.keys()].filter(p => /^xl\/(media|drawings)\//.test(p)).length, 0);
  assert.doesNotMatch(Buffer.from(parts.get('xl/worksheets/sheet1.xml')).toString('utf8'), /<row r="3"[^>]* ht="82.5"/);
  // A photo for a style that isn't on the sheet is never placed.
  const one = xlsxParts(await SS.workbookBytes(SS.buildWorkbook(loadExcelJS(), sh, { SS999: { data: PNG1, width: 1, height: 1, extension: 'png' } })));
  assert.equal([...one.keys()].filter(p => /^xl\/media\//.test(p)).length, 0);

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
  const sell = leakySell(), sh = SS.salesRows(sell, { today: TODAY });
  const photos = new Map(sh.rows.map(r => [r.photoKey, { data: PNG1, width: 1, height: 1, extension: 'png' }]));
  const parts = xlsxParts(await SS.workbookBytes(SS.buildWorkbook(loadExcelJS(), sh, photos, { print: 'c0ffee'.repeat(10) + 'beef' })));
  assert.ok(parts.has('xl/worksheets/sheet1.xml') && parts.has('xl/sharedStrings.xml') && parts.has('docProps/core.xml'), [...parts.keys()].join(', '));
  const all = text(parts);
  assert.ok(all.includes('SS700') && all.includes('Item SS100'), 'the scan reads the sheet text');
  noMarks(all, 'the Excel file');
  for (const f of new Set([...sell.styles, ...sell.lots].flatMap(x => x.flags || []).filter(f => f.includes('-')))) assert.ok(!all.includes(f), `the flag ${f} is in the Excel file`);
  // Its author is the tool, never a person.
  assert.match(Buffer.from(parts.get('docProps/core.xml')).toString('utf8'), /<dc:creator>Shalam Tool<\/dc:creator>[^]*<cp:lastModifiedBy>Shalam Tool<\/cp:lastModifiedBy>/);
});
