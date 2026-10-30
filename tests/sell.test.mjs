import test from 'node:test';
import assert from 'node:assert/strict';
import { core, loadXLSX, xlsxBytes, readBook } from './load.mjs';

const C = core('common');
const ST = core('stages');
const A = core('ats');
const SELL = core('sell');
const TODAY = C.ymd(2026, 10, 30);
const D = (m, d, y = 2026) => C.ymd(y, m, d);

// ---------- invented data (POs 9xxxxx, styles AB…, Factory A)

// An allocation as core/stages.js gives it (only the fields the selling view reads).
let nextId = 1;
function alloc(o = {}) {
  const id = 'a' + nextId++;
  const units = o.units ?? 1000, master = o.master === undefined ? 24 : o.master;
  return { id, bucket: 'inProd', units, cases: master ? Math.ceil(units / master) : null, master, style: 'AB100', styleRaw: 'AB100',
    upc: '000111222333', po: '900101', poShown: '', ptype: 'LA', customer: '', factory: 'Factory A', crd: null, od: null,
    etd: null, eta: null, stage: 'In production', flags: [], sheetRow: 10, round: '', containers: [], section: '',
    dates: { ctReady: null, ctRow: null, poBook: null, factoryCrd: null, crdExpected: null, shipBy: null, orderDate: null, original: null },
    ...o, ...(o.dates ? { dates: { ctReady: null, ctRow: null, poBook: null, factoryCrd: null, crdExpected: null, shipBy: null, orderDate: null, original: null, ...o.dates } } : {}) };
}
// A stage engine result around some allocations (its styles: one per style, with the counted UPCs).
function stagesOf(allocs) {
  const styles = new Map();
  for (const a of allocs) {
    if (!styles.has(a.style)) styles.set(a.style, { style: a.style, styleRaw: a.styleRaw, desc: 'Item ' + a.style, category: 'Lip',
      factory: a.factory, factories: [a.factory], upcs: [], allocs: [] });
    const s = styles.get(a.style);
    if (a.upc && !s.upcs.some(u => u.upc === a.upc)) s.upcs.push({ upc: a.upc });
    s.allocs.push(a);
  }
  return { ok: true, allocs, styles: [...styles.values()] };
}
// A parseAts result: { style: [inWh, orders] }.
function atsOf(map, reportDay = D(10, 29)) {
  const styles = new Map(Object.entries(map).map(([style, v]) => [style, { style, styleRaw: style, desc: 'ERP ' + style,
    category: 'Lip', inWh: v[0], orders: v[1], erpAts: null, warehouses: ['WH 1'], rows: [{ row: 3 }], unknown: v[0] == null || v[1] == null }]));
  return { ok: true, reportDay, reportDayFrom: 'file name', styles, issues: [], counts: {} };
}
const sell = (allocs, ats, extra = {}) => SELL.buildSell({ stages: stagesOf(allocs), ats, today: TODAY, ...extra });
const styleOf = (r, s = 'AB100') => r.styles.find(x => x.style === s);
const codes = r => r.issues.map(i => i.code);
// Every lot is taken + available; per style, the lots are what's available from China, arriving and taken by orders.
function sumsHold(r) {
  for (const l of r.lots) assert.equal(l.units, l.taken + l.available, `lot ${l.id}`);
  for (const s of r.styles) {
    const lu = s.lots.reduce((t, l) => t + l.units, 0);
    assert.equal(lu, s.fromChina + s.arriving + s.ordersTaken, `style ${s.style}`);
    assert.equal(s.total, (s.now ?? 0) + s.fromChina + s.arriving);
  }
  const t = r.totals;
  assert.equal(t.total, t.now + t.fromChina + t.arriving);
  assert.equal(t.fromChina, r.lots.filter(l => l.inChina).reduce((s, l) => s + l.available, 0));
  assert.equal(t.arriving, r.lots.filter(l => !l.inChina).reduce((s, l) => s + l.available, 0));
  assert.equal(t.ordersTaken, r.lots.reduce((s, l) => s + l.taken, 0));
}

// ---------- inputs

test('the stage numbers not loaded: not loaded, nothing else; today is required', () => {
  assert.deepEqual(SELL.buildSell({ stages: { ok: false, notLoaded: ['Container Tracking'] }, ats: null, today: TODAY }),
    { ok: false, today: TODAY, notLoaded: ['Container Tracking'] });
  assert.equal(SELL.buildSell({ stages: null, ats: null, today: TODAY }).ok, false);
  assert.throws(() => SELL.buildSell({ stages: stagesOf([]), ats: null }), /today/);
});

// ---------- lots: what can be sold, and what's sold

test('lots are the counted stages that are not sold; China Pickup and direct-to-customer units are Sold, never available', () => {
  const as = [
    alloc({ bucket: 'noCrd', units: 100 }), alloc({ bucket: 'inProd', units: 200 }), alloc({ bucket: 'ready', units: 300 }),
    alloc({ bucket: 'booked', units: 400 }), alloc({ bucket: 'loaded', units: 500 }),
    alloc({ bucket: 'water', units: 600, eta: D(11, 20) }), alloc({ bucket: 'port', units: 700, eta: D(10, 25) }),
    alloc({ bucket: 'pullFromChina', units: 800, ptype: 'China Pickup' }),
    alloc({ bucket: 'pullFromChina', units: 900, flags: ['pull-section'] }),
    alloc({ bucket: 'loaded', units: 50, ptype: 'China Pickup' }),      // never left in a stage by the engine: still Sold
    alloc({ bucket: 'unmatched', units: 70, stage: 'Booked' }), alloc({ bucket: 'unmatched', units: 80, stage: 'Landed' }),
    alloc({ bucket: 'landed', units: 1 }), alloc({ bucket: 'short', units: 2 }), alloc({ bucket: 'ctMore', units: 3 }),
    alloc({ bucket: 'shippedNotOnCt', units: 4 }), alloc({ bucket: 'planning', units: 5 }),
  ];
  const r = sell(as, null);
  assert.equal(r.ok, true);
  assert.equal(r.lots.length, 7);
  assert.deepEqual(r.lots.filter(l => l.inChina).map(l => l.status).sort(), ['Booked', 'In production', 'Loaded', 'No CRD', 'Ready']);
  assert.deepEqual(r.lots.filter(l => !l.inChina).map(l => l.status).sort(), ['At port', 'On the water']);
  const s = styleOf(r);
  assert.equal(s.fromChina, 1500);
  assert.equal(s.arriving, 1300);
  assert.equal(s.soldChina, 1750);
  assert.equal(s.unmatched, 70, 'Unmatched units are counted apart, not available; landed ones not at all');
  assert.equal(r.totals.soldChina, 1750);
  sumsHold(r);
});

// ---------- dates

test('pull-from date in China: 10 days before the ETD; else 5 days before the ready date; else about 75 days after the order', () => {
  const as = [
    alloc({ po: '900201', bucket: 'loaded', etd: D(12, 10), crd: D(11, 20), dates: { factoryCrd: D(11, 20), orderDate: D(8, 1) } }),
    alloc({ po: '900202', bucket: 'inProd', crd: D(12, 1), dates: { factoryCrd: D(12, 1), orderDate: D(8, 1) } }),
    alloc({ po: '900203', bucket: 'noCrd', dates: { orderDate: D(11, 1) } }),
    alloc({ po: '900204', bucket: 'noCrd' }),
  ];
  const r = sell(as, null);
  const lot = po => r.lots.find(l => l.po === po);
  assert.deepEqual([lot('900201').pullFrom, lot('900201').pullRule, lot('900201').estimate], [D(11, 30), 'etd', false]);
  assert.deepEqual([lot('900202').pullFrom, lot('900202').pullRule], [D(11, 26), 'crd']);
  const est = lot('900203');
  assert.deepEqual([est.pullFrom, est.pullRule, est.estimate], [D(1, 15, 2027), 'order', true]);
  assert.deepEqual(est.pullRange, { from: D(11, 1), to: D(1, 15, 2027) });
  assert.ok(est.flags.includes('estimate'));
  const none = lot('900204');
  assert.deepEqual([none.pullFrom, none.pullRule, none.availableDate], [null, 'none', null]);
  assert.ok(none.flags.includes('no-date'));
  assert.equal(SELL.ruleWords(lot('900201')), '10 days before ETD 12/10/2026');
  assert.equal(SELL.ruleWords(lot('900202')), '5 days before CRD 12/1/2026');
  assert.equal(SELL.ruleWords(est), 'about 1/15/2027: 75 days after the order date 11/1/2026');
  assert.equal(SELL.ruleWords(none), 'no ETD, CRD or order date');
  assert.deepEqual(r.counts.byRule, { etd: 1, crd: 1, order: 1, none: 1 });
  const nd = r.issues.filter(i => i.code === 'no-date');
  assert.equal(nd.length, 1);
  assert.equal(nd[0].po, '900204');
  assert.match(nd[0].message, /no ETD, no ready date \(CRD\) and no order date/);
  assert.equal(styleOf(r).firstPull, D(11, 26));
  assert.equal(styleOf(r).firstPullRule, 'crd');
});

test('a pull-from date before today is "now", and the date is kept', () => {
  const r = sell([alloc({ bucket: 'ready', crd: D(9, 1), dates: { shipBy: D(9, 1) } }), alloc({ po: '900102', bucket: 'booked', etd: D(11, 9) })], null);
  const past = r.lots.find(l => l.po === '900101'), edge = r.lots.find(l => l.po === '900102');
  assert.equal(past.pullFrom, D(8, 27));
  assert.equal(past.pullNow, true);
  assert.ok(past.flags.includes('now'));
  assert.equal(edge.pullFrom, D(10, 30), 'today itself is not "now"');
  assert.equal(edge.pullNow, false);
  assert.equal(styleOf(r).firstPull, D(8, 27));
  assert.equal(styleOf(r).firstPullNow, true);
  assert.equal(r.counts.pullNow, 1);
});

test('sailed lots are available on their ETA; no ETA is listed', () => {
  const r = sell([alloc({ bucket: 'water', etd: D(10, 1), eta: D(11, 18) }), alloc({ po: '900102', bucket: 'port', eta: null })], null);
  const w = r.lots.find(l => l.po === '900101'), p = r.lots.find(l => l.po === '900102');
  assert.deepEqual([w.arrival, w.arrivalRule, w.availableDate, w.pullFrom], [D(11, 18), 'eta', D(11, 18), null]);
  assert.deepEqual([p.arrival, p.arrivalRule, p.availableDate], [null, 'none', null]);
  assert.equal(SELL.ruleWords(w), 'ETA 11/18/2026');
  assert.equal(styleOf(r).firstArrival, D(11, 18));
  const nd = r.issues.find(i => i.code === 'no-date');
  assert.match(nd.message, /at port have no ETA/);
});

test("ready date: Container Tracking's Ready Date, else the PO Book's, Factory CRD, CRD Expected, Ship by; each source kept", () => {
  const r = sell([
    alloc({ po: '900301', dates: { ctReady: D(12, 5), factoryCrd: D(12, 1), crdExpected: D(11, 30), shipBy: D(12, 3) } }),
    alloc({ po: '900302', dates: { poBook: D(1, 10, 2027), factoryCrd: null, shipBy: null } }),
    alloc({ po: '900303', dates: { factoryCrd: D(12, 1), crdExpected: D(12, 4), shipBy: D(12, 2) } }),
    alloc({ po: '900304', dates: { crdExpected: D(12, 4), shipBy: D(12, 2) } }),
    alloc({ po: '900305', dates: { shipBy: D(12, 2) } }),
  ], null);
  const lot = po => r.lots.find(l => l.po === po);
  assert.deepEqual([lot('900301').readyDate, lot('900301').readyFrom], [D(12, 5), 'ctReady']);
  assert.deepEqual(lot('900301').readySources.map(s => s.key), ['ctReady', 'factoryCrd', 'crdExpected', 'shipBy']);
  assert.equal(lot('900301').dates.shipBy, D(12, 3));
  assert.deepEqual([lot('900302').readyDate, lot('900302').readyFrom], [D(1, 10, 2027), 'poBook']);
  assert.deepEqual([lot('900303').readyDate, lot('900303').readyFrom], [D(12, 1), 'factoryCrd']);
  assert.deepEqual([lot('900304').readyDate, lot('900304').readyFrom], [D(12, 4), 'crdExpected']);
  assert.deepEqual([lot('900305').readyDate, lot('900305').readyFrom], [D(12, 2), 'shipBy']);
  assert.ok(!codes(r).includes('dates-disagree'), 'within 7 days: no flag');
});

test('dates-disagree: sources more than 7 days apart, one line per PO + style naming each date; sailed lots not listed', () => {
  const r = sell([
    alloc({ po: '900401', units: 600, dates: { ctReady: D(12, 20), factoryCrd: D(12, 1), shipBy: D(11, 15) } }),
    alloc({ po: '900401', units: 400, bucket: 'booked', dates: { ctReady: D(12, 20), factoryCrd: D(12, 1), shipBy: D(11, 15) } }),
    alloc({ po: '900402', dates: { factoryCrd: D(12, 1), crdExpected: D(12, 8) } }),                 // exactly 7: fine
    alloc({ po: '900403', dates: { factoryCrd: D(12, 1), crdExpected: D(12, 9) } }),                 // 8: listed
    alloc({ po: '900404', bucket: 'water', eta: D(11, 20), dates: { factoryCrd: D(9, 1), shipBy: D(8, 1) } }),
  ], null);
  const dd = r.issues.filter(i => i.code === 'dates-disagree');
  assert.deepEqual(dd.map(i => i.po), ['900401', '900403']);
  assert.match(dd[0].message, /up to 35 days apart: Container Tracking's Ready Date 12\/20\/2026 \(Container Tracking\), Factory CRD 12\/1\/2026 \(PO Status Tracker\), Ship by 11\/15\/2026 \(Master Sheet Live\)\. The ready date used is 12\/20\/2026/);
  assert.equal(dd[0].sources.length, 3);
  assert.equal(r.lots.filter(l => l.flags.includes('dates-disagree')).length, 3);
  assert.ok(!r.lots.find(l => l.po === '900404').flags.includes('dates-disagree'));
  assert.equal(r.counts.datesDisagreeSailedLots, 1);
  assert.equal(r.counts.datesDisagreePairs, 2);
  assert.ok(styleOf(r).flags.includes('dates-disagree'));
});

// ---------- customer orders: stock first, then the earliest lot

test('customer orders come off stock first: Available now is In WH − orders, never below 0', () => {
  const r = sell([alloc({ bucket: 'inProd', units: 1000, crd: D(12, 1), dates: { factoryCrd: D(12, 1) } })], atsOf({ AB100: [5000, 1200] }));
  const s = styleOf(r);
  assert.deepEqual([s.inWh, s.orders, s.stockNet, s.now, s.fromChina, s.ordersTaken, s.oversold], [5000, 1200, 3800, 3800, 1000, 0, 0]);
  assert.equal(s.total, 4800);
  assert.equal(s.onReport, true);
  assert.equal(r.atsMissing, false);
  assert.equal(r.reportDay, D(10, 29));
  sumsHold(r);
});

test('orders beyond stock come off the lots in date order (no date last, then PO); what is left is oversold', () => {
  const as = [
    alloc({ po: '900503', units: 500, bucket: 'noCrd' }),                                              // no date: last
    alloc({ po: '900502', units: 300, bucket: 'inProd', crd: D(12, 1), dates: { factoryCrd: D(12, 1) } }),   // pull 11/26
    alloc({ po: '900501', units: 200, bucket: 'water', eta: D(11, 10) }),                             // arrives 11/10: first
    alloc({ po: '900500', units: 100, bucket: 'inProd', crd: D(12, 1), dates: { factoryCrd: D(12, 1) } }),   // same date, lower PO
  ];
  const r = sell(as, atsOf({ AB100: [1000, 1450] }));   // short 450
  const s = styleOf(r);
  assert.deepEqual(s.lots.map(l => l.po), ['900501', '900500', '900502', '900503']);
  assert.deepEqual(s.lots.map(l => [l.taken, l.available]), [[200, 0], [100, 0], [150, 150], [0, 500]]);
  assert.deepEqual([s.now, s.stockNet, s.ordersTaken, s.oversold, s.fromChina, s.arriving], [0, -450, 450, 0, 650, 0]);
  assert.equal(s.firstPull, D(11, 26), 'the first pull date of what is still available');
  assert.equal(s.firstArrival, null, 'the lot on the water is all taken');
  sumsHold(r);
  const over = sell(as.map(a => ({ ...a })), atsOf({ AB100: [0, 1300] }));
  const o = styleOf(over);
  assert.deepEqual([o.ordersTaken, o.oversold, o.fromChina + o.arriving], [1100, 200, 0]);
  assert.ok(o.flags.includes('oversold'));
  const iss = over.issues.find(i => i.code === 'oversold');
  assert.equal(iss.style, 'AB100');
  assert.match(iss.message, /customer orders \(1,300\) are 200 units more than the warehouse stock \(0\) and every open lot from China \(1,100 units\)/);
  assert.equal(over.totals.oversold, 200);
  assert.equal(over.counts.oversoldStyles, 1);
  sumsHold(over);
});

test('a style not on the ATS report: stock "not on report" (never 0), its lots in full, one summary line', () => {
  const r = sell([alloc({ style: 'AB200', styleRaw: 'AB200', units: 400 }), alloc({ style: 'AB201', styleRaw: 'AB201', bucket: 'pullFromChina', units: 10 })],
    atsOf({ AB100: [50, 0] }));
  const s = styleOf(r, 'AB200');
  assert.deepEqual([s.inWh, s.orders, s.stockNet, s.now, s.onReport], [null, null, null, null, false]);
  assert.equal(s.fromChina, 400);
  assert.ok(s.flags.includes('not-on-report') && s.flags.includes('orders-unknown'));
  assert.ok(s.lots[0].flags.includes('orders-unknown'));
  assert.ok(!styleOf(r, 'AB201').flags.includes('orders-unknown'), 'nothing available: no orders to take off');
  const ou = r.issues.filter(i => i.code === 'orders-unknown');
  assert.equal(ou.length, 1);
  assert.match(ou[0].message, /^1 style is not on the ATS report, so customer orders aren't taken off their lots from China \(400 units shown in full\): for example AB200\./);
  // A style on the ATS report with no lots: still a row, stock only.
  const st = styleOf(r, 'AB100');
  assert.deepEqual([st.now, st.fromChina, st.lots.length, st.total], [50, 0, 0, 50]);
  sumsHold(r);
});

test('the ATS report missing or failed: every style shows its lots in full, Available now unknown, atsMissing', () => {
  const as = [alloc({ units: 400 }), alloc({ style: 'AB300', styleRaw: 'AB300', units: 100, bucket: 'water', eta: D(11, 5) })];
  for (const ats of [null, { ok: false, error: 'ATS report: the report doesn\'t add up: rows 1, Grand Total 2.', styles: new Map() }]) {
    const r = sell(as.map(a => ({ ...a })), ats);
    assert.equal(r.atsMissing, true);
    assert.equal(r.reportDay, null);
    assert.equal(r.atsError, ats ? ats.error : '');
    assert.deepEqual(r.styles.map(s => [s.style, s.now, s.onReport, s.total]), [['AB100', null, null, 400], ['AB300', null, null, 100]]);
    assert.equal(r.totals.now, 0);
    assert.equal(r.totals.total, 500);
    const ou = r.issues.filter(i => i.code === 'orders-unknown');
    assert.equal(ou.length, 1);
    assert.match(ou[0].message, ats ? /^The ATS report couldn't be used/ : /^There's no ATS report/);
    sumsHold(r);
  }
});

test('a style whose stock on the report cannot be read is not 0: its lots in full', () => {
  const r = sell([alloc({ units: 300 })], atsOf({ AB100: [100, null] }));
  const s = styleOf(r);
  assert.deepEqual([s.inWh, s.orders, s.now, s.onReport, s.fromChina], [100, null, null, true, 300]);
  assert.ok(s.flags.includes('stock-unknown') && s.flags.includes('orders-unknown'));
  assert.match(r.issues.find(i => i.code === 'orders-unknown').message, /1 has a stock figure on it that can't be read/);
});

test('erp-only: styles on the ATS report and nowhere else are one summary line (UPC_MASTER counts as somewhere)', () => {
  const upc = { byStyle: new Map([['AB402', [{ upc: '000111222444', desc: 'From UPC_MASTER', category: 'Eye' }]]]) };
  const r = sell([alloc()], atsOf({ AB100: [0, 0], AB401: [10, 0], AB402: [5, 0], AB403: [0, 0] }), { upc });
  const e = r.issues.filter(i => i.code === 'erp-only');
  assert.equal(e.length, 1);
  assert.match(e[0].message, /^2 styles are on the ATS report but on no Master Sheet Live or Container Tracking line and not in UPC_MASTER \(1 with stock or orders\).*AB401, AB403/);
  assert.ok(styleOf(r, 'AB401').flags.includes('erp-only'));
  const u = styleOf(r, 'AB402');
  assert.deepEqual([u.desc, u.category, u.upcs], ['From UPC_MASTER', 'Eye', ['000111222444']]);
  assert.equal(styleOf(r, 'AB401').desc, 'ERP AB401', 'the report\'s description when nothing else has one');
});

// ---------- cases

test('cases: case pack from the lots (mixed when they differ); stock uses the latest Master Sheet Master', () => {
  const r = sell([
    alloc({ units: 100, master: 24, sheetRow: 10 }), alloc({ po: '900102', units: 48, master: 24, sheetRow: 12 }),
    alloc({ style: 'AB200', styleRaw: 'AB200', units: 100, master: 24, sheetRow: 20 }),
    alloc({ style: 'AB200', styleRaw: 'AB200', po: '900202', units: 100, master: 12, sheetRow: 30 }),
    alloc({ style: 'AB300', styleRaw: 'AB300', units: 100, master: null }),
    alloc({ style: 'AB400', styleRaw: 'AB400', bucket: 'landed', units: 10, master: 6, sheetRow: 5 }),
  ], atsOf({ AB100: [49, 0], AB200: [25, 0], AB300: [10, 0], AB400: [13, 1] }));
  const s1 = styleOf(r), s2 = styleOf(r, 'AB200'), s3 = styleOf(r, 'AB300'), s4 = styleOf(r, 'AB400');
  assert.equal(s1.casePack, 24);
  // Whole cases only (a part case is never offered as a case): 49 / 24 = 2 cases + 1 loose unit; 100 / 24 = 4 + 4.
  assert.deepEqual([s1.cases.now, s1.cases.fromChina, s1.cases.total], [2, 4 + 2, 8]);
  assert.deepEqual([s1.looseUnits.now, s1.looseUnits.fromChina, s1.looseUnits.total], [1, 4, 5]);
  assert.equal(s2.casePack, 'mixed');
  assert.ok(s2.flags.includes('mixed-case-pack'));
  assert.deepEqual([s2.cases.now, s2.looseUnits.now], [2, 1], 'stock: the latest Master Sheet line (row 30) has Master 12');
  assert.equal(s3.casePack, null);
  assert.deepEqual([s3.cases.now, s3.noCasePackUnits.now, s3.noCasePackUnits.fromChina, s3.looseUnits.total], [0, 10, 100, 0]);
  assert.deepEqual([s4.casePack, s4.cases.now, s4.looseUnits.now], [6, 2, 0], 'no lots: the latest Master Sheet Master');
  assert.equal(r.totals.cases.now, 2 + 2 + 0 + 2);
  assert.equal(r.totals.looseUnits.now, 1 + 1 + 0 + 0);
  assert.equal(r.totals.noCasePackUnits.now, 10);
  for (const l of r.lots) if (l.cases != null) assert.equal(l.cases * l.master + l.looseUnits, l.available, 'a lot is its whole cases and loose units');
});

// ---------- sums on a larger made-up mix

test('sums hold on a large random mix: lots, taken, available, oversold, totals', () => {
  let seed = 7;
  const rnd = k => (seed = (seed * 1103515245 + 12345) % 2147483648) % k;
  const buckets = ['noCrd', 'inProd', 'ready', 'booked', 'loaded', 'water', 'port', 'pullFromChina', 'landed', 'unmatched'];
  const as = [], stock = {};
  for (let i = 0; i < 600; i++) {
    const style = 'AB' + (100 + rnd(60)), bucket = buckets[rnd(buckets.length)];
    const day = rnd(4) ? D(9, 1) + rnd(200) : null;
    as.push(alloc({ style, styleRaw: style, po: String(900000 + rnd(500)), bucket, units: 1 + rnd(5000), master: rnd(5) ? 12 * (1 + rnd(4)) : null,
      etd: bucket === 'loaded' && rnd(2) ? day : null, eta: ['water', 'port'].includes(bucket) && rnd(5) ? day : null,
      stage: bucket === 'unmatched' && rnd(3) === 0 ? 'Landed' : 'Booked', crd: day,
      dates: { factoryCrd: day, shipBy: rnd(3) ? day : null, orderDate: rnd(2) ? D(6, 1) + rnd(100) : null } }));
  }
  for (let i = 0; i < 70; i++) if (rnd(4)) stock['AB' + (100 + i)] = [rnd(20000) - 1000, rnd(40000)];
  const r = sell(as, atsOf(stock));
  sumsHold(r);
  const lotUnits = r.lots.reduce((s, l) => s + l.units, 0);
  const counted = as.filter(a => SELL.IN_CHINA.includes(a.bucket) || SELL.SAILED.includes(a.bucket)).reduce((s, a) => s + a.units, 0);
  assert.equal(lotUnits, counted, 'every counted unit that is not sold is in a lot');
  assert.equal(r.totals.soldChina, as.filter(a => a.bucket === 'pullFromChina').reduce((s, a) => s + a.units, 0));
  let short = 0;
  for (const s of r.styles) if (s.stockNet != null && s.stockNet < 0) short -= s.stockNet;
  assert.equal(r.totals.ordersTaken + r.totals.oversold, short, 'every unit sold ahead of stock is taken off a lot or oversold');
  for (const s of r.styles) {
    // Taken lots are always the earliest: no lot with a later date is taken while an earlier one has units left.
    const ls = s.lots;
    for (let i = 1; i < ls.length; i++) if (ls[i].taken > 0) assert.equal(ls[i - 1].available, 0, `style ${s.style}`);
  }
});

// ---------- the stage engine's dates, end to end

// Master Sheet Live rows with Ship by on each PO block's first row (blocks of one or more lines, as the sheet has them).
const MH = ['Category', 'Sub', 'Product name', 'Style Number', 'Barcode', 'Master', 'Quantity', 'Po Number', 'Order Date', 'Ship by'];
let nextRow = 2;
function block(po, shipBy, lns) {
  return lns.map((l, i) => ({ ...l, po: l.po ?? po, row: nextRow++, blockFirst: i === 0, shipBy }));
}
function mrowsOf(lines) {
  const rows = [MH];
  // Dates as the sheet holds them: Excel serial numbers.
  for (const l of lines) rows[l.row - 1] = ['Lip', 'Gloss', 'Item', l.sty, null, l.master ?? 24, l.qty, l.blockFirst ? l.po : null, null, l.blockFirst ? l.shipBy + 25569 : null];
  for (let i = 0; i < rows.length; i++) if (!rows[i]) rows[i] = [];
  return rows;
}
function line(o) {
  const po = o.po, sty = o.sty ?? 'AB100';
  return { key: (po + '|' + sty).toUpperCase(), base: String(po).slice(0, 6), desc: 'Item', sup: 'Factory A', od: null, qty: 1000,
    price: 1, crd: null, ptype: 'LA', cust: '', cpo: '', src: 'open', status: 'open', bo: false, tracked: true, sty, ...o };
}
const blk = (sheet, po, rounds, lns, extra = {}) => ({ sheet, row: 50, section: rounds.length ? 'Round ' + rounds[0] : 'Planning A', rounds, po, poRaw: po,
  factory: 'Factory A', orderDate: null, shipBy: null, readyDate: null, needBy: '', ...extra,
  lines: lns.map((l, i) => ({ row: 51 + i, style: l.style ?? 'AB100', styleRaw: l.style ?? 'AB100', upc: null, upcRaw: '', master: null, qty: l.qty })) });
const ctOf = (blocks, rows = []) => {
  const rounds = { open: {}, complete: {} };
  for (const r of rows) for (const k of r.rounds) (rounds[r.sheet.toLowerCase()][k] = rounds[r.sheet.toLowerCase()][k] || []).push(r);
  return { ok: true, blocks, rounds, issues: [], counts: {} };
};

test("the stage engine carries every source's date: Container Tracking's Ready Date, tracker CRDs, Ship by, order date", () => {
  // Tracker: Factory CRD and CRD Expected per 6-digit PO + style.
  const T = new Map([['900601|AB100', { crd: D(12, 1), exp: D(11, 20), sup: 'Factory A' }], ['900602|AB100', { crd: null, exp: D(12, 15), sup: 'Factory A' }],
    ['900603|AB100', { crd: D(9, 1), exp: null, sup: 'Factory A' }], ['900604|AB100', { crd: D(12, 2), exp: null, sup: 'Factory A' }]]);
  const ls = [
    ...block('900601', D(11, 10), [line({ po: '900601', crd: D(12, 1), od: D(8, 1), qty: 1000 })]),
    ...block('900602', D(12, 20), [line({ po: '900602', crd: D(12, 15), od: D(8, 2), qty: 500 })]),
    // A backorder PO typed on Master Sheet Live: the PO Book uses its Ship by; the tracker's dates are the original PO's.
    ...block('900603-1', D(1, 5, 2027), [line({ po: '900603-1', key: '900603-1|AB100', crd: D(1, 5, 2027), od: D(8, 3), bo: true, qty: 300 })]),
    ...block('900604', D(12, 1), [line({ po: '900604', crd: D(12, 2), od: D(8, 4), qty: 200 })]),
  ];
  const ct = ctOf([
    blk('Loaded', '900601', [700], [{ qty: 1000 }], { readyDate: D(12, 18) }),
    blk('Booking', '900604', [], [{ qty: 200 }], { readyDate: D(12, 30), section: 'Planning A' }),   // planning: not booked yet
  ]);
  const R = ST.build({ po: { lines: ls, trk: { T, minPO: 900000 } }, mrows: mrowsOf(ls), ct, today: TODAY,
    sections: { planning: ['Planning A'] } });
  assert.equal(R.ok, true);
  assert.equal(R.balance.diff, 0);
  const a = po => R.allocs.find(x => x.po === po && x.sheetRow != null && x.bucket !== 'planning');
  assert.deepEqual(a('900601').dates, { ctReady: D(12, 18), ctRow: { sheet: 'Loaded', row: 50 }, poBook: null, factoryCrd: D(12, 1),
    crdExpected: D(11, 20), shipBy: D(11, 10), orderDate: D(8, 1), original: null });
  assert.equal(a('900601').bucket, 'loaded');
  assert.deepEqual([a('900602').dates.factoryCrd, a('900602').dates.crdExpected, a('900602').dates.shipBy, a('900602').dates.ctReady], [null, D(12, 15), D(12, 20), null]);
  const bo = a('900603-1').dates;
  assert.deepEqual([bo.shipBy, bo.factoryCrd, bo.crdExpected], [D(1, 5, 2027), null, null]);
  assert.deepEqual(bo.original, { factoryCrd: D(9, 1), crdExpected: null });
  const plan = a('900604');
  assert.equal(plan.bucket, 'inProd', 'on Container Tracking only in a planning section: the Master Sheet stage');
  assert.equal(plan.dates.ctReady, D(12, 30), "the planning section's Ready Date is the line's");

  const r = SELL.buildSell({ stages: R, ats: null, today: TODAY });
  const lot = po => r.lots.find(l => l.po === po);
  assert.deepEqual([lot('900601').readyDate, lot('900601').readyFrom], [D(12, 18), 'ctReady']);
  assert.ok(lot('900601').flags.includes('dates-disagree'), '11/10 to 12/18: more than 7 days');
  assert.deepEqual([lot('900602').readyDate, lot('900602').readyFrom, lot('900602').pullFrom, lot('900602').pullRule], [D(12, 15), 'crdExpected', D(12, 10), 'crd']);
  assert.ok(!lot('900602').flags.includes('dates-disagree'), '12/15 and 12/20: within 7 days');
  assert.deepEqual([lot('900603-1').readyDate, lot('900603-1').readyFrom], [D(1, 5, 2027), 'shipBy']);
  assert.ok(!lot('900603-1').flags.includes('dates-disagree'), "the original PO's 9/1 Factory CRD is not this backorder's");
  assert.deepEqual([lot('900604').readyDate, lot('900604').readyFrom], [D(12, 30), 'ctReady']);
  // Without Container Tracking's date the synced date is the PO Book's CRD.
  for (const l of r.lots) if (l.dates.ctReady == null) assert.equal(l.readyDate, l.crd, l.po);
  assert.equal(r.counts.readyNotCrd, 0);
});

test('end to end with a real-format ATS file: parseAts then buildSell', () => {
  const X = loadXLSX();
  const HEAD = ['Color Code', 'Style', 'Style Description', 'Warehouse', 'In WH', 'Total Customer Orders', 'ATS from in WH'];
  const rows = [['Inventory Report'], HEAD, ['XX', 'AB100', 'Item', 'WH 1', '600', '1000', '-400'], ['XX', 'AB100', 'Item', 'WH 2', 100, 0, '100'],
    ['Grand Total', null, null, null, 700, '', '']];
  const ats = A.parseAts(X, readBook(xlsxBytes({ Report: rows })), 'Inventory Report - On Hand_In Transit_In Production 10-29-2026 12-58 PM.xlsx', TODAY);
  assert.equal(ats.ok, true, ats.error);
  const r = SELL.buildSell({ stages: stagesOf([alloc({ units: 250, bucket: 'booked', etd: D(11, 30) }), alloc({ po: '900102', units: 250, bucket: 'inProd', crd: D(12, 31), dates: { factoryCrd: D(12, 31) } })]), ats, today: TODAY });
  const s = styleOf(r);
  assert.deepEqual([s.inWh, s.orders, s.stockNet, s.now, s.ordersTaken, s.fromChina, s.oversold], [700, 1000, -300, 0, 300, 200, 0]);
  assert.deepEqual(s.lots.map(l => [l.po, l.taken, l.available]), [['900101', 250, 0], ['900102', 50, 200]]);
  assert.equal(r.reportDay, D(10, 29));
  sumsHold(r);
});

// ---------- fixes after review (Oct 30)

test('ready date: a source more than a year ahead (a wrong year) is skipped for the next one, kept and still disagrees', () => {
  const r = sell([
    alloc({ po: '900801', bucket: 'booked', dates: { ctReady: D(10, 30, 2028), crdExpected: D(10, 23), shipBy: D(10, 23) } }),
    alloc({ po: '900802', bucket: 'booked', dates: { ctReady: D(10, 30, 2028) } }),   // the only date: used as it is
    alloc({ po: '900803', bucket: 'booked', dates: { ctReady: D(1, 5, 2014), factoryCrd: D(12, 1) } }),   // before 2015
  ], atsOf({ AB100: [0, 0] }));
  const lot = po => r.lots.find(l => l.po === po);
  assert.deepEqual([lot('900801').readyDate, lot('900801').readyFrom, lot('900801').pullFrom], [D(10, 23), 'crdExpected', D(10, 18)]);
  assert.deepEqual(lot('900801').readySources.map(s => [s.key, s.far]), [['ctReady', true], ['crdExpected', false], ['shipBy', false]]);
  assert.ok(lot('900801').flags.includes('dates-disagree'), 'the far date still makes the dates disagree');
  assert.ok(lot('900801').flags.includes('now'));
  assert.deepEqual([lot('900802').readyDate, lot('900802').readyFrom], [D(10, 30, 2028), 'ctReady']);
  assert.deepEqual([lot('900803').readyDate, lot('900803').readyFrom], [D(12, 1), 'factoryCrd']);
  const msg = r.issues.find(i => i.code === 'dates-disagree' && i.po === '900801').message;
  assert.match(msg, /Container Tracking's Ready Date 10\/30\/2028 is more than a year ahead or before 2015 \(a wrong year\?\), so it isn't used\. The ready date used is 10\/23\/2026 \(CRD Expected\)/);
  assert.equal(r.counts.readyFarSkipped, 2);
  sumsHold(r);
});

// A parseAts result with In Trans: { style: [inWh, orders, inTrans] }, and the In Trans column found.
const atsT = (map, reportDay = D(10, 29)) => {
  const a = atsOf(Object.fromEntries(Object.entries(map).map(([k, v]) => [k, [v[0], v[1]]])), reportDay);
  for (const [k, v] of Object.entries(map)) a.styles.get(k).inTrans = v[2] ?? null;
  return Object.assign(a, { columns: ['style', 'inWh', 'orders', 'inTrans'] });
};

test('erp-received: sailed lots with an ETA on or before the report day, beyond its In Trans, are listed (never taken off)', () => {
  const r = sell([
    alloc({ po: '900901', bucket: 'port', units: 1000, eta: D(8, 30) }),
    alloc({ po: '900902', bucket: 'water', units: 400, eta: D(11, 20) }),   // after the report day: not received
    alloc({ style: 'AB200', styleRaw: 'AB200', po: '900903', bucket: 'port', units: 600, eta: D(10, 20) }),
  ], atsT({ AB100: [1000, 0, 400], AB200: [0, 0, 600] }));
  const s = styleOf(r), s2 = styleOf(r, 'AB200');
  assert.deepEqual([s.now, s.arriving, s.total, s.erpReceived], [1000, 1400, 2400, 1000], 'the numbers stay as the contract says; the flag says why');
  assert.ok(s.flags.includes('erp-received'));
  assert.deepEqual(r.lots.filter(l => l.flags.includes('erp-received')).map(l => l.po), ['900901']);
  assert.ok(!s2.flags.includes('erp-received'), 'In Trans covers it: not received');
  const i = r.issues.filter(x => x.code === 'erp-received');
  assert.equal(i.length, 1);
  assert.match(i[0].message, /Style AB100: 1,000 units on the water or at port \(PO 900901, ETA 8\/30\/2026\) have an ETA on or before the ATS report's date \(10\/29\/2026\) but aren't In Trans on it \(In Trans 400\)/);
  assert.deepEqual([r.counts.erpReceivedStyles, r.counts.erpReceivedUnits, r.counts.erpReceivedAvailable], [1, 1000, 1000]);
  // Without the In Trans column nothing is said.
  const r2 = sell([alloc({ po: '900901', bucket: 'port', units: 1000, eta: D(8, 30) })], atsOf({ AB100: [1000, 0] }));
  assert.ok(!codes(r2).includes('erp-received') && !styleOf(r2).flags.includes('erp-received'));
  sumsHold(r);
});

test('erp-in-transit and orders-may-include-sold: an oversold style says what may cover it, and never "nothing is coming"', () => {
  const r = sell([
    alloc({ style: 'AB300', styleRaw: 'AB300', po: '901001', bucket: 'pullFromChina', units: 1104, ptype: 'China Pickup' }),
    alloc({ style: 'AB400', styleRaw: 'AB400', po: '901002', bucket: 'water', units: 1000, eta: D(11, 20) }),
  ], atsT({ AB200: [0, 1000, 1200], AB300: [0, 1104, null], AB400: [0, 3000, 1500] }));
  const s2 = styleOf(r, 'AB200'), s3 = styleOf(r, 'AB300'), s4 = styleOf(r, 'AB400');
  assert.deepEqual([s2.oversold, s2.erpInTransBeyond, s2.flags.includes('erp-in-transit')], [1000, 1200, true]);
  assert.deepEqual([s3.oversold, s3.soldInOrders, s3.flags.includes('orders-may-include-sold')], [1104, 1104, true]);
  assert.deepEqual([s4.oversold, s4.erpInTransBeyond, s4.ordersTaken], [2000, 500, 1000], 'In Trans beyond our 1,000 on the water');
  const ov = st => r.issues.find(i => i.code === 'oversold' && i.style === st).message;
  assert.match(ov('AB200'), /with no open lot from China\. Check before turning a buyer away: the ATS report shows 1,200 In Trans, 1,200 units more than we have on the water/);
  assert.match(ov('AB300'), /1,104 units of this style are sold to ship from China: if those orders are in the ERP's Total Customer Orders, they're taken off twice/);
  assert.ok(!/nothing is coming/i.test(r.issues.map(i => i.message).join(' ')));
  const sum = code => r.issues.filter(i => i.code === code);
  assert.equal(sum('erp-in-transit').length, 1);
  assert.match(sum('erp-in-transit')[0].message, /^2 styles \(2 of them oversold\): .*Up to 1,500 units of those orders/);
  assert.equal(sum('orders-may-include-sold').length, 1);
  assert.deepEqual([r.counts.erpInTransStyles, r.counts.erpInTransUnits, r.counts.soldInOrdersStyles, r.counts.soldInOrdersUnits], [2, 1500, 1, 1104]);
  // An oversold style with nothing that may cover it still says so plainly.
  const r2 = sell([], atsT({ AB500: [0, 50, null] }));
  assert.match(r2.issues.find(i => i.code === 'oversold').message, /with no open lot from China\. Nothing is available to sell\./);
  sumsHold(r);
});
