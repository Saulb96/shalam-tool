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
// Every lot is matched + received + in the ERP's transit (what the ERP's report already counts) + taken + available,
// none below 0; per style, what's left of the lots is what's available from China, arriving and taken by orders, and
// Arriving is the ERP's In Trans plus ours.
const KEEP = ['matched', 'received', 'inErpTransit', 'taken', 'available'];
function sumsHold(r) {
  for (const l of r.lots) {
    assert.equal(l.units, KEEP.reduce((t, k) => t + l[k], 0), `lot ${l.id}`);
    for (const k of KEEP) assert.ok(l[k] >= 0, `lot ${l.id} ${k}`);
    assert.ok(!(l.source === 'erp' && (l.matched || l.received || l.inErpTransit)), `lot ${l.id}: the ERP's own lots are never matched`);
  }
  for (const s of r.styles) {
    const left = s.lots.reduce((t, l) => t + l.taken + l.available, 0);
    assert.equal(left, s.fromChina + s.arriving + s.ordersTaken, `style ${s.style}`);
    assert.equal(s.total, (s.now ?? 0) + s.fromChina + s.arriving);
    assert.equal(s.arriving, s.arrivingErp + s.arrivingOurs, `style ${s.style}: arriving`);
    assert.equal(s.arrivingErp, s.lots.filter(l => l.source === 'erp').reduce((t, l) => t + l.available, 0));
  }
  const t = r.totals;
  assert.equal(t.total, t.now + t.fromChina + t.arriving);
  assert.equal(t.fromChina, r.lots.filter(l => l.inChina).reduce((s, l) => s + l.available, 0));
  assert.equal(t.arriving, r.lots.filter(l => !l.inChina).reduce((s, l) => s + l.available, 0));
  assert.equal(t.ordersTaken, r.lots.reduce((s, l) => s + l.taken, 0));
  assert.equal(t.arriving, t.arrivingErp + t.arrivingOurs);
  assert.ok(!r.issues.some(i => i.code === 'balance'), 'no balance issue');
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


// ---------- Arriving from the ERP's In Trans (Saul, Oct 30)

// A parseAts result with In Trans shipments: { style: [inWh, orders, [[qty, ETA, PO?], …], [backlog, …]] } (one report
// row each; an ETA of null is a part that can't be read; a shipment's own PO only when the report has a PO column), and
// the In Trans column found (the Import PO Backlog column too when any style gives one, even an empty list).
function atsE(map, reportDay = D(10, 29)) {
  const a = atsOf(Object.fromEntries(Object.entries(map).map(([k, v]) => [k, [v[0], v[1]]])), reportDay);
  for (const [k, v] of Object.entries(map)) {
    const ships = v[2] || [], s = a.styles.get(k), bl = v[3] || [];
    s.rows = [{ row: 3, warehouse: 'WH 1', inWh: v[0], orders: v[1], inTrans: ships.length ? ships.reduce((t, x) => t + x[0], 0) : null,
      inTransList: ships.map(x => x[0]), inTransBad: [], etaParts: ships.map(x => x[1] ?? null), eta: ships.map(x => x[1]).filter(d => d != null),
      backlog: bl.length ? bl.reduce((t, n) => t + n, 0) : null, backlogList: bl, poParts: ships.some(x => x[2]) ? ships.map(x => x[2] || '') : [] }];
    s.inTrans = s.rows[0].inTrans; s.backlog = s.rows[0].backlog;
  }
  const backlog = Object.values(map).some(v => v[3]);
  return Object.assign(a, { columns: ['style', 'inWh', 'orders', 'inTrans', 'eta', ...(backlog ? ['backlog'] : [])] });
}
const lotOf = (r, po) => r.lots.find(l => l.po === po);
const erpLots = (r, s = 'AB100') => r.lots.filter(l => l.style === s && l.source === 'erp');
// Every unit accounted for, per style, against the inputs worked out here: our lots hold the stage numbers' Incoming
// (every allocation in a counted stage that isn't sold), the ERP's lots its In Trans; what's sellable (from China +
// arriving + taken by orders) + matched + received + in the ERP's transit = Incoming + In Trans; In Trans = matched +
// landed, not received + taken off China + unplaced, each of its shipments placed in full.
// Then the checks the identities can't make (a unit counted twice or dropped still adds up), each against the inputs:
// received never more than In WH and only with an ETA on or before the report's day; landed by our Landed goods never
// more than those (not sold) dated from 60 days before the report's day up to today, landed by its shipments never more
// than those due by then; taken off China only from our lots still in China, never below the report's Import PO
// Backlog (sold goods count until they're loaded; a part below zero or not a number: no backlog); with a backlog, landed
// never more than the In Trans beyond what may come off China; In Trans left unplaced only when nothing more could come
// off China; and never In Trans landed or unplaced while our China goods beyond the backlog are still offered (offered
// twice; with no backlog, unplaced while any are).
const P = ['matched', 'landedOurs', 'landedDue', 'offChina', 'unplaced'];
function reconcile(r, allocs, ats) {
  sumsHold(r);
  const lotB = [...SELL.IN_CHINA, ...SELL.SAILED];
  const sold = a => a.ptype === 'China Pickup' || (a.flags || []).includes('pull-section');
  const sumU = (a, f = x => x.units) => a.reduce((t, x) => t + f(x), 0);
  const rep = r.reportDay;
  const backlogCol = !!(ats && Array.isArray(ats.columns) && ats.columns.includes('backlog'));
  for (const s of r.styles) {
    const mine = allocs.filter(a => a.style === s.style);
    const incoming = sumU(mine.filter(a => lotB.includes(a.bucket) && !sold(a)));
    const e = ats && ats.styles && ats.styles.get(s.style);
    const ships = s.arrivingFrom === 'erp' && e ? e.rows.flatMap(x => x.inTransList.map((n, i) => ({ units: n, eta: (x.etaParts || [])[i] ?? null }))) : [];
    const inTrans = sumU(ships);
    assert.equal(s.ourIncoming, incoming, `${s.style}: our lots hold Incoming`);
    assert.equal(s.lots.filter(l => l.source === 'ours').reduce((t, l) => t + l.units, 0), incoming, `${s.style}: our lots`);
    assert.equal(s.lots.filter(l => l.source === 'erp').reduce((t, l) => t + l.units, 0), inTrans, `${s.style}: the ERP's lots hold its In Trans`);
    assert.equal(s.fromChina + s.arriving + s.ordersTaken + s.matchedOurs + s.receivedDropped + s.inErpTransitTaken, incoming + inTrans, `${s.style}: Incoming + In Trans`);
    if (s.arrivingFrom === 'erp') {
      assert.equal(s.erpInTrans, s.matchedOurs + s.landedNotReceived + s.inErpTransitTaken + s.erpUnplaced, `${s.style}: In Trans = matched + landed + taken off China + unplaced`);
      assert.equal(s.landedNotReceived, s.landedOurs + s.landedDue, `${s.style}: landed`);
      assert.equal(s.inErpTransitTaken, s.takenOnVessel + s.takenSamePo + s.takenOther, `${s.style}: taken by tier`);
      const eL = s.lots.filter(l => l.source === 'erp'), placed = k => sumU(eL, l => l.place[k]);
      for (const l of eL) assert.equal(P.reduce((t, k) => t + l.place[k], 0), l.units, `${l.id}: placed in full`);
      assert.deepEqual([placed('matched'), placed('landedOurs'), placed('landedDue'), placed('offChina'), placed('unplaced')],
        [s.matchedOurs, s.landedOurs, s.landedDue, s.inErpTransitTaken, s.erpUnplaced], `${s.style}: the shipments' places`);
      // The checks the identities can't make.
      if (e && typeof e.inWh === 'number') assert.ok(s.receivedDropped <= Math.max(0, e.inWh), `${s.style}: received ${s.receivedDropped} > In WH ${e.inWh}`);
      const day = a => a.eta ?? a.etd ?? null;
      const recent = rep == null ? 0 : sumU(mine.filter(a => a.bucket === 'landed' && !sold(a) && day(a) != null && day(a) >= rep - 60 && day(a) <= Math.max(rep, TODAY)));
      assert.ok(s.landedOurs <= recent, `${s.style}: landed ${s.landedOurs} beyond our recent Landed goods ${recent}`);
      // A shipment due after the report's day is our Landed goods' only within 14 days of the date of one of them.
      for (const l of eL) if (l.place.landedOurs && !(l.eta != null && rep != null && l.eta <= rep)) {
        assert.ok(l.eta != null && mine.some(a => a.bucket === 'landed' && !sold(a) && day(a) != null && day(a) >= rep - 60 && day(a) <= Math.max(rep, TODAY) && l.eta <= day(a) + SELL.LANDED_LATE_DAYS), `${l.id}: landed, due ${l.eta}`);
      }
      assert.equal(s.landedLater, sumU(eL.filter(l => !(l.eta != null && rep != null && l.eta <= rep)), l => l.place.landedOurs), `${s.style}: landed later`);
      assert.ok(s.landedDue <= sumU(ships.filter(x => x.eta != null && rep != null && x.eta <= rep)), `${s.style}: landed beyond the shipments due by the report's day`);
      const china = sumU(mine.filter(a => SELL.IN_CHINA.includes(a.bucket) && !sold(a)));
      const soldChina = sumU(mine.filter(a => ((sold(a) && SELL.IN_CHINA.includes(a.bucket)) || (a.bucket === 'pullFromChina' && !['On the water', 'At port', 'Landed', 'Shipped'].includes(a.stage))) &&
        a.bucket !== 'loaded' && a.stage !== 'Loaded'));
      const parts = backlogCol && e ? e.rows.flatMap(x => x.backlogList || []) : [];
      const backlog = !backlogCol || !e || parts.some(n => !(n >= 0)) || e.rows.some(x => (x.backlogBad || []).length) ? null : sumU(parts, n => n);
      const room = backlog == null ? Infinity : Math.max(0, china + soldChina - backlog), held = china - s.inErpTransitTaken;
      assert.ok(s.inErpTransitTaken <= room, `${s.style}: ${s.inErpTransitTaken} taken off China, below the backlog`);
      if (s.erpUnplaced > 0) assert.ok(held === 0 || s.inErpTransitTaken === room, `${s.style}: unplaced while more could come off China`);
      if (backlog != null) assert.ok(s.landedNotReceived <= Math.max(0, inTrans - s.matchedOurs - Math.min(room, china)), `${s.style}: landed ${s.landedNotReceived} while the backlog says more has left China`);
      if (s.erpInTrans < s.ourSailed) assert.equal(s.landedNotReceived + s.erpUnplaced, 0, `${s.style}: the ERP has less: nothing landed or unplaced`);
      assert.equal(backlog == null ? Math.min(s.erpUnplaced, held) : Math.min(s.landedNotReceived + s.erpUnplaced, held, Math.max(0, held + soldChina - backlog)), 0, `${s.style}: offered twice`);
      assert.deepEqual([s.offeredTwice, s.unplacedInChina, s.erpBacklog], [0, Math.min(s.erpUnplaced, held), backlog], s.style);
      assert.equal(s.takenDueBy, sumU(eL.filter(l => l.eta != null && rep != null && l.eta <= rep), l => l.place.offChina), `${s.style}: taken off China on shipments due by the report's day`);
      assert.equal(s.flags.includes('erp-backlog-unknown'), backlogCol && !!e && backlog == null, `${s.style}: backlog unknown`);
      assert.equal(s.flags.includes('erp-transit-unplaced'), s.erpUnplaced > 0, s.style);
      assert.equal(r.issues.filter(i => i.code === 'erp-transit-unplaced' && i.style === s.style).length, s.erpUnplaced > 0 ? 1 : 0, `${s.style}: one Problems line`);
    } else assert.deepEqual([s.erpInTrans, s.matchedOurs, s.receivedDropped, s.inErpTransitTaken, s.landedNotReceived, s.erpUnplaced, s.keptExtra, s.arrivingErp, s.offeredTwice],
      [null, 0, 0, 0, 0, 0, 0, 0, 0], s.style);
  }
  for (const l of r.lots) {
    if (l.received) assert.ok(l.eta != null && rep != null && l.eta <= rep, `${l.id}: received with an ETA after the report's day`);
    if (l.inErpTransit) assert.ok(l.source === 'ours' && l.inChina && SELL.TIERS.includes(l.erpTier) && (l.erpTier !== 'vessel' || (['booked', 'loaded'].includes(l.bucket) && l.etd != null)), `${l.id}: in the ERP's transit`);
  }
  assert.equal(r.totals.offeredTwice, 0);
  assert.ok(!r.issues.some(i => ['maybe-twice', 'taken-in-backlog', 'erp-in-trans-unmatched'].includes(i.code)), 'the old cross-checks are gone');
}

test('ERP In Trans ≥ our sailed units: ours are matched (not lots any more); the excess due after the report day comes off Booked / Loaded lots by ETD first', () => {
  const as = [
    alloc({ po: '900901', bucket: 'water', units: 1000, eta: D(11, 10) }),
    alloc({ po: '900902', bucket: 'port', units: 500, eta: D(10, 25) }),
    alloc({ po: '900903', bucket: 'booked', units: 1200, etd: D(11, 20) }),
    alloc({ po: '900904', bucket: 'loaded', units: 800, etd: D(11, 5) }),
    alloc({ po: '900905', bucket: 'booked', units: 600, crd: D(11, 1), dates: { factoryCrd: D(11, 1) } }),   // no ETD: never taken
    alloc({ po: '900906', bucket: 'inProd', units: 2000, crd: D(12, 1), dates: { factoryCrd: D(12, 1) } }),
  ];
  const ats = atsE({ AB100: [5000, 5000, [[1500, D(11, 12)], [900, D(11, 28)], [400, D(12, 5)]]] });
  const r = sell(as, ats), s = styleOf(r);
  // The ERP's three shipments are the Arriving lots, dated by their ETA.
  const e = erpLots(r);
  assert.deepEqual(e.map(l => [l.units, l.eta, l.availableDate, l.status, l.inChina, l.arrivalRule]),
    [[1500, D(11, 12), D(11, 12), 'In transit (ERP)', false, 'eta'], [900, D(11, 28), D(11, 28), 'In transit (ERP)', false, 'eta'], [400, D(12, 5), D(12, 5), 'In transit (ERP)', false, 'eta']]);
  assert.equal(SELL.ruleWords(e[0]), 'ETA 11/12/2026');
  // Our sailed lots: matched, kept in the panel, nothing available.
  for (const po of ['900901', '900902']) {
    const l = lotOf(r, po);
    assert.deepEqual([l.matched, l.available, l.taken], [l.units, 0, 0], po);
    assert.ok(l.flags.includes('matched-erp') && s.lots.includes(l), po);
  }
  // 2,800 − 1,500 = 1,300 more In Trans, due after the report's day and nothing Landed of ours: off the Loaded lot
  // (ETD 11/5) first, then the Booked one (11/20).
  assert.deepEqual([lotOf(r, '900904').inErpTransit, lotOf(r, '900904').available, lotOf(r, '900904').erpTier], [800, 0, 'vessel']);
  assert.deepEqual([lotOf(r, '900903').inErpTransit, lotOf(r, '900903').available, lotOf(r, '900903').erpTier], [500, 700, 'vessel']);
  assert.deepEqual([s.takenOnVessel, s.takenSamePo, s.takenOther, s.erpUnplaced, s.landedNotReceived], [1300, 0, 0, 0, 0]);
  assert.deepEqual(erpLots(r).map(l => [l.place.matched, l.place.offChina]), [[1500, 0], [0, 900], [0, 400]], 'ours on the shipment of their sum');
  assert.ok(lotOf(r, '900904').flags.includes('in-erp-transit') && lotOf(r, '900903').flags.includes('in-erp-transit'));
  assert.deepEqual([lotOf(r, '900905').inErpTransit, lotOf(r, '900905').available], [0, 600], 'Booked with no ETD: not on a vessel yet');
  assert.deepEqual([s.arrivingFrom, s.erpInTrans, s.erpShipments, s.ourSailed, s.matchedOurs, s.inErpTransitTaken, s.landedNotReceived, s.receivedDropped, s.keptExtra],
    ['erp', 2800, 3, 1500, 1500, 1300, 0, 0, 0]);
  assert.deepEqual([s.arriving, s.arrivingErp, s.arrivingOurs, s.fromChina, s.now, s.total], [2800, 2800, 0, 3300, 0, 6100]);
  assert.equal(s.firstArrival, D(11, 12));
  assert.ok(s.flags.includes('in-erp-transit') && !s.flags.includes('landed-not-received'));
  assert.ok(!codes(r).includes('erp-received') && !codes(r).includes('erp-in-transit') && !codes(r).includes('balance'));
  assert.deepEqual([r.totals.arrivingErp, r.totals.arrivingOurs, r.totals.inErpTransitTaken, r.totals.matchedOurs, r.totals.erpInTrans, r.totals.ourIncoming],
    [2800, 0, 1300, 1500, 2800, 6100]);
  assert.deepEqual([r.counts.erpStyles, r.counts.erpLots, r.counts.erpMoreStyles, r.counts.matchedLots, r.counts.inErpTransitLots], [1, 3, 1, 2, 2]);
  reconcile(r, as, ats);
});

test('step 1: In Trans our sheets have landed in the last 60 days is landed, not received yet; nothing comes off China', () => {
  // A landed PO: the ERP's 5,040 In Trans due 10/22 is our Landed PO 900960 (ETA 10/22); the Booked 10,008 (ETD
  // 11/11) is its Import PO Backlog. Until Oct 30 the rule took the 5,040 off the Booked lot.
  const as = [
    alloc({ style: 'AB500', styleRaw: 'AB500', po: '900960', bucket: 'landed', stage: 'Landed', units: 5040, eta: D(10, 22) }),
    alloc({ style: 'AB500', styleRaw: 'AB500', po: '900961', bucket: 'booked', stage: 'Booked', units: 10008, etd: D(11, 11) }),
    // Five landed lots: five Landed lots (ETA 10/24) are the ERP's five shipments; the Loaded 1,344 and 480 sold (China
    // Pickup) are its backlog.
    ...[1248, 1440, 240, 336, 240].map((u, i) => alloc({ style: 'AB501', styleRaw: 'AB501', po: '90097' + i, bucket: 'landed', stage: 'Landed', units: u, eta: D(10, 24) })),
    alloc({ style: 'AB501', styleRaw: 'AB501', po: '900979', bucket: 'loaded', stage: 'Loaded', units: 1344, etd: D(10, 30) }),
    alloc({ style: 'AB501', styleRaw: 'AB501', po: '900978', bucket: 'pullFromChina', stage: 'Booked', ptype: 'China Pickup', units: 480 }),
  ];
  const ats = atsE({ AB500: [0, 768, [[5040, D(10, 22)]], [10008]],
    AB501: [0, 5328, [[240, D(10, 24)], [1440, D(10, 24)], [240, D(10, 24)], [336, D(10, 24)], [1248, D(10, 24)]], [480, 1344]] });
  const r = sell(as, ats), s = styleOf(r, 'AB500');
  assert.deepEqual([s.landedOurs, s.landedDue, s.landedNotReceived, s.inErpTransitTaken, s.erpUnplaced, s.erpBacklog, s.takeable], [5040, 0, 5040, 0, 0, 10008, 0]);
  assert.deepEqual(s.landedLots, [{ po: '900960', units: 5040, day: D(10, 22), dayFrom: 'eta' }]);
  assert.deepEqual([lotOf(r, '900961').inErpTransit, lotOf(r, '900961').available], [0, 10008], 'the Booked lot stays available from China');
  const [e] = erpLots(r, 'AB500');
  assert.deepEqual([e.place.landedOurs, e.taken, e.available], [5040, 768, 4272]);
  assert.ok(e.flags.includes('landed-not-received') && s.flags.includes('landed-not-received') && !s.flags.includes('in-erp-transit'));
  assert.deepEqual([s.fromChina, s.arriving, s.total, s.oversold], [10008, 4272, 14280, 0]);
  const s1 = styleOf(r, 'AB501');
  assert.deepEqual([s1.landedOurs, s1.inErpTransitTaken, s1.erpUnplaced, s1.ordersTaken, s1.oversold, s1.fromChina, s1.arriving], [3504, 0, 0, 4848, 480, 0, 0]);
  assert.deepEqual(erpLots(r, 'AB501').map(l => l.place.landedOurs), [240, 1440, 240, 336, 1248], 'each Landed lot on the shipment of its quantity');
  assert.deepEqual([lotOf(r, '900979').inErpTransit, lotOf(r, '900979').taken], [0, 1344]);
  assert.deepEqual([r.counts.landedOursStyles, r.counts.landedNotReceivedStyles, r.totals.landedOurs, r.totals.landedLater], [2, 2, 8544, 0]);
  assert.ok(!codes(r).includes('erp-transit-unplaced') && !codes(r).includes('balance'));
  reconcile(r, as, ats);
  // The window: Landed goods dated (ETA, else ETD) from 60 days before the report's day (8/30) up to today (10/30), the
  // latest first, never sold goods; the rest of the shipment due 10/25 is due by the report's day all the same.
  const bs = [
    alloc({ style: 'AB510', styleRaw: 'AB510', po: '900981', bucket: 'landed', stage: 'Landed', units: 1000, eta: D(8, 29) }),   // 61 days: no
    alloc({ style: 'AB510', styleRaw: 'AB510', po: '900982', bucket: 'landed', stage: 'Landed', units: 400, eta: D(8, 30) }),    // 60 days: yes
    alloc({ style: 'AB510', styleRaw: 'AB510', po: '900983', bucket: 'landed', stage: 'Landed', units: 300, etd: D(10, 1) }),    // ETD only: yes
    alloc({ style: 'AB510', styleRaw: 'AB510', po: '900984', bucket: 'landed', stage: 'Landed', units: 200, eta: D(10, 30) }),   // after the report's day, by today: yes
    alloc({ style: 'AB510', styleRaw: 'AB510', po: '900985', bucket: 'landed', stage: 'Landed', units: 250, eta: D(10, 31) }),   // after today: no
    alloc({ style: 'AB510', styleRaw: 'AB510', po: '900986', bucket: 'landed', stage: 'Landed', units: 500, eta: D(10, 20), ptype: 'China Pickup' }),   // sold: no
    alloc({ style: 'AB510', styleRaw: 'AB510', po: '900987', bucket: 'inProd', units: 3000, crd: D(12, 1), dates: { factoryCrd: D(12, 1) } }),
    // A shipment due after the report's day is our Landed goods' only within 14 days of their date (10/27): 11/10 is,
    // 11/11 isn't (it has left China: off the lot in production, whose 2,000 the backlog has only 1,700 of).
    alloc({ style: 'AB511', styleRaw: 'AB511', po: '900988', bucket: 'landed', stage: 'Landed', units: 500, eta: D(10, 27) }),
    alloc({ style: 'AB511', styleRaw: 'AB511', po: '900989', bucket: 'inProd', units: 2000, crd: D(12, 1), dates: { factoryCrd: D(12, 1) } }),
    // A landed lot received long ago: the ERP's 20,016 due 11/11 are our 11,136 on the water and our Loaded 8,880; our 19,920 Landed on
    // 9/29 are long received (In WH). Its backlog (96) lets 8,784 come off the Loaded lot; 96 stay unplaced, in China.
    alloc({ style: 'AB512', styleRaw: 'AB512', po: '900990', bucket: 'landed', stage: 'Landed', units: 19920, eta: D(9, 29) }),
    alloc({ style: 'AB512', styleRaw: 'AB512', po: '900991', bucket: 'water', stage: 'On the water', units: 11136, eta: D(11, 11) }),
    alloc({ style: 'AB512', styleRaw: 'AB512', po: '900992', bucket: 'loaded', stage: 'Loaded', units: 8880, etd: D(10, 31) }),
  ];
  const ats2 = atsE({ AB510: [0, 0, [[1500, D(10, 25)]], [3000]], AB511: [0, 0, [[100, D(10, 29)], [300, D(11, 10)], [300, D(11, 11)]], [1700]],
    AB512: [22416, 2064, [[20016, D(11, 11)]], [96]] });
  const r2 = sell(bs, ats2), t = styleOf(r2, 'AB510');
  assert.deepEqual([t.landedOurs, t.landedLater, t.landedDue, t.inErpTransitTaken, t.erpUnplaced, t.fromChina, t.arriving], [900, 0, 600, 0, 0, 3000, 1500]);
  assert.deepEqual(t.landedLots.map(x => [x.po, x.units, x.dayFrom]), [['900984', 200, 'eta'], ['900983', 300, 'etd'], ['900982', 400, 'eta']]);
  const t1 = styleOf(r2, 'AB511');
  assert.deepEqual([t1.landedOurs, t1.landedLater, t1.landedDue, t1.takenOther, t1.erpUnplaced, t1.fromChina], [400, 300, 0, 300, 0, 1700]);
  assert.deepEqual(erpLots(r2, 'AB511').map(l => [l.units, l.place.landedOurs, l.place.offChina]), [[100, 100, 0], [300, 300, 0], [300, 0, 300]]);
  const t2 = styleOf(r2, 'AB512');
  assert.deepEqual([t2.matchedOurs, t2.landedOurs, t2.takenOnVessel, t2.erpUnplaced, t2.unplacedInChina, t2.fromChina, t2.arriving, t2.now], [11136, 0, 8784, 96, 96, 96, 20016, 20352]);
  assert.deepEqual([r2.counts.landedLaterStyles, r2.totals.landedLater], [1, 300]);
  reconcile(r2, bs, ats2);
});

test('step 1: shipments due on or before the report day are landed, not received, even with no Landed goods of ours; ours on the water take the closest ETA', () => {
  const as = [
    alloc({ style: 'AB200', styleRaw: 'AB200', po: '900911', bucket: 'water', stage: 'On the water', units: 300, eta: D(11, 1) }),
    alloc({ style: 'AB200', styleRaw: 'AB200', po: '900912', bucket: 'loaded', stage: 'Loaded', units: 200, etd: D(11, 3) }),
    alloc({ style: 'AB200', styleRaw: 'AB200', po: '900913', bucket: 'inProd', units: 1000, crd: D(12, 1), dates: { factoryCrd: D(12, 1) } }),
  ];
  // 1,500 In Trans: our 300 on the water (ETA 11/1) are on the shipment due 10/25, the closest; its other 300 were due
  // by the report's day (10/29): landed, not received. The 900 due 11/20 have left China: 200 off the Loaded lot, 700 off
  // the one in production.
  const ats = atsE({ AB200: [100, 900, [[600, D(10, 25)], [900, D(11, 20)]]] });
  const r = sell(as, ats), s = styleOf(r, 'AB200');
  assert.deepEqual([s.erpInTrans, s.matchedOurs, s.landedOurs, s.landedDue, s.takenOnVessel, s.takenSamePo, s.takenOther, s.erpUnplaced], [1500, 300, 0, 300, 200, 0, 700, 0]);
  const [e1, e2] = erpLots(r, 'AB200');
  assert.deepEqual([e1.place, e2.place], [{ matched: 300, landedOurs: 0, landedDue: 300, offChina: 0, unplaced: 0 }, { matched: 0, landedOurs: 0, landedDue: 0, offChina: 900, unplaced: 0 }]);
  assert.ok(e1.flags.includes('erp-ours-sailed') && e1.flags.includes('landed-not-received') && e2.flags.includes('erp-off-china'));
  assert.deepEqual([lotOf(r, '900912').inErpTransit, lotOf(r, '900912').erpTier, lotOf(r, '900913').inErpTransit, lotOf(r, '900913').erpTier, lotOf(r, '900913').available],
    [200, 'vessel', 700, 'other', 300]);
  // Customer orders beyond stock (800): the ERP's 10/25 shipment, then its 11/20 one (before the lot in production's 11/26).
  assert.deepEqual([e1.taken, e1.available, e2.taken, e2.available], [600, 0, 200, 700]);
  assert.deepEqual([s.now, s.ordersTaken, s.oversold, s.arriving, s.fromChina, s.total, s.firstArrival], [0, 800, 0, 700, 300, 1000, D(11, 20)]);
  assert.deepEqual([r.counts.landedDueStyles, r.counts.inErpTransitTierLots.vessel, r.counts.inErpTransitTierLots.other], [1, 1, 1]);
  reconcile(r, as, ats);
  // Only In Trans on the report, none of ours: due by its day, landed; due later or with no ETA, unplaced (Arriving on
  // the ERP's word, one Problems line each).
  const ats2 = atsE({ AB201: [0, 0], AB202: [0, 0, [[240, D(11, 9)]]], AB203: [0, 0, [[120, D(10, 20)], [60, null]]] });
  const as2 = [alloc({ style: 'AB201', styleRaw: 'AB201', units: 50 })];
  const r2 = sell(as2, ats2), s2 = styleOf(r2, 'AB202'), s3 = styleOf(r2, 'AB203');
  assert.deepEqual([s2.arriving, s2.landedNotReceived, s2.erpUnplaced, s2.unplacedInChina, s2.firstArrival], [240, 0, 240, 0, D(11, 9)]);
  assert.deepEqual([s3.arriving, s3.landedDue, s3.erpUnplaced], [180, 120, 60], 'no ETA: not called landed');
  const un = r2.issues.filter(i => i.code === 'erp-transit-unplaced');
  assert.deepEqual(un.map(i => [i.style, i.units, i.inChina]), [['AB202', 240, 0], ['AB203', 60, 0]]);
  assert.match(un[0].message, /^Style AB202: the ERP counts 240 units in transit that our sheets don't have on the water, at port, landed or still to sell in China\. They're offered as Arriving on the ERP's word/);
  reconcile(r2, as2, ats2);
});

test('ERP In Trans below our sailed units: by ETA, received on or before the report day (dropped), the rest kept as ours', () => {
  const as = [
    alloc({ style: 'AB300', styleRaw: 'AB300', po: '900931', bucket: 'port', units: 600, eta: D(10, 20) }),
    alloc({ style: 'AB300', styleRaw: 'AB300', po: '900932', bucket: 'water', units: 400, eta: D(10, 27) }),
    alloc({ style: 'AB300', styleRaw: 'AB300', po: '900933', bucket: 'water', units: 1000, eta: D(11, 15) }),
    alloc({ style: 'AB300', styleRaw: 'AB300', po: '900934', bucket: 'water', units: 500, eta: null }),
    alloc({ style: 'AB300', styleRaw: 'AB300', po: '900935', bucket: 'booked', units: 700, etd: D(11, 25) }),
    // The report's day itself counts as received; the day after doesn't. No In Trans at all on the report.
    alloc({ style: 'AB301', styleRaw: 'AB301', po: '900936', bucket: 'port', units: 100, eta: D(10, 29) }),
    alloc({ style: 'AB301', styleRaw: 'AB301', po: '900937', bucket: 'water', units: 200, eta: D(10, 30) }),
  ];
  const ats = atsE({ AB300: [2000, 0, [[1200, D(11, 15)]]], AB301: [500, 0] });
  const r = sell(as, ats), s = styleOf(r, 'AB300');
  // Ours 2,500 sailed, the ERP 1,200 In Trans: 1,300 to explain, earliest ETA first.
  assert.deepEqual(['900931', '900932', '900933', '900934'].map(po => { const l = lotOf(r, po); return [l.received, l.notOnErp, l.matched, l.available]; }),
    [[600, 0, 0, 0], [400, 0, 0, 0], [0, 300, 700, 300], [0, 0, 500, 0]]);
  assert.ok(lotOf(r, '900931').flags.includes('received') && lotOf(r, '900933').flags.includes('not-on-erp-in-trans') && lotOf(r, '900933').flags.includes('matched-erp'));
  assert.ok(!lotOf(r, '900934').flags.includes('no-date'), 'a lot the ERP counts in full needs no date');
  assert.ok(!codes(r).includes('no-date'));
  assert.deepEqual([lotOf(r, '900935').inErpTransit, lotOf(r, '900935').available], [0, 700], 'the ERP has less: nothing comes off China');
  assert.deepEqual([s.erpInTrans, s.ourSailed, s.matchedOurs, s.receivedDropped, s.keptExtra, s.inErpTransitTaken, s.landedNotReceived], [1200, 2500, 1200, 1000, 300, 0, 0]);
  assert.deepEqual([s.arriving, s.arrivingErp, s.arrivingOurs, s.fromChina, s.now, s.total], [1500, 1200, 300, 700, 2000, 4200]);
  assert.ok(s.flags.includes('received') && s.flags.includes('not-on-erp-in-trans'));
  assert.equal(s.firstArrival, D(11, 15));
  const s1 = styleOf(r, 'AB301');
  assert.deepEqual([lotOf(r, '900936').received, lotOf(r, '900937').notOnErp, s1.erpInTrans, s1.arriving, s1.arrivingOurs, s1.arrivingErp, s1.arrivingFrom],
    [100, 200, 0, 200, 200, 0, 'erp']);
  assert.deepEqual([r.counts.erpLessStyles, r.counts.receivedStyles, r.counts.receivedLots, r.counts.keptExtraStyles, r.counts.keptExtraLots], [2, 2, 3, 2, 2]);
  assert.deepEqual([r.totals.receivedDropped, r.totals.keptExtra, r.totals.arrivingOurs], [1100, 500, 500]);
  reconcile(r, as, ats);
  // With the report's day from the page (the file's saved day), the boundary moves with it.
  const r2 = sell(as.map(a => ({ ...a })), atsE({ AB300: [2000, 0, [[1200, D(11, 15)]]], AB301: [500, 0] }, null), { reportDay: D(10, 30) });
  assert.deepEqual([styleOf(r2, 'AB301').receivedDropped, styleOf(r2, 'AB301').keptExtra], [300, 0]);
  // No report day at all: nothing can be called received, so the difference is all kept (and said).
  const r3 = sell(as.map(a => ({ ...a })), atsE({ AB301: [500, 0] }, null));
  assert.deepEqual([styleOf(r3, 'AB301').receivedDropped, styleOf(r3, 'AB301').keptExtra], [0, 300]);
  assert.ok(styleOf(r3, 'AB301').flags.includes('no-report-day'));
  const nd = r3.issues.filter(i => i.code === 'no-report-day');
  assert.equal(nd.length, 1);
  assert.match(nd[0].message, /^1 style \(0 of them oversold\): the ATS report's date isn't known, so none of the 300 units on the water or at port beyond its In Trans is called received/);
});

test('ERP In Trans below ours: a shipment and a lot of the same quantity are matched first; received never more than In WH', () => {
  // Three lots due 10/28 and two ERP shipments of 1,800 and 1,000: those are lots B and A (same quantities), so lot C
  // (1,200) is the one not on its In Trans. Its ETA passed, but In WH is 0: it can't be in the warehouse number, so it's
  // kept as Arriving from our data, never dropped.
  const as = [
    alloc({ style: 'AB320', styleRaw: 'AB320', po: '900941A', bucket: 'water', units: 1000, eta: D(10, 28) }),
    alloc({ style: 'AB320', styleRaw: 'AB320', po: '900941B', bucket: 'water', units: 1800, eta: D(10, 28) }),
    alloc({ style: 'AB320', styleRaw: 'AB320', po: '900941C', bucket: 'water', units: 1200, eta: D(10, 28) }),
  ];
  const ats = atsE({ AB320: [0, 5000, [[1800, D(10, 28)], [1000, D(10, 28)]]] });
  const r = sell(as, ats), s = styleOf(r, 'AB320');
  assert.deepEqual(['900941A', '900941B', '900941C'].map(po => { const l = lotOf(r, po); return [l.matched, l.received, l.notOnErp]; }),
    [[1000, 0, 0], [1800, 0, 0], [0, 0, 1200]]);
  assert.deepEqual([s.matchedOurs, s.receivedDropped, s.keptExtra, s.receivedCapped, s.arriving, s.arrivingOurs, s.ordersTaken, s.oversold],
    [2800, 0, 1200, 1200, 0, 0, 4000, 1000]);
  assert.ok(s.flags.includes('received-over-in-wh') && s.flags.includes('not-on-erp-in-trans') && !s.flags.includes('received'));
  assert.equal(r.counts.sameQtyMatchedLots, 2);
  const cap = r.issues.find(i => i.code === 'received-over-in-wh');
  assert.match(cap.message, /^1 style \(1 of them oversold\): 1,200 units on the water or at port have an ETA on or before the ATS report's date and aren't on its In Trans, but the style's In WH is less/);
  reconcile(r, as, ats);
  // In WH 700: 700 of lot C received (dropped), 500 kept.
  const ats2 = atsE({ AB320: [700, 5000, [[1800, D(10, 28)], [1000, D(10, 28)]]] });
  const r2 = sell(as.map(a => ({ ...a, flags: [] })), ats2), s2 = styleOf(r2, 'AB320');
  assert.deepEqual([s2.receivedDropped, s2.keptExtra, s2.receivedCapped], [700, 500, 500]);
  reconcile(r2, as, ats2);
  // Same quantity, different ETAs: the ERP's 700 due 10/31 is our 700 at port (due 10/20), not received; our 900 due
  // 11/20 is the one not on its In Trans yet (earliest ETA first alone would have dropped the 700 as received).
  const bs = [
    alloc({ style: 'AB321', styleRaw: 'AB321', po: '900942A', bucket: 'port', units: 700, eta: D(10, 20) }),
    alloc({ style: 'AB321', styleRaw: 'AB321', po: '900942B', bucket: 'water', units: 900, eta: D(11, 20) }),
    // Two of ours with the ERP's quantity: the one with the closest ETA is matched.
    alloc({ style: 'AB322', styleRaw: 'AB322', po: '900943A', bucket: 'port', units: 720, eta: D(10, 20) }),
    alloc({ style: 'AB322', styleRaw: 'AB322', po: '900943B', bucket: 'water', units: 720, eta: D(11, 13) }),
    alloc({ style: 'AB322', styleRaw: 'AB322', po: '900943C', bucket: 'water', units: 1000, eta: D(11, 20) }),
  ];
  const ats3 = atsE({ AB321: [5000, 0, [[700, D(10, 31)]]], AB322: [5000, 0, [[720, D(11, 13)]]] });
  const r3 = sell(bs, ats3), t1 = styleOf(r3, 'AB321'), t2 = styleOf(r3, 'AB322');
  assert.deepEqual([lotOf(r3, '900942A').matched, lotOf(r3, '900942B').notOnErp, t1.receivedDropped, t1.arriving], [700, 900, 0, 1600]);
  assert.deepEqual(['900943A', '900943B', '900943C'].map(po => { const l = lotOf(r3, po); return [l.matched, l.received, l.notOnErp]; }),
    [[0, 720, 0], [720, 0, 0], [0, 0, 1000]]);
  assert.deepEqual([t2.receivedDropped, t2.keptExtra, t2.arriving], [720, 1000, 1720]);
  reconcile(r3, bs, ats3);
});

test('step 2: In Trans due after the report day has left China: off our lots there, on a vessel first, then a PO Container Tracking shows sailed, then by pull-from date', () => {
  // A PO split in China: our 3,024 on the water is the ERP's shipment of that quantity; its 720 and 816 due 11/13 are the
  // goods Container Tracking has on the water under PO 900952 (no Master Sheet line), which our sheets still show Ready
  // on 900952A and 900952B. Until Oct 30 they were offered twice: Arriving, and from China (here taken by orders).
  const as = [
    alloc({ style: 'AB400', styleRaw: 'AB400', po: '900951', bucket: 'water', stage: 'On the water', units: 3024, eta: D(11, 13) }),
    alloc({ style: 'AB400', styleRaw: 'AB400', po: '900952A', bucket: 'ready', stage: 'Ready', units: 816, crd: D(9, 10), dates: { factoryCrd: D(9, 10) } }),
    alloc({ style: 'AB400', styleRaw: 'AB400', po: '900952B', bucket: 'ready', stage: 'Ready', units: 720, crd: D(9, 10), dates: { factoryCrd: D(9, 10) } }),
    alloc({ style: 'AB400', styleRaw: 'AB400', po: '900952', bucket: 'unmatched', stage: 'On the water', units: 1536, eta: D(11, 13), sheetRow: null, ptype: '', flags: ['near-miss'] }),
  ];
  const ats = atsE({ AB400: [0, 1536, [[3024, D(11, 3)], [720, D(11, 13)], [816, D(11, 13)]], []] });
  const r = sell(as, ats), s = styleOf(r, 'AB400');
  assert.deepEqual([s.matchedOurs, s.landedNotReceived, s.takenOnVessel, s.takenSamePo, s.takenOther, s.erpUnplaced, s.erpBacklog, s.takeable],
    [3024, 0, 0, 1536, 0, 0, 0, 1536]);
  assert.deepEqual(['900952A', '900952B'].map(po => [lotOf(r, po).inErpTransit, lotOf(r, po).erpTier, lotOf(r, po).available]), [[816, 'po', 0], [720, 'po', 0]]);
  assert.deepEqual(erpLots(r, 'AB400').map(l => [l.units, l.place.matched, l.place.offChina]), [[3024, 3024, 0], [720, 0, 720], [816, 0, 816]]);
  assert.deepEqual([s.fromChina, s.arriving, s.ordersTaken, s.total, s.oversold], [0, 3024, 1536, 3024, 0]);
  assert.ok(s.flags.includes('in-erp-transit') && !s.flags.includes('erp-transit-unplaced'));
  reconcile(r, as, ats);
  // Every tier, in order: on a vessel (earliest ETD first), then a PO whose base Container Tracking shows sailed or
  // landed (900975 landed in July, too long ago to be the In Trans itself), then the rest by pull-from date.
  const bs = [
    alloc({ style: 'AB410', styleRaw: 'AB410', po: '900971', bucket: 'booked', stage: 'Booked', units: 500, etd: D(11, 20) }),
    alloc({ style: 'AB410', styleRaw: 'AB410', po: '900972', bucket: 'loaded', stage: 'Loaded', units: 300, etd: D(11, 5) }),
    alloc({ style: 'AB410', styleRaw: 'AB410', po: '900973', bucket: 'ready', stage: 'Ready', units: 400, crd: D(10, 1), dates: { factoryCrd: D(10, 1) } }),
    alloc({ style: 'AB410', styleRaw: 'AB410', po: '900974', bucket: 'inProd', units: 600, crd: D(11, 15), dates: { factoryCrd: D(11, 15) } }),
    alloc({ style: 'AB410', styleRaw: 'AB410', po: '900975A', bucket: 'inProd', units: 200, crd: D(12, 20), dates: { factoryCrd: D(12, 20) } }),
    alloc({ style: 'AB410', styleRaw: 'AB410', po: '900975', bucket: 'landed', stage: 'Landed', units: 900, eta: D(7, 1) }),
    // The shipment's own PO (a PO column on the report) counts too: 900982-BCF before 900981's earlier date.
    alloc({ style: 'AB420', styleRaw: 'AB420', po: '900981', bucket: 'ready', stage: 'Ready', units: 500, crd: D(9, 1), dates: { factoryCrd: D(9, 1) } }),
    alloc({ style: 'AB420', styleRaw: 'AB420', po: '900982-BCF', bucket: 'ready', stage: 'Ready', units: 500, crd: D(9, 20), dates: { factoryCrd: D(9, 20) } }),
  ];
  const ats2 = atsE({ AB410: [0, 0, [[1700, D(11, 25)]]], AB420: [0, 0, [[500, D(11, 20), '900982']]] });
  const r2 = sell(bs, ats2), t = styleOf(r2, 'AB410');
  assert.deepEqual([t.takenOnVessel, t.takenSamePo, t.takenOther, t.erpUnplaced, t.landedOurs], [800, 200, 700, 0, 0]);
  assert.deepEqual(['900972', '900971', '900975A', '900973', '900974'].map(po => [lotOf(r2, po).inErpTransit, lotOf(r2, po).erpTier]),
    [[300, 'vessel'], [500, 'vessel'], [200, 'po'], [400, 'other'], [300, 'other']]);
  assert.deepEqual([lotOf(r2, '900982-BCF').inErpTransit, lotOf(r2, '900982-BCF').erpTier, lotOf(r2, '900981').inErpTransit, erpLots(r2, 'AB420')[0].erpPo], [500, 'po', 0, '900982']);
  reconcile(r2, bs, ats2);
});

test("the guard: never taking our China goods below the ATS report's Import PO Backlog; what's left stays Arriving, one Problems line per style", () => {
  const as = [
    // Backlog 1,000 of the 1,500 in China: only 500 may come off (the Loaded lot); 700 stay unplaced, as the Ready 1,000
    // are still in China on our sheets and the ERP's.
    alloc({ style: 'AB503', styleRaw: 'AB503', po: '900991', bucket: 'ready', stage: 'Ready', units: 1000, crd: D(9, 1), dates: { factoryCrd: D(9, 1) } }),
    alloc({ style: 'AB503', styleRaw: 'AB503', po: '900992', bucket: 'loaded', stage: 'Loaded', units: 500, etd: D(11, 5) }),
    // Sold goods count in the backlog too: 480 sold + 1,344 Loaded, backlog 480, so the 1,344 may come off.
    alloc({ style: 'AB504', styleRaw: 'AB504', po: '900993', bucket: 'loaded', stage: 'Loaded', units: 1344, etd: D(10, 30) }),
    alloc({ style: 'AB504', styleRaw: 'AB504', po: '900994', bucket: 'pullFromChina', stage: 'Booked', ptype: 'China Pickup', units: 480 }),
    // The backlog has all of it in China: nothing comes off.
    alloc({ style: 'AB506', styleRaw: 'AB506', po: '900996', bucket: 'ready', stage: 'Ready', units: 600, crd: D(9, 1), dates: { factoryCrd: D(9, 1) } }),
  ];
  const ats = atsE({ AB503: [0, 0, [[1200, D(11, 20)]], [1000]], AB504: [0, 0, [[1344, D(11, 13)]], [480]],
    AB505: [0, 0, [[800, D(11, 9)]], []], AB506: [0, 0, [[600, D(11, 9)]], [600]] });
  const r = sell(as, ats), s = k => styleOf(r, k);
  assert.deepEqual([s('AB503').takeable, s('AB503').takenOnVessel, s('AB503').takenOther, s('AB503').erpUnplaced, s('AB503').unplacedInChina, s('AB503').fromChina, s('AB503').arriving],
    [500, 500, 0, 700, 700, 1000, 1200]);
  assert.deepEqual([s('AB504').takeable, s('AB504').takenOnVessel, s('AB504').erpUnplaced], [1344, 1344, 0]);
  assert.deepEqual([s('AB505').erpUnplaced, s('AB505').unplacedInChina, s('AB505').arriving], [800, 0, 800]);
  assert.deepEqual([s('AB506').takeable, s('AB506').inErpTransitTaken, s('AB506').erpUnplaced, s('AB506').unplacedInChina, s('AB506').fromChina], [0, 0, 600, 600, 600]);
  assert.deepEqual([r.counts.backlogGuardStyles, r.counts.erpUnplacedStyles, r.totals.erpUnplaced, r.totals.unplacedInChina, r.totals.offeredTwice], [2, 3, 2100, 1300, 0]);
  const un = r.issues.filter(i => i.code === 'erp-transit-unplaced');
  assert.deepEqual(un.map(i => [i.style, i.units, i.inChina]), [['AB503', 700, 700], ['AB505', 800, 0], ['AB506', 600, 600]]);
  assert.match(un[0].message, /^Style AB503: the ERP counts 700 units in transit that our sheets still show in China; check before offering from China\. They're offered as Arriving, and our goods still in China stay available from China because the ATS report's Import PO Backlog still has 1,000 of the style there \(our sheets have 1,000, sold goods not loaded yet included, after 500 came off\)\.$/);
  assert.match(un[2].message, /^Style AB506: the ERP counts 600 units in transit that our sheets still show in China; check before offering from China\. .*\(our sheets have 600, sold goods not loaded yet included\)\.$/);
  assert.ok(erpLots(r, 'AB503')[0].flags.includes('erp-transit-unplaced') && erpLots(r, 'AB503')[0].flags.includes('erp-off-china'));
  reconcile(r, as, ats);
  // Without the Import PO Backlog column there's no guard: the rest comes off the Ready lot.
  const one = as.filter(a => a.style === 'AB503'), ats2 = atsE({ AB503: [0, 0, [[1200, D(11, 20)]]] });
  const t = styleOf(sell(one, ats2), 'AB503');
  assert.deepEqual([t.takenOnVessel, t.takenOther, t.erpUnplaced, t.erpBacklog, t.takeable], [500, 700, 0, null, null]);
  reconcile(sell(one, ats2), one, ats2);
});

// What the ERP itself has for a style, after its customer orders: In WH + In Trans + Import PO Backlog − orders. The
// tool's total plus what it holds back as sold (goods the backlog counts) must come to it when our sheets and the
// report agree.
const erpTotal = (inWh, orders, ships, backlog) => inWh + ships.reduce((t, x) => t + x[0], 0) + backlog.reduce((t, n) => t + n, 0) - orders;

test('step 1 gives way to the Import PO Backlog: In Trans it says has left China comes off China first, never offered twice', () => {
  const as = [
    // Booked lots the ERP has shipped: the ERP's three shipments due 10/28 are our Booked lots of those quantities (ETD 11/4); its backlog
    // (1,536) is our fourth Booked lot. Until now they were called landed and the Booked lots stayed: 5,088 offered twice.
    ...[['900A01', 1632], ['900A02', 1632], ['900A03', 1824]].map(([po, u]) => alloc({ style: 'AB600', styleRaw: 'AB600', po, bucket: 'booked', stage: 'Booked', units: u, etd: D(11, 4) })),
    alloc({ style: 'AB600', styleRaw: 'AB600', po: '900A04', bucket: 'booked', stage: 'Booked', units: 1536, etd: D(11, 12) }),
    // A landed lot already received: our Landed 5,008 (ETA 10/20) is the style's In WH, received; the ERP's 4,288 due 10/28 is our
    // Ready lot of that quantity (its backlog is blank: nothing in China). Until now our Landed goods took it.
    alloc({ style: 'AB601', styleRaw: 'AB601', po: '900A11', bucket: 'landed', stage: 'Landed', units: 5008, eta: D(10, 20) }),
    alloc({ style: 'AB601', styleRaw: 'AB601', po: '900A12', bucket: 'ready', stage: 'Ready', units: 4288, crd: D(9, 20), dates: { factoryCrd: D(9, 20) } }),
    alloc({ style: 'AB601', styleRaw: 'AB601', po: '900A13', bucket: 'water', stage: 'On the water', units: 752, eta: D(11, 11) }),
    // A stale Ready line (its goods landed long ago) beside a blank backlog: In Trans due 10/25 has left China.
    alloc({ style: 'AB602', styleRaw: 'AB602', po: '900A21', bucket: 'landed', stage: 'Landed', units: 900, eta: D(7, 1) }),
    alloc({ style: 'AB602', styleRaw: 'AB602', po: '900A22', bucket: 'ready', stage: 'Ready', units: 900, crd: D(6, 1), dates: { factoryCrd: D(6, 1) } }),
    // Only the part beyond what may come off is landed: backlog 1,000 of the 1,600 in China, 800 In Trans due 10/24 and
    // our Landed 800 (ETA 10/24): 600 off China (the Loaded lot first), 200 landed.
    alloc({ style: 'AB603', styleRaw: 'AB603', po: '900A31', bucket: 'landed', stage: 'Landed', units: 800, eta: D(10, 24) }),
    alloc({ style: 'AB603', styleRaw: 'AB603', po: '900A32', bucket: 'loaded', stage: 'Loaded', units: 600, etd: D(10, 31) }),
    alloc({ style: 'AB603', styleRaw: 'AB603', po: '900A33', bucket: 'inProd', units: 1000, crd: D(12, 1), dates: { factoryCrd: D(12, 1) } }),
    // Sold goods not loaded yet beyond the backlog can't come off (they aren't lots): what our unsold goods can't cover is
    // still landed, never left unplaced. 100 Ready + 500 sold (China Pickup, Booked), backlog 0, In Trans 400 due 10/25.
    alloc({ style: 'AB604', styleRaw: 'AB604', po: '900A41', bucket: 'landed', stage: 'Landed', units: 400, eta: D(10, 20) }),
    alloc({ style: 'AB604', styleRaw: 'AB604', po: '900A42', bucket: 'ready', stage: 'Ready', units: 100, crd: D(9, 1), dates: { factoryCrd: D(9, 1) } }),
    alloc({ style: 'AB604', styleRaw: 'AB604', po: '900A43', bucket: 'pullFromChina', stage: 'Booked', ptype: 'China Pickup', units: 500 }),
  ];
  const E = { AB600: [0, 5088, [[1632, D(10, 28)], [1632, D(10, 28)], [1824, D(10, 28)]], [1536]],
    AB601: [5008, 752, [[4288, D(10, 28)], [752, D(11, 5)]], []], AB602: [900, 0, [[900, D(10, 25)]], []],
    AB603: [0, 0, [[800, D(10, 24)]], [1000]], AB604: [0, 0, [[400, D(10, 25)]], []] };
  const ats = atsE(E);
  const r = sell(as, ats), s = k => styleOf(r, k);
  const a0 = s('AB600');
  assert.deepEqual([a0.landedNotReceived, a0.takenOnVessel, a0.takenDueBy, a0.erpUnplaced, a0.takeable], [0, 5088, 5088, 0, 5088]);
  assert.deepEqual(['900A01', '900A02', '900A03', '900A04'].map(po => lotOf(r, po).inErpTransit), [1632, 1632, 1824, 0], 'the earliest ETDs come off');
  assert.deepEqual([a0.now, a0.fromChina, a0.arriving, a0.ordersTaken, a0.total, a0.oversold], [0, 1536, 0, 5088, 1536, 0]);
  assert.equal(a0.total, erpTotal(...E.AB600), "the ERP's own total");
  assert.ok(erpLots(r, 'AB600').every(l => l.place.offChina === l.units && l.flags.includes('erp-off-china') && !l.flags.includes('landed-not-received')));
  const a1 = s('AB601');
  assert.deepEqual([a1.matchedOurs, a1.landedOurs, a1.landedDue, a1.inErpTransitTaken, a1.takenDueBy, a1.erpUnplaced], [752, 0, 0, 4288, 4288, 0]);
  assert.deepEqual([a1.now, a1.fromChina, a1.arriving, a1.total], [4256, 0, 5040, 9296]);
  assert.equal(a1.total, erpTotal(...E.AB601));
  const a2 = s('AB602');
  assert.deepEqual([a2.landedDue, a2.takenOther, a2.fromChina, a2.arriving, a2.total], [0, 900, 0, 900, 1800]);
  assert.equal(a2.total, erpTotal(...E.AB602));
  const a3 = s('AB603');
  assert.deepEqual([a3.takeable, a3.takenOnVessel, a3.takenOther, a3.landedOurs, a3.landedLots.map(x => x.units), a3.erpUnplaced, a3.fromChina, a3.arriving],
    [600, 600, 0, 200, [200], 0, 1000, 800]);
  assert.equal(a3.total, erpTotal(...E.AB603));
  const a4 = s('AB604');
  assert.deepEqual([a4.takeable, a4.inErpTransitTaken, a4.landedOurs, a4.erpUnplaced, a4.fromChina, a4.arriving], [600, 100, 300, 0, 0, 400]);
  assert.ok(!codes(r).includes('erp-transit-unplaced') && !codes(r).includes('balance'));
  assert.deepEqual([r.counts.takenDueByStyles, r.totals.takenDueBy, r.totals.offeredTwice], [5, 5088 + 4288 + 900 + 600 + 100, 0]);
  reconcile(r, as, ats);
  // The tradeoff: a backlog that has every China lot keeps step 1 as it was (the landed PO above); with no backlog
  // column, nothing says the goods have left China, so step 1 is never limited.
  const ats2 = atsE(Object.fromEntries(Object.entries(E).map(([k, v]) => [k, v.slice(0, 3)])));
  const r2 = sell(as, ats2), b0 = styleOf(r2, 'AB600');
  assert.deepEqual([b0.landedDue, b0.inErpTransitTaken, b0.fromChina, b0.arriving, b0.total], [5088, 0, 1536, 5088, 6624]);
  reconcile(r2, as, ats2);
});

test('sold goods count against the Import PO Backlog only until they are loaded', () => {
  // The report's backlog has sold goods booked or ready, never those already loaded (China Pickup or straight to the
  // customer). Ready 500 + 500 sold already loaded, backlog 500, In Trans 500 due 11/20: the Ready lot is the backlog,
  // so it stays; the In Trans is unplaced (listed), never taken off it.
  const as = [
    alloc({ style: 'AB610', styleRaw: 'AB610', po: '900B01', bucket: 'ready', stage: 'Ready', units: 500, crd: D(9, 1), dates: { factoryCrd: D(9, 1) } }),
    alloc({ style: 'AB610', styleRaw: 'AB610', po: '900B02', bucket: 'pullFromChina', stage: 'Loaded', units: 500, flags: ['pull-section'] }),
    // The same with the sold goods still booked: the backlog is those, so the Ready lot has left China.
    alloc({ style: 'AB611', styleRaw: 'AB611', po: '900B11', bucket: 'ready', stage: 'Ready', units: 500, crd: D(9, 1), dates: { factoryCrd: D(9, 1) } }),
    alloc({ style: 'AB611', styleRaw: 'AB611', po: '900B12', bucket: 'pullFromChina', stage: 'Booked', ptype: 'China Pickup', units: 500 }),
    // A sold lot at Loaded in a counted bucket (straight to the customer) doesn't count either.
    alloc({ style: 'AB612', styleRaw: 'AB612', po: '900B21', bucket: 'ready', stage: 'Ready', units: 300, crd: D(9, 1), dates: { factoryCrd: D(9, 1) } }),
    alloc({ style: 'AB612', styleRaw: 'AB612', po: '900B22', bucket: 'loaded', stage: 'Loaded', units: 200, etd: D(11, 1), flags: ['pull-section'] }),
  ];
  const E = { AB610: [0, 0, [[500, D(11, 20)]], [500]], AB611: [0, 0, [[500, D(11, 20)]], [500]], AB612: [0, 0, [[300, D(11, 20)]], [300]] };
  const ats = atsE(E);
  const r = sell(as, ats), s = k => styleOf(r, k);
  assert.deepEqual([s('AB610').chinaAll, s('AB610').takeable, s('AB610').inErpTransitTaken, s('AB610').erpUnplaced, s('AB610').unplacedInChina, s('AB610').fromChina, s('AB610').arriving],
    [500, 0, 0, 500, 500, 500, 500]);
  assert.equal(s('AB610').total, erpTotal(...E.AB610), "the ERP's own total (its sold goods already loaded are in neither)");
  assert.deepEqual([s('AB611').chinaAll, s('AB611').takeable, s('AB611').inErpTransitTaken, s('AB611').fromChina, s('AB611').arriving], [1000, 500, 500, 0, 500]);
  assert.equal(s('AB611').total + 500, erpTotal(...E.AB611), "the ERP's own total less the sold goods in its backlog");
  assert.deepEqual([s('AB612').chinaAll, s('AB612').takeable, s('AB612').inErpTransitTaken, s('AB612').erpUnplaced, s('AB612').fromChina], [300, 0, 0, 300, 300]);
  assert.match(r.issues.find(i => i.code === 'erp-transit-unplaced' && i.style === 'AB610').message, /Import PO Backlog still has 500 of the style there \(our sheets have 500, sold goods not loaded yet included\)/);
  reconcile(r, as, ats);
});

test('ERP In Trans below ours: In Trans due after the report day comes off China as far as the backlog says, and ours with an ETA passed are received instead', () => {
  // At port 1,000 (ETA 10/20, In WH 5,000) and Loaded 1,000 (ETD 11/1); the ERP has 600 In Trans due 11/15 and a backlog
  // of 400: the 600 are off the Loaded lot, and our 1,000 at port were received. Until now 600 were offered twice.
  const as = [
    alloc({ style: 'AB620', styleRaw: 'AB620', po: '900C01', bucket: 'port', stage: 'At port', units: 1000, eta: D(10, 20) }),
    alloc({ style: 'AB620', styleRaw: 'AB620', po: '900C02', bucket: 'loaded', stage: 'Loaded', units: 1000, etd: D(11, 1) }),
    // The same, In WH 700: only 300 more can have been received, so only 300 come off China (the rest as before).
    alloc({ style: 'AB621', styleRaw: 'AB621', po: '900C11', bucket: 'port', stage: 'At port', units: 1000, eta: D(10, 20) }),
    alloc({ style: 'AB621', styleRaw: 'AB621', po: '900C12', bucket: 'loaded', stage: 'Loaded', units: 1000, etd: D(11, 1) }),
    // The same, its In Trans due by the report's day: our goods at port, not received yet (nothing comes off China).
    alloc({ style: 'AB622', styleRaw: 'AB622', po: '900C21', bucket: 'port', stage: 'At port', units: 1000, eta: D(10, 20) }),
    alloc({ style: 'AB622', styleRaw: 'AB622', po: '900C22', bucket: 'loaded', stage: 'Loaded', units: 1000, etd: D(11, 1) }),
  ];
  const E = { AB620: [5000, 0, [[600, D(11, 15)]], [400]], AB621: [700, 0, [[600, D(11, 15)]], [400]], AB622: [5000, 0, [[600, D(10, 27)]], [400]] };
  const ats = atsE(E);
  const r = sell(as, ats), s = k => styleOf(r, k);
  assert.deepEqual([lotOf(r, '900C01').received, lotOf(r, '900C01').matched, lotOf(r, '900C02').inErpTransit, lotOf(r, '900C02').erpTier, lotOf(r, '900C02').available], [1000, 0, 600, 'vessel', 400]);
  assert.deepEqual([s('AB620').matchedOurs, s('AB620').receivedDropped, s('AB620').inErpTransitTaken, s('AB620').fromChina, s('AB620').arriving, s('AB620').total], [0, 1000, 600, 400, 600, 6000]);
  assert.equal(s('AB620').total, erpTotal(...E.AB620));
  assert.deepEqual(erpLots(r, 'AB620').map(l => [l.place.matched, l.place.offChina, l.flags.includes('erp-off-china')]), [[0, 600, true]]);
  assert.deepEqual([s('AB621').receivedDropped, s('AB621').matchedOurs, s('AB621').inErpTransitTaken, s('AB621').fromChina], [700, 300, 300, 700]);
  assert.deepEqual([s('AB622').receivedDropped, s('AB622').matchedOurs, s('AB622').inErpTransitTaken, s('AB622').fromChina], [400, 600, 0, 1000]);
  assert.deepEqual([r.counts.lessTakenStyles, r.counts.erpLessStyles], [2, 3]);
  assert.ok(!codes(r).includes('balance'));
  reconcile(r, as, ats);
  // No Import PO Backlog column: nothing says the Loaded lot has left China.
  const ats2 = atsE(Object.fromEntries(Object.entries(E).map(([k, v]) => [k, v.slice(0, 3)])));
  assert.deepEqual([styleOf(sell(as, ats2), 'AB620').inErpTransitTaken, styleOf(sell(as, ats2), 'AB620').receivedDropped], [0, 400]);
  reconcile(sell(as, ats2), as, ats2);
});

test('an Import PO Backlog part below zero or not a number is listed, and the style has no backlog (end to end through parseAts)', () => {
  const X = loadXLSX();
  const HEAD = ['Style', 'Warehouse', 'In WH', 'Total Customer Orders', 'In Trans Qty', 'Import PO Backlog Qty', 'ETA Date'];
  const rows = [['Inventory Report'], HEAD,
    ['AB630', 'WH 1', 3456, 5088, '5040, 1632, 1632', '-1632', '10/28/2026, 11/11/2026, 11/11/2026'],
    ['AB631', 'WH 1', 0, 0, '500', '48, soon', '11/20/2026'],
    ['AB632', 'WH 1', 0, 0, '500', '400', '11/20/2026'],
    ['Grand Total', null, 3456, 5088, null, null, null]];
  const ats = A.parseAts(X, readBook(xlsxBytes({ Report: rows })), 'Inventory Report 10-29-2026.xlsx', TODAY);
  assert.equal(ats.ok, true, ats.error);
  const nb = ats.issues.filter(i => i.code === 'negative-backlog'), nn = ats.issues.filter(i => i.code === 'not-a-number');
  assert.deepEqual([nb.map(i => i.style), nn.map(i => i.style), ats.counts.backlogUnreadable], [['AB630'], ['AB631'], 2]);
  assert.match(nb[0].message, /^Style AB630 \(row 3\): Import PO Backlog Qty -1,632 is below zero \(a correction in the ERP\?\), so the style's goods the ERP still has in China aren't known/);
  assert.match(nn[0].message, /^Style AB631 \(row 4\): Import PO Backlog Qty "soon" isn't a number/);
  assert.deepEqual([ats.styles.get('AB631').rows[0].backlogList, ats.styles.get('AB631').rows[0].backlogBad], [[48], ['soon']]);
  // A negative backlog: ours on the water 3,840 and 1,632, Ready 1,200; In Trans 5,040 + 1,632 + 1,632 and a backlog of
  // −1,632 (the ERP's own correction of one 1,632?). The backlog isn't used: the Ready lot comes off, 432 are unplaced, and
  // the Problems line names the backlog.
  const as = [
    alloc({ style: 'AB630', styleRaw: 'AB630', po: '900D01', bucket: 'water', stage: 'On the water', units: 3840, eta: D(10, 28) }),
    alloc({ style: 'AB630', styleRaw: 'AB630', po: '900D02', bucket: 'water', stage: 'On the water', units: 1632, eta: D(11, 12) }),
    alloc({ style: 'AB630', styleRaw: 'AB630', po: '900D03', bucket: 'ready', stage: 'Ready', units: 1200, crd: D(8, 15), dates: { factoryCrd: D(8, 15) } }),
    alloc({ style: 'AB631', styleRaw: 'AB631', po: '900D11', bucket: 'ready', stage: 'Ready', units: 500, crd: D(9, 1), dates: { factoryCrd: D(9, 1) } }),
    alloc({ style: 'AB632', styleRaw: 'AB632', po: '900D21', bucket: 'ready', stage: 'Ready', units: 500, crd: D(9, 1), dates: { factoryCrd: D(9, 1) } }),
  ];
  const r = sell(as, ats), s = k => styleOf(r, k);
  assert.deepEqual([s('AB630').erpBacklog, s('AB630').erpBacklogOff, s('AB630').takeable, s('AB630').matchedOurs, s('AB630').landedDue, s('AB630').inErpTransitTaken, s('AB630').erpUnplaced],
    [null, [-1632], null, 5472, 1200, 1200, 432]);
  assert.ok(s('AB630').flags.includes('erp-backlog-unknown') && s('AB631').flags.includes('erp-backlog-unknown') && !s('AB632').flags.includes('erp-backlog-unknown'));
  assert.match(r.issues.find(i => i.code === 'erp-transit-unplaced' && i.style === 'AB630').message, /The ATS report's Import PO Backlog for the style \(-1,632\) is below zero or not a number, so it isn't used; it may be the ERP's own correction of this In Trans\.$/);
  // "48, soon": not 48. With no backlog, the 500 come off the Ready lot; with its 400, only 100 may.
  assert.deepEqual([s('AB631').erpBacklog, s('AB631').inErpTransitTaken, s('AB632').erpBacklog, s('AB632').inErpTransitTaken, s('AB632').erpUnplaced], [null, 500, 400, 100, 400]);
  assert.equal(r.counts.backlogUnknownStyles, 2);
  reconcile(r, as, ats);
});

test('no report day: nothing is called landed; In Trans beyond ours comes off China, within the backlog', () => {
  const as = [
    alloc({ style: 'AB500', styleRaw: 'AB500', po: '900960', bucket: 'landed', stage: 'Landed', units: 5040, eta: D(10, 22) }),
    alloc({ style: 'AB500', styleRaw: 'AB500', po: '900961', bucket: 'booked', stage: 'Booked', units: 10008, etd: D(11, 11) }),
  ];
  const ats = atsE({ AB500: [0, 0, [[5040, D(10, 22)]], [10008]] }, null);
  const r = sell(as, ats), s = styleOf(r, 'AB500');
  assert.equal(r.reportDay, null);
  assert.deepEqual([s.landedNotReceived, s.inErpTransitTaken, s.erpUnplaced, s.unplacedInChina, s.fromChina, s.arriving], [0, 0, 5040, 5040, 10008, 5040]);
  assert.ok(s.flags.includes('no-report-day') && s.flags.includes('erp-transit-unplaced'));
  assert.match(r.issues.find(i => i.code === 'no-report-day').message,
    /^1 style \(0 of them oversold\): the ATS report's date isn't known, so none of its In Trans beyond our sailed goods is called landed, not received yet: it comes off our goods still in China instead, as far as its Import PO Backlog allows/);
  reconcile(r, as, ats);
  const ats2 = atsE({ AB500: [0, 0, [[5040, D(10, 22)]]] }, null), t = styleOf(sell(as, ats2), 'AB500');
  assert.deepEqual([t.takenOnVessel, t.erpUnplaced, t.fromChina], [5040, 0, 4968]);
  reconcile(sell(as, ats2), as, ats2);
});

test('ETA Date list that does not line up with In Trans Qty: no ETA for those units, listed by row (end to end through parseAts)', () => {
  const X = loadXLSX();
  const HEAD = ['Color Code', 'Style', 'Style Description', 'Warehouse', 'In WH', 'Total Customer Orders', 'In Trans Qty', 'ETA Date'];
  const rows = [['Inventory Report'], HEAD,
    ['XX', 'AB100', 'Item', 'WH 1', 0, 0, '1,000, 2,000', '11/14/2026'],            // 2 quantities, 1 date
    ['XX', 'AB200', 'Item', 'WH 1', 0, 0, '1200, 3400', '11/14/2026, soon'],         // lined up, one date unreadable
    ['XX', 'AB300', 'Item', 'WH 1', 0, 0, '1200, 3400', '11/14/2026, 12/1/2026'],    // fine
    ['XX', 'AB400', 'Item', 'WH 1', 0, 0, '500', null],                              // a quantity, no date
    ['Grand Total', null, null, null, 0, 0, null, null]];
  const ats = A.parseAts(X, readBook(xlsxBytes({ Report: rows })), 'Inventory Report - On Hand_In Transit_In Production 10-29-2026 12-58 PM.xlsx', TODAY);
  assert.equal(ats.ok, true, ats.error);
  assert.deepEqual(ats.styles.get('AB200').rows[0].etaParts, [D(11, 14), null], 'the unreadable date keeps its place');
  assert.deepEqual(ats.styles.get('AB200').rows[0].eta, [D(11, 14)]);
  const as = ['AB100', 'AB200', 'AB300', 'AB400'].map((k, i) => alloc({ style: k, styleRaw: k, po: '90095' + i, bucket: 'inProd', units: 10, crd: D(12, 1), dates: { factoryCrd: D(12, 1) } }));
  const r = sell(as, ats);
  assert.deepEqual(erpLots(r, 'AB100').map(l => [l.units, l.eta, l.flags.includes('erp-eta-mismatch'), l.flags.includes('no-date')]), [[1000, null, true, true], [2000, null, true, true]]);
  assert.deepEqual(erpLots(r, 'AB200').map(l => [l.units, l.eta]), [[1200, D(11, 14)], [3400, null]]);
  assert.deepEqual(erpLots(r, 'AB300').map(l => [l.units, l.eta]), [[1200, D(11, 14)], [3400, D(12, 1)]]);
  assert.deepEqual(erpLots(r, 'AB400').map(l => [l.units, l.eta]), [[500, null]]);
  assert.deepEqual([styleOf(r).arriving, styleOf(r).firstArrival], [3000, null], 'still counted as Arriving, with no date');
  assert.ok(styleOf(r).flags.includes('erp-eta-mismatch') && !styleOf(r, 'AB300').flags.includes('erp-eta-mismatch'));
  const mm = r.issues.filter(i => i.code === 'erp-eta-mismatch');
  assert.deepEqual(mm.map(i => [i.style, i.units]), [['AB100', 3000], ['AB200', 3400], ['AB400', 500]]);
  assert.match(mm[0].message, /^Style AB100: the ATS report's ETA Date list doesn't line up with its In Trans Qty list \(row 3: 2 In Trans quantities, 1 ETA date\), so 3,000 units In Trans have no arrival date/);
  assert.match(mm[1].message, /row 4: 1 ETA date that can't be read/);
  assert.ok(!codes(r).includes('no-date'), 'listed once, as a mismatch');
  assert.equal(r.counts.erpEtaMismatchStyles, 3);
  assert.equal(r.counts.erpLotsNoEta, 4);
  reconcile(r, as, ats);
});

test('In Trans that cannot be read, a style not on the report, no In Trans column, no report: Arriving is our own sailed lots', () => {
  const X = loadXLSX();
  const HEAD = ['Style', 'Warehouse', 'In WH', 'Total Customer Orders', 'In Trans Qty', 'ETA Date'];
  const rows = [['Inventory Report'], HEAD, ['AB100', 'WH 1', 0, 0, '1000, lots', '11/14/2026, 11/20/2026'], ['AB101', 'WH 1', 0, 0, '-500', '11/14/2026'],
    // Two quantities with no space: never one number of 10,001,400. "720,816" with a date for each part: one or two?
    ['AB103', 'WH 1', 0, 0, '1000,1400', '11/14/2026,12/01/2026'], ['AB104', 'WH 1', 0, 0, '720,816', '11/13/2026, 11/13/2026'],
    ['Grand Total', null, 0, 0, null, null]];
  const ats = A.parseAts(X, readBook(xlsxBytes({ Report: rows })), 'Inventory Report 10-29-2026.xlsx', TODAY);
  assert.equal(ats.ok, true, ats.error);
  const nn = ats.issues.filter(i => i.code === 'not-a-number');
  assert.equal(nn.length, 3);
  assert.match(nn[0].message, /^Style AB100 \(row 3\): In Trans Qty "lots" isn't a number, so the ERP's In Trans for the style isn't known/);
  assert.match(nn[1].message, /^Style AB103 \(row 5\): In Trans Qty "1000,1400" isn't a number/);
  assert.match(nn[2].message, /^Style AB104 \(row 6\): In Trans Qty "720,816" could be one quantity or two/);
  const neg = ats.issues.filter(i => i.code === 'negative-in-trans');
  assert.equal(neg.length, 1);
  assert.match(neg[0].message, /^Style AB101 \(row 4\): In Trans Qty -500 is below zero, so the ERP's In Trans for the style isn't known/);
  assert.deepEqual([ats.styles.get('AB103').rows[0].inTransList, ats.styles.get('AB104').rows[0].inTransList], [[], []]);
  assert.equal(ats.counts.inTransUnreadable, 4);
  const as = [alloc({ po: '900961', bucket: 'water', units: 700, eta: D(11, 2) }), alloc({ style: 'AB101', styleRaw: 'AB101', po: '900962', bucket: 'port', units: 300, eta: D(10, 1) }),
    alloc({ style: 'AB102', styleRaw: 'AB102', po: '900963', bucket: 'water', units: 400, eta: D(11, 3) })];
  const r = sell(as, ats);
  for (const k of ['AB100', 'AB101', 'AB103', 'AB104']) {
    const s = styleOf(r, k);
    assert.deepEqual([s.arrivingFrom, s.erpInTrans, s.arrivingOurs, s.arriving, erpLots(r, k).length], ['ours', null, s.arriving, s.arriving, 0], k);
    assert.ok(s.flags.includes('erp-in-trans-unknown'), k);
  }
  assert.deepEqual([styleOf(r, 'AB103').arriving, styleOf(r, 'AB104').arriving], [0, 0], 'no phantom units');
  assert.equal(styleOf(r).arriving, 700);
  assert.equal(styleOf(r, 'AB101').arriving, 300, 'an ETA long passed, but nothing to say it was received');
  // Not on the report: our sailed lots, in full, flagged as today.
  const s2 = styleOf(r, 'AB102');
  assert.deepEqual([s2.arrivingFrom, s2.arriving, s2.arrivingOurs, s2.onReport], ['ours', 400, 400, false]);
  assert.ok(s2.flags.includes('not-on-report') && s2.flags.includes('orders-unknown') && lotOf(r, '900963').flags.includes('orders-unknown'));
  assert.equal(r.counts.erpInTransUnknownStyles, 4);
  reconcile(r, as, ats);
  // The report without its In Trans Qty column: our sailed lots for every style, one line on Problems.
  const r2 = sell(as.map(a => ({ ...a })), atsOf({ AB100: [0, 0], AB101: [0, 0] }));
  assert.ok(r2.styles.every(s => s.arrivingFrom === 'ours'));
  assert.deepEqual(r2.styles.map(s => s.arriving), [700, 300, 400]);
  const nc = r2.issues.filter(i => i.code === 'no-in-trans-column');
  assert.equal(nc.length, 1);
  assert.match(nc[0].message, /^The ATS report has no "In Trans Qty" column, so Arriving comes from our shipping data/);
  reconcile(r2, as, null);
  // No report (or a red one): our sailed lots, orders not taken off, as before.
  for (const a of [null, { ok: false, error: 'red', styles: new Map() }]) {
    const r3 = sell(as.map(x => ({ ...x })), a);
    assert.ok(r3.atsMissing && r3.styles.every(s => s.arrivingFrom === 'ours' && s.erpInTrans === null));
    assert.deepEqual([r3.totals.arriving, r3.totals.arrivingOurs, r3.totals.arrivingErp], [1400, 1400, 0]);
    assert.ok(!codes(r3).includes('no-in-trans-column'));
    reconcile(r3, as, null);
  }
});

test('orders-may-include-sold stays: orders beyond stock no more than the goods sold to ship from China', () => {
  const r = sell([alloc({ style: 'AB300', styleRaw: 'AB300', po: '901001', bucket: 'pullFromChina', units: 1104, ptype: 'China Pickup' })],
    atsE({ AB300: [0, 1104] }));
  const s3 = styleOf(r, 'AB300');
  assert.deepEqual([s3.oversold, s3.soldInOrders, s3.flags.includes('orders-may-include-sold')], [1104, 1104, true]);
  assert.match(r.issues.find(i => i.code === 'oversold').message, /1,104 units of this style are sold to ship from China: if those orders are in the ERP's Total Customer Orders, they're taken off twice/);
  assert.equal(r.issues.filter(i => i.code === 'orders-may-include-sold').length, 1);
  assert.deepEqual([r.counts.soldInOrdersStyles, r.counts.soldInOrdersUnits], [1, 1104]);
  // An oversold style with nothing that may cover it still says so plainly.
  const r2 = sell([], atsE({ AB500: [0, 50] }));
  assert.match(r2.issues.find(i => i.code === 'oversold').message, /with no open lot from China\. Nothing is available to sell\./);
});

test('every unit accounted for on a large random mix with the ERP In Trans: Incoming, In Trans, received, landed, off China, unplaced', () => {
  // mulberry32: every bit of it varies, so styles, buckets and sizes spread out.
  let seed = 11;
  const rnd = k => { seed = (seed + 0x6D2B79F5) | 0; let x = Math.imul(seed ^ (seed >>> 15), 1 | seed); x = (x + Math.imul(x ^ (x >>> 7), 61 | x)) ^ x; return ((x ^ (x >>> 14)) >>> 0) % k; };
  const buckets = ['noCrd', 'inProd', 'ready', 'booked', 'loaded', 'water', 'port', 'pullFromChina', 'landed', 'unmatched'];
  const as = [], map = {};
  for (let i = 0; i < 900; i++) {
    const style = 'AB' + (100 + rnd(80)), bucket = buckets[rnd(buckets.length)], day = rnd(5) ? D(8, 15) + rnd(150) : null;
    // Stages as the engine gives them: Container Tracking's for goods that left China, a sold line's at either side.
    const stage = { water: 'On the water', port: 'At port', landed: 'Landed', unmatched: rnd(2) ? 'On the water' : 'Booked',
      pullFromChina: ['Booked', 'On the water', 'Loaded'][rnd(3)] }[bucket] || 'In production';
    as.push(alloc({ style, styleRaw: style, po: String(900000 + rnd(300)) + (rnd(3) ? '' : 'A'), bucket, stage, units: 1 + rnd(4000), master: rnd(4) ? 24 : null,
      ptype: bucket === 'landed' && !rnd(8) ? 'China Pickup' : 'LA',
      etd: ['booked', 'loaded', 'landed'].includes(bucket) && rnd(3) ? day : null, eta: ['water', 'port', 'landed', 'unmatched'].includes(bucket) && rnd(6) ? day : null,
      crd: day, dates: { factoryCrd: day, orderDate: rnd(2) ? D(6, 1) + rnd(100) : null } }));
  }
  for (let i = 0; i < 90; i++) {
    if (!rnd(5)) continue;   // some styles not on the report
    const n = rnd(5), ships = [];
    for (let j = 0; j < n; j++) ships.push([1 + rnd(6000), rnd(6) ? D(10, 1) + rnd(60) : null, rnd(6) ? '' : String(900000 + rnd(300))]);
    // The Import PO Backlog: a number, blank (0), or no column for the style's row.
    map['AB' + (100 + i)] = [rnd(20000), rnd(30000), ships, rnd(3) ? [rnd(9000)] : []];
  }
  const ats = atsE(map);
  const r = sell(as, ats);
  reconcile(r, as, ats);
  assert.ok(!codes(r).includes('balance'), 'the engine finds every unit too');
  const t = r.totals;
  assert.equal(t.arriving, t.arrivingErp + t.arrivingOurs);
  assert.equal(t.erpInTrans, t.matchedOurs + t.landedNotReceived + t.inErpTransitTaken + t.erpUnplaced);
  assert.equal(t.fromChina + t.arriving + t.ordersTaken + t.matchedOurs + t.receivedDropped + t.inErpTransitTaken, t.ourIncoming + t.erpInTrans);
  assert.equal(t.offeredTwice, 0);
  // Every branch shows up: more, less, landed (ours, due), received, kept, taken off China (each tier; on shipments due
  // by the report's day as the backlog says; with the ERP's less), unplaced (some kept in China by the backlog).
  for (const k of ['erpMoreStyles', 'erpLessStyles', 'landedOursStyles', 'landedDueStyles', 'erpUnplacedStyles', 'unplacedInChinaStyles',
    'backlogGuardStyles', 'receivedStyles', 'keptExtraStyles', 'inErpTransitStyles', 'receivedCappedStyles', 'takenDueByStyles', 'lessTakenStyles']) assert.ok(r.counts[k] > 0, k);
  for (const k of SELL.TIERS) assert.ok(r.counts.inErpTransitTierLots[k] > 0, k);
  // Received only with an ETA on or before the report's day; in the ERP's transit only our lots still in China.
  for (const l of r.lots) {
    if (l.received) assert.ok(l.eta != null && l.eta <= D(10, 29), l.id);
    if (l.inErpTransit) assert.ok(l.inChina && l.source === 'ours', l.id);
  }
  // The same mix with no Import PO Backlog column: nothing guards China, so nothing our sheets still show there is left
  // unplaced; and nothing limits landed, so at least as much is called landed.
  const ats2 = atsE(Object.fromEntries(Object.entries(map).map(([k, v]) => [k, v.slice(0, 3)])));
  const r2 = sell(as, ats2);
  reconcile(r2, as, ats2);
  assert.deepEqual([r2.totals.unplacedInChina, r2.totals.offeredTwice, r2.counts.backlogGuardStyles, r2.counts.lessTakenStyles, r2.totals.takenDueBy], [0, 0, 0, 0, 0]);
  assert.ok(r2.totals.landedNotReceived >= t.landedNotReceived);
  // And with a backlog part below zero on some styles: their backlog isn't used (no guard), every unit still accounted for.
  const ats3 = atsE(Object.fromEntries(Object.entries(map).map(([k, v], i) => [k, i % 4 ? v : [...v.slice(0, 3), [-(1 + i), ...v[3]]]])));
  const r3 = sell(as, ats3);
  reconcile(r3, as, ats3);
  assert.ok(r3.counts.backlogUnknownStyles > 0 && r3.styles.filter(s => s.flags.includes('erp-backlog-unknown')).every(s => s.erpBacklog === null && s.takeable === null));
  // Customer orders: off the earliest lots with units left first.
  for (const s of r.styles) {
    let kept = false;
    for (const l of s.lots) { if (kept && l.taken) assert.fail(`style ${s.style}: a later lot taken first`); if (l.available > 0) kept = true; }
  }
});

test('with the stage engine itself: our lots hold its Incoming per style, and nothing is called a balance problem', () => {
  const T = new Map([['900971|AB100', { crd: D(12, 1), exp: null, sup: 'Factory A' }], ['900972|AB100', { crd: D(12, 1), exp: null, sup: 'Factory A' }]]);
  const ls = [
    ...block('900971', D(11, 10), [line({ po: '900971', crd: D(12, 1), od: D(8, 1), qty: 1000 })]),
    ...block('900972', D(12, 20), [line({ po: '900972', crd: D(12, 1), od: D(8, 2), qty: 600 })]),
  ];
  const ct = ctOf([blk('Loaded', '900971', [700], [{ qty: 1000 }])], [{ sheet: 'Open', rounds: [700], row: 9, etd: D(10, 1), eta: D(10, 31), containers: ['TEST1234567'] }]);
  const R = ST.build({ po: { lines: ls, trk: { T, minPO: 900000 } }, mrows: mrowsOf(ls), ct, today: TODAY });
  assert.equal(R.ok, true);
  const st = R.styles.find(s => s.style === 'AB100');
  const ats = atsE({ AB100: [0, 0, [[1000, D(10, 31)], [900, D(11, 30)]]] });
  const r = SELL.buildSell({ stages: R, ats, today: TODAY });
  const s = styleOf(r);
  assert.equal(s.ourIncoming, st.units.incoming);
  assert.equal(s.erpInTrans, 1900);
  assert.equal(s.fromChina + s.arriving + s.ordersTaken + s.matchedOurs + s.receivedDropped + s.inErpTransitTaken, st.units.incoming + 1900);
  assert.ok(!codes(r).includes('balance'), JSON.stringify(r.issues.filter(i => i.code === 'balance')));
  sumsHold(r);
});
