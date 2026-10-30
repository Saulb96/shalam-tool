import test from 'node:test';
import assert from 'node:assert/strict';
import { core, loadXLSX, xlsxBytes, readBook } from './load.mjs';

const C = core('common');
const ST = core('stages');
const TODAY = C.ymd(2026, 10, 28);
const D = (m, d, y = 2026) => C.ymd(y, m, d);
const CODES = new Set(['near-miss', 'unmatched', 'ct-more', 'short', 'shipped-not-on-ct', 'msl-behind-ct', 'ct-behind-msl',
  'both-tabs', 'no-etd', 'etd-passed-booked', 'unknown-type', 'no-case-pack', 'part-case', 'balance', 'missing-column',
  'other-trip', 'negative-qty', 'complete-in-transit', 'ignored-section', 'below-tracker', 'section-not-listed']);

// ---------- invented data builders (POs 9xxxxx, styles AB…, Factory A, Customer 1)

// Master Sheet Live rows: header on row 1 and each line's Barcode / Master / Category / Sub on its own row number.
const MH = ['Category', 'Sub', 'Product name', 'Style Number', 'Barcode', 'Master', 'Quantity', 'Po Number'];
function mrowsFor(lines) {
  const rows = [MH];
  for (const l of lines) rows[l.row - 1] = [l.cat ?? 'Lip', l.subcat ?? 'Gloss', 'Item', l.sty, l.upc ?? null, l.master ?? null, l.qty, l.po];
  for (let i = 0; i < rows.length; i++) if (!rows[i]) rows[i] = [];
  return rows;
}
let nextRow = 2;
// A PO Book line as applyOverlay gives it. Extra fields (upc, master, cat, subcat) only feed mrowsFor.
function line(o = {}) {
  const po = o.po ?? '900101', sty = o.sty ?? 'AB100', row = o.row ?? nextRow++;
  return { key: (po + '|' + sty.replace(/\s+/g, '')).toUpperCase(), po, base: String(po).slice(0, 6), sty, desc: 'Item ' + sty,
    sup: 'Factory A', od: null, qty: 1000, price: 1, crd: null, ptype: 'LA', cust: 'Customer 1', cpo: '', src: 'open',
    status: 'open', bo: false, tracked: true, ...o, row };
}
const shippedLine = o => line({ src: 'round', status: 'shipped', ...o });

let nextCtRow = 100;
// A Booking/Loaded PO block as InvCT gives it.
function blk(sheet, po, rounds, lns, extra = {}) {
  const row = nextCtRow; nextCtRow += lns.length + 2;
  return { sheet, row, section: rounds.length ? 'Round ' + rounds.join('-') : 'Planning', rounds, po, poRaw: po,
    factory: 'Factory A', orderDate: null, shipBy: null, readyDate: null, needBy: '', ...extra,
    lines: lns.map((l, i) => ({ row: row + 1 + i, style: l.style ?? 'AB100', styleRaw: l.style ?? 'AB100', upc: l.upc ?? null,
      upcRaw: l.upc ?? '', master: l.master ?? null, qty: l.qty })) };
}
// An Open/Complete row as InvCT gives it.
const crow = (sheet, n, o = {}) => ({ sheet, row: o.row ?? 5, id: '', containerText: (o.containers || []).join(' '),
  containers: o.containers || [], roundText: String(n), rounds: Array.isArray(n) ? n : [n], etd: o.etd ?? null, eta: o.eta ?? null });
function ctOf(blocks, rows = []) {
  const rounds = { open: {}, complete: {} };
  for (const r of rows) for (const k of r.rounds) { const t = rounds[r.sheet.toLowerCase()]; (t[k] = t[k] || []).push(r); }
  return { ok: true, error: null, blocks, rounds, issues: [], counts: {} };
}
const fwdOf = recs => ({ ok: true, byContainer: new Map(recs.map(r => [r.container, { ids: ['1'], rows: [3], mode: 'FCL', etd: null, eta: null, ...r }])) });
const drayOf = recs => ({ ok: true, records: recs.map(r => ({ eta: null, availability: 'NOT AVAILABLE', actualPickup: null, actualDelivery: null, ...r })) });
const run = (lines, ct, extra = {}) => ST.build({ po: { lines }, mrows: mrowsFor(lines), ct, today: TODAY, ...extra });
const units = (r, pred) => r.allocs.filter(pred).reduce((s, a) => s + a.units, 0);
const inBucket = (r, b) => units(r, a => a.bucket === b);
const codes = r => r.issues.map(i => i.code);
const codesBut = (r, skip = ['no-case-pack']) => codes(r).filter(c => !skip.includes(c));
const balanced = r => {
  assert.equal(r.balance.diff, 0, 'Master Sheet balance');
  assert.deepEqual(r.balance.bad, []);
  assert.equal(r.balance.ct.diff, 0, 'Container Tracking balance');
  assert.ok(!codes(r).includes('balance'));
};
// The stage one Container Tracking line gets, read through an unmatched allocation (no Master Sheet line).
function stageOf({ sheet = 'Loaded', open = true, etd = null, eta = null, containers = ['TEST1234567'], fwd, dray, rounds = [350] } = {}) {
  const b = blk(sheet, '900990', rounds, [{ style: 'ZZ1', qty: 10 }]);
  const rows = rounds.length ? [crow(open ? 'Open' : 'Complete', rounds[0], { containers, etd, eta })] : [];
  const r = run([], ctOf([b], rows), { forwarder: fwd ? fwdOf(fwd) : null, drayage: dray ? drayOf(dray) : null });
  return { stage: r.allocs[0].stage, a: r.allocs[0], r };
}

// ---------- not loaded, inputs

test('Master Sheet lines or Container Tracking missing: not loaded, nothing else', () => {
  const ct = ctOf([]);
  assert.deepEqual(ST.build({ po: null, ct, today: TODAY }), { ok: false, notLoaded: ['Master Sheet Live'] });
  assert.deepEqual(ST.build({ po: { lines: [] }, ct: null, today: TODAY }), { ok: false, notLoaded: ['Container Tracking'] });
  assert.deepEqual(ST.build({ po: null, ct: { ok: false }, today: TODAY }), { ok: false, notLoaded: ['Master Sheet Live', 'Container Tracking'] });
  assert.throws(() => ST.build({ po: { lines: [] }, ct }), /today/);
});

test('a line without status takes it from src (round = shipped); mrows can come from po.mrows', () => {
  const l1 = line({ status: undefined, src: 'round', qty: 100 }), l2 = line({ status: undefined, src: 'open', sty: 'AB101', qty: 50 });
  const r = ST.build({ po: { lines: [l1, l2], mrows: mrowsFor([l1, l2]) }, ct: ctOf([]), today: TODAY });
  assert.equal(inBucket(r, 'shippedNotOnCt'), 100);
  assert.equal(inBucket(r, 'noCrd'), 50);
  assert.ok(!codes(r).includes('missing-column'));
  balanced(r);
});

test('Master Sheet columns missing: that fact is blank and one issue each, nothing else stops', () => {
  const l = line({ master: 48, upc: '000111222333' });
  const rows = mrowsFor([l]).map(r => [r[0], r[3], r[6], r[7]]);   // Category, style, qty, PO: no Barcode, Master, Sub
  const r = ST.build({ po: { lines: [l] }, mrows: rows, ct: ctOf([]), today: TODAY });
  const miss = r.issues.filter(i => i.code === 'missing-column');
  assert.equal(miss.length, 3);
  assert.match(miss.map(i => i.message).join(' '), /"Barcode".*"Master".*"Sub"/s);
  assert.ok(!codes(r).includes('no-case-pack'), 'the missing Master column is one issue, not one per style');
  assert.equal(r.allocs[0].cases, null);
  assert.equal(r.allocs[0].upc, null);
  assert.equal(r.styles[0].category, 'Lip');
});

// ---------- keys and matching

test('style is exact: inner spaces and -F are different items, listed as near-misses, never matched', () => {
  const ls = [line({ po: '900201', sty: 'AB 100', qty: 500, crd: D(12, 1) }), line({ po: '900202', sty: 'AB200-F', qty: 300, crd: D(12, 1) })];
  const ct = ctOf([blk('Booking', '900201', [], [{ style: 'AB100', qty: 500 }]), blk('Booking', '900202', [], [{ style: 'AB200', qty: 300 }])]);
  const r = run(ls, ct);
  assert.equal(inBucket(r, 'inProd'), 800, 'the Master Sheet lines keep their own stage');
  assert.equal(inBucket(r, 'unmatched'), 800);
  assert.equal(r.totals.units.unmatched, 800);
  assert.equal(r.totals.units.booked, 0);
  const near = r.issues.filter(i => i.code === 'near-miss');
  assert.deepEqual(near.map(i => i.style).sort(), ['AB100', 'AB200']);
  assert.equal(r.counts.nearMisses, 2);
  assert.ok(r.allocs.filter(a => a.bucket === 'unmatched').every(a => a.flags.includes('near-miss') && a.nearMiss));
  balanced(r);
});

test('near-miss on the 6-digit PO: same style, the PO written another way; an exact PO with a suffix only matches itself', () => {
  const ls = [line({ po: '900301', sty: 'AB100', qty: 100 }), line({ po: '900302-BCF', sty: 'AB200', qty: 100 })];
  const ct = ctOf([blk('Booking', '900301A', [], [{ style: 'AB100', qty: 100 }]), blk('Booking', '900302-BCF', [], [{ style: 'AB200', qty: 100 }])]);
  const r = run(ls, ct);
  const near = r.issues.find(i => i.code === 'near-miss');
  assert.equal(near.po, '900301A');
  assert.deepEqual(near.near, { po: '900301', style: 'AB100' });
  assert.equal(units(r, a => a.bucket === 'booked' && a.style === 'AB200'), 100, '900302-BCF matches 900302-BCF');
  balanced(r);
});

test('candidate POs: the PO Book PO, then the original Master Sheet PO (renamed, split, derived, #n repeat)', () => {
  const renamed = line({ po: '900401-1', srcPo: '900401', renamed: true, bo: true, sty: 'AB100', qty: 100, key: '900401|AB100' });
  const parent = line({ po: '900402', sty: 'AB200', qty: 60, key: '900402|AB200' });
  const split = line({ po: '900402-1', sty: 'AB200', qty: 40, key: 'split-1', split: true, parentKey: '900402|AB200', bo: true, row: parent.row });
  const derived = line({ po: '900403-1', sty: 'AB300', qty: 70, key: '900403|AB300#2', derived: true, bo: true });
  const ct = ctOf([
    blk('Booking', '900401', [], [{ style: 'AB100', qty: 100 }]),     // the original PO
    blk('Booking', '900402-1', [], [{ style: 'AB200', qty: 40 }]),    // the split's PO Book PO
    blk('Booking', '900403', [], [{ style: 'AB300', qty: 70 }]),      // the key's PO part, without #2
  ]);
  const r = run([renamed, parent, split, derived], ct);
  assert.equal(inBucket(r, 'unmatched'), 0);
  assert.equal(inBucket(r, 'booked'), 210);
  assert.equal(r.counts.groups, 3, 'the split stays in its parent\'s group');
  const a = r.allocs.find(x => x.style === 'AB100');
  assert.equal(a.po, '900401');
  assert.equal(a.poShown, '900401-1');
  assert.equal(a.from.poShown, 'PO Book');
  balanced(r);
});

// ---------- the allocation steps

test('step 1: shipped lines take sailed units (landed ones leave every total)', () => {
  const ls = [shippedLine({ po: '900501', qty: 1000 })];
  const ct = ctOf([blk('Loaded', '900501', [340], [{ qty: 1000 }])], [crow('Complete', 340, { containers: ['TEST1111111'], etd: D(8, 1), eta: D(9, 1) })]);
  const r = run(ls, ct);
  assert.equal(inBucket(r, 'landed'), 1000);
  assert.equal(r.totals.units.incoming, 0);
  assert.deepEqual(r.issues, []);
  balanced(r);
});

test('step 1 + 2: sailed beyond shipped goes to open lines at its stage (MSL behind); beyond open is listed, not counted', () => {
  const ls = [shippedLine({ po: '900601', qty: 1000, row: 40 }), line({ po: '900601', qty: 500, row: 41, crd: D(12, 1) })];
  const ct = ctOf([blk('Loaded', '900601', [350], [{ qty: 1200 }]), blk('Loaded', '900601', [351], [{ qty: 400 }])],
    [crow('Open', 350, { containers: ['TEST2222222'], etd: D(10, 1), eta: D(11, 5) }), crow('Open', 351, { containers: ['TEST3333333'], etd: D(10, 1), eta: D(10, 20) })]);
  const r = run(ls, ct);
  // Round 351 is at port (ETA passed), 350 on the water. Shipped (1,000) takes the most advanced first: 400 port + 600 water.
  assert.equal(units(r, a => a.status === 'shipped' && a.bucket === 'port'), 400);
  assert.equal(units(r, a => a.status === 'shipped' && a.bucket === 'water'), 600);
  // The open line (500) takes the next 500 sailed, flagged; the last 100 are more than Master Sheet holds.
  assert.equal(units(r, a => a.status === 'open' && a.bucket === 'water' && a.flags.includes('msl-behind-ct')), 500);
  assert.equal(inBucket(r, 'ctMore'), 100);
  assert.equal(r.totals.units.water, 1100);
  assert.equal(r.totals.units.port, 400);
  assert.deepEqual(codesBut(r).sort(), ['ct-more', 'msl-behind-ct']);
  balanced(r);
});

test('step 3: shipped units Container Tracking still shows unsailed count On the water (CT behind MSL)', () => {
  const ls = [shippedLine({ po: '900701', qty: 800 })];
  const ct = ctOf([blk('Loaded', '900701', [352], [{ qty: 800 }])], [crow('Open', 352, { containers: ['TEST4444444'], etd: D(11, 10), eta: D(12, 10) })]);
  const r = run(ls, ct);
  const a = r.allocs.find(x => x.bucket === 'water');
  assert.equal(a.units, 800);
  assert.equal(a.stage, 'On the water');
  assert.equal(a.ctStage, 'Loaded');
  assert.ok(a.flags.includes('ct-behind-msl'));
  assert.equal(a.from.stage, 'Master Sheet Live');
  assert.deepEqual(codesBut(r), ['ct-behind-msl']);
  balanced(r);
});

test('step 4: shipped units left are short-shipped when the group has Container Tracking units, else treated as landed', () => {
  const ls = [
    shippedLine({ po: '900801', qty: 1000 }),
    shippedLine({ po: '900802', sty: 'AB200', qty: 300, crd: D(9, 15) }),          // recent: listed
    shippedLine({ po: '900803', sty: 'AB300', qty: 200, crd: D(3, 1), od: D(1, 5) }), // old: balance only
  ];
  const ct = ctOf([blk('Loaded', '900801', [340], [{ qty: 700 }])], [crow('Complete', 340, { etd: D(8, 1), eta: D(9, 1) })]);
  const r = run(ls, ct);
  assert.equal(inBucket(r, 'landed'), 700);
  assert.equal(inBucket(r, 'short'), 300);
  assert.equal(inBucket(r, 'shippedNotOnCt'), 500);
  assert.equal(r.totals.units.incoming, 0);
  const nc = r.issues.filter(i => i.code === 'shipped-not-on-ct');
  assert.deepEqual(nc.map(i => i.po), ['900802']);
  assert.equal(r.issues.filter(i => i.code === 'short').length, 1);
  balanced(r);
});

test('step 5: open lines take unsailed units, Loaded first, earliest CRD first (no CRD last); extra is listed', () => {
  const ls = [
    line({ po: '900901', qty: 300, crd: null, row: 60 }),
    line({ po: '900901', qty: 300, crd: D(12, 20), row: 61 }),
    line({ po: '900901', qty: 300, crd: D(11, 5), row: 62 }),
  ];
  const ct = ctOf([blk('Booking', '900901', [], [{ qty: 500 }]), blk('Loaded', '900901', [353], [{ qty: 450 }])],
    [crow('Open', 353, { containers: ['TEST5555555'], etd: D(11, 20), eta: D(12, 20) })]);
  const r = run(ls, ct);
  const of = row => r.allocs.filter(a => a.sheetRow === row).map(a => `${a.bucket} ${a.units}`).sort();
  assert.deepEqual(of(62), ['loaded 300'], 'earliest CRD gets Loaded first');
  assert.deepEqual(of(61), ['booked 150', 'loaded 150']);
  assert.deepEqual(of(60), ['booked 300'], 'no CRD goes last');
  assert.equal(inBucket(r, 'ctMore'), 50);
  assert.equal(r.allocs.find(a => a.bucket === 'ctMore').stage, 'Booked');
  assert.ok(codes(r).includes('ct-more'));
  balanced(r);
});

test('step 6: open units left keep their Master Sheet stage; CRD today is In production, yesterday Ready', () => {
  const ls = [
    line({ po: '901001', sty: 'AB1', qty: 10, crd: null }),
    line({ po: '901001', sty: 'AB2', qty: 20, crd: TODAY }),
    line({ po: '901001', sty: 'AB3', qty: 30, crd: TODAY - 1 }),
    line({ po: '901001', sty: 'AB4', qty: 40, crd: TODAY + 1 }),
  ];
  const r = run(ls, ctOf([]));
  assert.deepEqual({ noCrd: r.totals.units.noCrd, inProd: r.totals.units.inProd, ready: r.totals.units.ready }, { noCrd: 10, inProd: 60, ready: 30 });
  assert.equal(r.totals.units.inChina, 100);
  assert.equal(r.totals.units.incoming, 100);
  assert.equal(r.allocs.find(a => a.style === 'AB3').stage, 'Ready');
  balanced(r);
});

test('a cap keeps the most advanced stage first: Landed, then At port, then On the water', () => {
  const ls = [line({ po: '901101', qty: 500 })];
  const ct = ctOf([
    blk('Loaded', '901101', [360], [{ qty: 300 }]), blk('Loaded', '901101', [361], [{ qty: 300 }]),
  ], [crow('Open', 360, { containers: ['TEST6666666'], etd: D(10, 1), eta: D(11, 30) }), crow('Complete', 361, { etd: D(9, 1), eta: D(10, 1) })]);
  const r = run(ls, ct);
  assert.equal(units(r, a => a.bucket === 'landed'), 300);
  assert.equal(units(r, a => a.bucket === 'water'), 200);
  assert.equal(units(r, a => a.bucket === 'ctMore' && a.stage === 'On the water'), 100);
  balanced(r);
});

// ---------- China Pickup, unknown type, unmatched

test('China Pickup: every unit goes to Pull from China whatever its stage, which is still recorded', () => {
  const ls = [line({ po: '901201', ptype: 'China Pickup', qty: 1000, crd: D(12, 1) }), shippedLine({ po: '901201', ptype: 'China Pickup', qty: 400, row: 90 })];
  const ct = ctOf([blk('Loaded', '901201', [362], [{ qty: 600 }])], [crow('Open', 362, { containers: ['TEST7777777'], etd: D(10, 1), eta: D(11, 20) })]);
  const r = run(ls, ct);
  assert.equal(r.totals.units.pullFromChina, 1400, 'shipped on the water 400 + open on the water 200 + open In production 800');
  assert.equal(r.totals.units.incoming, 0);
  assert.equal(r.totals.units.inChina, 0);
  const stages = r.allocs.filter(a => a.bucket === 'pullFromChina').map(a => a.stage).sort();
  assert.deepEqual(stages, ['In production', 'On the water', 'On the water']);
  const p = r.pull.find(x => x.ptype === 'China Pickup');
  assert.equal(p.ctStatus, 'On the water 200 · Not on tracker 800');
  balanced(r);
});

test('China Pickup shipped and not on Container Tracking is not listed (it never goes on it)', () => {
  const r = run([shippedLine({ po: '901251', ptype: 'China Pickup', qty: 100, crd: D(10, 1) })], ctOf([]));
  assert.equal(inBucket(r, 'shippedNotOnCt'), 100);
  assert.deepEqual(r.issues, []);
});

test('PO type unknown (not on the PO Status Tracker): counted as stock like LA, listed once per PO', () => {
  const ls = [line({ po: '901301', sty: 'AB1', tracked: false, qty: 10 }), line({ po: '901301', sty: 'AB2', tracked: false, qty: 20 })];
  const r = run(ls, ctOf([]));
  assert.equal(r.totals.units.noCrd, 30);
  assert.ok(r.allocs.every(a => a.ptype === 'Unknown' && a.flags.includes('unknown-type')));
  const u = r.issues.filter(i => i.code === 'unknown-type');
  assert.equal(u.length, 1);
  assert.match(u[0].message, /2 styles.*30 units/);
  assert.equal(r.pull.every(p => p.ptype === 'Unknown'), true);
});

test('unmatched Container Tracking units count as Unmatched by stage; landed ones are not counted or listed', () => {
  const ct = ctOf([
    blk('Loaded', '901401', [370], [{ style: 'ZZ1', qty: 100 }]),
    blk('Loaded', '901402', [371], [{ style: 'ZZ2', qty: 200 }]),
  ], [crow('Open', 370, { containers: ['TEST8888888'], etd: D(10, 1), eta: D(11, 20) }), crow('Complete', 371, { etd: D(8, 1), eta: D(9, 1) })]);
  const r = run([], ct);
  assert.equal(r.totals.units.unmatched, 100);
  assert.equal(r.counts.unmatchedLandedUnits, 200);
  assert.deepEqual(r.issues.filter(i => i.code !== 'no-case-pack').map(i => [i.code, i.po]), [['unmatched', '901401']]);
  assert.deepEqual(r.styles.map(s => s.style), ['ZZ1'], 'a landed unmatched style shows nowhere');
  assert.ok(r.allocs.some(a => a.bucket === 'unmatched' && a.stage === 'Landed'), 'but its allocation is kept for the balance');
  balanced(r);
});

// ---------- Container Tracking stages

test('the same PO + style + round on Booking and Loaded counts once, from Loaded', () => {
  const ls = [line({ po: '901501', qty: 1000 })];
  const ct = ctOf([blk('Booking', '901501', [380], [{ qty: 1000 }]), blk('Loaded', '901501', [380], [{ qty: 800 }])],
    [crow('Open', 380, { containers: ['TEST9999999'], etd: D(11, 10), eta: D(12, 10) })]);
  const r = run(ls, ct);
  assert.equal(r.totals.units.loaded, 800);
  assert.equal(r.totals.units.booked, 0);
  assert.equal(r.counts.bothTabs, 1);
  assert.equal(r.counts.bothTabsUnits, 1000);
  assert.equal(r.issues.filter(i => i.code === 'both-tabs').length, 1);
  assert.equal(inBucket(r, 'noCrd'), 200);
  balanced(r);
});

test('a round takes its least advanced container; a block with several rounds its least advanced round', () => {
  // Round 390: two rows (two containers), one at port, one on the water.
  const two = blk('Loaded', '901601', [390], [{ style: 'ZZ1', qty: 10 }]);
  const several = blk('Loaded', '901602', [391, 392], [{ style: 'ZZ2', qty: 10 }]);
  const rows = [
    crow('Open', 390, { containers: ['TEST0000001'], etd: D(9, 1), eta: D(10, 1), row: 7 }),
    crow('Open', 390, { containers: ['TEST0000002'], etd: D(10, 1), eta: D(11, 20), row: 8 }),
    crow('Complete', 391, { etd: D(8, 1), eta: D(9, 1) }),
    crow('Open', 392, { containers: ['TEST0000003'], etd: D(11, 20), eta: D(12, 20) }),
  ];
  const r = run([], ctOf([two, several], rows));
  const st = Object.fromEntries(r.allocs.map(a => [a.style, a.stage]));
  assert.deepEqual(st, { ZZ1: 'On the water', ZZ2: 'Loaded' });
  const a = r.allocs.find(x => x.style === 'ZZ1');
  assert.deepEqual(a.containers, ['TEST0000001', 'TEST0000002']);
  assert.deepEqual(a.ctRows.map(x => x.sheet + ' ' + x.row), [`Loaded ${two.lines[0].row}`, 'Open 7', 'Open 8']);
});

test('a round on both Open and Complete: Open wins', () => {
  const b = blk('Loaded', '901651', [393], [{ style: 'ZZ1', qty: 10 }]);
  const r = run([], ctOf([b], [crow('Complete', 393, { etd: D(8, 1), eta: D(9, 1) }), crow('Open', 393, { etd: D(10, 1), eta: D(11, 20) })]));
  assert.equal(r.allocs[0].stage, 'On the water');
  assert.ok(r.allocs[0].flags.includes('round-open-and-complete'));
});

test('report evidence per container, most advanced first', () => {
  const c = 'TEST1234567';
  assert.equal(stageOf({ dray: [{ container: c, actualDelivery: D(10, 20) }] }).stage, 'Landed', 'drayage delivered, even on Open');
  assert.equal(stageOf({ dray: [{ container: c, availability: 'DELIVERED' }] }).stage, 'Landed');
  assert.equal(stageOf({ dray: [{ container: c, actualPickup: D(10, 25) }] }).stage, 'At port');
  assert.equal(stageOf({ dray: [{ container: c, availability: 'AVAILABLE' }] }).stage, 'At port');
  assert.equal(stageOf({ fwd: [{ container: c, status: 'Arrived' }] }).stage, 'At port');
  assert.equal(stageOf({ fwd: [{ container: c, status: 'Sailed' }] }).stage, 'On the water');
  assert.equal(stageOf({ fwd: [{ container: c, status: 'Soon To Arrive' }] }).stage, 'On the water');
  const pd = stageOf({ etd: D(10, 1), eta: D(11, 20), fwd: [{ container: c, status: 'Pending Departure', etd: D(10, 30), eta: D(12, 1) }] });
  assert.equal(pd.stage, 'Loaded', 'Pending Departure beats a passed ETD (a rolled vessel)');
  assert.equal(pd.a.etd, D(10, 30), "the forwarder's ETD is shown when it covers the container");
  assert.equal(pd.a.from.etd, 'Forwarder');
  assert.equal(pd.a.from.stage, 'Forwarder');
  // Drayage beats the forwarder; a drayage row that is only "not available" decides nothing.
  assert.equal(stageOf({ fwd: [{ container: c, status: 'Sailed' }], dray: [{ container: c, availability: 'AVAILABLE' }] }).stage, 'At port');
  assert.equal(stageOf({ etd: D(10, 1), eta: D(11, 20), dray: [{ container: c, availability: 'NOT AVAILABLE' }] }).stage, 'On the water');
});

test('no report covers it: Complete = Landed, ETA passed = At port, ETD passed = On the water, else not sailed', () => {
  assert.equal(stageOf({ open: false }).stage, 'Landed');
  assert.equal(stageOf({ etd: D(9, 1), eta: D(10, 27) }).stage, 'At port');
  assert.equal(stageOf({ etd: D(10, 27), eta: D(11, 27) }).stage, 'On the water');
  assert.equal(stageOf({ etd: TODAY, eta: D(11, 27) }).stage, 'Loaded', 'an ETD of today has not passed yet');
  assert.equal(stageOf({ sheet: 'Booking', etd: D(11, 2), eta: D(12, 2) }).stage, 'Booked');
  const noCont = stageOf({ containers: [], etd: D(10, 1), eta: D(10, 20) });
  assert.equal(noCont.stage, 'At port', "a row with no container uses the row's own dates");
  assert.equal(noCont.a.from.eta, 'Container Tracking');
});

test('no round, or a round with no Open/Complete row: not sailed, no ETD, listed', () => {
  const r1 = stageOf({ sheet: 'Booking', rounds: [] });
  assert.equal(r1.stage, 'Booked');
  assert.equal(r1.a.etd, null);
  assert.ok(r1.r.issues.some(i => i.code === 'no-etd' && /not in a round/.test(i.message)));
  const b = blk('Loaded', '901701', [394], [{ style: 'ZZ1', qty: 10 }]);
  const r2 = run([], ctOf([b]));
  assert.equal(r2.allocs[0].stage, 'Loaded');
  assert.ok(r2.issues.some(i => i.code === 'no-etd' && /round 394 has no row/.test(i.message)));
  assert.equal(r2.counts.noEtd, 1);
});

test('on Booking with its ETD passed and no report saying it sailed: listed', () => {
  const t = stageOf({ sheet: 'Booking', etd: D(10, 1), eta: D(11, 20) });
  assert.equal(t.stage, 'On the water', 'no report covers it: the ETD decides');
  assert.ok(t.r.issues.some(i => i.code === 'etd-passed-booked'));
  const rolled = stageOf({ sheet: 'Booking', etd: D(10, 1), eta: D(11, 20), fwd: [{ container: 'TEST1234567', status: 'Pending Departure', eta: D(11, 25) }] });
  assert.equal(rolled.stage, 'Booked');
  assert.ok(rolled.r.issues.some(i => i.code === 'etd-passed-booked'));
  const sailed = stageOf({ sheet: 'Booking', etd: D(10, 1), eta: D(11, 20), fwd: [{ container: 'TEST1234567', status: 'Sailed', eta: D(11, 20) }] });
  assert.ok(!sailed.r.issues.some(i => i.code === 'etd-passed-booked'));
});

test('forwarder or drayage missing: stages from Container Tracking dates only, those blocks counted', () => {
  const { r } = stageOf({ etd: D(10, 1), eta: D(11, 20) });
  assert.deepEqual(r.counts.reportsMissing, ['Forwarder', 'Drayage']);
  assert.equal(r.counts.trackerOnly, 1);
  const withReport = stageOf({ etd: D(10, 1), eta: D(11, 20), fwd: [{ container: 'TEST1234567', status: 'Sailed', eta: D(11, 20) }], dray: [] });
  assert.deepEqual(withReport.r.counts.reportsMissing, []);
  assert.equal(withReport.r.counts.trackerOnly, 0);
  assert.equal(stageOf({ open: false }).r.counts.trackerOnly, 0, 'a landed block is not one of them');
});

test('a report trip far from the tracker row is another trip of a reused container: not used, listed', () => {
  const t = stageOf({ open: false, etd: D(3, 1), eta: D(4, 1), dray: [{ container: 'TEST1234567', availability: 'AVAILABLE', eta: D(10, 20) }] });
  assert.equal(t.stage, 'Landed');
  assert.ok(t.r.issues.some(i => i.code === 'other-trip'));
  assert.equal(t.r.counts.otherTrip, 1);
});

test('a round on Complete is Landed whatever the reports say; a report saying it has not arrived is listed once', () => {
  const c = 'TEST1234567';
  for (const status of ['Arrived', 'Sailed', 'Soon To Arrive', 'Pending Departure']) {
    const t = stageOf({ open: false, etd: D(10, 1), eta: D(10, 20), fwd: [{ container: c, status, etd: D(10, 2), eta: D(10, 22) }] });
    assert.equal(t.stage, 'Landed', status);
    assert.equal(t.a.eta, D(10, 22), "the forwarder's ETA is still the one shown");
    assert.equal(t.a.from.stage, 'Container Tracking');
    assert.equal(t.r.totals.units.incoming, 0);
    const iss = t.r.issues.filter(i => i.code === 'complete-in-transit');
    if (status === 'Arrived') {   // the forwarder keeps Arrived shipments for days after delivery: normal
      assert.equal(iss.length, 0);
      assert.ok(!t.a.flags.includes('complete-in-transit'));
      continue;
    }
    assert.ok(t.a.flags.includes('complete-in-transit'));
    assert.equal(iss.length, 1, status);
    assert.match(iss[0].message, new RegExp(`Round 350 .*Complete.*forwarder says ${status}.*${c}`));
    assert.equal(t.r.counts.completeInTransit, 1);
  }
  // Arrived, but drayage has it not delivered: listed.
  const ar = stageOf({ open: false, etd: D(10, 1), eta: D(10, 20), fwd: [{ container: c, status: 'Arrived', eta: D(10, 22) }], dray: [{ container: c, eta: D(10, 22) }] });
  assert.match(ar.r.issues.find(i => i.code === 'complete-in-transit').message, /forwarder says Arrived .*drayage says not available/);
  // Drayage covering it without a delivery is named too; drayage delivered agrees, nothing listed.
  const nd = stageOf({ open: false, etd: D(10, 1), eta: D(10, 20), fwd: [{ container: c, status: 'Soon To Arrive', eta: D(10, 22) }], dray: [{ container: c, eta: D(10, 22) }] });
  assert.match(nd.r.issues.find(i => i.code === 'complete-in-transit').message, /drayage says not available/);
  const del = stageOf({ open: false, etd: D(10, 1), eta: D(10, 20), dray: [{ container: c, availability: 'DELIVERED', eta: D(10, 22) }] });
  assert.equal(del.stage, 'Landed');
  assert.equal(del.a.from.stage, 'Drayage');
  assert.ok(!del.r.issues.some(i => i.code === 'complete-in-transit'));
  // A Complete round on Booking whose ETD passed isn't "still on Booking, ETD passed".
  const bk = stageOf({ sheet: 'Booking', open: false, etd: D(10, 1), eta: D(10, 20) });
  assert.ok(!bk.r.issues.some(i => i.code === 'etd-passed-booked'));
  // Two blocks on the same round: still one issue for its Complete row.
  const b1 = blk('Loaded', '900991', [350], [{ style: 'ZZ1', qty: 10 }]), b2 = blk('Booking', '900992', [350], [{ style: 'ZZ2', qty: 10 }]);
  const r = run([], ctOf([b1, b2], [crow('Complete', 350, { containers: [c], etd: D(10, 1), eta: D(10, 20) })]), { forwarder: fwdOf([{ container: c, status: 'Sailed', eta: D(10, 22) }]) });
  assert.equal(r.issues.filter(i => i.code === 'complete-in-transit').length, 1);
});

test('drayage on a row with no ETA yet: a port date before its ETD, or long after, is another trip', () => {
  const c = 'TEST1234567';
  const old = stageOf({ etd: D(11, 15), dray: [{ container: c, availability: 'DELIVERED', actualDelivery: D(6, 20), eta: D(6, 15) }] });
  assert.equal(old.stage, 'Loaded', 'an old delivery of a reused container number');
  assert.ok(old.r.issues.some(i => i.code === 'other-trip' && /Drayage/.test(i.message)));
  const late = stageOf({ etd: D(5, 1), dray: [{ container: c, availability: 'AVAILABLE', eta: D(10, 20) }] });
  assert.equal(late.stage, 'On the water', 'more than 90 days after the ETD: not this trip');
  const ok = stageOf({ etd: D(9, 20), dray: [{ container: c, availability: 'DELIVERED', actualDelivery: D(10, 25), eta: D(10, 20) }] });
  assert.equal(ok.stage, 'Landed', 'within the window: this trip');
  assert.equal(stageOf({ dray: [{ container: c, availability: 'AVAILABLE', eta: D(10, 20) }] }).stage, 'At port', 'no dates on the row: taken as this trip');
});

test('a yearless tracker date read a year off is still the same trip as the report: the report decides', () => {
  const c = 'TEST1234567', b = blk('Booking', '900993', [990], [{ style: 'ZZ1', qty: 100 }]);
  // "ETD 1/8, ETA 2/12" typed in late December: read as this year, the forwarder has next year's dates.
  const row = { ...crow('Open', 990, { containers: [c], etd: D(1, 8), eta: D(2, 12) }), etdYearless: true, etaYearless: true };
  const today = D(12, 28);
  const fw = fwdOf([{ container: c, status: 'Pending Departure', etd: D(1, 8, 2027), eta: D(2, 12, 2027) }]);
  const r = ST.build({ po: { lines: [] }, mrows: mrowsFor([]), ct: ctOf([b], [row]), forwarder: fw, today });
  assert.equal(r.allocs[0].stage, 'Booked');
  assert.equal(r.allocs[0].eta, D(2, 12, 2027));
  assert.ok(!r.issues.some(i => i.code === 'other-trip'));
  // Without the yearless mark the same dates are another trip.
  const r2 = ST.build({ po: { lines: [] }, mrows: mrowsFor([]), ct: ctOf([b], [{ ...row, etdYearless: false, etaYearless: false }]), forwarder: fw, today });
  assert.ok(r2.issues.some(i => i.code === 'other-trip'));
});

test('a shipped line the PO Book marks as a repeat of a line inside a Round: its own outcome, not short-shipped', () => {
  const ls = [shippedLine({ po: '900851', qty: 600, row: 70 }), shippedLine({ po: '900851', qty: 600, row: 90, key: '900851|AB100#2', dupRound: true }),
    shippedLine({ po: '900852', sty: 'AB200', qty: 500, row: 71 })];
  const ct = ctOf([blk('Loaded', '900851', [340], [{ qty: 600 }]), blk('Loaded', '900852', [340], [{ style: 'AB200', qty: 400 }])],
    [crow('Complete', 340, { etd: D(8, 1), eta: D(9, 1) })]);
  const r = run(ls, ct);
  assert.equal(inBucket(r, 'duplicate'), 600);
  assert.ok(r.allocs.find(a => a.bucket === 'duplicate').flags.includes('dup-round'));
  assert.equal(inBucket(r, 'short'), 100, 'a real shortfall is still short-shipped');
  assert.deepEqual(r.issues.filter(i => i.code === 'short').map(i => i.po), ['900852']);
  assert.equal(r.counts.duplicateUnits, 600);
  assert.equal(r.totals.units.incoming, 0);
  balanced(r);
  // When Container Tracking has both, the repeat takes the second shipment like any shipped line.
  const r2 = run(ls.slice(0, 2), ctOf([blk('Loaded', '900851', [340], [{ qty: 1200 }])], [crow('Complete', 340, { etd: D(8, 1), eta: D(9, 1) })]));
  assert.equal(inBucket(r2, 'landed'), 1200);
  assert.equal(inBucket(r2, 'duplicate'), 0);
});

test('shipped and not on Container Tracking, or short: names Container Tracking units of the same order written another way', () => {
  const ls = [shippedLine({ po: '900861', sty: 'AB15-F', qty: 100, crd: D(10, 1) }), shippedLine({ po: '900862-1', sty: 'AB16', qty: 100, crd: D(10, 1) })];
  const ct = ctOf([blk('Loaded', '900861', [341], [{ style: 'AB15', qty: 100 }]), blk('Loaded', '900862', [341], [{ style: 'AB16', qty: 100 }])],
    [crow('Complete', 341, { etd: D(9, 1), eta: D(10, 1) })]);
  const r = run(ls, ct);
  const snc = r.issues.filter(i => i.code === 'shipped-not-on-ct');
  assert.equal(snc.length, 2);
  for (const i of snc) {
    assert.doesNotMatch(i.message, /nowhere on Container Tracking/);
    assert.match(i.message, /Container Tracking does have 100 units of style AB1[56] on PO 90086[12] \(Landed\)/);
  }
  assert.equal(inBucket(r, 'shippedNotOnCt'), 200);
  balanced(r);
});

test('Container Tracking behind Master Sheet: says when the forwarder\'s Pending Departure kept it unsailed', () => {
  const c = 'TEST1234567';
  const ls = [shippedLine({ po: '900871', qty: 800 })];
  const ct = ctOf([blk('Loaded', '900871', [352], [{ qty: 800 }])], [crow('Open', 352, { containers: [c], etd: D(10, 26), eta: D(11, 26) })]);
  const r = run(ls, ct, { forwarder: fwdOf([{ container: c, status: 'Pending Departure', etd: D(10, 27), eta: D(11, 27) }]) });
  const i = r.issues.find(x => x.code === 'ct-behind-msl');
  assert.match(i.message, /forwarder still says Pending Departure \(ETD 10\/27\/2026\) for 800 units \(round 352\).*rolled/);
  assert.doesNotMatch(i.message, /Container Tracking still shows/);
  assert.equal(r.totals.units.water, 800);
  // No forwarder and an ETD still ahead: Container Tracking's own dates kept it unsailed.
  const r2 = run(ls, ctOf([blk('Loaded', '900871', [352], [{ qty: 800 }])], [crow('Open', 352, { containers: [c], etd: D(11, 10), eta: D(12, 10) })]));
  assert.match(r2.issues.find(x => x.code === 'ct-behind-msl').message, /Container Tracking still shows 800 units as Loaded/);
});

// ---------- Container Tracking sections with no round (config.js's sections) and POs below the tracker's lowest PO

// Invented section names, written the way config.js writes them (matched trimmed, any case, spaces as one).
const SECT = { pullFromChina: ['Direct  A', 'Direct B'], planning: ['Plan X'], ignore: ['Hidden Test'] };
const runS = (lines, ct, extra = {}) => run(lines, ct, { sections: SECT, ...extra });
const sec = (sheet, po, name, lns) => blk(sheet, po, [], lns, { section: name });

test('sections: a direct-to-customer section goes to Pull from China, capped by the group; status "Loaded · <section>"', () => {
  const ls = [line({ po: '900501', qty: 1000, crd: D(12, 1) }), line({ po: '900502', sty: 'AB200', qty: 5000, crd: D(12, 1) }),
    shippedLine({ po: '900503', sty: 'AB300', qty: 400, crd: D(10, 1) })];
  const ct = ctOf([sec('Loaded', '900501', ' DIRECT a ', [{ qty: 1200 }]), sec('Booking', '900502', 'Direct B', [{ style: 'AB200', qty: 3000 }]),
    sec('Loaded', '900503', 'Direct A', [{ style: 'AB300', qty: 400 }])]);
  const r = runS(ls, ct);
  const t = r.totals.units;
  assert.equal(t.pullFromChina, 1000 + 3000 + 400, 'LA POs too: never In China or Incoming');
  assert.equal(t.loaded + t.booked + t.water, 0);
  assert.equal(t.inProd, 2000, "the rest of 900502 keeps its Master Sheet stage");
  const more = r.allocs.filter(a => a.bucket === 'ctMore');
  assert.deepEqual(more.map(a => [a.po, a.units, a.section]), [['900501', 200, 'DIRECT a']], 'beyond the group: listed, not counted');
  assert.match(r.issues.find(i => i.code === 'ct-more').message, /200 units more than Master Sheet Live \(under "DIRECT a"\)/);
  const p1 = r.allocs.find(a => a.po === '900501' && a.bucket === 'pullFromChina');
  assert.equal(p1.stage, 'Loaded'); assert.equal(p1.section, 'DIRECT a'); assert.ok(p1.flags.includes('pull-section'));
  assert.ok(!p1.flags.includes('no-etd'));
  assert.equal(r.pull.find(p => p.po === '900501').ctStatus, 'Loaded · DIRECT a');
  assert.equal(r.pull.find(p => p.po === '900502').ctStatus, 'Booked (Direct B) 3,000 · Not on tracker 2,000');
  assert.deepEqual(r.pull.find(p => p.po === '900502').ctParts.map(p => [p.status, p.section, p.units]), [['Booked', 'Direct B', 3000], ['Not on tracker', '', 2000]]);
  // A shipped line: its units stay Pull from China (like a China Pickup PO), said so.
  const sh = r.allocs.find(a => a.po === '900503');
  assert.deepEqual([sh.bucket, sh.stage, sh.section, sh.flags.includes('ct-behind-msl')], ['pullFromChina', 'Loaded', 'Direct A', true]);
  assert.match(r.issues.find(i => i.code === 'ct-behind-msl').message, /400 units under "Direct A", which config\.js lists as going straight to the customer.*counted as Pull from China/);
  assert.equal(r.counts.pullSectionUnits, 1200 + 3000 + 400);
  assert.ok(!codes(r).some(c => c === 'no-etd' || c === 'section-not-listed'), codes(r).join(','));
  balanced(r);
});

test('sections: a planning section leaves the Master Sheet stage; its units are "planning", in no total and never matched', () => {
  const ls = [line({ po: '900601', qty: 1000, crd: D(12, 1) }), line({ po: '900602', sty: 'AB200', qty: 300, crd: null }),
    line({ po: '900603', sty: 'AB300', qty: 200, crd: D(9, 1) })];
  const ct = ctOf([sec('Booking', '900601', 'Plan X', [{ qty: 1000 }]), blk('Booking', '900601', [352], [{ qty: 400 }]),
    sec('Loaded', '900602', 'plan  x', [{ style: 'AB200', qty: 300 }]), sec('Booking', '900603', 'Plan X', [{ style: 'AB300', qty: 200 }]),
    sec('Booking', '900690', 'Plan X', [{ style: 'ZZ1', qty: 50 }])], [crow('Open', 352, { etd: D(11, 10), eta: D(12, 10) })]);
  const r = runS(ls, ct);
  const t = r.totals.units;
  assert.equal(t.booked, 400, 'only the round counts as Booked');
  assert.equal(t.inProd, 600); assert.equal(t.noCrd, 300); assert.equal(t.ready, 200);
  assert.equal(t.unmatched, 0, 'a planning line with no Master Sheet line is planning, not Unmatched');
  assert.equal(inBucket(r, 'planning'), 1000 + 300 + 200 + 50);
  assert.equal(r.counts.planningUnits, 1550);
  const pa = r.allocs.find(a => a.bucket === 'planning' && a.po === '900602');
  assert.deepEqual([pa.stage, pa.section, pa.flags.includes('planning-section')], ['Planned', 'plan  x', true]);
  const nc = r.allocs.find(a => a.po === '900602' && a.bucket === 'noCrd');
  assert.ok(nc.flags.includes('planning-section'));
  assert.equal(nc.section, 'plan  x', "the line's side panel names the planning section (No CRD · plan  x)");
  assert.equal(r.allocs.find(a => a.po === '900601' && a.bucket === 'inProd').section, 'Plan X');
  assert.equal(r.pull.find(p => p.po === '900602').ctStatus, 'Planned · plan  x');
  assert.equal(r.pull.find(p => p.po === '900601').ctStatus, 'Booked 400 · Planned (Plan X) 600');
  assert.ok(!codes(r).some(c => ['no-etd', 'section-not-listed', 'unmatched', 'ct-more'].includes(c)), codes(r).join(','));
  balanced(r);
});

test('sections: a shipped line whose only Container Tracking lines are planning lines is shipped, not on Container Tracking', () => {
  const ls = [shippedLine({ po: '900701', qty: 500, crd: D(10, 1) })];
  const r = runS(ls, ctOf([sec('Loaded', '900701', 'Plan X', [{ qty: 500 }])]));
  assert.equal(inBucket(r, 'shippedNotOnCt'), 500);
  assert.equal(inBucket(r, 'short'), 0);
  assert.equal(inBucket(r, 'planning'), 500);
  assert.match(r.issues.find(i => i.code === 'shipped-not-on-ct').message, /on Container Tracking only under "Plan X", which config\.js lists as planned, not booked yet/);
  const sn = r.allocs.find(a => a.bucket === 'shippedNotOnCt');
  assert.deepEqual(sn.flags, ['shipped-not-on-ct'], 'a shipped line is never flagged "planned, not booked yet"');
  balanced(r);
});

test('sections: an ignored section is in no total and no stage, with one Problems line', () => {
  const ls = [line({ po: '900801', qty: 100, crd: D(12, 1) })];
  const ct = ctOf([sec('Booking', '900801', 'hidden test', [{ qty: 100 }, { style: 'AB101', qty: 20 }]), sec('Loaded', '900802', 'Hidden Test', [{ style: 'ZZ2', qty: 30 }])]);
  const r = runS(ls, ct);
  assert.equal(r.totals.units.inProd, 100);
  assert.equal(r.totals.units.booked + r.totals.units.unmatched, 0);
  assert.equal(inBucket(r, 'ignored'), 150);
  const iss = r.issues.filter(i => i.code === 'ignored-section');
  assert.equal(iss.length, 1);
  assert.match(iss[0].message, /"hidden test" section is ignored, as config\.js says: 3 lines on 2 POs \(900801, 900802\), 150 units/, 'one name however it is cased');
  assert.ok(r.allocs.filter(a => a.bucket === 'ignored').every(a => a.flags.includes('ignored-section') && a.section));
  assert.equal(r.counts.ignoredUnits, 150);
  assert.ok(!codes(r).some(c => c === 'no-etd' || c === 'unmatched'), codes(r).join(','));
  balanced(r);
});

test("below the PO Status Tracker's lowest PO: not counted, one Problems line, before the section rules; skipped when unknown", () => {
  const c = 'TEST4444444';
  const blocks = [sec('Loaded', '800001', 'Direct A', [{ style: 'OLD1', qty: 500 }]), blk('Loaded', '800002', [350], [{ style: 'OLD2', qty: 300 }]),
    sec('Booking', '800003', 'Plan X', [{ style: 'OLD3', qty: 200 }]), blk('Loaded', '900881', [], [{ style: 'ZZ1', qty: 300 }], { section: null })];
  const rows = [crow('Complete', 350, { containers: [c], etd: D(8, 1), eta: D(9, 1) })];
  const fwd = fwdOf([{ container: c, status: 'Sailed', eta: D(9, 2) }]);   // would list complete-in-transit if it were staged
  const b = (trk, sections = SECT) => ST.build({ po: { lines: [], trk }, mrows: mrowsFor([]), ct: ctOf(blocks, rows), forwarder: fwd, today: TODAY, sections });
  const r = b({ minPO: 900000 });
  assert.equal(inBucket(r, 'belowTracker'), 1000, 'a direct-to-customer or planning line below the lowest PO is below the tracker');
  assert.equal(inBucket(r, 'pullFromChina') + inBucket(r, 'planning'), 0);
  assert.equal(r.totals.units.unmatched, 300, 'only the PO above the lowest PO is Unmatched');
  const iss = r.issues.filter(i => i.code === 'below-tracker');
  assert.equal(iss.length, 1);
  assert.match(iss[0].message, /3 lines on 3 POs below the PO Status Tracker's lowest PO \(900000\): 1,000 units, 700 of them not landed/);
  assert.ok(!r.issues.some(i => /^8000/.test(i.po)), 'no issue per PO below it');
  assert.ok(!codes(r).includes('complete-in-transit') && !codes(r).includes('other-trip'), codes(r).join(','));
  assert.equal(r.allocs.find(a => a.po === '800002').stage, 'Landed', 'its stage is still recorded');
  assert.deepEqual([r.counts.belowTrackerLines, r.counts.belowTrackerUnits, r.counts.belowTrackerNotLandedUnits], [3, 1000, 700]);
  balanced(r);
  // The tracker's lowest PO unknown: no such rule; the sections decide.
  for (const trk of [null, { minPO: Infinity }]) {
    const u = b(trk);
    assert.equal(inBucket(u, 'belowTracker'), 0);
    assert.equal(inBucket(u, 'planning'), 200);
    assert.equal(u.totals.units.unmatched, 500 + 300, 'unmatched direct-to-customer units stay Unmatched; landed ones are left out');
    assert.ok(!codes(u).includes('below-tracker'));
    balanced(u);
  }
});

test('a section with no round not on the lists: Booked/Loaded, one Problems line per section; before any section: no ETD as before', () => {
  const ls = [line({ po: '900901', qty: 600, crd: D(12, 1) })];
  const ct = ctOf([sec('Booking', '900901', 'Odd One', [{ qty: 600 }]), sec('Booking', '900902', 'Odd One', [{ style: 'ZZ1', qty: 100 }]),
    sec('Loaded', '900903', ' odd  ONE', [{ style: 'ZZ2', qty: 50 }]), sec('Loaded', '900905', 'Second Odd', [{ style: 'ZZ4', qty: 5 }]),
    blk('Booking', '900904', [], [{ style: 'ZZ3', qty: 10 }], { section: null }), sec('Booking', '900906', 'Round TBD', [{ style: 'ZZ5', qty: 7 }])]);
  const r = runS(ls, ct);
  assert.equal(r.totals.units.booked, 600);
  const nl = r.issues.filter(i => i.code === 'section-not-listed');
  assert.equal(nl.length, 3, 'one per section name');
  assert.equal(nl[0].message, "Section 'Odd One' on Booking and Loaded has no round and isn't in config.js's section list: counted as Booked/Loaded until it's added (3 POs, 750 units).");
  assert.match(nl[1].message, /^Section 'Second Odd' on Loaded /);
  assert.match(nl[2].message, /^Section 'Round TBD' on Booking /, 'a header that reads like a round with no number is a section too');
  const ne = r.issues.filter(i => i.code === 'no-etd');
  assert.deepEqual(ne.map(i => i.po), ['900904'], 'no ETD only before any section');
  assert.ok(r.allocs.find(a => a.po === '900901').flags.includes('section-not-listed'));
  assert.ok(r.allocs.find(a => a.po === '900906').flags.includes('section-not-listed'));
  assert.equal(r.allocs.find(a => a.po === '900901').section, 'Odd One');
  assert.equal(r.allocs.find(a => a.po === '900904').section, '');
  assert.equal(r.counts.sectionsNotListed, 3);
  assert.equal(r.counts.noEtd, 1);
  balanced(r);
  // Once it's on a list, it follows it.
  const p = runS(ls, ct, { sections: { ...SECT, planning: ['Plan X', 'round  tbd'] } });
  assert.equal(inBucket(p, 'planning'), 7);
  assert.equal(p.issues.filter(i => i.code === 'section-not-listed').length, 2);
  balanced(p);
});

test("sections missing: no section rules (Booked/Loaded, no ETD per block, nothing about the lists); the lowest-PO rule still applies", () => {
  const ls = [line({ po: '900951', qty: 600, crd: D(12, 1) }), line({ po: '900952', sty: 'AB200', qty: 300, crd: D(12, 1) })];
  const ct = ctOf([sec('Loaded', '900951', 'Direct A', [{ qty: 600 }]), sec('Booking', '900952', 'Plan X', [{ style: 'AB200', qty: 300 }]),
    sec('Booking', '900953', 'Hidden Test', [{ style: 'ZZ1', qty: 40 }]), sec('Booking', '900954', 'Odd One', [{ style: 'ZZ2', qty: 9 }])]);
  const r = run(ls, ct);
  assert.equal(JSON.stringify(run(ls, ct, { sections: null })), JSON.stringify(r), 'missing and null are the same');
  assert.equal(r.totals.units.loaded, 600); assert.equal(r.totals.units.booked, 300); assert.equal(r.totals.units.unmatched, 49);
  assert.equal(r.totals.units.pullFromChina, 0);
  assert.ok(!r.allocs.some(a => ['planning', 'ignored', 'belowTracker'].includes(a.bucket)));
  assert.equal(r.issues.filter(i => i.code === 'no-etd').length, 4);
  assert.ok(!codes(r).some(c => ['section-not-listed', 'ignored-section', 'below-tracker'].includes(c)), codes(r).join(','));
  assert.ok(r.allocs.filter(a => a.sheetRow != null).every(a => a.flags.includes('no-etd')));
  // Empty lists are lists: every section with no round is then not on them.
  const e = run(ls, ct, { sections: {} });
  assert.equal(e.issues.filter(i => i.code === 'section-not-listed').length, 4);
  assert.equal(e.issues.filter(i => i.code === 'no-etd').length, 0);
  assert.equal(e.totals.units.loaded + e.totals.units.booked, 900);
  balanced(r); balanced(e);
  // The PO Status Tracker's lowest PO is its own rule (rule 4), not a section rule: with it known, a PO below it is
  // below the tracker whatever config.js says; the rest is as above.
  const low = ctOf([...ct.blocks, sec('Loaded', '800951', 'Direct A', [{ style: 'OLD1', qty: 70 }])]);
  const b = ST.build({ po: { lines: ls, trk: { minPO: 900000 } }, mrows: mrowsFor(ls), ct: low, today: TODAY });
  assert.equal(inBucket(b, 'belowTracker'), 70);
  assert.equal(b.issues.filter(i => i.code === 'below-tracker').length, 1);
  assert.equal(b.totals.units.loaded, 600); assert.equal(b.totals.units.booked, 300); assert.equal(b.totals.units.unmatched, 49);
  assert.equal(b.issues.filter(i => i.code === 'no-etd').length, 4, 'the PO below the tracker raises no no-etd');
  assert.ok(!codes(b).some(c => ['section-not-listed', 'ignored-section'].includes(c)), codes(b).join(','));
  balanced(b);
});

test('sections: an open line takes a direct-to-customer section before a shipped line of the same PO and style does', () => {
  // A section never moves into a round, so a shipped line (inside a Round on Master Sheet Live) doesn't take it first.
  const ls = [shippedLine({ po: '900961', qty: 500, row: 600 }), line({ po: '900961', qty: 500, crd: D(12, 1), row: 601 })];
  const r = runS(ls, ctOf([sec('Loaded', '900961', 'Direct A', [{ qty: 500 }])]));
  assert.equal(units(r, a => a.status === 'open' && a.bucket === 'pullFromChina'), 500);
  assert.equal(units(r, a => a.status === 'shipped' && a.bucket === 'short'), 500, 'the shipped line has no Container Tracking line of its own');
  assert.equal(r.totals.units.pullFromChina, 500);
  assert.equal(r.totals.units.incoming, 0, 'nothing in In China or Incoming');
  assert.equal(r.pull.find(p => p.po === '900961').ctStatus, 'Loaded · Direct A');
  assert.ok(!codes(r).includes('ct-behind-msl'), codes(r).join(','));
  balanced(r);
  // What the open lines leave still goes to the shipped line, as Pull from China (capped by the group), listed.
  const l2 = [shippedLine({ po: '900962', qty: 500, row: 610 }), line({ po: '900962', qty: 300, crd: D(12, 1), row: 611 })];
  const r2 = runS(l2, ctOf([sec('Loaded', '900962', 'Direct A', [{ qty: 1000 }])]));
  assert.equal(units(r2, a => a.status === 'open' && a.bucket === 'pullFromChina'), 300);
  assert.equal(units(r2, a => a.status === 'shipped' && a.bucket === 'pullFromChina' && a.flags.includes('ct-behind-msl')), 500);
  assert.equal(inBucket(r2, 'ctMore'), 200);
  assert.match(r2.issues.find(i => i.code === 'ct-behind-msl').message, /500 units under "Direct A".*counted as Pull from China/);
  balanced(r2);
});

test('sections: a round and a direct-to-customer section on one PO and style: the round keeps its units first, whatever the row order', () => {
  const ls = [line({ po: '900971', qty: 1000, crd: D(12, 1) })];
  const R5 = sheet => blk(sheet, '900971', [355], [{ qty: 1000 }]), DA = sheet => sec(sheet, '900971', 'Direct A', [{ qty: 1000 }]);
  const rows = [crow('Open', 355, { etd: D(11, 20), eta: D(12, 20) })];
  const cases = [[DA('Loaded'), R5('Loaded')], [R5('Loaded'), DA('Loaded')], [R5('Booking'), DA('Loaded')], [DA('Booking'), R5('Loaded')]];
  for (const blocks of cases) {
    const r = runS(ls, ctOf(blocks, rows));
    const t = r.totals.units, rk = blocks.find(b => b.rounds.length).sheet;
    assert.equal(rk === 'Loaded' ? t.loaded : t.booked, 1000, blocks.map(b => b.sheet + ' ' + b.section).join(', '));
    assert.equal(t.pullFromChina, 0);
    assert.deepEqual(r.allocs.filter(a => a.bucket === 'ctMore').map(a => [a.units, a.section]), [[1000, 'Direct A']]);
    balanced(r);
  }
  // A shipped line takes the round (On the water, its Round row says it sailed); the section's units are beyond it.
  const s = runS([shippedLine({ po: '900971', qty: 1000, crd: D(10, 1) })], ctOf([R5('Booking'), DA('Loaded')], rows));
  assert.equal(s.totals.units.water, 1000);
  assert.deepEqual(s.allocs.filter(a => a.bucket === 'ctMore').map(a => [a.units, a.section]), [[1000, 'Direct A']]);
  assert.equal(s.totals.units.pullFromChina, 0);
  balanced(s);
});

test('sections: one section spelled two ways on Booking and Loaded is one section (counted once, named once)', () => {
  const ls = [line({ po: '900981', qty: 1000, crd: D(12, 1) })];
  // A direct-to-customer section: Booking's copy left out, with its both-tabs line, as for a round.
  const d = runS(ls, ctOf([sec('Booking', '900981', 'Direct A', [{ qty: 1000 }]), sec('Loaded', '900981', ' direct  a', [{ qty: 1000 }])]));
  assert.equal(d.totals.units.pullFromChina, 1000);
  assert.equal(inBucket(d, 'ctMore'), 0);
  assert.deepEqual(codesBut(d, ['no-case-pack', 'missing-column']), ['both-tabs']);
  assert.equal(d.counts.bothTabs, 1);
  balanced(d);
  // A planning section: placed once, named once, and no per-line both-tabs line (planning units have no Problems line).
  const p = runS(ls, ctOf([sec('Booking', '900981', 'Plan X', [{ qty: 1000 }]), sec('Loaded', '900981', 'plan x ', [{ qty: 1000 }])]));
  assert.equal(inBucket(p, 'planning'), 1000);
  assert.equal(p.pull.find(x => x.po === '900981').ctStatus, 'Planned · plan x');
  assert.ok(!codes(p).includes('both-tabs'), codes(p).join(','));
  balanced(p);
  // Below the tracker's lowest PO, or ignored: the one summary line only (no both-tabs, no negative-qty line).
  for (const [po, name, code] of [['800981', 'Direct A', 'below-tracker'], ['900982', 'Hidden Test', 'ignored-section']]) {
    const blocks = [sec('Booking', po, name, [{ style: 'ZZ1', qty: 40 }, { style: 'ZZ2', qty: -5 }]), sec('Loaded', po, name.toUpperCase(), [{ style: 'ZZ1', qty: 40 }])];
    const r = ST.build({ po: { lines: [], trk: { minPO: 900000 } }, mrows: mrowsFor([]), ct: ctOf(blocks), today: TODAY, sections: SECT });
    assert.deepEqual(codesBut(r, ['no-case-pack', 'missing-column']), [code], `${name}: ${codes(r).join(',')}`);
    assert.equal(units(r, a => a.ctRows.length > 0), 40, 'placed once');
    assert.equal(r.counts.negativeQty, 1, 'still counted');
    balanced(r);
  }
});

test('a negative quantity on Booking or Loaded is listed, never silently dropped', () => {
  const ls = [line({ po: '900891', qty: 1000, crd: D(12, 1) })];
  const ct = ctOf([blk('Loaded', '900891', [], [{ qty: 1000 }, { qty: -480 }])]);
  const r = run(ls, ct);
  const i = r.issues.find(x => x.code === 'negative-qty');
  assert.ok(i && i.sheet === 'Loaded' && /-480/.test(i.message), JSON.stringify(i));
  assert.equal(r.counts.negativeQty, 1);
  assert.equal(r.totals.units.loaded, 1000);
  balanced(r);
});

test('as the contract orders the steps: a shipped line with no Container Tracking line of its own takes an open line\'s booking', () => {
  // Step 3 (shipped takes unsailed units) runs before step 5 (open takes them). Documented here; Saul may change it.
  const ls = [shippedLine({ po: '900895', qty: 500, row: 300 }), line({ po: '900895', qty: 500, row: 301, crd: D(9, 1) })];
  const r = run(ls, ctOf([blk('Booking', '900895', [], [{ qty: 500 }])]));
  assert.equal(units(r, a => a.status === 'shipped' && a.bucket === 'water' && a.flags.includes('ct-behind-msl')), 500);
  assert.equal(units(r, a => a.status === 'open' && a.bucket === 'ready'), 500);
  assert.equal(r.totals.units.incoming, 1000);
  balanced(r);
});

// ---------- cases

test('cases: ceil(units / Master), part cases flagged with one summary line, no Master shows units only', () => {
  const ls = [
    line({ po: '901801', sty: 'AB1', qty: 96, master: 48 }),
    line({ po: '901801', sty: 'AB2', qty: 100, master: 48 }),
    line({ po: '901801', sty: 'AB3', qty: 30, master: 12 }),
    line({ po: '901801', sty: 'AB4', qty: 50 }),
  ];
  const r = run(ls, ctOf([]));
  const by = Object.fromEntries(r.allocs.map(a => [a.style, a]));
  assert.equal(by.AB1.cases, 2);
  assert.equal(by.AB2.cases, 3);
  assert.ok(by.AB2.flags.includes('part-case') && !by.AB1.flags.includes('part-case'));
  assert.equal(by.AB4.cases, null);
  assert.equal(r.totals.cases.noCrd, 2 + 3 + 3);
  assert.equal(r.totals.noCasePackUnits.noCrd, 50);
  assert.equal(r.issues.filter(i => i.code === 'part-case').length, 1);
  assert.match(r.issues.find(i => i.code === 'part-case').message, /^2 quantities/);
  assert.deepEqual(r.issues.filter(i => i.code === 'no-case-pack').map(i => i.style), ['AB4']);
  assert.equal(r.styles.find(s => s.style === 'AB4').noCasePack, true);
});

test("cases for Container Tracking-only units: the block line's Master, else the style's latest Master Sheet Master", () => {
  const ls = [shippedLine({ po: '901901', sty: 'AB1', qty: 10, master: 6, row: 200 }), shippedLine({ po: '901902', sty: 'AB1', qty: 10, master: 12, row: 201 })];
  const ct = ctOf([
    blk('Booking', '901990', [], [{ style: 'AB1', qty: 24 }]),               // unmatched: the style's latest Master (12)
    blk('Booking', '901991', [], [{ style: 'AB1', qty: 24, master: 8 }]),    // its own Master
  ]);
  const r = run(ls, ct);
  const u = r.allocs.filter(a => a.bucket === 'unmatched');
  assert.deepEqual(u.map(a => [a.po, a.cases, a.from.cases]), [['901990', 2, 'Master Sheet Live'], ['901991', 3, 'Container Tracking']]);
});

// ---------- outputs

test('styles: UPC rows sum to the style, UPC_MASTER description, several factories, next ETA', () => {
  const ls = [
    line({ po: '902001', sty: 'AB1', qty: 100, upc: '000111222333', crd: D(12, 1) }),
    line({ po: '902002', sty: 'AB1', qty: 50, upc: '000111222340', sup: 'Factory B', crd: D(12, 1) }),
  ];
  const ct = ctOf([blk('Loaded', '902001', [395], [{ style: 'AB1', qty: 60 }]), blk('Loaded', '902002', [396], [{ style: 'AB1', qty: 50 }])],
    [crow('Open', 395, { containers: ['TEST0000005'], etd: D(10, 1), eta: D(11, 20) }), crow('Open', 396, { containers: ['TEST0000006'], etd: D(10, 1), eta: D(11, 10) })]);
  const upc = { byStyle: new Map([['AB1', [{ upc: '000111222333', desc: 'Lip gloss', category: 'Lips', sub: 'Gloss' }]]]) };
  const r = run(ls, ct, { upc });
  const s = r.styles[0];
  assert.equal(s.desc, 'Lip gloss');
  assert.equal(s.from.desc, 'UPC_MASTER');
  assert.equal(s.category, 'Lips');
  assert.equal(s.factory, 'Several');
  assert.deepEqual(s.factories, ['Factory A', 'Factory B']);
  assert.equal(s.upcs.length, 2);
  for (const k of ST.TOTAL_KEYS) assert.equal(s.upcs.reduce((n, u) => n + u.units[k], 0), s.units[k], k);
  assert.equal(s.units.water, 110);
  assert.equal(s.units.inProd, 40);
  assert.equal(s.units.incoming, 150);
  assert.equal(s.nextEta, D(11, 10));
  assert.equal(s.allocs.length, r.allocs.length);
});

test('the side panel facts: PO shown, customer, CRD, round, containers, ETD/ETA, stage and the file each came from', () => {
  const ls = [line({ po: '902101-1', key: '902101|AB1', derived: true, bo: true, sty: 'AB1', qty: 100, crd: D(9, 1) })];
  const ct = ctOf([blk('Loaded', '902101-1', [397], [{ style: 'AB1', qty: 100 }])],
    [crow('Open', 397, { containers: ['TEST0000007'], etd: D(10, 1), eta: D(11, 20) })]);
  const r = run(ls, ct, { forwarder: fwdOf([{ container: 'TEST0000007', status: 'Sailed', etd: D(10, 2), eta: D(11, 22) }]) });
  const a = r.allocs[0];
  assert.equal(a.bucket, 'water');
  assert.equal(a.po, '902101');
  assert.equal(a.poShown, '902101-1');
  assert.equal(a.customer, 'Customer 1');
  assert.equal(a.round, '397');
  assert.deepEqual(a.containers, ['TEST0000007']);
  assert.equal(a.eta, D(11, 22));
  assert.deepEqual({ stage: a.from.stage, eta: a.from.eta, units: a.from.units, round: a.from.round, poShown: a.from.poShown },
    { stage: 'Forwarder', eta: 'Forwarder', units: 'Master Sheet Live', round: 'Container Tracking', poShown: 'PO Book' });
  assert.ok(a.id && typeof a.id === 'string');
});

test('Pull from China: one row per open line with its Container Tracking status, in plain words', () => {
  const ls = [line({ po: '902201', qty: 5040, crd: D(12, 1) }), line({ po: '902202', sty: 'AB2', qty: 10 }), shippedLine({ po: '902203', sty: 'AB3', qty: 10 })];
  const ct = ctOf([blk('Booking', '902201', [398], [{ qty: 3000 }])], [crow('Open', 398, { etd: D(11, 10), eta: D(12, 10) })]);
  const r = run(ls, ct);
  assert.equal(r.pull.length, 2);
  const p = r.pull.find(x => x.po === '902201');
  assert.equal(p.ctStatus, 'Booked 3,000 · Not on tracker 2,040');
  assert.equal(p.round, '398');
  assert.equal(p.etd, D(11, 10));
  assert.equal(r.pull.find(x => x.po === '902202').ctStatus, 'Not on tracker');
});

test('issues: source Incoming, known codes, plain messages', () => {
  const ls = [line({ po: '902301', qty: 10, tracked: false }), shippedLine({ po: '902302', qty: 10, crd: D(10, 1) })];
  const ct = ctOf([blk('Booking', '902390', [], [{ style: 'ZZ9', qty: 5 }])]);
  const r = run(ls, ct);
  assert.ok(r.issues.length >= 3);
  for (const i of r.issues) {
    assert.equal(i.source, 'Incoming');
    assert.ok(CODES.has(i.code), i.code);
    assert.ok(i.message.length > 20 && !/undefined|NaN|\[object/.test(i.message), i.message);
  }
});

// ---------- balance on a large random fixture

function rng(seed) { let s = seed >>> 0; return () => ((s = (s * 1664525 + 1013904223) >>> 0) / 2 ** 32); }
const SECTION_POOL = ['Planning', 'Direct A', 'direct b', 'Plan X', 'Hidden Test', 'Odd One', null];
test('balance: every Master Sheet unit ends in exactly one outcome on a large random fixture', () => {
  for (const seed of [1, 7, 42]) {
    const R = rng(seed), pick = a => a[Math.floor(R() * a.length)], int = (a, b) => a + Math.floor(R() * (b - a + 1));
    const lines = [], blocks = [], rows = [], fwd = [], dray = [];
    let row = 2;
    for (let i = 0; i < 400; i++) {
      const po = String(903000 + i), sty = 'AB' + int(1, 60) + pick(['', '', ' X', '-F']);
      const cp = R() < 0.1, tracked = R() < 0.9;
      for (let j = int(1, 3); j > 0; j--) {
        const shipped = R() < 0.4, kind = R();
        const extra = kind < 0.1 ? { po: po + '-1', srcPo: po, renamed: true, bo: true } : kind < 0.2 ? { po: po + '-2', derived: true, bo: true } : {};
        lines.push(line({ po, sty, qty: pick([0, 48, 100, 1000, 2400, 5040]), crd: pick([null, TODAY - 30, TODAY, TODAY + 40]),
          ptype: cp ? 'China Pickup' : 'LA', tracked, master: pick([null, 12, 48]), upc: pick([null, '000111222333']),
          key: `${po}|${sty.replace(/\s+/g, '')}${j > 1 ? '#' + j : ''}`, row: row++, ...(shipped ? { src: 'round', status: 'shipped', dupRound: j > 1 && kind > 0.85 } : {}), ...extra }));
      }
      for (let k = R() < 0.7 ? int(1, 3) : 0; k > 0; k--) {
        const n = int(400, 460), sheet = pick(['Booking', 'Loaded']);
        const ctPo = R() < 0.9 ? po : po + 'A', ctSty = R() < 0.9 ? sty : sty.replace(/\s+/g, '');
        const rounds = R() < 0.85 ? [n] : [];
        blocks.push(blk(sheet, ctPo, rounds, [{ style: ctSty, qty: pick([0, 50, 100, 1000, 3000]), master: pick([null, 24]) }], rounds.length ? {} : { section: pick(SECTION_POOL) }));
        if (R() < 0.8) {
          const c = 'TEST' + String(1000000 + i * 10 + k);
          rows.push(crow(pick(['Open', 'Open', 'Complete']), n, { containers: R() < 0.8 ? [c] : [], etd: TODAY + int(-60, 30), eta: TODAY + int(-30, 60), row: int(2, 300) }));
          if (R() < 0.4) fwd.push({ container: c, status: pick(['Pending Departure', 'Sailed', 'Soon To Arrive', 'Arrived', 'Odd']), eta: TODAY + int(-20, 50) });
          if (R() < 0.3) dray.push({ container: c, availability: pick(['DELIVERED', 'AVAILABLE', 'NOT AVAILABLE']), eta: TODAY + int(-20, 50), actualPickup: R() < 0.3 ? TODAY - 3 : null });
        }
      }
    }
    // Booking duplicates of Loaded lines (counted once) and unmatched lines.
    for (const b of blocks.filter(x => x.sheet === 'Loaded' && x.rounds.length).slice(0, 20)) blocks.push(blk('Booking', b.po, b.rounds, [{ style: b.lines[0].style, qty: 77 }]));
    for (let i = 0; i < 50; i++) blocks.push(blk(pick(['Booking', 'Loaded']), String(990000 + i), [int(400, 460)], [{ style: 'ZZ' + i, qty: 123 }]));
    // POs below the tracker's lowest PO (900000 below), in rounds and in sections.
    for (let i = 0; i < 30; i++) {
      const rounds = R() < 0.5 ? [int(400, 460)] : [];
      blocks.push(blk(pick(['Booking', 'Loaded']), String(800000 + i), rounds, [{ style: 'OLD' + i, qty: 77 }], rounds.length ? {} : { section: pick(SECTION_POOL) }));
    }
    const ct = ctOf(blocks, rows), forwarder = fwdOf(fwd), drayage = drayOf(dray);
    const snap = () => JSON.stringify([lines, ct, [...forwarder.byContainer], drayage]);
    const before = snap();
    const r = run(lines, ct, { forwarder, drayage });
    assert.equal(snap(), before, 'the inputs are not changed (the page rebuilds from cached data)');
    assert.equal(JSON.stringify(run(lines, ct, { forwarder, drayage })), JSON.stringify(r), 'the same inputs give the same result');
    // With the section rules and the tracker's lowest PO: both balances still 0, every outcome used.
    const rs = ST.build({ po: { lines, trk: { minPO: 900000 } }, mrows: mrowsFor(lines), ct, forwarder, drayage, today: TODAY, sections: SECT });
    assert.equal(snap(), before);
    balanced(rs);
    assert.equal(rs.balance.ct.units, r.balance.ct.units, 'the same Container Tracking units, placed differently');
    for (const b of ['pullFromChina', 'planning', 'ignored', 'belowTracker']) assert.ok(inBucket(rs, b) > 0, b);
    for (const a of rs.allocs) assert.ok(ST.BUCKETS.includes(a.bucket), a.bucket);
    for (const i of rs.issues) assert.ok(CODES.has(i.code), i.code);
    const perLineS = new Map();
    for (const a of rs.allocs) if (a.lineKey) perLineS.set(a.lineKey, (perLineS.get(a.lineKey) || 0) + a.units);
    for (const l of lines) assert.equal(perLineS.get(l.key) || 0, l.qty, l.key);
    assert.equal(rs.issues.filter(i => i.code === 'below-tracker').length, 1);
    assert.ok(rs.issues.filter(i => i.code === 'section-not-listed').length <= 2, 'at most one per unlisted section name (Planning, Odd One)');
    assert.equal(new Set(rs.allocs.map(a => a.id)).size, rs.allocs.length, 'allocation ids are unique');
    assert.equal(rs.pull.length, r.pull.length);
    balanced(r);
    assert.equal(r.balance.mslUnits, lines.reduce((s, l) => s + l.qty, 0));
    // Per line: its allocations add up to its quantity.
    const perLine = new Map();
    for (const a of r.allocs) if (a.lineKey) perLine.set(a.lineKey, (perLine.get(a.lineKey) || 0) + a.units);
    for (const l of lines) assert.equal(perLine.get(l.key) || 0, l.qty, l.key);
    // Totals are the sum of the counted allocations, and the derived keys add up.
    const t = r.totals.units;
    assert.equal(t.inChina, t.noCrd + t.inProd + t.ready + t.booked + t.loaded);
    assert.equal(t.incoming, t.inChina + t.water + t.port);
    const sumStyles = k => r.styles.reduce((s, x) => s + x.units[k], 0);
    for (const k of ST.TOTAL_KEYS) assert.equal(sumStyles(k), t[k], k);
    for (const a of r.allocs) assert.ok(ST.BUCKETS.includes(a.bucket), a.bucket);
    for (const i of r.issues) assert.ok(CODES.has(i.code), i.code);
    assert.equal(r.pull.length, lines.filter(l => l.status === 'open').length);
    assert.ok(r.counts.bothTabs >= 20);
    assert.ok(r.counts.unmatchedBlocks >= 50);
    assert.equal(new Set(r.allocs.map(a => a.id)).size, r.allocs.length, 'allocation ids are unique');
  }
});

// ---------- through the real parsers

test('through the real parsers: PO Book core, Container Tracking and forwarder workbooks', () => {
  const X = loadXLSX(), PB = core('pobook'), CT = core('ct'), SH = core('shipping');
  const ser = d => d + 25569;
  const sheetRows = (rows, name) => C.sheetRows(X, readBook(xlsxBytes({ [name]: rows })).Sheets[name]);
  const trows = sheetRows([
    ['PO#', 'Supplier', 'Order date', 'Item Description', 'Style #', 'Order Type', 'Customer', 'Customer 1st Ship Date', 'CRD Expected', 'Factory CRD'],
    ['900100', 'Factory A', null, 'Old', 'AB050', 'LA', 'Customer 1', null, null, ser(D(6, 1))],
    ['900101', 'Factory A', null, 'Gloss', 'AB100', 'LA', 'Customer 1', null, null, ser(D(9, 1))],
    ['900102', 'Factory B', null, 'Brush', 'AB200', 'China Pickup', 'Customer 2', null, null, ser(D(12, 1))],
    ['900103', 'Factory A', null, 'Liner', 'AB 300', 'LA', 'Customer 1', null, null, ser(D(12, 1))],
  ], 'PO Status Tracker');
  const MHR = ['Category', 'Sub', 'Product name', 'Style Number', 'Barcode', 'Master', 'Price', 'Quantity', 'Po Number', 'Factory', 'Order Date', 'Ship by', 'Notes'];
  const mrows = sheetRows([MHR,
    ['Lip', 'Gloss', 'Old', 'AB050', null, 50, 1, 500, '900100', 'Factory A', null, null, null],
    [null, null, 'Round 340', null, null, null, null, null, null, null, null, null, null],
    ['Lip', 'Gloss', 'Gloss', 'AB100', '000111222333', 48, 1, 4800, '900101', 'Factory A', null, null, null],
    ['Tool', 'Brush', 'Brush', 'AB200', null, 25, 1, 1000, '900102', 'Factory B', null, null, null],
    ['Eye', 'Liner', 'Liner', 'AB 300', null, 12, 1, 600, '900103', 'Factory A', null, null, null],
  ], 'Master Sheet Live');
  const trk = PB.parseTracker(trows), src = PB.parseMaster(mrows, trk);
  const lines = PB.applyOverlay(src, new Map(), PB.cfgFrom([]));
  const PH = ['Category', 'Style', 'Barcode', 'Master', 'Quantity', 'PO Number', 'Company'];
  const CH = ['ID#', 'Container #', 'Round #', 'ETD', 'ETA'];
  const wb = readBook(xlsxBytes({
    Booking: [PH, ['Round 351', null, null, null, null, null, null], [null, 'AB200', null, 25, 1000, '900102', 'Factory B'], [null, 'AB300', null, 12, 600, '900103', 'Factory A']],
    Loaded: [PH, ['Round 350', null, null, null, null, null, null], [null, 'AB100', '000111222333', 48, 3000, '900101', 'Factory A'],
      ['Round 340', null, null, null, null, null, null], [null, 'AB050', null, 50, 500, '900100', 'Factory A']],
    Open: [CH, [1, 'TEST1234567', 350, ser(D(10, 1)), ser(D(11, 20))]],
    Complete: [CH, [2, 'TEST7654321', 340, ser(D(7, 1)), ser(D(8, 1))]],
  }));
  const ct = CT.parseContainerTracking(X, wb, TODAY);
  assert.equal(ct.ok, true, ct.error);
  const fwb = readBook(xlsxBytes({ sheet1: [['Active shipments'], ['ID', 'Status', 'ETD', 'ETA', 'Container Numbers', 'PO Numbers'],
    ['S1', 'Sailed', ser(D(10, 2)), ser(D(11, 22)), 'TEST1234567', '900101']] }));
  const forwarder = SH.parseForwarder(X, fwb, TODAY);
  assert.equal(forwarder.ok, true, forwarder.error);
  const r = ST.build({ po: { lines, trk, mrows }, ct, forwarder, drayage: null, upc: null, today: TODAY });
  assert.equal(r.ok, true);
  const by = (sty, b) => units(r, a => a.style === sty && a.bucket === b);
  assert.equal(by('AB050', 'landed'), 500, 'shipped in Round 340, landed');
  assert.equal(by('AB100', 'water'), 3000, 'open, but the forwarder says sailed');
  assert.equal(by('AB100', 'ready'), 1800, 'the rest at its Master Sheet stage (CRD passed)');
  assert.equal(by('AB200', 'pullFromChina'), 1000);
  assert.equal(by('AB 300', 'inProd'), 600);
  assert.equal(by('AB300', 'unmatched'), 600, 'inner space: a different item');
  assert.equal(r.allocs.find(a => a.style === 'AB100' && a.bucket === 'water').cases, 63);
  assert.equal(r.styles.find(s => s.style === 'AB100').upcs[0].upc, '000111222333');
  assert.deepEqual([...new Set(codes(r))].sort(), ['msl-behind-ct', 'near-miss', 'no-etd', 'part-case']);
  assert.equal(r.totals.units.incoming, 3000 + 1800 + 600);
  balanced(r);
});
