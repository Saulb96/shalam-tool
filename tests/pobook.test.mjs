import test from 'node:test';
import assert from 'node:assert/strict';
import { readFileSync, existsSync } from 'node:fs';
import { createHash } from 'node:crypto';
import { core, root, loadXLSX, xlsxBytes, readBook } from './load.mjs';

const P = core('pobook');
const sha = s => createHash('sha256').update(s, 'utf8').digest('hex');

// Rows read back from a real .xlsx the way the PO Book reads a sheet.
function sheet(rows, name = 'S') {
  const ws = readBook(xlsxBytes({ [name]: rows })).Sheets[name];
  return loadXLSX().utils.sheet_to_json(ws, { header: 1, raw: true, defval: null, blankrows: true });
}

// Invented data, ported from Dev PLM's tests/test_po_book.py (the Python twin of the same code), so both agree.
const TH = ['PO#', 'Supplier', 'Order date', 'Item Description', 'Style #', 'Order Type', 'Customer',
  'Customer 1st Ship Date', 'CRD Expected', 'Factory CRD'];
const MH = ['Product name', 'Style Number', 'Price', 'Quantity', 'Po Number', 'Factory', 'Order Date', 'Ship by', 'Notes'];
const t = (po, sty, { sup = 'Factory A', cust = 'Customer 1', crd = null, exp = null, typ = 'LA', od = '6/1/2026' } = {}) =>
  [po, sup, od, `desc ${sty}`, sty, typ, cust, null, exp, crd];
const m = (name, sty, qty, { po = null, fac = '', sb = null, price = 1 } = {}) => [name, sty, price, qty, po, fac, null, sb, null];
const day = s => P.toDay(s);

const TRACKER = [TH,
  t(900300, 'X1'),
  t('900437', 'J1', { crd: '9/1/2026' }),
  t('900437', 'J2', { crd: '9/5/2026' }),
  t('900500', 'RM1-F', { sup: 'Factory B', cust: 'Customer 2', typ: 'China Pickup', exp: '10/1/2026' }),
  t('900600', 'SV1', { crd: '11/1/2026' }),
  t('bulk', 'Q1'), // no six digits: the PO Book leaves it out
  t('900650', null)]; // no style: left out
const MASTER = [MH,
  m('Old thing', 'X1', 10, { po: '900300' }),
  m('Round 1 (May)', null, null),
  m('Lip gloss', 'J1', 100, { po: '900437', fac: 'Factory A' }),
  m('Lip gloss 2', 'J2', 50), // no Po Number: still 900437, the block carries down
  m('  round 2', null, null), // the last Round (any case, leading spaces)
  m('Lip gloss 2 again', 'J2', 50, { po: '900437' }), // below the last Round, but already in one: shipped
  m('Brush', 'RM1', 30, { po: '900500', sb: '9/20/2026' }), // the tracker's RM1-F answers for RM1
  m('Bag', 'SV1', 10, { po: '900600' }),
  m('Bag 2', 'SV2', 5), // not on the PO Status Tracker: Master Sheet Live only
  m('Ancient', 'A1', 1, { po: '900200' }), // older than the tracker's first PO: left out
  m('Backorder', 'Z1', 1, { po: '900700-1', sb: '12/1/2026' }),
  m('No style', null, 3)];

const trk = () => P.parseTracker(sheet(TRACKER));
const parsed = () => P.parseMaster(sheet(MASTER), trk());
const by = (lines, key) => lines.find(l => l.key === key);

const item = (id, modified, fields) => ({ id: String(id), lastModifiedDateTime: modified, fields });
const OVERLAY = [
  item(1, '2026-10-01T10:00:00Z', { LineKey: '__settings', Kind: 'settings', Data: '{"src": {"tracker": {"d": "D1", "i": "I1"}}, "la": 20}' }),
  item(2, '2026-10-01T10:00:00Z', { LineKey: '900600|SV2', Kind: 'line', Shipped: true, ShippedDate: '2026-10-01' }),
  item(3, '2026-10-02T10:00:00Z', { LineKey: '900500-1|RM1', Kind: 'split', PONumber: '900500-1', Style: 'RM1',
    ParentKey: '900500|RM1', Qty: 10, CRDOverride: '2026-12-15' }),
  item(4, '2026-10-02T10:00:00Z', { LineKey: '900700-1|Z1', Kind: 'line', PONumber: '900701' }),
  item(5, '2026-09-01T10:00:00Z', { LineKey: '900700-1|Z1', Kind: 'line', PONumber: '999999' }), // older: ignored
  item(6, '2026-10-02T10:00:00Z', { Kind: 'line' })]; // no LineKey: ignored

test('the copied block stays inside the module', () => {
  assert.equal(typeof globalThis.PoBookCore, 'object');
  for (const name of ['S', 'num', 'toDay', 'DAY', 'parseTracker', 'parseMaster', 'cols']) assert.equal(name in globalThis, false, name);
});

test('dates read as the PO Book reads them', () => {
  assert.equal(P.fmtDay(P.toDay(45678)), '1/21/2025');
  assert.equal(P.fmtDay(P.toDay('6/2/26')), '6/2/2026');
  assert.equal(P.toDay('2026-06-02T10:00:00Z'), P.toDay('6/2/2026'));
  assert.equal(P.toDay(30000), null);
  assert.equal(P.toDay('TBD'), null);
  assert.equal(P.normSty(' rm 1-f '), 'RM1-F');
});

test('a missing column stops the read, as in the PO Book', () => {
  assert.throws(() => P.parseTracker(sheet([TH.slice(0, -1), t('900437', 'J1').slice(0, -1)])), /missing column "Factory CRD"/);
});

test('tracker lines need six digits and a style; -F answers for the style without it', () => {
  const k = trk();
  assert.equal(k.minPO, 900300);
  assert.deepEqual(new Set(k.all.map(x => x.po)), new Set(['900300', '900437', '900500', '900600']));
  assert.equal(k.T.get('900500|RM1'), k.T.get('900500|RM1-F'));
});

test('Round rows decide shipped and open (ported)', () => {
  const p = parsed();
  assert.deepEqual(Object.fromEntries(p.lines.map(l => [l.key, l.src])), {
    '900300|X1': 'round', '900437|J1': 'round', '900437|J2': 'round', '900437|J2#2': 'round',
    '900500|RM1': 'open', '900600|SV1': 'open', '900600|SV2': 'open', '900700-1|Z1': 'open' });
  assert.equal(by(p.lines, '900437|J2#2').dupRound, true); // same PO, style and qty already inside a Round
  const codes = new Set(p.issues.map(i => i.code));
  for (const c of ['dupround', 'mslonly', 'nostyle']) assert.ok(codes.has(c), c);
});

test('Master Sheet lines take the tracker\'s details (ported)', () => {
  const p = parsed();
  const rm = by(p.lines, '900500|RM1');
  assert.deepEqual([rm.tracked, rm.sup, rm.cust, rm.ptype, rm.crd], [true, 'Factory B', 'Customer 2', 'China Pickup', day('10/1/2026')]);
  const sv2 = by(p.lines, '900600|SV2');
  assert.deepEqual([sv2.tracked, sv2.cust, sv2.po], [false, '', '900600']);
  const z = by(p.lines, '900700-1|Z1');
  assert.ok(z.bo);
  assert.equal(z.crd, day('12/1/2026')); // a -N PO takes Master Sheet Live's Ship by
  assert.equal(by(p.lines, '900437|J1').row, 4);
});

test('older POs and an empty tracker leave lines out', () => {
  assert.ok(!parsed().lines.some(l => l.po === '900200'));
  assert.deepEqual(P.parseMaster(sheet(MASTER), P.parseTracker(sheet([TH]))).lines, []);
});

test('PO_Book_Lines: the newest item per line, and the settings item (ported)', () => {
  const ov = P.overlayFromItems(OVERLAY);
  assert.deepEqual([...ov.keys()], ['900600|SV2', '900500-1|RM1', '900700-1|Z1']);
  assert.equal(ov.get('900700-1|Z1').PONumber, '900701');
  assert.equal(ov.get('900700-1|Z1')._id, '4');
  assert.deepEqual(P.cfgFrom(OVERLAY), { la: 20, amber: 14, refresh: 5, src: { tracker: { d: 'D1', i: 'I1' } } });
  assert.deepEqual(P.cfgFrom([item(1, null, { LineKey: '__settings', Data: 'not json' })]), { la: 30, amber: 14, refresh: 5 });
  assert.deepEqual(P.cfgFrom([]), { la: 30, amber: 14, refresh: 5 });
  assert.equal(P.overlayFromItems([]).size, 0);
});

test('hand marks, splits, renamed POs and backorders (ported)', () => {
  const lines = P.applyOverlay(parsed(), P.overlayFromItems(OVERLAY), P.cfgFrom(OVERLAY));
  const st = Object.fromEntries(lines.map(l => [l.po + ' ' + l.sty, l.status]));
  assert.equal(st['900600 SV2'], 'shipped'); // Shipped ticked in the PO Book
  assert.equal(st['900600-1 SV1'], 'open'); // the rest of a partly shipped PO becomes its next backorder
  assert.ok(by(lines, '900600|SV1').derived);
  const split = by(lines, '900500-1|RM1'), parent = by(lines, '900500|RM1');
  assert.deepEqual([split.po, split.qty, split.crd, split.status], ['900500-1', 10, day('2026-12-15'), 'open']);
  assert.deepEqual([parent.qty, parent.splitOut], [20, 10]);
  const z = by(lines, '900700-1|Z1');
  assert.deepEqual([z.po, z.srcPo, z.renamed], ['900701', '900700-1', true]);
  assert.equal(by(lines, '900437|J1').status, 'shipped');
  assert.equal(by(lines, '900500|RM1').status, 'open');
});

test('a partly shipped PO reuses an all-open backorder (ported)', () => {
  const rows = sheet([MH, m('A', 'A1', 1, { po: '900437' }), m('B', 'B1', 2), m('Round 1', null, null),
    m('C', 'C1', 3, { po: '900437' }), m('D', 'D1', 4, { po: '900437-1' })]);
  const k = P.parseTracker(sheet([TH, t('900437', 'A1'), t('900437', 'B1'), t('900437', 'C1'), t('900437-1', 'D1')]));
  const lines = P.applyOverlay(P.parseMaster(rows, k), new Map(), P.cfgFrom([]));
  assert.deepEqual(new Set(lines.map(l => `${l.sty} ${l.po} ${l.status}`)),
    new Set(['A1 900437 shipped', 'B1 900437 shipped', 'C1 900437-1 open', 'D1 900437-1 open']));
});

test('countMaster gives every row a result, and its lines are parseMaster\'s', () => {
  const blank = Array(MH.length).fill(null);
  const rows = sheet([MH,
    blank, // 2 blank
    m('Stray', 'S0', 5), // 3 no PO number above it
    m('Old', 'O1', 1, { po: '900100' }), // 4 older than the tracker's lowest PO
    m('Round 7', null, null), // 5 Round row
    m('After round', 'AR1', 2), // 6 no PO number since the Round row
    m('Thing', 'J1', 'n/a', { po: '900437' }), // 7 line, quantity not a number
    m(null, null, 4), // 8 no style, no name: the PO Book says nothing
    m('Named', null, 4), // 9 no style
    blank, // 10 blank inside a block
    m('Sample', 'T1', 1, { po: 'TBD' }), // 11 PO doesn't start with 6 digits
    m('Line 2', 'J2', null, { po: '900437' }), // 12 line, quantity blank
    m('Line 3', 'J3', 7)]); // 13 line
  const k = trk(), r = P.countMaster(rows, k);
  assert.equal(r.ok, true);
  assert.equal(r.rowsRead, rows.length - 1);
  assert.equal(Object.values(r.counts).reduce((a, b) => a + b, 0), r.rowsRead);
  assert.deepEqual(r.counts, { line: 3, round: 1, blank: 2, noPo: 2, badPo: 1, belowMin: 1, noStyle: 2 });
  const lineRows = r.results.filter(x => x.result === 'line').map(x => x.row);
  assert.deepEqual(lineRows, [7, 12, 13]);
  assert.deepEqual(P.parseMaster(rows, k).lines.map(l => l.row).sort((a, b) => a - b), lineRows);
  assert.deepEqual(r.issues.map(i => `${i.row} ${i.code}`), ['3 nopo', '6 nopo', '7 qty', '8 nostyle', '9 nostyle', '11 badpo', '12 qty']);
  assert.ok(r.issues.every(i => i.source === 'Master Sheet Live' && i.message));
});

test('countMaster agrees with parseMaster on the ported sheet', () => {
  const rows = sheet(MASTER), k = trk(), r = P.countMaster(rows, k);
  assert.equal(Object.values(r.counts).reduce((a, b) => a + b, 0), rows.length - 1);
  assert.equal(r.counts.line, P.parseMaster(rows, k).lines.length);
  assert.equal(r.counts.round, 2);
  assert.equal(r.counts.belowMin, 1);
});

test('countMaster: a missing column or nothing counted is never success', () => {
  const r = P.countMaster(sheet([MH.filter(h => h !== 'Quantity' && h !== 'Notes'), ['x']]), trk());
  assert.equal(r.ok, false);
  assert.deepEqual(r.missing, ['Quantity', 'Notes']);
  assert.match(r.error, /missing columns "Quantity", "Notes"\. Headers seen: Product name, Style Number/);
  const e = P.countMaster(sheet(MASTER), P.parseTracker(sheet([TH])));
  assert.equal(e.ok, false);
  assert.match(e.error, /PO Status Tracker gave no PO numbers/);
  assert.equal(e.counts.belowMin, e.rowsRead - e.counts.round - e.counts.blank);
  const none = P.countMaster(sheet([MH, m('Round 1', null, null)]), trk());
  assert.equal(none.ok, false);
  assert.match(none.error, /no PO lines found in 1 rows/);
});

test('countTracker gives every tracker row a result', () => {
  const blank = Array(TH.length).fill(null);
  const rows = sheet([...TRACKER, blank, t('900437', 'J1'), t(null, 'K1')]);
  const r = P.countTracker(rows);
  assert.equal(r.ok, true);
  assert.equal(r.rowsRead, rows.length - 1);
  assert.deepEqual(r.counts, { tracked: 6, blank: 1, badPo: 2, noStyle: 1 });
  assert.equal(r.counts.tracked, P.parseTracker(rows).all.length);
  assert.deepEqual(r.issues.map(i => `${i.row} ${i.code}`), ['7 badpo', '8 nostyle', '10 dup', '11 badpo']);
  const bad = P.countTracker(sheet([['PO#', 'Supplier'], [900101, 'Factory A']]));
  assert.equal(bad.ok, false);
  assert.match(bad.error, /missing columns "Order date", .*"Factory CRD"\. Headers seen: PO#, Supplier/);
  assert.equal(P.countTracker(sheet([TH])).ok, false);
});

test('styles the PO Book reads as one item, and PO cells holding more than one PO, are listed', () => {
  const tracker = sheet([TH, t('900101', 'AB100-F'), t('900102', 'RM 2201'), t('900102', 'RM2201'), t('900103A/B', 'CD1'),
    t('900104-BCF', 'EF1')]);
  const tc = P.countTracker(tracker);
  assert.deepEqual(tc.issues.map(i => `${i.row} ${i.code}`), ['4 style-alias', '5 po-text']);
  assert.match(tc.issues[0].message, /"RM 2201" and "RM2201"/);
  const k = P.parseTracker(tracker);
  const rows = sheet([MH, m('Gloss', 'AB100', 5, { po: '900101' }), // the tracker only has AB100-F
    m('Brush', 'CD1', 5, { po: '900103A/B/C' }), m('Brush 2', 'CD2', 5), // one PO cell holding three POs
    m('Bag', 'EF1', 5, { po: '900104-BCF' })]);
  const r = P.countMaster(rows, k);
  assert.deepEqual(r.issues.map(i => `${i.row} ${i.code}`), ['2 style-alias', '3 po-text']);
  assert.match(r.issues[0].message, /"AB100" takes its customer, CRD and China Pickup \/ LA from "AB100-F"/);
  assert.equal(r.counts.line, 4); // still counted as the PO Book counts them
  assert.equal(r.counts.line, P.parseMaster(rows, k).lines.length);
});

test('coreText finds the text between the markers, and round-trips', () => {
  const A = '/*CORE' + '-START*/', B = '/*CORE' + '-END*/';
  assert.equal(P.coreText(null), null);
  assert.equal(P.coreText('<script>nothing here</script>'), null);
  assert.equal(P.coreText(`x${A}only a start`), null);
  assert.equal(P.coreText(`<p>${A}\nconst a=1;\n${B}</p>`), '\nconst a=1;\n');
  assert.equal(P.coreText(`${A}\r\nconst a=1;\r\n${B}`), '\nconst a=1;\n'); // a Windows clone reads the same
  const own = P.coreText(readFileSync(root('core', 'pobook.js'), 'utf8'));
  assert.ok(own && own.includes('function parseMaster('));
  assert.equal(P.coreText(`<html>${A}${own}${B}</html>`), own);
  assert.equal(sha(own), P.CORE_SHA256); // the block in this file is the one the hash was taken from
});

const LIVE = process.env.PO_BOOK_HTML || root('..', 'po-book', 'index.html');
test('CORE_SHA256 matches the live PO Book', { skip: !existsSync(LIVE) && 'no local PO Book clone next to this repo' }, () => {
  const live = P.coreText(readFileSync(LIVE, 'utf8'));
  assert.ok(live, 'the PO Book has no CORE markers');
  assert.equal(sha(live), P.CORE_SHA256, 'the PO Book\'s rules changed since core/pobook.js was copied');
});
