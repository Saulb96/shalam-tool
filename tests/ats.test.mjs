import test from 'node:test';
import assert from 'node:assert/strict';
import { core, loadXLSX, xlsxBytes, readBook } from './load.mjs';

const C = core('common');
const A = core('ats');
const X = loadXLSX();
const TODAY = C.ymd(2026, 10, 30);
const NAME = 'Inventory Report - On Hand_In Transit_In Production 10-29-2026 12-58 PM.xlsx';

// ---------- invented data (styles AB…, warehouses WH 1 / WH 2), in the real report's layout
const TITLE = ['Inventory Report - On Hand_In Transit_In Production'];
const HEAD = ['Color Code', 'Style', 'Style Description', 'Category', 'Sub Category', 'Warehouse', 'In WH',
  'Total Customer Orders', 'ATS from in WH', 'ATS from In Trans', 'In Trans Qty', 'Import PO Backlog Qty', 'ETA Date', 'ETD Date'];
// One style row; numbers as text, as the ERP writes them (In WH and orders as numbers, ATS as text, like the sample).
const r = (style, inWh, orders, o = {}) => ['XX', style, o.desc ?? 'Item ' + style, o.cat ?? 'Lip', o.sub ?? 'Gloss', o.wh ?? 'WH 1',
  inWh, orders, o.ats ?? (typeof inWh === 'number' && typeof orders === 'number' ? String(inWh - orders) : '0'), '0',
  o.inTrans ?? null, o.backlog ?? null, o.eta ?? null, o.etd ?? null];
const grand = (inWh, cell = 0) => { const g = Array(HEAD.length).fill(null); g[cell] = 'Grand Total'; g[6] = inWh; g[7] = ''; g[8] = ''; return g; };
const book = (rows, sheet = 'Report', extra = {}) => readBook(xlsxBytes({ ...extra, [sheet]: rows }));
const parse = (rows, name = NAME, sheet, extra) => A.parseAts(X, book(rows, sheet, extra), name, TODAY);
const codes = res => res.issues.map(i => i.code);

test('reads the report: title on row 1, header by name, styles summed over warehouses, Grand Total checks', () => {
  const rows = [TITLE, HEAD,
    r('AB100', 1200, 200), r('AB101', 0, 0), r('AB102', 500, 800),
    r('AB103', 300, 100, { wh: 'WH 1' }), r('AB103', 50, 25, { wh: 'WH 2' }),
    grand(2050)];
  const res = parse(rows);
  assert.equal(res.ok, true, res.error);
  assert.equal(res.headerRow, 2);
  assert.equal(res.sheet, 'Report');
  assert.equal(res.reportDay, C.ymd(2026, 10, 29));
  assert.equal(res.reportDayFrom, 'file name');
  assert.equal(res.styles.size, 4);
  const s = res.styles.get('AB103');
  assert.equal(s.inWh, 350);
  assert.equal(s.orders, 125);
  assert.equal(s.erpAts, 225);
  assert.deepEqual(s.warehouses, ['WH 1', 'WH 2']);
  assert.deepEqual(s.rows.map(x => x.row), [6, 7]);
  assert.equal(s.desc, 'Item AB103');
  assert.equal(res.styles.get('AB102').inWh - res.styles.get('AB102').orders, -300, 'sold ahead stays negative');
  assert.equal(res.counts.styleRows, 5);
  assert.equal(res.counts.grandTotalRows, 1);
  assert.equal(res.counts.inWh, 2050);
  assert.equal(res.counts.multiWarehouseStyles, 1);
  assert.deepEqual(res.grandTotal.inWh, 2050);
  assert.deepEqual(codes(res), []);
});

test('numbers stored as text are read ("-1234", "1,200", "(50)"); text that is not a number is listed, never 0', () => {
  const rows = [TITLE, HEAD,
    r('AB100', '1,200', '200', { ats: '1000' }),
    r('AB101', '-40', '0', { ats: '-40' }),
    r('AB102', 100, 'n/a'),
    r('AB103', 70, '(50)', { ats: '120' }),
    r('AB104', 10, 5, { ats: 'lots' }),
    grand(1340)];
  const res = parse(rows);
  assert.equal(res.ok, true, res.error);
  assert.equal(res.styles.get('AB100').inWh, 1200);
  assert.equal(res.styles.get('AB101').inWh, -40);
  assert.equal(res.styles.get('AB103').orders, -50);
  const bad = res.styles.get('AB102');
  assert.equal(bad.orders, null, 'not a number: unknown, never 0');
  assert.equal(bad.unknown, true);
  const nan = res.issues.filter(i => i.code === 'not-a-number');
  assert.deepEqual(nan.map(i => i.style).sort(), ['AB102', 'AB104']);
  assert.match(nan.find(i => i.style === 'AB102').message, /Total Customer Orders "n\/a".*isn't known/);
  assert.match(nan.find(i => i.style === 'AB104').message, /ERP's ATS isn't compared/);
  assert.equal(res.styles.get('AB104').unknown, false, "only the ERP's ATS is unreadable: stock is still known");
  assert.ok(!codes(res).includes('erp-ats-differs'), 'an unreadable ERP ATS is not compared');
});

test('a blank In WH or Total Customer Orders on a style row: that style is unknown, one summary line', () => {
  const rows = [TITLE, HEAD, r('AB100', 100, null), r('AB101', null, 5, { ats: '' }), r('AB102', 10, 0), grand(110)];
  const res = parse(rows);
  assert.equal(res.ok, true, res.error);
  assert.equal(res.styles.get('AB100').orders, null);
  assert.equal(res.styles.get('AB101').inWh, null);
  assert.equal(res.styles.get('AB102').unknown, false);
  const nan = res.issues.filter(i => i.code === 'not-a-number');
  assert.equal(nan.length, 1);
  assert.match(nan[0].message, /^2 style rows have a blank In WH or Total Customer Orders/);
  assert.equal(res.counts.unknownStyles, 2);
});

test("the Grand Total row is the check: rows that don't add up fail the report", () => {
  const res = parse([TITLE, HEAD, r('AB100', 100, 0), r('AB101', 50, 0), grand(160)]);
  assert.equal(res.ok, false);
  assert.match(res.error, /the report doesn't add up: rows 150, Grand Total 160/);
  // An In WH that isn't a number can't be checked either.
  const res2 = parse([TITLE, HEAD, r('AB100', 100, 0), r('AB101', 'fifty', 0), grand(150)]);
  assert.equal(res2.ok, false);
  assert.match(res2.error, /rows 100, Grand Total 150 \(1 row has an In WH that isn't a number: row 4\)/);
  // Found in either of the first two cells, in any case; the Grand Total row is never a style.
  const res3 = parse([TITLE, HEAD, r('AB100', 100, 0), grand(100, 1).map(v => (v === 'Grand Total' ? 'GRAND TOTAL' : v))]);
  assert.equal(res3.ok, true, res3.error);
  assert.equal(res3.styles.size, 1);
});

test('no Grand Total row: listed, not a failure; a second one is listed and not counted', () => {
  const res = parse([TITLE, HEAD, r('AB100', 100, 0), r('AB101', 5, 0)]);
  assert.equal(res.ok, true, res.error);
  assert.deepEqual(codes(res), ['grand-total']);
  assert.match(res.issues[0].message, /no "Grand Total" row/);
  const res2 = parse([TITLE, HEAD, r('AB100', 100, 0), grand(100), grand(999)]);
  assert.equal(res2.ok, true, res2.error);
  assert.deepEqual(codes(res2), ['grand-total']);
  assert.match(res2.issues[0].message, /another "Grand Total" row/);
});

test("erp-ats-differs: the ERP's ATS from in WH against In WH − Total Customer Orders, per style", () => {
  const rows = [TITLE, HEAD,
    r('AB100', 0, 300, { ats: '0', wh: '' }),          // orders with no stock and no warehouse: the ERP shows 0
    r('AB101', 100, 40, { ats: '60' }),
    r('AB102', 100, 0, { ats: '60', wh: 'WH 1' }), r('AB102', 20, 80, { ats: '-60', wh: 'WH 2' }),   // summed: ERP 0, tool 120 − 80 = 40
    grand(220)];
  const res = parse(rows);
  assert.equal(res.ok, true, res.error);
  const d = res.issues.filter(i => i.code === 'erp-ats-differs');
  assert.deepEqual(d.map(i => i.style), ['AB100', 'AB102']);
  assert.match(d[0].message, /"ATS from in WH" is 0, but In WH 0 − Total Customer Orders 300 = -300\. The tool uses -300\./);
  assert.match(d[1].message, /rows 5, 6/);
  assert.equal(res.counts.erpAtsDiffers, 2);
});

test('negative In WH (physical stock below 0) is listed and kept as it is', () => {
  const res = parse([TITLE, HEAD, r('AB100', -7000, 0, { ats: '-7000' }), r('AB101', 7100, 0), grand(100)]);
  assert.equal(res.ok, true, res.error);
  assert.equal(res.styles.get('AB100').inWh, -7000);
  const neg = res.issues.filter(i => i.code === 'negative-in-wh');
  assert.equal(neg.length, 1);
  assert.equal(neg[0].style, 'AB100');
  assert.match(neg[0].message, /In WH is -7,000 \(row 3\)/);
});

test('a row with units but no style is listed and still part of the Grand Total check', () => {
  const res = parse([TITLE, HEAD, r('AB100', 100, 0), r('', 25, 0), r('', 0, 0), grand(125)]);
  assert.equal(res.ok, true, res.error);
  assert.equal(res.styles.size, 1);
  assert.equal(res.counts.noStyleRows, 2);
  const ns = res.issues.filter(i => i.code === 'no-style');
  assert.equal(ns.length, 1, 'an empty-looking row is only counted');
  assert.match(ns[0].message, /Row 4 has no style but 25 units in the warehouse/);
});

test('header found by name in the first 10 rows; the "Report" sheet, else the first sheet', () => {
  const rows = [TITLE, [], ['Run by someone'], HEAD, r('AB100', 10, 0), grand(10)];
  const notes = { Notes: [['nothing here']] };
  const res = parse(rows, NAME, 'Report', notes);   // Notes is the first sheet, Report the second
  assert.equal(res.ok, true, res.error);
  assert.equal(res.sheet, 'Report');
  assert.equal(res.headerRow, 4);
  assert.equal(res.styles.size, 1);
  const other = parse(rows, NAME, 'Sheet A');
  assert.equal(other.ok, true, other.error);
  assert.equal(other.sheet, 'Sheet A', 'no Report sheet: the first sheet is read');
  const wrong = parse(rows, NAME, 'Sheet A', notes);
  assert.equal(wrong.ok, false, 'no Report sheet and the first sheet is something else: fails');
  assert.match(wrong.error, /column in the first 10 rows of sheet "Notes"/);
});

test('a required column missing, or no style rows: the report fails with a plain message', () => {
  const noOrders = [TITLE, HEAD.map(h => (h === 'Total Customer Orders' ? 'Orders?' : h)), r('AB100', 10, 0), grand(10)];
  const res = parse(noOrders);
  assert.equal(res.ok, false);
  assert.match(res.error, /no "Total Customer Orders" column in row 2/);
  assert.deepEqual(res.missing, ['Total Customer Orders']);
  const deep = [TITLE, [], [], [], [], [], [], [], [], [], HEAD, r('AB100', 10, 0)];
  const res2 = parse(deep);
  assert.equal(res2.ok, false, 'a header below row 10 is not found');
  assert.match(res2.error, /no "Style", "In WH", "Total Customer Orders" column in the first 10 rows/);
  const empty = parse([TITLE, HEAD, grand(0)]);
  assert.equal(empty.ok, false);
  assert.match(empty.error, /no style rows below the header \(row 2\)/);
  assert.equal(A.parseAts(X, { SheetNames: [] }, NAME, TODAY).ok, false);
});

test("the report's day comes from its file name (m-d-yyyy, with or without the time), else none", () => {
  assert.equal(A.dayFromName(NAME), C.ymd(2026, 10, 29));
  assert.equal(A.dayFromName('Inventory Report 1-5-2027.xlsx'), C.ymd(2027, 1, 5));
  assert.equal(A.dayFromName('Inventory Report - On Hand_In Transit_In Production 11-02-2026 9-05 AM (1).xlsx'), C.ymd(2026, 11, 2));
  assert.equal(A.dayFromName('Inventory Report.xlsx'), null);
  assert.equal(A.dayFromName('Report 13-45-2026.xlsx'), null, 'not a date');
  const res = parse([TITLE, HEAD, r('AB100', 10, 0), grand(10)], 'ats.xlsx');
  assert.equal(res.ok, true);
  assert.equal(res.reportDay, null);
  assert.equal(res.reportDayFrom, 'none');
});

test('In Trans and Import PO Backlog lists are read for cross-checks only, never added to stock', () => {
  const rows = [TITLE, HEAD,
    r('AB100', 100, 0, { inTrans: '1200, 3400', backlog: '5000', eta: '11/14/2026, 12/1/2026', etd: '10/20/2026' }),
    r('AB101', 0, 0, { inTrans: '1,200, 300' }),
    grand(100)];
  const res = parse(rows);
  assert.equal(res.ok, true, res.error);
  const s = res.styles.get('AB100');
  assert.equal(s.inWh, 100);
  assert.equal(s.inTrans, 4600);
  assert.equal(s.backlog, 5000);
  assert.deepEqual(s.rows[0].eta, [C.ymd(2026, 11, 14), C.ymd(2026, 12, 1)]);
  assert.deepEqual(s.rows[0].etd, [C.ymd(2026, 10, 20)]);
  assert.equal(res.styles.get('AB101').inTrans, 1500, 'a thousands comma stays inside its number');
});

test('a Grand Total (or Total) row is never a style, whichever cell holds the label', () => {
  // The label under Style (a layout with other columns first), and a bare "Total" under Style.
  const underStyle = grand(1500, 1), bare = grand(1500, 1); bare[1] = 'Total';
  for (const g of [underStyle, bare]) {
    const res = parse([TITLE, HEAD, r('AB100', 1000, 0), r('AB101', 500, 0), g]);
    assert.equal(res.ok, true, res.error);
    assert.deepEqual([...res.styles.keys()], ['AB100', 'AB101'], `"${g[1]}" is not a style`);
    assert.equal(res.counts.inWh, 1500);
    assert.equal(res.counts.grandTotalRows, 1);
    assert.deepEqual(codes(res), []);
  }
  // In a later column too (Style Description), and in any case.
  const late = grand(1500, 2); late[2] = 'GRAND TOTAL';
  const res = parse([TITLE, HEAD, r('AB100', 1000, 0), r('AB101', 500, 0), late]);
  assert.deepEqual([res.ok, res.styles.size, res.counts.grandTotalRows], [true, 2, 1]);
});

test('negative Total Customer Orders (a credit or return) is listed and kept as it is', () => {
  const res = parse([TITLE, HEAD, r('AB100', 100, -50), r('AB101', 10, 0), grand(110)]);
  assert.equal(res.ok, true, res.error);
  assert.equal(res.styles.get('AB100').orders, -50);
  assert.deepEqual(codes(res), ['negative-orders']);
  assert.equal(res.issues[0].style, 'AB100');
  assert.match(res.issues[0].message, /Total Customer Orders is -50 \(row 3\), below zero/);
  assert.equal(res.counts.negativeOrders, 1);
});

test("the fields found are listed, so a blank In Trans can be told from a missing column", () => {
  const res = parse([TITLE, HEAD, r('AB100', 100, 0), grand(100)]);
  assert.ok(res.columns.includes('inTrans') && res.columns.includes('style') && res.columns.includes('inWh'), res.columns.join(','));
  const head = HEAD.filter(h => h !== 'In Trans Qty'), row = r('AB100', 100, 0).filter((_, i) => HEAD[i] !== 'In Trans Qty');
  const g = grand(100).filter((_, i) => HEAD[i] !== 'In Trans Qty');
  const res2 = parse([TITLE, head, row, g]);
  assert.equal(res2.ok, true, res2.error);
  assert.ok(!res2.columns.includes('inTrans'));
});
