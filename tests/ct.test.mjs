import test from 'node:test';
import assert from 'node:assert/strict';
import { core, loadXLSX, xlsxBytes, readBook } from './load.mjs';

const C = core('common');
const CT = core('ct');
const X = loadXLSX();

// Invented data only. Rows are written as { header: value } so the same rows can be laid out in any column order.
const BOOK_H = ['Category', 'Sub', 'Brand', 'Product Name', 'Style', 'Barcode', 'Inner', 'Master', 'Price', 'Quantity', null,
  'PO Number', 'Company', null, 'PO Amount ', null, 'Notes', 'CBM', 'ready Date'];
const LOAD_H = ['Category', 'Sub ', 'Brand', 'Product name ', 'Style', 'Barcode', 'Inner', 'Master', 'Price', 'Quantity',
  'Old Style name', 'Po Number', 'Company', 'Order Date', 'PO Amount', 'Ship by', 'Notes', 'Ready Date ', 'Need by'];
const CONT_H = ['ID#', 'Container #', 'Round #', 'ETD', 'ETA', 'Factory', 'PO#', 'P.L/ C.I.', 'Sent Docs Forwarder', 'telex',
  'Freight released', 'Customs Cleared', 'FDA Released', 'DO Sent to trucker', 'Packing list sent to WH', 'Notes'];

const ser = (y, m, d) => C.ymd(y, m, d) + 25569; // Excel serial, as a typed date is stored
const day = s => C.dayOf(s);
const TODAY = C.ymd(2026, 10, 28);
const lay = (header, rows) => [header, ...rows.map(o => o == null ? [] : Array.isArray(o) ? o : header.map(h => h == null ? null : (o[h.trim()] ?? null)))];

const BOOKING = [
  { Style: 'AB-100', Quantity: 1200, 'PO Number': 900101, Company: 'Factory A' }, // row 2: before any section
  { Style: 'AB-101', Quantity: 600 },
  { Category: 'Round 350' },
  { Style: 'AB-200', Barcode: '000111222333', Master: 48, Quantity: 2400, 'PO Number': 900102, Company: ' Factory b ', 'ready Date': ser(2026, 10, 1) },
  { Style: 'ab-201 ', Barcode: 'N/A', Quantity: '1,200' },
  null, // row 7: blank
  { Category: 'Round 351-352' },
  { Style: 'AB-300', Quantity: 500, 'PO Number': '900103 -bcf' },
  { Style: 'AB-301', Quantity: 50, 'PO Number': '900103-BCF' }, // same PO, same section: fine
  { Style: 'AB-202', Quantity: 10, 'PO Number': 900102 }, // row 11: PO already under Round 350
  { Category: 'Planning' },
  { Style: 'AB-400', Quantity: 100 }, // row 13: no PO above it
  { Style: 'AB-500', Quantity: 'TBD', 'PO Number': 900104 }, // row 14
  { Style: 'AB-501' }, // row 15: blank quantity
  { 'Product Name': 'Total', 'PO Amount': 999 }, // row 16: not a line
  { 'Product Name': 'Loose', Quantity: 50 }, // row 17: quantity with no style
  { Category: 'Round TBD' }, // row 18
  null, null, null, [''], // empty formatted rows below the data
];
const LOADED = [
  { Category: 'Loaded' },
  { Category: 'Round 349' },
  { Style: 'AB-600', Quantity: 3000, 'Po Number': 900105, Company: 'Factory C', 'Order Date': ser(2026, 7, 1), 'Ship by': '9/15', 'Ready Date': 'soon', 'Need by': 'Nov' },
  { Category: 'Round 350' },
  { Style: 'AB-200', Quantity: 2400, 'Po Number': 900102 }, // row 6: also on Booking, round 350
  { Category: 'Round 360' }, // row 7: nothing on Open or Complete
  { Style: 'AB-700', Quantity: 10, 'Po Number': 900106 },
];
const OPEN = [
  null,
  { 'ID#': 1, 'Container #': 'TEST1234567', 'Round #': 350, ETD: ser(2026, 10, 1), ETA: ser(2026, 11, 5), Factory: 'Factory B', 'PO#': '900102' },
  { 'ID#': 2, 'Container #': 'test7654321 / TEST1111111', 'Round #': 350, ETD: ser(2026, 10, 3), ETA: '11/9', 'Freight released': 'NA', 'Customs Cleared': 'hold' },
  { 'Container #': 'TEST2222222', 'Round #': '351-352', ETA: '1/5' }, // no ETD, yearless ETA long past
  { 'Container #': 'TEST3333333' }, // row 6: no round
  { 'Container #': 'x', 'Round #': 'TBD' }, // row 7
];
const COMPLETE = [
  { 'ID#': 'Complete Shipments' },
  { 'ID#': 3, 'Container #': 'TEST4444444', 'Round #': 349, ETA: '7/15', 'Packing list sent to WH': '7/20' },
  { 'ID#': 4, 'Container #': 'Air Shipment', 'Round #': 348, ETD: ser(2026, 6, 1), ETA: ser(2026, 6, 10) },
];

const DONE = [{ 'Container #': 'TEST0000009', 'Round #': 300, ETD: ser(2026, 5, 1), ETA: ser(2026, 6, 1) }];

function book(over = {}) {
  const sheets = {
    'Booking vs Master PO Comp': [['PO', 'On Both sheets']],
    Booking: lay(over.bookH || BOOK_H, over.booking || BOOKING),
    Loaded: lay(over.loadH || LOAD_H, over.loaded || LOADED),
    Open: lay(over.openH || CONT_H, over.open || OPEN),
    Complete: lay(over.completeH || CONT_H, over.complete || COMPLETE),
  };
  for (const k of over.drop || []) delete sheets[k];
  return readBook(xlsxBytes(sheets));
}
const parse = (over, today = TODAY) => CT.parseContainerTracking(X, book(over), today);
const codes = r => r.issues.reduce((o, i) => (o[i.code] = (o[i.code] || 0) + 1, o), {});

test('round headers', () => {
  assert.deepEqual(CT.parseRounds('Round 302-306'), [302, 303, 304, 305, 306]);
  assert.deepEqual(CT.parseRounds('Round 298 & 299'), [298, 299]);
  assert.deepEqual(CT.parseRounds('Round 358 incomplete'), [358]);
  assert.deepEqual(CT.parseRounds(' round 145 - sent to someone '), [145]);
  assert.deepEqual(CT.parseRounds('Round 360 - 2nd box'), []); // not a range, and no number left: listed by the parser
  assert.deepEqual(CT.parseRounds('Planning'), []);
  assert.deepEqual(CT.parseRounds('Round TBD'), []);
});

test('blocks and lines from Booking and Loaded', () => {
  const r = parse();
  assert.equal(r.ok, true, r.error);
  const bk = r.blocks.filter(b => b.sheet === 'Booking');
  assert.deepEqual(bk.map(b => [b.row, b.po, b.section, b.rounds]), [
    [2, '900101', null, []], [5, '900102', 'Round 350', [350]], [9, '900103-BCF', 'Round 351-352', [351, 352]],
    [10, '900103-BCF', 'Round 351-352', [351, 352]], [11, '900102', 'Round 351-352', [351, 352]], [14, '900104', 'Planning', []]]);
  const b5 = bk[1];
  assert.equal(b5.factory, 'Factory b'); // as written, ends trimmed, never mapped
  assert.equal(b5.readyDate, day('10/1/2026'));
  assert.equal(b5.orderDate, null); // Booking's order date column has no header
  assert.equal(b5.shipBy, null);
  assert.equal(b5.needBy, '');
  assert.deepEqual(b5.lines, [
    { row: 5, style: 'AB-200', styleRaw: 'AB-200', upc: '000111222333', upcRaw: '000111222333', master: 48, qty: 2400 },
    { row: 6, style: 'AB-201', styleRaw: 'ab-201', upc: null, upcRaw: 'N/A', master: null, qty: 1200 }]);
  assert.equal(bk[2].poRaw, '900103 -bcf');
  assert.deepEqual(bk[5].lines, []); // its quantity is text: listed, never 0
  assert.ok(r.blocks.every(b => b.lines.every(l => typeof l.qty === 'number' && !Number.isNaN(l.qty))));

  const ld = r.blocks.filter(b => b.sheet === 'Loaded');
  assert.deepEqual(ld.map(b => [b.row, b.po, b.rounds]), [[4, '900105', [349]], [6, '900102', [350]], [8, '900106', [360]]]);
  assert.equal(ld[0].orderDate, day('7/1/2026'));
  assert.equal(ld[0].shipBy, day('9/15/2026')); // yearless: this year, listed
  assert.equal(ld[0].readyDate, null); // "soon": listed
  assert.equal(ld[0].needBy, 'Nov');
});

test('every row read gets a result', () => {
  const { counts } = parse();
  assert.deepEqual(counts.Booking, { rowsRead: 17, blankRows: 1, sectionRows: 4, styleRows: 10, counted: 7, notCounted: 3, otherRows: 2, blocks: 6, units: 5960 });
  assert.deepEqual(counts.Loaded, { rowsRead: 7, blankRows: 0, sectionRows: 4, styleRows: 3, counted: 3, notCounted: 0, otherRows: 0, blocks: 3, units: 5410 });
  assert.deepEqual(counts.Open, { rowsRead: 6, blankRows: 1, labelRows: 0, roundRows: 3, noRoundRows: 2, rounds: 3 });
  assert.deepEqual(counts.Complete, { rowsRead: 3, blankRows: 0, labelRows: 1, roundRows: 2, noRoundRows: 0, rounds: 2 });
  for (const s of ['Booking', 'Loaded']) {
    const n = counts[s];
    assert.equal(n.rowsRead, n.blankRows + n.sectionRows + n.styleRows + n.otherRows);
    assert.equal(n.styleRows, n.counted + n.notCounted);
  }
  for (const s of ['Open', 'Complete']) {
    const n = counts[s];
    assert.equal(n.rowsRead, n.blankRows + n.labelRows + n.roundRows + n.noRoundRows);
  }
});

test('Open and Complete keep every row under each of its rounds', () => {
  const { rounds } = parse();
  assert.deepEqual(Object.keys(rounds.open).map(Number).sort(), [350, 351, 352]);
  assert.equal(rounds.open[350].length, 2);
  const [a, b] = rounds.open[350];
  assert.deepEqual(a.containers, ['TEST1234567']);
  assert.equal(a.etd, day('10/1/2026'));
  assert.equal(a.eta, day('11/5/2026'));
  assert.equal(a.factoryText, 'Factory B');
  assert.equal(a.poText, '900102');
  assert.equal(a.id, '1');
  assert.deepEqual(b.containers, ['TEST7654321', 'TEST1111111']);
  assert.equal(b.eta, day('11/9/2026'));
  assert.equal(b.etaYearless, true);
  assert.equal(b.freightReleased, null); // "NA": nothing to report, not listed
  assert.equal(b.customsCleared, null); // "hold": listed
  assert.equal(rounds.open[351][0], rounds.open[352][0]);
  assert.equal(rounds.open[351][0].roundText, '351-352');
  const c = rounds.complete[349][0];
  assert.equal(c.sheet, 'Complete');
  assert.equal(c.row, 3);
  assert.equal(c.plSentToWh, day('7/20/2026'));
  assert.deepEqual(rounds.complete[348][0].containers, []);
  assert.equal(rounds.complete[348][0].containerText, 'Air Shipment');
});

test('issues raised, each with a reason', () => {
  const r = parse();
  assert.deepEqual(codes(r), {
    'po-two-sections': 1, 'no-po': 1, 'qty-not-number': 2, 'qty-no-style': 1, 'no-style': 1, 'round-no-number': 1, 'yearless-date': 5,
    'not-a-date': 2, 'no-round': 1, 'round-not-number': 1, 'round-no-open-complete': 1, 'on-both-tabs': 1 });
  const at = (code) => r.issues.filter(i => i.code === code).map(i => [i.sheet, i.row]);
  assert.deepEqual(at('po-two-sections'), [['Booking', 11]]);
  assert.deepEqual(at('no-po'), [['Booking', 13]]);
  assert.deepEqual(at('qty-not-number'), [['Booking', 14], ['Booking', 15]]);
  assert.deepEqual(at('qty-no-style'), [['Booking', 17]]);
  assert.deepEqual(at('no-style'), [['Booking', 16]]); // a product name and an amount, but no style or quantity
  assert.deepEqual(at('round-no-number'), [['Booking', 18]]);
  assert.deepEqual(at('round-no-open-complete'), [['Loaded', 8]]);
  assert.deepEqual(at('no-round'), [['Open', 6]]);
  assert.deepEqual(at('round-not-number'), [['Open', 7]]);
  assert.deepEqual(at('not-a-date'), [['Loaded', 4], ['Open', 4]]);
  assert.deepEqual(at('yearless-date'), [['Loaded', 4], ['Open', 4], ['Open', 5], ['Complete', 3], ['Complete', 3]]);
  const both = r.issues.find(i => i.code === 'on-both-tabs');
  assert.deepEqual([both.po, both.style, both.round, both.row, both.loadedRow], ['900102', 'AB-200', 350, 5, 6]);
  assert.ok(r.issues.every(i => i.source === 'Container Tracking' && i.message.length > 10));
  assert.match(r.issues.find(i => i.code === 'qty-not-number').message, /"TBD" is not a number/);
});

test('columns are found by header name, in any order', () => {
  const base = parse();
  const shuffled = parse({
    bookH: ['Notes', 'PO Number', 'Quantity', 'Inserted', 'Style', 'Company', 'Category', 'Product Name', 'Barcode', 'Master', 'ready Date', 'PO Amount'],
    openH: ['Notes', 'ETA', 'ETD', 'Round #', 'Container #', 'ID#', 'Customs Cleared', 'Freight released', 'Packing list sent to WH', 'Factory', 'PO#'],
    completeH: ['Round #', 'Container #', 'ETA', 'ETD', 'ID#', 'Packing list sent to WH'],
  });
  assert.equal(shuffled.ok, true, shuffled.error);
  assert.deepEqual(shuffled.blocks, base.blocks);
  assert.deepEqual(shuffled.rounds, base.rounds);
  assert.deepEqual(shuffled.counts, base.counts);
});

test('a missing column or tab stops the parse and names it', () => {
  const r = parse({ bookH: BOOK_H.map(h => h === 'Quantity' ? 'Amount' : h), openH: CONT_H.filter(h => h !== 'ETA') });
  assert.equal(r.ok, false);
  assert.match(r.error, /Booking tab has no "Quantity" column/);
  assert.match(r.error, /headers seen in row 1: Category, Sub, Brand/);
  assert.match(r.error, /Open tab has no "ETA" column/);
  assert.deepEqual(r.errorDetail.map(d => [d.sheet, d.missing]), [['Booking', ['Quantity']], ['Open', ['ETA']]]);
  assert.ok(r.errorDetail[0].seen.includes('Amount'));
  assert.deepEqual(r.blocks, []);

  const t = parse({ drop: ['Complete'] });
  assert.equal(t.ok, false);
  assert.match(t.error, /no "Complete" tab \(tabs seen: Booking vs Master PO Comp, Booking, Loaded, Open\)/);
});

test('a file with no lines is a failure, not "nothing in transit"', () => {
  const r = parse({ booking: [], loaded: [{ Category: 'Loaded' }] });
  assert.equal(r.ok, false);
  assert.match(r.error, /no lines were read from Booking or Loaded/);
  const o = parse({ open: [], complete: [{ 'ID#': 'Complete Shipments' }] });
  assert.equal(o.ok, false);
  assert.match(o.error, /no container rows were read/);
});

test('dates typed without a year, around New Year', () => {
  const jan10 = C.ymd(2027, 1, 10);
  const open = [
    { 'Container #': 'TEST0000001', 'Round #': 400, ETD: '12/20', ETA: '1/15' },
    { 'Container #': 'TEST0000002', 'Round #': 401, ETA: '12/28' },
    { 'Container #': 'TEST0000003', 'Round #': 402, ETD: ser(2026, 12, 1), ETA: '12/30' },
  ];
  const r = parse({ open }, jan10);
  const rec = n => r.rounds.open[n][0];
  assert.equal(rec(400).etd, day('12/20/2026')); // more than 60 days ahead -> last year
  assert.equal(rec(400).eta, day('1/15/2027')); // before its ETD -> next year
  assert.equal(rec(401).eta, day('12/28/2026')); // more than 120 days ahead -> last year
  assert.equal(rec(402).eta, day('12/30/2026')); // takes the ETD's year
  assert.ok(rec(400).etdYearless && rec(400).etaYearless && !rec(402).etdYearless);
  const yl = r.issues.filter(i => i.code === 'yearless-date' && i.sheet === 'Open');
  assert.equal(yl.length, 4);
  assert.match(yl[0].message, /ETD "12\/20" has no year; read as 12\/20\/2026/);
});

test('a yearless ETA with no ETD, long past: next year on Open, kept on Complete', () => {
  const r = parse();
  assert.equal(r.rounds.open[351][0].eta, day('1/5/2027'));
  const c = parse({ complete: [{ 'Container #': 'TEST0000004', 'Round #': 349, ETA: '1/5' }] });
  assert.equal(c.rounds.complete[349][0].eta, day('1/5/2026'));
  assert.equal(c.rounds.complete[349][0].etaYearless, true);
});

test('the same PO + style + round on both tabs is listed once', () => {
  const booking = [{ Category: 'Round 350' },
    { Style: 'AB-200', Quantity: 100, 'PO Number': 900102 }, { Style: 'AB-200', Quantity: 50 }];
  const loaded = [{ Category: 'Round 349-350' }, { Style: 'ab-200', Quantity: 150, 'Po Number': '900102' },
    { Category: 'Round 351' }, { Style: 'AB-200', Quantity: 5, 'Po Number': '900102' }];
  const r = parse({ booking, loaded, complete: DONE, open: [{ 'Container #': 'TEST0000005', 'Round #': '349-351', ETD: ser(2026, 10, 1), ETA: ser(2026, 11, 1) }] });
  assert.equal(r.ok, true, r.error);
  assert.deepEqual(codes(r), { 'on-both-tabs': 1 });
});

test('a PO repeated on Loaded, or in one Booking section, is not listed', () => {
  const booking = [{ Category: 'Round 350' }, { Style: 'AB-1', Quantity: 1, 'PO Number': 900107 }, { Style: 'AB-2', Quantity: 1, 'PO Number': 900107 }];
  const loaded = [{ Category: 'Round 350' }, { Style: 'AB-3', Quantity: 1, 'Po Number': 900108 },
    { Category: 'Round 351' }, { Style: 'AB-4', Quantity: 1, 'Po Number': 900108 }];
  const r = parse({ booking, loaded, complete: DONE, open: [{ 'Container #': 'TEST0000006', 'Round #': '350-351', ETD: ser(2026, 10, 1), ETA: ser(2026, 11, 1) }] });
  assert.deepEqual(codes(r), {});
});

// Open and Complete rows the tests below reuse: rounds 300 and 301 are on Complete, 302 on Open.
const ROUNDS_DONE = [{ 'Container #': 'TEST0000009', 'Round #': '300-301', ETD: ser(2026, 5, 1), ETA: ser(2026, 6, 1) }];
const OPEN_ONE = [{ 'Container #': 'TEST0000010', 'Round #': 302, ETD: ser(2026, 10, 1), ETA: ser(2026, 11, 1) }];

test('a row read as a section header that looks like a product line is listed', () => {
  const loaded = [
    { Category: 'Lip', Style: 'AB-1', Quantity: 10, 'Po Number': 900101 },
    { Category: 'Round 300' },
    { Style: 'AB-2', Quantity: 5, 'Po Number': 900102 },
    { Category: 'Lip', 'Product name': 'Gloss', Price: 1.5 }, // row 5: a product line missing style and quantity
    { Style: 'AB-3', Quantity: 5, 'Po Number': 900103 },
    { Category: 'Round 301', Quantity: 999 }, // row 7: a round header with a subtotal in it
    { Style: 'AB-4', Quantity: 5, 'Po Number': 900104 },
    { Category: 'Lip' }, // row 9: only a Category, but product lines use "Lip"
    { Style: 'AB-5', Quantity: 5, 'Po Number': 900105 },
  ];
  const r = parse({ loaded, complete: ROUNDS_DONE, open: OPEN_ONE });
  assert.equal(r.ok, true, r.error);
  const odd = r.issues.filter(i => i.code === 'odd-section');
  assert.deepEqual(odd.map(i => [i.sheet, i.row]), [['Loaded', 5], ['Loaded', 7], ['Loaded', 9]]);
  assert.match(odd[0].message, /also has Product name, Price/);
  assert.match(odd[1].message, /looks like a round header.*stay under "Lip"/);
  // What is read is unchanged: a section is still "Category filled, Style and Quantity empty".
  assert.deepEqual(r.blocks.filter(b => b.sheet === 'Loaded').map(b => [b.po, b.section]),
    [['900101', null], ['900102', 'Round 300'], ['900103', 'Lip'], ['900104', 'Lip'], ['900105', 'Lip']]);
});

test('product lines with no style, rows with no round, and bad barcodes are listed with their rows', () => {
  const loaded = [{ Category: 'Round 300' },
    { Style: 'AB-1', Quantity: 10, 'Po Number': 900101, Barcode: '12345678901' }, // row 3: 11 digits: listed, still counted
    { Style: 'AB-2', Quantity: 10, Barcode: 'N/A' }, // read as blank
    { 'Product name': 'Gloss', Price: 2, 'PO Amount': 20 }, // row 5: no style, no quantity
    { 'Po Number': 900102 }, // row 6: a PO row with its lines below: not listed
    { Style: 'AB-3', Quantity: 1 }];
  const open = [...OPEN_ONE, { 'ID#': 7, Factory: 'Factory A', 'PO#': '900101' }, { 'ID#': 'In transit' }];
  const r = parse({ loaded, open, complete: ROUNDS_DONE });
  const at = code => r.issues.filter(i => i.code === code).map(i => [i.sheet, i.row]);
  assert.deepEqual(at('bad-upc'), [['Loaded', 3]]);
  assert.deepEqual(at('no-style').filter(x => x[0] === 'Loaded'), [['Loaded', 5]]);
  assert.deepEqual(at('no-round'), [['Open', 3]]); // ID#, Factory and PO# but no round
  assert.equal(r.counts.Open.labelRows, 1); // "In transit" alone in ID#: a title
  assert.equal(r.counts.Loaded.counted, 3);
  assert.equal(r.blocks[0].lines[0].upc, null);
});

test('a yearless date on Complete is never after today', () => {
  const complete = [
    { 'Container #': 'TEST0000011', 'Round #': 300, ETA: '12/3', 'Freight released': '12/5' },
    { 'Container #': 'TEST0000012', 'Round #': 301, ETD: '11/20', ETA: '12/20' },
  ];
  const r = parse({ complete, open: OPEN_ONE });
  const a = r.rounds.complete[300][0], b = r.rounds.complete[301][0];
  assert.equal(a.eta, day('12/3/2025'));
  assert.equal(a.freightReleased, day('12/5/2025'));
  assert.equal(b.etd, day('11/20/2025'));
  assert.equal(b.eta, day('12/20/2025'));
  assert.match(r.issues.find(i => i.code === 'yearless-date' && i.row === 2).message, /read as 12\/3\/2025/);
  // On Open the same ETA is still to come.
  const o = parse({ open: [{ 'Container #': 'TEST0000013', 'Round #': 302, ETA: '12/3' }], complete: ROUNDS_DONE });
  assert.equal(o.rounds.open[302][0].eta, day('12/3/2026'));
});

test('dates far from today, odd round headers, bad containers and no ETA on Open are listed', () => {
  const loaded = [{ Category: 'Round 300 - sent 11/2' }, // row 2: 11 and 2 are a date, not rounds
    { Style: 'AB-1', Quantity: 1, 'Po Number': 900101, 'Ready Date': ser(2028, 3, 1) }, // row 3: two years out
    { Category: 'Round 301 & 302' }, // two rounds: fine
    { Style: 'AB-2', Quantity: 1, 'Po Number': 900102 }];
  const open = [{ 'Container #': 'TEST0000014', 'Round #': 302, ETD: ser(2026, 10, 1) }, // row 2: no ETA yet
    { 'Container #': 'ABCD123456', 'Round #': '303 (2 boxes)', ETD: ser(2026, 10, 1), ETA: ser(2026, 11, 1) }]; // row 3
  const complete = [{ 'Container #': 'TEST0000015', 'Round #': '300-301', ETD: ser(2026, 5, 1), ETA: ser(2026, 6, 1), 'Freight released': ser(2010, 6, 1) }];
  const r = parse({ loaded, open, complete });
  const at = code => r.issues.filter(i => i.code === code).map(i => [i.sheet, i.row]);
  assert.deepEqual(r.blocks.filter(b => b.sheet === 'Loaded').map(b => b.rounds), [[300], [301, 302]]);
  assert.deepEqual(at('round-odd'), [['Loaded', 2], ['Open', 3]]);
  assert.deepEqual(Object.keys(r.rounds.open).map(Number), [302, 303]);
  assert.deepEqual(at('date-far'), [['Loaded', 3], ['Complete', 2]]);
  assert.deepEqual(at('no-eta'), [['Open', 2]]);
  assert.deepEqual(at('bad-container'), [['Open', 3]]);
  assert.deepEqual(CT.readRounds('Round 371 - Sent to someone to update'), { rounds: [371], odd: false });
  assert.deepEqual(CT.readRounds('Round 302-306'), { rounds: [302, 303, 304, 305, 306], odd: false });
});

test('a PO + style on both tabs where one side is not in a round is listed; a split across rounds is not', () => {
  const booking = [{ Category: 'Unshipped' }, { Style: 'AB-1', Quantity: 5, 'PO Number': 900101 },
    { Category: 'Round 301' }, { Style: 'AB-2', Quantity: 5, 'PO Number': 900102 }];
  const loaded = [{ Category: 'Round 300' }, { Style: 'AB-1', Quantity: 5, 'Po Number': 900101 },
    { Style: 'AB-2', Quantity: 5, 'Po Number': 900102 }];
  const r = parse({ booking, loaded, complete: ROUNDS_DONE, open: OPEN_ONE });
  const x = r.issues.filter(i => i.code === 'on-both-tabs-other-round');
  assert.deepEqual(x.map(i => [i.sheet, i.row, i.po, i.style, i.loadedRow]), [['Booking', 3, '900101', 'AB-1', 3]]);
});

test('the page can load it after common.js', () => {
  assert.equal(globalThis.InvCT, CT);
  assert.equal(typeof globalThis.InvCT.parseContainerTracking, 'function');
  assert.throws(() => CT.parseContainerTracking(X, book(), undefined), /today must be a day number/);
});
