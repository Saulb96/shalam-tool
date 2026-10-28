import test from 'node:test';
import assert from 'node:assert/strict';
import { core, loadXLSX, xlsxBytes, readBook } from './load.mjs';

const C = core('common');
const S = core('shipping');
const X = loadXLSX();

const day = s => C.dayOf(s);
const serial = s => day(s) + 25569;           // an Excel date cell
const TODAY = day('10/28/2026');
const codes = issues => issues.map(i => i.code).sort();

// ------------------------------------------------------------ PO text

test('poKeys expands free-text PO cells like the Late & Short Function', () => {
  assert.deepEqual(S.poKeys('900529A-B-C'), ['900529A', '900529B', '900529C']);
  assert.deepEqual(S.poKeys('900420A---900420D'), ['900420A', '900420B', '900420C', '900420D']);
  assert.deepEqual(S.poKeys('900502ABCDEG'), ['900502A', '900502B', '900502C', '900502D', '900502E', '900502G']);
  assert.deepEqual(S.poKeys('900052A-D'), ['900052A', '900052B', '900052C', '900052D']);
  assert.deepEqual(S.poKeys('900107A/B'), ['900107A', '900107B']);
  assert.deepEqual(S.poKeys('9004360'), ['900436O']);                        // zero typed for the letter O
  assert.deepEqual(S.poKeys('1234567'), []);                                 // 7 digits not ending in 0: not a PO
  assert.deepEqual(S.poKeys('900453A/900453B/9 - 900453A/900453B/900453C'), ['900453A', '900453B', '900453C']);
  assert.deepEqual(S.poKeys('900101, 900102\n900103a-b'), ['900101', '900102', '900103A', '900103B']);
  assert.deepEqual(S.poKeys(''), []);
  assert.deepEqual(S.poKeys(null), []);
});

test('poKeys keeps -BCF / -1 suffixes whole and prefers POs the tracker knows', () => {
  assert.deepEqual(S.poKeys('900447-BCF'), ['900447-BCF']);
  assert.deepEqual(S.poKeys('900447-bcf/900508-1'), ['900447-BCF', '900508-1']);
  assert.deepEqual(S.poKeys('900502AB'), ['900502A', '900502B']);
  assert.deepEqual(S.poKeys('900502AB', new Set(['900502AB'])), ['900502AB']);
  assert.deepEqual(S.poKeys('900529A-B', ['900529A-B']), ['900529A-B']);   // an array works too
});

// ------------------------------------------------------------ forwarder

const FWD_HEAD = ['ID', 'Status', 'Mode', 'ETD', 'ETA', 'Final ETA', 'Supplier Name', 'Container Numbers', 'PO Numbers',
  'Customs Released', 'Freight Released', 'Cargo Ready Date', 'Cartons', 'Total Cartons', 'LFD', 'Discharged', 'Picked Up Destination'];
const H = Object.fromEntries(FWD_HEAD.map((h, i) => [h, i]));
function fwdRow(o) {
  const r = FWD_HEAD.map(() => null);
  for (const [k, v] of Object.entries(o)) r[H[k]] = v;
  return r;
}
const fwdBook = rows => readBook(xlsxBytes({ Sheet1: [['Active Shipment Report', null, 'printed 10/28/2026'], FWD_HEAD, ...rows] }));

test('forwarder: header on row 2, one record per shipment, containers lined up with their date lines', () => {
  const wb = fwdBook([
    FWD_HEAD.map(() => ' '),                                   // a row of spaces under the header
    fwdRow({ ID: 5001, Status: 'Sailed', Mode: 'FCL', ETD: serial('10/1/2026'), ETA: serial('11/1/2026'), 'Final ETA': serial('11/3/2026'),
      'Supplier Name': 'Factory A\nFactory B', 'Container Numbers': 'AAAU1000001\nAAAU1000002', 'PO Numbers': '900101/900102A-B\n900103',
      'Customs Released': 'Yes', 'Freight Released': 'No', 'Cargo Ready Date': serial('9/25/2026'), Cartons: '120\n80', 'Total Cartons': 200,
      LFD: '2026-11-10\n', Discharged: '2026-11-04\n2026-11-05', 'Picked Up Destination': null }),
    fwdRow({ ID: 5002, Status: 'Pending Departure', ETD: serial('11/5/2026'), ETA: serial('12/5/2026'), 'PO Numbers': '900104' }),
    fwdRow({ ID: 5003, Status: 'Soon To Arrive', ETD: serial('10/5/2026'), 'Final ETA': serial('11/2/2026'),
      'Container Numbers': '\nAAAU1000003', LFD: '2026-11-20\n2026-11-21', 'PO Numbers': '900105' }),
  ]);
  const r = S.parseForwarder(X, wb, TODAY);
  assert.equal(r.ok, true, r.error);
  assert.equal(r.headerRow, 2);
  assert.ok(r.headers.includes('Container Numbers'));
  assert.deepEqual(r.counts, { rowsRead: 4, records: 3, containers: 3, blank: 1, noId: 0, noContainer: 1 });
  assert.deepEqual(r.issues, []);

  const a = r.records[0];
  assert.equal(a.id, '5001');
  assert.equal(a.row, 4);
  assert.equal(a.status, 'Sailed');
  assert.equal(a.mode, 'FCL');
  assert.equal(a.etd, day('10/1/2026'));
  assert.equal(a.eta, day('11/1/2026'));
  assert.equal(a.finalEta, day('11/3/2026'));
  assert.equal(a.cargoReady, day('9/25/2026'));
  assert.deepEqual(a.containers, ['AAAU1000001', 'AAAU1000002']);
  assert.deepEqual(a.perContainer.map(e => [e.container, e.lfd, e.discharged, e.pickedUp]), [
    ['AAAU1000001', day('11/10/2026'), day('11/4/2026'), null],
    ['AAAU1000002', null, day('11/5/2026'), null],        // a blank line is a container with no LFD yet
  ]);
  assert.equal(a.poText, '900101/900102A-B\n900103');
  assert.equal(a.customsReleased, true);
  assert.equal(a.freightReleased, false);
  assert.deepEqual(a.cartons, [120, 80]);
  assert.equal(a.totalCartons, 200);
  assert.equal(a.supplier, 'Factory A\nFactory B');

  const pending = r.records[1];
  assert.equal(pending.pseudo, true);
  assert.deepEqual(pending.containers, ['#5002']);
  assert.equal(pending.customsReleased, null);           // blank is not known, not "no"
  assert.equal(pending.totalCartons, null);

  // The container on the second line takes the second line of the date columns.
  assert.equal(r.records[2].perContainer[0].lfd, day('11/21/2026'));

  const v = r.byContainer.get('AAAU1000001');
  assert.equal(v.eta, day('11/3/2026'));                 // Final ETA, else ETA
  assert.deepEqual(v.ids, ['5001']);
  assert.deepEqual(v.poTexts, ['900101/900102A-B\n900103']);
  assert.ok(r.byContainer.has('#5002'));
  assert.equal(r.byContainer.get('#5002').eta, day('12/5/2026'));
  assert.deepEqual(r.unknownStatuses, []);
});

test('forwarder: two containers on one line fall back to their order in the cell', () => {
  const wb = fwdBook([fwdRow({ ID: 5010, Status: 'Arrived', ETD: serial('9/1/2026'), ETA: serial('10/1/2026'),
    'Container Numbers': 'AAAU1000011, AAAU1000012', LFD: '2026-10-08\n2026-10-09', 'PO Numbers': '900110' })]);
  const r = S.parseForwarder(X, wb, TODAY);
  assert.deepEqual(r.records[0].perContainer.map(e => [e.container, e.lfd]),
    [['AAAU1000011', day('10/8/2026')], ['AAAU1000012', day('10/9/2026')]]);
});

test('forwarder: a reused container number keeps the later trip; one shared within 14 days is read as one', () => {
  const wb = fwdBook([
    fwdRow({ ID: 6001, Status: 'Sailed', ETD: serial('10/20/2026'), ETA: serial('11/20/2026'), 'Container Numbers': 'AAAU2000001', 'PO Numbers': '900201' }),
    fwdRow({ ID: 6002, Status: 'Arrived', ETD: serial('8/1/2026'), ETA: serial('9/1/2026'), 'Container Numbers': 'AAAU2000001', 'PO Numbers': '900202',
      LFD: '2026-09-08' }),
    fwdRow({ ID: 6003, Status: 'Arrived', ETD: serial('7/1/2026'), ETA: serial('8/1/2026'), 'Container Numbers': 'AAAU2000002', 'PO Numbers': '900203' }),
    fwdRow({ ID: 6004, Status: 'Sailed', ETD: serial('10/1/2026'), ETA: serial('11/1/2026'), 'Container Numbers': 'AAAU2000002', 'PO Numbers': '900204' }),
    fwdRow({ ID: 6005, Status: 'Sailed', ETD: serial('10/2/2026'), ETA: serial('11/2/2026'), 'Container Numbers': 'AAAU2000003', 'PO Numbers': '900205',
      'Customs Released': 'Yes', 'Freight Released': 'Yes', LFD: '2026-11-12' }),
    fwdRow({ ID: 6006, Status: 'Sailed', ETD: serial('10/3/2026'), ETA: serial('11/4/2026'), 'Container Numbers': 'AAAU2000003', 'PO Numbers': '900206',
      'Customs Released': 'No', 'Freight Released': 'Yes' }),
  ]);
  const r = S.parseForwarder(X, wb, TODAY);
  assert.equal(r.ok, true, r.error);
  // the earlier trip listed after the later one is skipped
  assert.deepEqual(r.byContainer.get('AAAU2000001').ids, ['6001']);
  assert.equal(r.records[1].perContainer[0].olderTrip, true);
  assert.equal(r.records[0].perContainer[0].olderTrip, false);
  // the later trip listed after the earlier one replaces it
  assert.deepEqual(r.byContainer.get('AAAU2000002').ids, ['6004']);
  assert.equal(r.records[2].perContainer[0].olderTrip, true);
  // consolidated: POs from both, the later row's status and ETA, released only when both are
  const m = r.byContainer.get('AAAU2000003');
  assert.deepEqual(m.ids, ['6005', '6006']);
  assert.deepEqual(m.poTexts, ['900205', '900206']);
  assert.equal(m.eta, day('11/4/2026'));
  assert.equal(m.lfd, day('11/12/2026'));
  assert.equal(m.customsReleased, false);
  assert.equal(m.freightReleased, true);
  assert.deepEqual(codes(r.issues), ['container-on-two-shipments', 'container-reused', 'container-reused']);
  assert.equal(r.counts.containers, 3);
});

test('forwarder: every row gets a result; problems are listed, never guessed', () => {
  const wb = fwdBook([
    fwdRow({ ID: 7001, Status: 'Arrived', ETD: serial('9/1/2026'), ETA: serial('10/1/2026'), 'PO Numbers': '900301' }),     // no container
    fwdRow({ ID: 7002, Status: 'Delayed', ETD: 'soon', ETA: serial('11/1/2026'), 'Container Numbers': 'AAAU3000001\nTBA', 'PO Numbers': '900302',
      'Customs Released': 'Partly', Cartons: '100\nabout 50', 'Total Cartons': 'n/a' }),
    fwdRow({ Status: 'Sailed', 'PO Numbers': '900303' }),                                                                    // no ID
    fwdRow({ ID: 7001, Status: '', ETD: serial('9/1/2026'), ETA: '11/2', 'Container Numbers': 'AAAU3000002', 'PO Numbers': '900304' }),
  ]);
  const r = S.parseForwarder(X, wb, TODAY);
  assert.equal(r.ok, true, r.error);
  const c = r.counts;
  assert.equal(c.rowsRead, c.blank + c.noId + c.records);
  assert.deepEqual(c, { rowsRead: 4, records: 3, containers: 2, blank: 0, noId: 1, noContainer: 1 });
  assert.deepEqual(codes(r.issues), ['bad-container', 'bad-date', 'bad-number', 'bad-number', 'bad-yes-no', 'duplicate-id',
    'no-container', 'no-id', 'no-status', 'unknown-status', 'yearless-date']);
  assert.deepEqual(r.unknownStatuses, ['Delayed']);
  const bad = r.records[1];
  assert.equal(bad.etd, null);
  assert.equal(bad.customsReleased, false);
  assert.deepEqual(bad.cartons, [100, null]);
  assert.equal(bad.totalCartons, null);
  assert.equal(r.records[2].eta, day('11/2/2026'));      // yearless: today's year, and listed
  for (const i of r.issues) { assert.equal(i.source, 'Forwarder'); assert.ok(i.row >= 3 && i.message); }
});

test('forwarder and drayage: a date over a year from today, or before 2015, is kept and listed', () => {
  const wb = fwdBook([fwdRow({ ID: 7101, Status: 'Sailed', ETD: serial('9/1/2026'), ETA: serial('10/1/2029'), 'Container Numbers': 'AAAU3000003', 'PO Numbers': '900305' })]);
  const r = S.parseForwarder(X, wb, TODAY);
  assert.deepEqual(codes(r.issues), ['date-far']);
  assert.equal(r.records[0].eta, day('10/1/2029'));
  assert.match(r.issues[0].message, /reads as 10\/1\/2029, more than a year from today or before 2015/);
  const d = S.parseDrayage(['Container#,ETA,Availability,Actual Pick Up Date,Actual Delivery Date', 'AAAU3000004,10/20/2014,AVAILABLE,,'].join('\n'), TODAY);
  assert.deepEqual(codes(d.issues), ['date-far']);
});

test('forwarder: missing columns are fatal and name the column and the headers seen', () => {
  const head = FWD_HEAD.filter(h => h !== 'PO Numbers');
  let r = S.parseForwarder(X, readBook(xlsxBytes({ Sheet1: [head, [1, 'Sailed']] })), TODAY);
  assert.equal(r.ok, false);
  assert.match(r.error, /"PO Numbers"/);
  assert.match(r.error, /Headers seen on row 1: ID, Status/);
  assert.deepEqual(r.missing, ['PO Numbers']);

  const noEta = FWD_HEAD.filter(h => h !== 'ETA' && h !== 'Final ETA');
  r = S.parseForwarder(X, readBook(xlsxBytes({ Sheet1: [noEta] })), TODAY);
  assert.equal(r.ok, false);
  assert.deepEqual(r.missing, ['Final ETA or ETA']);

  const etaOnly = FWD_HEAD.filter(h => h !== 'Final ETA');
  r = S.parseForwarder(X, readBook(xlsxBytes({ Sheet1: [etaOnly, [1, 'Sailed', null, serial('10/1/2026'), serial('11/1/2026'), null,
    'AAAU4000001', '900401']] })), TODAY);
  assert.equal(r.ok, true, r.error);
  assert.equal(r.byContainer.get('AAAU4000001').eta, day('11/1/2026'));

  r = S.parseForwarder(X, readBook(xlsxBytes({ Sheet1: [['Report'], ['ID', 'Status', 'Containers']] })), TODAY);
  assert.equal(r.ok, false);
  assert.match(r.error, /Container Numbers/);
  assert.deepEqual(r.seen, ['ID', 'Status', 'Containers']);
});

test('forwarder: an empty report, or one with no container numbers, is not ok', () => {
  let r = S.parseForwarder(X, fwdBook([]), TODAY);
  assert.equal(r.ok, false);
  assert.match(r.error, /no shipments/);
  r = S.parseForwarder(X, fwdBook([fwdRow({ ID: 8001, Status: 'Pending Departure', ETD: serial('11/9/2026'), ETA: serial('12/9/2026') })]), TODAY);
  assert.equal(r.ok, false);
  assert.match(r.error, /no container numbers/);
  assert.equal(r.counts.records, 1);
});

// ------------------------------------------------------------ CSV

test('parseCsv follows RFC 4180', () => {
  const text = '﻿a,"b, with comma","say ""hi"""\r\n1,"two\nlines",3\r\n\r\nx,="0123",y"z\n"",last';
  assert.deepEqual(S.parseCsv(text), [
    ['a', 'b, with comma', 'say "hi"'],
    ['1', 'two\nlines', '3'],
    [''],                                   // a blank line is a row, so row numbers match Excel
    ['x', '="0123"', 'y"z'],                // a quote that doesn't open a field is kept as written
    ['', 'last'],
  ]);
  assert.deepEqual(S.parseCsv('a,b\n'), [['a', 'b']]);
  assert.deepEqual(S.parseCsv('a,\n'), [['a', '']]);
  assert.deepEqual(S.parseCsv('"a"b,c'), [['ab', 'c']]);
  assert.deepEqual(S.parseCsv(''), []);
  assert.throws(() => S.parseCsv('a,"never closed\n1,2'), /quoted value/);
});

// ------------------------------------------------------------ drayage

const DRAY_HEAD = ['Customer Name', 'Container#', 'BL Number', 'Ref#', 'ETA', 'LFD', 'Availability', 'Terminal', 'Deliver To',
  'Scheduled Pick Up Date', 'Actual Pick Up Date', 'Proposed Deliver Date', 'Actual Delivery Date', 'Empty Notify Date',
  'Empty Return Date', 'Remarks'];
const q = v => (v === '' || /^=/.test(v) ? v : '"' + String(v).replace(/"/g, '""') + '"');
const drayCsv = rows => [DRAY_HEAD.map(q).join(','), ...rows.map(r => r.map(q).join(','))].join('\n') + '\n';
const drow = (o) => DRAY_HEAD.map(h => o[h] ?? '');

test('drayage: one record per container, ="..." unwrapped, availability uppercased', () => {
  const text = drayCsv([
    drow({ 'Customer Name': 'Customer 1', 'Container#': 'aaau5000001', 'BL Number': 'BL0001', 'Ref#': '="900501"', ETA: '10/20/2026',
      LFD: '10/27/2026', Availability: 'Delivered', Terminal: 'T1', 'Scheduled Pick Up Date': '10/22/2026', 'Actual Pick Up Date': '10/22/2026',
      'Proposed Deliver Date': '10/23/2026', 'Actual Delivery Date': '10/23/2026', 'Empty Return Date': '10/25/2026' }),
    drow({ 'Container#': 'AAAU5000002', 'Ref#': '900502/900503', ETA: '11/2/2026', Availability: 'NOT AVAILABLE' }),
  ]);
  const r = S.parseDrayage(text, TODAY);
  assert.equal(r.ok, true, r.error);
  assert.equal(r.headerRow, 1);
  assert.deepEqual(r.counts, { rowsRead: 2, records: 2, blank: 0, noContainer: 0, duplicates: 0 });
  assert.deepEqual(r.issues, []);
  assert.deepEqual(r.records[0], {
    row: 2, container: 'AAAU5000001', ref: '900501', eta: day('10/20/2026'), lfd: day('10/27/2026'), availability: 'DELIVERED',
    terminal: 'T1', scheduledPickup: day('10/22/2026'), actualPickup: day('10/22/2026'), proposedDelivery: day('10/23/2026'),
    actualDelivery: day('10/23/2026'), emptyReturn: day('10/25/2026'), bl: 'BL0001',
  });
  assert.equal(r.records[1].actualPickup, null);
  assert.equal(r.records[1].ref, '900502/900503');
});

test('drayage: a container on two rows keeps the latest ETA; every row gets a result', () => {
  const text = drayCsv([
    drow({ 'Container#': 'AAAU6000001', ETA: '11/1/2026', Availability: 'AVAILABLE' }),
    drow({ 'Container#': 'AAAU6000001', ETA: '6/1/2026', Availability: 'DELIVERED' }),       // older trip: dropped
    drow({ 'Container#': 'AAAU6000002', ETA: '9/1/2026', Availability: 'DELIVERED' }),
    drow({ 'Container#': 'AAAU6000002', ETA: '9/1/2026', Availability: 'Acknowledged' }),    // same ETA: the later row wins
    drow({ 'Container#': 'not yet', ETA: '11/9/2026' }),
    drow({}),
    drow({ 'Container#': 'AAAU6000003', ETA: 'next week', LFD: '11/3', Availability: '' }),
  ]) + ',"AAAU6000004",,,"11/5/2026",,"AVAILABLE",,,,,,,,,,,\n';     // 18 values under 16 headers
  const r = S.parseDrayage(text, TODAY);
  assert.equal(r.ok, true, r.error);
  const c = r.counts;
  assert.deepEqual(c, { rowsRead: 8, records: 4, blank: 1, noContainer: 1, duplicates: 2 });
  assert.equal(c.rowsRead, c.blank + c.noContainer + c.records + c.duplicates);
  const by = Object.fromEntries(r.records.map(x => [x.container, x]));
  assert.equal(by.AAAU6000001.availability, 'AVAILABLE');
  assert.equal(by.AAAU6000002.row, 5);
  assert.equal(by.AAAU6000003.eta, null);
  assert.equal(by.AAAU6000003.lfd, day('11/3/2026'));
  assert.deepEqual(r.unknownAvailability, ['ACKNOWLEDGED']);
  assert.deepEqual(codes(r.issues), ['bad-date', 'duplicate-container', 'duplicate-container', 'field-count', 'no-availability',
    'no-container', 'unknown-availability', 'yearless-date']);
  assert.equal(r.issues.find(i => i.code === 'field-count').row, 9);
});

test('drayage: missing columns, an empty file and a broken file are not ok', () => {
  let r = S.parseDrayage(drayCsv([]).replace('"Actual Delivery Date"', '"Delivered On"'), TODAY);
  assert.equal(r.ok, false);
  assert.match(r.error, /"Actual Delivery Date"/);
  assert.match(r.error, /Headers seen on row 1: Customer Name, Container#/);
  assert.deepEqual(r.missing, ['Actual Delivery Date']);

  r = S.parseDrayage(drayCsv([]), TODAY);
  assert.equal(r.ok, false);
  assert.match(r.error, /no container rows/);

  assert.equal(S.parseDrayage('', TODAY).ok, false);
  r = S.parseDrayage(drayCsv([]) + '"AAAU7000001,10/1/2026\n', TODAY);
  assert.equal(r.ok, false);
  assert.match(r.error, /quoted value/);
});

// ------------------------------------------------------------ newest usable file

test('pickNewest: newest first, lock files and other types skipped, ages in New York days', () => {
  const files = [
    { id: '1', name: 'report-a.xlsx', lastModifiedDateTime: '2026-10-23T15:00:00Z' },   // Friday
    { id: '2', name: '~$report-b.xlsx', lastModifiedDateTime: '2026-10-28T12:00:00Z' },
    { id: '3', name: 'report-c.XLSX', lastModifiedDateTime: '2026-10-28T03:00:00Z' },   // 11 pm on 10/27 in New York
    { id: '4', name: 'notes.csv', lastModifiedDateTime: '2026-10-28T12:00:00Z' },
    { id: '5', name: 'old.xlsx', lastModifiedDateTime: '2026-10-19T12:00:00Z' },
    { id: '6', name: 'archive.xlsx', folder: { childCount: 3 }, lastModifiedDateTime: '2026-10-28T13:00:00Z' },
  ];
  const c = S.pickNewest(files, '.xlsx', TODAY);
  assert.deepEqual(c.map(x => x.name), ['report-c.XLSX', 'report-a.xlsx', 'old.xlsx']);
  assert.deepEqual(c.map(x => [x.ageDays, x.businessDaysOld, x.tooOld, x.amber]), [
    [1, 1, false, false],
    [5, 3, false, true],
    [9, 7, true, true],
  ]);
  assert.equal(c[0].day, day('10/27/2026'));
  assert.equal(c[0].item, files[2]);
  assert.deepEqual(S.pickNewest(files, '.csv', TODAY).map(x => x.id), ['4']);
  assert.deepEqual(S.pickNewest([], '.csv', TODAY), []);
  // Friday's report read on Monday is one business day old: not amber
  assert.equal(S.pickNewest([files[0]], '.xlsx', day('10/26/2026'))[0].amber, false);
});
