// The page in headless Edge against a stand-in for SharePoint: a fake sign-in (never a real one), a fake Graph that
// answers from invented workbooks, a test config.js with invented locations, and checks on the chips, banners, the
// error email, the read-only rules and the tabs. The step 2 scenarios (s2-…) compare every stage number the page shows
// with core/stages.js run here on the same files; the Available to sell scenarios (sell-…, ats…) compare the selling
// view with core/ats.js and core/sell.js run here on the same files, and check Saul's pull-from date rule lot by lot
// on their own (10 days before the ETD; else 5 days before the CRD; else about 75 days after the order date).
//   node tests/headless/run.mjs [scenario,scenario…]      (default: all but s2-real and sell-real)
//   REAL_DIR=<folder> [REAL_LISTS=<lists.json>] node tests/headless/run.mjs s2-real
//   REAL_DIR=<folder> REAL_ATS=<ATS report.xlsx> node tests/headless/run.mjs sell-real
// s2-real and sell-real run only when named: the page on local copies of the real files (served on 127.0.0.1 only;
// numbers printed, nothing else). <folder> as for tools/check-real.mjs; the lists default to <folder>/lists.json;
// sell-real serves the ATS report as saved yesterday, with yesterday's date in its name.
// Edge runs InPrivate, with sync off and a new throwaway profile; each scenario gets its own fresh browser context;
// every request that isn't to 127.0.0.1 is blocked. Screenshots go to .tmp/. Exits 1 when any check fails.
import { createServer } from 'node:http';
import { createServer as netServer } from 'node:net';
import { spawn } from 'node:child_process';
import { readFileSync, writeFileSync, existsSync, mkdirSync, mkdtempSync, rmSync, statSync } from 'node:fs';
import { createHash } from 'node:crypto';
import { tmpdir } from 'node:os';
import { join, resolve, sep, extname } from 'node:path';
import { loadXLSX, core, root, xlsxBytes } from '../load.mjs';

const EDGE = process.env.EDGE_PATH || 'C:/Program Files (x86)/Microsoft/Edge/Application/msedge.exe';
const REPO = resolve(root());
const OUT = join(REPO, '.tmp');
const GRAPH = 'https://graph.microsoft.com/v1.0';
const MAIL_TO = 'saul@example.com';
// Invented locations: two sharing links and four /sites/… paths, on sites and libraries the stand-in below knows.
const TEST_WEB = 'https://contoso-test.sharepoint.example';
const LOC = {
  msl: `${TEST_WEB}/:x:/s/TestSite/EaTestMasterSheetLink001?e=test01`,
  tracker: `${TEST_WEB}/:x:/s/TestSite/EaTestTrackerLink0002?e=test02`,
  ct: '/sites/TestSite/Shared Documents/Logistics/Test Container Sheet.xlsx',
  forwarder: '/sites/TestReports/Shared Documents/Shipping/Forwarder',
  drayage: '/sites/TestReports/Shared Documents/Shipping/Drayage',
  ats: '/sites/TestSite/Shared Documents/Stock/ATS',
};
const CT_V2 = '/sites/TestSite/Shared Documents/Logistics/Test Container Sheet v2.xlsx';
const CT_PATH = '/drives/drv-src/root:/Logistics/Test Container Sheet.xlsx';   // how Container Tracking is read: by its path in its library
// Container Tracking sections with no round, with invented names (the real ones are in config.js).
const TEST_SECTIONS = { pullFromChina: ['Customer Direct'], planning: ['Planning A'], ignore: ['Hidden Rows Test'] };
const TEST_CONFIG = { tenantId: '00000000-0000-4000-8000-000000000000', clientId: '11111111-1111-4111-8111-111111111111',
  sources: LOC, sections: TEST_SECTIONS, errorEmailTo: MAIL_TO, errorEmailName: 'Saul' };
const ACCOUNT = { name: 'Test User', username: 'test.user@example.com', homeAccountId: 'test-home', localAccountId: 'test-local', tenantId: TEST_CONFIG.tenantId };
// As a real Graph token lists them: the sign-in scopes (email, openid, profile) come along with the three asked for.
const READ_SCP = 'email Mail.Send openid profile Sites.Read.All User.Read';
const SITE_ID = 'site-prod', SRC = 'drv-src', REP = 'drv-rep';
const XOPT = { type: 'array', cellFormula: false, cellHTML: false, cellStyles: false };   // as app.js reads files
const KEYS = ['msl', 'tracker', 'pobook', 'ct', 'forwarder', 'drayage', 'ats', 'upc'];
const NAMES = { msl: 'Master Sheet Live', tracker: 'PO Status Tracker', pobook: 'PO Book lines', ct: 'Container Tracking',
  forwarder: 'Forwarder', drayage: 'Drayage', ats: 'ATS', upc: 'UPC_MASTER' };

const X = loadXLSX();
const C = core('common'), PB = core('pobook'), CT = core('ct'), SH = core('shipping'), UPC = core('upc'), ST = core('stages');
const ATS = core('ats'), SELL = core('sell');
const TODAY = C.todayNY(Date.now());   // the page works its dates out from the same clock
const iso = days => new Date(Date.now() - days * 864e5).toISOString();
const fmtN = v => (typeof v === 'number' ? v.toLocaleString('en-US') : String(v));
const sleep = ms => new Promise(r => setTimeout(r, ms));
const sha = t => createHash('sha256').update(t, 'utf8').digest('hex');

// ------------------------------------------------------------------ invented source files

function buildFixtures(today) {
  const bytes = new Map();
  const TH = ['PO#', 'Supplier', 'Order date', 'Item Description', 'Style #', 'Item Type', 'Order Type', 'Customer',
    'Customer 1st Ship Date', 'CRD Expected', 'Factory CRD', 'Delayed CRD'];
  const t = (po, sty, type = 'Stock', cust = 'Customer 1') => [po, 'Factory A', '6/1/2026', 'Item ' + sty, sty, 'Cosmetics', type, cust, '12/1/2026', '11/1/2026', '11/5/2026', null];
  bytes.set('tracker.xlsx', xlsxBytes({ 'PO Status Tracker': [TH, t(900100, 'ZZ-900'), t(900101, 'AB-100'), t(900101, 'AB-200'),
    t(900102, 'CD-300', 'China Pickup', 'Customer 2'), t(900103, 'GH-500'), t('bulk', 'QQ-1'), []] }));

  const MH = ['Category', 'Sub', 'Brand', 'Product name', 'Style Number', 'Barcode', 'Inner', 'Master', 'Price', 'Quantity',
    'Old Style name', 'Po Number', 'Factory', 'Order Date', 'PO Amount', 'Ship by', 'Notes'];
  const m = (name, sty, qty, po, upc) => ['Lip', 'Gloss', 'Brand X', name, sty, upc || null, 6, 48, 1.5, qty, null, po, po ? 'Factory A' : null, po ? '6/1/2026' : null, null, '11/1/2026', null];
  const mslRows = [
    m('Old cream', 'ZZ-900', 240, '900100', '000000000031'),
    m('Round 1', null, null, null),
    m('Lip gloss', 'AB-100', 480, '900101', '000000000017'),
    m('Lip gloss 2', 'AB-200', 96, null, '000000000024'),
    m('Brush', 'CD-300', 'TBD', '900102'),
    m('Bag', 'EF-400', 12, null, '000000000048'),
    [],
    m('Sample bag', null, 5, null)];
  bytes.set('msl.xlsx', xlsxBytes({ 'Master Sheet Live': [MH, ...mslRows] }));
  bytes.set('msl-bad.xlsx', xlsxBytes({ 'Master Sheet Live': [MH.map(h => (h === 'Quantity' ? 'Qty' : h)), ...mslRows] }));   // no "Quantity"

  // Container Tracking, laid out like the real one (Booking's order date and ship by columns have no header).
  const BOOK = ['Category', 'Sub', 'Brand', 'Product Name', 'Style', 'Barcode', 'Inner', 'Master', 'Price', 'Quantity', null,
    'PO Number', 'Company', null, 'PO Amount', null, 'Notes'];
  const LOAD = ['Category', 'Sub', 'Brand', 'Product name', 'Style', 'Barcode', 'Inner', 'Master', 'Price', 'Quantity',
    'Old Style name', 'Po Number', 'Company', 'Order Date', 'PO Amount', 'Ship by', 'Notes', 'Ready Date', 'Need by'];
  const CONT = ['ID#', 'Container #', 'Round #', 'ETD', 'ETA', 'Factory', 'PO#', 'P.L/ C.I.', 'Sent Docs Forwarder', 'telex',
    'Freight released', 'Customs Cleared', 'FDA Released', 'DO Sent to trucker', 'Packing list sent to WH', 'Notes'];
  const row = (h, o) => h.map(x => (x == null ? null : (o[x] ?? null)));
  const ser = (y, mo, d) => Date.UTC(y, mo - 1, d) / 864e5 + 25569;
  const ct = qtyHeader => {
    const bh = BOOK.map(h => (h === 'Quantity' ? qtyHeader : h));
    return xlsxBytes({
      'Booking vs Master PO Comp': [['PO', 'Style']],
      Booking: [bh, row(bh, { Category: 'Round 351' }),
        row(bh, { Style: 'AB-100', Barcode: '000000000017', Master: 48, [qtyHeader]: 480, 'PO Number': 900101, Company: 'Factory A' }),
        row(bh, { Style: 'AB-200', Barcode: '000000000024', Master: 48, [qtyHeader]: 96 })],
      Loaded: [LOAD, row(LOAD, { Category: 'Loaded' }), row(LOAD, { Category: 'Round 350' }),
        row(LOAD, { Style: 'ZZ-900', Barcode: '000000000031', Master: 24, Quantity: 240, 'Po Number': 900100, Company: 'Factory B', 'Ready Date': ser(2026, 9, 1) })],
      Open: [CONT, row(CONT, { 'ID#': 1, 'Container #': 'AAAU1234567', 'Round #': 351, ETD: ser(2026, 11, 2), ETA: ser(2026, 12, 2), Factory: 'Factory A', 'PO#': '900101' })],
      Complete: [CONT, row(CONT, { 'ID#': 'Complete Shipments' }),
        row(CONT, { 'ID#': 2, 'Container #': 'BBBU7654321', 'Round #': 350, ETD: ser(2026, 9, 5), ETA: ser(2026, 10, 5), Factory: 'Factory B', 'PO#': '900100' })],
    });
  };
  bytes.set('ct.xlsx', ct('Quantity'));
  bytes.set('ct-bad.xlsx', ct('Pieces'));   // Booking has no "Quantity" column

  const FWD = ['ID', 'Status', 'Mode', 'ETD', 'ETA', 'Final ETA', 'Supplier Name', 'Container Numbers', 'PO Numbers', 'Customs Released', 'Freight Released'];
  bytes.set('fwd.xlsx', xlsxBytes({ Sheet1: [['Active Shipment Report'], FWD,
    [101, 'Sailed', 'FCL', '11/2/2026', '12/2/2026', '12/3/2026', 'Factory A', 'AAAU1234567', '900101', 'No', 'No'],
    [102, 'Pending Departure', 'FCL', '11/20/2026', '12/20/2026', null, 'Factory B', null, '900102', null, null]] }));

  const DH = ['Customer Name', 'Container#', 'BL Number', 'Ref#', 'ETA', 'LFD', 'Availability', 'Terminal', 'Deliver To',
    'Scheduled Pick Up Date', 'Actual Pick Up Date', 'Proposed Deliver Date', 'Actual Delivery Date', 'Empty Notify Date', 'Empty Return Date', 'Remarks'];
  const DR = ['Customer 1,BBBU7654321,BL0001,="900100",10/5/2026,10/9/2026,DELIVERED,Terminal 1,Warehouse 1,10/6/2026,10/6/2026,10/7/2026,10/7/2026,,10/9/2026,',
    'Customer 1,CCCU1111111,BL0002,="900101-BCF",10/20/2026,10/24/2026,AVAILABLE,Terminal 1,Warehouse 1,,,,,,,'];
  bytes.set('dray.csv', Buffer.from([DH.join(','), ...DR].join('\n') + '\n'));
  // Step 2: drayage says round 351's container is at the port (the forwarder says it sailed); and one with no
  // "Availability" column, which fails.
  const portRow = 'Customer 1,AAAU1234567,BL0003,="900101",12/2/2026,12/6/2026,AVAILABLE,Terminal 1,Warehouse 1,,,,,,,';
  bytes.set('dray-port.csv', Buffer.from([DH.join(','), ...DR, portRow].join('\n') + '\n'));
  bytes.set('dray-bad.csv', Buffer.from([DH.map(h => (h === 'Availability' ? 'Avail' : h)).join(','), ...DR, portRow].join('\n') + '\n'));

  const pob = [
    { id: '1', lastModifiedDateTime: iso(2), fields: { Title: 'Settings', LineKey: '__settings', Kind: 'settings', Data: '{"la":30}' } },
    { id: '2', lastModifiedDateTime: iso(3), fields: { Title: 'Line', LineKey: '900101|AB-200', Kind: 'line', PONumber: '900101', Style: 'AB-200', Shipped: false } },
    // the newer item for the same line wins: AB-200 marked shipped by hand
    { id: '3', lastModifiedDateTime: iso(1), fields: { Title: 'Line', LineKey: '900101|AB-200', Kind: 'line', PONumber: '900101', Style: 'AB-200', Shipped: true, ShippedDate: '2026-10-01', ShippedBy: 'Test User' } },
    { id: '4', lastModifiedDateTime: iso(1), fields: { Title: 'No line key' } },
  ];
  const u = (id, sty, upc, desc) => ({ id, fields: { Title: sty, UPC: upc, Style: sty, Description: desc, Category: 'Lip', SubCategory: 'Gloss', Brand: 'Brand X' } });
  const upcPages = [[u('1', 'AB-100', '000000000017', 'Lip gloss'), u('2', 'AB-200', '000000000024', 'Lip gloss 2'), u('3', 'CD-300', '12345', 'Brush')],
    [u('4', 'ZZ-900', '000000000031', 'Old cream'), u('5', 'EF-400', '000000000048', 'Bag')]];

  // Step 2, a big invented set: 2,400 open styles over 240 POs (every 20th PO a China Pickup), 600 of them on a Booking
  // round, 400 on a Loaded round that has sailed, one Container Tracking line with no Master Sheet Live line, one style
  // on two POs with two UPCs, some with no CRD, no Master or a part case.
  {
    const N = 2400, PER = 10, pad = i => String(i).padStart(4, '0');
    const crd = i => (i % 7 === 0 ? null : i % 3 === 0 ? '8/1/2026' : '3/1/2027');
    const brows = [MH, m('Shipped cream', 'SH-1', 240, '900200', '000000009991'), m('Round 1', null, null, null)];
    const trows = [TH, [900200, 'Factory A', '6/1/2026', 'Item SH-1', 'SH-1', 'Cosmetics', 'Stock', 'Customer 1', '12/1/2026', '9/1/2026', '9/1/2026', null]];
    for (let i = 0; i < N; i++) {
      const po = 900200 + Math.floor(i / PER), sty = 'ST' + pad(i), fac = 'Factory ' + 'ABCD'[i % 4];
      brows.push(['Cat ' + (i % 5), 'Sub', 'Brand X', 'Item ' + i, sty, String(100000000000 + i), 6, i % 50 === 7 ? null : 24, 1.5, i % 11 === 3 ? 1210 : 1200, null, String(po), fac, '6/1/2026', null, crd(i) ? '11/1/2026' : null, null]);
      const type = Math.floor(i / PER) % 20 === 0 ? 'China Pickup' : 'Stock';
      trows.push([po, fac, '6/1/2026', 'Item ' + sty, sty, 'Cosmetics', type, 'Customer ' + ((i % 3) + 1), '12/1/2026', crd(i), crd(i), null]);
    }
    brows.push(['Cat 1', 'Sub', 'Brand X', 'Item 1b', 'ST0001', '200000000001', 6, 24, 1.5, 600, null, '900450', 'Factory B', '6/1/2026', null, '11/1/2026', null]);
    trows.push([900450, 'Factory B', '6/1/2026', 'Item ST0001', 'ST0001', 'Cosmetics', 'Stock', 'Customer 1', '12/1/2026', '3/1/2027', '3/1/2027', null]);
    bytes.set('msl-big.xlsx', xlsxBytes({ 'Master Sheet Live': brows }));
    bytes.set('trk-big.xlsx', xlsxBytes({ 'PO Status Tracker': trows }));
    const bk = [BOOK, row(BOOK, { Category: 'Round 360' })];
    for (let i = 0; i < 600; i++) bk.push(row(BOOK, { Style: 'ST' + pad(i), Barcode: String(100000000000 + i), Master: 24, Quantity: 1200, 'PO Number': 900200 + Math.floor(i / PER), Company: 'Factory A' }));
    // A 13-digit UPC: the "500 unmatched" tag must still show after it in the 200px Style cell.
    bk.push(row(BOOK, { Style: 'UN-1', Barcode: '1234567890128', Master: 24, Quantity: 500, 'PO Number': 900999, Company: 'Factory Z' }));
    // Sections with no round (the test config.js's invented names): straight to the customer (PO 900301, one line with
    // 300 more than the sheet), planned (PO 900302), hidden (PO 900303, and a style with no line); on Loaded, one not on
    // the lists (PO 900304), and a PO below the tracker's lowest (900200) in round 361.
    const stl = (i, q, fac) => row(BOOK, { Style: 'ST' + pad(i), Barcode: String(100000000000 + i), Master: 24, Quantity: q, 'PO Number': 900200 + Math.floor(i / PER), Company: fac });
    bk.push(row(BOOK, { Category: 'Customer Direct' }));
    for (let i = 1010; i < 1015; i++) bk.push(stl(i, i === 1014 ? 1500 : 1200, 'Factory C'));
    bk.push(row(BOOK, { Category: 'Planning A' }));
    for (let i = 1020; i < 1025; i++) bk.push(stl(i, 1200, 'Factory C'));
    bk.push(row(BOOK, { Category: 'Hidden Rows Test' }), stl(1030, 1200, 'Factory C'),
      row(BOOK, { Style: 'HID-1', Master: 24, Quantity: 240, 'PO Number': 900303, Company: 'Factory C' }));
    const ld = [LOAD, row(LOAD, { Category: 'Loaded' }), row(LOAD, { Category: 'Round 361' })];
    for (let i = 600; i < 1000; i++) ld.push(row(LOAD, { Style: 'ST' + pad(i), Barcode: String(100000000000 + i), Master: 24, Quantity: 1200, 'Po Number': 900200 + Math.floor(i / PER), Company: 'Factory B' }));
    ld.push(row(LOAD, { Style: 'OLD-1', Master: 24, Quantity: 700, 'Po Number': 900150, Company: 'Factory B' }));
    ld.push(row(LOAD, { Category: 'Odd Section' }), row(LOAD, { Style: 'ST1040', Barcode: '100000001040', Master: 24, Quantity: 1200, 'Po Number': 900304, Company: 'Factory B' }));
    bytes.set('ct-big.xlsx', xlsxBytes({
      'Booking vs Master PO Comp': [['PO', 'Style']], Booking: bk, Loaded: ld,
      Open: [CONT, row(CONT, { 'ID#': 1, 'Container #': 'DDDU2222222', 'Round #': 360, ETD: ser(2027, 1, 5), ETA: ser(2027, 2, 5), Factory: 'Factory A' }),
        row(CONT, { 'ID#': 2, 'Container #': 'AAAU1234567', 'Round #': 361, ETD: ser(2026, 11, 2), ETA: ser(2026, 12, 2), Factory: 'Factory B' })],
      Complete: [CONT, row(CONT, { 'ID#': 'Complete Shipments' }), row(CONT, { 'ID#': 3, 'Container #': 'EEEU3333333', 'Round #': 300, ETD: ser(2026, 6, 5), ETA: ser(2026, 7, 5) })],
    }));
  }

  // ---- Available to sell. The ERP's weekly ATS report as the real one is laid out: sheet "Report", a title row, the
  // header, one row per style and warehouse (numbers sometimes stored as text), and a "Grand Total" row the rows' In WH
  // must add up to (o.grand: another total, so it doesn't).
  const AH = ['Color Code', 'Style', 'Style Description', 'Category', 'Sub Category', 'Warehouse', 'In WH', 'Total Customer Orders', 'ATS from in WH',
    'ATS from In Trans', 'In Trans Qty', 'Import PO Backlog Qty', 'ETA Date', 'ETD Date'];
  const ar = (sty, inWh, ord, ats, wh = 'LA') => ['C1', sty, 'Desc ' + sty, 'Lip', 'Gloss', wh, inWh, ord,
    ats !== undefined ? ats : typeof inWh === 'number' && typeof ord === 'number' ? inWh - ord : null, null, null, null, null, null];
  const atsBook = (rows, o = {}) => {
    const sum = j => rows.reduce((s, r) => s + (Number(String(r[j] ?? '').replace(/,/g, '')) || 0), 0);
    return xlsxBytes({ Report: [['Inventory Report - On Hand_In Transit_In Production (invented)'], o.header || AH, ...rows,
      ['Grand Total', null, null, null, null, null, o.grand ?? sum(6), sum(7), null, null, null, null, null, null]] });
  };
  // The standard small files: AB-100's orders take its lot on the water and more; ZZ-900 on two warehouses; QQ-77 on
  // the report only.
  const smallAts = [ar('AB-100', 100, 700), ar('AB-200', '50', 0, '50'), ar('ZZ-900', 1000, 200), ar('QQ-77', 300, 0), ar('ZZ-900', 5, 0, 5, 'NJ')];
  bytes.set('ats.xlsx', atsBook(smallAts));
  bytes.set('ats-nocol.xlsx', atsBook(smallAts, { header: AH.map(h => (h === 'Total Customer Orders' ? 'Orders Total' : h)) }));

  // A small set for the selling view, dated around today so each pull-from rule gives a date of its own:
  //   SL-101 Booked, ETD in 30 days (10 days before the ETD) · SL-102 two lots in production, CRDs in 40 and 90 days
  //   (5 days before the CRD), customer orders beyond its stock take the first lot and part of the second · SL-103 no
  //   CRD and no Ship by (about 75 days after the order date), not on the ATS report · SL-104 Ready, CRD passed ("now"),
  //   oversold · SL-105 shipped, on the water (its ETA) · SL-106 Factory CRD and CRD Expected 20 days apart (dates
  //   disagree) · SL-107 China Pickup (sold) · SL-108 on the ATS report only · SL-109 Ready, CRD passed ("now"), part
  //   taken by orders.
  const sd = d => C.fmtDay(today + d), sser = d => today + d + 25569;
  const sellLines = [   // [style, PO, quantity, Master, order date, CRD Expected, Factory CRD, Ship by, order type, UPC] (days from today)
    ['SL-105', 900506, 2400, 24, -90, -40, -40, -40, 'Stock', '000000005055'],
    ['SL-101', 900501, 1200, 24, -60, 42, 40, 41, 'Stock', '000000005017'],
    ['SL-102', 900502, 1200, 12, -60, 40, 40, 40, 'Stock', '000000005024'],
    ['SL-102', 900503, 480, 12, -60, 90, 90, 90, 'Stock', '000000005024'],
    ['SL-103', 900504, 600, 12, -10, null, null, null, 'Stock', '000000005031'],
    ['SL-104', 900505, 1200, 24, -60, -20, -20, -20, 'Stock', '000000005048'],
    ['SL-106', 900507, 960, 24, -60, 60, 40, 40, 'Stock', '000000005062'],
    ['SL-107', 900508, 720, 24, -60, 30, 30, 30, 'China Pickup', '000000005079'],
    ['SL-109', 900509, 600, 12, -60, -12, -12, -12, 'Stock', '000000005093']];
  const sOpt = d => (d == null ? null : sd(d));
  bytes.set('trk-sell.xlsx', xlsxBytes({ 'PO Status Tracker': [TH, ...sellLines.map(([sty, po, , , od, exp, fcrd, , type]) =>
    [po, 'Factory S', sd(od), 'Item ' + sty, sty, 'Cosmetics', type, type === 'China Pickup' ? 'Customer 2' : 'Customer 1', sd(120), sOpt(exp), sOpt(fcrd), null])] }));
  const sm = ([sty, po, qty, master, od, , , sb, , upc]) => ['Lip', 'Gloss', 'Brand S', 'Item ' + sty, sty, upc, 6, master, 2, qty, null, String(po), 'Factory S', sd(od), null, sOpt(sb), null];
  bytes.set('msl-sell.xlsx', xlsxBytes({ 'Master Sheet Live': [MH, sm(sellLines[0]), m('Round 1', null, null, null), ...sellLines.slice(1).map(sm)] }));
  bytes.set('ct-sell.xlsx', xlsxBytes({
    'Booking vs Master PO Comp': [['PO', 'Style']],
    Booking: [BOOK, row(BOOK, { Category: 'Round 380' }),
      row(BOOK, { Style: 'SL-101', Barcode: '000000005017', Master: 24, Quantity: 1200, 'PO Number': 900501, Company: 'Factory S' })],
    Loaded: [LOAD, row(LOAD, { Category: 'Loaded' }), row(LOAD, { Category: 'Round 381' }),
      row(LOAD, { Style: 'SL-105', Barcode: '000000005055', Master: 24, Quantity: 2400, 'Po Number': 900506, Company: 'Factory S', 'Ready Date': sser(-38) })],
    Open: [CONT, row(CONT, { 'ID#': 1, 'Container #': 'GGGU1111111', 'Round #': 380, ETD: sser(30), ETA: sser(60), Factory: 'Factory S', 'PO#': '900501' }),
      row(CONT, { 'ID#': 2, 'Container #': 'FFFU5555555', 'Round #': 381, ETD: sser(-5), ETA: sser(20), Factory: 'Factory S', 'PO#': '900506' })],
    Complete: [CONT, row(CONT, { 'ID#': 'Complete Shipments' }), row(CONT, { 'ID#': 3, 'Container #': 'EEEU3333333', 'Round #': 300, ETD: ser(2026, 6, 5), ETA: ser(2026, 7, 5) })],
  }));
  // SL-101's "ATS from in WH" isn't In WH − Total Customer Orders (listed); SL-105's numbers are stored as text.
  const sellAts = [ar('SL-101', 1000, 200, 700), ar('SL-102', 100, 1600), ar('SL-104', 0, 5000), ar('SL-105', '300', '0', '300'),
    ar('SL-106', 0, 0), ar('SL-107', 0, 0), ar('SL-108', 50, 0), ar('SL-109', 0, 100)];
  bytes.set('ats-sell.xlsx', atsBook(sellAts));
  bytes.set('ats-sell-bad.xlsx', atsBook(sellAts, { grand: 9999 }));

  // The big set's report: every third style, some with more customer orders than stock (taken off their lots, or
  // oversold), and three styles on the report only.
  {
    const rows = [];
    for (let i = 0; i < 2400; i += 3) rows.push(ar('ST' + String(i).padStart(4, '0'), (i % 5) * 100, (i % 4) * 400));
    rows.push(ar('ERP-1', 500, 0), ar('ERP-2', 0, 240), ar('ERP-3', 1200, 1200));
    bytes.set('ats-big.xlsx', atsBook(rows));
  }

  // The PO Book page, served at /po-book/index.html: its CORE block as copied into core/pobook.js, or changed.
  const block = PB.coreText(readFileSync(join(REPO, 'core', 'pobook.js'), 'utf8'));
  const poBook = body => `<!doctype html><html><head><title>PO Book</title></head><body><script>\n/*CORE-START*/${body}/*CORE-END*/\n</script></body></html>`;
  return { bytes, pob, upcPages, poBookSame: poBook(block), poBookChanged: poBook(block + '\n/* a rule changed */'), blockMatches: block != null && sha(block) === PB.CORE_SHA256 };
}

// Local copies of the real files for s2-real, under names the stand-in serves (read from REAL_DIR, never printed).
function addRealFiles(fx, dir, listsPath) {
  const names = { 'real-msl.xlsx': 'Master Sheet Live.xlsx', 'real-trk.xlsx': 'PO Status Tracker.xlsx', 'real-ct.xlsx': 'Container Tracking Sheet.xlsx',
    'real-fwd.xlsx': 'Forwarder.xlsx', 'real-dray.csv': 'Drayage.csv' };
  for (const [k, f] of Object.entries(names)) fx.bytes.set(k, readFileSync(join(dir, f)));
  const L = JSON.parse(readFileSync(listsPath, 'utf8'));
  if (!Array.isArray(L.UPC_MASTER) || !Array.isArray(L.PO_Book_Lines)) throw new Error(`${listsPath} needs UPC_MASTER and PO_Book_Lines arrays`);
  const half = Math.ceil(L.UPC_MASTER.length / 2);   // two pages, so the page follows @odata.nextLink as on SharePoint
  fx.real = { pob: L.PO_Book_Lines, upcPages: [L.UPC_MASTER.slice(0, half), L.UPC_MASTER.slice(half)] };
}

// The stage engine on the same files, with the input app.js gives it: data.po (PO Book lines with the marks, and the
// raw Master Sheet Live rows), and each other source's parse result, or null where the page doesn't use it
// (fwd: null / dray: null for a report that failed or is too old).
function stagesFor(fx, today, f = {}) {
  f = { msl: 'msl.xlsx', trk: 'tracker.xlsx', ct: 'ct.xlsx', fwd: 'fwd.xlsx', dray: 'dray.csv', pob: fx.pob, upc: fx.upcPages.flat(), sections: TEST_SECTIONS, ...f };
  const sheet = (name, sh) => C.sheetRows(X, C.findSheet(X.read(fx.bytes.get(name), { ...XOPT, sheets: [sh] }), sh).ws);
  const trows = sheet(f.trk, 'PO Status Tracker'), mrows = sheet(f.msl, 'Master Sheet Live');
  const trk = PB.parseTracker(trows), src = PB.parseMaster(mrows, trk), ov = PB.overlayFromItems(f.pob), cfg = PB.cfgFrom(f.pob);
  const po = { trk, src, lines: PB.applyOverlay(src, ov, cfg), ov, cfg, counts: PB.countMaster(mrows, trk).counts, mrows };
  const ok = r => (r && r.ok ? r : null);
  const ct = CT.parseContainerTracking(X, X.read(fx.bytes.get(f.ct), XOPT), today);
  const fwd = f.fwd ? SH.parseForwarder(X, X.read(fx.bytes.get(f.fwd), XOPT), today) : null;
  const dray = f.dray ? SH.parseDrayage(new TextDecoder().decode(fx.bytes.get(f.dray)), today) : null;
  const upc = f.upc ? UPC.fromListItems(f.upc) : null;
  return ST.build({ po, mrows, ct: ok(ct), forwarder: ok(fwd), drayage: ok(dray), upc: ok(upc), today, sections: f.sections });
}
// Available to sell on the same files, as app.js builds it: the stage engine, the ATS report read with core/ats.js under
// the name the stand-in gives it (f.ats: the world's ats option; its date is in the name), and UPC_MASTER. A report
// that fails is passed as null, as on the page. { R, A (parseAts's result, or null), S (buildSell's, or null) }.
function sellFor(fx, today, f = {}) {
  const R = stagesFor(fx, today, f);
  const A = f.ats && !f.ats.empty ? ATS.parseAts(X, X.read(fx.bytes.get(f.ats.file || 'ats.xlsx'), XOPT), atsName(f.ats), today) : null;
  const upc = UPC.fromListItems(f.upc || fx.upcPages.flat());
  const S = R.ok ? SELL.buildSell({ stages: R, ats: A && A.ok ? A : null, today, upc: upc.ok ? upc : null }) : null;
  return { R, A, S };
}
// Saul's rule, worked out here on its own from each lot's dates (not with core/sell.js's code): still in China, 10 days
// before the ETD; no ETD, 5 days before the ready date (the first of Container Tracking's Ready Date, the PO Book's own
// date, Factory CRD, CRD Expected, Ship by); no ready date, 75 days after the order date (an estimate, its range from the
// order date). Sailed: the ETA. A date before today is "now"; ready dates more than 7 days apart disagree (in China).
// Also each lot's units = taken + available. Answers one line per lot that breaks a rule.
const READY_KEYS = ['ctReady', 'poBook', 'factoryCrd', 'crdExpected', 'shipBy'];
const isDay = v => typeof v === 'number' && Number.isFinite(v);
// A source more than a year ahead or before 2015 (a wrong year) isn't used, unless every source is.
const readyPick = (src, today) => src.find(x => !(x > today + 366 || x < C.ymd(2015, 1, 1))) ?? (src.length ? src[0] : null);
function ruleMismatches(S, today) {
  const bad = [];
  for (const l of S.lots) {
    const d = l.dates || {}, src = l.dates ? READY_KEYS.map(k => d[k]).filter(isDay) : isDay(l.crd) ? [l.crd] : [];
    const ready = readyPick(src, today), why = [];
    if ((l.readyDate ?? null) !== ready) why.push(`ready date ${l.readyDate}, not ${ready}`);
    if (l.inChina) {
      const w = isDay(l.etd) ? [l.etd - 10, 'etd'] : ready != null ? [ready - 5, 'crd'] : isDay(d.orderDate) ? [d.orderDate + 75, 'order'] : [null, 'none'];
      if ((l.pullFrom ?? null) !== w[0] || l.pullRule !== w[1]) why.push(`pull-from ${l.pullFrom} (${l.pullRule}), not ${w[0]} (${w[1]})`);
      if (!!l.estimate !== (w[1] === 'order')) why.push('estimate');
      if (w[1] === 'order' && !(l.pullRange && l.pullRange.from === d.orderDate && l.pullRange.to === d.orderDate + 75)) why.push('the range is not the order date to 75 days after it');
      if ((l.availableDate ?? null) !== w[0]) why.push('available date is not the pull-from date');
      if (l.flags.includes('now') !== (w[0] != null && w[0] < today)) why.push('"now"');
      const spread = src.length > 1 ? Math.max(...src) - Math.min(...src) : 0;
      if (l.flags.includes('dates-disagree') !== spread > 7) why.push(`dates-disagree with a spread of ${spread} days`);
    } else if ((l.availableDate ?? null) !== (isDay(l.eta) ? l.eta : null) || l.flags.includes('dates-disagree')) why.push('sailed: not its ETA');
    if (l.units !== l.taken + l.available || l.taken < 0 || l.available < 0) why.push(`units ${l.units} ≠ taken ${l.taken} + available ${l.available}`);
    if (why.length) bad.push(`${l.style} PO ${l.po}: ${why.join('; ')}`);
  }
  return bad;
}
// The selling engine's sums: the styles add up to the totals, each style's lots to its numbers, stock after orders,
// customer orders taken off the earliest lots first, the lots to the stage engine's Incoming and sold to its Pull from
// China. Answers what doesn't hold.
function sellSums(S, R) {
  const bad = [], T = S.totals, sum = (a, f) => a.reduce((p, x) => p + (f(x) || 0), 0);
  for (const k of ['now', 'fromChina', 'arriving', 'total', 'soldChina', 'oversold', 'ordersTaken']) if (sum(S.styles, x => x[k]) !== T[k]) bad.push(`styles' ${k} ${sum(S.styles, x => x[k])} ≠ total ${T[k]}`);
  if (T.total !== T.now + T.fromChina + T.arriving) bad.push('total ≠ now + from China + arriving');
  const lotDay = l => (l.availableDate == null ? Infinity : l.availableDate);
  for (const s of S.styles) {
    const ls = s.lots, lu = sum(ls, l => l.units);
    if (lu !== s.fromChina + s.arriving + s.ordersTaken) bad.push(`${s.style}: lots ${lu} ≠ from China + arriving + taken`);
    if (s.fromChina !== sum(ls.filter(l => l.inChina), l => l.available) || s.arriving !== sum(ls.filter(l => !l.inChina), l => l.available)) bad.push(`${s.style}: lot availability`);
    if (sum(ls, l => l.taken) !== s.ordersTaken) bad.push(`${s.style}: taken ≠ orders taken`);
    const known = isDay(s.inWh) && isDay(s.orders) && !s.flags.includes('stock-unknown');
    if (known && (s.now !== Math.max(0, s.inWh - s.orders) || s.ordersTaken + s.oversold !== Math.max(0, s.orders - s.inWh))) bad.push(`${s.style}: stock after orders`);
    if (!known && (s.ordersTaken || s.oversold || s.now)) bad.push(`${s.style}: orders taken off with stock not known`);
    // Taken from the earliest first (no date last): once a lot keeps something, no later lot gives any.
    let kept = false;
    for (const l of ls.slice().sort((a, b) => (lotDay(a) > lotDay(b)) - (lotDay(a) < lotDay(b)))) {
      if (kept && l.taken) { bad.push(`${s.style}: a later lot taken before an earlier one`); break; }
      if (l.available > 0) kept = true;
    }
  }
  if (R && R.ok) {
    if (sum(S.lots, l => l.units) !== R.totals.units.incoming) bad.push(`lots ${sum(S.lots, l => l.units)} ≠ Incoming ${R.totals.units.incoming}`);
    if (T.soldChina !== R.totals.units.pullFromChina) bad.push(`sold ${T.soldChina} ≠ Pull from China ${R.totals.units.pullFromChina}`);
  }
  return bad;
}
// The section lists in the real config.js (public), run with a stand-in window: s2-real serves them, so the page and
// the engine here read the real files as the live page does.
function realSections() {
  const win = {};
  new Function('window', readFileSync(join(REPO, 'config.js'), 'utf8'))(win);
  return win.INV_CONFIG ? win.INV_CONFIG.sections : undefined;
}

// Every issue code and flag core/stages.js, core/ats.js and core/sell.js can give, without plain words in app.js's
// CODE_LABELS (the page would show the bare code). Read from the files' text.
function unlabelledCodes() {
  const app = readFileSync(join(REPO, 'app.js'), 'utf8');
  const m = /const CODE_LABELS = (\{[^]*?\n {2}\});/.exec(app);
  if (!m) return ['(no CODE_LABELS object in app.js)'];
  const labels = new Function(`return ${m[1]};`)();
  const codes = new Set();
  for (const f of ['stages', 'ats', 'sell']) {
    const eng = readFileSync(join(REPO, 'core', f + '.js'), 'utf8');
    for (const x of eng.matchAll(/\bissue\('([\w-]+)'/g)) codes.add(x[1]);
    for (const x of eng.matchAll(/'([a-z]+(?:-[a-z]+)+)'/g)) codes.add(x[1]);   // the hyphenated codes and flags
    // flags pushed onto a lot or style, one-word ones too ("now", "taken", "oversold")
    for (const x of eng.matchAll(/flags\.push\(([^)]*)\)/g)) for (const y of x[1].matchAll(/'([\w-]+)'/g)) codes.add(y[1]);
  }
  return [...codes].filter(k => !labels[k]).sort();
}

// What the page should show, from the same parsers run here on the same files, the way app.js calls them.
function expected(fx, today) {
  const sheet = (name, sh) => C.sheetRows(X, C.findSheet(X.read(fx.bytes.get(name), { ...XOPT, sheets: [sh] }), sh).ws);
  const trows = sheet('tracker.xlsx', 'PO Status Tracker'), mrows = sheet('msl.xlsx', 'Master Sheet Live');
  const tc = PB.countTracker(trows), trk = PB.parseTracker(trows), cm = PB.countMaster(mrows, trk), src = PB.parseMaster(mrows, trk);
  const ov = PB.overlayFromItems(fx.pob), lines = PB.applyOverlay(src, ov, PB.cfgFrom(fx.pob));
  const plain = PB.applyOverlay(src, new Map(), PB.cfgFrom([]));
  const open = ls => ls.filter(l => l.status === 'open').length;
  const ct = CT.parseContainerTracking(X, X.read(fx.bytes.get('ct.xlsx'), XOPT), today);
  const fwd = SH.parseForwarder(X, X.read(fx.bytes.get('fwd.xlsx'), XOPT), today);
  const dray = SH.parseDrayage(new TextDecoder().decode(fx.bytes.get('dray.csv')), today);
  const upc = UPC.fromListItems(fx.upcPages.flat());
  const sum = a => a.reduce((s, v) => s + (v || 0), 0);
  const e = {
    tracker: { read: tc.rowsRead, counted: tc.counts.tracked, issues: tc.issues.length + src.issues.filter(i => i.code === 'trkonly').length },
    msl: { read: cm.rowsRead, counted: cm.counts.line, issues: cm.issues.length + src.issues.filter(i => i.code !== 'nostyle' && i.code !== 'trkonly').length,
      lines: lines.length, open: open(lines), shipped: lines.length - open(lines), openPlain: open(plain) },
    // Items with no LineKey are listed, one issue each.
    pobook: { read: fx.pob.length, counted: ov.size, issues: fx.pob.filter(i => !i.fields.LineKey).length },
    // Rows read on all four tabs; rows used = style lines on Booking and Loaded + container rows on Open and Complete.
    ct: { read: sum(['Booking', 'Loaded', 'Open', 'Complete'].map(k => ct.counts[k].rowsRead)),
      counted: ct.counts.Booking.counted + ct.counts.Loaded.counted + ct.counts.Open.roundRows + ct.counts.Complete.roundRows, issues: ct.issues.length, ok: ct.ok },
    forwarder: { read: fwd.counts.rowsRead, counted: fwd.counts.records, issues: fwd.issues.length, ok: fwd.ok },
    drayage: { read: dray.counts.rowsRead, counted: dray.counts.records, issues: dray.issues.length, ok: dray.ok },
    upc: { read: upc.counts.rows, counted: upc.counts.used, issues: upc.issues.length, badupc: upc.issues.filter(i => i.code === 'badupc').length },
  };
  // The stage engine's own issues join the list on Problems, under "Incoming".
  const R = stagesFor(fx, today);
  e.incoming = { issues: R.ok ? R.issues.length : 0 };
  // The ATS folder is empty in the standard world: Available to sell without the report (its issues under "Available
  // to sell"; none from the report).
  const S = R.ok ? SELL.buildSell({ stages: R, ats: null, today, upc }) : null;
  e.ats = { issues: 0 };
  e.sell = { issues: S && S.ok ? S.issues.length : 0 };
  e.totalIssues = sum(Object.values(e).map(x => x.issues));
  e.sane = tc.ok && cm.ok && ct.ok && fwd.ok && dray.ok && upc.ok && e.msl.open !== e.msl.openPlain && e.totalIssues > 0 &&
    R.ok && R.balance.diff === 0 && e.incoming.issues > 0 && !!S && S.ok && S.atsMissing && e.sell.issues > 0;
  Object.defineProperty(e, 'R', { value: R, enumerable: false });   // kept out of the printed summary
  Object.defineProperty(e, 'S', { value: S, enumerable: false });
  return e;
}

// ------------------------------------------------------------------ the stand-in for SharePoint (Graph and downloads)

let DL = '';          // the download host: another port on 127.0.0.1, so downloads are cross-origin like SharePoint's
let world = null;     // the SharePoint state of the scenario being run

function newWorld(fx, o = {}) {
  return { fx, o: { mslFile: 'msl.xlsx', trkFile: 'tracker.xlsx', ctFile: 'ct.xlsx', fwdFile: 'fwd.xlsx', drayFile: 'dray.csv', drayOldFile: 'dray.csv',
    fwdAge: 0, tracker: 200, poBook: 'same', mail: 202, ats: { empty: true }, slowCt: false, sources: {}, real: false, ...o },
    calls: [], downloads: [], pageGets: [], unanswered: [], pending: new Set(), aborted: [], poBookGets: 0 };
}
// The ATS folder's reports, named as the ERP names them, with the day the report was made: "Inventory Report - On
// Hand_In Transit_In Production 10-28-2026 12-58 PM.xlsx" (o.age days ago; `older` days before that for the older one),
// or with no date in the name (o.noDate), when the page takes the day the file was saved.
function atsName(o, older = 0) {
  if (o && o.noDate) return older ? 'ats-report-older.xlsx' : 'ats-report-new.xlsx';
  const dt = new Date((TODAY - ((o && o.age) || 0) - older) * 864e5);
  return `Inventory Report - On Hand_In Transit_In Production ${dt.getUTCMonth() + 1}-${dt.getUTCDate()}-${dt.getUTCFullYear()} 12-58 PM.xlsx`;
}
// The config.js the page gets: the invented locations, with the scenario's changes.
const configFor = w => ({ ...TEST_CONFIG, sources: { ...LOC, ...((w && w.o.sources) || {}) }, ...(w && w.o.real ? { sections: realSections() } : {}) });

const J = (status, json) => ({ status, json });
const notFound = () => J(404, { error: { code: 'itemNotFound', message: 'The resource could not be found.' } });
const denied = () => J(403, { error: { code: 'accessDenied', message: 'Access denied.' } });
const webUrl = id => `https://example.invalid/files/${id}`;
const by = { user: { displayName: 'Test Saver' } };
function fileMeta(id, name, fx, age) {
  return { id, name, webUrl: webUrl(id), lastModifiedDateTime: iso(age), lastModifiedBy: by, file: {}, eTag: `"${id}-1"`,
    parentReference: { driveId: SRC }, '@microsoft.graph.downloadUrl': `${DL}/fx/${fx}` };
}
const folderMeta = (id, name) => ({ id, name, webUrl: webUrl(id), folder: { childCount: 3 }, lastModifiedDateTime: iso(1) });
const listMeta = (id, name) => ({ id, displayName: name, name, webUrl: webUrl(id), lastModifiedDateTime: iso(1), lastModifiedBy: by });
const kid = (id, name, age, extra) => ({ id, name, lastModifiedDateTime: iso(age), lastModifiedBy: by, webUrl: webUrl(id), file: {}, ...extra });

// /sites/<site> paths go to the host app.js gives them; the sites, their libraries and two sharing links.
const SITES = { Production: SITE_ID, TestSite: 'site-test', TestReports: 'site-rep' };
const DRIVES = {
  'site-test': [{ id: 'drv-assets', name: 'Site Assets', webUrl: `${TEST_WEB}/sites/TestSite/SiteAssets` },
    { id: SRC, name: 'Documents', webUrl: `${TEST_WEB}/sites/TestSite/Shared%20Documents` }],
  'site-rep': [{ id: REP, name: 'Documents', webUrl: `${TEST_WEB}/sites/TestReports/Shared%20Documents` }],
};
const SHARES = { [LOC.msl]: ['itm-msl', 'Master Sheet Live.xlsx'], [LOC.tracker]: ['itm-trk', 'PO Status Tracker.xlsx'] };
const isResolve = x => /^\/shares\/|^\/sites\/[^/]+:\/sites\/(?!Production$)|\/drives$|\/root:\//.test(x.path);
// Finding a location afresh: a sharing link, a site or its libraries (a path inside a known library is not one).
const isLookup = x => /^\/shares\/|^\/sites\/[^/]+:\/sites\/(?!Production$)|\/drives$/.test(x.path);

function graph(w, method, path, q) {
  const o = w.o;
  if (method === 'GET') {
    let m;
    if ((m = /^\/sites\/met1sol\.sharepoint\.com:\/sites\/([^/]+)$/.exec(path))) return SITES[m[1]] ? J(200, { id: SITES[m[1]] }) : notFound();
    if ((m = /^\/sites\/([^/]+)\/drives$/.exec(path)) && DRIVES[m[1]]) return J(200, { value: DRIVES[m[1]] });
    if ((m = /^\/shares\/u!([\w-]+)\/driveItem$/.exec(path))) {
      const hit = SHARES[Buffer.from(m[1], 'base64url').toString('utf8')];
      if (!hit) return notFound();
      // o.shares 403: Graph won't open sharing links for this sign-in (Microsoft lists only write permissions for it).
      if (o.shares === 403) return denied();
      return J(200, { id: hit[0], name: hit[1], webUrl: webUrl(hit[0]), file: {}, parentReference: { driveId: SRC } });
    }
    // o.ctReplaced: the sheet was renamed to "… OLD.xlsx" and a new file saved under the old name (a new id).
    if (path === `/drives/${SRC}/root:/Logistics/Test Container Sheet.xlsx`) return J(200, o.ctReplaced ? fileMeta('itm-ct-new', 'Test Container Sheet.xlsx', o.ctFile, 0) : fileMeta('itm-ct', 'Test Container Sheet.xlsx', o.ctFile, 0));
    if (path === `/drives/${SRC}/items/itm-ct-new` && o.ctReplaced) return J(200, fileMeta('itm-ct-new', 'Test Container Sheet.xlsx', o.ctFile, 0));
    if (path === `/drives/${SRC}/root:/Logistics/Test Container Sheet v2.xlsx`) return J(200, fileMeta('itm-ct2', 'Test Container Sheet v2.xlsx', 'ct.xlsx', 0));
    if (path === `/drives/${REP}/root:/Shipping/Forwarder`) return J(200, folderMeta('itm-fwd-dir', 'Forwarder'));
    if (path === `/drives/${REP}/root:/Shipping/Drayage`) return J(200, folderMeta('itm-dray-dir', 'Drayage'));
    if (path === `/drives/${SRC}/root:/Stock/ATS` || path === `/drives/${SRC}/items/itm-ats-dir`) return o.ats ? J(200, folderMeta('itm-ats-dir', 'ATS')) : notFound();
    if (path === `/drives/${SRC}/items/itm-ats-dir/children` && o.ats) {
      const lock = kid('itm-ats-lock', '~$ats-report-new.xlsx', 0), sub = { id: 'itm-ats-sub', name: 'Old', folder: { childCount: 2 }, lastModifiedDateTime: iso(0) };
      if (o.ats.empty) return J(200, { value: [lock, sub] });
      // A newer file that isn't Excel is skipped: the newest Excel file is the report.
      // o.ats.resaved: the older report was opened and saved again after the newest one (it must not win).
      return J(200, { value: [kid('itm-ats-txt', 'read-me.txt', 0), kid('itm-ats-1', atsName(o.ats), Math.max(0, o.ats.age)),
        kid('itm-ats-0', atsName(o.ats, 7), o.ats.resaved ? 0 : Math.max(0, o.ats.age) + 7), lock, sub] });
    }
    // The newest report, with its download link (the invented workbook o.ats.file); the older one only if asked for.
    if (path === `/drives/${SRC}/items/itm-ats-1` && o.ats && !o.ats.empty) return J(200, fileMeta('itm-ats-1', atsName(o.ats), o.ats.file || 'ats.xlsx', Math.max(0, o.ats.age)));
    if (path === `/drives/${SRC}/items/itm-ats-0` && o.ats && !o.ats.empty) return J(200, fileMeta('itm-ats-0', atsName(o.ats, 7), 'ats.xlsx', o.ats.resaved ? 0 : Math.max(0, o.ats.age) + 7));
    if (path === `/drives/${SRC}/items/itm-msl`) return J(200, fileMeta('itm-msl', 'Master Sheet Live.xlsx', o.mslFile, 1));
    if (path === `/drives/${SRC}/items/itm-trk`) return o.tracker === 403 ? denied() : J(200, fileMeta('itm-trk', 'PO Status Tracker.xlsx', o.trkFile, 1));
    if (path === `/drives/${SRC}/items/itm-ct`) return J(200, fileMeta('itm-ct', o.ctReplaced ? 'Test Container Sheet OLD.xlsx' : 'Test Container Sheet.xlsx', o.ctFile, o.ctReplaced ? 5 : 0));
    if (path === `/drives/${SRC}/items/itm-ct2`) return J(200, fileMeta('itm-ct2', 'Test Container Sheet v2.xlsx', 'ct.xlsx', 0));
    if (path === `/drives/${REP}/items/itm-fwd-dir`) return J(200, folderMeta('itm-fwd-dir', 'Forwarder'));
    if (path === `/drives/${REP}/items/itm-fwd-dir/children`) {
      // two pages; the newest file is on the second, so the page must follow @odata.nextLink
      if (q.get('$skiptoken') === 'p2') return J(200, { value: [kid('itm-fwd-new', 'forwarder-report-new.xlsx', o.fwdAge)] });
      return J(200, { value: [kid('itm-fwd-old', 'forwarder-report-old.xlsx', o.fwdAge + 3), kid('itm-fwd-lock', '~$forwarder-report-new.xlsx', 0),
        { id: 'itm-fwd-sub', name: 'Archive', folder: { childCount: 9 }, lastModifiedDateTime: iso(0) }, kid('itm-fwd-txt', 'read-me.txt', 0)],
      '@odata.nextLink': `${GRAPH}/drives/${REP}/items/itm-fwd-dir/children?$top=999&$skiptoken=p2` });
    }
    if (path === `/drives/${REP}/items/itm-fwd-new`) return J(200, fileMeta('itm-fwd-new', 'forwarder-report-new.xlsx', o.fwdFile, o.fwdAge));
    if (path === `/drives/${REP}/items/itm-fwd-old`) return J(200, fileMeta('itm-fwd-old', 'forwarder-report-old.xlsx', 'fwd.xlsx', o.fwdAge + 3));
    if (path === `/drives/${REP}/items/itm-dray-dir`) return J(200, folderMeta('itm-dray-dir', 'Drayage'));
    if (path === `/drives/${REP}/items/itm-dray-dir/children`) {
      return J(200, { value: [kid('itm-dray-old', 'drayage-report-old.csv', 1), kid('itm-dray-new', 'drayage-report-new.csv', 0), kid('itm-dray-x', 'drayage-report-new.xlsx', 0)] });
    }
    if (path === `/drives/${REP}/items/itm-dray-new`) return J(200, fileMeta('itm-dray-new', 'drayage-report-new.csv', o.drayFile, 0));
    // asked for only when the newest one can't be read
    if (path === `/drives/${REP}/items/itm-dray-old`) return J(200, fileMeta('itm-dray-old', 'drayage-report-old.csv', o.drayOldFile, 1));
    if (path === `/sites/${SITE_ID}/lists/PO_Book_Lines`) return J(200, listMeta('lst-pob', 'PO_Book_Lines'));
    if (path === `/sites/${SITE_ID}/lists/lst-pob/items`) return J(200, { value: o.real ? w.fx.real.pob : w.fx.pob });
    if (path === `/sites/${SITE_ID}/lists/UPC_MASTER`) return J(200, listMeta('lst-upc', 'UPC_MASTER'));
    if (path === `/sites/${SITE_ID}/lists/lst-upc/items`) {
      const pages = o.real ? w.fx.real.upcPages : w.fx.upcPages;
      if (q.get('$skiptoken') === 'p2') return J(200, { value: pages[1] });
      return J(200, { value: pages[0], '@odata.nextLink': `${GRAPH}/sites/${SITE_ID}/lists/lst-upc/items?$expand=fields&$top=999&$skiptoken=p2` });
    }
  }
  if (method === 'POST' && path === '/me/sendMail') {
    return o.mail === 202 ? { status: 202 } : J(o.mail, { error: { code: 'ErrorAccessDenied', message: 'Access is denied. Check credentials and try again.' } });
  }
  w.unanswered.push(`${method} ${path}`);
  return method === 'GET' ? notFound() : J(405, { error: { code: 'notAllowed', message: 'The stand-in answers reads and the email only.' } });
}

const TYPES = { '.html': 'text/html; charset=utf-8', '.js': 'text/javascript; charset=utf-8', '.css': 'text/css; charset=utf-8', '.png': 'image/png', '.json': 'application/json', '.svg': 'image/svg+xml' };
const SERVED = /^(index\.html|app\.js|app\.css|favicon\.png|core\/[\w.-]+\.js|vendor\/[\w./-]+)$/;

function pageServer() {
  return createServer((req, res) => {
    const [rawPath, rawQ = ''] = req.url.split('?');
    let path;
    try { path = decodeURIComponent(rawPath); } catch (e) { res.writeHead(400); res.end(); return; }
    if (path.startsWith('/__graph/')) {
      const chunks = [];
      req.on('data', c => chunks.push(c));
      req.on('end', () => {
        const text = Buffer.concat(chunks).toString('utf8');
        let body = null;
        try { body = text ? JSON.parse(text) : null; } catch (e) { body = text; }
        const gpath = path.slice('/__graph'.length), q = new URLSearchParams(rawQ);
        const w = world;
        const r = w ? graph(w, req.method, gpath, q) : notFound();
        if (w) w.calls.push({ method: req.method, path: gpath, query: rawQ, auth: req.headers.authorization || '', prefer: req.headers.prefer || '', body, status: r.status });
        res.writeHead(r.status, r.json ? { 'Content-Type': 'application/json' } : {});
        res.end(r.json ? JSON.stringify(r.json) : undefined);
      });
      return;
    }
    if (world) world.pageGets.push({ method: req.method, path });
    if (req.method !== 'GET') { res.writeHead(405); res.end(); return; }
    if (path === '/config.js') {
      res.writeHead(200, { 'Content-Type': TYPES['.js'], 'Cache-Control': 'no-store' });
      res.end(`window.INV_CONFIG = ${JSON.stringify(configFor(world), null, 2)};\n`);
      return;
    }
    if (path === '/po-book/index.html') {
      if (world) world.poBookGets++;
      const v = world && world.o.poBook;
      if (v !== 'same' && v !== 'changed') { res.writeHead(404); res.end('Not found'); return; }
      res.writeHead(200, { 'Content-Type': TYPES['.html'] });
      res.end(v === 'same' ? world.fx.poBookSame : world.fx.poBookChanged);
      return;
    }
    const rel = path === '/' ? 'index.html' : path.replace(/^\/+/, '');
    const file = resolve(REPO, rel);
    if (!SERVED.test(rel) || !file.startsWith(REPO + sep) || !existsSync(file) || !statSync(file).isFile()) { res.writeHead(404); res.end('Not found'); return; }
    res.writeHead(200, { 'Content-Type': TYPES[extname(file)] || 'application/octet-stream', 'Cache-Control': 'no-store' });
    res.end(readFileSync(file));
  });
}

// Pre-signed download links: no sign-in needed, cross-origin with CORS open, and no Authorization header expected.
function downloadServer() {
  return createServer((req, res) => {
    const cors = { 'Access-Control-Allow-Origin': '*' };
    const name = decodeURIComponent(req.url.split('?')[0]).replace(/^\/fx\//, '');
    const w = world;
    if (w) w.downloads.push({ name, method: req.method, auth: !!req.headers.authorization });
    if (req.method === 'OPTIONS') { res.writeHead(204, cors); res.end(); return; }
    const buf = w && w.fx.bytes.get(name);
    if (!buf || req.method !== 'GET') { res.writeHead(404, cors); res.end(); return; }
    const send = () => { res.writeHead(200, { ...cors, 'Content-Type': 'application/octet-stream', 'Content-Length': buf.length }); res.end(buf); };
    if (w.o.slowCt && name === w.o.ctFile) {
      w.pending.add(name);
      const timer = setTimeout(() => { w.pending.delete(name); send(); }, 25000);
      res.on('close', () => { if (w.pending.delete(name)) { clearTimeout(timer); w.aborted.push(name); } });
      return;
    }
    send();
  });
}

// ------------------------------------------------------------------ what runs in the page before its own scripts

// opts.scp: the permissions the fake sign-in's tokens carry (the page reads them from the token, as from a real one);
// MSAL's result lists them too, as the sign-in service's reply does. opts.opaque: tokens the page can't decode.
// opts.noStorage: browser storage blocked, as in some locked-down browsers.
function initScript(pageOrigin, opts = {}) {
  return `(() => {
  if (window.__headless) return;
  const H = window.__headless = { blocked: [], requests: [], msal: [], msalSetTried: 0, config: null, tokens: 0, forced: 0, scp: ${JSON.stringify(opts.scp || READ_SCP)}, opaque: ${!!opts.opaque} };
  const GRAPH = ${JSON.stringify(GRAPH)}, PAGE = ${JSON.stringify(pageOrigin)}, DL = ${JSON.stringify(DL)};
  const realFetch = window.fetch.bind(window);
  // Graph goes to the local stand-in; the page and the download host pass; anything else is refused and recorded.
  // Every call is logged with its method and the address the page asked for.
  window.fetch = function (input, init) {
    const url = typeof input === 'string' ? input : (input && input.url) || String(input);
    const method = String((init && init.method) || (input && typeof input === 'object' && input.method) || 'GET').toUpperCase();
    H.requests.push({ method, url });
    if (url === GRAPH || url.startsWith(GRAPH + '/')) return realFetch(PAGE + '/__graph' + url.slice(GRAPH.length), init);
    let o = '';
    try { o = new URL(url, location.href).origin; } catch (e) { /* not a URL */ }
    if (o === PAGE || o === DL) return realFetch(input, init);
    H.blocked.push(url);
    return Promise.reject(new TypeError('Failed to fetch'));
  };
  // The 5-minute refresh is held here (any timer of 2 minutes or more) so a check can run it now with fireLong().
  const realSet = window.setTimeout.bind(window), realClear = window.clearTimeout.bind(window);
  const long = new Map();
  let seq = 1e9;
  window.setTimeout = function (fn, ms, ...args) {
    if (typeof fn === 'function' && Number(ms) >= 120000) { const id = ++seq; long.set(id, { fn, args }); return id; }
    return realSet(fn, ms, ...args);
  };
  window.clearTimeout = function (id) { if (!long.delete(id)) realClear(id); };
  H.longWaiting = () => long.size;
  H.fireLong = () => { const due = [...long.values()]; long.clear(); for (const t of due) t.fn(...t.args); return due.length; };
  ${opts.noStorage ? `Object.defineProperty(window, 'localStorage', { configurable: true, get() { throw new DOMException('Storage is blocked.', 'SecurityError'); } });` : ''}
  const b64u = o => btoa(JSON.stringify(o)).replace(/=+$/, '').replace(/\\+/g, '-').replace(/\\//g, '_');
  const account = ${JSON.stringify(ACCOUNT)};
  class InteractionRequiredAuthError extends Error { constructor(m) { super(m); this.name = 'InteractionRequiredAuthError'; } }
  const never = name => () => { H.msal.push(name); return Promise.reject(new Error(name + ' is not allowed in the headless check')); };
  class PublicClientApplication {
    constructor(cfg) { H.config = JSON.parse(JSON.stringify(cfg)); this.active = null; }
    async initialize() {}
    async handleRedirectPromise() { return null; }
    getAllAccounts() { return [account]; }
    getActiveAccount() { return this.active; }
    setActiveAccount(a) { this.active = a; }
    async acquireTokenSilent(req) {
      H.tokens++; if (req && req.forceRefresh) H.forced++;
      // A scenario sets H.tokenFail to make the quiet renewal fail the way MSAL's hidden-frame timeout does.
      if (H.tokenFail) { const e = new Error('Token renewal operation failed due to timeout.'); e.errorCode = H.tokenFail; throw e; }
      const t = H.opaque ? 'headless-opaque-token-' + H.tokens : b64u({ alg: 'none', typ: 'JWT' }) + '.' + b64u({ aud: 'https://graph.microsoft.com', scp: H.scp, name: account.name }) + '.headless';
      return { accessToken: t, account, scopes: H.scp.split(' ').filter(Boolean) };
    }
  }
  for (const n of ['loginRedirect', 'loginPopup', 'acquireTokenRedirect', 'acquireTokenPopup', 'logoutRedirect', 'logoutPopup']) PublicClientApplication.prototype[n] = never(n);
  const fake = Object.freeze({ PublicClientApplication, InteractionRequiredAuthError, headlessFake: true });
  // The vendored msal-browser assigns window.msal when it loads: the setter ignores it, so the fake stays.
  Object.defineProperty(window, 'msal', { configurable: false, enumerable: true, get: () => fake, set: () => { H.msalSetTried++; } });
})();`;
}

// Everything the checks read from the page, in one call.
function snapshot() {
  const q = s => document.querySelector(s);
  const txt = el => (el ? el.textContent.replace(/\s+/g, ' ').trim() : '');
  const shown = el => !!el && !el.hidden && !el.closest('[hidden]') && getComputedStyle(el).display !== 'none';
  const chips = {}, cards = {}, srows = {};
  for (const k of ['msl', 'tracker', 'pobook', 'ct', 'forwarder', 'drayage', 'ats', 'upc']) {
    const el = q('#chip-' + k);
    chips[k] = el ? { state: el.dataset.state, text: txt(el), title: el.getAttribute('title') || '', href: el.getAttribute('href') || '' } : null;
    const cd = q('#card-' + k);
    cards[k] = cd ? { text: txt(cd), pairs: [...cd.querySelectorAll('dl.cnt dt:not(.h)')].map(dt => [txt(dt), txt(dt.nextElementSibling)]) } : null;
    srows[k] = txt(q('#srow-' + k));
  }
  const issRows = [...document.querySelectorAll('#issTable tbody tr')];
  const H = window.__headless || {};
  let store = null;
  try { store = Object.fromEntries(Object.keys(localStorage).filter(k => k.startsWith('inv:')).map(k => [k, localStorage.getItem(k)])); } catch (e) { store = null; }
  return {
    title: document.title, appShown: shown(q('#app')), signinShown: shown(q('#signin')), busy: q('#app') && q('#app').getAttribute('aria-busy'),
    busyShown: shown(q('#busy')), health: txt(q('#health')), chips, cards, srows, cardCount: document.querySelectorAll('#srcCards .scard').length,
    banners: [...document.querySelectorAll('#banners .banner')].map(b => ({ kind: (b.className.match(/\b(err|warn|info)\b/) || [])[1] || '', text: txt(b), id: b.id || '',
      buttons: [...b.querySelectorAll('button')].map(txt) })),
    cProb: txt(q('#cProb')), pqInfo: txt(q('#pqInfo')),
    issueGroups: issRows.filter(r => r.classList.contains('gh')).map(txt),
    issueRows: issRows.filter(r => !r.classList.contains('gh') && !r.querySelector('.more,.empty')).length,
    srcRows: document.querySelectorAll('#srcTable tr[id^="srow-"]').length,
    tabs: Object.fromEntries(['sell', 'inv', 'pull', 'problems'].map(t => [t, shown(q('#tab-' + t))])),
    tabOrder: [...document.querySelectorAll('#tabs .tab')].map(b => b.dataset.tab), srcBoxIn: (q('#srcBox') && q('#srcBox').closest('.tabpane') || {}).id || '',
    who: txt(q('#who')), sourcesButton: !!q('#btnSources'), dialog: !!q('dialog'),
    toast: txt(q('#toast')), toastShown: shown(q('#toast')),
    msalFake: !!(window.msal && window.msal.headlessFake), msalSetTried: H.msalSetTried || 0, msalInteractive: H.msal || [],
    msalConfig: H.config, fetchBlocked: H.blocked || [], requests: H.requests || [], tokens: H.tokens || 0, forced: H.forced || 0,
    loading: [...document.querySelectorAll('.chip.src[data-state="load"]')].map(x => x.id), store,
  };
}
const READY = `(() => { if (window.__stale) return false; const a = document.getElementById('app'), h = document.getElementById('health');
  return !!a && !a.hidden && a.getAttribute('aria-busy') !== 'true' && !document.querySelector('.chip.src[data-state="load"]') && !!(h && h.textContent.trim()); })()`;
const IDLE = `(() => { const a = document.getElementById('app'), b = document.getElementById('busy');
  return !!a && !a.hidden && a.getAttribute('aria-busy') !== 'true' && !!b && b.hidden && !document.querySelector('.chip.src[data-state="load"]'); })()`;

// ------------------------------------------------------------------ Edge over the DevTools protocol

class Cdp {
  constructor(url) {
    this.seq = 0; this.waiting = new Map(); this.listeners = new Set();
    this.ready = new Promise((res, rej) => {
      const ws = this.ws = new WebSocket(url);
      ws.onopen = () => res();
      ws.onerror = () => rej(new Error('Could not connect to Edge over the DevTools protocol.'));
      ws.onclose = () => { for (const w of this.waiting.values()) w.reject(new Error('Edge closed the DevTools connection.')); this.waiting.clear(); };
      ws.onmessage = e => {
        const m = JSON.parse(typeof e.data === 'string' ? e.data : Buffer.from(e.data).toString('utf8'));
        if (m.id) {
          const w = this.waiting.get(m.id);
          if (!w) return;
          this.waiting.delete(m.id); clearTimeout(w.timer);
          if (m.error) w.reject(new Error(`${w.method}: ${m.error.message}`)); else w.resolve(m.result);
        } else for (const fn of this.listeners) fn(m);
      };
    });
  }
  send(method, params = {}, sessionId) {
    const id = ++this.seq;
    return new Promise((resolve, reject) => {
      const timer = setTimeout(() => { this.waiting.delete(id); reject(new Error(`${method}: no answer from Edge in 30 s`)); }, 30000);
      this.waiting.set(id, { resolve, reject, timer, method });
      this.ws.send(JSON.stringify({ id, method, params, ...(sessionId ? { sessionId } : {}) }));
    });
  }
  on(fn) { this.listeners.add(fn); return () => this.listeners.delete(fn); }
}

function freePort() {
  return new Promise((res, rej) => { const s = netServer(); s.once('error', rej); s.listen(0, '127.0.0.1', () => { const p = s.address().port; s.close(() => res(p)); }); });
}

async function launchEdge() {
  if (!existsSync(EDGE)) throw new Error(`Edge isn't at ${EDGE} (set EDGE_PATH).`);
  const profile = mkdtempSync(join(tmpdir(), 'shalam-headless-'));
  const port = await freePort();
  const proc = spawn(EDGE, ['--headless=new', '--inprivate', '--disable-sync', `--user-data-dir=${profile}`, `--remote-debugging-port=${port}`,
    '--no-first-run', '--no-default-browser-check', '--disable-background-networking', '--disable-component-update',
    '--disable-extensions', '--disable-gpu', '--mute-audio', 'about:blank'], { stdio: 'ignore' });
  let version = null;
  for (let i = 0; i < 100 && !version; i++) {
    try { const r = await fetch(`http://127.0.0.1:${port}/json/version`); if (r.ok) version = await r.json(); } catch (e) { await sleep(150); }
  }
  if (!version) { proc.kill(); throw new Error("Edge didn't start its DevTools endpoint."); }
  const cdp = new Cdp(version.webSocketDebuggerUrl);
  await cdp.ready;
  return { cdp, proc, profile, version: version.Browser };
}

// A Edge window that signs in to anything, or turns on sync, fails the run.
async function signInWindows(cdp) {
  const { targetInfos } = await cdp.send('Target.getTargets');
  return targetInfos.map(t => t.url).filter(u => /sync|signin|sign-in|login\.|microsoftonline|account/i.test(u));
}

const ALL_NET = [];   // every request any tab made in the whole run, as Edge saw it

// A tab in its own new browser context, so its browser storage starts empty (the email marks, the saved locations).
async function openTab(cdp, init) {
  const { browserContextId } = await cdp.send('Target.createBrowserContext', { disposeOnDetach: true });
  const { targetId } = await cdp.send('Target.createTarget', { url: 'about:blank', browserContextId });
  const { sessionId } = await cdp.send('Target.attachToTarget', { targetId, flatten: true });
  const t = { targetId, sessionId, browserContextId, blocked: [], net: [], console: [], exceptions: [] };
  t.send = (m, p) => cdp.send(m, p, sessionId);
  t.off = cdp.on(m => {
    if (m.sessionId !== sessionId) return;
    const p = m.params || {};
    if (m.method === 'Fetch.requestPaused') {
      const r = { method: p.request.method, url: p.request.url };
      t.net.push(r); ALL_NET.push(r);
      let u = null;
      try { u = new URL(p.request.url); } catch (e) { /* refused below */ }
      if (u && ((u.protocol === 'http:' && u.hostname === '127.0.0.1') || u.protocol === 'data:' || u.protocol === 'blob:')) t.send('Fetch.continueRequest', { requestId: p.requestId }).catch(() => {});
      else { t.blocked.push(p.request.url); t.send('Fetch.failRequest', { requestId: p.requestId, errorReason: 'BlockedByClient' }).catch(() => {}); }
    } else if (m.method === 'Runtime.consoleAPICalled' && p.type === 'error') {
      t.console.push((p.args || []).map(a => a.description || a.value || '').join(' ').slice(0, 300));
    } else if (m.method === 'Runtime.exceptionThrown') {
      const d = p.exceptionDetails || {};
      t.exceptions.push(String((d.exception && d.exception.description) || d.text || 'error').slice(0, 300));
    }
  });
  await t.send('Fetch.enable', { patterns: [{ urlPattern: '*' }] });
  await t.send('Page.enable');
  await t.send('Runtime.enable');
  await t.send('Emulation.setDeviceMetricsOverride', { width: 1440, height: 1000, deviceScaleFactor: 1, mobile: false });
  await t.send('Page.addScriptToEvaluateOnNewDocument', { source: init });
  await t.send('Page.bringToFront').catch(() => {});
  t.eval = async expr => {
    const r = await t.send('Runtime.evaluate', { expression: expr, returnByValue: true, awaitPromise: true });
    if (r.exceptionDetails) throw new Error('In the page: ' + ((r.exceptionDetails.exception && r.exceptionDetails.exception.description) || r.exceptionDetails.text));
    return r.result.value;
  };
  t.waitFor = async (expr, what, ms = 45000) => {
    const end = Date.now() + ms;
    while (Date.now() < end) { if (await t.eval(expr).catch(() => false)) return true; await sleep(150); }
    throw new Error(`Timed out after ${ms / 1000} s waiting for ${what}.`);
  };
  t.snap = () => t.eval(`(${snapshot})()`);
  t.click = sel => t.eval(`(() => { const el = document.querySelector(${JSON.stringify(sel)}); if (!el) throw new Error('no ${sel.replace(/'/g, '')}'); el.click(); return true; })()`);
  t.reload = async () => {
    await t.eval('window.__stale = 1');
    await t.send('Page.reload', { ignoreCache: true });
    await t.waitFor(READY, 'the page to load again');
  };
  t.shot = async name => {
    const m = await t.send('Page.getLayoutMetrics');
    const s = m.cssContentSize || m.contentSize;
    const r = await t.send('Page.captureScreenshot', { format: 'png', captureBeyondViewport: true,
      clip: { x: 0, y: 0, width: Math.ceil(s.width), height: Math.min(Math.ceil(s.height), 6000), scale: 1 } });
    const file = join(OUT, `headless-${name}.png`);
    writeFileSync(file, Buffer.from(r.data, 'base64'));
    return file;
  };
  t.close = async () => {
    t.off();
    await cdp.send('Target.closeTarget', { targetId }).catch(() => {});
    await cdp.send('Target.disposeBrowserContext', { browserContextId }).catch(() => {});
  };
  return t;
}

// Runs the page's 5-minute refresh now and waits for it to finish. Answers how many held timers ran.
async function runRefresh(t, w, expectCalls = true) {
  const from = w.calls.length;
  const fired = await t.eval('window.__headless.fireLong()');
  const end = Date.now() + 15000;
  while (expectCalls && fired && w.calls.length === from && Date.now() < end) await sleep(100);
  await t.waitFor(IDLE, 'the refresh to finish', 30000);
  return { fired, calls: w.calls.slice(from) };
}

// ------------------------------------------------------------------ scenarios

const sendMails = calls => calls.filter(c => c.method === 'POST' && c.path === '/me/sendMail');
const banner = (s, kind) => s.banners.filter(b => b.kind === kind).map(b => b.text).join(' | ');
const pair = (card, label) => { const p = card && card.pairs.find(x => x[0] === label); return p ? p[1] : null; };
const rc = e => `${fmtN(e.read)} read / ${fmtN(e.counted)} counted`;
const mailText = c => (c && c.body && c.body.message) || {};
const mailKey = (today, k) => `inv:mailed:${C.isoDay(today)}:${k}`;
// The read-only rule, three ways: as the page asked (the fake fetch), as Edge sent it, and as the stand-in got it.
const okPage = r => r.method === 'GET' || (r.method === 'POST' && r.url === GRAPH + '/me/sendMail');
const okNet = origin => r => r.method === 'GET' || (r.method === 'POST' && r.url === `${origin}/__graph/me/sendMail`);
const okCall = x => x.method === 'GET' || (x.method === 'POST' && x.path === '/me/sendMail');

function greenExcept(c, s, except) {
  for (const k of ['msl', 'tracker', 'pobook', 'ct', 'forwarder', 'drayage', 'upc'].filter(k => !except.includes(k))) {
    c(`${NAMES[k]} chip green`, s.chips[k] && s.chips[k].state === 'ok', s.chips[k] ? `${s.chips[k].state}: ${s.chips[k].text}` : 'no chip');
  }
}
function oneEmail(c, s, calls, k, problem, today) {
  const source = NAMES[k], sends = sendMails(calls);
  c('exactly one sendMail', sends.length === 1, `${sends.length} sent`);
  const msg = mailText(sends[0]);
  c(`the subject names ${source} and the problem`, msg.subject && msg.subject.includes(`: ${source} — ${problem}`), msg.subject);
  c('sent to errorEmailTo from config.js, plain text, kept in Sent Items',
    sends[0] && msg.toRecipients && msg.toRecipients.length === 1 && msg.toRecipients[0].emailAddress.address === MAIL_TO &&
    msg.body && msg.body.contentType === 'Text' && sends[0].body.saveToSentItems === true, JSON.stringify(msg.toRecipients));
  const key = mailKey(today, k);
  c(`this browser noted today's email ("${key}")`, !!(s.store && s.store[key]), JSON.stringify(s.store));
  return { sends, msg };
}
function stopChecks(c, s, w, scpWords) {
  const stop = s.banners.find(b => b.id === 'stopBanner');
  c('red stop banner with Sign in again', !!stop && stop.kind === 'err' && /Stopped: nothing is loaded/.test(stop.text) && stop.buttons.includes('Sign in again'), JSON.stringify(s.banners));
  for (const wd of scpWords) c(`the banner names ${wd}`, !!stop && stop.text.includes(wd), stop && stop.text);
  c('not one request to SharePoint (no Graph call, no download)', w.calls.length === 0 && w.downloads.length === 0,
    w.calls.map(x => `${x.method} ${x.path}`).concat(w.downloads.map(d => d.name)).join('; '));
  c('no source loads: every chip grey "not read"', KEYS.every(k => s.chips[k] && s.chips[k].state === 'off' && /not read/.test(s.chips[k].text)),
    KEYS.map(k => s.chips[k] && `${s.chips[k].state} ${s.chips[k].text}`).join(' | '));
  c('no email', sendMails(w.calls).length === 0);
  c('health line says it stopped', /Stopped/.test(s.health), s.health);
}

// ---- step 2: what the checks read from the Inventory, Pull from China and Problems tabs, in one call
const S2 = `(() => {
  const q = s => document.querySelector(s), txt = el => (el ? el.textContent.replace(/\\s+/g, ' ').trim() : '');
  const shown = el => !!el && !el.hidden && !el.closest('[hidden]') && getComputedStyle(el).display !== 'none';
  const r1 = q('#invTable tbody tr.r');
  return {
    tiles: [...document.querySelectorAll('#tiles .tile')].map(t => ({ k: t.dataset.t || '', cls: t.className, n: txt(t.querySelector('.n')), cs: txt(t.querySelector('.cs')), lbl: txt(t.querySelector('.lbl')) })),
    invTr: document.querySelectorAll('#invTable tbody tr.r').length, invFirst: r1 ? +r1.dataset.i : -1,
    invEmpty: txt(q('#invTable tbody td.empty')), info: txt(q('#fInfo')), fbarIn: (q('#fbar').parentElement || {}).id || '',
    pullTr: document.querySelectorAll('#pullTable tbody tr.r').length, pullEmpty: txt(q('#pullTable tbody td.empty')), cPull: txt(q('#cPull')),
    panel: shown(q('#panel')), panelTitle: txt(q('#panelTitle')), panelSub: txt(q('#panelSub')), panelTr: document.querySelectorAll('#panelTable tbody tr.r').length,
    panelHead: [...document.querySelectorAll('#panelTable thead th')].map(txt), panelFirst: txt(q('#panelTable tbody tr.r')),
    balance: { cls: q('#balance').className, text: txt(q('#balance')) },
    lead: shown(q('#invState')) ? txt(q('#invState')) : '', pullLead: shown(q('#pullState')) ? txt(q('#pullState')) : '',
    scrollInv: q('#invPane').scrollTop, fq: q('#fq').value, fFac: q('#fFac').value,
    heads: [...document.querySelectorAll('#invTable thead th')].map(th => ({ t: txt(th), sort: th.getAttribute('aria-sort') || '' })),
    issGroups: [...document.querySelectorAll('#issTable tr.gh')].map(txt), issRows: document.querySelectorAll('#issTable tbody tr:not(.gh)').length,
    cProb: txt(q('#cProb')), pqInfo: txt(q('#pqInfo')), health: txt(q('#health')),
    banners: [...document.querySelectorAll('#banners .banner')].map(b => ({ id: b.id, cls: b.className, text: txt(b) })),
  };
})()`;
const s2 = t => t.eval(S2);
const TILE_KEYS = ['inChina', 'water', 'port', 'incoming', 'pullFromChina', 'unmatched'];
const tileN = (s, k) => { const x = s.tiles.find(y => y.k === k); return x ? x.n : null; };
// A tile's number reads "12,345 units".
const unitsIs = (txt, v) => typeof txt === 'string' && txt.replace(/units$/, '').trim() === v.toLocaleString('en-US');
const tilesAre = (s, units) => !!units && TILE_KEYS.every(k => unitsIs(tileN(s, k), units[k]));
const tilesSay = (s, units) => TILE_KEYS.map(k => `${k} ${tileN(s, k)} vs ${units ? units[k] : '?'}`).join('; ');
const liveTiles = s => s.tiles.filter(x => TILE_KEYS.includes(x.k));
const notLoadedEverywhere = s => liveTiles(s).length === TILE_KEYS.length && liveTiles(s).every(x => x.n === 'not loaded' && !/\d/.test(x.cs));
const setVal = (t, id, v) => t.eval(`(() => { const el = document.getElementById(${JSON.stringify(id)}); el.value = ${JSON.stringify(v)}; el.dispatchEvent(new Event(el.tagName === 'SELECT' ? 'change' : 'input', { bubbles: true })); })()`);
const scrollTo = (t, id, y) => t.eval(`(() => { const p = document.getElementById(${JSON.stringify(id)}); p.scrollTop = ${y}; p.dispatchEvent(new Event('scroll')); return p.scrollTop; })()`);
const key = (t, k) => t.eval(`document.dispatchEvent(new KeyboardEvent('keydown', { key: ${JSON.stringify(k)}, bubbles: true }))`);
// The tiles' numbers from the engine's allocations that pass a filter, added up as the page does (landed Unmatched left out).
function tileSums(R, pred) {
  const g = Object.fromEntries(TILE_KEYS.map(k => [k, 0]));
  for (const a of R.allocs) {
    if (!pred(a) || (a.bucket === 'unmatched' && a.stage === 'Landed')) continue;
    const ks = ST.IN_CHINA.includes(a.bucket) ? ['inChina', 'incoming'] : a.bucket === 'water' || a.bucket === 'port' ? [a.bucket, 'incoming']
      : a.bucket === 'pullFromChina' || a.bucket === 'unmatched' ? [a.bucket] : [];
    for (const k of ks) g[k] += a.units;
  }
  return g;
}
const BIG = { mslFile: 'msl-big.xlsx', trkFile: 'trk-big.xlsx', ctFile: 'ct-big.xlsx' };
const BIG_FILES = { msl: 'msl-big.xlsx', trk: 'trk-big.xlsx', ct: 'ct-big.xlsx' };
const REAL = { mslFile: 'real-msl.xlsx', trkFile: 'real-trk.xlsx', ctFile: 'real-ct.xlsx', fwdFile: 'real-fwd.xlsx', drayFile: 'real-dray.csv', real: true };
const SELLW = { mslFile: 'msl-sell.xlsx', trkFile: 'trk-sell.xlsx', ctFile: 'ct-sell.xlsx' };
const SELL_FILES = { msl: 'msl-sell.xlsx', trk: 'trk-sell.xlsx', ct: 'ct-sell.xlsx' };
const SELL_ATS = { age: 2, file: 'ats-sell.xlsx' };

// ---- Available to sell: what the checks read from its tab and the side panel, in one call. A cell's value leaves out
// its second line (.u) and tags; `u` is the second line.
const SL = `(() => {
  const q = s => document.querySelector(s), txt = el => (el ? el.textContent.replace(/\\s+/g, ' ').trim() : '');
  const shown = el => !!el && !el.hidden && !el.closest('[hidden]') && getComputedStyle(el).display !== 'none';
  const plain = el => { if (!el) return ''; const x = el.cloneNode(true); x.querySelectorAll('.u,.tag,.ar').forEach(y => y.remove()); return x.textContent.replace(/\\s+/g, ' ').trim(); };
  const sub = el => { const u = el && el.querySelector('.u'); return u ? u.textContent.replace(/\\s+/g, ' ').trim() : ''; };
  const ph = [...document.querySelectorAll('#panelTable thead th')].map(plain);
  const on = q('#tabs .tab.on');
  return {
    tab: on ? on.dataset.tab : '', shown: shown(q('#tab-sell')), lead: shown(q('#sellState')) ? txt(q('#sellState')) : '',
    note: txt(q('#sellNote')), noteCls: (q('#sellNote .banner') || {}).className || '',
    tiles: [...document.querySelectorAll('#sellTiles .tile')].map(t => ({ k: t.dataset.t || '', btn: t.tagName === 'BUTTON', cls: t.className, n: txt(t.querySelector('.n')),
      cs: [...t.querySelectorAll('.cs')].map(txt), lbl: txt(t.querySelector('.lbl')), title: t.title || '' })),
    heads: [...document.querySelectorAll('#sellTable thead th')].map(th => ({ t: plain(th), sort: th.getAttribute('aria-sort') || '' })),
    tr: document.querySelectorAll('#sellTable tbody tr.r').length, first: q('#sellTable tbody tr.r') ? +q('#sellTable tbody tr.r').dataset.i : -1,
    empty: txt(q('#sellTable tbody td.empty')),
    rows: [...document.querySelectorAll('#sellTable tbody tr.r')].slice(0, 80).map(tr => {
      const tds = [...tr.children], sp = tds.find(td => td.colSpan > 1), b = k => txt(tr.querySelector('button.nb[data-b="' + k + '"]')), tot = tds[tds.length - 2];
      return { i: +tr.dataset.i, style: txt(tr.querySelector('button.sb')), upc: sub(tds[0]), tags: [...tr.querySelectorAll('.tag')].map(txt), cells: tds.length,
        span: sp ? plain(sp) : null, inWh: sp ? null : plain(tds[3]), orders: sp ? null : plain(tds[4]), nowCell: sp ? null : plain(tds[5]), casePack: plain(tds[2]),
        now: b('now'), fromChina: b('fromChina'), arriving: b('arriving'), total: b('total'), sold: b('soldChina'),
        totalCell: plain(tot), totalCls: tot.className, totalTitle: tot.title || '', firstPull: plain(tds[sp ? 5 : 7]), firstArrival: plain(tds[sp ? 7 : 9]) };
    }),
    info: txt(q('#fInfo')), fbarIn: (q('#fbar').parentElement || {}).id || '', availShown: shown(q('#fAvailWrap')), stageShown: shown(q('#fStage')), typeShown: shown(q('#fType')),
    panel: shown(q('#panel')), panelLbl: txt(q('#panelLbl')), panelTitle: txt(q('#panelTitle')), panelSub: txt(q('#panelSub')), panelNote: txt(q('#panelNote')),
    panelStock: shown(q('#panelStock')) ? txt(q('#panelStock')) : '', panelHead: ph, panelTr: document.querySelectorAll('#panelTable tbody tr.r').length,
    lots: [...document.querySelectorAll('#panelTable tbody tr.r')].slice(0, 80).map(tr => Object.fromEntries(ph.map((h, j) => {
      const td = tr.children[j]; return [h, { v: plain(td), u: sub(td), title: (td && td.title) || '', cls: (td && td.className) || '' }]; }))),
    scroll: q('#sellPane').scrollTop,
  };
})()`;
const sl = t => t.eval(SL);
const SELL_TILE_KEYS = ['now', 'fromChina', 'arriving', 'total'];
const sellTile = (p, k) => p.tiles.find(x => x.k === k) || { n: null, cs: [], cls: '', lbl: '' };
const sellTilesSay = (p, S) => [...SELL_TILE_KEYS, 'oversold'].map(k => `${k} "${sellTile(p, k).n}" vs ${S ? S.totals[k] : '?'}`).join('; ');
// "1,234 cases + 7 loose units + 56 units with no case pack", as the tiles write it (whole cases only).
const casesSay = (c, np, lo) => [c ? `${fmtN(c)} case${c === 1 ? '' : 's'}` : '', lo ? `${fmtN(lo)} loose unit${lo === 1 ? '' : 's'}` : '',
  np ? `${fmtN(np)} units with no case pack` : ''].filter(Boolean).join(' + ');
const nz = v => (v ? fmtN(v) : '');
// A date in running text: m/d this year, m/d/yyyy otherwise (as "10 days before ETD 11/14").
const md = (d, today) => (new Date(d * 864e5).getUTCFullYear() === new Date(today * 864e5).getUTCFullYear() ? C.fmtDay(d).replace(/\/\d{4}$/, '') : C.fmtDay(d));
// What the side panel's "Available from" cell must say for a lot (Saul's rule in words), worked out here: the date
// shown (or "now"), and the words that must be in its tooltip.
function whenSays(l, today) {
  if (!l.inChina) return l.eta == null ? { main: 'no ETA yet', words: [] } : { main: C.fmtDay(l.eta), words: [`ETA ${md(l.eta, today)}`] };
  if (l.pullFrom == null) return { main: 'no date', words: [] };
  const date = (l.pullRule === 'order' ? 'about ' : '') + C.fmtDay(l.pullFrom);
  const words = l.pullRule === 'etd' ? [`10 days before ETD ${md(l.etd, today)}`] : l.pullRule === 'crd' ? [`5 days before CRD ${md(l.readyDate, today)}`]
    : ['75 days after the order date', md(l.dates.orderDate, today), md(l.pullFrom, today)];
  return { main: l.pullFrom < today ? 'now' : date, words: [date, ...words] };
}
// A style's First pull date cell, worked out here.
const firstPullSays = (s, today) => (s.firstPull == null ? (s.lots.some(l => l.inChina && l.available > 0) ? 'no date' : '')
  : s.firstPull < today ? 'now' : (s.firstPullEstimate ? 'about ' : '') + C.fmtDay(s.firstPull));
// The table's columns, in order.
const SELL_HEADS = ['Style', 'Description', 'Case pack', 'In warehouse', 'Customer orders', 'Available now', 'Available from China', 'First pull date', 'Arriving', 'First arrival',
  'Total available', 'Sold (ships from China)'];
// Every style row drawn against the engine: stock, after orders, from China, arriving, total, sold, dates. Answers what differs.
function sellRowsDiffer(p, S, today, stock) {
  const by = new Map(S.styles.map(s => [s.styleRaw || s.style, s])), bad = [];
  for (const r of p.rows) {
    const s = by.get(r.style);
    if (!s) { bad.push(`${r.style}: not in the engine`); continue; }
    const want = { fromChina: nz(s.fromChina), arriving: nz(s.arriving), sold: nz(s.soldChina), firstPull: firstPullSays(s, today), firstArrival: s.firstArrival == null ? '' : C.fmtDay(s.firstArrival) };
    if (stock === 'ok' && !s.flags.includes('not-on-report') && !s.flags.includes('stock-unknown')) Object.assign(want, { inWh: nz(s.inWh), orders: nz(s.orders), now: nz(s.now), total: nz(s.total) });
    else if (stock === 'ok') Object.assign(want, { span: s.flags.includes('not-on-report') ? 'not on report' : undefined, total: nz(s.total) });
    else Object.assign(want, { span: stock, totalCell: stock });
    for (const [k, v] of Object.entries(want)) if (v !== undefined && r[k] !== v) bad.push(`${r.style} ${k}: "${r[k]}", not "${v}"`);
  }
  return bad;
}

const SCENARIOS = [
  {
    name: 'ok', title: 'everything fine', world: {},
    async run({ t, w, c, e }) {
      let s = await t.snap();
      c('signed in with the fake account (no real sign-in)', s.appShown && !s.signinShown && s.who.includes(ACCOUNT.name), s.who);
      greenExcept(c, s, []);
      c('ATS chip grey "no report yet", with where to save it', s.chips.ats && s.chips.ats.state === 'off' && /no report yet/.test(s.chips.ats.text) &&
        s.chips.ats.title.includes('Save the weekly ATS report into Stock/ATS on the TestSite site.'), s.chips.ats && `${s.chips.ats.text} | ${s.chips.ats.title}`);
      for (const k of ['msl', 'tracker', 'pobook', 'ct', 'forwarder', 'drayage', 'upc']) {
        c(`${NAMES[k]} chip shows "${rc(e[k])}"`, s.chips[k] && s.chips[k].text.includes(rc(e[k])), s.chips[k] && s.chips[k].text);
      }
      c('chips show who saved the file', s.chips.ct.text.includes('Test Saver'), s.chips.ct.text);
      c('Container Tracking chip links to the file', s.chips.ct.href === webUrl('itm-ct'), s.chips.ct.href);
      const shares = w.calls.filter(x => x.path.startsWith('/shares/'));
      c('both sharing links resolved through /shares/u!…/driveItem, without the redeem header',
        shares.length === 2 && shares.every(x => !/redeem/i.test(x.prefer)), shares.map(x => `${x.path} prefer="${x.prefer}"`).join('; '));
      c('each /sites/… path resolved through its site and the library matched by its address',
        ['/sites/met1sol.sharepoint.com:/sites/TestSite', '/sites/met1sol.sharepoint.com:/sites/TestReports', '/sites/site-test/drives', '/sites/site-rep/drives',
          `/drives/${SRC}/root:/Logistics/Test Container Sheet.xlsx`, `/drives/${REP}/root:/Shipping/Forwarder`, `/drives/${REP}/root:/Shipping/Drayage`, `/drives/${SRC}/root:/Stock/ATS`]
          .every(p => w.calls.some(x => x.path === p && x.status === 200)), w.calls.filter(isResolve).map(x => `${x.status} ${x.path}`).join('; '));
      const locKeys = Object.keys(s.store || {}).filter(k => k.startsWith('inv:loc:'));
      c('the resolved ids are kept in this browser under "inv:loc:" + the exact location text',
        Object.values(LOC).every(v => locKeys.includes('inv:loc:' + v)) && locKeys.length === 6, locKeys.join(' ; '));
      const msl = s.cards.msl;
      c(`Master Sheet Live card: ${e.msl.lines} lines, ${e.msl.open} open, ${e.msl.shipped} shipped (PO Book marks applied)`,
        pair(msl, 'Lines') === String(e.msl.lines) && pair(msl, 'Open') === String(e.msl.open) && pair(msl, 'Shipped') === String(e.msl.shipped),
        JSON.stringify(msl && msl.pairs));
      c(`PO Book lines card: ${e.pobook.read} items read, ${e.pobook.counted} line marked`,
        pair(s.cards.pobook, 'List items read') === String(e.pobook.read) && pair(s.cards.pobook, 'Lines marked in the PO Book') === String(e.pobook.counted),
        JSON.stringify(s.cards.pobook && s.cards.pobook.pairs));
      c('Forwarder: the newest .xlsx is used (it is on page 2 of the folder listing)', s.chips.forwarder.title.includes('forwarder-report-new.xlsx'), s.chips.forwarder.title);
      c('Forwarder: lock file, folder and other file types skipped', !w.calls.some(x => /itm-fwd-(lock|sub|txt)/.test(x.path)));
      c('Drayage: the newest .csv is used', s.chips.drayage.title.includes('drayage-report-new.csv'), s.chips.drayage.title);
      c('UPC_MASTER: both pages of list items read', pair(s.cards.upc, 'Rows') === String(e.upc.read), JSON.stringify(s.cards.upc && s.cards.upc.pairs));
      const lists = w.calls.filter(x => /\/lists\/lst-(pob|upc)\/items/.test(x.path) && !/skiptoken/.test(x.query));
      c('list items asked for with $expand=fields and $top=999', lists.length === 2 && lists.every(x => /\$expand=fields/.test(decodeURIComponent(x.query)) && /\$top=999/.test(x.query)), lists.map(x => x.query).join(' ; '));
      c('PO Book copy check ran and matched (PO Book lines chip stays green)', w.poBookGets >= 1 && s.chips.pobook.state === 'ok', `fetched ${w.poBookGets}×`);
      c('no banner', s.banners.length === 0, JSON.stringify(s.banners));
      c('health line: all sources loaded', /All sources loaded/.test(s.health), s.health);
      c('no email', sendMails(w.calls).length === 0);
      c('no ⚙ Sources button and no popup in the page', !s.sourcesButton && !s.dialog);
      c(`Problems count = every issue the parsers, the stage engine and the selling engine found (${e.totalIssues}, ${e.incoming.issues} of them Incoming, ${e.sell.issues} Available to sell)`, s.cProb === String(e.totalIssues), s.cProb);
      c('no console errors', t.console.length === 0, t.console.join(' | '));
      c('Available to sell is the first tab and is shown at start, with the sources box (a card per source) under it',
        s.tabOrder.join(',') === 'sell,inv,pull,problems' && s.tabs.sell && !s.tabs.inv && s.srcBoxIn === 'tab-sell' && s.cardCount === 8, `${s.tabOrder} · ${JSON.stringify(s.tabs)} · box in ${s.srcBoxIn} · ${s.cardCount} cards`);
      t.shots.push(await t.shot('ok'));

      await t.click('#tabs .tab[data-tab="inv"]');
      s = await t.snap();
      c('Where it is: the sources box moves with it', s.tabs.inv && !s.tabs.sell && s.srcBoxIn === 'tab-inv', `${JSON.stringify(s.tabs)} · box in ${s.srcBoxIn}`);
      await t.click('#tabs .tab[data-tab="pull"]');
      s = await t.snap();
      c('Pull from China tab draws', s.tabs.pull && !s.tabs.sell && !s.tabs.inv && !s.tabs.problems, JSON.stringify(s.tabs));
      await t.click('#tabs .tab[data-tab="problems"]');
      s = await t.snap();
      c('Problems tab: sources table with 8 rows', s.tabs.problems && s.srcRows === 8, `${s.srcRows} rows`);
      c(`Problems tab: every issue listed (${e.totalIssues})`, s.issueRows === e.totalIssues, `${s.issueRows} rows`);
      const groups = ['tracker', 'msl', 'pobook', 'ct', 'forwarder', 'drayage', 'upc', 'incoming', 'sell', 'ats'].filter(k => e[k].issues > 0).length;
      c(`Problems tab: issues grouped by source (${groups} groups, Incoming and Available to sell included) with counts in plain words, not codes`,
        s.issueGroups.length === groups && s.issueGroups.some(g => /^UPC_MASTER · \d+ issues?/.test(g) && /UPC not 12-13 digits \d/.test(g)) &&
        !s.issueGroups.some(g => /\bbadupc\b|\bmslonly\b|\bdupround\b/.test(g)), s.issueGroups.join(' | '));
      const inc = s.issueGroups.find(g => g.startsWith('Incoming')) || '';
      const incCodes = [...new Set(e.R.issues.map(i => i.code))];
      c(`Problems tab: the Incoming group has the stage engine's ${e.incoming.issues} issues (${incCodes.join(', ')}), in plain words`,
        new RegExp(`^Incoming · ${e.incoming.issues} issues?`).test(inc) && !incCodes.some(k => inc.includes(k) || inc.includes(k.replace(/-/g, ' '))), inc);
      const sg = s.issueGroups.find(g => g.startsWith('Available to sell')) || '';
      c(`Problems tab: right after it, "Available to sell" with the selling engine's ${e.sell.issues} (no ATS report: customer orders not known), in plain words`,
        s.issueGroups[1] === sg && new RegExp(`^Available to sell · ${e.sell.issues} issues?`).test(sg) && /Customer orders not known/.test(sg) && !/orders-unknown/.test(sg), s.issueGroups.join(' | '));
      c('PO Book lines: the item with no LineKey is listed', s.issueGroups.some(g => /^PO Book lines · 1 issue/.test(g) && /Item with no LineKey 1/.test(g)), s.issueGroups.join(' | '));
      await t.eval(`(() => { const q = document.getElementById('pq'); q.value = 'badupc'; q.dispatchEvent(new Event('input', { bubbles: true })); })()`);
      await sleep(500);
      s = await t.snap();
      c('Problems filter box narrows the list', s.issueRows === e.upc.badupc && new RegExp(`^${e.upc.badupc} of ${e.totalIssues} issues match`).test(s.pqInfo), `${s.issueRows} rows, "${s.pqInfo}"`);
      t.shots.push(await t.shot('ok-problems'));
      await t.click('#tabs .tab[data-tab="inv"]');
    },
  },
  {
    name: 'ok-reload', title: 'reload in the same browser: the saved locations are reused, the last tab shown again', after: 'ok',
    async run({ t, w, c }) {
      const from = w.calls.length;
      await t.reload();
      const s = await t.snap(), calls = w.calls.slice(from);
      c('the tab last used (Where it is) is shown again, kept in this browser (inv:tab2)', s.tabs.inv && !s.tabs.sell && (s.store || {})['inv:tab2'] === 'inv', `${JSON.stringify(s.tabs)} · ${JSON.stringify(s.store && s.store['inv:tab2'])}`);
      await t.click('#tabs .tab[data-tab="sell"]');
      greenExcept(c, s, []);
      c('no location found afresh (no /shares, site or library lookups)', !calls.some(isLookup), calls.filter(isLookup).map(x => x.path).join('; '));
      c('the sharing links opened straight from their saved ids', ['itm-msl', 'itm-trk'].every(id => calls.some(x => x.path === `/drives/${SRC}/items/${id}`)), calls.map(x => x.path).join('; '));
      c('each /sites/… path looked up by its path in its saved library, not by a saved id',
        [`/drives/${SRC}/root:/Logistics/Test Container Sheet.xlsx`, `/drives/${REP}/root:/Shipping/Forwarder`, `/drives/${REP}/root:/Shipping/Drayage`, `/drives/${SRC}/root:/Stock/ATS`]
          .every(p => calls.some(x => x.path === p && x.status === 200)) && !calls.some(x => /\/items\/itm-(ct|fwd-dir|dray-dir|ats-dir)$/.test(x.path)),
        calls.map(x => x.path).join('; '));
    },
  },
  {
    name: 'ct-replaced', title: 'Container Tracking renamed away and a new file saved under its name', after: 'ok-reload',
    async run({ t, w, c }) {
      w.o.ctReplaced = true;
      const from = w.calls.length, dls = w.downloads.length;
      await t.reload();
      const s = await t.snap(), calls = w.calls.slice(from);
      c('Container Tracking reads the new file at the path in config.js', s.chips.ct.state === 'ok' && s.chips.ct.href === webUrl('itm-ct-new') && !s.chips.ct.title.includes('OLD'),
        `${s.chips.ct.href} | ${s.chips.ct.title}`);
      c('the old file (same id, new name) is not opened', !calls.some(x => x.path === `/drives/${SRC}/items/itm-ct`), calls.map(x => x.path).join('; '));
      c('the new file was downloaded', w.downloads.slice(dls).some(d => d.name === 'ct.xlsx'));
      w.o.ctReplaced = false;
    },
  },
  {
    name: 'cfg-change', title: 'config.js points Container Tracking somewhere else: resolved afresh', after: 'ok-reload',
    async run({ t, w, c }) {
      w.o.sources = { ct: CT_V2 };
      const from = w.calls.length;
      await t.reload();
      const s = await t.snap(), calls = w.calls.slice(from);
      c('the new location text is resolved (its library path is looked up)', calls.some(x => x.path === `/drives/${SRC}/root:/Logistics/Test Container Sheet v2.xlsx` && x.status === 200),
        calls.filter(isResolve).map(x => x.path).join('; '));
      c('Container Tracking chip green, reading the new file', s.chips.ct.state === 'ok' && s.chips.ct.title.includes('Test Container Sheet v2.xlsx'), s.chips.ct.title);
      c('the other locations are not found afresh', !calls.some(x => x.path.startsWith('/shares/') || x.path === '/sites/met1sol.sharepoint.com:/sites/TestReports' || x.path === '/sites/site-rep/drives'),
        calls.filter(isLookup).map(x => x.path).join('; '));
      const keys = Object.keys(s.store || {}).filter(k => k.startsWith('inv:loc:'));
      c('the old location is dropped from this browser, the new one kept', keys.includes('inv:loc:' + CT_V2) && !keys.includes('inv:loc:' + LOC.ct), keys.join(' ; '));
    },
  },
  {
    name: 'ctbad', title: 'Container Tracking has no "Quantity" column', world: { ctFile: 'ct-bad.xlsx', poBook: 'missing' },
    async run({ t, w, c, today }) {
      const s = await t.snap();
      const ct = s.chips.ct;
      c('Container Tracking chip red', ct && ct.state === 'err', ct && `${ct.state}: ${ct.text}`);
      c('chip says "missing column" and names "Quantity"', /missing column/.test(ct.text) && /"Quantity"/.test(ct.title), ct.title);
      c('Problems lists the headers seen', /headers seen[^]*Pieces/i.test(s.srows.ct), s.srows.ct);
      const red = banner(s, 'err');
      c('red banner: stage numbers not loaded, naming Container Tracking and "Quantity"', /stage numbers are not loaded/i.test(red) && red.includes('Container Tracking') && red.includes('Quantity'), red);
      greenExcept(c, s, ['ct']);
      c("PO Book lines chip says nothing when the PO Book page can't be fetched", s.chips.pobook.state === 'ok' && w.poBookGets >= 1, s.chips.pobook.text);
      const { msg } = oneEmail(c, s, w.calls, 'ct', 'missing column "Quantity"', today);
      const body = (msg.body && msg.body.content) || '';
      c('email body: file, problem, headers seen, who had the page open, link',
        body.includes('Test Container Sheet.xlsx') && body.includes('Quantity') && /headers seen[^]*Pieces/i.test(body) &&
        body.includes(ACCOUNT.name) && body.includes(ACCOUNT.username) && body.includes(webUrl('itm-ct')), body);
      c('page says "Saul has been emailed about this."', red.includes('Saul has been emailed about this.') && s.srows.ct.includes('Saul has been emailed about this.'), red);
      t.shots.push(await t.shot('ctbad'));
    },
  },
  {
    name: 'reload', title: 'same day: a reload, then the 5-minute refresh, send nothing more', after: 'ctbad',
    async run({ t, w, c, today }) {
      const from = w.calls.length;
      await t.reload();
      let s = await t.snap();
      const calls = w.calls.slice(from);
      c('Container Tracking chip still red', s.chips.ct.state === 'err', s.chips.ct.text);
      c('the reload read the sources again', calls.some(x => x.path === CT_PATH), `${calls.length} calls`);
      c('no second sendMail after the reload', sendMails(calls).length === 0, `${sendMails(calls).length} sent`);
      c('page says Saul was already emailed today', banner(s, 'err').includes('Saul was already emailed about Container Tracking today.'), banner(s, 'err'));
      const r = await runRefresh(t, w);
      s = await t.snap();
      c('the 5-minute refresh was waiting, and ran', r.fired === 1 && r.calls.some(x => x.path === CT_PATH), `fired ${r.fired}, ${r.calls.length} calls`);
      c('no sendMail from the 5-minute refresh', sendMails(r.calls).length === 0, `${sendMails(r.calls).length} sent`);
      c('one email in total for the day', sendMails(w.calls).length === 1, `${sendMails(w.calls).length} sent`);
      c('Container Tracking still red after the refresh', s.chips.ct.state === 'err', s.chips.ct.text);
      c(`the mark is still there ("${mailKey(today, 'ct')}")`, !!(s.store && s.store[mailKey(today, 'ct')]));
      t.shots.push(await t.shot('reload'));
    },
  },
  {
    name: 'nostorage', title: 'browser storage blocked: one email per page session', world: { ctFile: 'ct-bad.xlsx' }, init: { noStorage: true },
    async run({ t, w, c }) {
      let s = await t.snap();
      c('storage really is blocked in this page', s.store === null);
      c('the page still loads: Container Tracking red, the rest green', s.chips.ct.state === 'err' && s.chips.msl.state === 'ok' && s.chips.upc.state === 'ok', `${s.chips.ct.state} ${s.chips.msl.state} ${s.chips.upc.state}`);
      c('one sendMail', sendMails(w.calls).length === 1, `${sendMails(w.calls).length} sent`);
      const r = await runRefresh(t, w);
      c('the 5-minute refresh ran and sent nothing more (the in-memory mark)', r.fired === 1 && r.calls.length > 0 && sendMails(r.calls).length === 0, `fired ${r.fired}, ${sendMails(r.calls).length} sent`);
      const from = w.calls.length;
      await t.reload();
      s = await t.snap();
      c('a reload is a new page session, so it may send once more', sendMails(w.calls.slice(from)).length === 1, `${sendMails(w.calls.slice(from)).length} sent`);
    },
  },
  {
    name: 'mslbad', title: 'Master Sheet Live has no "Quantity" column', world: { mslFile: 'msl-bad.xlsx' },
    async run({ t, w, c, today }) {
      const s = await t.snap(), m = s.chips.msl;
      c('Master Sheet Live chip red "missing column", naming "Quantity"', m.state === 'err' && /missing column/.test(m.text) && /"Quantity"/.test(m.title), `${m.state}: ${m.title}`);
      c('Problems row has a "Headers seen:" line', /Headers seen: .*\bQty\b/.test(s.srows.msl), s.srows.msl);
      const red = banner(s, 'err');
      c('red banner: stage numbers not loaded, naming Master Sheet Live', /stage numbers are not loaded/i.test(red) && red.includes('Master Sheet Live') && red.includes('Quantity'), red);
      const p = await s2(t);
      c('every stage tile says "not loaded", with no partial number', notLoadedEverywhere(p), JSON.stringify(liveTiles(p).map(x => `${x.n} ${x.cs}`)));
      greenExcept(c, s, ['msl']);
      const { msg } = oneEmail(c, s, w.calls, 'msl', 'missing column "Quantity"', today);
      const body = (msg.body && msg.body.content) || '';
      c('email body has a "Headers seen:" line', /^Headers seen: .*\bQty\b/m.test(body), body);
    },
  },
  {
    name: 'fwdold', title: 'the newest forwarder file is 9 days old', world: { fwdAge: 9 },
    async run({ t, w, c, today }) {
      const s = await t.snap(), f = s.chips.forwarder;
      c('Forwarder chip amber "not used"', f.state === 'warn' && /not used/.test(f.text), `${f.state}: ${f.text}`);
      c('the chip names the newest file and its age', f.title.includes('forwarder-report-new.xlsx') && f.title.includes('9 days old'), f.title);
      const amber = banner(s, 'warn');
      c("amber banner: forwarder not used, and what the stages can't see without it", /Forwarder report not used:/.test(amber) && /once Container Tracking's ETD passes/.test(amber), amber);
      c('no red banner', !banner(s, 'err'), banner(s, 'err'));
      c('the old file is not opened or downloaded', !w.calls.some(x => /itm-fwd-(new|old)$/.test(x.path)) && !w.downloads.some(d => d.name === 'fwd.xlsx'),
        w.downloads.map(d => d.name).join(','));
      greenExcept(c, s, ['forwarder']);
      oneEmail(c, s, w.calls, 'forwarder', 'not used', today);
      t.shots.push(await t.shot('fwdold'));
    },
  },
  {
    name: 'atsold', title: 'the newest ATS report is 12 days old: read and used, amber, Saul emailed', world: { ats: { age: 12 } },
    async run({ t, w, c, today }) {
      const s = await t.snap(), a = s.chips.ats, name = atsName(w.o.ats);
      c('ATS chip amber "12 days old"', a.state === 'warn' && /12 days old/.test(a.text), `${a.state}: ${a.text}`);
      c('ATS chip names the newest Excel file: the lock file, the subfolder and a newer .txt skipped', a.title.includes(name) && !a.title.includes('~$') && !a.title.includes('read-me'), a.title);
      c('the newest report is downloaded once and read; the older one and the .txt are not opened',
        w.downloads.filter(d => d.name === 'ats.xlsx').length === 1 && w.calls.some(x => x.path === `/drives/${SRC}/items/itm-ats-1`) && !w.calls.some(x => /itm-ats-(0|txt)$/.test(x.path)),
        `${w.downloads.map(d => d.name).join(',')} · ${w.calls.filter(x => /itm-ats/.test(x.path)).map(x => x.path).join('; ')}`);
      c('the card: report date from the file name, 12 days old', pair(s.cards.ats, 'Report date') === `${C.fmtDay(today - 12)} (from the file name)` && pair(s.cards.ats, 'Age') === '12 days',
        JSON.stringify(s.cards.ats && s.cards.ats.pairs));
      c('health line names ATS', s.health.includes('ATS'), s.health);
      greenExcept(c, s, []);
      oneEmail(c, s, w.calls, 'ats', 'the newest report is 12 days old', today);
      const { S } = sellFor(w.fx, today, { ats: w.o.ats }), p = await sl(t);
      c(`Available to sell: an amber note says the report is 12 days old; customer orders are still taken off (Available now ${fmtN(S.totals.now)})`,
        /\bwarn\b/.test(p.noteCls) && /The ATS report is 12 days old/.test(p.note) && p.note.includes(`report date ${C.fmtDay(today - 12)}`) && !S.atsMissing && unitsIs(sellTile(p, 'now').n, S.totals.now),
        `${p.noteCls}: ${p.note} · ${sellTile(p, 'now').n}`);
    },
  },
  {
    name: 'atsfound', title: 'a recent ATS report: downloaded, read with its Grand Total check, chip green', world: { ats: { age: 2 } },
    async run({ t, w, c, today }) {
      const s = await t.snap(), a = s.chips.ats, { A, S } = sellFor(w.fx, today, { ats: w.o.ats });
      c('the report (read here) adds up, with its date in the name', A.ok && A.reportDay === today - 2 && A.counts.styleRows === 5 && A.styles.size === 4, JSON.stringify(A.counts));
      c(`ATS chip green: "${A.counts.rowsRead} read / ${A.counts.styleRows} counted", naming the file`, a.state === 'ok' && a.text.includes(`${fmtN(A.counts.rowsRead)} read / ${fmtN(A.counts.styleRows)} counted`) && a.title.includes(atsName(w.o.ats)),
        `${a.state}: ${a.text} | ${a.title}`);
      c('the card: report date from the file name, 2 days old, styles on more than one row summed', pair(s.cards.ats, 'Report date') === `${C.fmtDay(today - 2)} (from the file name)` &&
        pair(s.cards.ats, 'Age') === '2 days' && pair(s.cards.ats, 'Styles on more than one row (summed)') === '1', JSON.stringify(s.cards.ats && s.cards.ats.pairs));
      c('no email, no amber or red', sendMails(w.calls).length === 0 && !banner(s, 'err') && !banner(s, 'warn'), JSON.stringify(s.banners));
      c('downloaded once', w.downloads.filter(d => d.name === 'ats.xlsx').length === 1, w.downloads.map(d => d.name).join(','));
      const p = await sl(t);
      c("Available to sell: no note; the tiles are the engine's with the report", !p.note && SELL_TILE_KEYS.every(k => unitsIs(sellTile(p, k).n, S.totals[k])), `${p.note} · ${sellTilesSay(p, S)}`);
    },
  },
  {
    name: 'atsresaved', title: 'an older ATS report was opened and saved again after the newest: the newest by the date in its name is used', world: { ats: { age: 2, resaved: true } },
    async run({ t, w, c, today }) {
      const s = await t.snap(), a = s.chips.ats, name = atsName(w.o.ats);
      c('ATS chip green, naming the report dated 2 days ago (not the one saved last)', a.state === 'ok' && a.title.includes(name) && !a.title.includes(atsName(w.o.ats, 7)), `${a.state}: ${a.text} | ${a.title}`);
      c('only the newest report is opened', w.calls.some(x => x.path === `/drives/${SRC}/items/itm-ats-1`) && !w.calls.some(x => /itm-ats-0$/.test(x.path)),
        w.calls.filter(x => /itm-ats/.test(x.path)).map(x => x.path).join('; '));
      c('the card: report date from the file name, 2 days old', pair(s.cards.ats, 'Report date') === `${C.fmtDay(today - 2)} (from the file name)` && pair(s.cards.ats, 'Age') === '2 days',
        JSON.stringify(s.cards.ats && s.cards.ats.pairs));
      c('no email', sendMails(w.calls).length === 0);
    },
  },
  {
    name: 'atsfuture', title: "the ATS report's file name is dated 3 days ahead: amber, said plainly, no email", world: { ats: { age: -3 } },
    async run({ t, w, c, today }) {
      const s = await t.snap(), a = s.chips.ats;
      c('ATS chip amber "date in the future"', a.state === 'warn' && /date in the future/.test(a.text), `${a.state}: ${a.text}`);
      c("the chip's tooltip names the date and says to check it", a.title.includes(C.fmtDay(today + 3)) && /after today/.test(a.title), a.title);
      c('no email', sendMails(w.calls).length === 0);
      greenExcept(c, s, []);
    },
  },
  {
    name: 'atsamber', title: 'the newest ATS report is 9 days old, with no date in its name: amber, no email', world: { ats: { age: 9, noDate: true } },
    async run({ t, w, c, today }) {
      const s = await t.snap(), a = s.chips.ats;
      c('ATS chip amber "9 days old"', a.state === 'warn' && /9 days old/.test(a.text), `${a.state}: ${a.text}`);
      c("the card: the report date is the day it was saved (no date in the file name)", pair(s.cards.ats, 'Report date') === `${C.fmtDay(today - 9)} (the day it was saved: no date in the file name)`,
        JSON.stringify(s.cards.ats && s.cards.ats.pairs));
      c('no email before 10 days', sendMails(w.calls).length === 0);
      greenExcept(c, s, []);
    },
  },
  {
    name: 'atsnocol', title: 'the ATS report has no "Total Customer Orders" column', world: { ats: { age: 1, file: 'ats-nocol.xlsx' } },
    async run({ t, w, c, today }) {
      const s = await t.snap(), a = s.chips.ats;
      c('ATS chip red "missing column", naming "Total Customer Orders"', a.state === 'err' && /missing column/.test(a.text) && a.title.includes('"Total Customer Orders"'), `${a.state}: ${a.text} | ${a.title}`);
      greenExcept(c, s, []);
      oneEmail(c, s, w.calls, 'ats', 'missing column "Total Customer Orders"', today);
      const p = await sl(t);
      c('Available to sell: a red note, Available now "not loaded"', /\berr\b/.test(p.noteCls) && /The ATS report couldn't be used/.test(p.note) && sellTile(p, 'now').n === 'not loaded', `${p.noteCls}: ${p.note} · ${sellTile(p, 'now').n}`);
    },
  },
  {
    name: 'atsmissing', title: 'no ATS folder yet', world: { ats: null },
    async run({ t, w, c }) {
      const s = await t.snap(), a = s.chips.ats;
      c('ATS chip grey "no report yet"', a.state === 'off' && /no report yet/.test(a.text), `${a.state}: ${a.text}`);
      c('the chip says where to save it', /save it into Stock\/ATS on the TestSite site/.test(a.text) && a.title.includes('Save the weekly ATS report into Stock/ATS on the TestSite site.'), `${a.text} | ${a.title}`);
      c('the card says it too', s.cards.ats.text.includes('Save the weekly ATS report into Stock/ATS on the TestSite site.'), s.cards.ats.text);
      c('the folder was only looked for (a 404), never made', w.calls.some(x => x.path === `/drives/${SRC}/root:/Stock/ATS` && x.status === 404));
      c('no email, no red or amber', sendMails(w.calls).length === 0 && !banner(s, 'err') && !banner(s, 'warn'), JSON.stringify(s.banners));
      greenExcept(c, s, []);
      t.shots.push(await t.shot('atsmissing'));
    },
  },
  {
    name: 'noaccess', title: 'no access to the PO Status Tracker (403)', world: { tracker: 403 },
    async run({ t, w, c, today }) {
      const s = await t.snap();
      c('PO Status Tracker chip red "no access"', s.chips.tracker.state === 'err' && /no access/.test(s.chips.tracker.text), s.chips.tracker.text);
      c('Master Sheet Live grey "waiting" (read with the tracker)', s.chips.msl.state === 'off' && /waiting/.test(s.chips.msl.text), s.chips.msl.text);
      const red = banner(s, 'err');
      c('red banner lists the tracker and Master Sheet Live', /stage numbers are not loaded/i.test(red) && red.includes('PO Status Tracker') && red.includes('Master Sheet Live'), red);
      const p = await s2(t);
      c('every stage tile says "not loaded"', notLoadedEverywhere(p), JSON.stringify(liveTiles(p).map(x => x.n)));
      greenExcept(c, s, ['tracker', 'msl']);
      c('the chip links to the sharing link from config.js', s.chips.tracker.href === LOC.tracker, s.chips.tracker.href);
      const { msg } = oneEmail(c, s, w.calls, 'tracker', 'no access', today);
      const body = (msg.body && msg.body.content) || '';
      c('email written for Saul: whose account, what fixes it, what it holds up',
        body.includes(`${ACCOUNT.name}'s account can't open`) && /What fixes it: Give Test User access to it, or check its location in config\.js\./.test(body) &&
        /Until then: The stage numbers stay off/.test(body) && !/Your account/.test(body) && /^Shalam Tool: PO Status Tracker — no access$/.test(msg.subject), body);
      t.shots.push(await t.shot('noaccess'));
    },
  },
  {
    name: 'linkrefused', title: 'Graph refuses both sharing links (403), as it may for a read-only sign-in', world: { shares: 403 },
    async run({ t, w, c }) {
      const s = await t.snap();
      for (const k of ['msl', 'tracker']) {
        const x = s.chips[k];
        c(`${NAMES[k]} chip red "link refused", not "no access"`, x.state === 'err' && /link refused/.test(x.text) && !/no access/.test(x.text), `${x.state}: ${x.text}`);
        c(`${NAMES[k]}: says the read-only sign-in may be why, and that a /sites/… path avoids it`,
          x.title.includes("read-only sign-in isn't allowed to open sharing links") && x.title.includes('/sites/<site>/<library>/<path>'), x.title);
      }
      const red = banner(s, 'err');
      c('red banner names both', /stage numbers are not loaded/i.test(red) && red.includes('Master Sheet Live') && red.includes('PO Status Tracker'), red);
      greenExcept(c, s, ['msl', 'tracker']);
      const sends = sendMails(w.calls);
      c('one email each for the two', sends.length === 2 && ['Master Sheet Live', 'PO Status Tracker'].every(n => sends.some(x => mailText(x).subject === `Shalam Tool: ${n} — link refused`)),
        sends.map(x => mailText(x).subject).join(' | '));
      const body = (mailText(sends[0]).body || {}).content || '';
      c('the email gives the right fix (not "give access")', /What fixes it: Check Test User can open the file\. If they can, the read-only sign-in can't use sharing links: replace the link in config\.js with the file's \/sites\/<site>\/<library>\/<path> location\./.test(body) &&
        !/Give Test User access/.test(body) && !/Your account/.test(body), body);
      t.shots.push(await t.shot('linkrefused'));
    },
  },
  {
    name: 'atsbadlib', title: 'the ATS location in config.js names a library that does not exist', world: { sources: { ats: '/sites/TestSite/Shared Docs/Stock/ATS' } },
    async run({ t, w, c, today }) {
      const s = await t.snap(), a = s.chips.ats;
      c('ATS chip red "not found", naming the library, not "no report yet"', a.state === 'err' && /not found/.test(a.text) && !/no report yet/.test(a.text) &&
        a.title.includes('No library called "Shared Docs" on /sites/TestSite'), `${a.state}: ${a.text} | ${a.title}`);
      greenExcept(c, s, []);
      const { msg } = oneEmail(c, s, w.calls, 'ats', 'not found', today);
      const body = (msg.body && msg.body.content) || '';
      c("the email's fix is about its location, not the weekly report", /What fixes it: Put it back, or update its location in config\.js/.test(body), body);
    },
  },
  {
    name: 'atsbadsite', title: 'the ATS location in config.js names a site that does not exist', world: { sources: { ats: '/sites/TestSit/Shared Documents/Stock/ATS' } },
    async run({ t, w, c, today }) {
      const s = await t.snap(), a = s.chips.ats;
      c('ATS chip red "not found", naming the site, not "no report yet"', a.state === 'err' && /not found/.test(a.text) && !/no report yet/.test(a.text) &&
        a.title.includes("There's no site /sites/TestSit on SharePoint"), `${a.state}: ${a.text} | ${a.title}`);
      greenExcept(c, s, []);
      oneEmail(c, s, w.calls, 'ats', 'not found', today);
    },
  },
  {
    name: 'mailfail', title: 'the email itself fails (sendMail 403)', world: { ctFile: 'ct-bad.xlsx', mail: 403 },
    async run({ t, w, c, today }) {
      let s = await t.snap();
      c('Container Tracking chip red and says the email failed', s.chips.ct.state === 'err' && /email failed/.test(s.chips.ct.text), s.chips.ct.text);
      c('Problems row says emailing Saul failed', /Emailing Saul failed/.test(s.srows.ct), s.srows.ct);
      c('it blames the mailbox, not SharePoint', /Emailing Saul failed: this account's mailbox didn't allow sending \(403/.test(s.srows.ct) &&
        !/Emailing Saul failed: Your account can't open it on SharePoint/.test(s.srows.ct), s.srows.ct);
      c('one sendMail tried', sendMails(w.calls).length === 1, `${sendMails(w.calls).length}`);
      c('the browser mark is taken back, so a later visit can try again', !!s.store && !s.store[mailKey(today, 'ct')], JSON.stringify(s.store));
      const r = await runRefresh(t, w);
      c('the 5-minute refresh ran and did not retry (no new sendMail this page session)', r.fired === 1 && r.calls.length > 0 && sendMails(r.calls).length === 0, `${sendMails(r.calls).length} sent`);
      const from = w.calls.length;
      await t.reload();
      s = await t.snap();
      c('a later visit (reload) tries once more', sendMails(w.calls.slice(from)).length === 1, `${sendMails(w.calls.slice(from)).length} sent`);
    },
  },
  {
    name: 'pobook-changed', title: "the PO Book's rules changed since the copy", world: { poBook: 'changed' },
    async run({ t, w, c }) {
      const s = await t.snap(), p = s.chips.pobook;
      c('PO Book lines chip amber "PO Book rules changed"', p.state === 'warn' && /PO Book rules changed/.test(p.text), `${p.state}: ${p.text}`);
      c('no email for it', sendMails(w.calls).length === 0);
      c('no red banner', !banner(s, 'err'), banner(s, 'err'));
      greenExcept(c, s, ['pobook']);
    },
  },
  {
    name: 'blank', title: 'a blank location in config.js', world: { sources: { ct: '', drayage: '   ' } },
    async run({ t, w, c }) {
      const s = await t.snap();
      for (const k of ['ct', 'drayage']) {
        c(`${NAMES[k]} chip grey "not set"`, s.chips[k].state === 'off' && /not set/.test(s.chips[k].text) && /blank/.test(s.chips[k].title), `${s.chips[k].state}: ${s.chips[k].text} | ${s.chips[k].title}`);
      }
      c('nothing looked up for them', !w.calls.some(x => /Logistics|Drayage|itm-ct|itm-dray/.test(x.path)), w.calls.filter(x => /Logistics|Drayage|itm-ct|itm-dray/.test(x.path)).map(x => x.path).join('; '));
      c('no email', sendMails(w.calls).length === 0);
      c('info banner: stage numbers not loaded, Container Tracking blank in config.js', /stage numbers are not loaded: the location of Container Tracking is blank in config\.js/.test(banner(s, 'info')), banner(s, 'info'));
      const p = await s2(t);
      c('every stage tile says "not loaded"', notLoadedEverywhere(p), JSON.stringify(liveTiles(p).map(x => x.n)));
      c('no red banner', !banner(s, 'err'), banner(s, 'err'));
      c('health line counts the sources not set', /2 sources not set in config\.js/.test(s.health), s.health);
      greenExcept(c, s, ['ct', 'drayage']);
      t.shots.push(await t.shot('blank'));
    },
  },
  {
    name: 'trkblank', title: "the PO Status Tracker's location is blank in config.js", world: { sources: { tracker: '' } },
    async run({ t, w, c }) {
      const s = await t.snap();
      c('PO Status Tracker chip grey "not set"', s.chips.tracker.state === 'off' && /not set/.test(s.chips.tracker.text), `${s.chips.tracker.state}: ${s.chips.tracker.text}`);
      c('Master Sheet Live grey "waiting", saying the tracker is blank in config.js', s.chips.msl.state === 'off' && /waiting/.test(s.chips.msl.text) &&
        s.chips.msl.title.includes('whose location is blank in config.js'), `${s.chips.msl.state}: ${s.chips.msl.title}`);
      const red = banner(s, 'err');
      c('the red banner names the real cause: the tracker is blank in config.js', red.includes('Master Sheet Live') && red.includes('PO Status Tracker: its location is blank in config.js.'), red);
      c('no email', sendMails(w.calls).length === 0);
      greenExcept(c, s, ['msl', 'tracker']);
    },
  },
  {
    name: 'readonly', title: 'the sign-in carries Sites.ReadWrite.All and Files.ReadWrite.All', world: {},
    init: { scp: 'User.Read Sites.Read.All Sites.ReadWrite.All Files.ReadWrite.All Mail.Send' }, noGraph: true,
    async run({ t, w, c }) {
      const s = await t.snap();
      stopChecks(c, s, w, ['Shalam Tool must be read-only, but its sign-in still allows changes to SharePoint (Sites.ReadWrite.All, Files.ReadWrite.All). In Entra, remove that permission from the Shalam Tool app and grant consent again.']);
      c('a fresh token was asked for once before stopping (a cached one may predate the fix)', s.forced >= 1, `${s.tokens} tokens, ${s.forced} forced`);
      c('no 5-minute refresh waiting', (await t.eval('window.__headless.longWaiting()')) === 0);
      t.shots.push(await t.shot('readonly'));
    },
  },
  {
    name: 'manage', title: 'the sign-in carries Sites.Manage.All (and no Sites.Read.All)', world: {},
    init: { scp: 'User.Read Sites.Manage.All Mail.Send' }, noGraph: true,
    async run({ t, w, c }) {
      const s = await t.snap();
      stopChecks(c, s, w, ['still allows changes to SharePoint (Sites.Manage.All)']);
    },
  },
  {
    name: 'noread', title: 'the sign-in is missing Sites.Read.All', world: {},
    init: { scp: 'User.Read Mail.Send' }, noGraph: true,
    async run({ t, w, c }) {
      const s = await t.snap();
      stopChecks(c, s, w, ['missing the Sites.Read.All permission', 'User.Read, Mail.Send']);
    },
  },
  {
    name: 'selected', title: 'the sign-in carries write-capable permissions with no "Write" in their names', world: {},
    init: { scp: 'User.Read Sites.Read.All Sites.Selected Files.SelectedOperations.Selected Mail.Send.Shared Mail.Send' }, noGraph: true,
    async run({ t, w, c }) {
      const s = await t.snap();
      stopChecks(c, s, w, ['still allows changes to SharePoint (Sites.Selected, Files.SelectedOperations.Selected, Mail.Send.Shared)']);
    },
  },
  {
    name: 'opaque', title: "a token the page can't decode: the sign-in service's list of permissions is used", world: {}, init: { opaque: true },
    async run({ t, w, c }) {
      const s = await t.snap();
      c('no stop banner', !s.banners.some(b => b.id === 'stopBanner'), JSON.stringify(s.banners));
      greenExcept(c, s, []);
    },
  },
  {
    name: 'opaque-write', title: "a token the page can't decode, with Sites.ReadWrite.All in the sign-in service's list", world: {},
    init: { opaque: true, scp: 'User.Read Sites.Read.All Sites.ReadWrite.All Mail.Send' }, noGraph: true,
    async run({ t, w, c }) {
      const s = await t.snap();
      stopChecks(c, s, w, ['still allows changes to SharePoint (Sites.ReadWrite.All)']);
    },
  },
  {
    name: 'opaque-none', title: "a token the page can't decode and no list of permissions", world: {}, init: { opaque: true, scp: ' ' }, noGraph: true,
    async run({ t, w, c }) {
      const s = await t.snap();
      stopChecks(c, s, w, ["couldn't read which permissions its sign-in has"]);
    },
  },
  {
    name: 'readonly-later', title: 'a later token allows changes (Entra changed while the page was open)', world: { ctFile: 'ct-bad.xlsx' },
    async run({ t, w, c }) {
      let s = await t.snap();
      c('first load ran with a read-only sign-in', s.chips.msl.state === 'ok' && sendMails(w.calls).length === 1, s.chips.msl.state);
      const from = w.calls.length, dls = w.downloads.length;
      await t.eval(`window.__headless.scp = 'User.Read Sites.Read.All Sites.ReadWrite.All Mail.Send'`);
      await t.click('#btnRefresh');
      await t.waitFor(`!!document.getElementById('stopBanner')`, 'the stop banner', 15000);
      await t.waitFor(IDLE, 'the load to stop', 10000);
      s = await t.snap();
      const stop = s.banners.find(b => b.id === 'stopBanner');
      c('red stop banner naming Sites.ReadWrite.All', !!stop && stop.text.includes('allows changes to SharePoint (Sites.ReadWrite.All)'), JSON.stringify(s.banners));
      c('nothing requested with that token', w.calls.length === from && w.downloads.length === dls, w.calls.slice(from).map(x => `${x.method} ${x.path}`).join('; '));
      c('what was read before is no longer shown: every chip grey "not read"', KEYS.every(k => s.chips[k].state === 'off' && /not read/.test(s.chips[k].text)),
        KEYS.map(k => `${s.chips[k].state} ${s.chips[k].text}`).join(' | '));
      const r = await runRefresh(t, w, false);
      await sleep(500);
      c('no 5-minute refresh is left waiting, and nothing more is asked for', r.fired === 0 && w.calls.length === from, `fired ${r.fired}; ${r.calls.map(x => x.path).join('; ')}`);
    },
  },
  {
    name: 'cancel', title: 'Cancel during a slow download', world: { slowCt: true }, wait: 'busy',
    async run({ t, w, c }) {
      await t.waitFor(`(() => { const b = document.getElementById('busy'); return !!b && !b.hidden; })()`, 'the loading line', 20000);
      const end = Date.now() + 20000;
      while (!w.pending.size && Date.now() < end) await sleep(100);
      let s = await t.snap();
      c('a load in progress shows Cancel', s.busyShown && s.busy === 'true' && w.pending.size === 1, `busy ${s.busy}, pending ${[...w.pending]}`);
      await t.click('#btnCancelBusy');
      await t.waitFor(IDLE, 'the load to stop', 10000);
      await sleep(300);
      s = await t.snap();
      c('the download was stopped', w.aborted.includes('ct.xlsx'), `aborted: ${w.aborted.join(',') || 'none'}`);
      c('Container Tracking chip says cancelled', s.chips.ct.state === 'off' && /cancelled/.test(s.chips.ct.text), `${s.chips.ct.state}: ${s.chips.ct.text}`);
      c('no chip left loading, loading line gone', !s.loading.length && !s.busyShown, s.loading.join(','));
      c('page says the refresh was cancelled', /cancelled/i.test(s.toast), s.toast);
      const notReached = ['forwarder', 'drayage', 'ats', 'upc'].filter(k => /not set/.test(s.chips[k].text));
      c("sources not read before the Cancel don't claim to be \"not set\"", !notReached.length,
        notReached.map(k => `${NAMES[k]}: ${s.chips[k].state} "${s.chips[k].text}"`).join('; '));
      c('they say they were not read because of the Cancel', ['forwarder', 'drayage', 'ats', 'upc'].every(k => /not read \(cancelled\)/.test(s.chips[k].text)),
        ['forwarder', 'drayage', 'ats', 'upc'].map(k => s.chips[k].text).join(' | '));
      c("the toast doesn't claim older data is shown on a first load", !/loaded before|read before/i.test(s.toast), s.toast);
      c('the health line names what was not read', /Not read this time/.test(s.health) && s.health.includes('Container Tracking'), s.health);
      c('no email after a cancel', !sendMails(w.calls).length);
      t.shots.push(await t.shot('cancel'));
    },
  },
  {
    name: 'authfail', title: "a refresh where the sign-in can't be renewed quietly (timed_out)", world: {},
    async run({ t, w, c }) {
      await t.eval(`window.__headless.tokenFail = 'timed_out'`);
      await t.click('#btnRefresh');
      await t.waitFor(`/sign-in needs renewing/.test(document.getElementById('banners').textContent)`, 'the sign-in banner', 15000);
      await t.waitFor(IDLE, 'the load to stop', 10000);
      const s = await t.snap();
      const red = banner(s, 'err');
      c('red banner asks to sign in again, with a button', /sign-in needs renewing \(timed_out\)/.test(red) && /Sign in again/.test(red), red);
      const redChips = KEYS.filter(k => s.chips[k].state === 'err');
      c('no source turned red over a sign-in problem', !redChips.length, redChips.map(k => `${NAMES[k]}: ${s.chips[k].text}`).join('; '));
      c('the chips keep what was read before, marked with when', ['msl', 'ct', 'upc'].every(k => s.chips[k].state === 'ok' && /as read/.test(s.chips[k].text)),
        ['msl', 'ct', 'upc'].map(k => `${s.chips[k].state}: ${s.chips[k].text}`).join(' | '));
      c('no email tried', !sendMails(w.calls).length);
      t.shots.push(await t.shot('authfail'));
    },
  },

  // ---------------------------------------------------------------- step 2: Incoming
  {
    name: 's2-small', title: 'step 2, the standard invented files: tiles equal the engine, the balance line, Incoming issues on Problems', world: {},
    async run({ t, c, e }) {
      const R = e.R;
      await t.click('#tabs .tab[data-tab="inv"]'); await sleep(200);
      let s = await s2(t);
      c('the engine (run here) built the small files, balance 0', R.ok && R.balance.diff === 0 && R.balance.ct.diff === 0, JSON.stringify(R.ok ? R.balance : R.notLoaded));
      c("every tile equals the engine's total", tilesAre(s, R.totals.units), tilesSay(s, R.totals.units));
      const go = s.tiles.find(x => /\bgo\b/.test(x.cls));
      c('one grey tile instead of the step 3 ones: "On hand and available: Available to sell →"', !!go && /\boff\b/.test(go.cls) && go.lbl === 'On hand and available' && go.n === 'Available to sell →' &&
        !s.tiles.some(x => /step 3/.test(x.cs + x.n)), JSON.stringify(s.tiles.map(x => `${x.lbl}: ${x.n} ${x.cs}`)));
      c('no banner, no lead card', !s.banners.length && !s.lead, JSON.stringify(s.banners) + ' ' + s.lead);
      c('health line: all sources loaded, with the Unmatched units', /All sources loaded/.test(s.health) && /Unmatched/.test(s.health), s.health);
      c('the Inventory table draws its rows', s.invTr > 0, s.invTr);
      const oh = await t.eval(`[...document.querySelectorAll('#invTable tbody tr.r:not(.upc)')].map(r => r.children[4].textContent.trim())`);
      c('On hand says "no report yet" on every row (the ATS folder is empty), never 0 or blank', oh.length > 0 && oh.every(x => x === 'no report yet'), JSON.stringify(oh));
      await t.click('#tiles [data-go="sell"]'); await sleep(200);
      c('the grey tile opens Available to sell', (await t.snap()).tabs.sell);
      await t.click('#tabs .tab[data-tab="inv"]'); await sleep(200);
      await t.click('#tabs .tab[data-tab="problems"]');
      s = await s2(t);
      c('balance line green: "Every unit adds up: … units on Master Sheet Live = …"', /\bok\b/.test(s.balance.cls) &&
        s.balance.text.startsWith(`Every unit adds up: ${R.balance.mslUnits.toLocaleString('en-US')} units on Master Sheet Live =`), `${s.balance.cls}: ${s.balance.text}`);
      c(`the Incoming group comes first, with its ${R.issues.length} issues`, !!s.issGroups[0] && s.issGroups[0].startsWith(`Incoming · ${R.issues.length} issue`), s.issGroups.join(' | '));
      t.shots.push(await t.shot('s2-small-problems'));
      await t.click('#tabs .tab[data-tab="inv"]');
    },
  },
  {
    name: 's2-big', title: 'step 2, 2,400 styles: tiles, windowed rows, sort, filters, side panel, refresh, Pull from China, checked problems', world: { ...BIG },
    async run({ t, w, c, today }) {
      const R = stagesFor(w.fx, today, BIG_FILES);
      c('the engine (run here) built the big files, balance 0', R.ok && R.styles.length > 2000 && R.balance.diff === 0 && R.balance.ct.diff === 0,
        JSON.stringify({ ok: R.ok, styles: R.styles && R.styles.length, bal: R.balance }));
      await t.click('#tabs .tab[data-tab="inv"]'); await sleep(300);
      let s = await s2(t);
      c('no lead card: the stage numbers are loaded', !s.lead, s.lead);
      for (const k of TILE_KEYS) c(`tile ${k} = the engine's ${R.totals.units[k].toLocaleString('en-US')}`, unitsIs(tileN(s, k), R.totals.units[k]), tileN(s, k));
      const inc = s.tiles.find(x => x.k === 'incoming'), cIn = R.totals.cases.incoming, npIn = R.totals.noCasePackUnits.incoming;
      c(`Incoming tile: "${cIn} cases + ${npIn} units with no case pack"`, npIn > 0 && inc && inc.cs === `${cIn.toLocaleString('en-US')} cases + ${npIn.toLocaleString('en-US')} units with no case pack`, inc && inc.cs);
      c(`info line: ${R.styles.length} styles`, s.info.startsWith(`${R.styles.length.toLocaleString('en-US')} styles`), s.info);
      c('windowed: far fewer rows drawn than styles', s.invTr > 10 && s.invTr < 150, s.invTr);
      const want = ['Style', 'Description', 'Category', 'Factory', 'On hand', 'No CRD', 'In production', 'Ready', 'Booked', 'Loaded', 'In China', 'On the water', 'At port', 'Incoming', 'Sold, ships from China', 'Cases incoming', 'Next ETA'];
      c('the columns, in order', s.heads.length === want.length && want.every((h, i) => s.heads[i].t.startsWith(h)), s.heads.map(h => h.t).join('|'));
      const two = await t.eval(`[...document.querySelectorAll('#invTable tbody tr.upc')].map(r => r.textContent.trim().slice(0, 14))`);
      c('ST0001 (two POs, two UPCs) has its two UPC lines under it', two.length === 2 && two.some(x => x.startsWith('100000000001')) && two.some(x => x.startsWith('200000000001')), JSON.stringify(two));
      const lay = await t.eval(`(() => { const p = document.getElementById('invPane'); p.scrollLeft = 600; p.dispatchEvent(new Event('scroll'));
        const td = document.querySelector('#invTable tbody tr.r td.fz'), th = document.querySelector('#invTable thead th');
        const r = { wide: p.scrollWidth > p.clientWidth, left: Math.round(td.getBoundingClientRect().left - p.getBoundingClientRect().left - p.clientLeft),
          th: getComputedStyle(th).position, page: document.documentElement.scrollWidth <= window.innerWidth + 1 };
        p.scrollLeft = 0; p.dispatchEvent(new Event('scroll')); return r; })()`);
      c('the table scrolls in its own pane: Style frozen, header sticky, no page-wide sideways scroll', lay.th === 'sticky' && lay.page && (!lay.wide || Math.abs(lay.left) <= 2), JSON.stringify(lay));
      const fit = await t.eval(`(() => { const p = document.getElementById('invPane').getBoundingClientRect(); return { bottom: Math.round(p.bottom), h: innerHeight }; })()`);
      c('the Inventory pane ends inside the window, so its own scroll bars are on screen', fit.bottom <= fit.h && fit.bottom > fit.h - 40, JSON.stringify(fit));
      const rowsH = await t.eval(`(() => { const r = [...document.querySelectorAll('#invTable tbody tr.r')], h = x => x.getBoundingClientRect().height.toFixed(1);
        return { style: [...new Set(r.filter(x => !x.classList.contains('upc')).map(h))], upc: [...new Set(r.filter(x => x.classList.contains('upc')).map(h))] }; })()`);
      c('UPC lines are drawn the same height as style rows (the windowed rows keep one height)', rowsH.upc.length >= 1 && rowsH.style.length === 1 && rowsH.upc.every(x => x === rowsH.style[0]), JSON.stringify(rowsH));
      t.shots.push(await t.shot('s2-big'));
      await scrollTo(t, 'invPane', 40000); await sleep(500);
      s = await s2(t);
      c('scrolling draws later rows, still windowed', s.invFirst > 600 && s.invTr < 150, `first ${s.invFirst}, ${s.invTr} rows`);
      await t.click('#invTable th[data-s="incoming"]'); await sleep(200);
      s = await s2(t);
      const hi = s.heads.find(h => h.t.startsWith('Incoming'));
      c('header sort: Incoming, biggest first, back to the top', hi && hi.sort === 'descending' && s.invFirst === 0 && s.scrollInv === 0, JSON.stringify(hi) + ' ' + s.invFirst);
      await t.click('#invTable th[data-s="style"]'); await sleep(100);

      // Filters change the rows and the tiles
      const one = R.styles.find(x => x.style === 'ST0012');
      await setVal(t, 'fq', 'ST0012'); await sleep(500);
      s = await s2(t);
      c('search: one style', s.info.startsWith('1 style') && s.invTr === 1, s.info);
      c("search: the tiles are that style's numbers, marked filtered", tilesAre(s, one.units) && s.tiles.filter(x => /filtered/.test(x.lbl)).length === 6, tilesSay(s, one.units));
      await setVal(t, 'fq', '100000000013'); await sleep(500);
      s = await s2(t);
      c('search by UPC finds its style', s.invTr === 1 && /ST0013/.test(await t.eval(`document.querySelector('#invTable tbody tr.r').textContent`)), s.info);
      await setVal(t, 'fq', 'UN-1'); await sleep(500);
      const tag = await t.eval(`(() => { const td = document.querySelector('#invTable tbody tr.r td.fz'), g = td && td.querySelector('.tag'); if (!g) return null;
        const a = td.getBoundingClientRect(), b = g.getBoundingClientRect(); return { inside: b.left >= a.left && b.right <= a.right + 0.5, w: Math.round(b.width), text: g.textContent.trim(), upc: (td.querySelector('.u') || {}).textContent || '', title: td.title }; })()`);
      c('the "500 unmatched" tag shows whole after a 13-digit UPC in the 200px Style cell, and the cell tooltip says it', !!tag && tag.inside && tag.w > 20 && tag.text === '500 unmatched' && tag.upc === '1234567890128' && /500 units unmatched/.test(tag.title), JSON.stringify(tag));
      await setVal(t, 'fq', 'ST07'); await sleep(500);   // sailed on round 361: numbers and a Next ETA on every row
      const fmt = await t.eval(`(() => { const rows = [...document.querySelectorAll('#invTable tbody tr.r')];
        const txt = el => { const x = el.cloneNode(true); x.querySelectorAll('.np').forEach(y => y.remove()); return x.textContent.trim(); };
        const nums = rows.flatMap(r => [...r.querySelectorAll('td.num:not(.off):not(.nr)')].map(txt)).filter(Boolean), eta = rows.map(r => r.lastElementChild.textContent.trim()).filter(Boolean);
        const oh = rows.filter(r => !r.classList.contains('upc')).map(r => r.children[4].textContent.trim());
        return { rows: rows.length, n: nums.length, bad: nums.filter(x => !/^[1-9]\\d{0,2}(,\\d{3})*$/.test(x)).slice(0, 5), comma: nums.filter(x => x.includes(',')).length,
          eta: eta.length, badEta: eta.filter(x => !/^\\d{1,2}\\/\\d{1,2}\\/\\d{4}$/.test(x)).slice(0, 5), onHand: [...new Set(oh)] }; })()`);
      c('numbers have thousands separators and a blank instead of 0; Next ETA is m/d/yyyy; On hand says "no report yet" (no ATS report)',
        fmt.rows > 0 && fmt.n > 0 && !fmt.bad.length && fmt.comma > 0 && fmt.eta === fmt.rows && !fmt.badEta.length && fmt.onHand.join() === 'no report yet', JSON.stringify(fmt));
      await setVal(t, 'fq', ''); await sleep(400);
      await setVal(t, 'fFac', 'Factory B'); await sleep(300);
      s = await s2(t);
      const gB = tileSums(R, a => a.factory === 'Factory B');
      c("factory filter: the tiles are Factory B's allocations", tilesAre(s, gB), tilesSay(s, gB));
      await setVal(t, 'fFac', ''); await setVal(t, 'fType', 'China Pickup'); await sleep(300);
      s = await s2(t);
      const gC = tileSums(R, a => a.ptype === 'China Pickup');
      c('PO type China Pickup: only Pull from China', tilesAre(s, gC) && gC.pullFromChina > 0 && gC.inChina === 0, tilesSay(s, gC));
      await setVal(t, 'fType', ''); await setVal(t, 'fCat', 'Cat 2'); await sleep(300);
      s = await s2(t);
      const nCat = R.styles.filter(x => x.category === 'Cat 2' && TILE_KEYS.some(k => x.units[k])).length, gCat = tileSums(R, a => (R.styles.find(x => x.style === a.style) || {}).category === 'Cat 2');
      c(`category Cat 2: ${nCat} styles, tiles their allocations`, s.info.startsWith(`${nCat.toLocaleString('en-US')} style`) && tilesAre(s, gCat), `${s.info} · ${tilesSay(s, gCat)}`);
      await setVal(t, 'fCat', ''); await setVal(t, 'fStage', 'booked'); await sleep(300);
      s = await s2(t);
      const nBooked = R.styles.filter(x => x.units.booked > 0).length;
      c(`stage Booked: ${nBooked} styles, In China = the Booked total`, s.info.startsWith(`${nBooked.toLocaleString('en-US')} style`) && unitsIs(tileN(s, 'inChina'), R.totals.units.booked) && unitsIs(tileN(s, 'water'), 0), `${s.info} · ${tileN(s, 'inChina')}`);
      await t.click('#btnClearF'); await sleep(300);
      s = await s2(t);
      c('Clear filters: every style again', s.info.startsWith(`${R.styles.length.toLocaleString('en-US')} styles`) && !s.fq && tilesAre(s, R.totals.units), s.info);

      // The side panel
      const nb = await t.eval(`(() => { const b = document.querySelector('#invTable tbody tr.r button.nb'); b.click(); return b.textContent.trim(); })()`);
      await sleep(200);
      s = await s2(t);
      c('a number opens the side panel with its lines, adding up to the number', s.panel && s.panelTr >= 1 && /ST0000/.test(s.panelTitle) && s.panelSub.includes(`${nb} units counted`), `${s.panel} ${s.panelTitle} · ${s.panelSub} (clicked ${nb})`);
      c('the panel has the columns', ['UPC', 'PO', 'PO in the PO Book', 'Customer', 'Counts in', 'Units', 'Cases', 'CRD', 'Round', 'Containers', 'ETD', 'ETA', 'Stage', 'Flags', 'From'].every(h => s.panelHead.includes(h)), s.panelHead.join('|'));
      c('the panel says which file each fact came from', /Master Sheet Live/.test(s.panelFirst), s.panelFirst);
      // Each fact's file under its value, in plain sight (not only in a tooltip): Units, CRD and Stage on the first row.
      const under = await t.eval(`(() => { const trs = [...document.querySelectorAll('#panelTable tbody tr.r')], hs = [...document.querySelectorAll('#panelTable thead th')].map(th => th.textContent.trim());
        const cell = (tr, h) => tr.children[hs.findIndex(x => x.startsWith(h))];
        const val = (tr, h) => { const td = cell(tr, h), x = td && td.cloneNode(true); if (!x) return null; x.querySelectorAll('.u').forEach(y => y.remove()); return x.textContent.trim(); };
        const u = (tr, h) => { const td = cell(tr, h), s = td && td.querySelector('.u'); return s ? s.textContent.trim() : null; };
        const crd = trs.map(tr => [val(tr, 'CRD'), u(tr, 'CRD')]);
        return { units: u(trs[0], 'Units'), stage: u(trs[0], 'Stage'), crdShown: crd.filter(x => x[0]).length, crdOk: crd.every(([v, s]) => (v ? /PO Status Tracker|Master Sheet Live|PO Book/.test(s) : !s)),
          oneHeight: new Set(trs.map(r => Math.round(r.getBoundingClientRect().height))).size === 1 }; })()`);
      c('each fact shows the file it came from under it (Units, Stage; every CRD shown); every panel row one height', under.units === 'Master Sheet Live' && !!under.stage && under.crdOk && under.oneHeight, JSON.stringify(under));
      c('the panel is a side panel, not a popup', !(await t.eval(`!!document.querySelector('dialog[open]')`)));
      t.shots.push(await t.shot('s2-big-panel'));
      await key(t, 'Escape'); await sleep(100);
      c('Esc closes the panel', !(await s2(t)).panel);
      await t.click('#tiles button.tile[data-t="unmatched"]'); await sleep(200);
      s = await s2(t);
      c('the Unmatched tile lists UN-1 (500 units), "unmatched", never "counted"', s.panel && s.panelTr === 1 && /UN-1/.test(s.panelFirst) && /500/.test(s.panelFirst) &&
        s.panelSub.includes(`${R.totals.units.unmatched.toLocaleString('en-US')} units unmatched`) && !/units counted/.test(s.panelSub), `${s.panelTitle} · ${s.panelSub} · ${s.panelFirst}`);
      await t.click('#btnPanelClose'); await sleep(100);
      c('the Close button closes it', !(await s2(t)).panel);
      await t.click('#tiles button.tile[data-t="inChina"]'); await sleep(200);
      s = await s2(t);
      c('the In China tile lists its lines, adding up to the tile', s.panel && s.panelTr > 0 && s.panelSub.includes(`${R.totals.units.inChina.toLocaleString('en-US')} units counted`), s.panelSub);
      await t.click('#btnPanelX'); await sleep(100);
      c('✕ closes it', !(await s2(t)).panel);
      await t.click('#invTable tbody tr.r button.sb'); await sleep(100);
      s = await s2(t);
      c('a style opens every line behind it', s.panel && /every line/.test(s.panelTitle), s.panelTitle);
      // the 5-minute refresh waits while the panel is open
      let from = w.calls.length;
      const fired = await t.eval('window.__headless.fireLong()');
      await sleep(1500);
      c('the 5-minute refresh waits while the panel is open', fired >= 1 && w.calls.length === from, `fired ${fired}, ${w.calls.length - from} calls`);
      await t.click('#btnPanelX');
      const end = Date.now() + 15000;
      while (w.calls.length === from && Date.now() < end) await sleep(100);
      await t.waitFor(IDLE, 'the refresh to finish', 30000);
      c('closing the panel runs the refresh that waited', w.calls.length > from, `${w.calls.length - from} calls`);

      // A refresh keeps the filters, the sort, the place and the open panel
      await setVal(t, 'fq', 'ST01'); await sleep(400);
      await t.click('#invTable th[data-s="incoming"]'); await sleep(100);
      await scrollTo(t, 'invPane', 600); await sleep(300);
      const before = await s2(t);
      await t.click('#invTable tbody tr.r button.nb'); await sleep(100);
      const pBefore = (await s2(t)).panelTitle;
      from = w.calls.length;
      await t.click('#btnRefresh');
      const end2 = Date.now() + 15000;
      while (w.calls.length === from && Date.now() < end2) await sleep(100);
      await t.waitFor(IDLE, 'the refresh to finish', 30000);
      s = await s2(t);
      c('Refresh keeps the search, the sort and the place', s.fq === 'ST01' && s.info === before.info && Math.abs(s.scrollInv - before.scrollInv) < 2 && s.heads.find(h => h.t.startsWith('Incoming')).sort === 'descending',
        `${s.fq} · ${s.info} vs ${before.info} · scroll ${s.scrollInv} vs ${before.scrollInv}`);
      c("…and the open panel's selection", s.panel && s.panelTitle === pBefore && s.panelTr >= 1, `${s.panelTitle} vs ${pBefore}`);
      await key(t, 'Escape');
      await t.click('#btnClearF'); await sleep(200);
      await t.click('#invTable th[data-s="style"]'); await sleep(100);

      // Pull from China
      await t.click('#tabs .tab[data-tab="pull"]'); await sleep(300);
      s = await s2(t);
      c('Pull from China: the filters move there, rows windowed', s.fbarIn === 'tab-pull' && s.pullTr > 10 && s.pullTr < 150 && !s.pullLead, `${s.fbarIn} ${s.pullTr}`);
      c(`Pull from China: ${R.pull.length} open lines`, s.cPull === R.pull.length.toLocaleString('en-US') && s.info.startsWith(`${R.pull.length.toLocaleString('en-US')} open lines`), `${s.cPull} · ${s.info}`);
      await setVal(t, 'fType', 'China Pickup'); await sleep(300);
      s = await s2(t);
      const nCp = R.pull.filter(r => r.ptype === 'China Pickup').length;
      c(`Pull from China, PO type China Pickup: ${nCp} lines`, nCp > 0 && s.cPull === nCp.toLocaleString('en-US'), s.cPull);
      await t.click('#pullTable tbody tr.r button.nb'); await sleep(200);
      s = await s2(t);
      c("a line's units open its allocations", s.panel && s.panelTr >= 1 && /^PO 9002/.test(s.panelTitle), `${s.panelTitle} ${s.panelTr}`);
      t.shots.push(await t.shot('s2-big-pull'));
      await key(t, 'Escape');
      await t.click('#btnClearF');

      // Container Tracking sections with no round, by the test config.js's invented names
      const uOf = p => R.allocs.filter(p).reduce((n, a) => n + a.units, 0);
      const secR = { pull: uOf(a => a.section === 'Customer Direct' && a.bucket === 'pullFromChina'), more: uOf(a => a.section === 'Customer Direct' && a.bucket === 'ctMore'),
        planning: uOf(a => a.bucket === 'planning'), ignored: uOf(a => a.bucket === 'ignored'), below: uOf(a => a.bucket === 'belowTracker'),
        odd: uOf(a => a.section === 'Odd Section' && a.bucket === 'loaded'), notListed: R.issues.filter(i => i.code === 'section-not-listed').length };
      c('the engine (run here) uses the test sections: Customer Direct to Pull from China, Planning A planned, Hidden Rows Test ignored, Odd Section Loaded and listed, PO 900150 below the tracker',
        JSON.stringify(secR) === JSON.stringify({ pull: 6000, more: 300, planning: 6000, ignored: 1440, below: 700, odd: 1200, notListed: 1 }), JSON.stringify(secR));
      const stCell = () => t.eval(`(() => { const td = document.querySelector('#pullTable tbody tr.r td.st'); return td ? { title: td.title, cls: td.className } : null; })()`);
      await setVal(t, 'fq', 'ST1010'); await sleep(400);
      let stc = await stCell();
      c('Pull from China: a line under Customer Direct says "Booked · Customer Direct"', !!stc && stc.title === 'Booked · Customer Direct', JSON.stringify(stc));
      await setVal(t, 'fq', 'ST1020'); await sleep(400);
      stc = await stCell();
      c('…a line only under Planning A says "Planned · Planning A", in grey', !!stc && stc.title === 'Planned · Planning A' && /\bnt\b/.test(stc.cls), JSON.stringify(stc));
      await setVal(t, 'fq', ''); await setVal(t, 'fStage', 'pullFromChina'); await sleep(300);
      s = await s2(t);
      const nPull = R.pull.filter(r => r.ptype === 'China Pickup' || r.flags.includes('pull-section')).length;
      c(`Pull from China, stage Pull from China: the ${nPull} lines of China Pickup POs and Customer Direct`, nPull > 0 && s.cPull === nPull.toLocaleString('en-US'), `${s.cPull} vs ${nPull}`);
      await t.click('#btnClearF'); await sleep(200);
      await t.click('#tabs .tab[data-tab="inv"]'); await sleep(200);
      await setVal(t, 'fq', 'ST1010'); await sleep(400);
      await t.eval(`document.querySelector('#invTable tbody tr.r button.sb').click()`); await sleep(200);
      const pr = await t.eval(`(() => { const hs = [...document.querySelectorAll('#panelTable thead th')].map(th => th.textContent.trim()), tr = document.querySelector('#panelTable tbody tr.r');
        if (!tr) return null; const cell = h => { const x = tr.children[hs.findIndex(y => y.startsWith(h))].cloneNode(true); x.querySelectorAll('.u').forEach(y => y.remove()); return x.textContent.trim(); };
        return { counts: cell('Counts in'), stage: cell('Stage'), flags: cell('Flags') }; })()`);
      c('the side panel: counts in "Sold, ships from China", Stage "Booked · Customer Direct", flag "Straight to the customer from China"',
        !!pr && pr.counts === 'Sold, ships from China' && pr.stage === 'Booked · Customer Direct' && /Straight to the customer from China/.test(pr.flags), JSON.stringify(pr));
      await key(t, 'Escape');
      await setVal(t, 'fq', 'ST1020'); await sleep(400);
      s = await s2(t);
      const r1020 = R.styles.find(x => x.style === 'ST1020');
      c('a style only under Planning A keeps its Master Sheet stage (nothing Booked)', !!r1020 && r1020.units.booked === 0 && r1020.units.inChina === 1200 && tilesAre(s, r1020.units), tilesSay(s, r1020 && r1020.units));
      await setVal(t, 'fq', ''); await sleep(300);
      // Its Pull from China line's side panel names the planning section with the Master Sheet stage.
      await t.click('#tabs .tab[data-tab="pull"]'); await sleep(300);
      await setVal(t, 'fq', 'ST1020'); await sleep(400);
      await t.click('#pullTable tbody tr.r button.nb'); await sleep(200);
      const pp = await t.eval(`(() => { const hs = [...document.querySelectorAll('#panelTable thead th')].map(th => th.textContent.trim()), tr = document.querySelector('#panelTable tbody tr.r');
        if (!tr) return null; const cell = h => { const x = tr.children[hs.findIndex(y => y.startsWith(h))].cloneNode(true); x.querySelectorAll('.u').forEach(y => y.remove()); return x.textContent.trim(); };
        return { stage: cell('Stage'), flags: cell('Flags') }; })()`);
      c('…its Pull from China line\'s side panel: Stage "<its CRD stage> · Planning A", flag "Planned, not booked yet"',
        !!pp && /^(No CRD|In production|Ready) · Planning A$/.test(pp.stage) && /Planned, not booked yet/.test(pp.flags), JSON.stringify(pp));
      await key(t, 'Escape');
      await setVal(t, 'fq', ''); await sleep(300);
      await t.click('#tabs .tab[data-tab="inv"]'); await sleep(200);
      // The stage filter on an outcome outside the totals: the style's tag keeps proper names' capitals.
      await setVal(t, 'fStage', 'belowTracker'); await setVal(t, 'fq', 'OLD-1'); await sleep(400);
      const tagOf = () => t.eval(`(() => { const g = document.querySelector('#invTable tbody tr.r td.fz .tag'); return g ? { text: g.textContent.trim(), title: g.title } : null; })()`);
      let tg = await tagOf();
      c('stage filter "Older than the PO Status Tracker": the tag reads "700 older than the PO Status Tracker"', !!tg && tg.text === '700 older than the PO Status Tracker' &&
        tg.title === '700 units older than the PO Status Tracker: not counted in any column', JSON.stringify(tg));
      await setVal(t, 'fStage', 'ignored'); await setVal(t, 'fq', 'ST1030'); await sleep(400);
      tg = await tagOf();
      c('…and "Ignored section": "1,200 in an ignored section"', !!tg && tg.text === '1,200 in an ignored section', JSON.stringify(tg));
      await t.click('#btnClearF'); await sleep(200);
      await t.click('#tabs .tab[data-tab="problems"]'); await sleep(200);
      const pt = await t.eval(`document.getElementById('issTable').textContent`);
      c('Problems: one line each for the section not on the list, the ignored section and the PO below the tracker, in plain words',
        pt.includes("Section 'Odd Section' on Loaded has no round and isn't in config.js's section list") && pt.includes('Section not in the list') &&
        pt.includes('Ignored section') && pt.includes('"Hidden Rows Test" section is ignored') && pt.includes('Older than the PO Status Tracker') &&
        R.issues.filter(i => ['section-not-listed', 'ignored-section', 'below-tracker'].includes(i.code)).length === 3, pt.slice(0, 200));
      const bal = (await s2(t)).balance.text;
      c('the balance line names the planned, ignored and older Container Tracking units apart (the planned ones aren\'t on Problems: it says where they are)',
        /listed below\): /.test(bal) && /6,000 units planned, not booked yet \(config\.js; not listed below: they show in each style's side panel\)/.test(bal) && /1,440 units in an ignored section/.test(bal) &&
        /700 units on POs below the PO Status Tracker's lowest PO/.test(bal), bal);
      await t.click('#tabs .tab[data-tab="inv"]'); await sleep(200);

      // Problems: the balance line, checked marks, new since the last load
      await t.click('#tabs .tab[data-tab="problems"]'); await sleep(200);
      s = await s2(t);
      c('balance line green', /\bok\b/.test(s.balance.cls) && s.balance.text.startsWith(`Every unit adds up: ${R.balance.mslUnits.toLocaleString('en-US')} units on Master Sheet Live`), s.balance.text);
      c('Incoming issues first', !!s.issGroups[0] && s.issGroups[0].startsWith('Incoming'), s.issGroups.join(' | '));
      const prob0 = s.cProb;
      const firstMsg = await t.eval(`document.querySelector('#issTable input.ck').closest('tr').querySelector('td.msg').textContent`);
      const has = () => t.eval(`[...document.querySelectorAll('#issTable td.msg')].some(td => td.textContent === ${JSON.stringify(firstMsg)})`);
      await t.click('#issTable input.ck'); await sleep(200);
      s = await s2(t);
      const ck = Object.keys((await t.snap()).store || {}).filter(k => k.startsWith('inv:checked:'));
      c('a checked problem goes away, kept in this browser as inv:checked:<source>:<code>:<po>:<style>:<row>', !(await has()) && ck.length === 1 && /^inv:checked:Incoming:[\w-]+:[^:]*:[^:]*:[^:]*$/.test(ck[0]), `${firstMsg.slice(0, 60)}; ${ck.join(',')}`);
      c('the Problems count drops by one', +s.cProb.replace(/,/g, '') === +prob0.replace(/,/g, '') - 1, `${prob0} -> ${s.cProb}`);
      await t.reload();
      await t.click('#tabs .tab[data-tab="problems"]'); await sleep(200);
      s = await s2(t);
      c('still checked after a reload', !(await has()) && /1 checked, hidden/.test(s.pqInfo), `${s.issRows} · ${s.pqInfo}`);
      await t.click('#pqChecked'); await sleep(100);
      c('"Show checked" brings it back', await has());
      await t.click('#pqChecked'); await sleep(100);
      await t.eval(`(() => { const k = Object.keys(localStorage).find(x => x.startsWith('inv:checked:')); localStorage.setItem(k, 'an older message'); })()`);
      await t.reload();
      await t.click('#tabs .tab[data-tab="problems"]'); await sleep(200);
      c('a checked problem whose message changed comes back', await has(), (await s2(t)).pqInfo);
      const seen0 = await t.eval(`localStorage.getItem('inv:seen')`);
      c("the issues last seen on Problems are kept in this browser (inv:seen)", !!seen0 && !!JSON.parse(seen0).Incoming && JSON.parse(seen0).Incoming.length > 0, (seen0 || '').slice(0, 80));
      // Leaving Problems (or the page while on it) marks what it showed as seen, so the stored list is edited after.
      await t.click('#tabs .tab[data-tab="inv"]'); await sleep(100);
      await t.eval(`(() => { const j = JSON.parse(localStorage.getItem('inv:seen')); j.Incoming = j.Incoming.slice(1); localStorage.setItem('inv:seen', JSON.stringify(j)); })()`);
      await t.reload();
      // The 5-minute refresh runs before anyone opens Problems: the "new" mark must survive it.
      const rf = await runRefresh(t, w);
      await t.click('#tabs .tab[data-tab="problems"]'); await sleep(200);
      const nw = await t.eval(`(() => { const rows = [...document.querySelectorAll('#issTable tbody tr')]; const g = rows.findIndex(r => r.classList.contains('gh'));
        return { n: document.querySelectorAll('#issTable tr.isnew').length, firstNew: !!(rows[g + 1] && rows[g + 1].classList.contains('isnew')), head: rows[g] ? rows[g].textContent : '' }; })()`);
      c('an issue new since Problems was last looked at is marked, first in its group, counted in the group line, after a 5-minute refresh', rf.fired >= 1 && nw.n === 1 && nw.firstNew && /1 new/.test(nw.head), JSON.stringify({ ...nw, fired: rf.fired }));
      await t.click('#tabs .tab[data-tab="inv"]'); await sleep(100);
      const seen1 = await t.eval(`localStorage.getItem('inv:seen')`);
      c('leaving Problems marks what it showed as seen', !!seen1 && JSON.parse(seen1).Incoming.length === JSON.parse(seen0).Incoming.length, `${JSON.parse(seen1 || '{}').Incoming?.length} vs ${JSON.parse(seen0).Incoming.length}`);
    },
  },
  {
    name: 's2-ctbad', title: 'step 2: Container Tracking fails, every stage number says "not loaded"', world: { ...BIG, ctFile: 'ct-bad.xlsx' },
    async run({ t, c }) {
      const p = await sl(t);
      c('Available to sell says "not loaded" everywhere: a lead card saying why, every tile, no rows', p.shown && /Available to sell is not loaded/.test(p.lead) && /Container Tracking \(missing column\)/.test(p.lead) &&
        p.tiles.length === 5 && p.tiles.every(x => x.n === 'not loaded' && !/\d/.test(x.cs.join(''))) && p.tr === 0 && /^Not loaded/.test(p.empty), `${p.lead} · ${JSON.stringify(p.tiles.map(x => x.n))} · ${p.tr} · ${p.empty}`);
      await t.click('#tabs .tab[data-tab="inv"]'); await sleep(200);
      let s = await s2(t);
      c('every stage tile says "not loaded", with no partial number', notLoadedEverywhere(s), JSON.stringify(liveTiles(s).map(x => `${x.n} ${x.cs}`)));
      c('the lead card says why', /Stage numbers not loaded/.test(s.lead) && /Container Tracking \(missing column\)/.test(s.lead), s.lead);
      c('the table says not loaded, no rows', s.invTr === 0 && /^Not loaded/.test(s.invEmpty), s.invEmpty);
      t.shots.push(await t.shot('s2-ctbad'));
      await t.click('#tabs .tab[data-tab="pull"]'); await sleep(200);
      s = await s2(t);
      c('Pull from China says not loaded', s.pullTr === 0 && /^Not loaded/.test(s.pullEmpty) && /Not loaded/.test(s.pullLead), `${s.pullEmpty} | ${s.pullLead}`);
      await t.click('#tabs .tab[data-tab="problems"]'); await sleep(200);
      s = await s2(t);
      c('balance line: not loaded', /\boff\b/.test(s.balance.cls) && /not loaded/.test(s.balance.text), s.balance.text);
      c('no Incoming issues group', !s.issGroups.some(g => g.startsWith('Incoming')), s.issGroups.join(' | '));
      await t.click('#tabs .tab[data-tab="inv"]');
    },
  },
  {
    name: 's2-fwdold', title: 'step 2: the forwarder report is 9 days old, stages from Container Tracking dates only', world: { ...BIG, fwdAge: 9 },
    async run({ t, w, c, today }) {
      const R = stagesFor(w.fx, today, { ...BIG_FILES, fwd: null });
      const s = await s2(t);
      const b = s.banners.find(x => x.id === 'reportsBanner');
      c(`amber banner: the Forwarder report isn't used, what that hides, and the ${R.counts.trackerOnly} PO entries no report covers`, !!b && /\bwarn\b/.test(b.cls) && /Forwarder report not used:/.test(b.text) &&
        /rolled vessels can't be seen/.test(b.text) && !/Drayage/.test(b.text.split('Forwarder report:')[0]) &&
        b.text.includes(`${R.counts.trackerOnly.toLocaleString('en-US')} PO ${R.counts.trackerOnly === 1 ? 'entry' : 'entries'} on Booking and Loaded ${R.counts.trackerOnly === 1 ? 'has' : 'have'} no report covering`), b && b.text);
      c('numbers still shown, equal to the engine without the forwarder', tilesAre(s, R.totals.units), tilesSay(s, R.totals.units));
      t.shots.push(await t.shot('s2-fwdold'));
    },
  },
  {
    name: 's2-refresh', title: 'step 2: a source that fails on a refresh is never used from the last read (drayage, then Container Tracking, then both back)',
    world: { drayFile: 'dray-port.csv', drayOldFile: 'dray-port.csv' },
    async run({ t, w, c, today }) {
      const R = stagesFor(w.fx, today, { dray: 'dray-port.csv' }), R0 = stagesFor(w.fx, today, { dray: null });
      c('the invented files exercise it: the drayage report moves units from On the water to At port', R.ok && R0.ok && R.totals.units.port > 0 &&
        R0.totals.units.port < R.totals.units.port && R0.totals.units.water > R.totals.units.water, JSON.stringify([R.totals && R.totals.units, R0.totals && R0.totals.units]));
      let s = await s2(t);
      c('first load: the tiles equal the engine with drayage', tilesAre(s, R.totals.units), tilesSay(s, R.totals.units));
      c('no shipping-reports banner', !s.banners.some(b => b.id === 'reportsBanner'), JSON.stringify(s.banners));
      // Both drayage files lose their "Availability" column.
      w.o.drayFile = w.o.drayOldFile = 'dray-bad.csv';
      let r = await runRefresh(t, w);
      s = await s2(t);
      let snap = await t.snap();
      c('the refresh ran; the Drayage chip is red', r.fired === 1 && snap.chips.drayage.state === 'err', `fired ${r.fired}; ${snap.chips.drayage.state}: ${snap.chips.drayage.text}`);
      c("the tiles equal the engine WITHOUT drayage: the last read's copy isn't used", tilesAre(s, R0.totals.units), tilesSay(s, R0.totals.units));
      const b = s.banners.find(x => x.id === 'reportsBanner');
      c(`amber banner: the Drayage report isn't used (deliveries can't be seen), ${R0.counts.trackerOnly} PO entries no report covers`, !!b && /\bwarn\b/.test(b.cls) &&
        /Drayage report not used:/.test(b.text) && /pickups and deliveries can't be seen/.test(b.text) && !/Container Tracking dates only/.test(b.text) &&
        b.text.includes(`${R0.counts.trackerOnly.toLocaleString('en-US')} PO ${R0.counts.trackerOnly === 1 ? 'entry' : 'entries'}`), b && b.text);
      // Container Tracking fails as well: nothing partial is left on the page.
      w.o.ctFile = 'ct-bad.xlsx';
      r = await runRefresh(t, w);
      s = await s2(t);
      c('Container Tracking fails on a refresh: every tile says "not loaded"', r.fired === 1 && notLoadedEverywhere(s), JSON.stringify(liveTiles(s).map(x => `${x.n} ${x.cs}`)));
      c('…the table too, and no shipping-reports banner over numbers that are gone', s.invTr === 0 && /^Not loaded/.test(s.invEmpty) && !s.banners.some(x => x.id === 'reportsBanner' && /PO entr/.test(x.text)),
        `${s.invEmpty} | ${JSON.stringify(s.banners.map(x => x.id))}`);
      // Both come back.
      w.o.ctFile = 'ct.xlsx'; w.o.drayFile = w.o.drayOldFile = 'dray-port.csv';
      r = await runRefresh(t, w);
      s = await s2(t); snap = await t.snap();
      c('both back on the next refresh: the tiles equal the engine with drayage again', r.fired === 1 && tilesAre(s, R.totals.units) && snap.chips.drayage.state === 'ok' && snap.chips.ct.state === 'ok', tilesSay(s, R.totals.units));
      c('no shipping-reports banner', !s.banners.some(x => x.id === 'reportsBanner'), JSON.stringify(s.banners));
      // A Pull from China panel follows its line when the line's stage changes on a Refresh (its allocation ids change).
      const moved = R.pull.find(p => /At port/.test(p.ctStatus) && R0.pull.some(q => q.lineKey === p.lineKey && !/At port/.test(q.ctStatus)));
      c('the invented files have an open line that drayage moves to At port', !!moved, JSON.stringify(R.pull.map(p => p.ctStatus)));
      if (moved) {
        await t.click('#tabs .tab[data-tab="pull"]'); await sleep(200);
        await setVal(t, 'fq', moved.po); await sleep(400);
        await t.eval(`(() => { const tr = [...document.querySelectorAll('#pullTable tbody tr.r')].find(r => r.textContent.includes(${JSON.stringify(moved.styleRaw || moved.style)})); tr.querySelector('button.nb').click(); })()`);
        await sleep(200);
        const p0 = await s2(t);
        w.o.drayFile = w.o.drayOldFile = 'dray-bad.csv';
        const from = w.calls.length;
        await t.click('#btnRefresh');
        const end = Date.now() + 15000;
        while (w.calls.length === from && Date.now() < end) await sleep(100);
        await t.waitFor(IDLE, 'the refresh to finish', 30000);
        const p1 = await s2(t), note = await t.eval(`document.getElementById('panelNote').textContent`);
        c("a Pull from China panel keeps its line after a Refresh moves the line's stage (At port, then On the water without drayage)",
          p0.panel && p1.panel && /At port/.test(p0.panelFirst) && p1.panelTr >= 1 && /On the water/.test(p1.panelFirst) && !note, `${p0.panelFirst} -> ${p1.panelFirst} (${p1.panelTr} rows) ${note}`);
        await key(t, 'Escape');
        w.o.drayFile = w.o.drayOldFile = 'dray-port.csv';
        await setVal(t, 'fq', ''); await sleep(300);
        await t.click('#tabs .tab[data-tab="inv"]'); await sleep(200);
      }
      t.shots.push(await t.shot('s2-refresh'));
    },
  },

  // ---------------------------------------------------------------- Available to sell
  {
    name: 'sell-ok', title: 'Available to sell with a fresh ATS report: tiles, rows, each pull-from rule in the side panel, filters, Problems, the other tabs',
    world: { ...SELLW, ats: SELL_ATS },
    async run({ t, w, c, today }) {
      const { R, A, S } = sellFor(w.fx, today, { ...SELL_FILES, ats: SELL_ATS });
      c('the engines (run here) built the invented sell files with the report', R.ok && R.balance.diff === 0 && R.balance.ct.diff === 0 && A.ok && !!S && S.ok && !S.atsMissing,
        JSON.stringify({ R: R.ok && R.balance.diff, A: A.ok || A.error, S: S && S.ok }));
      const lot = (k, po) => S.lots.find(l => l.style === k && String(l.po) === String(po)), sty = k => S.styles.find(x => x.style === k) || {};
      const L = { etd: lot('SL-101', 900501), crd: lot('SL-102', 900502), crd2: lot('SL-102', 900503), order: lot('SL-103', 900504), now: lot('SL-104', 900505),
        eta: lot('SL-105', 900506), dis: lot('SL-106', 900507), now2: lot('SL-109', 900509) };
      const shape = {
        etd: !!L.etd && L.etd.pullRule === 'etd' && L.etd.pullFrom === today + 20 && L.etd.status === 'Booked',
        crd: !!L.crd && L.crd.pullRule === 'crd' && L.crd.pullFrom === today + 35 && L.crd.taken === 1200 && L.crd.available === 0,
        crd2: !!L.crd2 && L.crd2.pullFrom === today + 85 && L.crd2.taken === 300 && L.crd2.available === 180,
        order: !!L.order && L.order.pullRule === 'order' && L.order.pullFrom === today + 65 && L.order.estimate && L.order.status === 'No CRD',
        now: !!L.now && L.now.pullRule === 'crd' && L.now.flags.includes('now') && L.now.status === 'Ready' && L.now.available === 0,
        nowLeft: !!L.now2 && L.now2.flags.includes('now') && L.now2.taken === 100 && L.now2.available === 500,
        eta: !!L.eta && !L.eta.inChina && L.eta.status === 'On the water' && L.eta.availableDate === today + 20,
        dis: !!L.dis && L.dis.flags.includes('dates-disagree') && L.dis.readyFrom === 'factoryCrd',
        oversold: sty('SL-104').oversold === 3800, notOnReport: (sty('SL-103').flags || []).includes('not-on-report'),
        erpOnly: (sty('SL-108').flags || []).includes('erp-only'), sold: sty('SL-107').soldChina === 720, atsIssue: A.issues.some(i => i.code === 'erp-ats-differs') };
      c('the invented files exercise every rule: ETD, CRD, the order date, "now", ETA, ready dates that disagree, orders over two lots, oversold, not on the report, report only, sold',
        Object.values(shape).every(Boolean), JSON.stringify(shape));
      const rules = ruleMismatches(S, today), sums = sellSums(S, R);
      c("Saul's pull-from rule and the ready date hold on every lot (worked out here on their own)", !rules.length, rules.join(' | '));
      c('the selling sums hold (styles, lots, stock after orders, earliest lots first, Incoming, sold)', !sums.length, sums.join(' | '));

      const s = await t.snap();
      let p = await sl(t);
      c('Available to sell is shown at start', p.tab === 'sell' && p.shown && s.tabs.sell, `${p.tab} ${p.shown}`);
      c(`ATS chip green: "${A.counts.rowsRead} read / ${A.counts.styleRows} counted"`, s.chips.ats.state === 'ok' && s.chips.ats.text.includes(`${fmtN(A.counts.rowsRead)} read / ${fmtN(A.counts.styleRows)} counted`),
        `${s.chips.ats.state}: ${s.chips.ats.text}`);
      c('no note over the tab, no lead card', !p.note && !p.lead, `${p.note} ${p.lead}`);
      c('the filter bar is on it, with "Only styles with something available" and without the stage and PO type filters', p.fbarIn === 'tab-sell' && p.availShown && !p.stageShown && !p.typeShown,
        JSON.stringify([p.fbarIn, p.availShown, p.stageShown, p.typeShown]));
      for (const k of SELL_TILE_KEYS) {
        const x = sellTile(p, k), cs = casesSay(S.totals.cases[k], S.totals.noCasePackUnits[k], S.totals.looseUnits[k]);
        c(`tile ${x.lbl || k}: the engine's ${fmtN(S.totals[k])} units, "${cs}"`, unitsIs(x.n, S.totals[k]) && x.cs[0] === cs, `${x.n} · ${x.cs.join(' | ')}`);
      }
      const ov = sellTile(p, 'oversold'), nOv = S.styles.filter(x => x.oversold > 0).length;
      c(`tile Oversold styles: ${nOv}, "${fmtN(S.totals.oversold)} units oversold", in red`, ov.n.replace(/styles?$/, '') === String(nOv) && ov.cs[0] === `${fmtN(S.totals.oversold)} units oversold` && /\bbad\b/.test(ov.cls),
        `${ov.n} · ${ov.cs} · ${ov.cls}`);
      c('From China, Arriving and Total open a panel; Available now and Oversold do not', ['fromChina', 'arriving', 'total'].every(k => sellTile(p, k).btn) && !sellTile(p, 'now').btn && !sellTile(p, 'oversold').btn,
        JSON.stringify(p.tiles.map(x => [x.k, x.btn])));
      c('the columns, in order', p.heads.map(h => h.t).join('|') === SELL_HEADS.join('|'), p.heads.map(h => h.t).join('|'));
      c(`one row per style (${S.styles.length}), sorted by style`, p.tr === S.styles.length && p.rows.map(r => r.style).join() === S.styles.map(x => x.styleRaw).sort().join(), p.rows.map(r => r.style).join());
      const diff = sellRowsDiffer(p, S, today, 'ok');
      c("every row: in warehouse, customer orders, available now, from China, first pull date, arriving, first arrival, total and sold are the engine's (blank for 0)", !diff.length, diff.join(' | '));
      const row = k => p.rows.find(r => r.style === k) || { tags: [] };
      c('SL-103 (not on the report): "not on report" across the stock columns, a tag, its Total in muted type said to be from China and arriving only',
        row('SL-103').span === 'not on report' && row('SL-103').tags.includes('not on ATS report') && /\bpt\b/.test(row('SL-103').totalCls) && /from China and arriving only/.test(row('SL-103').totalTitle),
        JSON.stringify(row('SL-103')));
      c(`First pull date: SL-103's is the estimate "about ${C.fmtDay(today + 65)}"; SL-109's "now" (its CRD passed); SL-104's blank (customer orders took its lot)`,
        row('SL-103').firstPull === `about ${C.fmtDay(today + 65)}` && row('SL-109').firstPull === 'now' && row('SL-104').firstPull === '',
        `${row('SL-103').firstPull} · ${row('SL-109').firstPull} · ${row('SL-104').firstPull}`);
      c('SL-104: tagged "oversold 3,800", Available now blank', row('SL-104').tags.includes('oversold 3,800') && row('SL-104').now === '' && row('SL-104').nowCell === '', JSON.stringify(row('SL-104')));
      c('SL-108 (on the report only): a row with its stock, nothing coming', row('SL-108').inWh === '50' && row('SL-108').now === '50' && !row('SL-108').fromChina && !row('SL-108').arriving, JSON.stringify(row('SL-108')));
      c('SL-102: case pack 12, its UPC under the style', row('SL-102').casePack === '12' && row('SL-102').upc === '000000005024', JSON.stringify(row('SL-102')));
      t.shots.push(await t.shot('sell-ok'));

      // The side panel of every style with lots: each lot in date order, the date a buyer can have it and its rule in words.
      const panelBad = [];
      for (const st of S.styles.filter(x => x.lots.length)) {
        await t.eval(`[...document.querySelectorAll('#sellTable tbody tr.r button.sb')].find(x => x.textContent.trim() === ${JSON.stringify(st.styleRaw)}).click()`);
        await sleep(120);
        p = await sl(t);
        if (!p.panel || p.panelTitle !== `${st.styleRaw}: available to sell`) { panelBad.push(`${st.style}: panel "${p.panelTitle}"`); continue; }
        if (p.lots.length !== st.lots.length) panelBad.push(`${st.style}: ${p.lots.length} lots shown, ${st.lots.length} in the engine`);
        st.lots.forEach((l, j) => {
          const r = p.lots[j] || {}, cell = h => r[h] || { v: '', u: '', title: '' }, w2 = whenSays(l, today), af = cell('Available from');
          const got = { status: cell('Status').v, po: cell('PO').v, units: cell('Units').v, taken: cell('Taken by orders').v, avail: cell('Available').v };
          const want = { status: l.status, po: String(l.po), units: fmtN(l.units), taken: nz(l.taken), avail: nz(l.available) };
          for (const k of Object.keys(want)) if (got[k] !== want[k]) panelBad.push(`${st.style} lot ${j + 1} ${k}: "${got[k]}", not "${want[k]}"`);
          if (af.v !== w2.main || !w2.words.every(x => af.title.includes(x))) panelBad.push(`${st.style} PO ${l.po} Available from: "${af.v}" / "${af.title}", not "${w2.main}" with ${JSON.stringify(w2.words)}`);
          const rd = cell('Ready date (CRD)');
          if (!rd.v.startsWith(l.readyDate == null ? 'none' : C.fmtDay(l.readyDate)) || /dates disagree/.test(rd.v) !== l.flags.includes('dates-disagree')) panelBad.push(`${st.style} PO ${l.po} ready date: "${rd.v}"`);
          if (l.readySources.length > 1 && !l.readySources.every(x => rd.u.includes(md(x.day, today)))) panelBad.push(`${st.style} PO ${l.po} ready date sources: "${rd.u}"`);
          if (cell('ETD').v !== (l.etd == null ? '' : C.fmtDay(l.etd)) || cell('ETA').v !== (l.eta == null ? '' : C.fmtDay(l.eta))) panelBad.push(`${st.style} PO ${l.po} ETD "${cell('ETD').v}", ETA "${cell('ETA').v}"`);
        });
        // The stock line under the lots.
        const stockWant = (st.flags || []).includes('not-on-report') ? ['not on the ATS report'] : [`${fmtN(st.inWh)} in warehouse − ${fmtN(st.orders)} customer orders = ${fmtN(st.inWh - st.orders)} net`,
          ...(st.ordersTaken ? [`(${fmtN(st.ordersTaken)} taken)`] : []), ...(st.oversold ? [`${fmtN(st.oversold)} more than every lot holds: oversold`] : [])];
        if (!stockWant.every(x => p.panelStock.includes(x))) panelBad.push(`${st.style} stock line: "${p.panelStock}", not ${JSON.stringify(stockWant)}`);
        await key(t, 'Escape'); await sleep(60);
      }
      c('the side panel of every style with lots: each lot in date order with its status, PO, units, taken by orders, available, the pull-from date ("now" once passed) or ETA with its rule in words, the ready date (every source when they differ), ETD and ETA; then the stock line',
        !panelBad.length, panelBad.join(' | '));
      // Two looked at closely: SL-102 (orders take its first lot and part of the second) and SL-106 (ready dates disagree).
      await t.eval(`[...document.querySelectorAll('#sellTable tbody tr.r button.sb')].find(x => x.textContent.trim() === 'SL-102').click()`); await sleep(150);
      p = await sl(t);
      c(`SL-102: "2 lots · ${fmtN(sty('SL-102').fromChina)} available from China · ${fmtN(sty('SL-102').ordersTaken)} taken by customer orders", "all taken" under the first lot`, p.panelLbl === 'What can be sold' &&
        p.panelSub === `2 lots · ${fmtN(sty('SL-102').fromChina)} available from China · ${fmtN(sty('SL-102').ordersTaken)} taken by customer orders` && !!p.lots[0] && p.lots[0]['Available'].u === 'all taken',
        `${p.panelLbl} · ${p.panelSub} · ${p.lots[0] && JSON.stringify(p.lots[0]['Available'])}`);
      t.shots.push(await t.shot('sell-ok-panel'));
      await key(t, 'Escape'); await sleep(60);
      c('Esc closes the panel', !(await sl(t)).panel);
      await t.eval(`[...document.querySelectorAll('#sellTable tbody tr.r button.sb')].find(x => x.textContent.trim() === 'SL-106').click()`); await sleep(150);
      p = await sl(t);
      const rd6 = (p.lots[0] || {})['Ready date (CRD)'] || {};
      c(`SL-106: ready date ${C.fmtDay(today + 40)} "dates disagree", Factory CRD used, CRD Expected ${md(today + 60, today)} beside it`,
        rd6.v === `${C.fmtDay(today + 40)} dates disagree` && rd6.u.includes(`Factory CRD ${md(today + 40, today)} (used)`) && rd6.u.includes(`CRD Expected ${md(today + 60, today)}`), JSON.stringify(rd6));
      await t.click('#btnPanelClose'); await sleep(60);

      // The tiles' panels, and Sold (ships from China), which isn't lots: its number opens the stage lines behind it.
      await t.click('#sellTiles button.tile[data-t="fromChina"]'); await sleep(150);
      p = await sl(t);
      const china = S.lots.filter(l => l.inChina);
      c(`the From China tile lists every lot still in China (${china.length}), with a Style column, adding up to the tile`, p.panel && p.panelTitle === 'Available from China: every style' &&
        p.panelHead[0] === 'Style' && p.panelTr === china.length && p.panelSub.includes(`${fmtN(S.totals.fromChina)} available from China`), `${p.panelTitle} · ${p.panelHead} · ${p.panelTr} · ${p.panelSub}`);
      await key(t, 'Escape'); await sleep(60);
      await t.click('#sellTiles button.tile[data-t="total"]'); await sleep(150);
      p = await sl(t);
      c(`the Total tile lists every lot and adds the ${fmtN(S.totals.now)} available now from stock`, p.panel && p.panelTr === S.lots.length && p.panelStock.includes(`Plus ${fmtN(S.totals.now)} available now from stock`),
        `${p.panelTr} · ${p.panelStock}`);
      await key(t, 'Escape'); await sleep(60);
      await t.eval(`[...document.querySelectorAll('#sellTable tbody tr.r')].find(x => x.querySelector('button.sb').textContent.trim() === 'SL-107').querySelector('button.nb[data-b="soldChina"]').click()`); await sleep(150);
      p = await sl(t);
      c("SL-107's Sold number opens its lines: 720 units counted in Sold, ships from China", p.panel && p.panelTitle === 'SL-107 · Sold, ships from China' && p.panelSub.includes('720 units counted'), `${p.panelTitle} · ${p.panelSub}`);
      await key(t, 'Escape'); await sleep(60);
      // The stage and PO type filters (set on Where it is, hidden here) don't empty that panel.
      await setVal(t, 'fStage', 'inChina'); await setVal(t, 'fType', 'LA'); await sleep(200);
      await t.eval(`[...document.querySelectorAll('#sellTable tbody tr.r')].find(x => x.querySelector('button.sb').textContent.trim() === 'SL-107').querySelector('button.nb[data-b="soldChina"]').click()`); await sleep(150);
      p = await sl(t);
      c("with a stage and PO type chosen on Where it is (hidden here), SL-107's Sold still opens its 720 units, and the panel names no hidden filter",
        p.panel && p.panelSub.includes('720 units counted') && !/stage|PO type/.test(p.panelNote || ''), `${p.panelSub} · ${p.panelNote}`);
      await key(t, 'Escape'); await sleep(60);
      await setVal(t, 'fStage', ''); await setVal(t, 'fType', ''); await sleep(200);

      // Filters: the rows and the tiles follow them.
      await setVal(t, 'fq', 'sl-102'); await sleep(400);
      p = await sl(t);
      const s102 = sty('SL-102');
      c('search: one style; the tiles are its numbers, marked filtered', p.tr === 1 && p.rows[0].style === 'SL-102' && p.info === `1 style of ${S.styles.length}` &&
        unitsIs(sellTile(p, 'fromChina').n, s102.fromChina) && unitsIs(sellTile(p, 'total').n, s102.total) && /filtered/.test(sellTile(p, 'total').lbl), `${p.tr} · ${p.info} · ${sellTilesSay(p, null)}`);
      await setVal(t, 'fq', '000000005031'); await sleep(400);
      p = await sl(t);
      c('search by UPC finds its style', p.tr === 1 && p.rows[0].style === 'SL-103', JSON.stringify(p.rows.map(r => r.style)));
      await setVal(t, 'fq', ''); await sleep(300);
      await t.eval(`(() => { const el = document.getElementById('fAvail'); el.checked = true; el.dispatchEvent(new Event('change', { bubbles: true })); })()`); await sleep(300);
      p = await sl(t);
      const avail = S.styles.filter(x => x.now > 0 || x.fromChina > 0 || x.arriving > 0);
      c(`"Only styles with something available": ${avail.length} of ${S.styles.length} (not the all-taken or sold-only ones), the tiles unchanged`,
        p.tr === avail.length && p.info === `${avail.length} styles of ${S.styles.length}` && !p.rows.some(r => r.style === 'SL-104' || r.style === 'SL-107') && SELL_TILE_KEYS.every(k => unitsIs(sellTile(p, k).n, S.totals[k])),
        `${p.tr} · ${p.info} · ${p.rows.map(r => r.style)}`);
      await t.click('#btnClearF'); await sleep(300);
      p = await sl(t);
      c('Clear filters clears it', p.tr === S.styles.length && !(await t.eval(`document.getElementById('fAvail').checked`)), `${p.tr}`);
      // The 5-minute refresh reads the report again and keeps the search.
      await setVal(t, 'fq', 'sl-102'); await sleep(400);
      const r5 = await runRefresh(t, w);
      p = await sl(t);
      c('the 5-minute refresh reads the ATS report again, keeps the search and shows the same numbers', r5.fired === 1 && r5.calls.some(x => x.path === `/drives/${SRC}/items/itm-ats-1`) &&
        (await t.eval(`document.getElementById('fq').value`)) === 'sl-102' && p.tr === 1 && unitsIs(sellTile(p, 'total').n, s102.total), `fired ${r5.fired} · ${p.tr} · ${sellTilesSay(p, null)}`);
      await setVal(t, 'fq', ''); await sleep(300);

      // Pull from China: each open line's pull-from date and what's left after orders; China Pickup "Sold".
      await t.click('#tabs .tab[data-tab="pull"]'); await sleep(250);
      const pull = await t.eval(`(() => { const hs = [...document.querySelectorAll('#pullTable thead th')].map(th => { const x = th.cloneNode(true); x.querySelectorAll('.ar').forEach(y => y.remove()); return x.textContent.trim(); });
        const ix = h => hs.indexOf(h), plain = td => { const x = td.cloneNode(true); x.querySelectorAll('.u').forEach(y => y.remove()); return x.textContent.trim(); };
        return { hs, rows: [...document.querySelectorAll('#pullTable tbody tr.r')].map(tr => ({ po: plain(tr.children[0]), style: plain(tr.children[1]),
          when: ix('Pull-from date') < 0 ? null : plain(tr.children[ix('Pull-from date')]), avail: ix('Available after orders') < 0 ? null : plain(tr.children[ix('Available after orders')]) })) }; })()`);
      const pr = (k, po) => pull.rows.find(r => r.style === k && r.po === String(po)) || {};
      c('Pull from China: the Pull-from date and Available after orders columns', pull.hs.includes('Pull-from date') && pull.hs.includes('Available after orders'), pull.hs.join('|'));
      const pullWant = [['SL-101', 900501, whenSays(L.etd, today).main, '1,200'], ['SL-102', 900502, whenSays(L.crd, today).main, 'all taken'], ['SL-102', 900503, whenSays(L.crd2, today).main, '180'],
        ['SL-103', 900504, whenSays(L.order, today).main, '600'], ['SL-104', 900505, 'now', 'all taken'], ['SL-109', 900509, 'now', '500'], ['SL-107', 900508, 'Sold', 'Sold']];
      const pullBad = pullWant.filter(([k, po, wh, av]) => pr(k, po).when !== wh || pr(k, po).avail !== av).map(([k, po, wh, av]) => `${k} PO ${po}: "${pr(k, po).when}" / "${pr(k, po).avail}", not "${wh}" / "${av}"`);
      c('…each open line: its pull-from date ("now" once passed, "about" for an estimate) and what is left after orders ("all taken"); the China Pickup line "Sold"', !pullBad.length, pullBad.join(' | '));
      t.shots.push(await t.shot('sell-ok-pull'));

      // Where it is: On hand is In WH from the report.
      await t.click('#tabs .tab[data-tab="inv"]'); await sleep(250);
      const oh = await t.eval(`Object.fromEntries([...document.querySelectorAll('#invTable tbody tr.r:not(.upc)')].map(r => [r.querySelector('button.sb').textContent.trim(), r.children[4].textContent.trim()]))`);
      c('Where it is: On hand is In WH from the ATS report (text numbers read; blank for 0; "not on report" for SL-103)', oh['SL-101'] === '1,000' && oh['SL-102'] === '100' && oh['SL-104'] === '' && oh['SL-105'] === '300' && oh['SL-103'] === 'not on report',
        JSON.stringify(oh));
      c("Where it is: its tiles are still the stage engine's", tilesAre(await s2(t), R.totals.units), tilesSay(await s2(t), R.totals.units));

      // Problems: Available to sell and the ATS report, each in plain words.
      await t.click('#tabs .tab[data-tab="problems"]'); await sleep(250);
      const sp = await s2(t), order = sp.issGroups.map(g => g.split(' · ')[0]);
      const gSell = sp.issGroups.find(g => g.startsWith('Available to sell')) || '', gAts = sp.issGroups.find(g => g.startsWith('ATS report')) || '';
      const sellCodes = [...new Set(S.issues.map(i => i.code))], atsCodes = [...new Set(A.issues.map(i => i.code))];
      c(`Problems: "Available to sell" (${S.issues.length}: ${sellCodes.join(', ')}), then "ATS report" (${A.issues.length}: ${atsCodes.join(', ')}), after Incoming and before the sources, in plain words`,
        new RegExp(`^Available to sell · ${S.issues.length} issues?`).test(gSell) && new RegExp(`^ATS report · ${A.issues.length} issues?`).test(gAts) &&
        order.indexOf('Available to sell') === order.indexOf('Incoming') + 1 && order.indexOf('ATS report') === order.indexOf('Available to sell') + 1 &&
        ![...sellCodes, ...atsCodes].some(k => (gSell + gAts).includes(k)), sp.issGroups.join(' | '));
      const pt = await t.eval(`document.getElementById('issTable').textContent`);
      c('…naming the oversold style, the PO and style whose ready dates disagree, the style on the report only and the one whose ERP ATS differs',
        pt.includes('Style SL-104: customer orders (5,000) are 3,800 units more') && pt.includes('PO 900507, style SL-106: the ready dates are up to 20 days apart') && /SL-108/.test(pt) && /Style SL-101 \(row \d+\): the ERP's "ATS from in WH" is 700/.test(pt),
        pt.slice(0, 300));
      c('no "page sums differ" issue (the rows add up to the engine\'s totals)', !/page adds the styles up/.test(pt));
      await t.click('#tabs .tab[data-tab="sell"]'); await sleep(150);
    },
  },
  {
    name: 'sell-missing', title: "Available to sell with no ATS report yet: customer orders aren't taken off, said so, nothing partial", world: { ...SELLW },
    async run({ t, w, c, today }) {
      const { R, S } = sellFor(w.fx, today, SELL_FILES);
      c('the engine (run here) without the report', R.ok && !!S && S.ok && S.atsMissing, JSON.stringify({ R: R.ok, S: S && S.atsMissing }));
      const s = await t.snap();
      let p = await sl(t);
      c('ATS chip grey "no report yet"', s.chips.ats.state === 'off' && /no report yet/.test(s.chips.ats.text), s.chips.ats.text);
      c("an amber note: customer orders aren't taken off until the weekly ATS report is in Stock/ATS on the TestSite site", /\bwarn\b/.test(p.noteCls) &&
        /Customer orders aren't taken off until the weekly ATS report is in Stock\/ATS on the TestSite site/.test(p.note), `${p.noteCls}: ${p.note}`);
      c('Available now, Total available and Oversold say "no report yet", with no number', ['now', 'total', 'oversold'].every(k => sellTile(p, k).n === 'no report yet' && !/\d/.test(sellTile(p, k).cs.join(''))),
        JSON.stringify(p.tiles.map(x => [x.k, x.n, x.cs])));
      c(`From China and Arriving: the lots in full (${fmtN(S.totals.fromChina)}, ${fmtN(S.totals.arriving)}), marked "customer orders not taken off"`,
        ['fromChina', 'arriving'].every(k => unitsIs(sellTile(p, k).n, S.totals[k]) && sellTile(p, k).cs.includes('customer orders not taken off')), sellTilesSay(p, S));
      const diff = sellRowsDiffer(p, S, today, 'no report yet');
      c('every row: "no report yet" across the stock columns and in Total; from China and arriving in full', p.tr === S.styles.length && !diff.length, `${p.tr} rows · ${diff.join(' | ')}`);
      await t.eval(`[...document.querySelectorAll('#sellTable tbody tr.r button.sb')].find(x => x.textContent.trim() === 'SL-102').click()`); await sleep(150);
      p = await sl(t);
      c('SL-102: both lots in full, nothing taken; the stock line says why', p.lots.length === 2 && p.lots.every(r => r['Taken by orders'].v === '') && p.lots[0]['Available'].v === '1,200' &&
        p.panelStock.startsWith('Stock: no report yet.') && /Customer orders aren't taken off/.test(p.panelStock), `${JSON.stringify(p.lots.map(r => [r['Taken by orders'].v, r['Available'].v]))} · ${p.panelStock}`);
      await key(t, 'Escape'); await sleep(60);
      t.shots.push(await t.shot('sell-missing'));
      await t.click('#tabs .tab[data-tab="problems"]'); await sleep(250);
      const sp = await s2(t), g = sp.issGroups.find(x => x.startsWith('Available to sell')) || '';
      c(`Problems: "Available to sell" has the one line saying customer orders aren't known (${S.issues.length})`, new RegExp(`^Available to sell · ${S.issues.length} issues?`).test(g) && /Customer orders not known/.test(g) &&
        !sp.issGroups.some(x => x.startsWith('ATS report')), sp.issGroups.join(' | '));
      await t.click('#tabs .tab[data-tab="sell"]'); await sleep(100);
      c('no email', sendMails(w.calls).length === 0);
    },
  },
  {
    name: 'sell-bad', title: "Available to sell with an ATS report that doesn't add up: red, Saul emailed, Available now not loaded", world: { ...SELLW, ats: { age: 2, file: 'ats-sell-bad.xlsx' } },
    async run({ t, w, c, today }) {
      const { R, A, S } = sellFor(w.fx, today, { ...SELL_FILES, ats: w.o.ats });
      c("the report (read here) doesn't add up, so the engine works without it", !!A && !A.ok && /doesn't add up: rows 1,450, Grand Total 9,999/.test(A.error) && !!S && S.atsMissing, A && A.error);
      const s = await t.snap();
      let p = await sl(t);
      c('ATS chip red "doesn\'t add up", with both totals in its tooltip', s.chips.ats.state === 'err' && /doesn't add up/.test(s.chips.ats.text) && /rows 1,450, Grand Total 9,999/.test(s.chips.ats.title),
        `${s.chips.ats.state}: ${s.chips.ats.text} | ${s.chips.ats.title}`);
      greenExcept(c, s, []);
      oneEmail(c, s, w.calls, 'ats', "doesn't add up", today);
      c("a red note: the report couldn't be used, why, and that Available now says not loaded", /\berr\b/.test(p.noteCls) && /The ATS report couldn't be used:/.test(p.note) && /doesn't add up/.test(p.note) &&
        /Available now says "not loaded"/.test(p.note), `${p.noteCls}: ${p.note}`);
      c('Available now, Total available and Oversold "not loaded"; From China and Arriving the lots in full', ['now', 'total', 'oversold'].every(k => sellTile(p, k).n === 'not loaded') &&
        ['fromChina', 'arriving'].every(k => unitsIs(sellTile(p, k).n, S.totals[k])), sellTilesSay(p, S));
      const diff = sellRowsDiffer(p, S, today, 'not loaded');
      c('every row: "not loaded" across the stock columns and in Total', p.tr === S.styles.length && !diff.length, `${p.tr} rows · ${diff.join(' | ')}`);
      c('the stage numbers are not held up (no red stage banner)', !/stage numbers are not loaded/i.test(banner(s, 'err')), banner(s, 'err'));
      await t.click('#tabs .tab[data-tab="inv"]'); await sleep(250);
      const q = await s2(t), oh = await t.eval(`[...new Set([...document.querySelectorAll('#invTable tbody tr.r:not(.upc)')].map(r => r.children[4].textContent.trim()))]`);
      c('Where it is: the stage tiles as the engine; On hand "not loaded"', tilesAre(q, R.totals.units) && oh.join() === 'not loaded', `${tilesSay(q, R.totals.units)} · ${JSON.stringify(oh)}`);
      t.shots.push(await t.shot('sell-bad'));
      await t.click('#tabs .tab[data-tab="sell"]'); await sleep(100);
    },
  },
  {
    name: 'sell-big', title: 'Available to sell on 2,400 styles with an ATS report: tiles, windowed rows, scroll, sort, only available', world: { ...BIG, ats: { age: 1, file: 'ats-big.xlsx' } },
    async run({ t, w, c, today }) {
      const { R, A, S } = sellFor(w.fx, today, { ...BIG_FILES, ats: w.o.ats });
      c('the engines (run here): over 2,400 styles, customer orders taken off lots, some oversold, three on the report only', R.ok && A.ok && !!S && S.ok && S.styles.length > 2400 && S.totals.ordersTaken > 0 &&
        S.totals.oversold > 0 && S.counts.erpOnlyStyles === 3, JSON.stringify({ styles: S && S.styles.length, taken: S && S.totals.ordersTaken, oversold: S && S.totals.oversold, erp: S && S.counts.erpOnlyStyles }));
      const rules = ruleMismatches(S, today), sums = sellSums(S, R);
      c("Saul's pull-from rule and the ready date hold on every lot (worked out here on their own)", !rules.length, rules.slice(0, 5).join(' | '));
      c('the selling sums hold', !sums.length, sums.slice(0, 5).join(' | '));
      let p = await sl(t);
      for (const k of SELL_TILE_KEYS) c(`tile ${k} = the engine's ${fmtN(S.totals[k])}`, unitsIs(sellTile(p, k).n, S.totals[k]), sellTile(p, k).n);
      const nOv = S.styles.filter(x => x.oversold > 0).length;
      c(`tile Oversold styles: ${nOv}`, sellTile(p, 'oversold').n.replace(/styles?$/, '') === String(nOv), sellTile(p, 'oversold').n);
      c(`info line: ${fmtN(S.styles.length)} styles; windowed: far fewer rows drawn`, p.info === `${fmtN(S.styles.length)} styles` && p.tr > 10 && p.tr < 150, `${p.info} · ${p.tr}`);
      let diff = sellRowsDiffer(p, S, today, 'ok');
      c('the rows drawn are the engine\'s', !diff.length, diff.slice(0, 5).join(' | '));
      const lay = await t.eval(`(() => { const p = document.getElementById('sellPane'); p.scrollLeft = 600; p.dispatchEvent(new Event('scroll'));
        const td = document.querySelector('#sellTable tbody tr.r td.fz'), th = document.querySelector('#sellTable thead th'), r = p.getBoundingClientRect();
        const o = { wide: p.scrollWidth > p.clientWidth, left: Math.round(td.getBoundingClientRect().left - r.left - p.clientLeft), th: getComputedStyle(th).position,
          page: document.documentElement.scrollWidth <= window.innerWidth + 1, bottom: Math.round(r.bottom), h: innerHeight };
        p.scrollLeft = 0; p.dispatchEvent(new Event('scroll')); return o; })()`);
      c('the table scrolls in its own pane, ending inside the window: Style frozen, header sticky, no page-wide sideways scroll',
        lay.th === 'sticky' && lay.page && (!lay.wide || Math.abs(lay.left) <= 2) && lay.bottom <= lay.h && lay.bottom > lay.h - 40, JSON.stringify(lay));
      t.shots.push(await t.shot('sell-big'));
      await scrollTo(t, 'sellPane', 40000); await sleep(400);
      p = await sl(t);
      diff = sellRowsDiffer(p, S, today, 'ok');
      c('scrolling draws later rows, still windowed, still the engine\'s', p.first > 600 && p.tr < 150 && !diff.length, `first ${p.first}, ${p.tr} rows · ${diff.slice(0, 3).join(' | ')}`);
      await t.click('#sellTable th[data-s="total"]'); await sleep(200);
      p = await sl(t);
      const hi = p.heads.find(h => h.t === 'Total available'), max = Math.max(...S.styles.map(x => x.total));
      c(`header sort: Total available, biggest first (${fmtN(max)}), back to the top`, !!hi && hi.sort === 'descending' && p.first === 0 && p.scroll === 0 && p.rows[0].total === fmtN(max), `${JSON.stringify(hi)} · ${p.first} · ${p.rows[0] && p.rows[0].total}`);
      await t.click('#sellTable th[data-s="style"]'); await sleep(150);
      await t.eval(`(() => { const el = document.getElementById('fAvail'); el.checked = true; el.dispatchEvent(new Event('change', { bubbles: true })); })()`); await sleep(300);
      p = await sl(t);
      const nAv = S.styles.filter(x => x.now > 0 || x.fromChina > 0 || x.arriving > 0).length;
      c(`"Only styles with something available": ${fmtN(nAv)} of ${fmtN(S.styles.length)}`, p.info === `${fmtN(nAv)} styles of ${fmtN(S.styles.length)}`, p.info);
      await t.click('#btnClearF'); await sleep(200);
      await t.click('#tabs .tab[data-tab="problems"]'); await sleep(250);
      const pt = await t.eval(`document.getElementById('issTable').textContent`), sp = await s2(t);
      const gSell = sp.issGroups.find(g => g.startsWith('Available to sell')) || '';
      c(`Problems: Available to sell has the engine's ${S.issues.length} issues; no "page sums differ" issue`, new RegExp(`^Available to sell · ${fmtN(S.issues.length)} issues?`).test(gSell) && !/page adds the styles up/.test(pt), gSell.slice(0, 160));
      await t.click('#tabs .tab[data-tab="sell"]'); await sleep(100);
    },
  },
  {
    name: 'sell-real', title: 'Available to sell on local copies of the real files and the real ATS report (only when named, with REAL_DIR and REAL_ATS; numbers printed, nothing else)',
    world: { ...REAL, ats: { age: 1, file: 'real-ats.xlsx' } }, optIn: true,
    async run({ t, w, c, today }) {
      const f = { msl: 'real-msl.xlsx', trk: 'real-trk.xlsx', ct: 'real-ct.xlsx', fwd: 'real-fwd.xlsx', dray: 'real-dray.csv', pob: w.fx.real.pob, upc: w.fx.real.upcPages.flat(), sections: realSections(), ats: w.o.ats };
      const { R, A, S } = sellFor(w.fx, today, f);
      c('the engines (run here) on the real files: balance 0, the ATS report adds up', R.ok && R.balance.diff === 0 && R.balance.ct.diff === 0 && !!A && A.ok && !!S && S.ok && !S.atsMissing,
        JSON.stringify({ R: R.ok && R.balance.diff, A: A && (A.ok || A.error), S: S && S.ok }));
      const rules = ruleMismatches(S, today), sums = sellSums(S, R);
      c("Saul's pull-from rule and the ready date hold on every real lot (worked out here on their own)", !rules.length, `${rules.length} lots, e.g. ${rules.slice(0, 3).join(' | ')}`);
      c('the selling sums hold on the real files', !sums.length, sums.slice(0, 5).join(' | '));
      let p = await sl(t);
      const s = await t.snap();
      c('ATS chip green', s.chips.ats.state === 'ok', `${s.chips.ats.state}: ${s.chips.ats.text}`);
      for (const k of SELL_TILE_KEYS) c(`tile ${k} = the engine's ${fmtN(S.totals[k])}`, unitsIs(sellTile(p, k).n, S.totals[k]), sellTile(p, k).n);
      const nOv = S.styles.filter(x => x.oversold > 0).length;
      c(`tile Oversold styles: ${nOv}`, sellTile(p, 'oversold').n.replace(/styles?$/, '') === String(nOv), sellTile(p, 'oversold').n);
      c('rows windowed', p.tr > 10 && p.tr < 150, p.tr);
      const diff = sellRowsDiffer(p, S, today, 'ok');
      c("the rows drawn are the engine's", !diff.length, `${diff.length} cells differ`);
      const ms = await t.eval(`(() => { const el = document.getElementById('fAvail'); const t0 = performance.now(); el.checked = true; el.dispatchEvent(new Event('change', { bubbles: true })); const a = performance.now() - t0;
        const t1 = performance.now(); el.checked = false; el.dispatchEvent(new Event('change', { bubbles: true })); return [a, performance.now() - t1]; })()`);
      const sm = await t.eval(`(() => { const p = document.getElementById('sellPane'); const t0 = performance.now(); p.scrollTop = 30000; p.dispatchEvent(new Event('scroll')); return new Promise(r => requestAnimationFrame(() => r(performance.now() - t0))); })()`);
      const tm = await t.eval(`(() => { const t0 = performance.now(); document.querySelector('#sellTiles button.tile[data-t="fromChina"]').click(); return performance.now() - t0; })()`);
      p = await sl(t);
      c('the From China tile opens its lots, adding up to the tile', p.panel && p.panelTr > 0 && p.panelSub.includes(`${fmtN(S.totals.fromChina)} available from China`), p.panelSub.replace(/[^\d,· a-z]/gi, ''));
      c('"only available" redraws in under 300 ms; the tile panel opens in under 500 ms', ms[0] < 300 && ms[1] < 300 && tm < 500, `${ms.map(x => x.toFixed(0))} · ${tm.toFixed(0)}`);
      await key(t, 'Escape');
      await t.click('#tabs .tab[data-tab="problems"]'); await sleep(300);
      const pt = await t.eval(`document.getElementById('issTable').textContent`);
      c('no "page sums differ" issue (the rows add up to the engine\'s totals)', !/page adds the styles up/.test(pt));
      const T = S.totals, K = T.cases, N = T.noCasePackUnits, cnt = S.counts, codes = {};
      for (const i of S.issues) codes[i.code] = (codes[i.code] || 0) + 1;
      console.log('REAL Available to sell (units / cases / units with no case pack):');
      for (const k of ['now', 'fromChina', 'arriving', 'total', 'soldChina']) console.log(`  ${k}: ${fmtN(T[k])} / ${fmtN(K[k])} / ${fmtN(N[k])}`);
      console.log(`  oversold ${fmtN(T.oversold)} units on ${fmtN(nOv)} styles · customer orders taken off lots ${fmtN(T.ordersTaken)}`);
      console.log(`REAL lots ${fmtN(cnt.lots)} (in China ${fmtN(cnt.lotsInChina)}, sailed ${fmtN(cnt.lotsSailed)}) · pull rule ${Object.entries(cnt.byRule).map(([k, v]) => `${k} ${fmtN(v)} (${fmtN(cnt.byRuleUnits[k])} units)`).join(', ')} · "now" ${fmtN(cnt.pullNow)} · sailed ${Object.entries(cnt.sailedByRule).map(([k, v]) => `${k} ${fmtN(v)}`).join(', ')}`);
      console.log(`REAL dates disagree: ${fmtN(cnt.datesDisagreeLots)} lots on ${fmtN(cnt.datesDisagreePairs)} PO + style · styles ${fmtN(S.styles.length)}, with stock now ${fmtN(cnt.stylesWithStock)}, not on the report ${fmtN(cnt.stylesNotOnReport)}, report only ${fmtN(cnt.erpOnlyStyles)}`);
      console.log('REAL Available to sell issue codes: ' + Object.entries(codes).sort((a, b) => b[1] - a[1]).map(([k, v]) => `${k} ${v}`).join(', '));
      console.log(`REAL page: ${p.tr} rows drawn; only available ${ms.map(x => x.toFixed(0)).join('/')} ms, scroll ${sm.toFixed(0)} ms, From China panel ${tm.toFixed(0)} ms`);
      await t.click('#tabs .tab[data-tab="sell"]');
    },
  },
  {
    name: 's2-real', title: 'step 2 on local copies of the real files (only when named, with REAL_DIR; numbers printed, nothing else)', world: { ...REAL }, optIn: true,
    async run({ t, w, c, today }) {
      const R = stagesFor(w.fx, today, { msl: 'real-msl.xlsx', trk: 'real-trk.xlsx', ct: 'real-ct.xlsx', fwd: 'real-fwd.xlsx', dray: 'real-dray.csv', pob: w.fx.real.pob, upc: w.fx.real.upcPages.flat(), sections: realSections() });
      c("config.js's section lists are served to the page and used here", !!realSections() && Array.isArray(realSections().pullFromChina), JSON.stringify(Object.keys(realSections() || {})));
      await t.click('#tabs .tab[data-tab="inv"]'); await sleep(300);
      let s = await s2(t);
      c('the engine (run here) is ok on the real files, balance 0', R.ok && R.balance.diff === 0 && R.balance.ct.diff === 0, JSON.stringify(R.ok ? R.balance.diff : R.notLoaded));
      for (const k of TILE_KEYS) c(`tile ${k} = the engine's ${R.totals.units[k].toLocaleString('en-US')}`, unitsIs(tileN(s, k), R.totals.units[k]), tileN(s, k));
      c('rows windowed', s.invTr > 10 && s.invTr < 150, s.invTr);
      const ms = await t.eval(`(() => { const el = document.getElementById('fFac'); const v = [...el.options].map(o => o.value).find(x => x && x !== '__none__'); const t0 = performance.now(); el.value = v; el.dispatchEvent(new Event('change', { bubbles: true })); const a = performance.now() - t0; const t1 = performance.now(); el.value = ''; el.dispatchEvent(new Event('change', { bubbles: true })); return [a, performance.now() - t1]; })()`);
      c('a filter redraws in under 300 ms', ms[0] < 300 && ms[1] < 300, ms.map(x => x.toFixed(0)).join(' / '));
      const sm = await t.eval(`(() => { const p = document.getElementById('invPane'); const t0 = performance.now(); p.scrollTop = 20000; p.dispatchEvent(new Event('scroll')); return new Promise(r => requestAnimationFrame(() => r(performance.now() - t0))); })()`);
      const tm = await t.eval(`(() => { const t0 = performance.now(); document.querySelector('#tiles button.tile[data-t="inChina"]').click(); return performance.now() - t0; })()`);
      s = await s2(t);
      c('the In China tile opens its lines, adding up to the tile', s.panel && s.panelTr > 0 && s.panelSub.includes(`${R.totals.units.inChina.toLocaleString('en-US')} units counted`), s.panelSub);
      await key(t, 'Escape');
      await t.click('#tabs .tab[data-tab="problems"]'); await sleep(300);
      s = await s2(t);
      c('balance line green', /\bok\b/.test(s.balance.cls), s.balance.cls);
      c("no \"page sums differ\" issue (the page's sums of the allocations equal the engine's totals)", !(await t.eval(`[...document.querySelectorAll('#issTable td')].some(td => /page adds the allocations/.test(td.textContent))`)));
      const codes = {};
      for (const i of R.issues) codes[i.code] = (codes[i.code] || 0) + 1;
      const U = R.totals.units, K = R.totals.cases, N = R.totals.noCasePackUnits;
      console.log('REAL totals (units / cases / units with no case pack):');
      for (const k of Object.keys(U)) console.log(`  ${k}: ${fmtN(U[k])} / ${fmtN(K[k])} / ${fmtN(N[k])}`);
      console.log(`REAL balance: groups ${fmtN(R.balance.groups)}, Master Sheet Live ${fmtN(R.balance.mslUnits)} = outcomes ${fmtN(R.balance.outcomeUnits)}, diff ${R.balance.diff}, bad ${R.balance.bad.length}; Container Tracking ${fmtN(R.balance.ct.units)} = placed ${fmtN(R.balance.ct.outcomeUnits)}, diff ${R.balance.ct.diff}`);
      console.log('REAL issue codes: ' + Object.entries(codes).sort((a, b) => b[1] - a[1]).map(([k, v]) => `${k} ${v}`).join(', '));
      console.log(`REAL page: styles ${fmtN(R.styles.length)}, allocations ${fmtN(R.allocs.length)}, Pull from China ${fmtN(R.pull.length)}, rows drawn ${s.invTr}; filter ${ms.map(x => x.toFixed(0)).join('/')} ms, scroll ${sm.toFixed(0)} ms, In China panel ${tm.toFixed(0)} ms; Problems ${s.cProb} open`);
    },
  },
];

// ------------------------------------------------------------------ main

async function main() {
  const pick = (process.argv[2] || '').split(',').map(x => x.trim()).filter(Boolean);
  const unknown = pick.filter(n => !SCENARIOS.some(s => s.name === n));
  if (unknown.length) { console.log(`Unknown scenario: ${unknown.join(', ')}. Known: ${SCENARIOS.map(s => s.name).join(', ')}`); process.exit(2); }
  const want = new Set(pick.length ? pick : SCENARIOS.filter(s => !s.optIn).map(s => s.name));
  for (let grew = true; grew;) {   // a scenario that runs after another needs that one (and its own) first
    grew = false;
    for (const s of SCENARIOS) if (want.has(s.name) && s.after && !want.has(s.after)) { want.add(s.after); grew = true; }
  }
  const run = SCENARIOS.filter(s => want.has(s.name));

  mkdirSync(OUT, { recursive: true });
  const today = TODAY;
  const fx = buildFixtures(today);
  if (want.has('s2-real') || want.has('sell-real')) {
    const dir = process.env.REAL_DIR, lists = process.env.REAL_LISTS || (dir && join(dir, 'lists.json'));
    if (!dir || !existsSync(dir) || !existsSync(lists)) { console.log('s2-real and sell-real need REAL_DIR (the folder of real files) and its lists.json (or REAL_LISTS).'); process.exit(2); }
    addRealFiles(fx, resolve(dir), resolve(lists));
    if (want.has('sell-real')) {
      const ats = process.env.REAL_ATS;
      if (!ats || !existsSync(ats)) { console.log('sell-real needs REAL_ATS (the ERP\'s ATS report, .xlsx).'); process.exit(2); }
      fx.bytes.set('real-ats.xlsx', readFileSync(ats));
    }
  }
  const e = expected(fx, today);
  const results = [];
  const setup = { name: 'setup checks', title: 'fixtures, harness and the whole run', checks: [] };
  const sc = (label, ok, detail) => setup.checks.push({ label, ok: !!ok, detail: ok ? '' : String(detail ?? '') });
  sc('core/pobook.js block hashes to CORE_SHA256 (the fake PO Book page serves it)', fx.blockMatches);
  sc('the invented files parse and exercise the PO Book marks, issues and the stage engine', e.sane, JSON.stringify(e));
  const gaps = unlabelledCodes();
  sc('every issue code and flag core/stages.js, core/ats.js and core/sell.js give has plain words in app.js CODE_LABELS', !gaps.length, gaps.join(', '));
  const real = readFileSync(join(REPO, 'config.js'), 'utf8');
  sc('the test config.js is not the real one (invented locations, test client id)', !Object.values(LOC).some(v => real.includes(v)) && !real.includes(TEST_CONFIG.clientId) && !/sharepoint\.example/.test(real));
  results.push(setup);

  const page = pageServer(), dl = downloadServer();
  await new Promise(r => page.listen(0, '127.0.0.1', r));
  await new Promise(r => dl.listen(0, '127.0.0.1', r));
  const ORIGIN = `http://127.0.0.1:${page.address().port}`;
  DL = `http://127.0.0.1:${dl.address().port}`;
  const allowedNet = okNet(ORIGIN);
  const allCalls = [], allDownloads = [];

  let edge = null;
  try {
    edge = await launchEdge();
    console.log(`Edge ${edge.version}, headless, InPrivate, sync off · page ${ORIGIN} · downloads ${DL}`);
    const early = await signInWindows(edge.cdp);
    sc('Edge opened no sign-in or sync window', !early.length, early.join(', '));
    let t = null, w = null;
    for (const s of run) {
      const checks = [];
      const c = (label, ok, detail) => checks.push({ label, ok: !!ok, detail: ok ? '' : String(detail ?? '') });
      const r = { name: s.name, title: s.title, checks };
      results.push(r);
      try {
        if (!s.after) {
          if (t) await t.close();
          if (w) { allCalls.push(...w.calls); allDownloads.push(...w.downloads); }
          w = world = newWorld(fx, s.world);
          t = await openTab(edge.cdp, initScript(ORIGIN, s.init));
          t.shots = [];
          await t.send('Page.navigate', { url: `${ORIGIN}/index.html` });
          if (s.wait !== 'busy') await t.waitFor(READY, 'the sources to load');
        }
        const before = { blocked: t.blocked.length, exceptions: t.exceptions.length, net: s.after ? t.net.length : 0,
          calls: s.after ? w.calls.length : 0, downloads: s.after ? w.downloads.length : 0 };
        await s.run({ t, w, c, e, today });
        // Checks every scenario gets.
        const snap = await t.snap();
        const calls = w.calls.slice(before.calls), dls = w.downloads.slice(before.downloads), net = t.net.slice(before.net);
        const bad = [...snap.requests.filter(x => !okPage(x)).map(x => `page asked ${x.method} ${x.url}`),
          ...net.filter(x => !allowedNet(x)).map(x => `Edge sent ${x.method} ${x.url}`),
          ...calls.filter(x => !okCall(x)).map(x => `Graph got ${x.method} ${x.path}`),
          ...dls.filter(d => d.method !== 'GET').map(d => `download host got ${d.method} ${d.name}`),
          ...w.pageGets.filter(x => x.method !== 'GET').map(x => `page host got ${x.method} ${x.path}`)];
        c('read-only: no request but GETs and POST https://graph.microsoft.com/v1.0/me/sendMail', net.length > 0 && !bad.length, bad.join('; ') || 'no requests seen');
        c('title is "M1S · …"', /^M1S · /.test(snap.title), snap.title);
        c('fake sign-in in place: the vendored msal loaded and could not replace it', snap.msalFake && snap.msalSetTried >= 1, `fake ${snap.msalFake}, set tried ${snap.msalSetTried}`);
        const cfg = snap.msalConfig && snap.msalConfig.auth;
        c('sign-in set up with the test config.js; redirect is the folder address', !!cfg && cfg.clientId === TEST_CONFIG.clientId &&
          cfg.authority.endsWith(TEST_CONFIG.tenantId) && cfg.redirectUri === ORIGIN + '/', JSON.stringify(cfg));
        c('no interactive sign-in, sign-out or token popup', !snap.msalInteractive.length, snap.msalInteractive.join(','));
        const blocked = [...t.blocked.slice(before.blocked), ...snap.fetchBlocked];
        c('no request left 127.0.0.1', !blocked.length, blocked.join(', '));
        const ex = t.exceptions.slice(before.exceptions);
        c('no uncaught errors in the page', !ex.length, ex.join(' | '));
        c('every Graph call carried a bearer token', (s.noGraph || calls.length > 0) && calls.every(x => /^Bearer \S+/.test(x.auth)), `${calls.length} calls`);
        c('file downloads went through the download link, without the token', dls.every(d => !d.auth) && !calls.some(x => x.method === 'GET' && /\/content$/.test(x.path)),
          dls.filter(d => d.auth).map(d => d.name).join(','));
        c('every Graph call was one the stand-in knows', !w.unanswered.length, w.unanswered.join('; '));
        c('Edge opened no sign-in or sync window', !(await signInWindows(edge.cdp)).length);
      } catch (err) {
        c('scenario ran to the end', false, err.message);
        if (t) { try { t.shots.push(await t.shot(s.name + '-error')); } catch (e2) { /* page gone */ } }
      }
      r.shots = t ? t.shots.splice(0) : [];
    }
    if (t) await t.close();
    if (w) { allCalls.push(...w.calls); allDownloads.push(...w.downloads); }
    // The guards themselves: a request to another host must be refused by both. 127.0.0.2 is still this machine, so
    // nothing leaves it even if a guard were broken.
    world = newWorld(fx, {});
    t = await openTab(edge.cdp, initScript(ORIGIN));
    await t.send('Page.navigate', { url: `${ORIGIN}/index.html` });
    await t.waitFor(READY, 'the page to load');
    await t.eval(`(() => { new Image().src = 'http://127.0.0.2:9/guard-probe.png'; fetch('http://127.0.0.2:9/guard-probe').catch(() => {}); })()`);
    await sleep(800);
    const probe = await t.snap();
    sc('the request guard refuses a page request to another host', t.blocked.some(u => u.includes('guard-probe.png')), t.blocked.join(', ') || 'nothing refused');
    sc('the fake fetch refuses another host', probe.fetchBlocked.some(u => u.includes('guard-probe')), probe.fetchBlocked.join(', ') || 'nothing refused');
    await t.close();
    // The read-only rule over the whole run, every tab: and the one POST was really seen, so the check is live.
    const badNet = ALL_NET.filter(x => !allowedNet(x) && !/guard-probe/.test(x.url));
    sc(`whole run: every one of ${fmtN(ALL_NET.length)} requests was a GET or the email POST`, ALL_NET.length > 0 && !badNet.length, badNet.map(x => `${x.method} ${x.url}`).join('; '));
    sc('whole run: the stand-in got no write (no PUT, PATCH, DELETE, or POST other than /me/sendMail)', !allCalls.filter(x => !okCall(x)).length && !allDownloads.some(d => d.method !== 'GET'),
      allCalls.filter(x => !okCall(x)).map(x => `${x.method} ${x.path}`).join('; '));
    // A few scenarios run by name may send no email at all; the whole run always does.
    if (!pick.length || allCalls.some(x => x.method === 'POST' && x.path === '/me/sendMail')) {
      sc('whole run: the email POST was seen (the check above is live)', ALL_NET.some(x => x.method === 'POST' && /\/__graph\/me\/sendMail$/.test(x.url)));
    }
  } catch (err) {
    sc('Edge started and was driven', false, err.message);
  } finally {
    if (edge) {
      await edge.cdp.send('Browser.close').catch(() => {});
      await new Promise(r => { if (edge.proc.exitCode != null) r(); else { edge.proc.once('exit', r); setTimeout(r, 5000); } });
      try { edge.proc.kill(); } catch (e2) { /* already gone */ }
      for (let i = 0; i < 5; i++) { try { rmSync(edge.profile, { recursive: true, force: true }); break; } catch (e2) { await sleep(500); } }
    }
    page.close(); dl.close();
  }

  let failed = 0, total = 0;
  for (const r of results) {
    const bad = r.checks.filter(x => !x.ok).length;
    console.log(`\n${bad ? 'FAIL' : 'PASS'}  ${r.name}: ${r.title} (${r.checks.length - bad}/${r.checks.length})`);
    for (const x of r.checks) {
      total++;
      if (!x.ok) failed++;
      console.log(`  ${x.ok ? 'ok  ' : 'FAIL'}  ${x.label}${x.ok ? '' : `\n          got: ${x.detail}`}`);
    }
    for (const f of r.shots || []) console.log(`  screenshot ${f}`);
  }
  console.log(`\n${total - failed}/${total} checks passed${failed ? `, ${failed} failed` : ''}.`);
  process.exit(failed ? 1 : 0);
}

main().catch(err => { console.error(err); process.exit(1); });
