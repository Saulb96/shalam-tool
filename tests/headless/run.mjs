// The page in headless Edge against a stand-in for SharePoint: a fake sign-in (never a real one), a fake Graph that
// answers from invented workbooks, and checks on the chips, banners, the error email and the tabs.
//   node tests/headless/run.mjs [scenario,scenario…]      (default: all)
// Edge runs InPrivate, with sync off and a new throwaway profile; every request that isn't to 127.0.0.1 is blocked.
// Screenshots go to .tmp/. Exits 1 when any check fails.
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
const TEST_CONFIG = { tenantId: '00000000-0000-4000-8000-000000000000', clientId: '11111111-1111-4111-8111-111111111111' };
const ACCOUNT = { name: 'Test User', username: 'test.user@example.com', homeAccountId: 'test-home', localAccountId: 'test-local', tenantId: TEST_CONFIG.tenantId };
const MAIL_TO = 'saul@example.com';
const SITE_ID = 'site-prod', PROD = 'drv-prod', SRC = 'drv-src', REP = 'drv-rep';
const XOPT = { type: 'array', cellFormula: false, cellHTML: false, cellStyles: false };   // as app.js reads files
const KEYS = ['msl', 'tracker', 'pobook', 'ct', 'forwarder', 'drayage', 'ats', 'upc'];
const NAMES = { msl: 'Master Sheet Live', tracker: 'PO Status Tracker', pobook: 'PO Book lines', ct: 'Container Tracking',
  forwarder: 'Forwarder', drayage: 'Drayage', ats: 'ATS', upc: 'UPC_MASTER' };

const X = loadXLSX();
const C = core('common'), PB = core('pobook'), CT = core('ct'), SH = core('shipping'), UPC = core('upc');
const iso = days => new Date(Date.now() - days * 864e5).toISOString();
const fmtN = v => (typeof v === 'number' ? v.toLocaleString('en-US') : String(v));
const sleep = ms => new Promise(r => setTimeout(r, ms));
const sha = t => createHash('sha256').update(t, 'utf8').digest('hex');

// ------------------------------------------------------------------ invented source files

function buildFixtures() {
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
  bytes.set('dray.csv', Buffer.from([DH.join(','),
    'Customer 1,BBBU7654321,BL0001,="900100",10/5/2026,10/9/2026,DELIVERED,Terminal 1,Warehouse 1,10/6/2026,10/6/2026,10/7/2026,10/7/2026,,10/9/2026,',
    'Customer 1,CCCU1111111,BL0002,="900101-BCF",10/20/2026,10/24/2026,AVAILABLE,Terminal 1,Warehouse 1,,,,,,,'].join('\n') + '\n'));

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

  const src = k => ({ link: `https://example.invalid/files/${k}`, driveId: k.includes('dir') ? REP : SRC, itemId: k, name: '', webUrl: `https://example.invalid/files/${k}` });
  const settings = { version: 1, sources: { msl: src('itm-msl'), tracker: src('itm-trk'), ct: src('itm-ct'), forwarder: src('itm-fwd-dir'), drayage: src('itm-dray-dir') },
    admins: [ACCOUNT.username], errorEmailTo: MAIL_TO, updatedBy: ACCOUNT.name, updatedAt: iso(20) };
  bytes.set('settings.json', Buffer.from(JSON.stringify(settings)));

  // The PO Book page, served at /po-book/index.html: its CORE block as copied into core/pobook.js, or changed.
  const block = PB.coreText(readFileSync(join(REPO, 'core', 'pobook.js'), 'utf8'));
  const poBook = body => `<!doctype html><html><head><title>PO Book</title></head><body><script>\n/*CORE-START*/${body}/*CORE-END*/\n</script></body></html>`;
  return { bytes, pob, upcPages, poBookSame: poBook(block), poBookChanged: poBook(block + '\n/* a rule changed */'), blockMatches: block != null && sha(block) === PB.CORE_SHA256 };
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
  e.totalIssues = sum(Object.values(e).map(x => x.issues));
  e.sane = tc.ok && cm.ok && ct.ok && fwd.ok && dray.ok && upc.ok && e.msl.open !== e.msl.openPlain && e.totalIssues > 0;
  return e;
}

// ------------------------------------------------------------------ the stand-in for SharePoint (Graph and downloads)

let DL = '';          // the download host: another port on 127.0.0.1, so downloads are cross-origin like SharePoint's
let world = null;     // the SharePoint state of the scenario being run

function newWorld(fx, o = {}) {
  return { fx, o: { settings: true, mslFile: 'msl.xlsx', ctFile: 'ct.xlsx', fwdAge: 0, tracker: 200, poBook: 'same', mail: 202, ats: null, slowCt: false, ...o },
    folders: new Set(['Inventory', ...(o.ats ? ['Inventory/ATS'] : []), ...(o.errorsDir ? ['Inventory/Errors'] : [])]), markers: new Map(),
    calls: [], downloads: [], unanswered: [], pending: new Set(), aborted: [], poBookGets: 0 };
}

const J = (status, json) => ({ status, json });
const notFound = () => J(404, { error: { code: 'itemNotFound', message: 'The resource could not be found.' } });
const webUrl = id => `https://example.invalid/files/${id}`;
const by = { user: { displayName: 'Test Saver' } };
function fileMeta(id, name, fx, age) {
  return { id, name, webUrl: webUrl(id), lastModifiedDateTime: iso(age), lastModifiedBy: by, file: {}, eTag: `"${id}-1"`,
    '@microsoft.graph.downloadUrl': `${DL}/fx/${fx}` };
}
const folderMeta = (id, name) => ({ id, name, webUrl: webUrl(id), folder: { childCount: 3 }, lastModifiedDateTime: iso(1) });
const listMeta = (id, name) => ({ id, displayName: name, name, webUrl: webUrl(id), lastModifiedDateTime: iso(1), lastModifiedBy: by });
const kid = (id, name, age, extra) => ({ id, name, lastModifiedDateTime: iso(age), lastModifiedBy: by, webUrl: webUrl(id), file: {}, ...extra });

function graph(w, method, path, q, body) {
  const o = w.o;
  if (method === 'GET') {
    if (path === '/sites/met1sol.sharepoint.com:/sites/Production') return J(200, { id: SITE_ID });
    if (path === `/sites/${SITE_ID}/drive`) return J(200, { id: PROD });
    if (path === `/drives/${PROD}/root:/Inventory/settings.json`) {
      if (o.settings === 403) return J(403, { error: { code: 'accessDenied', message: 'Access denied.' } });
      return o.settings ? J(200, { id: 'itm-settings', name: 'settings.json', eTag: '"settings-1"', file: {}, '@microsoft.graph.downloadUrl': `${DL}/fx/settings.json` }) : notFound();
    }
    if (path === `/drives/${PROD}/root:/Inventory:/children`) {   // what the page checks when settings.json is missing
      return J(200, { value: [...w.folders].filter(f => /^Inventory\/[^/]+$/.test(f)).map(f => folderMeta('dir-' + f.replace(/\W/g, '-'), f.split('/')[1])) });
    }
    if (path === `/drives/${SRC}/items/itm-msl`) return J(200, fileMeta('itm-msl', 'Master Sheet Live.xlsx', o.mslFile, 1));
    if (path === `/drives/${SRC}/items/itm-trk`) {
      return o.tracker === 403 ? J(403, { error: { code: 'accessDenied', message: 'Access denied.' } }) : J(200, fileMeta('itm-trk', 'PO Status Tracker.xlsx', 'tracker.xlsx', 1));
    }
    if (path === `/drives/${SRC}/items/itm-ct`) return J(200, fileMeta('itm-ct', 'Container Tracking Sheet.xlsx', o.ctFile, 0));
    if (path === `/drives/${REP}/items/itm-fwd-dir`) return J(200, folderMeta('itm-fwd-dir', 'Forwarder'));
    if (path === `/drives/${REP}/items/itm-fwd-dir/children`) {
      // two pages; the newest file is on the second, so the page must follow @odata.nextLink
      if (q.get('$skiptoken') === 'p2') return J(200, { value: [kid('itm-fwd-new', 'forwarder-report-new.xlsx', o.fwdAge)] });
      return J(200, { value: [kid('itm-fwd-old', 'forwarder-report-old.xlsx', o.fwdAge + 3), kid('itm-fwd-lock', '~$forwarder-report-new.xlsx', 0),
        { id: 'itm-fwd-sub', name: 'Archive', folder: { childCount: 9 }, lastModifiedDateTime: iso(0) }, kid('itm-fwd-txt', 'read-me.txt', 0)],
      '@odata.nextLink': `${GRAPH}/drives/${REP}/items/itm-fwd-dir/children?$top=999&$skiptoken=p2` });
    }
    if (path === `/drives/${REP}/items/itm-fwd-new`) return J(200, fileMeta('itm-fwd-new', 'forwarder-report-new.xlsx', 'fwd.xlsx', o.fwdAge));
    if (path === `/drives/${REP}/items/itm-fwd-old`) return J(200, fileMeta('itm-fwd-old', 'forwarder-report-old.xlsx', 'fwd.xlsx', o.fwdAge + 3));
    if (path === `/drives/${REP}/items/itm-dray-dir`) return J(200, folderMeta('itm-dray-dir', 'Drayage'));
    if (path === `/drives/${REP}/items/itm-dray-dir/children`) {
      return J(200, { value: [kid('itm-dray-old', 'drayage-report-old.csv', 1), kid('itm-dray-new', 'drayage-report-new.csv', 0), kid('itm-dray-x', 'drayage-report-new.xlsx', 0)] });
    }
    if (path === `/drives/${REP}/items/itm-dray-new`) return J(200, fileMeta('itm-dray-new', 'drayage-report-new.csv', 'dray.csv', 0));
    if (path === `/sites/${SITE_ID}/lists/PO_Book_Lines`) return J(200, listMeta('lst-pob', 'PO_Book_Lines'));
    if (path === `/sites/${SITE_ID}/lists/lst-pob/items`) return J(200, { value: w.fx.pob });
    if (path === `/sites/${SITE_ID}/lists/UPC_MASTER`) return J(200, listMeta('lst-upc', 'UPC_MASTER'));
    if (path === `/sites/${SITE_ID}/lists/lst-upc/items`) {
      if (q.get('$skiptoken') === 'p2') return J(200, { value: w.fx.upcPages[1] });
      return J(200, { value: w.fx.upcPages[0], '@odata.nextLink': `${GRAPH}/sites/${SITE_ID}/lists/lst-upc/items?$expand=fields&$top=999&$skiptoken=p2` });
    }
  }
  if (path === `/drives/${PROD}/root/children` && method === 'POST') return makeFolder(w, '', body);
  const m = /^\/drives\/drv-prod\/root:\/(.+?)(?::\/(content|children))?$/.exec(path);
  if (m) {
    const [, p, op] = m, parent = p.split('/').slice(0, -1).join('/');
    if (method === 'GET' && !op && w.folders.has(p)) return J(200, folderMeta('dir-' + p.replace(/\W/g, '-'), p.split('/').pop()));
    if (method === 'GET' && op === 'children' && p === 'Inventory/ATS' && w.folders.has(p)) {
      const a = o.ats || {};
      return J(200, { value: [kid('itm-ats-1', 'ats-report-new.xlsx', a.age), kid('itm-ats-0', 'ats-report-older.xlsx', a.age + 7), kid('itm-ats-lock', '~$ats-report-new.xlsx', 0)] });
    }
    if (method === 'POST' && op === 'children') return makeFolder(w, p, body);
    if (method === 'PUT' && op === 'content' && parent === 'Inventory/Errors') {
      if (!w.folders.has(parent)) return notFound();   // path uploads don't create parent folders
      if (q.get('@microsoft.graph.conflictBehavior') === 'fail' && w.markers.has(p)) return J(409, { error: { code: 'nameAlreadyExists', message: 'The specified item name already exists.' } });
      w.markers.set(p, body);
      return J(201, { id: 'itm-marker-' + w.markers.size, name: p.split('/').pop(), eTag: '"m-1"' });
    }
    if (method === 'DELETE' && !op && w.markers.has(p)) { w.markers.delete(p); return { status: 204 }; }
    if (method === 'GET' && !op) return notFound();
  }
  if (method === 'POST' && path === '/me/sendMail') {
    return o.mail === 202 ? { status: 202 } : J(o.mail, { error: { code: 'ErrorAccessDenied', message: 'Access is denied. Check credentials and try again.' } });
  }
  w.unanswered.push(`${method} ${path}`);
  return notFound();
}
function makeFolder(w, parent, body) {
  const name = body && body.name, p = parent ? parent + '/' + name : name;
  if (!name || !(body.folder)) return J(400, { error: { code: 'invalidRequest', message: 'A folder needs a name.' } });
  if (parent && !w.folders.has(parent)) return notFound();
  if (w.folders.has(p)) return J(409, { error: { code: 'nameAlreadyExists', message: 'The specified item name already exists.' } });
  w.folders.add(p);
  return J(201, folderMeta('dir-' + p.replace(/\W/g, '-'), name));
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
        const r = w ? graph(w, req.method, gpath, q, body) : notFound();
        if (w) w.calls.push({ method: req.method, path: gpath, query: rawQ, auth: req.headers.authorization || '', body, status: r.status });
        res.writeHead(r.status, r.json ? { 'Content-Type': 'application/json' } : {});
        res.end(r.json ? JSON.stringify(r.json) : undefined);
      });
      return;
    }
    if (req.method !== 'GET') { res.writeHead(405); res.end(); return; }
    if (path === '/config.js') {
      res.writeHead(200, { 'Content-Type': TYPES['.js'], 'Cache-Control': 'no-store' });
      res.end(`window.INV_CONFIG = ${JSON.stringify(TEST_CONFIG)};\n`);
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

// Pre-signed download links: no sign-in needed, cross-origin with CORS open, and no Authorization header expected
// (a preflight asking to send one is refused, as SharePoint's download host would not need it).
function downloadServer() {
  return createServer((req, res) => {
    const cors = { 'Access-Control-Allow-Origin': '*' };
    if (req.method === 'OPTIONS') { res.writeHead(204, cors); res.end(); return; }
    const name = decodeURIComponent(req.url.split('?')[0]).replace(/^\/fx\//, '');
    const w = world;
    const buf = w && w.fx.bytes.get(name);
    if (w) w.downloads.push({ name, auth: !!req.headers.authorization });
    if (!buf) { res.writeHead(404, cors); res.end(); return; }
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

function initScript(pageOrigin) {
  const b64u = o => Buffer.from(JSON.stringify(o)).toString('base64url');
  const token = `${b64u({ alg: 'none', typ: 'JWT' })}.${b64u({ aud: 'https://graph.microsoft.com', scp: 'User.Read Sites.ReadWrite.All Mail.Send', name: ACCOUNT.name })}.headless`;
  return `(() => {
  if (window.__headless) return;
  const H = window.__headless = { blocked: [], msal: [], msalSetTried: 0, config: null };
  const GRAPH = ${JSON.stringify(GRAPH)}, PAGE = ${JSON.stringify(pageOrigin)}, DL = ${JSON.stringify(DL)};
  const realFetch = window.fetch.bind(window);
  // Graph goes to the local stand-in; the page and the download host pass; anything else is refused and recorded.
  window.fetch = function (input, init) {
    const url = typeof input === 'string' ? input : (input && input.url) || String(input);
    if (url === GRAPH || url.startsWith(GRAPH + '/')) return realFetch(PAGE + '/__graph' + url.slice(GRAPH.length), init);
    let o = '';
    try { o = new URL(url, location.href).origin; } catch (e) { /* not a URL */ }
    if (o === PAGE || o === DL) return realFetch(input, init);
    H.blocked.push(url);
    return Promise.reject(new TypeError('Failed to fetch'));
  };
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
      // A scenario sets H.tokenFail to make the quiet renewal fail the way MSAL's hidden-frame timeout does.
      if (H.tokenFail) { const e = new Error('Token renewal operation failed due to timeout.'); e.errorCode = H.tokenFail; throw e; }
      return { accessToken: ${JSON.stringify(token)}, account, scopes: (req && req.scopes) || [] };
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
  return {
    title: document.title, appShown: shown(q('#app')), signinShown: shown(q('#signin')), busy: q('#app') && q('#app').getAttribute('aria-busy'),
    busyShown: shown(q('#busy')), health: txt(q('#health')), chips, cards, srows, cardCount: document.querySelectorAll('#srcCards .scard').length,
    banners: [...document.querySelectorAll('#banners .banner')].map(b => ({ kind: (b.className.match(/\b(err|warn|info)\b/) || [])[1] || '', text: txt(b) })),
    cProb: txt(q('#cProb')), pqInfo: txt(q('#pqInfo')),
    issueGroups: issRows.filter(r => r.classList.contains('gh')).map(txt),
    issueRows: issRows.filter(r => !r.classList.contains('gh') && !r.querySelector('.more,.empty')).length,
    srcRows: document.querySelectorAll('#srcTable tr[id^="srow-"]').length,
    tabs: Object.fromEntries(['inv', 'pull', 'problems'].map(t => [t, shown(q('#tab-' + t))])),
    sourcesButton: shown(q('#btnSources')), who: txt(q('#who')),
    dlgOpen: !!(q('#dlg') && q('#dlg').open), dlgFields: [...document.querySelectorAll('#dlg input[id^="sf_"], #dlg textarea[id^="sf_"]')].map(x => x.id),
    dlgAdmins: (q('#sf_admins') || {}).value || '', dlgCancel: !!q('#dlg #dCancel'),
    toast: txt(q('#toast')), toastShown: shown(q('#toast')),
    msalFake: !!(window.msal && window.msal.headlessFake), msalSetTried: H.msalSetTried || 0, msalInteractive: H.msal || [],
    msalConfig: H.config, fetchBlocked: H.blocked || [], loading: [...document.querySelectorAll('.chip.src[data-state="load"]')].map(x => x.id),
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

async function openTab(cdp, init) {
  const { targetId } = await cdp.send('Target.createTarget', { url: 'about:blank' });
  const { sessionId } = await cdp.send('Target.attachToTarget', { targetId, flatten: true });
  const t = { targetId, sessionId, blocked: [], console: [], exceptions: [] };
  t.send = (m, p) => cdp.send(m, p, sessionId);
  t.off = cdp.on(m => {
    if (m.sessionId !== sessionId) return;
    const p = m.params || {};
    if (m.method === 'Fetch.requestPaused') {
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
  t.shot = async name => {
    const m = await t.send('Page.getLayoutMetrics');
    const s = m.cssContentSize || m.contentSize;
    const r = await t.send('Page.captureScreenshot', { format: 'png', captureBeyondViewport: true,
      clip: { x: 0, y: 0, width: Math.ceil(s.width), height: Math.min(Math.ceil(s.height), 6000), scale: 1 } });
    const file = join(OUT, `headless-${name}.png`);
    writeFileSync(file, Buffer.from(r.data, 'base64'));
    return file;
  };
  t.close = async () => { t.off(); await cdp.send('Target.closeTarget', { targetId }).catch(() => {}); };
  return t;
}

// ------------------------------------------------------------------ scenarios

const isWrite = c => c.method !== 'GET';
const markerPuts = calls => calls.filter(c => c.method === 'PUT' && c.path.startsWith(`/drives/${PROD}/root:/Inventory/Errors/`));
const sendMails = calls => calls.filter(c => c.method === 'POST' && c.path === '/me/sendMail');
const banner = (s, kind) => s.banners.filter(b => b.kind === kind).map(b => b.text).join(' | ');
const pair = (card, label) => { const p = card && card.pairs.find(x => x[0] === label); return p ? p[1] : null; };
const rc = e => `${fmtN(e.read)} read / ${fmtN(e.counted)} counted`;
const mailText = c => (c && c.body && c.body.message) || {};

function greenExcept(c, s, except) {
  for (const k of ['msl', 'tracker', 'pobook', 'ct', 'forwarder', 'drayage', 'upc'].filter(k => !except.includes(k))) {
    c(`${NAMES[k]} chip green`, s.chips[k] && s.chips[k].state === 'ok', s.chips[k] ? `${s.chips[k].state}: ${s.chips[k].text}` : 'no chip');
  }
}
function oneEmail(c, w, calls, source, problem, today) {
  const puts = markerPuts(calls), sends = sendMails(calls);
  const name = `Inventory/Errors/${C.isoDay(today)} ${source}.json`;
  c(`exactly one marker file created: "${name.split('/').pop()}"`, puts.length === 1 && puts[0].status === 201 && puts[0].path.endsWith(`root:/${name}:/content`),
    puts.map(p => `${p.status} ${p.path}`).join('; ') || 'none');
  c('the marker is created with conflictBehavior=fail', puts.length === 1 && /@microsoft\.graph\.conflictBehavior=fail/.test(decodeURIComponent(puts[0].query)), puts[0] && puts[0].query);
  c('exactly one sendMail', sends.length === 1, `${sends.length} sent`);
  c('the marker is written before the email', puts.length === 1 && sends.length === 1 && calls.indexOf(puts[0]) < calls.indexOf(sends[0]));
  const msg = mailText(sends[0]);
  c(`the subject names ${source} and the problem`, msg.subject && msg.subject.includes(`: ${source} — ${problem}`), msg.subject);
  c('sent to the error address, plain text, kept in Sent Items',
    sends[0] && msg.toRecipients && msg.toRecipients.length === 1 && msg.toRecipients[0].emailAddress.address === MAIL_TO &&
    msg.body && msg.body.contentType === 'Text' && sends[0].body.saveToSentItems === true, JSON.stringify(msg.toRecipients));
  return { puts, sends, msg };
}

const SCENARIOS = [
  {
    name: 'ok', title: 'everything fine', world: {},
    async run({ t, w, c, e }) {
      let s = await t.snap();
      c('signed in with the fake account (no real sign-in)', s.appShown && !s.signinShown && s.who.includes(ACCOUNT.name), s.who);
      greenExcept(c, s, []);
      c('ATS chip grey "no report yet"', s.chips.ats && s.chips.ats.state === 'off' && /no report yet/.test(s.chips.ats.text), s.chips.ats && s.chips.ats.text);
      for (const k of ['msl', 'tracker', 'pobook', 'ct', 'forwarder', 'drayage', 'upc']) {
        c(`${NAMES[k]} chip shows "${rc(e[k])}"`, s.chips[k] && s.chips[k].text.includes(rc(e[k])), s.chips[k] && s.chips[k].text);
      }
      c('chips show who saved the file', s.chips.ct.text.includes('Test Saver'), s.chips.ct.text);
      c('Container Tracking chip links to the file', s.chips.ct.href === webUrl('itm-ct'), s.chips.ct.href);
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
      c('no email: nothing written to SharePoint', !w.calls.some(isWrite), w.calls.filter(isWrite).map(x => `${x.method} ${x.path}`).join('; '));
      c(`Problems count = every issue the parsers found (${e.totalIssues})`, s.cProb === String(e.totalIssues), s.cProb);
      c('no console errors', t.console.length === 0, t.console.join(' | '));
      c('Inventory tab: a card per source', s.tabs.inv && s.cardCount === 8, `${s.cardCount} cards`);
      t.shots.push(await t.shot('ok'));

      await t.click('#tabs .tab[data-tab="pull"]');
      s = await t.snap();
      c('Pull from China tab draws', s.tabs.pull && !s.tabs.inv && !s.tabs.problems, JSON.stringify(s.tabs));
      await t.click('#tabs .tab[data-tab="problems"]');
      s = await t.snap();
      c('Problems tab: sources table with 8 rows', s.tabs.problems && s.srcRows === 8, `${s.srcRows} rows`);
      c(`Problems tab: every issue listed (${e.totalIssues})`, s.issueRows === e.totalIssues, `${s.issueRows} rows`);
      const groups = ['tracker', 'msl', 'pobook', 'ct', 'forwarder', 'drayage', 'upc'].filter(k => e[k].issues > 0).length;
      c(`Problems tab: issues grouped by source (${groups} groups) with counts in plain words, not codes`,
        s.issueGroups.length === groups && s.issueGroups.some(g => /^UPC_MASTER · \d+ issues?/.test(g) && /UPC not 12-13 digits \d/.test(g)) &&
        !s.issueGroups.some(g => /\bbadupc\b|\bmslonly\b|\bdupround\b/.test(g)), s.issueGroups.join(' | '));
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
    name: 'ctbad', title: 'Container Tracking has no "Quantity" column', world: { ctFile: 'ct-bad.xlsx', poBook: 'missing' },
    async run({ t, w, c, today }) {
      const s = await t.snap();
      const ct = s.chips.ct;
      c('Container Tracking chip red', ct && ct.state === 'err', ct && `${ct.state}: ${ct.text}`);
      c('chip says "missing column" and names "Quantity"', /missing column/.test(ct.text) && /"Quantity"/.test(ct.title), ct.title);
      // Container Tracking puts the headers seen in its message; the PO Book group adds a "Headers seen:" line.
      c('Problems lists the headers seen', /headers seen[^]*Pieces/i.test(s.srows.ct), s.srows.ct);
      const red = banner(s, 'err');
      c('red banner: stage numbers not loaded, naming Container Tracking and "Quantity"', /stage numbers are not loaded/i.test(red) && red.includes('Container Tracking') && red.includes('Quantity'), red);
      greenExcept(c, s, ['ct']);
      c("PO Book lines chip says nothing when the PO Book page can't be fetched", s.chips.pobook.state === 'ok' && w.poBookGets >= 1, s.chips.pobook.text);
      const folderMade = w.calls.findIndex(x => x.method === 'POST' && x.path === `/drives/${PROD}/root:/Inventory:/children` && x.status === 201);
      const { puts, msg } = oneEmail(c, w, w.calls, 'Container Tracking', 'missing column "Quantity"', today);
      c('Inventory/Errors made before the marker (path uploads make no folders)', folderMade >= 0 && puts.length && folderMade < w.calls.indexOf(puts[0]), `folder POST at ${folderMade}`);
      const body = (msg.body && msg.body.content) || '';
      c('email body: file, problem, headers seen, who had the page open, link',
        body.includes('Container Tracking Sheet.xlsx') && body.includes('Quantity') && /headers seen[^]*Pieces/i.test(body) &&
        body.includes(ACCOUNT.name) && body.includes(ACCOUNT.username) && body.includes(webUrl('itm-ct')), body);
      c('page says "Saul has been emailed about this."', red.includes('Saul has been emailed about this.') && s.srows.ct.includes('Saul has been emailed about this.'), red);
      t.shots.push(await t.shot('ctbad'));
    },
  },
  {
    name: 'reload', title: 'same day, reload: the marker answers 409', after: 'ctbad',
    async run({ t, w, c }) {
      const from = w.calls.length;
      await t.eval('window.__stale = 1');
      await t.send('Page.reload', { ignoreCache: true });
      await t.waitFor(READY, 'the page to load again');
      const s = await t.snap(), calls = w.calls.slice(from);
      c('Container Tracking chip still red', s.chips.ct.state === 'err', s.chips.ct.text);
      const puts = markerPuts(calls);
      c('one marker create tried, refused with 409 (already sent today)', puts.length === 1 && puts[0].status === 409, puts.map(p => p.status).join(','));
      c('no second sendMail', sendMails(calls).length === 0, `${sendMails(calls).length} sent`);
      c('one email in total for the day', sendMails(w.calls).length === 1, `${sendMails(w.calls).length} sent`);
      c('Inventory/Errors not made again', !calls.some(x => x.method === 'POST' && x.path.endsWith(':/children')));
      c('page says Saul was already emailed today', banner(s, 'err').includes('Saul was already emailed about Container Tracking today.'), banner(s, 'err'));
      t.shots.push(await t.shot('reload'));
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
      greenExcept(c, s, ['msl']);
      const { msg } = oneEmail(c, w, w.calls, 'Master Sheet Live', 'missing column "Quantity"', today);
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
      c('amber banner: forwarder not used, stages from Container Tracking dates only', amber.includes('Forwarder report') && /Container Tracking dates only/.test(amber), amber);
      c('no red banner', !banner(s, 'err'), banner(s, 'err'));
      c('the old file is not opened or downloaded', !w.calls.some(x => /itm-fwd-(new|old)$/.test(x.path)) && !w.downloads.some(d => d.name === 'fwd.xlsx'),
        w.downloads.map(d => d.name).join(','));
      greenExcept(c, s, ['forwarder']);
      oneEmail(c, w, w.calls, 'Forwarder', 'not used', today);
      t.shots.push(await t.shot('fwdold'));
    },
  },
  {
    name: 'atsold', title: 'the newest ATS report is 12 days old', world: { ats: { age: 12 } },
    async run({ t, w, c, today }) {
      const s = await t.snap(), a = s.chips.ats;
      c('ATS chip amber "12 days old"', a.state === 'warn' && /12 days old/.test(a.text), `${a.state}: ${a.text}`);
      c('ATS chip names the newest report, lock file skipped', a.title.includes('ats-report-new.xlsx') && !a.title.includes('~$'), a.title);
      c('ATS report is not downloaded in step 1', !w.downloads.some(d => /ats/.test(d.name)));
      c('health line names ATS', s.health.includes('ATS'), s.health);
      greenExcept(c, s, []);
      oneEmail(c, w, w.calls, 'ATS', 'the newest report is 12 days old', today);
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
      greenExcept(c, s, ['tracker', 'msl']);
      const { msg } = oneEmail(c, w, w.calls, 'PO Status Tracker', 'no access', today);
      const body = (msg.body && msg.body.content) || '';
      c('email written for Saul: whose account, which file, what fixes it, what it holds up',
        body.includes(`${ACCOUNT.name}'s account can't open`) && /What fixes it: Give Test User access/.test(body) &&
        /Until then: The stage numbers stay off/.test(body) && !/Your account/.test(body) && /^Shalam Tool: PO Status Tracker — no access$/.test(msg.subject), body);
      t.shots.push(await t.shot('noaccess'));
    },
  },
  {
    name: 'mailfail', title: 'the email itself fails (sendMail 403)', world: { ctFile: 'ct-bad.xlsx', mail: 403 },
    async run({ t, w, c }) {
      const s = await t.snap();
      c('Container Tracking chip red and says the email failed', s.chips.ct.state === 'err' && /email failed/.test(s.chips.ct.text), s.chips.ct.text);
      c('Problems row says emailing Saul failed', /Emailing Saul failed/.test(s.srows.ct), s.srows.ct);
      c('it blames the mailbox, not SharePoint', /Emailing Saul failed: this account's mailbox didn't allow sending \(403/.test(s.srows.ct) &&
        !/Emailing Saul failed: Your account can't open it on SharePoint/.test(s.srows.ct), s.srows.ct);
      c('one sendMail tried', sendMails(w.calls).length === 1, `${sendMails(w.calls).length}`);
      const del = w.calls.filter(x => x.method === 'DELETE');
      c('the marker is taken back so a later visit can retry', del.length === 1 && del[0].status === 204 && w.markers.size === 0, del.map(x => `${x.status} ${x.path}`).join('; '));
    },
  },
  {
    name: 'pobook-changed', title: "the PO Book's rules changed since the copy", world: { poBook: 'changed' },
    async run({ t, w, c }) {
      const s = await t.snap(), p = s.chips.pobook;
      c('PO Book lines chip amber "PO Book rules changed"', p.state === 'warn' && /PO Book rules changed/.test(p.text), `${p.state}: ${p.text}`);
      c('no email for it', sendMails(w.calls).length === 0 && markerPuts(w.calls).length === 0);
      c('no red banner', !banner(s, 'err'), banner(s, 'err'));
      greenExcept(c, s, ['pobook']);
    },
  },
  {
    name: 'setup', title: 'no settings.json yet: the Sources popup', world: { settings: false },
    async run({ t, w, c }) {
      let s = await t.snap();
      c('every chip grey "not set up"', KEYS.every(k => s.chips[k] && s.chips[k].state === 'off' && /not set up/.test(s.chips[k].text)), KEYS.map(k => s.chips[k] && s.chips[k].state).join(','));
      c('info banner asks for the sources to be set up', /set up/.test(banner(s, 'info')), banner(s, 'info'));
      c('⚙ Sources shown to the first visitor', s.sourcesButton);
      await t.click('#btnSources');
      await t.waitFor(`document.getElementById('dlg').open`, 'the Sources popup');
      s = await t.snap();
      const want = ['sf_msl', 'sf_tracker', 'sf_ct', 'sf_forwarder', 'sf_drayage', 'sf_ats', 'sf_mail', 'sf_admins'];
      c('Sources popup has a field per source, the error email and the admins', want.every(id => s.dlgFields.includes(id)), s.dlgFields.join(','));
      c('your own address is filled in as the first admin', s.dlgAdmins.includes(ACCOUNT.username), s.dlgAdmins);
      c('the popup has Cancel', s.dlgCancel);
      t.shots.push(await t.shot('setup-sources'));
      await t.click('#dCancel');
      await t.waitFor(`!document.getElementById('dlg').open`, 'the popup to close', 5000).catch(() => {});
      s = await t.snap();
      c('Cancel closes the popup', !s.dlgOpen);
      c('nothing written to SharePoint', !w.calls.some(isWrite), w.calls.filter(isWrite).map(x => `${x.method} ${x.path}`).join('; '));
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
      const notReached = ['forwarder', 'drayage', 'ats', 'upc'].filter(k => /not set up/.test(s.chips[k].text));
      c("sources not read before the Cancel don't claim to be \"not set up\"", !notReached.length,
        notReached.map(k => `${NAMES[k]}: ${s.chips[k].state} "${s.chips[k].text}"`).join('; '));
      c('they say they were not read because of the Cancel', ['forwarder', 'drayage', 'ats', 'upc'].every(k => /not read \(cancelled\)/.test(s.chips[k].text)),
        ['forwarder', 'drayage', 'ats', 'upc'].map(k => s.chips[k].text).join(' | '));
      c("the toast doesn't claim older data is shown on a first load", !/loaded before|read before/i.test(s.toast), s.toast);
      c('the health line names what was not read', /Not read this time/.test(s.health) && s.health.includes('Container Tracking'), s.health);
      c('no email after a cancel', !sendMails(w.calls).length && !markerPuts(w.calls).length);
      t.shots.push(await t.shot('cancel'));
    },
  },
  {
    name: 'settings403', title: "the sources file can't be opened (403) on the first load", world: { settings: 403 },
    async run({ t, w, c }) {
      const s = await t.snap();
      const red = banner(s, 'err');
      c("red banner: couldn't load the sources, no access to the Inventory folder", /Couldn't load the sources/.test(red) && /can't open the Inventory folder/.test(red), red);
      const unset = KEYS.filter(k => /not set up/.test(s.chips[k].text));
      c('no chip claims "not set up"', !unset.length, unset.map(k => `${NAMES[k]}: "${s.chips[k].text}"`).join('; '));
      c('every chip grey "not read"', KEYS.every(k => s.chips[k].state === 'off' && /not read/.test(s.chips[k].text)), KEYS.map(k => s.chips[k].text).join(' | '));
      c('⚙ Sources not offered', !s.sourcesButton);
      c('nothing written to SharePoint', !w.calls.some(isWrite), w.calls.filter(isWrite).map(x => `${x.method} ${x.path}`).join('; '));
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
      c('no email tried', !sendMails(w.calls).length && !markerPuts(w.calls).length);
      t.shots.push(await t.shot('authfail'));
    },
  },
  {
    name: 'nosettings', title: 'settings.json is gone but the Inventory folder is in use', world: { settings: false, errorsDir: true },
    async run({ t, w, c }) {
      const s = await t.snap();
      const red = banner(s, 'err');
      c('red banner: the sources file is missing', /settings\.json is missing/.test(red), red);
      c('no info banner inviting a new first admin', !banner(s, 'info'), banner(s, 'info'));
      c('⚙ Sources not offered to whoever opens the page', !s.sourcesButton);
      c('health line says the sources file is missing', /Sources file missing/.test(s.health), s.health);
      c('nothing written to SharePoint', !w.calls.some(isWrite), w.calls.filter(isWrite).map(x => `${x.method} ${x.path}`).join('; '));
    },
  },
];

// ------------------------------------------------------------------ main

async function main() {
  const pick = (process.argv[2] || '').split(',').map(x => x.trim()).filter(Boolean);
  const unknown = pick.filter(n => !SCENARIOS.some(s => s.name === n));
  if (unknown.length) { console.log(`Unknown scenario: ${unknown.join(', ')}. Known: ${SCENARIOS.map(s => s.name).join(', ')}`); process.exit(2); }
  const want = new Set(pick.length ? pick : SCENARIOS.map(s => s.name));
  for (const s of SCENARIOS) if (want.has(s.name) && s.after) want.add(s.after);
  const run = SCENARIOS.filter(s => want.has(s.name));

  mkdirSync(OUT, { recursive: true });
  const fx = buildFixtures();
  const today = C.todayNY(Date.now());
  const e = expected(fx, today);
  const results = [];
  const setup = { name: 'setup checks', title: 'fixtures and harness', checks: [] };
  const sc = (label, ok, detail) => setup.checks.push({ label, ok: !!ok, detail: ok ? '' : String(detail ?? '') });
  sc('core/pobook.js block hashes to CORE_SHA256 (the fake PO Book page serves it)', fx.blockMatches);
  sc('the invented files parse and exercise the PO Book marks and issues', e.sane, JSON.stringify(e));
  results.push(setup);

  const page = pageServer(), dl = downloadServer();
  await new Promise(r => page.listen(0, '127.0.0.1', r));
  await new Promise(r => dl.listen(0, '127.0.0.1', r));
  const ORIGIN = `http://127.0.0.1:${page.address().port}`;
  DL = `http://127.0.0.1:${dl.address().port}`;

  let edge = null;
  try {
    edge = await launchEdge();
    console.log(`Edge ${edge.version}, headless, InPrivate, sync off · page ${ORIGIN} · downloads ${DL}`);
    const early = await signInWindows(edge.cdp);
    sc('Edge opened no sign-in or sync window', !early.length, early.join(', '));
    const init = initScript(ORIGIN);
    let t = null, w = null;
    for (const s of run) {
      const checks = [];
      const c = (label, ok, detail) => checks.push({ label, ok: !!ok, detail: ok ? '' : String(detail ?? '') });
      const r = { name: s.name, title: s.title, checks };
      results.push(r);
      try {
        if (!s.after) {
          if (t) await t.close();
          w = world = newWorld(fx, s.world);
          t = await openTab(edge.cdp, init);
          t.shots = [];
          await t.send('Page.navigate', { url: `${ORIGIN}/index.html` });
          if (s.wait !== 'busy') await t.waitFor(READY, 'the sources to load');
        }
        const before = { blocked: t.blocked.length, exceptions: t.exceptions.length, calls: s.after ? w.calls.length : 0, downloads: s.after ? w.downloads.length : 0 };
        await s.run({ t, w, c, e, today });
        // Checks every scenario gets.
        const snap = await t.snap();
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
        const calls = w.calls.slice(before.calls), dls = w.downloads.slice(before.downloads);
        c('every Graph call carried a bearer token', calls.length > 0 && calls.every(x => /^Bearer \S+/.test(x.auth)), `${calls.length} calls`);
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
    // The guards themselves: a request to another host must be refused by both. 127.0.0.2 is still this machine, so
    // nothing leaves it even if a guard were broken.
    world = newWorld(fx, { settings: false });
    t = await openTab(edge.cdp, init);
    await t.send('Page.navigate', { url: `${ORIGIN}/index.html` });
    await t.waitFor(READY, 'the page to load');
    await t.eval(`(() => { new Image().src = 'http://127.0.0.2:9/guard-probe.png'; fetch('http://127.0.0.2:9/guard-probe').catch(() => {}); })()`);
    await sleep(800);
    const probe = await t.snap();
    sc('the request guard refuses a page request to another host', t.blocked.some(u => u.includes('guard-probe.png')), t.blocked.join(', ') || 'nothing refused');
    sc('the fake fetch refuses another host', probe.fetchBlocked.some(u => u.includes('guard-probe')), probe.fetchBlocked.join(', ') || 'nothing refused');
    await t.close();
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
