// Runs the page's parsers, the stage engine and the selling engine on local copies of the real source files and prints
// only their structure and numbers: each header row, the counts, issue codes with up to 3 row numbers, the Incoming
// totals and balance, Available to sell, and whether core/pobook.js still matches the PO Book. Nothing else from the
// files is printed (no cell values, no issue messages); style codes and PO numbers appear next to numbers, nothing more.
//   node tools/check-real.mjs <folder> [po-book index.html] [--lists <lists.json>] [--ats <ERP report.xlsx>]
//                             [--today yyyy-mm-dd] [--no-reports | --no-forwarder | --no-drayage]
// <folder> holds Master Sheet Live.xlsx, PO Status Tracker.xlsx, Container Tracking Sheet.xlsx, Forwarder.xlsx and
// Drayage.csv. The PO Book file defaults to $PO_BOOK_HTML, then ../po-book/index.html next to this repo.
// --lists: the SharePoint lists the page reads, saved as JSON { UPC_MASTER: [items], PO_Book_Lines: [items] } (Graph
// list items with their fields); defaults to <folder>/lists.json when it's there. Without it the PO Book's marks and
// UPC_MASTER are left out, so open/shipped and descriptions can differ from the page.
// --ats: the ERP's "Inventory Report - On Hand_In Transit_In Production" (sheet "Report"), read with core/ats.js as the
// page reads it (its date from the file name, else the day the file was saved, as on the page), then Available to sell
// (core/sell.js) with it: Arriving from its In Trans, and how our own lots meet it (matched, received, in the ERP's
// transit, landed, kept), reconciled per style with Incoming and the report's In Trans (a failure if they don't add up);
// and compared per style with Incoming: In Trans Qty vs On the water + At port, Import PO
// Backlog Qty vs In China (LA POs only). Available to sell is also worked out without the report (as when it's missing),
// with or without --ats. Its sums, Saul's pull-from date rule and the ready dates are checked lot by lot here, on their
// own, not with the engine's code.
// --no-reports: the stages without the forwarder and drayage reports, as the page shows them when both are missing;
// --no-forwarder / --no-drayage: without that one report only (as when it is too old, or fails to read).
// The Container Tracking section lists come from config.js (run with a stand-in window), as the page passes them;
// --no-sections leaves them out (sections with no round all Booked/Loaded; POs below the tracker's lowest still apart).
// Exit code 1 when a file is missing or fails to parse (the ATS report too), counts don't add up, the balance is off,
// Available to sell doesn't add up or breaks a rule, or the PO Book's rules changed. Differences from the ERP are
// printed, never a failure.
import { readFileSync, existsSync, statSync } from 'node:fs';
import { createHash } from 'node:crypto';
import { join, resolve, basename } from 'node:path';
import { loadXLSX, core, root } from '../tests/load.mjs';

const XOPT = { type: 'array', cellFormula: false, cellHTML: false, cellStyles: false };   // as app.js reads files
const FILES = { msl: 'Master Sheet Live.xlsx', tracker: 'PO Status Tracker.xlsx', ct: 'Container Tracking Sheet.xlsx',
  forwarder: 'Forwarder.xlsx', drayage: 'Drayage.csv' };
const USAGE = 'Usage: node tools/check-real.mjs <folder> [po-book index.html] [--lists <lists.json>] [--ats <ERP report.xlsx>] [--today yyyy-mm-dd] [--no-reports | --no-forwarder | --no-drayage] [--no-sections]';

// Options take their value after "=" or as the next argument (--ats <file>); the rest are the two paths.
const VALUED = ['today', 'lists', 'ats'], FLAGS = ['no-reports', 'no-forwarder', 'no-drayage', 'no-sections', 'help'];
const opt = {}, pos = [];
for (let i = 2, a = process.argv; i < a.length; i++) {
  if (!a[i].startsWith('--')) { pos.push(a[i]); continue; }
  const [k, v] = [a[i].slice(2).split('=')[0], a[i].includes('=') ? a[i].slice(a[i].indexOf('=') + 1) : undefined];
  if (VALUED.includes(k)) {
    const val = v ?? a[++i];
    if (val == null || val.startsWith('--')) { console.log(`--${k} needs a value.\n${USAGE}`); process.exit(2); }
    opt[k] = val;
  } else if (FLAGS.includes(k)) opt[k] = true;
  else { console.log(`Unknown option --${k}.\n${USAGE}`); process.exit(2); }
}
const [folderArg, poBookArg] = pos;
if (!folderArg || opt.help || pos.length > 2) {
  console.log(USAGE);
  process.exit(folderArg && opt.help ? 0 : 2);
}
const folder = resolve(folderArg);

const XLSX = loadXLSX();
const C = core('common'), PB = core('pobook'), CT = core('ct'), SH = core('shipping'), UPC = core('upc'), ST = core('stages');
const ATS = core('ats'), SELL = core('sell');
const today = opt.today ? C.dayFromIso(opt.today) : C.todayNY(Date.now());
if (today == null) { console.log(`--today ${opt.today} isn't a yyyy-mm-dd date.`); process.exit(2); }
const listsPath = opt.lists ? resolve(opt.lists) : existsSync(join(folder, 'lists.json')) ? join(folder, 'lists.json') : null;
const atsPath = opt.ats ? resolve(opt.ats) : null;

let failures = 0;
const fails = msg => { failures++; console.log(`  FAILED: ${msg}`); };
const n = v => (typeof v === 'number' && Number.isFinite(v) ? v.toLocaleString('en-US') : String(v));
const secs = t0 => ((performance.now() - t0) / 1000).toFixed(2) + ' s';
const head = t => console.log(`\n== ${t}`);

// A header row as "A Name | B Name | K (blank) | …", up to its last named cell.
function headerLine(row) {
  const cells = (row || []).map(v => C.S(v));
  let last = cells.length - 1;
  while (last >= 0 && !cells[last]) last--;
  return cells.slice(0, last + 1).map((h, i) => `${C.colLetter(i)} ${h || '(blank)'}`).join(' | ') || '(empty)';
}
// "rows read 10 = a 4 + b 6   adds up: yes", failing when the parts don't sum to the total.
function sumLine(label, total, parts) {
  const sum = parts.reduce((s, [, v]) => s + (v || 0), 0);
  const okSum = sum === total;
  console.log(`  ${label} ${n(total)} = ${parts.map(([k, v]) => `${k} ${n(v)}`).join(' + ')}   adds up: ${okSum ? 'yes' : 'NO'}`);
  if (!okSum) fails(`${label}: the parts add up to ${n(sum)}, not ${n(total)}`);
}
function otherCounts(counts, skip) {
  const rest = Object.entries(counts || {}).filter(([k, v]) => !skip.includes(k) && typeof v === 'number');
  if (rest.length) console.log(`  ${rest.map(([k, v]) => `${k} ${n(v)}`).join(' · ')}`);
}
// Issue codes with their count and up to 3 places (sheet and row only, never the message).
function issueLines(issues, where) {
  const by = new Map();
  for (const i of issues || []) { const k = i.code || '(no code)'; (by.get(k) || by.set(k, []).get(k)).push(i); }
  console.log(`  issues: ${n((issues || []).length)}`);
  for (const [code, list] of [...by].sort((a, b) => b[1].length - a[1].length || a[0].localeCompare(b[0]))) {
    const ex = [...new Set(list.map(where).filter(Boolean))].slice(0, 3);
    console.log(`    ${code.padEnd(24)} ${n(list.length).padStart(6)}${ex.length ? '   e.g. ' + ex.join(', ') : ''}`);
  }
}
const rowOf = i => (i.row == null || i.row === '' ? '' : `row ${i.row}`);
const sheetRowOf = i => [i.sheet, i.row == null || i.row === '' ? '(no row)' : `row ${i.row}`].filter(Boolean).join(' ');

function bytesOf(k) {
  const p = join(folder, FILES[k]);
  if (!existsSync(p)) { fails(`${FILES[k]} isn't in ${folder}`); return null; }
  const buf = readFileSync(p);
  if (!buf.length) { fails(`${FILES[k]} is empty (0 bytes)`); return null; }
  return new Uint8Array(buf.buffer, buf.byteOffset, buf.byteLength);
}
function readBook(bytes, extra) { return XLSX.read(bytes, { ...XOPT, ...(extra || {}) }); }
// config.js run with a stand-in window, as the page runs it; its `sections` (undefined when it has none).
function configSections() {
  const win = {};
  new Function('window', readFileSync(root('config.js'), 'utf8'))(win);
  return win.INV_CONFIG ? win.INV_CONFIG.sections : undefined;
}
// One sheet's rows from A1, as app.js reads Master Sheet Live and the tracker.
function readSheet(bytes, sheet) {
  let wb = readBook(bytes, { sheets: [sheet] }), hit = C.findSheet(wb, sheet);
  if (!hit || !hit.ws) {
    const names = readBook(bytes, { bookSheets: true }).SheetNames || [];
    const real = names.find(x => C.S(x).toLowerCase() === sheet.toLowerCase());
    if (!real) throw new Error(`no sheet named "${sheet}" (sheets: ${names.join(', ') || 'none'})`);
    wb = readBook(bytes, { sheets: [real] }); hit = C.findSheet(wb, sheet);
  }
  return { rows: C.sheetRows(XLSX, hit.ws), ref: hit.ws['!ref'] || '' };
}

console.log(`Shalam Tool real-file check · today ${C.fmtDay(today)} (New York) · ${folder}`);
console.log('Prints header names, counts, issue codes and row numbers only.');

// ---------------- PO Status Tracker and Master Sheet Live, read as the PO Book reads them
let trk = null, trkonly = [];
head(`PO Status Tracker (${FILES.tracker}, sheet "PO Status Tracker")`);
const tb = bytesOf('tracker');
let trows = null;
if (tb) {
  try {
    const t0 = performance.now();
    const r = readSheet(tb, 'PO Status Tracker'); trows = r.rows;
    console.log(`  sheet range ${r.ref}; read in ${secs(t0)}`);
    console.log(`  header row 1: ${headerLine(trows[0])}`);
    const tc = PB.countTracker(trows);
    if (tc.missing.length) fails(`missing column${tc.missing.length > 1 ? 's' : ''} ${tc.missing.map(x => `"${x}"`).join(', ')}`);
    else {
      sumLine('rows read', tc.rowsRead, Object.entries(tc.counts));
      trk = PB.parseTracker(trows);
      console.log(`  parseTracker: ${n(trk.all.length)} tracked lines (${n(trk.T.size)} PO+style keys with -F aliases)`);
      if (trk.all.length !== tc.counts.tracked) fails(`parseTracker read ${n(trk.all.length)} lines but countTracker counted ${n(tc.counts.tracked)}`);
      issueLines(tc.issues, rowOf);
      if (!tc.ok) fails(tc.error);
      else console.log('  result: ok');
    }
  } catch (e) { fails(`can't read it: ${e.message}`); }
}

head(`Master Sheet Live (${FILES.msl}, sheet "Master Sheet Live")`);
const mb = bytesOf('msl');
let mrows = null, msrc = null, mcounts = null;   // kept for the stage engine below
if (mb) {
  try {
    const t0 = performance.now();
    const r = readSheet(mb, 'Master Sheet Live');
    mrows = r.rows;
    console.log(`  sheet range ${r.ref}; read in ${secs(t0)}`);
    console.log(`  header row 1: ${headerLine(mrows[0])}`);
    if (!trk) fails('not counted: it is read together with the PO Status Tracker, which failed');
    else {
      const cm = PB.countMaster(mrows, trk);
      if (cm.missing.length) fails(`missing column${cm.missing.length > 1 ? 's' : ''} ${cm.missing.map(x => `"${x}"`).join(', ')}`);
      else {
        sumLine('rows read', cm.rowsRead, Object.entries(cm.counts));
        const src = PB.parseMaster(mrows, trk);
        const lines = PB.applyOverlay(src, new Map(), PB.cfgFrom([]));
        const open = lines.filter(l => l.status === 'open').length;
        console.log(`  parseMaster: ${n(src.lines.length)} lines (countMaster counted ${n(cm.counts.line)}: ${src.lines.length === cm.counts.line ? 'same' : 'DIFFERENT'})`);
        if (src.lines.length !== cm.counts.line) fails('parseMaster and countMaster disagree on the number of lines');
        console.log(`  without the PO Book's marks: open ${n(open)}, shipped ${n(lines.length - open)}`);
        // As the page lists them: countMaster's issues (with rows), then parseMaster's own except 'nostyle' (a repeat
        // of countMaster's) and 'trkonly' (listed under the tracker).
        const own = src.issues.filter(i => i.code !== 'nostyle' && i.code !== 'trkonly');
        trkonly = src.issues.filter(i => i.code === 'trkonly');
        // parseMaster's issues carry no row field; 'dupround' names its row in the text.
        const rowIn = i => { const m = /\brow (\d+)/.exec(i.issue || ''); return m ? `row ${m[1]}` : '(no row)'; };
        issueLines([...cm.issues, ...own], i => (i.row ? `row ${i.row}` : rowIn(i)));
        if (!cm.ok || !src.lines.length) fails(cm.error || 'no lines');
        else { console.log('  result: ok'); msrc = src; mcounts = cm.counts; }
      }
    }
  } catch (e) { fails(`can't read it: ${e.message}`); }
}
if (trkonly.length) console.log(`\n  PO Status Tracker, also listed by the PO Book: trkonly ${n(trkonly.length)} (no row numbers)`);

// ---------------- Container Tracking
head(`Container Tracking (${FILES.ct})`);
const cb = bytesOf('ct');
const used = { ct: null, forwarder: null, drayage: null, upc: null };   // each parse result the page would use, or null
if (cb) {
  try {
    const t0 = performance.now();
    const wb = readBook(cb);
    console.log(`  tabs: ${wb.SheetNames.join(' | ')}`);
    for (const tab of ['Booking', 'Loaded', 'Open', 'Complete']) {
      const hit = C.findSheet(wb, tab);
      console.log(`  ${tab} header row 1: ${hit ? headerLine(C.sheetRows(XLSX, hit.ws)[0]) : '(no such tab)'}`);
    }
    const r = CT.parseContainerTracking(XLSX, wb, today);
    console.log(`  read in ${secs(t0)}`);
    const c = r.counts || {};
    for (const tab of ['Booking', 'Loaded']) {
      const x = c[tab];
      if (!x) continue;
      sumLine(`${tab} rows read`, x.rowsRead, [['blank', x.blankRows], ['section', x.sectionRows], ['style', x.styleRows], ['other', x.otherRows]]);
      sumLine(`${tab} style rows`, x.styleRows, [['counted', x.counted], ['not counted', x.notCounted]]);
      otherCounts(x, ['rowsRead', 'blankRows', 'sectionRows', 'styleRows', 'otherRows', 'counted', 'notCounted']);
    }
    for (const tab of ['Open', 'Complete']) {
      const x = c[tab];
      if (!x) continue;
      sumLine(`${tab} rows read`, x.rowsRead, [['blank', x.blankRows], ['label', x.labelRows], ['round', x.roundRows], ['no round', x.noRoundRows]]);
      otherCounts(x, ['rowsRead', 'blankRows', 'labelRows', 'roundRows', 'noRoundRows']);
    }
    const blocks = r.blocks || [];
    const inRound = blocks.filter(b => b.rounds && b.rounds.length).length;
    console.log(`  PO blocks ${n(blocks.length)} (${n(inRound)} in a round) · lines ${n(blocks.reduce((s, b) => s + b.lines.length, 0))}`);
    issueLines(r.issues, sheetRowOf);
    if (!r.ok) fails(r.error);
    else { console.log('  result: ok'); used.ct = r; }
  } catch (e) { fails(`can't read it: ${e.message}`); }
}

// ---------------- Forwarder and drayage
head(`Forwarder (${FILES.forwarder})`);
const fb = bytesOf('forwarder');
if (fb) {
  try {
    const t0 = performance.now();
    const wb = readBook(fb);
    const r = SH.parseForwarder(XLSX, wb, today);
    console.log(`  first sheet "${wb.SheetNames[0]}"; read in ${secs(t0)}`);
    if (r.headerRow) console.log(`  header row ${r.headerRow}: ${headerLine(C.sheetRows(XLSX, wb.Sheets[wb.SheetNames[0]])[r.headerRow - 1])}`);
    const x = r.counts || {};
    if (r.headerRow) {
      sumLine('rows read', x.rowsRead, [['blank', x.blank], ['no ID', x.noId], ['records', x.records]]);
      console.log(`  real containers ${n(x.containers)} · shipments with no container ${n(x.noContainer)} · byContainer keys ${n(r.byContainer.size)} · unknown statuses ${n(r.unknownStatuses.length)}`);
    }
    issueLines(r.issues, rowOf);
    if (!r.ok) fails(r.error);
    else { console.log('  result: ok'); used.forwarder = r; }
  } catch (e) { fails(`can't read it: ${e.message}`); }
}

head(`Drayage (${FILES.drayage})`);
const db = bytesOf('drayage');
if (db) {
  try {
    const t0 = performance.now();
    const text = new TextDecoder('utf-8').decode(db);   // as app.js decodes it
    const r = SH.parseDrayage(text, today);
    console.log(`  read in ${secs(t0)}`);
    if (r.headerRow) console.log(`  header row ${r.headerRow}: ${headerLine(SH.parseCsv(text)[r.headerRow - 1])}`);
    const x = r.counts || {};
    if (r.headerRow) {
      sumLine('rows read', x.rowsRead, [['blank', x.blank], ['no container', x.noContainer], ['records', x.records], ['duplicates', x.duplicates]]);
      console.log(`  unknown availability values ${n(r.unknownAvailability.length)}`);
    }
    issueLines(r.issues, rowOf);
    if (!r.ok) fails(r.error);
    else { console.log('  result: ok'); used.drayage = r; }
  } catch (e) { fails(`can't read it: ${e.message}`); }
}

// ---------------- the SharePoint lists, from a saved copy
let pobItems = null;
if (!listsPath) {
  console.log('\n== UPC_MASTER and PO_Book_Lines: no --lists file (and no lists.json in the folder), so the PO Book\'s marks and');
  console.log('   UPC_MASTER are left out below: open/shipped, descriptions and categories can differ from the page.');
} else {
  head(`SharePoint lists (${listsPath})`);
  let L = null;
  try { L = JSON.parse(readFileSync(listsPath, 'utf8')); } catch (e) { fails(`can't read it: ${e.message}`); }
  const items = k => { const v = L && L[k]; return Array.isArray(v) ? v : v && Array.isArray(v.value) ? v.value : null; };
  if (L) {
    pobItems = items('PO_Book_Lines');
    if (!pobItems) fails('it has no PO_Book_Lines array');
    else {
      const ov = PB.overlayFromItems(pobItems), keys = new Set(pobItems.map(i => C.S(i.fields && i.fields.LineKey)).filter(Boolean));
      const noKey = pobItems.filter(i => !C.S(i.fields && i.fields.LineKey)).length;
      console.log(`  PO_Book_Lines: ${n(pobItems.length)} items read · ${n(ov.size)} lines marked in the PO Book · settings item ${keys.has('__settings') ? 'yes' : 'no'} · no LineKey ${n(noKey)}`);
      if (pobItems.length > 1 && !keys.size) fails('no PO_Book_Lines item has a LineKey (the page stops on this)');
    }
    const upcItems = items('UPC_MASTER');
    if (!upcItems) fails('it has no UPC_MASTER array');
    else {
      const t0 = performance.now();
      const u = UPC.fromListItems(upcItems);
      console.log(`  UPC_MASTER: read in ${secs(t0)}`);
      otherCounts(u.counts, []);
      issueLines(u.issues, i => (i.row == null || i.row === '' ? '' : `item ${i.row}`));
      if (!u.ok) fails(u.error || 'UPC_MASTER gave nothing usable');
      else { console.log('  result: ok'); used.upc = u; }
    }
  }
}

// ---------------- Incoming: the stage engine on the same files, with the input app.js gives it
head('Incoming (core/stages.js)');
const LABEL = { noCrd: 'No CRD', inProd: 'In production', ready: 'Ready', booked: 'Booked', loaded: 'Loaded', inChina: 'In China',
  water: 'On the water', port: 'At port', incoming: 'Incoming', pullFromChina: 'Pull from China', unmatched: 'Unmatched' };
let R = null, sections;
const blockers = [!trk && 'PO Status Tracker', !msrc && 'Master Sheet Live', !used.ct && 'Container Tracking'].filter(Boolean);
if (blockers.length) fails(`not loaded: ${blockers.join(', ')} didn't load, so the page shows "not loaded" for every stage number`);
else {
  const ov = pobItems ? PB.overlayFromItems(pobItems) : null, cfg = pobItems ? PB.cfgFrom(pobItems) : null;
  const lines = ov && cfg ? PB.applyOverlay(msrc, ov, cfg) : msrc.lines;   // as app.js: without the list, parseMaster's lines
  const open = lines.filter(l => (l.status || (l.src === 'round' ? 'shipped' : 'open')) === 'open').length;
  const fwd = opt['no-reports'] || opt['no-forwarder'] ? null : used.forwarder, dray = opt['no-reports'] || opt['no-drayage'] ? null : used.drayage;
  const left = ['no-reports', 'no-forwarder', 'no-drayage'].filter(k => opt[k]).map(k => '--' + k).join(' ');
  console.log(`  lines ${n(lines.length)} (open ${n(open)}, shipped ${n(lines.length - open)})${ov ? " with the PO Book's marks" : " WITHOUT the PO Book's marks (no --lists)"}`);
  console.log(`  forwarder ${fwd ? 'used' : 'not used'} · drayage ${dray ? 'used' : 'not used'}${left ? ` (${left})` : ''} · UPC_MASTER ${used.upc ? 'used' : 'not used'} · today ${C.fmtDay(today)}`);
  if (fwd || dray) console.log('  (the reports are used whatever their age here; the page refuses one over 7 days old)');
  // config.js's section lists, as the page passes them (window.INV_CONFIG.sections).
  try { sections = opt['no-sections'] ? undefined : configSections(); } catch (e) { fails(`can't read config.js: ${e.message}`); }
  const lists = sections ? Object.entries(sections).map(([k, v]) => `${k}: ${Array.isArray(v) ? v.join(' | ') : '(not a list)'}`).join(' · ') : '';
  console.log(`  sections ${opt['no-sections'] ? 'NOT used (--no-sections)' : sections ? `from config.js: ${lists}` : 'missing from config.js (no section rules)'} · PO Status Tracker's lowest PO ${trk && Number.isFinite(trk.minPO) ? trk.minPO : 'unknown'}`);
  try {
    const t0 = performance.now();
    R = ST.build({ po: { trk, src: msrc, lines, ov, cfg, counts: mcounts, mrows }, mrows, ct: used.ct, forwarder: fwd, drayage: dray, upc: used.upc, today, sections });
    console.log(`  built in ${secs(t0)}`);
  } catch (e) { fails(`the stage engine stopped: ${e.message}`); }
  if (R && R.ok === false) { fails(`the engine says not loaded: ${(R.notLoaded || []).join(', ')}`); R = null; }
}
if (R) {
  const U = R.totals.units, K = R.totals.cases, NP = R.totals.noCasePackUnits;
  console.log(`\n  ${'Totals'.padEnd(16)} ${'units'.padStart(12)} ${'cases'.padStart(10)} ${'units, no case pack'.padStart(20)}`);
  for (const k of ST.TOTAL_KEYS) console.log(`  ${LABEL[k].padEnd(16)} ${n(U[k]).padStart(12)} ${n(K[k]).padStart(10)} ${n(NP[k]).padStart(20)}`);

  // The totals add up three ways: the derived keys, the styles, and the allocations themselves.
  const tsum = (o, ks) => ks.reduce((s, k) => s + o[k], 0);
  const sumUp = [];
  for (const [what, o] of [['units', U], ['cases', K], ['no-case-pack units', NP]]) {
    if (o.inChina !== tsum(o, ST.IN_CHINA)) sumUp.push(`${what}: In China ${n(o.inChina)} isn't No CRD + In production + Ready + Booked + Loaded (${n(tsum(o, ST.IN_CHINA))})`);
    if (o.incoming !== o.inChina + o.water + o.port) sumUp.push(`${what}: Incoming ${n(o.incoming)} isn't In China + On the water + At port (${n(o.inChina + o.water + o.port)})`);
  }
  const fromStyles = Object.fromEntries(ST.TOTAL_KEYS.map(k => [k, R.styles.reduce((s, x) => s + (x.units[k] || 0), 0)]));
  const offStyles = ST.TOTAL_KEYS.filter(k => fromStyles[k] !== U[k]);
  if (offStyles.length) sumUp.push(`the styles add up to other totals: ${offStyles.map(k => `${LABEL[k]} ${n(fromStyles[k])}`).join(', ')}`);
  const upcOff = R.styles.filter(s => s.upcs.length && ST.TOTAL_KEYS.some(k => s.upcs.reduce((t, u) => t + (u.units[k] || 0), 0) !== (s.units[k] || 0))).length;
  if (upcOff) sumUp.push(`${n(upcOff)} styles whose UPC lines don't add up to the style`);
  const BUCKET_TOTALS = { water: ['water', 'incoming'], port: ['port', 'incoming'], pullFromChina: ['pullFromChina'], unmatched: ['unmatched'] };
  for (const k of ST.IN_CHINA) BUCKET_TOTALS[k] = [k, 'inChina', 'incoming'];
  const fromAllocs = Object.fromEntries(ST.TOTAL_KEYS.map(k => [k, 0]));
  const outside = {};
  for (const a of R.allocs) {
    const hidden = a.bucket === 'unmatched' && a.stage === 'Landed';
    if (!hidden && BUCKET_TOTALS[a.bucket]) for (const k of BUCKET_TOTALS[a.bucket]) fromAllocs[k] += a.units;
    else { const k = hidden ? 'unmatched, landed (hidden)' : a.bucket === 'ctMore' ? `ctMore, ${a.stage === 'Landed' ? 'landed' : 'not landed'}` : a.bucket; outside[k] = (outside[k] || 0) + a.units; }
  }
  const offAllocs = ST.TOTAL_KEYS.filter(k => fromAllocs[k] !== U[k]);
  if (offAllocs.length) sumUp.push(`the allocations add up to other totals: ${offAllocs.map(k => `${LABEL[k]} ${n(fromAllocs[k])}`).join(', ')}`);
  console.log(`  totals add up (In China and Incoming from their parts; the styles; their UPC lines; the allocations): ${sumUp.length ? 'NO' : 'yes'}`);
  for (const x of sumUp) fails(x);
  console.log(`  outside every total (units): ${Object.entries(outside).sort().map(([k, v]) => `${k} ${n(v)}`).join(' · ') || 'none'}`);
  // Units in Container Tracking sections with no round, by section and where they went. Only names on config.js's
  // lists are printed (they're public); any other section is "another section".
  const listed = new Set(Object.values(sections || {}).filter(Array.isArray).flat());
  const known = s => { const k = C.S(s).replace(/\s+/g, ' ').toLowerCase(); return [...listed].some(x => C.S(x).replace(/\s+/g, ' ').toLowerCase() === k); };
  const bySec = {};
  // Container Tracking units only (a Master Sheet line at its own stage under a planning section also names it).
  for (const a of R.allocs) if (a.section && a.ctRows.length) {
    const k = known(a.section) ? a.section : '(another section)';
    const m = bySec[k] || (bySec[k] = {});
    const b = a.bucket === 'unmatched' && a.stage === 'Landed' ? 'unmatched, landed' : a.bucket;
    m[b] = (m[b] || 0) + a.units;
  }
  console.log('  sections with no round (units, by where they went):');
  for (const [s, m] of Object.entries(bySec).sort()) console.log(`    ${s.padEnd(18)} ${Object.entries(m).sort((p, q) => q[1] - p[1]).map(([k, v]) => `${k} ${n(v)}`).join(' · ')}`);
  console.log(`  Container Tracking-only outcomes: planning ${n(R.counts.planningUnits)} · ignored ${n(R.counts.ignoredUnits)} · below the tracker's lowest PO ${n(R.counts.belowTrackerUnits)} on ${n(R.counts.belowTrackerLines)} lines (${n(R.counts.belowTrackerNotLandedUnits)} not landed)` +
    ` · sections not on the lists ${n(R.counts.sectionsNotListed)} (${n(R.counts.sectionNotListedUnits)} units counted Booked/Loaded)`);

  const b = R.balance;
  console.log('\n  Balance');
  console.log(`    Master Sheet Live: ${n(b.groups)} PO + style groups, ${n(b.mslUnits)} units = ${n(b.outcomeUnits)} in a stage or outcome · difference ${n(b.diff)} · groups off ${n(b.bad.length)}`);
  console.log(`    Container Tracking: ${n(b.ct.units)} units on its lines = ${n(b.ct.outcomeUnits)} placed · difference ${n(b.ct.diff)}`);
  if (b.diff !== 0 || b.bad.length) fails(`the Master Sheet Live balance is off by ${n(b.diff)} (${n(b.bad.length)} groups)`);
  if (b.ct.diff !== 0) fails(`the Container Tracking balance is off by ${n(b.ct.diff)}`);

  console.log('\n  Counts');
  for (const [k, v] of Object.entries(R.counts)) console.log(`    ${k}: ${Array.isArray(v) ? v.join(', ') || '(none)' : n(v)}`);
  const by = (list, f) => Object.entries(list.reduce((m, x) => { const k = f(x); m[k] = (m[k] || 0) + 1; return m; }, {})).sort((p, q) => q[1] - p[1]).map(([k, v]) => `${k} ${n(v)}`).join(' · ');
  console.log(`    styles ${n(R.styles.length)} (${n(R.styles.filter(s => s.upcs.length > 1).length)} with more than one UPC) · allocations ${n(R.allocs.length)} · Pull from China rows ${n(R.pull.length)}`);
  console.log(`    allocations by bucket: ${by(R.allocs, a => a.bucket)}`);
  console.log(`    Pull from China by PO type: ${by(R.pull, p => p.ptype)} · by status: ${by(R.pull, p => (p.ctParts.length > 1 ? '(a mix)' : p.ctParts.length ? p.ctParts[0].status + (p.ctParts[0].section ? ' · a section' : '') : 'Not on tracker'))}`);
  // The Pull from China tile against the Pull from China tab (open lines only): the difference is shipped lines.
  const pfc = R.allocs.filter(a => a.bucket === 'pullFromChina'), sumU = a => a.reduce((s, x) => s + x.units, 0);
  const pfcOpen = sumU(pfc.filter(a => a.status === 'open')), pfcShipped = pfc.filter(a => a.sheetRow != null && a.status !== 'open');
  const pfcCt = sumU(pfc.filter(a => a.sheetRow == null));
  // The tab's stage filter picks a line (all its units) when its PO is China Pickup or it has a direct-to-customer section.
  const tabLines = R.pull.filter(p => p.ptype === 'China Pickup' || p.flags.includes('pull-section')), tabKeys = new Set(tabLines.map(p => p.lineKey));
  const tabUnits = sumU(tabLines), tabPfc = sumU(pfc.filter(a => a.status === 'open' && tabKeys.has(a.lineKey)));
  const pfcAdds = pfcOpen + sumU(pfcShipped) + pfcCt === U.pullFromChina;
  console.log(`    Pull from China tile ${n(U.pullFromChina)} = open lines (on the Pull from China tab) ${n(pfcOpen)} + shipped lines ${n(sumU(pfcShipped))} on ${n(new Set(pfcShipped.map(a => a.lineKey)).size)} lines (not on the tab; each listed as ct-behind-msl)${pfcCt ? ` + Container Tracking only ${n(pfcCt)}` : ''}: ${pfcAdds ? 'adds up' : 'DOES NOT add up'}`);
  console.log(`    the tab's "Pull from China" stage filter: ${n(tabLines.length)} lines, ${n(tabUnits)} units = ${n(tabPfc)} in Pull from China + ${n(tabUnits - tabPfc)} of those lines at another stage (partly covered)` +
    (pfcOpen !== tabPfc ? ` · ${n(pfcOpen - tabPfc)} open-line units in Pull from China on lines it doesn't pick` : ''));
  if (!pfcAdds) fails('the Pull from China tile is not open lines + shipped lines');
  const ids = new Set(R.allocs.map(a => a.id));
  if (ids.size !== R.allocs.length) fails(`${n(R.allocs.length - ids.size)} allocations share an id`);
  console.log('');
  issueLines(R.issues, i => [i.sheet, i.row == null || i.row === '' ? '' : `row ${i.row}`].filter(Boolean).join(' '));
  if (R.issues.some(i => i.code === 'balance')) fails('the engine raised a balance issue');
}

// ---------------- the ATS report, read as the page reads it (core/ats.js)
let atsR = null, atsDayUsed = null;   // the report, and its day (file name, else saved) as the page takes it
if (atsPath) {
  head(`ATS report (core/ats.js, ${basename(atsPath)})`);
  if (!existsSync(atsPath)) fails(`${atsPath} isn't there`);
  else {
    try {
      const t0 = performance.now();
      const buf = readFileSync(atsPath);
      const wb = readBook(new Uint8Array(buf.buffer, buf.byteOffset, buf.byteLength));
      const r = ATS.parseAts(XLSX, wb, basename(atsPath), today);
      console.log(`  sheet "${r.sheet || '(none)'}" of ${wb.SheetNames.length}; read in ${secs(t0)}`);
      if (r.headerRow) console.log(`  header row ${r.headerRow}: ${headerLine(C.sheetRows(XLSX, wb.Sheets[r.sheet])[r.headerRow - 1])}`);
      // Its date: from the file name, else the day the file was saved, as on the page (amber after 8 days, an email after 10).
      const saved = C.todayNY(statSync(atsPath).mtimeMs), day = r.reportDay ?? saved, age = today - day;
      atsDayUsed = day;
      console.log(`  report date ${C.fmtDay(day)} (${r.reportDay != null ? 'from the file name' : 'the day the file was saved: no date in the file name'}) · ${n(age)} day${age === 1 ? '' : 's'} old` +
        (age > 10 ? ': the page would email Saul (over 10 days)' : age > 8 ? ': amber on the page (over 8 days)' : ''));
      const x = r.counts || {};
      if (r.headerRow) {
        sumLine('rows read', x.rowsRead, [['blank', x.blankRows], ['Grand Total', x.grandTotalRows], ['no style', x.noStyleRows], ['style', x.styleRows]]);
        const known = [...r.styles.values()].filter(s => typeof s.inWh === 'number' && typeof s.orders === 'number');
        console.log(`  styles ${n(x.styles)} (${n(x.multiWarehouseStyles)} on more than one row, summed; ${n(x.unknownStyles)} whose stock can't be read) · In WH ${n(x.inWh)} · Total Customer Orders ${n(x.orders)}`);
        console.log(`  In WH − Total Customer Orders per style: above 0 on ${n(known.filter(s => s.inWh > s.orders).length)} styles, below 0 on ${n(known.filter(s => s.inWh < s.orders).length)} (sold ahead of stock)`);
        const g = r.grandTotal;
        if (g) {
          const gt = typeof g.inWh === 'number' && !Number.isNaN(g.inWh) ? g.inWh : null;
          console.log(`  Grand Total (row ${g.row}): In WH ${gt == null ? 'blank or not a number' : n(gt)} · the rows add up to it: ${gt === x.inWh ? 'yes' : 'NO'}`);
        } else console.log('  no Grand Total row: nothing to check the rows against');
      }
      issueLines(r.issues, i => [i.style, i.row == null || i.row === '' ? '' : `row ${i.row}`].filter(Boolean).join(' '));
      if (!r.ok) fails(r.error || "the ATS report couldn't be read");
      else { console.log('  result: ok'); atsR = r; }
    } catch (e) { fails(`can't read it: ${e.message}`); }
  }
}

// ---------------- Available to sell (core/sell.js): with the report and without it
head('Available to sell (core/sell.js)');
console.log("  Saul's pull-from date for goods still in China: 10 days before the ETD; no ETD, 5 days before the ready date (CRD);");
console.log('  no CRD, about 75 days after the order date (an estimate). Goods on the water or at port: their ETA.');
if (!R) fails('not worked out: the stage numbers did not load');
else {
  const runs = [];
  if (atsPath) runs.push(atsR ? ['with the ATS report', atsR] : null);
  runs.push(['without the ATS report (as the page shows it when the report is missing or red)', null]);
  for (const run of runs) {
    if (!run) { console.log("\n  with the ATS report: not worked out, the report didn't load (above)"); continue; }
    let S = null;
    try {
      const t0 = performance.now();
      // As the page: the report's day from its file name, else the day it was saved.
      S = SELL.buildSell({ stages: R, ats: run[1], today, upc: used.upc, reportDay: run[1] ? atsDayUsed : null });
      console.log(`\n  ${run[0]}: built in ${secs(t0)}${S.ok && !S.atsMissing ? ` · report date ${S.reportDay != null ? C.fmtDay(S.reportDay) : 'not known'}${run[1] && run[1].reportDay == null ? ' (the day it was saved: no date in the file name)' : ''}` : ''}`);
    } catch (e) { fails(`the selling engine stopped: ${e.message}`); continue; }
    if (!S || !S.ok) { fails(`the selling engine gave no numbers (${(S && (S.notLoaded || []).join(', ')) || 'no answer'})`); continue; }
    sellReport(S, R, !run[1]);
  }
}

// The checks the reconciliation can't make (a unit counted twice, or dropped, still adds up), worked out here from the
// files on their own and set against the engine's numbers (its cross-checks are listed on Problems, never counted:
// Saul decides). Per style on the report with In Trans read:
// - In Trans beyond our sailed and on-vessel goods, due after the report date, while Container Tracking has the style
//   on the water or at port with no Master Sheet line and the sheet still has goods in China: maybe offered twice;
// - taken into the ERP's transit although its Import PO Backlog holds that many of the style's China goods (sold too);
// - our own Landed goods due shortly before the report date, which could be that In Trans instead of China goods;
// - received never more than In WH, and the engine's split of In Trans the same as worked out here.
// Style codes and numbers only.
function erpCrossChecks(S, R, erpS) {
  const rep = S.reportDay, isD = v => typeof v === 'number' && Number.isFinite(v);
  const IN = ['noCrd', 'inProd', 'ready', 'booked', 'loaded'], SAIL = ['water', 'port'], GONE = ['On the water', 'At port', 'Landed', 'Shipped'];
  const soldA = a => a.bucket === 'pullFromChina' || ((IN.includes(a.bucket) || SAIL.includes(a.bucket)) && (a.ptype === 'China Pickup' || (a.flags || []).includes('pull-section')));
  const by = new Map();
  for (const a of R.allocs) (by.get(a.style) || by.set(a.style, []).get(a.style)).push(a);
  const sum = (a, f = x => x.units) => a.reduce((p, x) => p + (f(x) || 0), 0);
  const backlogCol = !!(atsR && Array.isArray(atsR.columns) && atsR.columns.includes('backlog'));
  const L = { twice: [], loose: [], backlog: [], landedOurs: [], received: [], differ: [] };
  for (const s of erpS) {
    const e = atsR.styles.get(s.style), al = by.get(s.style) || [];
    const ships = e.rows.flatMap(x => (x.inTransList || []).map((q, i) => ({ units: q, eta: (x.etaParts || [])[i] ?? null })));
    const E = sum(ships);
    const sailed = al.filter(a => SAIL.includes(a.bucket) && !soldA(a) && a.units > 0), china = al.filter(a => IN.includes(a.bucket) && !soldA(a) && a.units > 0);
    const ours = sum(sailed), due = list => sum(list.filter(y => isD(y.eta) && isD(rep) && y.eta <= rep));
    let taken = 0, landed = 0, unm = 0;
    if (E >= ours) {
      let x = E - ours;
      taken = Math.min(x, sum(china.filter(a => ['booked', 'loaded'].includes(a.bucket) && isD(a.etd))));
      x -= taken;
      landed = Math.min(x, Math.max(0, due(ships) - due(sailed)));
      unm = x - landed;
    }
    const held = sum(china) - taken;
    const ctSailed = sum(al.filter(a => ['unmatched', 'ctMore'].includes(a.bucket) && ['On the water', 'At port'].includes(a.stage)));
    const twice = Math.min(unm, ctSailed, held);
    const soldChina = sum(al.filter(a => soldA(a) && (IN.includes(a.bucket) || (a.bucket === 'pullFromChina' && !GONE.includes(a.stage)))));
    const backlog = sum(e.rows.flatMap(r => (r.backlogList || []).filter(v => typeof v === 'number' && v > 0)), v => v);
    const inBacklog = backlogCol && taken ? Math.max(0, taken - Math.max(0, sum(china) + soldChina - backlog)) : 0;
    const ourLanded = sum(al.filter(a => a.bucket === 'landed' && isD(a.eta) && isD(rep) && a.eta <= rep && a.eta > rep - 45));
    if (twice) L.twice.push([s.style, twice]);
    if (Math.min(unm, held)) L.loose.push([s.style, Math.min(unm, held)]);
    if (inBacklog) L.backlog.push([s.style, inBacklog]);
    if (taken && ourLanded) L.landedOurs.push([s.style, Math.min(taken, ourLanded)]);
    if (isD(s.inWh) && s.receivedDropped > Math.max(0, s.inWh)) L.received.push([s.style, s.receivedDropped - Math.max(0, s.inWh)]);
    if (taken !== s.inErpTransitTaken || landed !== s.landedNotReceived || unm !== s.erpUnmatched || twice !== s.maybeTwice || inBacklog !== s.takenInBacklog) {
      L.differ.push(`${s.style} (here taken ${n(taken)} landed ${n(landed)} not sailed ${n(unm)} twice ${n(twice)} backlog ${n(inBacklog)}; engine ${n(s.inErpTransitTaken)} ${n(s.landedNotReceived)} ${n(s.erpUnmatched)} ${n(s.maybeTwice)} ${n(s.takenInBacklog)})`);
    }
  }
  const line = list => `${n(sum(list, x => x[1]))} units on ${n(list.length)} styles${list.length ? ` (most: ${list.slice().sort((a, b) => b[1] - a[1]).slice(0, 6).map(x => `${x[0]} ${n(x[1])}`).join(' · ')})` : ''}`;
  console.log("    cross-checks the reconciliation can't make (worked out here from the files; listed on Problems, never counted, Saul decides):");
  console.log(`      maybe offered twice: In Trans due after the report date beyond our sailed and on-vessel goods, with Container Tracking on the water / at port under a PO with no Master Sheet line and goods still in China: ${line(L.twice)}`);
  console.log(`        for scale, the same without Container Tracking's word (In Trans not sailed on ours that goods still in China could be): ${line(L.loose)}`);
  console.log(`      taken into the ERP's transit although its Import PO Backlog holds that many of the style's China goods: ${backlogCol ? line(L.backlog) : 'no Import PO Backlog column'}`);
  console.log(`        for scale, taken where our own Landed goods due in the 45 days before the report date could be that In Trans instead: ${line(L.landedOurs)}`);
  console.log(`      received more than the style's In WH (never allowed): ${line(L.received)}`);
  console.log(`      the engine's split of In Trans and its cross-checks against these, style by style: ${L.differ.length ? `DIFFER on ${n(L.differ.length)} (e.g. ${L.differ.slice(0, 3).join(' | ')})` : 'the same on every style'}`);
  if (L.received.length) fails(`Available to sell: ${n(L.received.length)} styles have more received than In WH`);
  if (L.differ.length) fails(`Available to sell: the engine's split of the ERP's In Trans differs from the one worked out here on ${n(L.differ.length)} styles`);
  const T = S.totals;
  if (T.maybeTwice !== sum(L.twice, x => x[1]) || T.takenInBacklog !== sum(L.backlog, x => x[1])) fails("Available to sell: the engine's cross-check totals aren't the ones worked out here");
}

// The selling engine's numbers, and their checks: every sum, Saul's rule and the ready date, lot by lot, worked out here
// on their own. Style codes and PO numbers next to numbers only.
function sellReport(S, R, noReport) {
  const T = S.totals, K = T.cases || {}, NP = T.noCasePackUnits || {}, LO = T.looseUnits || {}, cnt = S.counts || {};
  const ROWS = [['now', 'Available now'], ['fromChina', 'Available from China'], ['arriving', 'Arriving'], ['arrivingErp', '  of it ERP In Trans'],
    ['arrivingOurs', '  of it ours'], ['total', 'Total available'], ['soldChina', 'Sold, ships from China'], ['oversold', 'Oversold'], ['ordersTaken', 'Orders taken off lots']];
  console.log(`    ${'Totals'.padEnd(24)} ${'units'.padStart(12)} ${'whole cases'.padStart(12)} ${'loose units'.padStart(12)} ${'units, no case pack'.padStart(20)}`);
  for (const [k, w] of ROWS) {
    const hasCases = ['now', 'fromChina', 'arriving', 'total', 'soldChina'].includes(k), hasLoose = ['now', 'fromChina', 'arriving', 'total'].includes(k);
    console.log(`    ${w.padEnd(24)} ${n(T[k]).padStart(12)} ${(hasCases ? n(K[k]) : '').padStart(12)} ${(hasLoose ? n(LO[k]) : '').padStart(12)} ${(hasCases ? n(NP[k]) : '').padStart(20)}`);
  }
  const sumU = (a, f = l => l.units) => a.reduce((p, l) => p + (f(l) || 0), 0);
  if (!noReport) {
    // Arriving from the ERP's In Trans, and how our own lots meet it (Saul, Oct 30), worked out here as well.
    const st0 = S.styles, erpS = st0.filter(s => s.arrivingFrom === 'erp'), erpL = S.lots.filter(l => l.source === 'erp');
    const sumS = (a, k) => a.reduce((p, s) => p + (s[k] || 0), 0);
    console.log("    Arriving from the ATS report's In Trans (each quantity with its ETA, the same snapshot as the stock):");
    console.log(`      styles on the report with In Trans read: ${n(cnt.erpStyles)} (In Trans can't be read on ${n(cnt.erpInTransUnknownStyles)}${cnt.noInTransColumn ? '; NO In Trans column' : ''}) · ERP shipments ${n(erpL.length)}, ${n(sumU(erpL))} units (${n(cnt.erpLotsNoEta)} with no ETA: the ETA list doesn't line up on ${n(cnt.erpEtaMismatchStyles)} styles)`);
    console.log(`      ERP In Trans ≥ our sailed units: ${n(cnt.erpMoreStyles)} styles (with sailed lots of ours) · less: ${n(cnt.erpLessStyles)} styles`);
    console.log(`      our sailed units matched to the ERP's In Trans (counted there, not again): ${n(T.matchedOurs)} units on ${n(cnt.matchedLots)} lots`);
    console.log(`      In Trans beyond ours, taken off Booked / Loaded lots with an ETD (earliest ETD first): ${n(T.inErpTransitTaken)} units on ${n(cnt.inErpTransitLots)} lots, ${n(cnt.inErpTransitStyles)} styles`);
    console.log(`      In Trans beyond that on shipments due on or before the report date: landed, not received yet: ${n(T.landedNotReceived)} units on ${n(cnt.landedNotReceivedStyles)} styles`);
    console.log(`      In Trans beyond that, due after the report date (or no ETA): not sailed on our shipping data: ${n(T.erpUnmatched)} units on ${n(cnt.erpUnmatchedStyles)} styles`);
    console.log(`      ERP In Trans < ours: shipments matched first to one of ours of the same quantity: ${n(cnt.sameQtyMatchedLots)} lots`);
    console.log(`      ours with an ETA on or before the report date, not on its In Trans: received (in In WH), dropped: ${n(T.receivedDropped)} units on ${n(cnt.receivedLots)} lots, ${n(cnt.receivedStyles)} styles`);
    console.log(`      ours not on its In Trans yet (a later ETA or none, or more than the style's In WH): kept as Arriving from our data: ${n(T.keptExtra)} units on ${n(cnt.keptExtraLots)} lots, ${n(cnt.keptExtraStyles)} styles` +
      ` (ETA passed but more than In WH: ${n(T.receivedCapped)} units on ${n(cnt.receivedCappedStyles)} styles)${cnt.noReportDayStyles ? ` (${n(cnt.noReportDayStyles)} styles with no report date)` : ''}`);
    // Reconciliation, per style and in total: our lots = Incoming; the ERP's lots = its In Trans; In Trans = matched +
    // in its transit + landed; sellable (from China + arriving + taken) + matched + received + in its transit = Incoming + In Trans.
    const inTransOf = s => { const e = atsR && atsR.styles.get(s.style); return e && s.arrivingFrom === 'erp' ? e.rows.reduce((p, x) => p + (x.inTransList || []).reduce((q, v) => q + v, 0), 0) : 0; };
    const incOf = new Map(R.styles.map(s => [s.style, s.units.incoming]));
    const recBad = [];
    for (const s of st0) {
      const inc = incOf.get(s.style) || 0, it = inTransOf(s);
      const why = [sumU(s.lots.filter(l => l.source === 'ours')) !== inc ? 'our lots ≠ Incoming' : '', sumU(s.lots.filter(l => l.source === 'erp')) !== it ? "the ERP's lots ≠ its In Trans" : '',
        s.arrivingFrom === 'erp' && s.erpInTrans !== s.matchedOurs + s.inErpTransitTaken + s.landedNotReceived + s.erpUnmatched ? 'In Trans ≠ matched + in transit + landed + not sailed on ours' : '',
        s.fromChina + s.arriving + s.ordersTaken + s.matchedOurs + s.receivedDropped + s.inErpTransitTaken !== inc + it ? 'sellable + matched + received + in transit ≠ Incoming + In Trans' : ''].filter(Boolean);
      if (why.length) recBad.push(`${s.style} (${why.join(', ')})`);
    }
    const inc = R.totals.units.incoming, it = sumS(erpS, 'erpInTrans'), sell = T.fromChina + T.arriving + T.ordersTaken;
    console.log(`      reconciliation: Incoming ${n(inc)} + ERP In Trans ${n(it)} = ${n(inc + it)} = sellable ${n(sell)} (from China + arriving + taken by orders) + matched ${n(T.matchedOurs)} + received ${n(T.receivedDropped)} + in the ERP's transit ${n(T.inErpTransitTaken)}` +
      ` (= ${n(sell + T.matchedOurs + T.receivedDropped + T.inErpTransitTaken)}): ${sell + T.matchedOurs + T.receivedDropped + T.inErpTransitTaken === inc + it ? 'yes' : 'NO'}`);
    const itParts = T.matchedOurs + T.inErpTransitTaken + T.landedNotReceived + T.erpUnmatched;
    console.log(`      ERP In Trans ${n(it)} = matched ${n(T.matchedOurs)} + in its transit ${n(T.inErpTransitTaken)} + landed, not received ${n(T.landedNotReceived)} + not sailed on ours ${n(T.erpUnmatched)}: ${it === itParts ? 'yes' : 'NO'}` +
      ` · per style, all four hold on every style: ${recBad.length ? `NO on ${n(recBad.length)} (e.g. ${recBad.slice(0, 3).join(' | ')})` : 'yes'}`);
    if (recBad.length || sell + T.matchedOurs + T.receivedDropped + T.inErpTransitTaken !== inc + it || it !== itParts) fails("Available to sell: the ERP's In Trans and our lots don't reconcile");
    if (atsR && it !== sumS([...atsR.styles.values()].filter(e => erpS.some(s => s.style === e.style)), 'inTrans')) fails("Available to sell: the ERP's lots aren't the report's In Trans for those styles");
    erpCrossChecks(S, R, erpS);
    console.log(`    orders beyond stock no more than the units sold to ship from China (maybe the same orders; listed, open question): ${n(cnt.soldInOrdersStyles)} styles (${n(cnt.soldInOrdersOversold)} oversold), ${n(cnt.soldInOrdersUnits)} units`);
  }
  if (noReport) console.log('    (no report: the page shows Available now, Total available and Oversold as "no report yet"; From China and Arriving are before customer orders)');

  // Lots: where they are, their rule, "now", ready dates (ours; the ERP's shipments have only their ETA).
  const lots = S.lots.filter(l => l.source !== 'erp'), china = lots.filter(l => l.inChina), sailed = lots.filter(l => !l.inChina);
  console.log(`    our lots ${n(lots.length)}: in China ${n(china.length)} (${n(sumU(china))} units), on the water or at port ${n(sailed.length)} (${n(sumU(sailed))} units) · the ERP's In Trans lots ${n(S.lots.length - lots.length)}`);
  const RULE = { etd: '10 days before the ETD', crd: '5 days before the CRD', order: 'about 75 days after the order date', none: 'no date' };
  console.log('    lots in China by pull-from rule (lots / units):');
  for (const k of ['etd', 'crd', 'order', 'none']) {
    const ls = china.filter(l => l.pullRule === k);
    console.log(`      ${k.padEnd(6)} ${RULE[k].padEnd(36)} ${n(ls.length).padStart(7)} / ${n(sumU(ls)).padStart(11)}`);
  }
  const nowL = china.filter(l => l.flags.includes('now')), byStatus = {};
  for (const l of nowL) byStatus[l.status] = (byStatus[l.status] || 0) + 1;
  console.log(`    pull-from date already passed ("now"): ${n(nowL.length)} lots, ${n(sumU(nowL))} units (${n(sumU(nowL, l => l.available))} left after orders) · ${Object.entries(byStatus).map(([k, v]) => `${k} ${n(v)}`).join(' · ') || 'none'}`);
  const noEta = sailed.filter(l => l.availableDate == null);
  console.log(`    sailed lots by arrival: ETA ${n(sailed.length - noEta.length)} · no ETA ${n(noEta.length)}`);
  const RF = { ctReady: "Container Tracking's Ready Date", poBook: "the PO Book's own date", factoryCrd: 'Factory CRD', crdExpected: 'CRD Expected', shipBy: 'Ship by', crd: 'CRD', none: 'none' };
  const rf = {};
  for (const l of lots) rf[l.readyFrom || 'none'] = (rf[l.readyFrom || 'none'] || 0) + 1;
  console.log(`    ready date taken from (lots): ${['ctReady', 'poBook', 'factoryCrd', 'crdExpected', 'shipBy', 'crd', 'none'].filter(k => rf[k]).map(k => `${RF[k]} ${n(rf[k])}`).join(' · ')}` +
    ` · a source more than a year ahead or before 2015 skipped on ${n(cnt.readyFarSkipped)} lots`);
  const dis = china.filter(l => l.flags.includes('dates-disagree')), pairs = new Set(dis.map(l => l.po + '|' + l.style));
  const spreads = dis.map(l => l.readySpread).sort((a, b) => a - b), pct = q => spreads[Math.min(spreads.length - 1, Math.floor(spreads.length * q))];
  const sailedDis = sailed.filter(l => l.readySources.length > 1 && Math.max(...l.readySources.map(x => x.day)) - Math.min(...l.readySources.map(x => x.day)) > 7).length;
  console.log(`    ready dates more than 7 days apart (dates-disagree): ${n(dis.length)} lots in China on ${n(pairs.size)} PO + style pairs` +
    (spreads.length ? ` · spread median ${n(pct(0.5))} days, 90th percentile ${n(pct(0.9))}, most ${n(spreads[spreads.length - 1])}` : '') + ` · sailed lots that disagree (not listed) ${n(sailedDis)}`);
  for (const k of ['etd', 'crd', 'order']) {
    const l = china.find(x => x.pullRule === k && !x.flags.includes('now')) || china.find(x => x.pullRule === k);
    if (l) console.log(`    e.g. ${k}: ${l.style} PO ${l.po} ${l.status} ${n(l.units)} units → ${l.pullNow ? 'now (' : ''}${C.fmtDay(l.pullFrom)}${l.pullNow ? ')' : ''}: ${SELL.ruleWords(l)}`);
  }

  // Styles: a few sanity numbers.
  const st = S.styles, flag = f => st.filter(s => s.flags.includes(f));
  const nor = flag('not-on-report'), norLots = nor.filter(s => s.lots.length);
  console.log(`    styles ${n(st.length)}: on the ATS report ${n(cnt.stylesOnReport)} · not on it ${n(nor.length)} (${n(norLots.length)} with lots, ${n(norLots.reduce((p, s) => p + sumU(s.lots), 0))} units shown in full${norLots.length ? ': ' + norLots.slice(0, 5).map(s => s.style).join(', ') + (norLots.length > 5 ? '…' : '') : ''})` +
    ` · stock can't be read ${n(flag('stock-unknown').length)} · on the report only ${n(flag('erp-only').length)}`);
  const ov = st.filter(s => s.oversold > 0).sort((a, b) => b.oversold - a.oversold);
  console.log(`    with stock now ${n(st.filter(s => s.now > 0).length)} · with something to sell ${n(st.filter(s => s.total > 0).length)} · oversold ${n(ov.length)}${ov.length ? ` (most: ${ov.slice(0, 5).map(s => `${s.style} ${n(s.oversold)}`).join(' · ')})` : ''}`);
  const all = S.lots;
  console.log(`    lots taken by customer orders ${n(all.filter(l => l.taken > 0).length)} (all of it ${n(all.filter(l => l.taken > 0 && !l.available).length)}; the ERP's In Trans ${n(all.filter(l => l.taken > 0 && l.source === 'erp').length)}) · orders not known on ${n(flag('orders-unknown').length)} styles`);

  // The checks: sums, and Saul's rule worked out here, lot by lot.
  const bad = [], isD = v => typeof v === 'number' && Number.isFinite(v), sumS = k => st.reduce((p, s) => p + (s[k] || 0), 0);
  for (const k of ['now', 'fromChina', 'arriving', 'arrivingErp', 'arrivingOurs', 'total', 'soldChina', 'oversold', 'ordersTaken']) if (sumS(k) !== T[k]) bad.push(`the styles add up to ${k} ${n(sumS(k))}, not ${n(T[k])}`);
  if (T.total !== T.now + T.fromChina + T.arriving) bad.push('Total available isn\'t now + from China + arriving');
  if (T.arriving !== T.arrivingErp + T.arrivingOurs) bad.push("Arriving isn't the ERP's In Trans + ours");
  for (const k of ['now', 'fromChina', 'arriving']) if (st.reduce((p, s) => p + ((s.cases || {})[k] || 0), 0) !== K[k]) bad.push(`the styles' cases add up to another ${k}`);
  if (sumU(lots) !== R.totals.units.incoming) bad.push(`our lots hold ${n(sumU(lots))} units, Incoming is ${n(R.totals.units.incoming)}`);
  if (T.soldChina !== R.totals.units.pullFromChina) bad.push(`sold ${n(T.soldChina)} isn't the stage engine's Pull from China ${n(R.totals.units.pullFromChina)}`);
  if (new Set(all.map(l => l.id)).size !== all.length) bad.push('two lots share an id');
  if (st.reduce((p, s) => p + s.lots.length, 0) !== all.length) bad.push("the styles' lots aren't every lot once");
  const lotDay = l => (l.availableDate == null ? Infinity : l.availableDate);
  const styleBad = { lots: [], stock: [], order: [] };
  for (const s of st) {
    if (sumU(s.lots, l => l.taken + l.available) !== s.fromChina + s.arriving + s.ordersTaken || s.fromChina !== sumU(s.lots.filter(l => l.inChina), l => l.available) ||
      s.arriving !== sumU(s.lots.filter(l => !l.inChina), l => l.available) || sumU(s.lots, l => l.taken) !== s.ordersTaken) styleBad.lots.push(s.style);
    const known = isD(s.inWh) && isD(s.orders) && !s.flags.includes('stock-unknown');
    if (known ? s.now !== Math.max(0, s.inWh - s.orders) || s.ordersTaken + s.oversold !== Math.max(0, s.orders - s.inWh) : s.ordersTaken || s.oversold || s.now) styleBad.stock.push(s.style);
    let kept = false;   // customer orders come off the earliest lots first (no date last)
    for (const l of s.lots.slice().sort((a, b) => (lotDay(a) > lotDay(b)) - (lotDay(a) < lotDay(b)))) {
      if (kept && l.taken) { styleBad.order.push(s.style); break; }
      if (l.available > 0) kept = true;
    }
  }
  if (styleBad.lots.length) bad.push(`${n(styleBad.lots.length)} styles whose lots don't add up to from China + arriving + taken (e.g. ${styleBad.lots.slice(0, 3).join(', ')})`);
  if (styleBad.stock.length) bad.push(`${n(styleBad.stock.length)} styles whose stock after orders is off (e.g. ${styleBad.stock.slice(0, 3).join(', ')})`);
  if (styleBad.order.length) bad.push(`${n(styleBad.order.length)} styles where orders took a later lot before an earlier one (e.g. ${styleBad.order.slice(0, 3).join(', ')})`);
  const ruleBad = [];
  // Every lot (ours and the ERP's): units = matched + received + in the ERP's transit + taken + available, none below 0.
  const KEEP = ['matched', 'received', 'inErpTransit', 'taken', 'available'];
  for (const l of all) {
    const d = l.dates || {}, src = l.dates ? ['ctReady', 'poBook', 'factoryCrd', 'crdExpected', 'shipBy'].map(k => d[k]).filter(isD) : isD(l.crd) ? [l.crd] : [];
    // A source more than a year ahead or before 2015 (a wrong year) isn't used, unless every source is.
    const ready = src.find(x => !(x > today + 366 || x < C.ymd(2015, 1, 1))) ?? (src.length ? src[0] : null), why = [];
    if ((l.readyDate ?? null) !== ready) why.push('ready date');
    if (l.units !== KEEP.reduce((p, k) => p + (l[k] || 0), 0) || KEEP.some(k => l[k] < 0)) why.push('units ≠ matched + received + in transit + taken + available');
    if (l.received && !(isD(l.eta) && l.eta <= S.reportDay)) why.push('received with an ETA after the report date');
    if (l.inErpTransit && !(['booked', 'loaded'].includes(l.bucket) && isD(l.etd))) why.push("in the ERP's transit but not Booked / Loaded with an ETD");
    if (l.inChina) {
      const w = isD(l.etd) ? [l.etd - 10, 'etd'] : ready != null ? [ready - 5, 'crd'] : isD(d.orderDate) ? [d.orderDate + 75, 'order'] : [null, 'none'];
      if ((l.pullFrom ?? null) !== w[0] || l.pullRule !== w[1] || !!l.estimate !== (w[1] === 'order') || (l.availableDate ?? null) !== w[0]) why.push(`pull-from ${C.fmtDay(l.pullFrom) || 'none'} (${l.pullRule}), not ${C.fmtDay(w[0]) || 'none'} (${w[1]})`);
      if (w[1] === 'order' && !(l.pullRange && l.pullRange.from === d.orderDate && l.pullRange.to === w[0])) why.push('range');
      if (l.flags.includes('now') !== (w[0] != null && w[0] < today)) why.push('"now"');
      if (l.flags.includes('dates-disagree') !== (src.length > 1 && Math.max(...src) - Math.min(...src) > 7)) why.push('dates-disagree');
    } else if ((l.availableDate ?? null) !== (isD(l.eta) ? l.eta : null)) why.push('not its ETA');
    if (why.length) ruleBad.push(`${l.style} PO ${l.po}: ${why.join(', ')}`);
  }
  console.log(`    sums hold (styles = totals; each style's lots = from China + arriving + taken; stock after orders; earliest lots first;` +
    ` our lots = Incoming ${n(R.totals.units.incoming)}; sold = Pull from China ${n(R.totals.units.pullFromChina)}): ${bad.length ? 'NO' : 'yes'}`);
  console.log(`    Saul's rule, the ready date and units = matched + received + in transit + taken + available, worked out here for each of the ${n(all.length)} lots: ${ruleBad.length ? `${n(ruleBad.length)} DIFFER` : 'all the same'}`);
  if (ruleBad.length) bad.push(`${n(ruleBad.length)} lots break Saul's rule, the ready-date order or their units (e.g. ${ruleBad.slice(0, 3).join(' | ')})`);
  for (const x of bad) fails(`Available to sell: ${x}`);
  console.log('');
  issueLines(S.issues, i => [i.style, i.po ? `PO ${i.po}` : ''].filter(Boolean).join(' '));
}

// ---------------- the ERP's report, per style, against Incoming
if (atsPath) {
  head(`ERP cross-check (${atsPath})`);
  if (!existsSync(atsPath)) fails(`${atsPath} isn't there`);
  else if (!R) fails('not compared: the stage numbers did not load');
  else {
    try { erpCheck(atsPath, R); } catch (e) { fails(`can't read it: ${e.message}`); }
  }
}

function erpCheck(path, R) {
  const buf = readFileSync(path);
  const wb = readBook(new Uint8Array(buf.buffer, buf.byteOffset, buf.byteLength));
  const hit = C.findSheet(wb, 'Report') || { name: wb.SheetNames[0], ws: wb.Sheets[wb.SheetNames[0]] };
  const rows = C.sheetRows(XLSX, hit.ws);
  const SPEC = { style: { names: ['Style'], required: true }, inTrans: { names: ['In Trans Qty'], required: true },
    backlog: { names: ['Import PO Backlog Qty'], required: true }, wh: { names: ['Warehouse'] }, inWh: { names: ['In WH'] },
    eta: { names: ['ETA Date'] } };
  // The header row: the first of rows 1-10 with every required column (the real report has a title row first).
  let hr = -1, fc = null;
  for (let i = 0; i < Math.min(10, rows.length) && hr < 0; i++) { const f = C.findColumns(rows[i], SPEC); if (!f.missing.length) { hr = i; fc = f; } }
  if (hr < 0) { fails(`sheet "${hit.name}": no header row with Style, In Trans Qty and Import PO Backlog Qty in rows 1-10`); return; }
  console.log(`  sheet "${hit.name}", header row ${hr + 1}: ${headerLine(rows[hr])}`);
  const col = fc.cols;
  // A cell holds one quantity, or a list (one per shipment, "a, b, c", lined up with the ETA / ETD lists): summed.
  let badParts = 0;
  const qty = v => {
    if (v == null || C.S(v) === '') return 0;
    if (typeof v === 'number') return v;
    let s = 0;
    for (const p of C.S(v).split(/,\s+/)) { const x = C.num(p); if (typeof x === 'number' && Number.isFinite(x)) s += x; else if (x !== null) badParts++; }
    return s;
  };
  const erp = new Map(), cnt = { blank: 0, total: 0, noStyle: 0, style: 0 }, wh = {};
  for (let i = hr + 1; i < rows.length; i++) {
    const r = rows[i] || [];
    if (!r.some(v => C.S(v) !== '')) { cnt.blank++; continue; }
    const sty = C.normStyle(r[col.style]);
    if (!sty) { if (r.some(v => /^grand total$/i.test(C.S(v)))) cnt.total++; else cnt.noStyle++; continue; }
    cnt.style++;
    const w = col.wh != null ? C.S(r[col.wh]) || '(blank)' : '(no column)';
    wh[w] = (wh[w] || 0) + 1;
    const e = erp.get(sty) || erp.set(sty, { inTrans: 0, backlog: 0, inWh: 0, rows: 0, transPast: 0, transFuture: 0, transNoEta: 0 }).get(sty);
    e.inTrans += qty(r[col.inTrans]); e.backlog += qty(r[col.backlog]); e.inWh += col.inWh != null ? qty(r[col.inWh]) : 0; e.rows++;
    // In Trans split by the ETA it lines up with (one ETA per shipment): before today, or today and later.
    const parts = C.S(r[col.inTrans]) === '' ? [] : typeof r[col.inTrans] === 'number' ? [r[col.inTrans]] : C.S(r[col.inTrans]).split(/,\s+/);
    const etas = col.eta == null || C.S(r[col.eta]) === '' ? [] : typeof r[col.eta] === 'number' ? [r[col.eta]] : C.S(r[col.eta]).split(/,\s+/);
    parts.forEach((p, j) => {
      const x = typeof p === 'number' ? p : C.num(p);
      if (typeof x !== 'number' || !Number.isFinite(x)) return;
      const d = parts.length === etas.length ? C.toDay(etas[j], C.dayYear(today)) : null;
      if (!d || d.day == null) e.transNoEta += x; else if (d.day < today) e.transPast += x; else e.transFuture += x;
    });
  }
  sumLine('rows read', rows.length - hr - 1, [['blank', cnt.blank], ['Grand Total', cnt.total], ['no style', cnt.noStyle], ['style', cnt.style]]);
  console.log(`  styles ${n(erp.size)} (${n([...erp.values()].filter(e => e.rows > 1).length)} on more than one row) · rows by warehouse: ${Object.entries(wh).sort((a, b) => b[1] - a[1]).map(([k, v]) => `${k} ${n(v)}`).join(' · ')}`);
  if (badParts) console.log(`  quantities that aren't numbers (left out): ${n(badParts)}`);

  // Ours per style: On the water + At port (stock: China Pickup goes to Pull from China), and In China on LA POs.
  const ours = new Map(), side = { unknownInChina: 0, unmatchedSailed: 0, unmatchedInChina: 0 };
  const mine = s => ours.get(s) || ours.set(s, { inTrans: 0, backlog: 0 }).get(s);
  for (const a of R.allocs) {
    if (a.bucket === 'water' || a.bucket === 'port') mine(a.style).inTrans += a.units;
    else if (ST.IN_CHINA.includes(a.bucket)) { if (a.ptype === 'LA') mine(a.style).backlog += a.units; else if (a.ptype === 'Unknown') side.unknownInChina += a.units; }
    else if (a.bucket === 'unmatched' && a.stage !== 'Landed') { if (a.stage === 'On the water' || a.stage === 'At port') side.unmatchedSailed += a.units; else side.unmatchedInChina += a.units; }
  }
  // Every unit we have per style, by where it is, to show where a style the ERP has more of sits on our side.
  const where = new Map();
  for (const a of R.allocs) {
    const k = a.bucket === 'unmatched' ? `Unmatched ${a.stage}` : ST.IN_CHINA.includes(a.bucket) ? (a.ptype === 'LA' ? 'In China (LA)' : 'In China (other)')
      : { water: 'On the water', port: 'At port', landed: 'landed', short: 'short-shipped', shippedNotOnCt: 'shipped, not on Container Tracking',
        ctMore: 'beyond Master Sheet Live', pullFromChina: 'Pull from China', duplicate: 'repeat of a shipped line',
        planning: 'planned, not booked yet', ignored: 'ignored section', belowTracker: "below the tracker's lowest PO" }[a.bucket] || a.bucket;
    const m = where.get(a.style) || where.set(a.style, {}).get(a.style);
    m[k] = (m[k] || 0) + a.units;
  }
  const loose = s => s.replace(/\s+/g, '').replace(/-F$/, '');
  const looseOurs = new Set([...where.keys()].map(loose));
  // The ERP's other columns, for a style where we are higher.
  const ERP_OTHER = { inTrans: 'In Trans Qty', backlog: 'Import PO Backlog Qty', inWh: 'In WH' };
  const compare = (field, erpName, ourName, ourKeys) => {
    const all = new Set([...erp.keys(), ...ours.keys()]);
    let te = 0, to = 0, both = 0, agree = 0, same = 0, onlyErp = 0, onlyOurs = 0, notOnReport = 0;
    let gap = 0, absent = 0, absentUnits = 0, looseHit = 0, looseUnits = 0, over = 0, offReport = 0, offReportUnits = 0;
    const gapWhere = {}, overWhere = {};
    const diffs = [];
    for (const s of all) {
      const e = erp.has(s) ? erp.get(s)[field] : 0, o = ours.has(s) ? ours.get(s)[field] : 0;
      te += e; to += o;
      if (!e && !o) continue;
      both++;
      if (e === o) same++;
      if (Math.abs(o - e) <= 0.05 * Math.max(Math.abs(o), Math.abs(e))) agree++;
      if (e && !o) onlyErp++;
      if (o && !e) { onlyOurs++; if (!erp.has(s)) notOnReport++; }
      if (e > o) {
        gap += e - o;
        const w = where.get(s);
        if (!w) {
          absent++; absentUnits += e - o;
          if (looseOurs.has(loose(s))) { looseHit++; looseUnits += e - o; }
        } else for (const [k, v] of Object.entries(w)) if (!ourKeys.includes(k)) {
          const g = gapWhere[k] || (gapWhere[k] = { styles: 0, units: 0 });
          g.styles++; g.units += Math.min(v, e - o);   // each style counts at most its own difference
        }
      }
      if (o > e) {
        over += o - e;
        const x = erp.get(s);
        if (!x) { offReport++; offReportUnits += o - e; } else for (const k of Object.keys(ERP_OTHER)) if (k !== field && x[k] > 0) {
          const g = overWhere[k] || (overWhere[k] = { styles: 0, units: 0 });
          g.styles++; g.units += Math.min(x[k], o - e);
        }
      }
      diffs.push({ s, e, o, d: o - e });
    }
    console.log(`\n  ${erpName} vs our ${ourName}`);
    console.log(`    totals: ERP ${n(te)} · ours ${n(to)} · ours − ERP ${n(to - te)}`);
    console.log(`    styles with a quantity on either side ${n(both)}: within 5% ${n(agree)} (exactly equal ${n(same)}) · ERP only ${n(onlyErp)} · ours only ${n(onlyOurs)} (${n(notOnReport)} not on the ERP report at all)`);
    console.log(`    where the ERP is higher (${n(gap)} units in all): ${n(absentUnits)} on ${n(absent)} styles we have no units of at all` +
      ` (${n(looseHit)} of them, ${n(looseUnits)} units, are a style we have written with other spaces or -F)`);
    console.log(`      the rest, by where else we have those styles (styles, up to that many units of the difference; a style can be in several):`);
    console.log(`        ${Object.entries(gapWhere).sort((p, q) => q[1].units - p[1].units).map(([k, v]) => `${k} ${n(v.styles)} / ${n(v.units)}`).join(' · ') || 'none'}`);
    console.log(`    where ours is higher (${n(over)} units in all): ${n(offReportUnits)} on ${n(offReport)} styles not on the ERP report; the rest, by what the ERP shows`);
    console.log(`      for those styles instead (styles / up to that many units): ${Object.entries(overWhere).sort((p, q) => q[1].units - p[1].units).map(([k, v]) => `${ERP_OTHER[k]} ${n(v.styles)} / ${n(v.units)}`).join(' · ') || 'none'}`);
    console.log('    10 biggest differences (style: ERP / ours / ours − ERP):');
    for (const x of diffs.sort((p, q) => Math.abs(q.d) - Math.abs(p.d) || p.s.localeCompare(q.s)).slice(0, 10)) {
      console.log(`      ${x.s.padEnd(18)} ${n(x.e).padStart(10)} ${n(x.o).padStart(10)} ${((x.d > 0 ? '+' : '') + n(x.d)).padStart(11)}${erp.has(x.s) ? '' : '   (not on the ERP report)'}`);
    }
  };
  compare('inTrans', 'In Trans Qty', 'On the water + At port', ['On the water', 'At port']);
  compare('backlog', 'Import PO Backlog Qty', 'In China on LA POs', ['In China (LA)']);

  // The differences by their two definitional causes. The ERP keeps goods In Trans until it receives them (we take
  // them off once landed), and counts goods booked on a vessel with an ETD as In Trans (we count them In China); its
  // Backlog includes China Pickup POs (our Pull from China).
  const tot = f => [...erp.values()].reduce((s, e) => s + e[f], 0);
  const ourTot = f => [...per.values()].reduce((s, e) => s + e[f], 0);
  const per = new Map(), mineX = s => per.get(s) || per.set(s, { landedRecent: 0, umLandedRecent: 0, sailed: 0, bookedEtd: 0, inChinaLa: 0, pullNotSailed: 0 }).get(s);
  for (const a of R.allocs) {
    const recent = a.eta != null && a.eta >= today - 30;
    if (a.bucket === 'landed' && recent) mineX(a.style).landedRecent += a.units;
    else if (a.bucket === 'unmatched' && a.stage === 'Landed' && recent) mineX(a.style).umLandedRecent += a.units;
    else if (a.bucket === 'water' || a.bucket === 'port') mineX(a.style).sailed += a.units;
    if (ST.IN_CHINA.includes(a.bucket) && a.ptype === 'LA') {
      mineX(a.style).inChinaLa += a.units;
      if ((a.bucket === 'booked' || a.bucket === 'loaded') && a.etd != null) mineX(a.style).bookedEtd += a.units;
    }
    if (a.bucket === 'pullFromChina' && a.stage !== 'On the water' && a.stage !== 'At port') mineX(a.style).pullNotSailed += a.units;
  }
  // Per style, how much of an ERP quantity our figures explain, taking ours in order (each up to what is left).
  const explain = (ef, ofs) => {
    const got = ofs.map(() => 0);
    let all = 0;
    for (const [s, e0] of erp) {
      let e = ef(e0); all += e;
      const o = per.get(s);
      ofs.forEach((of, j) => { const x = o ? Math.max(0, Math.min(e, of(o))) : 0; got[j] += x; e -= x; });
    }
    return { all, got, rest: all - got.reduce((p, x) => p + x, 0) };
  };
  console.log('\n  The differences by cause (per style, each of ours counted up to what the ERP still has left to explain)');
  console.log(`    ERP In Trans ${n(tot('inTrans'))} = ETA before today ${n(tot('transPast'))} + ETA today or later ${n(tot('transFuture'))}${tot('transNoEta') ? ` + no ETA it lines up with ${n(tot('transNoEta'))}` : ''}`);
  const past = explain(e => e.transPast, [o => o.landedRecent, o => o.umLandedRecent]);
  console.log(`    ETA before today ${n(past.all)} (the ERP hasn't received them yet): ${n(past.got[0])} is goods we count as landed with an ETA in the last 30 days,` +
    ` ${n(past.got[1])} more is Unmatched landed; ${n(past.rest)} not explained`);
  const fut = explain(e => e.transFuture, [o => o.sailed, o => o.bookedEtd]);
  console.log(`    ETA today or later ${n(fut.all)}: ${n(fut.got[0])} is our On the water + At port, ${n(fut.got[1])} more is our Booked / Loaded in a round with an ETD` +
    ` (the ERP counts booked-on-a-vessel as in transit); ${n(fut.rest)} not explained`);
  const bl = explain(e => e.backlog, [o => o.inChinaLa - o.bookedEtd, o => o.pullNotSailed, o => o.bookedEtd]);
  console.log(`    Import PO Backlog ${n(bl.all)}: ${n(bl.got[0])} is our In China (LA) not yet on a vessel with an ETD, ${n(bl.got[1])} more is our Pull from China` +
    ` (China Pickup POs and direct-to-customer sections), ${n(bl.got[2])} more is our Booked / Loaded with an ETD; ${n(bl.rest)} not explained`);
  console.log(`    ours, for scale: landed with an ETA in the last 30 days ${n(ourTot('landedRecent'))} · On the water + At port ${n(ourTot('sailed'))} · Booked / Loaded with an ETD ${n(ourTot('bookedEtd'))}` +
    ` · In China (LA) ${n(ourTot('inChinaLa'))} · Pull from China not sailed ${n(ourTot('pullNotSailed'))}`);
  console.log(`\n  Not in either comparison (ours): In China on POs of unknown type ${n(side.unknownInChina)} · Unmatched, not landed: at On the water / At port ${n(side.unmatchedSailed)}, Booked / Loaded ${n(side.unmatchedInChina)}`);
}

// ---------------- the copied PO Book code
head('PO Book copy (core/pobook.js)');
const sha = t => createHash('sha256').update(t, 'utf8').digest('hex');
const ours = PB.coreText(readFileSync(root('core', 'pobook.js'), 'utf8'));
const oursOk = ours != null && sha(ours) === PB.CORE_SHA256;
console.log(`  the block in core/pobook.js hashes to CORE_SHA256: ${oursOk ? 'yes' : 'NO'}`);
if (!oursOk) fails('core/pobook.js was edited between its markers, or CORE_SHA256 is wrong');
const pbPath = poBookArg ? resolve(poBookArg) : process.env.PO_BOOK_HTML ? resolve(process.env.PO_BOOK_HTML) : root('..', 'po-book', 'index.html');
if (!existsSync(pbPath)) console.log(`  not compared: no PO Book file at ${pbPath} (pass its path as the second argument)`);
else {
  const theirs = PB.coreText(readFileSync(pbPath, 'utf8'));
  if (theirs == null) fails(`${pbPath} has no CORE-START / CORE-END markers`);
  else {
    const same = sha(theirs) === PB.CORE_SHA256;
    console.log(`  ${pbPath}: ${same ? 'same rules as the copy' : 'CHANGED since the copy: re-copy the block into core/pobook.js and update CORE_SHA256'}`);
    if (!same) failures++;
  }
}

console.log(failures ? `\n${failures} problem${failures > 1 ? 's' : ''} found.` : '\nEverything read and added up.');
process.exit(failures ? 1 : 0);
