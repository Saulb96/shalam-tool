// Runs the page's parsers on local copies of the real source files and prints only their structure: each header row,
// the counts, issue codes with up to 3 row numbers, and whether core/pobook.js still matches the PO Book.
// Nothing else from the files is printed (no cell values, no issue messages).
//   node tools/check-real.mjs <folder> [po-book index.html] [--today=yyyy-mm-dd]
// <folder> holds Master Sheet Live.xlsx, PO Status Tracker.xlsx, Container Tracking Sheet.xlsx, Forwarder.xlsx and
// Drayage.csv. The PO Book file defaults to $PO_BOOK_HTML, then ../po-book/index.html next to this repo.
// Exit code 1 when a file is missing or fails to parse, counts don't add up, or the PO Book's rules changed.
import { readFileSync, existsSync } from 'node:fs';
import { createHash } from 'node:crypto';
import { join, resolve } from 'node:path';
import { loadXLSX, core, root } from '../tests/load.mjs';

const XOPT = { type: 'array', cellFormula: false, cellHTML: false, cellStyles: false };   // as app.js reads files
const FILES = { msl: 'Master Sheet Live.xlsx', tracker: 'PO Status Tracker.xlsx', ct: 'Container Tracking Sheet.xlsx',
  forwarder: 'Forwarder.xlsx', drayage: 'Drayage.csv' };

const args = process.argv.slice(2);
const opt = Object.fromEntries(args.filter(a => a.startsWith('--')).map(a => { const [k, v] = a.slice(2).split('='); return [k, v ?? true]; }));
const [folderArg, poBookArg] = args.filter(a => !a.startsWith('--'));
if (!folderArg || opt.help) {
  console.log('Usage: node tools/check-real.mjs <folder> [po-book index.html] [--today=yyyy-mm-dd]');
  process.exit(folderArg ? 0 : 2);
}
const folder = resolve(folderArg);

const XLSX = loadXLSX();
const C = core('common'), PB = core('pobook'), CT = core('ct'), SH = core('shipping');
const today = opt.today ? C.dayFromIso(opt.today) : C.todayNY(Date.now());
if (today == null) { console.log(`--today=${opt.today} isn't a yyyy-mm-dd date.`); process.exit(2); }

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
if (mb) {
  try {
    const t0 = performance.now();
    const r = readSheet(mb, 'Master Sheet Live'), mrows = r.rows;
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
        console.log(`  without the PO Book's marks (PO_Book_Lines isn't read locally): open ${n(open)}, shipped ${n(lines.length - open)}`);
        // As the page lists them: countMaster's issues (with rows), then parseMaster's own except 'nostyle' (a repeat
        // of countMaster's) and 'trkonly' (listed under the tracker).
        const own = src.issues.filter(i => i.code !== 'nostyle' && i.code !== 'trkonly');
        trkonly = src.issues.filter(i => i.code === 'trkonly');
        // parseMaster's issues carry no row field; 'dupround' names its row in the text.
        const rowIn = i => { const m = /\brow (\d+)/.exec(i.issue || ''); return m ? `row ${m[1]}` : '(no row)'; };
        issueLines([...cm.issues, ...own], i => (i.row ? `row ${i.row}` : rowIn(i)));
        if (!cm.ok || !src.lines.length) fails(cm.error || 'no lines');
        else console.log('  result: ok');
      }
    }
  } catch (e) { fails(`can't read it: ${e.message}`); }
}
if (trkonly.length) console.log(`\n  PO Status Tracker, also listed by the PO Book: trkonly ${n(trkonly.length)} (no row numbers)`);

// ---------------- Container Tracking
head(`Container Tracking (${FILES.ct})`);
const cb = bytesOf('ct');
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
    else console.log('  result: ok');
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
    else console.log('  result: ok');
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
    else console.log('  result: ok');
  } catch (e) { fails(`can't read it: ${e.message}`); }
}

console.log('\n== UPC_MASTER and PO_Book_Lines are SharePoint lists: not checked here (the page reads them).');

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
