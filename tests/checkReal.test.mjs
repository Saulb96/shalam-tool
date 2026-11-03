// tools/check-real.mjs writes its sample sales sheet (real numbers) only outside this repo, which is public. It compares
// the real folders, so another spelling of the repo's own folder is refused before anything is read or written. Its
// report's pictures section prints counts only: never a style code or a link.
import test from 'node:test';
import assert from 'node:assert/strict';
import { spawnSync, execFileSync } from 'node:child_process';
import { mkdtempSync, mkdirSync, symlinkSync, unlinkSync, lstatSync, existsSync, rmSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { root, core, loadExcelJS } from './load.mjs';

const run = (...args) => spawnSync(process.execPath, [root('tools/check-real.mjs'), ...args], { encoding: 'utf8', timeout: 120000 });
// The repo folder's short (8.3) name, where Windows keeps one ('' when it doesn't).
function shortName(p) {
  if (process.platform !== 'win32') return '';
  try {
    const s = execFileSync('cmd.exe', ['/d', '/s', '/c', `"for %I in ("${p}") do @echo %~sI"`], { encoding: 'utf8', windowsVerbatimArguments: true }).trim();
    return s && s.toLowerCase() !== p.toLowerCase() ? s : '';
  } catch (e) { return ''; }
}

test('check-real: --sales-out inside this repo is refused however its path is spelled (short name, junction, ..)', () => {
  const tmp = mkdtempSync(join(tmpdir(), 'shalam-check-')), src = join(tmp, 'no files'), link = join(tmp, 'repo link');
  mkdirSync(src);
  let linked = false;
  try {
    // A junction (a link on other systems) to the repo folder, outside it.
    try { symlinkSync(root(), link, 'junction'); linked = true; } catch (e) { /* not allowed here: the other spellings still run */ }
    const short = shortName(root());
    const inside = [root('sample.xlsx'), join(root('tests'), '..', 'sample.xlsx'), root('tests', 'sample.xlsx'),
      ...(linked ? [join(link, 'sample.xlsx'), join(link, 'tests', 'sample.xlsx')] : []), ...(short ? [join(short, 'sample.xlsx')] : [])];
    assert.ok(process.platform !== 'win32' || (linked && short), `the junction (${linked}) and the short name (${short || 'none'}) were both tried`);
    for (const p of inside) {
      const r = run(src, '--sales-out', p);
      assert.equal(r.status, 2, `${p}: ${r.stdout}${r.stderr}`);
      assert.match(r.stdout, /is inside this repo \(.*\); write the sample somewhere else/, p);
    }
    assert.ok(!existsSync(root('sample.xlsx')) && !existsSync(root('tests', 'sample.xlsx')), 'nothing written in the repo');
    // A folder that isn't there: refused, not made.
    const nowhere = run(src, '--sales-out', join(tmp, 'missing', 'sample.xlsx'));
    assert.equal(nowhere.status, 2);
    assert.match(nowhere.stdout, /the folder .* isn't there; make it first/);
    // Outside the repo: taken (no sheet here, since the folder has no files, so nothing is written).
    const out = run(src, '--sales-out', join(tmp, 'sample.xlsx'));
    assert.doesNotMatch(out.stdout, /inside this repo|isn't there/);
    assert.equal(out.status, 1, 'the missing files fail the check');
    assert.ok(!existsSync(join(tmp, 'sample.xlsx')));
  } finally {
    // The link first, on its own (never through it), then the folder, only once the link is surely gone.
    if (linked) unlinkSync(link);
    let gone = true;
    try { lstatSync(link); gone = false; } catch (e) { /* gone */ }
    if (gone) rmSync(tmp, { recursive: true, force: true });
  }
});

// An invented report with pictures on two style rows and one linked to an invented picture server (External).
async function pictureReport() {
  const EP = core('erpPictures'), E = loadExcelJS(), wb = new E.Workbook(), ws = wb.addWorksheet('Report');
  const PNG = Uint8Array.from(Buffer.from('iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==', 'base64'));
  for (const r of [['Inventory Report'], ['Image', 'Style', 'Warehouse', 'In WH', 'Total Customer Orders'], [null, 'CK101', 'WH A', 5, 1],
    [null, 'CK102', 'WH A', 7, 0], [null, 'CK103', 'WH A', 2, 0], [null, 'Grand Total', null, 14, 1]]) ws.addRow(r);
  const ids = [PNG, Uint8Array.from([...PNG, 0]), Uint8Array.from([...PNG, 1])].map(buffer => wb.addImage({ buffer, extension: 'png' }));
  ids.forEach((id, i) => ws.addImage(id, { tl: { col: 0, row: 2 + i }, br: { col: 1, row: 3 + i } }));
  const bytes = new Uint8Array(await wb.xlsx.writeBuffer()), parts = [];
  for (const e of EP.zipEntries(bytes)) {
    if (e.name.endsWith('/') || e.name === 'xl/media/image3.png') continue;
    let b = await EP.inflateEntry(bytes, e);
    if (e.name === 'xl/drawings/_rels/drawing1.xml.rels') {
      b = Buffer.from(Buffer.from(b).toString('utf8').replace('Target="../media/image3.png"', 'Target="https://pictures.example.invalid/CK103.png?SIGNATURE=abc" TargetMode="External"'));
    }
    parts.push([e.name, b]);
  }
  return EP.storedZip(parts);
}

test("check-real: --pictures prints the report's picture counts and checks its numbers against the whole-file read", async () => {
  const tmp = mkdtempSync(join(tmpdir(), 'shalam-check-')), src = join(tmp, 'no files');
  mkdirSync(src);
  try {
    const file = join(tmp, 'Example_Inventory_Report_11-02-2026_04-59_PM.xlsx');
    writeFileSync(file, await pictureReport());
    const r = run(src, '--pictures', file, '--today', '2026-11-02');
    assert.equal(r.status, 1, 'the missing files fail the check');
    assert.match(r.stdout, /== The report's pictures \(core\/erpPictures\.js, Example_Inventory_Report_11-02-2026_04-59_PM\.xlsx\)/);
    assert.match(r.stdout, /bytes: under 2 MB: the page reads it as before and takes no pictures from it · read here in [\d.]+ s \(the whole file: [\d.]+ s\)/);
    assert.match(r.stdout, /pictures 2 \(styles with one\) · drawing anchors 3 · linked outside the file 1 \(skipped, never opened\) · kinds the browser can't draw 0 · on rows with no style 0/);
    assert.match(r.stdout, /kinds: png 2 · /);
    assert.match(r.stdout, /its numbers are exactly the whole-file read's \(every style, count and issue\): yes/);
    assert.match(r.stdout, /missing 0 · damaged 0 · far too big 0\n/);
    assert.match(r.stdout, /the numbers' part on its own \(the page keeps it for a big report\): [\d,]+ bytes \([\d.]+% of the file\), \d+ parts, no file properties, drawing or picture among them: yes · read again, its numbers are the whole-file read's: yes/);
    assert.doesNotMatch(r.stdout, /FAILED: (its pictures|the numbers read with|can't read its pictures)/);
    for (const s of ['CK101', 'CK102', 'CK103', 'pictures.example', 'SIGNATURE']) assert.ok(!r.stdout.includes(s), `${s} is printed`);
  } finally { rmSync(tmp, { recursive: true, force: true }); }
});
