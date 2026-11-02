// tools/check-real.mjs writes its sample sales sheet (real numbers) only outside this repo, which is public. It compares
// the real folders, so another spelling of the repo's own folder is refused before anything is read or written.
import test from 'node:test';
import assert from 'node:assert/strict';
import { spawnSync, execFileSync } from 'node:child_process';
import { mkdtempSync, mkdirSync, symlinkSync, unlinkSync, lstatSync, existsSync, rmSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { root } from './load.mjs';

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
