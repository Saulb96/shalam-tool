// The page only reads, except the error email and saving the sales sheet into the one library in config.js. These read
// the page's own files as text and run its guards on request shapes; the headless check watches the requests it really
// makes.
import test from 'node:test';
import assert from 'node:assert/strict';
import { readFileSync, readdirSync } from 'node:fs';
import vm from 'node:vm';
import { root } from './load.mjs';

const text = p => readFileSync(root(p), 'utf8');
const app = text('app.js');
const pageFiles = ['app.js', 'index.html', 'config.js', ...readdirSync(root('core')).filter(f => f.endsWith('.js')).map(f => 'core/' + f)];
const GRAPH = 'https://graph.microsoft.com/v1.0';
// A piece of app.js, from `from` up to (not including) the next `to` after it.
function slice(from, to) {
  const a = app.indexOf(from), b = app.indexOf(to, a + from.length);
  assert.ok(a > 0 && b > a, `app.js: "${from.trim()}" … "${to.trim()}" not found`);
  return app.slice(a, b);
}
const fnText = name => slice(`  async function ${name}(`, '\n  }\n');

test('every request leaves through g() or fetchRetry(), and both check it with the guard first', () => {
  const fetches = app.split('\n').map((l, i) => [i + 1, l]).filter(([, l]) => /\bfetch\s*\(/.test(l) && !/^\s*\/\//.test(l));
  const ok = [/const r = await fetch\(\/\^https\?:\/\.test\(url\) \? url : GRAPH \+ url, \{ \.\.\.opt, headers \}\);/,   // g()
    /const r = await fetch\(url, opt\);/,                                                                              // fetchRetry()
    /await fetch\(location\.origin \+ '\/po-book\/index\.html', \{ cache: 'no-store', signal: timed\(signal, 15000\) \}\)/];   // the PO Book page (a GET)
  assert.equal(fetches.length, ok.length, fetches.map(([n, l]) => `app.js:${n} ${l.trim()}`).join('\n'));
  for (const [n, l] of fetches) assert.ok(ok.some(re => re.test(l)), `app.js:${n} fetches outside g() and fetchRetry(): ${l.trim()}`);
  assert.match(fnText('g'), /^  async function g\(url, opt = \{\}, tries = \{ auth: 0, busy: 0 \}\) \{\n    const method = String\(opt\.method \|\| 'GET'\)\.toUpperCase\(\);\n    const refused = refusal\(method, url, opt, true\);\n    if \(refused\) throw new Error\(refused\);\n    const t = await token\(tries\.auth > 0\);/);
  assert.match(fnText('fetchRetry'), /^  async function fetchRetry\(url, opt\) \{\n    const refused = refusal\(opt && opt\.method, url, opt, false\);\n    if \(refused\) throw new Error\(refused\);\n    for \(/);
  for (const f of pageFiles) {
    assert.doesNotMatch(text(f), /\b(PATCH|DELETE|MERGE)\b/, `${f} names a write method`);
    assert.doesNotMatch(text(f), /XMLHttpRequest|sendBeacon|<form\b|:\/content\b/, `${f} has another way to write`);
    if (f !== 'app.js') assert.doesNotMatch(text(f), /\bPUT\b|conflictBehavior/, `${f} names a write`);
  }
  // Only app.js talks to the network; the parsers and the sales sheet never do.
  for (const f of pageFiles.filter(x => x.startsWith('core/'))) assert.doesNotMatch(text(f), /\bfetch\s*\(/, f);
});

test('the methods the page names: the email, the sales sheet\'s upload session, batches of reads, and PUTs to the upload address', () => {
  const lines = app.split('\n').map((l, i) => [i + 1, l]).filter(([, l]) => /\bmethod\s*:/.test(l));
  for (const [n, l] of lines) {
    const ms = [...l.matchAll(/\bmethod\s*:\s*'([A-Za-z]+)'/g)].map(m => m[1]);
    assert.ok(ms.length, `app.js:${n} sets a method that isn't written out: ${l.trim()}`);
    for (const m of ms) {
      assert.ok(['GET', 'POST', 'PUT'].includes(m), `app.js:${n} ${m}`);
      if (m === 'POST') assert.ok(/g\('\/me\/sendMail', \{ method: 'POST'/.test(l) || /'\/\$batch', \{ method: 'POST'/.test(l) ||
        /^    const j = await gj\(path, \{ method: 'POST', signal, body: JSON\.stringify\(\{ item: \{ '@microsoft\.graph\.conflictBehavior': 'replace' \} \}\) \}\);$/.test(l), `app.js:${n} ${l.trim()}`);
      if (m === 'PUT') assert.match(l, /fetchRetry\(/, `app.js:${n}: a PUT goes through fetchRetry() (no sign-in token)`);
    }
  }
  assert.ok(lines.some(([, l]) => /g\('\/me\/sendMail', \{ method: 'POST'/.test(l)), 'the error email');
  assert.equal((app.match(/conflictBehavior/g) || []).length, 1, 'one upload session');
  assert.match(fnText('uploadSession'), /const path = sessionPath\(\);/);
});

// The guard, run in a sandbox on request shapes. state: what salesDrive() and uploadSession() would have set.
function guard(state = {}, cfg = { salesFile: 'Sales Test.xlsx' }) {
  const ctx = vm.createContext({ CFG: cfg, GRAPH, URL });
  vm.runInContext(app.split('\n').find(l => /^  const encPath = /.test(l)), ctx);
  vm.runInContext(slice('  // ---------- what may leave the page ----------', "  // The sales sheet's library (config.js salesLibrary"), ctx);
  vm.runInContext(`allowed.salesDrive = ${JSON.stringify(state.drive || '')}; allowed.upload = ${JSON.stringify(state.upload || '')};`, ctx);
  const refusal = vm.runInContext('refusal', ctx);
  return (method, url, opt, graph = true) => refusal(method, url, opt, graph);
}
const batch = (...reqs) => ({ body: JSON.stringify({ requests: reqs.map((r, i) => ({ id: String(i + 1), ...r })) }) });
const refused = r => assert.match(r, /^Shalam Tool only reads \(apart from the error email and saving the sales sheet\), so it refused /, r || '(allowed)');

test('the guard: any GET, the email, a batch of GETs; everything else refused', () => {
  const ok = guard();
  for (const graph of [true, false]) {
    assert.equal(ok('GET', '/sites/x/drives', {}, graph), '');
    assert.equal(ok(undefined, 'https://contoso.sharepoint.com/download?k=1', {}, graph), '');
    for (const m of ['PATCH', 'DELETE', 'MERGE', 'HEAD', 'OPTIONS', 'delete']) refused(ok(m, '/drives/b!src/items/1', {}, graph));
  }
  assert.equal(ok('POST', '/me/sendMail', { body: '{}' }), '');
  assert.equal(ok('POST', GRAPH + '/me/sendMail', { body: '{}' }), '');
  refused(ok('POST', '/me/sendMail', { body: '{}' }, false));
  refused(ok('POST', '/users/someone/sendMail', { body: '{}' }));
  refused(ok('POST', '/me/sendMail?x=1', { body: '{}' }));
  // A batch: allowed only when every request in it is a GET, with no body.
  assert.equal(ok('POST', '/$batch', batch({ method: 'GET', url: '/drives/b!p/root:/thumb/AB1.jpg' }, { method: 'GET', url: '/drives/b!p/root:/thumb/AB2.jpg' })), '');
  assert.equal(ok('POST', GRAPH + '/$batch', batch({ method: 'get', url: '/me' })), '');
  refused(ok('POST', '/$batch', batch({ method: 'GET', url: '/me' }, { method: 'POST', url: '/me/sendMail', body: {} })));
  refused(ok('POST', '/$batch', batch({ method: 'DELETE', url: '/drives/b!p/items/1' })));
  refused(ok('POST', '/$batch', batch({ url: '/me' })));
  refused(ok('POST', '/$batch', batch({ method: 'GET', url: '/me', body: { a: 1 } })));
  refused(ok('POST', '/$batch', batch()));
  refused(ok('POST', '/$batch', { body: 'not json' }));
  refused(ok('POST', '/$batch', { body: { requests: [{ id: '1', method: 'GET', url: '/me' }] } }));
  refused(ok('POST', '/$batch', batch({ method: 'GET', url: '/me' }), false));
  // Nothing about the sales sheet is allowed until its library and session are known.
  refused(ok('POST', '/drives//root:/Sales%20Test.xlsx:/createUploadSession', {}));
  refused(ok('PUT', 'https://contoso.sharepoint.com/upload', {}, false));
});

test('the guard: the one upload session for the sales sheet in its library, and PUTs only to the address it returned', () => {
  const UP = 'https://contoso.sharepoint.com/sites/Reports/_api/v2.0/drives/b!sales/items/up/uploadSession?guid=abc&tempauth=xyz';
  const ok = guard({ drive: 'b!sales', upload: UP });
  const SESSION = '/drives/b!sales/root:/Sales%20Test.xlsx:/createUploadSession';
  assert.equal(ok('POST', SESSION, { body: '{}' }), '');
  assert.equal(ok('POST', GRAPH + SESSION, { body: '{}' }), '');
  refused(ok('POST', SESSION, { body: '{}' }, false));                                       // not without the sign-in token
  refused(ok('POST', '/drives/b!other/root:/Sales%20Test.xlsx:/createUploadSession', {}));   // another library
  refused(ok('POST', '/drives/b!sales/root:/Other.xlsx:/createUploadSession', {}));          // another file
  refused(ok('POST', '/drives/b!sales/root:/Sales Test.xlsx:/createUploadSession', {}));     // the name not as encoded
  refused(ok('POST', '/drives/b!sales/root:/Folder/Sales%20Test.xlsx:/createUploadSession', {}));
  refused(ok('POST', SESSION + '?x=1', {}));
  refused(ok('POST', '/drives/b!sales/items/root/children', { body: '{}' }));                 // a new file another way
  refused(ok('PUT', SESSION, {}));
  // PUT: the upload address only, with no sign-in token.
  assert.equal(ok('PUT', UP, { headers: { 'Content-Range': 'bytes 0-9/10' }, body: new Uint8Array(10) }, false), '');
  assert.equal(ok('put', UP, {}, false), '');
  refused(ok('PUT', UP, { headers: { Authorization: 'Bearer t' } }, false));
  refused(ok('PUT', UP, { headers: { authorization: 'Bearer t' } }, false));
  refused(ok('PUT', UP, {}, true));                                                          // through g(), which adds the token
  refused(ok('PUT', UP + '&more=1', {}, false));                                             // somewhere else
  refused(ok('PUT', 'https://contoso.sharepoint.com/sites/Reports/Shared%20Documents/x.xlsx', {}, false));
  refused(ok('PUT', GRAPH + '/drives/b!sales/root:/Sales%20Test.xlsx:/content', {}, false));
  refused(ok('PATCH', UP, {}, false));
  refused(ok('DELETE', UP, {}, false));
  refused(ok('POST', UP, {}, false));
  // An upload address that isn't on SharePoint is never allowed, even if it was recorded.
  for (const u of ['https://contoso.example/upload', 'http://contoso.sharepoint.com/upload', 'https://contoso.sharepoint.com.example/upload', 'not a url']) {
    refused(guard({ drive: 'b!sales', upload: u })('PUT', u, {}, false));
  }
  // A blank salesFile: no session at all.
  refused(guard({ drive: 'b!sales' }, { salesFile: '  ' })('POST', '/drives/b!sales/root::/createUploadSession', {}));
});

test('only salesDrive() and uploadSession() set what the guard allows, and the upload address must be on SharePoint', () => {
  const drive = fnText('salesDrive'), session = fnText('uploadSession');
  const sets = re => [...app.matchAll(re)].length;
  assert.equal(sets(/allowed\.salesDrive = /g), 2);
  assert.equal((drive.match(/allowed\.salesDrive = /g) || []).length, 2);
  assert.match(drive, /allowed\.salesDrive = await pathDrive\(loc, /);
  assert.equal(sets(/allowed\.upload = /g), 2);
  assert.match(session, /allowed\.upload = '';[^]*if \(!onSharePoint\(u\)\) throw [^]*allowed\.upload = u;\n    return u;$/);
  assert.match(session, /'@microsoft\.graph\.conflictBehavior': 'replace'/);
});

test('sign-in asks for read permissions only, plus Mail.Send; it never asks for the saving permission', () => {
  const m = /const SCOPES = (\[[^\]]*\]);/.exec(app);
  assert.ok(m, 'SCOPES not found');
  assert.deepEqual(JSON.parse(m[1].replace(/'/g, '"')), ['User.Read', 'Sites.Read.All', 'Mail.Send']);
  assert.doesNotMatch(app.replace(/^\s*\/\/.*$/gm, ''), /scopes: \[[^\]]*Selected/);
  // The stop rule: other than Mail.Send and exactly Lists.SelectedOperations.Selected, any scope with Write, Manage or
  // FullControl in it, or that isn't a plain read (or a sign-in scope); and Sites.Read.All required.
  assert.match(app, /const SALES_SCOPE = \/\^Lists\\\.SelectedOperations\\\.Selected\$\/i;/);
  assert.match(app, /!\/\^Mail\\\.Send\$\/i\.test\(x\) && !SALES_SCOPE\.test\(x\) && \(\/Write\|Manage\|FullControl\/i\.test\(x\) \|\| !\(isRead\(x\) \|\| OIDC\.test\(x\)\)\)/);
  assert.match(app, /const isRead = x => x\.split\('\.'\)\.some\(p => \/\^Read\(Basic\)\?\$\/i\.test\(p\)\);/);
  assert.match(app, /Shalam Tool may only read \(apart from the error email and saving the sales sheet\), but its sign-in still allows changes to SharePoint \(\$\{writes\.join\(', '\)\}\)\. In Entra, remove that permission from the Shalam Tool app and grant consent again\./);
  assert.match(app, /\^Sites\\\.Read\\\.All\$/);
});

test('the stop rule, run on permission lists (from the token, or from the sign-in service when the token is opaque)', () => {
  const from = app.indexOf('  const OIDC = '), endMark = "    return '';\n  }", to = app.indexOf(endMark, from) + endMark.length;
  assert.ok(from > 0 && to > from, 'scopeProblem not found in app.js');
  const ctx = { atob: s => Buffer.from(s, 'base64').toString('binary'), mailScopeMissing: null, salesScopeMissing: null };
  vm.runInNewContext(app.slice(from, to), ctx);
  const jwt = scp => ['{"alg":"none"}', JSON.stringify({ scp })].map(x => Buffer.from(x).toString('base64url')).join('.') + '.sig';
  const check = (scp, opaque) => ctx.scopeProblem(opaque ? { accessToken: 'opaque', scopes: scp.split(' ').filter(Boolean) } : { accessToken: jwt(scp), scopes: [] });
  for (const opaque of [false, true]) {
    assert.equal(check('email Mail.Send openid profile Sites.Read.All User.Read offline_access', opaque), '');
    assert.deepEqual([ctx.mailScopeMissing, ctx.salesScopeMissing], [false, true], 'no saving permission: saving is off');
    assert.equal(check('email Mail.Send openid profile Sites.Read.All User.Read Lists.SelectedOperations.Selected', opaque), '');
    assert.deepEqual([ctx.mailScopeMissing, ctx.salesScopeMissing], [false, false], 'the saving permission turns saving on');
    assert.equal(check('Sites.Read.All lists.selectedoperations.selected', opaque), '');
    assert.deepEqual([ctx.mailScopeMissing, ctx.salesScopeMissing], [true, false]);
    assert.equal(check('User.ReadBasic.All Sites.Read.All Mail.ReadBasic', opaque), '');
    assert.deepEqual([ctx.mailScopeMissing, ctx.salesScopeMissing], [true, true]);
    for (const [scp, named] of [['User.Read Sites.Read.All Sites.ReadWrite.All Mail.Send', 'Sites.ReadWrite.All'],
      ['User.Read Sites.Read.All Files.ReadWrite Mail.Send', 'Files.ReadWrite'], ['Sites.Read.All Sites.FullControl.All', 'Sites.FullControl.All'],
      ['Sites.Read.All Sites.Selected Mail.Send', 'Sites.Selected'], ['Sites.Read.All ListItems.SelectedOperations.Selected', 'ListItems.SelectedOperations.Selected'],
      ['Sites.Read.All Files.SelectedOperations.Selected Lists.SelectedOperations.Selected', 'Files.SelectedOperations.Selected'],
      ['Sites.Read.All Lists.SelectedOperations.Selected Sites.ReadWrite.All', 'Sites.ReadWrite.All'],
      ['Sites.Read.All Lists.ReadWrite.All', 'Lists.ReadWrite.All'], ['Sites.Read.All Sites.Manage.All', 'Sites.Manage.All'],
      ['Sites.Read.All Lists.SelectedOperations', 'Lists.SelectedOperations'], ['Sites.Read.All Lists.SelectedOperations.Selected.All', 'Lists.SelectedOperations.Selected.All'],
      ['Sites.Read.All Lists.SelectedOperations.Selected.Write', 'Lists.SelectedOperations.Selected.Write'],
      ['Sites.Read.All Mail.Send.Shared Mail.Send', 'Mail.Send.Shared'], ['Sites.Read.All Directory.AccessAsUser.All', 'Directory.AccessAsUser.All'],
      ['Sites.Read.All Sites.Create.All', 'Sites.Create.All'], ['Sites.Read.All Files.ReadWrite.AppFolder', 'Files.ReadWrite.AppFolder']]) {
      const msg = check(scp, opaque);
      assert.match(msg, /^Shalam Tool may only read \(apart from the error email and saving the sales sheet\), but its sign-in still allows changes to SharePoint \(/, scp);
      assert.ok(msg.includes(`(${named})`), `${scp} -> ${msg}`);
    }
    assert.match(check('User.Read Mail.Send', opaque), /missing the Sites\.Read\.All permission \(it has: User\.Read, Mail\.Send\)/);
    assert.match(check('Lists.SelectedOperations.Selected', opaque), /missing the Sites\.Read\.All permission/, 'saving alone never reads');
  }
  assert.match(check('', true), /couldn't read which permissions its sign-in has/);
  assert.match(ctx.scopeProblem({ accessToken: 'opaque' }), /couldn't read which permissions its sign-in has/);
});

test('config.js holds the sign-in, every source location, the sales sheet\'s library, file and photos, and the error email', () => {
  const ctx = { window: {} };
  vm.runInNewContext(text('config.js'), ctx);
  const c = ctx.window.INV_CONFIG;
  assert.ok(c, 'window.INV_CONFIG not set');
  assert.match(c.tenantId, /^[0-9a-f-]{36}$/);
  assert.match(c.clientId, /^[0-9a-f-]{36}$/);
  assert.deepEqual(Object.keys(c.sources).sort(), ['ats', 'ct', 'drayage', 'forwarder', 'msl', 'tracker']);
  for (const [k, v] of Object.entries(c.sources)) {
    assert.equal(typeof v, 'string', k);
    assert.ok(v === '' || /^https:\/\/[^/]+\.sharepoint\.com\/:[a-z]:\//.test(v) || /^\/sites\/[^/]+\/[^/]+\/./.test(v), `${k}: "${v}" isn't a sharing link or a /sites/<site>/<library>/<path>`);
  }
  // The sales sheet: a library itself (no folder: the one place the page may write), a plain .xlsx name, and Quote
  // Tracker's photo store.
  assert.equal(c.salesLibrary, '/sites/Reports/Shalam Sales tool');
  assert.match(c.salesLibrary, /^\/sites\/[^/]+\/[^/]+$/);
  assert.equal(c.salesFile, 'Available to sell.xlsx');
  assert.match(c.salesFile, /^[^/\\:*?"<>|]+\.xlsx$/);
  assert.equal(c.photos, '/sites/Production/Shared Documents/Quote Sheets/Quote Images');
  assert.match(c.errorEmailTo, /^[^@\s]+@[^@\s]+\.[^@\s]+$/);
  assert.ok(String(c.errorEmailName || '').trim(), 'errorEmailName is blank');
});

test('no settings file, Sources popup, marker files or admin list left in the page', () => {
  for (const f of ['app.js', 'index.html', 'app.css']) {
    assert.doesNotMatch(text(f), /settings\.json|btnSources|openSources|saveSources|ensureFolder|Inventory\/Errors|admins|<dialog/i, f);
  }
});

test('index.html is a whole document that loads config.js first and app.js last; ExcelJS only when a sheet is built', () => {
  const h = text('index.html');
  assert.match(h, /^<!doctype html>\s*<html lang="en">\s*<head>[^]*<\/head>\s*<body>[^]*<\/body>\s*<\/html>\s*$/i);
  assert.match(h, /<title>M1S · Shalam Tool<\/title>/);
  const scripts = [...h.matchAll(/<script src="([^"]+)"><\/script>/g)].map(m => m[1]);
  assert.equal(scripts[0], 'config.js');
  assert.equal(scripts[scripts.length - 1], 'app.js');
  assert.equal(scripts.indexOf('core/salesSheet.js'), scripts.indexOf('core/sell.js') + 1);
  assert.ok(scripts.indexOf('core/common.js') < scripts.indexOf('core/salesSheet.js'));
  assert.ok(!scripts.some(s => /exceljs/i.test(s)), 'ExcelJS is loaded lazily, not at the start');
  assert.match(slice('  function loadExcelJS(', '\n  }\n'), /s\.src = 'vendor\/exceljs\/exceljs\.min\.js';/);
  assert.match(app, /\['InvSalesSheet', 'core\/salesSheet\.js'\]/);
  assert.doesNotMatch(h, /<!--/, 'a comment marker in index.html');
});
