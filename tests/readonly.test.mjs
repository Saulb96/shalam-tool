// The page must stay read-only: the only request it may make that isn't a GET is the error email.
// These read the page's own files as text; the headless check watches the requests it really makes.
import test from 'node:test';
import assert from 'node:assert/strict';
import { readFileSync, readdirSync } from 'node:fs';
import vm from 'node:vm';
import { root } from './load.mjs';

const text = p => readFileSync(root(p), 'utf8');
const app = text('app.js');
const pageFiles = ['app.js', 'index.html', 'config.js', ...readdirSync(root('core')).filter(f => f.endsWith('.js')).map(f => 'core/' + f)];

test('the one request method in the page is the POST that sends the error email', () => {
  const lines = app.split('\n');
  const methods = lines.map((l, i) => [i + 1, l]).filter(([, l]) => /\bmethod\s*:/.test(l));
  assert.equal(methods.length, 1, methods.map(([n, l]) => `app.js:${n} ${l.trim()}`).join('\n'));
  assert.match(methods[0][1], /g\('\/me\/sendMail', \{ method: 'POST'/);
  for (const f of pageFiles) {
    assert.doesNotMatch(text(f), /\b(PUT|PATCH|DELETE)\b/, `${f} names a write method`);
    assert.doesNotMatch(text(f), /XMLHttpRequest|sendBeacon|<form\b|conflictBehavior|:\/content\b/, `${f} has another way to write`);
  }
  // Only app.js talks to the network; the parsers never do.
  for (const f of pageFiles.filter(x => x.startsWith('core/'))) assert.doesNotMatch(text(f), /\bfetch\s*\(/, f);
});

test('the Graph helper refuses anything but a GET or the email POST, before asking for a token', () => {
  const g = app.slice(app.indexOf('async function g(url'), app.indexOf('const t = await token(tries.auth > 0);'));
  assert.match(g, /if \(method !== 'GET' && !\(method === 'POST' && url === '\/me\/sendMail'\)\) throw/);
});

test('sign-in asks for read permissions only, plus Mail.Send', () => {
  const m = /const SCOPES = (\[[^\]]*\]);/.exec(app);
  assert.ok(m, 'SCOPES not found');
  assert.deepEqual(JSON.parse(m[1].replace(/'/g, '"')), ['User.Read', 'Sites.Read.All', 'Mail.Send']);
  // The stop rule: other than Mail.Send, any scope with Write, Manage or FullControl in it, or that isn't a plain read
  // (or a sign-in scope); and Sites.Read.All required.
  assert.match(app, /!\/\^Mail\\\.Send\$\/i\.test\(x\) && \(\/Write\|Manage\|FullControl\/i\.test\(x\) \|\| !\(isRead\(x\) \|\| OIDC\.test\(x\)\)\)/);
  assert.match(app, /const isRead = x => x\.split\('\.'\)\.some\(p => \/\^Read\(Basic\)\?\$\/i\.test\(p\)\);/);
  assert.match(app, /Shalam Tool must be read-only, but its sign-in still allows changes to SharePoint \(\$\{writes\.join\(', '\)\}\)\. In Entra, remove that permission from the Shalam Tool app and grant consent again\./);
  assert.match(app, /\^Sites\\\.Read\\\.All\$/);
});

test('the stop rule, run on permission lists (from the token, or from the sign-in service when the token is opaque)', () => {
  const from = app.indexOf('  const OIDC = '), endMark = "    return '';\n  }", to = app.indexOf(endMark, from) + endMark.length;
  assert.ok(from > 0 && to > from, 'scopeProblem not found in app.js');
  const ctx = { atob: s => Buffer.from(s, 'base64').toString('binary'), mailScopeMissing: null };
  vm.runInNewContext(app.slice(from, to), ctx);
  const jwt = scp => ['{"alg":"none"}', JSON.stringify({ scp })].map(x => Buffer.from(x).toString('base64url')).join('.') + '.sig';
  const check = (scp, opaque) => ctx.scopeProblem(opaque ? { accessToken: 'opaque', scopes: scp.split(' ').filter(Boolean) } : { accessToken: jwt(scp), scopes: [] });
  for (const opaque of [false, true]) {
    assert.equal(check('email Mail.Send openid profile Sites.Read.All User.Read offline_access', opaque), '');
    assert.equal(ctx.mailScopeMissing, false);
    assert.equal(check('User.ReadBasic.All Sites.Read.All Mail.ReadBasic', opaque), '');
    assert.equal(ctx.mailScopeMissing, true);
    for (const [scp, named] of [['User.Read Sites.Read.All Sites.ReadWrite.All Mail.Send', 'Sites.ReadWrite.All'],
      ['User.Read Sites.Read.All Files.ReadWrite Mail.Send', 'Files.ReadWrite'], ['Sites.Read.All Sites.FullControl.All', 'Sites.FullControl.All'],
      ['Sites.Read.All Sites.Selected Mail.Send', 'Sites.Selected'], ['Sites.Read.All ListItems.SelectedOperations.Selected', 'ListItems.SelectedOperations.Selected'],
      ['Sites.Read.All Mail.Send.Shared Mail.Send', 'Mail.Send.Shared'], ['Sites.Read.All Directory.AccessAsUser.All', 'Directory.AccessAsUser.All'],
      ['Sites.Read.All Sites.Create.All', 'Sites.Create.All'], ['Sites.Read.All Files.ReadWrite.AppFolder', 'Files.ReadWrite.AppFolder']]) {
      const msg = check(scp, opaque);
      assert.match(msg, /^Shalam Tool must be read-only, but its sign-in still allows changes to SharePoint \(/, scp);
      assert.ok(msg.includes(`(${named})`), `${scp} -> ${msg}`);
    }
    assert.match(check('User.Read Mail.Send', opaque), /missing the Sites\.Read\.All permission \(it has: User\.Read, Mail\.Send\)/);
  }
  assert.match(check('', true), /couldn't read which permissions its sign-in has/);
  assert.match(ctx.scopeProblem({ accessToken: 'opaque' }), /couldn't read which permissions its sign-in has/);
});

test('config.js holds the sign-in, every source location and the error email', () => {
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
  assert.match(c.errorEmailTo, /^[^@\s]+@[^@\s]+\.[^@\s]+$/);
  assert.ok(String(c.errorEmailName || '').trim(), 'errorEmailName is blank');
});

test('no settings file, Sources popup, marker files or admin list left in the page', () => {
  for (const f of ['app.js', 'index.html', 'app.css']) {
    assert.doesNotMatch(text(f), /settings\.json|btnSources|openSources|saveSources|ensureFolder|Inventory\/Errors|admins|<dialog/i, f);
  }
});

test('index.html is a whole document that loads config.js first and app.js last', () => {
  const h = text('index.html');
  assert.match(h, /^<!doctype html>\s*<html lang="en">\s*<head>[^]*<\/head>\s*<body>[^]*<\/body>\s*<\/html>\s*$/i);
  assert.match(h, /<title>M1S · Shalam Tool<\/title>/);
  const scripts = [...h.matchAll(/<script src="([^"]+)"><\/script>/g)].map(m => m[1]);
  assert.equal(scripts[0], 'config.js');
  assert.equal(scripts[scripts.length - 1], 'app.js');
  assert.doesNotMatch(h, /<!--/, 'a comment marker in index.html');
});
