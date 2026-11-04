// Shalam Sales (Saul, Nov 4): the page a salesperson opens. It reads only the two sales sheets Shalam Tool saves into
// Shalam Sales tool (config.js salesLibrary: salesFile and chinaFile), shows them with their photos and filters them
// (brand, category, sub category, a search, on hand or in transit, available by a date: core/salesView.js). Download
// makes a file of exactly what is shown, in this browser; nothing is saved anywhere. The page never writes: any request
// but a GET is refused before it leaves the page. A salesperson needs read access to that one library and nothing else.
(function () {
  'use strict';

  const CFG = window.INV_CONFIG || {};
  const SCOPES = ['User.Read', 'Sites.Read.All'];
  const GRAPH = 'https://graph.microsoft.com/v1.0';
  const HOST = 'met1sol.sharepoint.com';
  const REFRESH_MS = 5 * 60000;
  // Saved more than this long ago, the files are called out of date (Shalam Tool saves them every 5 minutes while open).
  const STALE_MS = 24 * 3600000;
  // This page's own address is its sign-in redirect (registered in Entra next to Shalam Tool's).
  const REDIRECT = location.origin + location.pathname;

  // MSAL v5 finishes popup and hidden-frame sign-ins through its bridge script: this copy of the page only hands the
  // reply back to the main window and does nothing else.
  const framed = window.self !== window.top;
  const authReply = /[?#&](code|error)=/.test(location.href) &&
    (framed || (!!window.opener && /^msal\./.test(window.name || '')));
  if (authReply) {
    document.documentElement.classList.add('authreply');
    const s = document.createElement('script');
    s.src = 'vendor/msal/msal-redirect-bridge.min.js';
    s.onload = () => { try { Promise.resolve(window.msalRedirectBridge.broadcastResponseToMainFrame()).catch(() => {}); } catch (e) { /* the main window times out and says so */ } };
    document.head.appendChild(s);
    return;
  }

  const SHEETS = [
    { kind: 'sales', label: 'On hand ATS', file: () => String(CFG.salesFile || '').trim() },
    { kind: 'china', label: 'Pull from China', file: () => String(CFG.chinaFile || '').trim() },
  ];
  const $ = id => document.getElementById(id);
  const esc = s => String(s ?? '').replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
  const fmtN = n => (typeof n === 'number' && Number.isFinite(n) ? n.toLocaleString('en-US') : '');
  const plural = (n, w) => `${fmtN(n)} ${w}${n === 1 ? '' : 's'}`;
  const isDay = v => typeof v === 'number' && Number.isFinite(v);
  const safeDecode = s => { try { return decodeURIComponent(s); } catch (e) { return s; } };
  const tail = u => safeDecode(String(u || '')).replace(/\/+$/, '').split('/').pop();
  const encPath = p => p.split('/').filter(Boolean).map(encodeURIComponent).join('/');
  const abortErr = () => new DOMException('Cancelled', 'AbortError');
  const isAbort = e => !!e && e.name === 'AbortError';
  function fmtLong(iso) {
    const d = new Date(iso);
    return iso && !isNaN(d) ? d.toLocaleString('en-US', { month: 'numeric', day: 'numeric', year: 'numeric', hour: 'numeric', minute: '2-digit' }) : '';
  }
  const store = {
    get(k) { try { return localStorage.getItem('sales:' + k); } catch (e) { return null; } },
    set(k, v) { try { localStorage.setItem('sales:' + k, v); } catch (e) { /* storage blocked: nothing kept */ } },
  };
  function problem(message, extra) { return Object.assign(new Error(message), extra || {}); }
  function sleep(ms, signal) {
    return new Promise((resolve, reject) => {
      if (signal && signal.aborted) { reject(abortErr()); return; }
      const t = setTimeout(resolve, ms);
      if (signal) signal.addEventListener('abort', () => { clearTimeout(t); reject(abortErr()); }, { once: true });
    });
  }

  // ---------- sign-in ----------
  let pca = null, account = null;
  function showSignin(msg, disable) {
    $('app').hidden = true; $('signin').hidden = false;
    if (msg) $('signinMsg').textContent = msg;
    $('btnSignin').disabled = !!disable;
  }
  async function initAuth() {
    const M = window.msal;
    if (!CFG.clientId) { showSignin("This page isn't set up yet: the app's client ID in config.js is blank.", true); return; }
    if (!M || !M.PublicClientApplication) { showSignin("The sign-in library didn't load. Check the connection and reload the page.", true); return; }
    let replyErr = null;
    try {
      pca = new M.PublicClientApplication({
        auth: { clientId: CFG.clientId, authority: 'https://login.microsoftonline.com/' + CFG.tenantId, redirectUri: REDIRECT, postLogoutRedirectUri: REDIRECT },
        cache: { cacheLocation: 'localStorage' },
      });
      if (pca.initialize) await pca.initialize();
      const res = await pca.handleRedirectPromise().catch(e => { console.error(e); replyErr = e; return null; });
      account = (res && res.account) || (pca.getActiveAccount && pca.getActiveAccount()) || pca.getAllAccounts()[0] || null;
    } catch (e) {
      console.error(e);
      showSignin("Sign-in couldn't start: " + (e.message || e), true);
      return;
    }
    if (account) { if (pca.setActiveAccount) pca.setActiveAccount(account); start(); }
    else showSignin(replyErr ? "Sign-in didn't finish: " + (replyErr.errorMessage || replyErr.message || replyErr) : null);
  }
  async function signIn() {
    try {
      if (framed) {
        const r = await pca.loginPopup({ scopes: SCOPES, prompt: 'select_account' });
        account = r.account; pca.setActiveAccount(account); start();
      } else await pca.loginRedirect({ scopes: SCOPES, redirectStartPage: location.href, prompt: 'select_account' });
    } catch (e) { $('signinMsg').textContent = "Sign-in didn't finish: " + (e.errorMessage || e.message || e); }
  }
  function signOut() {
    const req = { account, postLogoutRedirectUri: REDIRECT };
    Promise.resolve(framed ? pca.logoutPopup(req) : pca.logoutRedirect(req)).catch(e => toast("Couldn't sign out: " + (e.message || e), 6000));
  }
  function needsInteraction(e) {
    const I = window.msal && window.msal.InteractionRequiredAuthError;
    if (I && e instanceof I) return true;
    return /interaction_required|consent_required|login_required|no_tokens_found|InteractionRequired/i.test(String((e && (e.errorCode || e.name)) || '') + ' ' + String((e && e.message) || ''));
  }
  // An access token. When it can't be renewed quietly, sign in again (by redirect the page leaves, so this never settles).
  async function token(force) {
    try {
      return (await pca.acquireTokenSilent({ scopes: SCOPES, account, forceRefresh: !!force })).accessToken;
    } catch (e) {
      if (!needsInteraction(e)) throw problem("Your sign-in couldn't be renewed. Reload the page; if it keeps happening, sign out and in again.", { auth: true });
      if (framed) return (await pca.acquireTokenPopup({ scopes: SCOPES, account })).accessToken;
      await pca.acquireTokenRedirect({ scopes: SCOPES, account, redirectStartPage: location.href });
      return new Promise(() => {});
    }
  }

  // ---------- Graph: reads only ----------
  // The guard: anything but a GET never leaves the page.
  const refusal = method => {
    const m = String(method || 'GET').toUpperCase();
    return m === 'GET' ? '' : `Shalam Sales only reads, so it refused a ${m} request.`;
  };
  // 401: new token once. 429/503: wait for Retry-After, up to 3 times. Errors carry .status.
  async function g(url, opt = {}, tries = { auth: 0, busy: 0 }) {
    const refused = refusal(opt.method);
    if (refused) throw new Error(refused);
    const t = await token(tries.auth > 0);
    const r = await fetch(/^https?:/.test(url) ? url : GRAPH + url, { ...opt, method: 'GET', headers: { Authorization: 'Bearer ' + t } });
    if (r.status === 401 && tries.auth === 0) return g(url, opt, { ...tries, auth: 1 });
    if ((r.status === 429 || r.status === 503) && tries.busy < 3) {
      const ra = parseInt((r.headers && r.headers.get('Retry-After')) || '', 10);
      await sleep(Math.min(60, ra > 0 ? ra : 2 ** (tries.busy + 1)) * 1000, opt.signal);
      return g(url, opt, { ...tries, busy: tries.busy + 1 });
    }
    if (!r.ok) {
      let msg = '';
      try { const j = await r.json(); msg = j.error && (j.error.message || j.error.code); } catch (e) { /* no body */ }
      throw Object.assign(new Error(`SharePoint answered ${r.status}${msg ? ': ' + msg : ''}`), { status: r.status });
    }
    return r;
  }
  const gj = async (url, opt) => (await g(url, opt)).json();
  async function gAll(url, signal) {
    const out = [];
    while (url) {
      const j = await gj(url, { signal });
      out.push(...((j && j.value) || []));
      url = j && j['@odata.nextLink'];
    }
    return out;
  }
  // File contents only through the pre-signed download link (no token: the link carries its own key).
  async function download(meta, signal) {
    const url = meta && meta['@microsoft.graph.downloadUrl'];
    if (!url) throw problem(`SharePoint gave no download link for ${(meta && meta.name) || 'the file'}.`);
    for (let busy = 0; ; busy++) {
      const r = await fetch(url, { signal, cache: 'no-store' });
      if ((r.status === 429 || r.status === 503) && busy < 3) { await sleep(2 ** (busy + 1) * 1000, signal); continue; }
      if (!r.ok) throw Object.assign(new Error(`The download failed (${r.status}).`), { status: r.status });
      return r.arrayBuffer();
    }
  }
  // The library (config.js salesLibrary, a /sites/<site>/<library>): its drive id. One the account can't open isn't
  // among the site's libraries at all.
  async function salesDrive(signal) {
    const text = String(CFG.salesLibrary || '').trim(), parts = text.split('/').filter(Boolean);
    if (parts.length !== 3 || parts[0].toLowerCase() !== 'sites') throw problem("This page isn't set up: salesLibrary in config.js isn't a /sites/<site>/<library>. Tell Saul.");
    const [, site, lib] = parts;
    let sid;
    try { sid = (await gj(`/sites/${HOST}:/sites/${encodeURIComponent(site)}?$select=id`, { signal })).id; }
    catch (e) {
      if (e.status === 404 || e.status === 403) throw problem(`Your account can't open ${lib}. Ask Saul for access.`, { noAccess: true });
      throw e;
    }
    const drives = await gAll(`/sites/${sid}/drives?$select=id,name,webUrl`, signal);
    const d = drives.find(x => tail(x.webUrl).toLowerCase() === lib.toLowerCase()) || drives.find(x => String(x.name || '').toLowerCase() === lib.toLowerCase());
    if (!d) throw problem(`Your account can't open ${lib}. Ask Saul for access.`, { noAccess: true });
    return d.id;
  }

  // ExcelJS (vendor/exceljs): loaded the first time a file is read.
  let excelLoading = null;
  function loadExcelJS() {
    if (window.ExcelJS) return Promise.resolve(window.ExcelJS);
    if (!excelLoading) {
      excelLoading = new Promise((resolve, reject) => {
        const s = document.createElement('script');
        s.src = 'vendor/exceljs/exceljs.min.js';
        s.onload = () => (window.ExcelJS ? resolve(window.ExcelJS) : reject(new Error("The Excel library loaded but didn't start.")));
        s.onerror = () => reject(new Error("The Excel library (vendor/exceljs) didn't load. Check the connection and try again."));
        document.head.appendChild(s);
      }).catch(e => { excelLoading = null; throw e; });
    }
    return excelLoading;
  }

  // ---------- the two files ----------
  // Per kind: { meta (the file as SharePoint has it), sheet (InvSalesSheet.readWorkbook's), urls (each row's photo
  // address, by photo key), err (why there's none, in plain words) }.
  const data = { sales: null, china: null };
  let loadCtl = null, refreshTimer = null;
  async function loadAll() {
    if (loadCtl) return;
    const ctl = loadCtl = new AbortController(), signal = ctl.signal;
    clearTimeout(refreshTimer);
    $('app').setAttribute('aria-busy', 'true');
    busy('Opening Shalam Sales tool…');
    try {
      const drive = await salesDrive(signal);
      for (const K of SHEETS) {
        const file = K.file();
        if (!file) { setData(K.kind, { err: `${K.label} isn't set up yet (no file name in config.js).` }); continue; }
        busy(`Reading ${K.label}…`);
        let meta;
        try { meta = await gj(`/drives/${drive}/root:/${encPath(file)}`, { signal }); }
        catch (e) {
          if (e.status === 404) { setData(K.kind, { err: `${K.label} hasn't been saved into Shalam Sales tool yet.` }); continue; }
          throw e;
        }
        const was = data[K.kind];
        if (was && was.sheet && was.meta && was.meta.eTag === meta.eTag) { was.meta = meta; continue; }
        const bytes = await download(meta, signal);
        const E = await loadExcelJS();
        const wb = new E.Workbook();
        await wb.xlsx.load(bytes);
        if (signal.aborted) throw abortErr();
        const sheet = InvSalesSheet.readWorkbook(wb);
        if (sheet.kind !== K.kind) throw problem(`${file} in Shalam Sales tool isn't the ${K.label} sheet. Tell Saul.`);
        setData(K.kind, { meta, sheet });
      }
      $('health').innerHTML = healthText();
    } catch (e) {
      if (isAbort(e)) { toast('Stopped. The page shows what it read before.'); }
      else {
        console.error(e);
        const msg = /failed to fetch|networkerror|load failed/i.test(String(e.message || '')) ? "Couldn't reach SharePoint. Check the connection; the page tries again in a few minutes." : e.message || String(e);
        for (const K of SHEETS) if (!data[K.kind] || !data[K.kind].sheet) setData(K.kind, { err: msg });
        $('health').innerHTML = `<span class="r">${esc(msg)}</span>`;
      }
    } finally {
      loadCtl = null;
      busy('');
      $('app').setAttribute('aria-busy', 'false');
      render();
      refreshTimer = setTimeout(loadAll, REFRESH_MS);
    }
  }
  // A file's data, its photos as addresses this page made from the file (one per style, the old ones let go).
  function setData(kind, d) {
    const was = data[kind];
    if (was && was.urls) for (const u of was.urls.values()) URL.revokeObjectURL(u);
    const urls = new Map();
    if (d.sheet) d.sheet.rows.forEach((r, i) => {
      const p = d.sheet.photos[i];
      if (p && !urls.has(r.photoKey)) urls.set(r.photoKey, URL.createObjectURL(new Blob([p.data], { type: `image/${p.extension === 'jpeg' ? 'jpeg' : p.extension}` })));
    });
    data[kind] = { ...d, urls };
  }
  function healthText() {
    return SHEETS.filter(K => data[K.kind] && data[K.kind].meta).map(K => {
      const at = data[K.kind].meta.lastModifiedDateTime, old = at && Date.now() - new Date(at).getTime() > STALE_MS;
      return `<span class="${old ? 'a' : ''}">${esc(K.label)} saved <b>${esc(fmtLong(at))}</b>${old ? ' (not saved since: the numbers may be out of date)' : ''}</span>`;
    }).join(' · ');
  }
  function busy(msg) {
    $('busy').hidden = !msg;
    $('busyMsg').textContent = msg || '';
  }

  // ---------- filters and the table ----------
  const V = window.InvSalesView || { ALL: '*' };
  let tab = 'sales';
  const f = { q: '', brand: V.ALL, category: V.ALL, sub: V.ALL, where: 'all', by: null };
  const dayOfInput = v => { const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(v || ''); return m ? InvCommon.ymd(+m[1], +m[2], +m[3]) : null; };
  const fmtDay = d => InvCommon.fmtDay(d);
  function fillSelect(el, all, list, none, value) {
    const opts = [[V.ALL, all], ...list.map(v => [v, v || none])];
    el.innerHTML = opts.map(([v, t]) => `<option value="${esc(v)}">${esc(t)}</option>`).join('');
    const keep = opts.some(([v]) => v === value) ? value : V.ALL;
    el.value = keep;
    return keep;
  }
  function viewNow() {
    const d = data[tab];
    return d && d.sheet ? V.view(d.sheet, { ...f, where: tab === 'sales' ? f.where : 'all' }) : null;
  }
  const whenText = (w, est) => (w === 'now' ? 'now' : isDay(w) ? (est ? `about ${fmtDay(w)}` : fmtDay(w)) : 'date to come');
  const COLS = {
    sales: [
      ['Photo'], ['Style', 'fz'], ['Description'], ['Brand'], ['Category'], ['Sub category'], ['Master/Inner'],
      ['On hand', 'num'], ['In transit', 'num'], ['Available from'], ['Total', 'num'], ['Notes'],
    ],
    china: [['Photo'], ['Style', 'fz'], ['Description'], ['Brand'], ['Category'], ['Sub category'], ['Master/Inner'], ['Units', 'num'], ['Pull from date'], ['Notes']],
  };
  function rowHtml(r, urls) {
    const u = urls.get(r.photoKey);
    const pic = u ? `<img class="pic big" src="${esc(u)}" alt="" loading="lazy">` : '<span class="pic big none">no photo</span>';
    const num = n => `<td class="num">${n ? fmtN(n) : ''}</td>`;
    const head = `<td>${pic}</td><td class="fz"><b>${esc(r.style)}</b>${r.upc ? `<div class="upc">${esc(r.upc)}</div>` : ''}</td><td class="desc">${esc(r.desc)}</td>` +
      `<td>${esc(r.brand)}</td><td>${esc(r.category)}</td><td>${esc(r.sub)}</td><td class="nw">${esc(r.pack)}</td>`;
    const note = `<td>${r.note ? `<span class="pill warn">${esc(r.note)}</span>` : ''}</td>`;
    if (tab === 'china') return `<tr>${head}${num(r.units)}<td class="nw">${esc(whenText(r.pullFrom, r.estimate))}</td>${note}</tr>`;
    return `<tr>${head}${num(r.onHand)}${num(r.inTransit)}<td class="nw">${r.inTransit > 0 ? esc(whenText(r.transitFrom, false)) : ''}</td>${num(r.total)}${note}</tr>`;
  }
  function render() {
    for (const b of document.querySelectorAll('#tabs .tab')) b.classList.toggle('on', b.dataset.tab === tab);
    for (const K of SHEETS) {
      const d = data[K.kind], c = $(K.kind === 'sales' ? 'cSales' : 'cChina');
      c.hidden = !(d && d.sheet); c.textContent = d && d.sheet ? fmtN(d.sheet.rows.length) : '';
    }
    const d = data[tab], K = SHEETS.find(x => x.kind === tab);
    $('lead').hidden = !!(d && d.sheet);
    $('fbar').hidden = $('pane').hidden = !(d && d.sheet);
    if (!(d && d.sheet)) {
      $('leadH').textContent = d ? `${K.label} can't be shown` : `Loading ${K.label}…`;
      $('leadP').textContent = d ? d.err || '' : '';
      renderDl();
      return;
    }
    const ch = V.choices(d.sheet);
    f.brand = fillSelect($('fBrand'), 'All brands', ch.brands, 'No brand', f.brand);
    f.category = fillSelect($('fCat'), 'All categories', ch.categories, 'No category', f.category);
    f.sub = fillSelect($('fSub'), 'All sub categories', ch.subs[f.category] || ch.subs[V.ALL], 'No sub category', f.sub);
    $('fWhere').hidden = tab !== 'sales';
    $('btnClearF').hidden = !V.filtered(f);
    const v = viewNow(), cols = COLS[tab];
    $('grid').querySelector('thead').innerHTML = `<tr>${cols.map(([t, cls]) => `<th${cls ? ` class="${cls}"` : ''}>${esc(t)}</th>`).join('')}</tr>`;
    $('grid').querySelector('tbody').innerHTML = v.rows.length ? v.rows.map(r => rowHtml(r, d.urls)).join('')
      : `<tr><td class="empty" colspan="${cols.length}">Nothing matches these filters.</td></tr>`;
    const t = v.totals;
    $('fInfo').textContent = `${plural(v.counts.styles, 'style')} · ${plural(v.counts.rows, 'line')} · ` +
      (tab === 'sales' ? `${fmtN(t.onHand)} on hand · ${fmtN(t.inTransit)} in transit · ${fmtN(t.total)} units` : `${fmtN(t.units)} units`);
    renderDl(v);
  }

  // ---------- Download: what is shown, made here ----------
  let dl = null;
  function renderDl(v) {
    const K = SHEETS.find(x => x.kind === tab), btn = $('btnDl');
    if (v === undefined) v = viewNow();
    btn.disabled = !!dl || !v || !v.rows.length;
    btn.classList.toggle('china', tab === 'china');
    $('btnDlText').textContent = `Download ${K.label}${v && V.filtered(f) ? ' (filtered)' : ''}`;
    btn.title = v ? `${plural(v.counts.rows, 'line')} as shown, with their photos` : '';
    $('btnDlX').hidden = !dl;
    $('dlMsg').textContent = dl ? 'Making the file…' : '';
  }
  async function downloadNow() {
    const v = viewNow(), K = SHEETS.find(x => x.kind === tab);
    if (dl || !v || !v.rows.length) return;
    const me = dl = { stop: false };
    renderDl(v);
    try {
      const E = await loadExcelJS();
      const wb = InvSalesSheet.buildWorkbook(E, { ok: true, kind: v.kind, rows: v.rows, byDate: v.byDate }, v.photos);
      const bytes = await InvSalesSheet.workbookBytes(wb);
      if (me.stop) return;
      const name = `${K.label}${V.filtered(f) ? ' (filtered)' : ''}.xlsx`;
      const url = URL.createObjectURL(new Blob([bytes], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }));
      const a = document.createElement('a');
      a.href = url; a.download = name;
      document.body.appendChild(a); a.click(); a.remove();
      setTimeout(() => URL.revokeObjectURL(url), 60000);
      toast(`${name} downloaded: ${plural(v.counts.rows, 'line')}, ${plural(v.photos.filter(Boolean).length, 'photo')}.`, 5000);
    } catch (e) {
      console.error(e);
      if (!me.stop) toast(`The file couldn't be made: ${e.message || e}`, 8000);
    } finally {
      if (dl === me) dl = null;
      renderDl();
    }
  }

  function toast(t, ms = 3500) {
    const el = $('toast');
    el.textContent = t; el.style.display = 'block';
    clearTimeout(el._t); el._t = setTimeout(() => { el.style.display = 'none'; }, ms);
  }

  // ---------- start ----------
  function start() {
    $('signin').hidden = true; $('app').hidden = false;
    $('btnRefresh').hidden = false; $('who').hidden = false;
    $('who').innerHTML = `<b>${esc(account.name || account.username || '')}</b><button type="button" class="ghost" id="btnSignout" title="Sign out">Sign out</button>`;
    $('btnSignout').onclick = signOut;
    const t = store.get('tab');
    if (SHEETS.some(K => K.kind === t)) tab = t;
    render();
    loadAll();
  }
  function bindUi() {
    $('btnSignin').onclick = signIn;
    $('btnRefresh').onclick = () => loadAll();
    $('btnCancelBusy').onclick = () => { if (loadCtl) loadCtl.abort(); };
    $('tabs').addEventListener('click', e => {
      const b = e.target.closest('.tab');
      if (!b || b.dataset.tab === tab) return;
      tab = b.dataset.tab; store.set('tab', tab);
      render();
    });
    let typing = null;
    $('fq').addEventListener('input', () => { clearTimeout(typing); typing = setTimeout(() => { f.q = $('fq').value; render(); }, 200); });
    $('fBrand').onchange = () => { f.brand = $('fBrand').value; render(); };
    $('fCat').onchange = () => { f.category = $('fCat').value; f.sub = V.ALL; render(); };
    $('fSub').onchange = () => { f.sub = $('fSub').value; render(); };
    $('fWhere').onchange = () => { f.where = $('fWhere').value; render(); };
    $('fBy').onchange = () => { f.by = dayOfInput($('fBy').value); render(); };
    $('btnClearF').onclick = () => {
      Object.assign(f, { q: '', brand: V.ALL, category: V.ALL, sub: V.ALL, where: 'all', by: null });
      $('fq').value = ''; $('fWhere').value = 'all'; $('fBy').value = '';
      render();
    };
    $('btnDl').onclick = downloadNow;
    $('btnDlX').onclick = () => { if (dl) { dl.stop = true; dl = null; renderDl(); toast('Download cancelled.'); } };
  }
  // Drifting dots behind the glass, as on Shalam Tool; still under prefers-reduced-motion.
  function startBg() {
    const c = $('bg'), x = c && c.getContext && c.getContext('2d');
    if (!x) return;
    let seed = 20261104;
    const rnd = () => { seed |= 0; seed = seed + 0x6D2B79F5 | 0; let t = Math.imul(seed ^ seed >>> 15, 1 | seed); t = t + Math.imul(t ^ t >>> 7, 61 | t) ^ t; return ((t ^ t >>> 14) >>> 0) / 4294967296; };
    const P = Array.from({ length: 70 }, () => ({ x: rnd(), y: rnd(), r: 1 + rnd() * 2.4, vx: (rnd() - 0.5) * 0.00016, vy: (rnd() - 0.5) * 0.00016, h: rnd() < 0.5 ? '47,155,214' : '138,111,209' }));
    const still = window.matchMedia && matchMedia('(prefers-reduced-motion: reduce)').matches;
    function draw() {
      c.width = innerWidth; c.height = innerHeight;
      x.clearRect(0, 0, c.width, c.height);
      for (const p of P) {
        if (!still) { p.x = (p.x + p.vx + 1) % 1; p.y = (p.y + p.vy + 1) % 1; }
        x.beginPath(); x.arc(p.x * c.width, p.y * c.height, p.r, 0, 7); x.fillStyle = `rgba(${p.h},.45)`; x.fill();
      }
      if (!still) requestAnimationFrame(draw);
    }
    draw();
    addEventListener('resize', () => { if (still) draw(); });
  }
  function boot() {
    bindUi();
    startBg();
    const missing = [['InvCommon', 'core/common.js'], ['InvSalesSheet', 'core/salesSheet.js'], ['InvSalesView', 'core/salesView.js']].filter(([n]) => !window[n]).map(p => p[1]);
    if (missing.length) { showSignin(`Part of the page didn't load (${missing.join(', ')}). Reload the page; if it keeps happening, tell Saul.`, true); return; }
    initAuth();
  }
  boot();
})();
