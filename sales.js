// Shalam Sales (Saul, Nov 4; made simple Nov 25: "his login just show him latest ATS import and the 2 download
// buttons"): the page a salesperson opens. It looks only at the two sales sheets Shalam Tool saves into Shalam Sales tool
// (config.js salesLibrary: salesFile and chinaFile), says which ATS report they were made from and when they were saved,
// and hands each file over exactly as it is in the library. Nothing of the sheets is shown on the page. The page never
// writes: any request but a GET is refused before it leaves the page. A salesperson needs read access to that one
// library and nothing else.
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

  // The spreadsheet library (vendor/xlsx-js-style): loaded the first time a file is read, only to read its properties.
  let xlsxLoading = null;
  function loadXlsx() {
    if (window.XLSX) return Promise.resolve(window.XLSX);
    if (!xlsxLoading) {
      xlsxLoading = new Promise((resolve, reject) => {
        const s = document.createElement('script');
        s.src = 'vendor/xlsx-js-style/xlsx.bundle.js';
        s.onload = () => (window.XLSX ? resolve(window.XLSX) : reject(new Error("The spreadsheet library loaded but didn't start.")));
        s.onerror = () => reject(new Error("The spreadsheet library didn't load."));
        document.head.appendChild(s);
      }).catch(e => { xlsxLoading = null; throw e; });
    }
    return xlsxLoading;
  }
  // The ATS report a file was made from, as Shalam Tool wrote it into the file's properties (null: not said, or the
  // file can't be read for it; the file is still handed over).
  async function reportDayIn(bytes) {
    try {
      const X = await loadXlsx();
      return InvSalesSheet.reportDayOf((X.read(new Uint8Array(bytes), { type: 'array', bookProps: true }).Props || {}).Subject);
    } catch (e) { console.error(e); return null; }
  }

  // ---------- the two files ----------
  // Per kind: { meta (the file as SharePoint has it), bytes (the file itself, as downloaded), reportDay (the ATS report
  // it was made from, or null), err (why there's none, in plain words) }.
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
        if (!file) { data[K.kind] = { err: `${K.label} isn't set up yet (no file name in config.js).` }; continue; }
        busy(`Getting ${K.label}…`);
        let meta;
        try { meta = await gj(`/drives/${drive}/root:/${encPath(file)}`, { signal }); }
        catch (e) {
          if (e.status === 404) { data[K.kind] = { err: `${K.label} hasn't been saved into Shalam Sales tool yet.` }; continue; }
          throw e;
        }
        const was = data[K.kind];
        if (was && was.bytes && was.meta && was.meta.eTag === meta.eTag) { was.meta = meta; continue; }
        const bytes = await download(meta, signal);
        const reportDay = await reportDayIn(bytes);
        if (signal.aborted) throw abortErr();
        data[K.kind] = { meta, bytes, reportDay };
        render();
      }
      $('health').textContent = '';
    } catch (e) {
      if (isAbort(e)) { toast('Stopped. The page keeps the files it got before.'); }
      else {
        console.error(e);
        const msg = /failed to fetch|networkerror|load failed/i.test(String(e.message || '')) ? "Couldn't reach SharePoint. Check the connection; the page tries again in a few minutes." : e.message || String(e);
        for (const K of SHEETS) if (!data[K.kind] || !data[K.kind].bytes) data[K.kind] = { err: msg };
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
  function busy(msg) {
    $('busy').hidden = !msg;
    $('busyMsg').textContent = msg || '';
  }

  // ---------- what the page shows: the latest ATS import, and a button for each file ----------
  const fmtDay = d => InvCommon.fmtDay(d);
  const EL = { sales: ['btnDlSales', 'infoSales'], china: ['btnDlChina', 'infoChina'] };
  function render() {
    const days = SHEETS.map(K => (data[K.kind] && data[K.kind].bytes ? data[K.kind].reportDay : null)).filter(isDay);
    const got = SHEETS.some(K => data[K.kind] && data[K.kind].bytes);
    $('impH').textContent = days.length ? `Latest ATS import: ${fmtDay(Math.max(...days))}` : 'Latest ATS import';
    $('impP').textContent = days.length === 2 && days[0] !== days[1] ? 'The two files are from different ATS reports: each says its own below.'
      : days.length ? 'The files below are made from that ATS report by Shalam Tool.'
        : got ? "The files don't say yet which ATS report they were made from. They will once Shalam Tool saves them again." : loadCtl ? 'Getting the files…' : '';
    for (const K of SHEETS) {
      const d = data[K.kind], [btn, info] = EL[K.kind].map($);
      btn.disabled = !(d && d.bytes);
      if (!d) { info.className = 'selinfo'; info.textContent = loadCtl ? 'Getting the file…' : ''; continue; }
      if (!d.bytes) { info.className = 'selinfo r'; info.textContent = d.err || ''; continue; }
      const at = d.meta.lastModifiedDateTime, old = at && Date.now() - new Date(at).getTime() > STALE_MS;
      info.className = 'selinfo' + (old ? ' a' : '');
      info.textContent = `${isDay(d.reportDay) ? `ATS report of ${fmtDay(d.reportDay)} · ` : ''}saved ${fmtLong(at)}${old ? ' (not saved since: the numbers may be out of date)' : ''}`;
    }
  }

  // ---------- Download: the file as it is in the library ----------
  function downloadNow(kind) {
    const d = data[kind], K = SHEETS.find(x => x.kind === kind);
    if (!d || !d.bytes) return;
    const name = d.meta.name || K.file();
    const url = URL.createObjectURL(new Blob([d.bytes], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }));
    const a = document.createElement('a');
    a.href = url; a.download = name;
    document.body.appendChild(a); a.click(); a.remove();
    setTimeout(() => URL.revokeObjectURL(url), 60000);
    toast(`${name} downloaded.`, 5000);
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
    loadAll();
    render();
  }
  function bindUi() {
    $('btnSignin').onclick = signIn;
    $('btnRefresh').onclick = () => loadAll();
    $('btnCancelBusy').onclick = () => { if (loadCtl) loadCtl.abort(); };
    $('btnDlSales').onclick = () => downloadNow('sales');
    $('btnDlChina').onclick = () => downloadNow('china');
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
    const missing = [['InvCommon', 'core/common.js'], ['InvSalesSheet', 'core/salesSheet.js']].filter(([n]) => !window[n]).map(p => p[1]);
    if (missing.length) { showSignin(`Part of the page didn't load (${missing.join(', ')}). Reload the page; if it keeps happening, tell Saul.`, true); return; }
    initAuth();
  }
  boot();
})();
