// Shalam Tool: sign-in, reading every source from SharePoint, the source chips, the failure rules, the error email
// and the Problems tab. Parsing lives in core/*.js; this file fetches, calls the parsers and draws.
// The page is read-only: the one request it makes that isn't a GET is the error email (POST /me/sendMail).
(function () {
  'use strict';

  const CFG = window.INV_CONFIG || {};
  const SCOPES = ['User.Read', 'Sites.Read.All', 'Mail.Send'];
  const GRAPH = 'https://graph.microsoft.com/v1.0';
  const HOST = 'met1sol.sharepoint.com';
  const SITE = HOST + ':/sites/Production';   // where the PO_Book_Lines and UPC_MASTER lists are
  const REFRESH_MS = 5 * 60000;
  const MAX_REPORT_DAYS = 7;   // forwarder and drayage files older than this aren't used
  const AMBER_BDAYS = 2;       // ... and are amber from this many business days old
  const ATS_AMBER_DAYS = 8, ATS_MAIL_DAYS = 10;
  const ISSUE_PAGE = 200;
  const NY = 'America/New_York';
  // A visit to .../index.html signs in as the folder address, so only one redirect address needs registering.
  const REDIRECT = location.origin + location.pathname.replace(/index\.html$/i, '');
  const POB_FIELDS = 'Title,LineKey,Kind,PONumber,Style,ParentKey,ShippedDate,ShippedBy,CRDOverride,Qty,Shipped,Data';
  const UPC_FIELDS = 'Title,UPC,Style,Description,Category,SubCategory,Brand';
  const XOPT = { type: 'array', cellFormula: false, cellHTML: false, cellStyles: false };

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

  // kind file / folder: its location (a SharePoint link or a /sites/… path) is in config.js under sources.<k>.
  const SOURCES = [
    { k: 'msl', name: 'Master Sheet Live', kind: 'file', sheet: 'Master Sheet Live', group: true },
    { k: 'tracker', name: 'PO Status Tracker', kind: 'file', sheet: 'PO Status Tracker', group: true },
    { k: 'pobook', name: 'PO Book lines', kind: 'list', list: 'PO_Book_Lines', group: true },
    { k: 'ct', name: 'Container Tracking', kind: 'file' },
    { k: 'forwarder', name: 'Forwarder', kind: 'folder', ext: '.xlsx' },
    { k: 'drayage', name: 'Drayage', kind: 'folder', ext: '.csv' },
    { k: 'ats', name: 'ATS', kind: 'folder' },
    { k: 'upc', name: 'UPC_MASTER', kind: 'list', list: 'UPC_MASTER' },
  ];
  const BY_K = Object.fromEntries(SOURCES.map(x => [x.k, x]));
  const CORE_KEYS = ['msl', 'tracker', 'pobook', 'ct'];   // red here: the stage numbers are not loaded
  const REPORT_KEYS = ['forwarder', 'drayage'];            // trouble here: amber banner, numbers still shown
  const locText = k => String((CFG.sources && CFG.sources[k]) || '').trim();

  const LABELS = { rowsRead: 'Rows read', sectionRows: 'Round / section header rows', styleRows: 'Style rows', counted: 'Counted',
    notCounted: 'Not counted (listed on Problems)', roundRows: 'Round rows', records: 'Records', containers: 'Containers', rows: 'Rows',
    styles: 'Styles', upcs: 'UPCs', badUpc: 'Bad UPCs (listed, not fixed)', noStyle: 'No style',
    upcOnly: 'Only a UPC (one line on Problems)', dupUpc: 'UPC on two styles',
    noUpc: 'No UPC', used: 'Items used', repeats: 'Same UPC and style twice', blank: 'Blank rows', blankRows: 'Blank rows',
    otherRows: 'Rows with no style (product lines listed on Problems)', blocks: 'POs', units: 'Units', rounds: 'Rounds',
    noId: 'No shipment ID', noContainer: 'No container number', duplicates: 'Repeated containers',
    labelRows: 'Title rows (like "Complete Shipments")', noRoundRows: 'Rows with no Round # (listed on Problems)' };
  // Words that differ by source.
  const SRC_LABELS = { forwarder: { records: 'Shipments' }, drayage: { records: 'Containers' } };

  // Plain words for the parsers' issue codes, shown on the Problems tab (the code stays in the tooltip and the filter).
  const CODE_LABELS = {
    // Master Sheet Live and the PO Status Tracker (the PO Book's rules)
    nopo: 'No PO number above it', badpo: "PO doesn't start with 6 digits", nostyle: 'No style', qty: 'Quantity blank or not a number',
    dup: 'Same PO and style twice', mslonly: 'Not on the PO Status Tracker', trkonly: 'On the PO Status Tracker only',
    dupround: 'Below the last Round, already shipped', 'style-alias': 'Style read as another style', 'po-text': 'More than one PO in the cell',
    // Container Tracking
    'round-no-number': 'Round header with no number', 'round-odd': 'Round header with other numbers', 'odd-section': 'Header row that may be a product line',
    'not-a-date': 'Not a date', 'yearless-date': 'Date with no year', 'date-far': 'Date far from today', 'po-two-sections': 'PO under two sections',
    'qty-no-style': 'Quantity with no style', 'no-style': 'Product line with no style', 'no-po': 'No PO number above it',
    'qty-not-number': 'Quantity blank or not a number', 'bad-upc': 'Barcode not 12-13 digits', 'no-round': 'No Round #',
    'round-not-number': 'Round # not a number', 'no-eta': 'No ETA yet', 'bad-container': 'No container number',
    'round-no-open-complete': 'Round not on Open or Complete', 'on-both-tabs': 'On Booking and Loaded, same round',
    'on-both-tabs-other-round': 'On Booking and Loaded, different sections',
    // Forwarder and drayage
    'bad-date': 'Not a date', 'bad-number': 'Not a number', 'bad-yes-no': 'Not Yes or No',
    'container-listed-twice': 'Container listed twice', 'container-on-two-shipments': 'Container on two shipments',
    'container-reused': 'Container number reused', 'duplicate-container': 'Container on two rows', 'duplicate-id': 'Shipment ID twice',
    'field-count': 'Row with too many or too few cells', 'no-availability': 'No availability', 'no-container': 'No container number',
    'no-id': 'No shipment ID', 'no-status': 'No status', 'several-containers': 'Several containers on one line',
    'unknown-availability': 'Availability not known', 'unknown-status': 'Status not known', 'newer-file-not-read': "Newer file couldn't be read",
    // UPC_MASTER and PO Book lines
    noupc: 'No UPC', badupc: 'UPC not 12-13 digits', dupupc: 'UPC on two styles', repeat: 'Same UPC and style twice',
    upconly: 'Rows with only a UPC',
    'no-linekey': 'Item with no LineKey',
  };
  const codeLabel = c => CODE_LABELS[c] || (c ? String(c).replace(/-/g, ' ') : '(no code)');

  // ---------- small helpers ----------
  const $ = id => document.getElementById(id);
  const esc = s => String(s ?? '').replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
  const fmtN = n => (typeof n === 'number' && Number.isFinite(n) ? n.toLocaleString('en-US') : '');
  const fmtV = v => (typeof v === 'number' ? fmtN(v) : String(v ?? ''));
  const abortErr = () => new DOMException('Cancelled', 'AbortError');
  const isAbort = e => !!e && e.name === 'AbortError';
  const yieldUi = () => new Promise(r => setTimeout(r, 0));
  const encPath = p => p.split('/').filter(Boolean).map(encodeURIComponent).join('/');
  const who = x => (x && x.user && (x.user.displayName || x.user.email)) || '';
  const safeDecode = s => { try { return decodeURIComponent(s); } catch (e) { return s; } };
  const tail = u => safeDecode(String(u || '')).replace(/\/+$/, '').split('/').pop();
  const firstNum = (o, keys) => { if (!o) return null; for (const k of keys) if (typeof o[k] === 'number') return o[k]; return null; };
  const sumOf = a => { const v = a.filter(x => typeof x === 'number'); return v.length ? v.reduce((p, x) => p + x, 0) : null; };
  const plural = (n, w) => `${fmtN(n)} ${w}${n === 1 ? '' : 's'}`;
  function sleep(ms, signal) {
    return new Promise((resolve, reject) => {
      if (signal && signal.aborted) return reject(abortErr());
      const t = setTimeout(resolve, ms);
      if (signal) signal.addEventListener('abort', () => { clearTimeout(t); reject(abortErr()); }, { once: true });
    });
  }
  function fmtShort(iso) {
    const d = new Date(iso);
    return iso && !isNaN(d) ? d.toLocaleString('en-US', { month: 'numeric', day: 'numeric', hour: 'numeric', minute: '2-digit' }) : '';
  }
  function fmtLong(iso) {
    const d = new Date(iso);
    return iso && !isNaN(d) ? d.toLocaleString('en-US', { month: 'numeric', day: 'numeric', year: 'numeric', hour: 'numeric', minute: '2-digit' }) : '';
  }
  // Browser storage, every key starting "inv:". set answers whether it was kept (false when storage is blocked).
  const store = {
    get(k) { try { return localStorage.getItem('inv:' + k); } catch (e) { return null; } },
    set(k, v) { try { localStorage.setItem('inv:' + k, v); return true; } catch (e) { return false; } },
    del(k) { try { localStorage.removeItem('inv:' + k); } catch (e) { /* storage blocked: nothing kept to remove */ } },
    // Removes the keys under prefix that keep() turns down (old email marks, locations no longer in config.js).
    prune(prefix, keep) {
      try {
        const p = 'inv:' + prefix, gone = [];
        for (let i = 0; i < localStorage.length; i++) { const k = localStorage.key(i); if (k && k.startsWith(p) && !keep(k.slice(p.length))) gone.push(k); }
        for (const k of gone) localStorage.removeItem(k);
      } catch (e) { /* storage blocked */ }
    },
  };
  const human = (k, own) => (own && own[k]) || LABELS[k] || cap(String(k).replace(/([a-z])([A-Z])/g, (m, a, b) => a + ' ' + b.toLowerCase()));
  // Number fields of a parser's counts object as [label, value] rows; nested objects (one per tab) get a heading.
  function countRows(obj, own) {
    const out = [];
    if (!obj || typeof obj !== 'object') return out;
    const nested = [];
    for (const [k, v] of Object.entries(obj)) {
      if (typeof v === 'number') out.push([human(k, own), v]);
      else if (v && typeof v === 'object' && !Array.isArray(v) && !(v instanceof Map) && !(v instanceof Set)) nested.push([k, v]);
    }
    for (const [k, v] of nested) { const sub = countRows(v, own); if (sub.length) out.push({ h: human(k, own) }, ...sub); }
    return out;
  }
  const dataRowCount = rows => (rows || []).slice(1).filter(r => r && r.some(v => v != null && window.InvCommon.S(v) !== '')).length;
  const headerList = rows => ((rows && rows[0]) || []).map(v => window.InvCommon.S(v)).filter(Boolean).join(' | ');

  // ---------- sign-in ----------
  let pca = null, account = null;

  function showSignin(msg, disable) {
    $('signin').hidden = false; $('app').hidden = true;
    if (msg) $('signinMsg').textContent = msg;
    $('btnSignin').disabled = !!disable;
  }

  async function initAuth() {
    const M = window.msal;
    if (!CFG.clientId) {
      showSignin("This page isn't set up yet: the app's client ID in config.js is blank. Saul adds it after registering the app in Entra.", true);
      return;
    }
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
      } else {
        await pca.loginRedirect({ scopes: SCOPES, redirectStartPage: location.href, prompt: 'select_account' });
      }
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
  // Any other failure (a timed-out renewal, no network to the sign-in service) is tagged .auth, so the page asks to sign
  // in again instead of marking the sources red.
  // The sign-in result (accessToken, scopes).
  async function rawToken(force) {
    try {
      return await pca.acquireTokenSilent({ scopes: SCOPES, account, forceRefresh: !!force });
    } catch (e) {
      if (!needsInteraction(e)) { if (e && typeof e === 'object') e.auth = true; throw e; }
      if (framed) return await pca.acquireTokenPopup({ scopes: SCOPES, account });
      await pca.acquireTokenRedirect({ scopes: SCOPES, account, redirectStartPage: location.href });
      return new Promise(() => {});
    }
  }
  // Every token is checked before it is used: one that allows changes, or can't read SharePoint, stops the page (.stop).
  // A cached token may predate a change in Entra, so a bad one is renewed once before stopping.
  async function token(force) {
    let r = await rawToken(force), bad = scopeProblem(r);
    if (bad && !force) { r = await rawToken(true); bad = scopeProblem(r); }
    if (bad) { scopeFatal = bad; throw Object.assign(new Error(bad), { stop: true }); }
    return r.accessToken;
  }

  // Shalam Tool must be read-only. Mail.Send (the error email) is the one permission allowed that isn't a read.
  // The permissions come from the token's scp claim; Microsoft calls Graph tokens opaque, so if one can't be read,
  // from the list the sign-in service returned with it (MSAL's scopes). The same rules apply to either.
  const OIDC = /^(openid|profile|email|offline_access)$/i;
  // A plain read: one part of its name is Read or ReadBasic (User.Read, Sites.Read.All, Mail.ReadBasic).
  const isRead = x => x.split('.').some(p => /^Read(Basic)?$/i.test(p));
  function scopeProblem(res) {
    let scp = null;
    try {
      const p = String(res && res.accessToken).split('.')[1];
      scp = JSON.parse(atob(p.replace(/-/g, '+').replace(/_/g, '/'))).scp;
    } catch (e) { /* not readable: the sign-in service's list below */ }
    if (typeof scp !== 'string' && res && Array.isArray(res.scopes) && res.scopes.length) scp = res.scopes.join(' ');
    if (typeof scp !== 'string') return "Shalam Tool couldn't read which permissions its sign-in has, so it can't make sure it is read-only and doesn't load anything. Click Sign in again; if this keeps happening, tell Saul.";
    const list = scp.split(/\s+/).filter(Boolean);
    // Anything with Write, Manage or FullControl in it, and anything else that isn't a plain read: Sites.Selected and
    // the *.SelectedOperations.Selected permissions name no Write but can carry write access.
    const writes = list.filter(x => !/^Mail\.Send$/i.test(x) && (/Write|Manage|FullControl/i.test(x) || !(isRead(x) || OIDC.test(x))));
    if (writes.length) return `Shalam Tool must be read-only, but its sign-in still allows changes to SharePoint (${writes.join(', ')}). In Entra, remove that permission from the Shalam Tool app and grant consent again.`;
    if (!list.some(x => /^Sites\.Read\.All$/i.test(x))) return `Shalam Tool's sign-in is missing the Sites.Read.All permission (it has: ${list.join(', ') || 'none'}), so it can't read SharePoint. In Entra, add Sites.Read.All to the Shalam Tool app and grant consent again.`;
    mailScopeMissing = !list.some(x => /^Mail\.Send$/i.test(x));
    return '';
  }

  // ---------- Graph ----------
  // 401: new token once. 429/503: wait for Retry-After, up to 3 times. Errors carry .status.
  // Read-only: a GET, or the one POST that sends the error email. Anything else is refused before it leaves the page.
  async function g(url, opt = {}, tries = { auth: 0, busy: 0 }) {
    const method = String(opt.method || 'GET').toUpperCase();
    if (method !== 'GET' && !(method === 'POST' && url === '/me/sendMail')) throw new Error(`Shalam Tool is read-only, so it refused a ${method} request.`);
    const t = await token(tries.auth > 0);
    const headers = { Authorization: 'Bearer ' + t, ...(typeof opt.body === 'string' ? { 'Content-Type': 'application/json' } : {}), ...(opt.headers || {}) };
    const r = await fetch(/^https?:/.test(url) ? url : GRAPH + url, { ...opt, headers });
    if (r.status === 401 && tries.auth === 0) return g(url, opt, { ...tries, auth: 1 });
    if ((r.status === 429 || r.status === 503) && tries.busy < 3) {
      const ra = parseInt((r.headers && r.headers.get('Retry-After')) || '', 10);
      await sleep(Math.min(60, ra > 0 ? ra : 2 ** (tries.busy + 1)) * 1000, opt.signal);
      return g(url, opt, { ...tries, busy: tries.busy + 1 });
    }
    if (!r.ok) {
      let text = '', msg = '';
      try { text = await r.text(); } catch (e) { /* no body */ }
      try { const j = JSON.parse(text); msg = j.error && (j.error.message || j.error.code); } catch (e) { msg = text.slice(0, 200); }
      const err = new Error(`SharePoint answered ${r.status}${msg ? ': ' + msg : ''}`);
      err.status = r.status; err.detail = msg || '';
      throw err;
    }
    return r;
  }
  // A plain fetch (download links carry their own key) that waits out 429/503 like g(), up to 3 times.
  async function fetchRetry(url, opt) {
    for (let busy = 0; ; busy++) {
      const r = await fetch(url, opt);
      if ((r.status !== 429 && r.status !== 503) || busy >= 3) return r;
      const ra = parseInt((r.headers && r.headers.get('Retry-After')) || '', 10);
      await sleep(Math.min(60, ra > 0 ? ra : 2 ** (busy + 1)) * 1000, opt && opt.signal);
    }
  }
  // The cycle's signal plus a time limit, for steps that must not hang the page (where the browser supports it).
  const timed = (signal, ms) => (typeof AbortSignal.any === 'function' && typeof AbortSignal.timeout === 'function'
    ? AbortSignal.any([signal, AbortSignal.timeout(ms)].filter(Boolean)) : signal);
  const gj = async (url, opt) => { const r = await g(url, opt); return r.status === 204 || r.status === 202 ? null : r.json(); };
  async function gAll(url, signal, onPage) {
    const out = [];
    while (url) {
      const j = await gj(url, { signal });
      out.push(...((j && j.value) || []));
      url = j && j['@odata.nextLink'];
      if (onPage) onPage(out.length);
    }
    return out;
  }
  // File contents only through the pre-signed download link: the /content redirect is blocked by CORS.
  async function download(meta, signal) {
    const url = meta && meta['@microsoft.graph.downloadUrl'];
    if (!url) throw problem("can't read", `SharePoint gave no download link for ${(meta && meta.name) || 'the file'}.`);
    const r = await fetchRetry(url, { signal, cache: 'no-store' });
    if (!r.ok) { const e = new Error(`The download failed (${r.status}).`); e.status = r.status; throw e; }
    return r.arrayBuffer();
  }

  // ---------- the Production site (its lists) ----------
  let siteId = null;
  async function ensureSite(signal) {
    if (siteId) return siteId;
    try { siteId = (await gj(`/sites/${SITE}?$select=id`, { signal })).id; }
    catch (e) {
      if (e.status === 404) throw problem('not found', "The Production site couldn't be found on SharePoint (moved or renamed).");
      throw e;
    }
    return siteId;
  }

  // ---------- where each file and folder is: config.js ----------
  // A sharing link, a SharePoint address, or a /sites/<site>/<library>/<path>.
  function readLocation(text) {
    const t = String(text || '').trim();
    if (!t) return null;
    if (/^\/sites\//i.test(t)) return { kind: 'path', host: HOST, path: t };
    if (/^https?:\/\//i.test(t)) {
      let u;
      try { u = new URL(t); } catch (e) { return { kind: 'bad' }; }
      const id = u.searchParams.get('id');
      if (id && /^\/sites\//i.test(id)) return { kind: 'path', host: u.host, path: id };
      if (/^\/:[a-z]:\//i.test(u.pathname)) return { kind: 'share', url: t };
      const p = safeDecode(u.pathname);
      if (/^\/sites\//i.test(p) && !/\.aspx$/i.test(p)) return { kind: 'path', host: u.host, path: p };
      return { kind: 'share', url: t };
    }
    return { kind: 'bad' };
  }
  const shareId = u => 'u!' + btoa(unescape(encodeURIComponent(u))).replace(/=+$/, '').replace(/\//g, '_').replace(/\+/g, '-');
  const SEL = '?$select=id,parentReference';
  // A 404 for the file or folder itself (not its site or library) carries .missing: for ATS that means no report yet.
  const missingItem = e => { if (e && e.status === 404) e.missing = true; return e; };
  async function shareItem(loc, signal) {
    let it;
    // No "redeem" header: the link only resolves for people who already have access.
    try { it = await gj(`/shares/${shareId(loc.url)}/driveItem${SEL}`, { signal }); }
    catch (e) {
      // Microsoft lists only write permissions for opening a sharing link, so a 403 may be the read-only sign-in itself.
      if (e.status === 403) throw problem('link refused', "SharePoint refused the sharing link in config.js (403). Either your account can't open the file, or Shalam Tool's read-only sign-in isn't allowed to open sharing links; the file's /sites/<site>/<library>/<path> location in config.js doesn't need one.", { status: 403 });
      throw missingItem(e);
    }
    const driveId = it.parentReference && it.parentReference.driveId;
    if (!driveId) throw new Error("SharePoint didn't say which library this link is in.");
    return { driveId, item: await gj(`/drives/${driveId}/items/${it.id}`, { signal }) };
  }
  const pathParts = loc => loc.path.split('/').filter(Boolean);
  // The library a /sites/<site>/<library>/<path> is in: the site, then the library matched by the end of its webUrl.
  async function pathDrive(loc, src, signal) {
    const parts = pathParts(loc), site = parts[1], lib = parts[2];
    let sid;
    try { sid = (await gj(`/sites/${loc.host}:/sites/${encodeURIComponent(site)}?$select=id`, { signal })).id; }
    catch (e) {
      if (e.status === 404) throw problem('not found', `There's no site /sites/${site} on SharePoint, so the location in config.js for ${src.name} is wrong (or the site was renamed).`);
      throw e;
    }
    const drives = await gAll(`/sites/${sid}/drives?$select=id,name,webUrl`, signal);
    const d = drives.find(x => tail(x.webUrl).toLowerCase() === lib.toLowerCase()) || drives.find(x => String(x.name || '').toLowerCase() === lib.toLowerCase());
    if (!d) throw problem('not found', `No library called "${lib}" on /sites/${site}. Its libraries: ${drives.map(x => tail(x.webUrl)).join(', ') || 'none'}.`);
    return d.id;
  }
  async function pathItem(driveId, loc, signal) {
    const rest = pathParts(loc).slice(3).join('/');
    try { return await gj(rest ? `/drives/${driveId}/root:/${encPath(rest)}` : `/drives/${driveId}/root`, { signal }); }
    catch (e) { throw missingItem(e); }
  }
  async function pathAt(loc, src, signal) {
    const driveId = await pathDrive(loc, src, signal);
    return { driveId, item: await pathItem(driveId, loc, signal) };
  }

  // A source's file or folder, as Graph gives it. What a location resolves to is kept in this browser, keyed by the
  // exact location text, so a change in config.js is resolved afresh. A sharing link follows its file wherever it
  // moves, so it is reopened by its saved id; a /sites/… path is looked up by its path every time (only its library is
  // saved), because a file renamed away and replaced by a new one under the old name keeps its id.
  // Answers null when the location is blank.
  async function locate(k, s, signal) {
    const text = locText(k), src = BY_K[k];
    if (!text) return null;
    const key = 'loc:' + text;
    let ref = null;
    try { ref = JSON.parse(store.get(key) || 'null'); } catch (e) { ref = null; }
    if (ref && ref.name) { s.file = ref.name; s.webUrl = ref.webUrl || ''; }   // kept if it can't be opened now
    else if (/^https?:/i.test(text)) s.webUrl = text;
    const loc = readLocation(text);
    if (!loc || loc.kind === 'bad') throw problem("can't read the location", `Its location in config.js ("${text}") isn't a SharePoint link or a /sites/… path.`);
    if (loc.kind === 'path' && pathParts(loc).length < 3) throw problem("can't read the location", 'A path needs the site, the library and the folder, like /sites/Production/Shared Documents/Folder.');
    let at = null;
    try {
      if (ref && ref.driveId && (loc.kind === 'path' || ref.itemId)) {
        try { at = { driveId: ref.driveId, item: loc.kind === 'path' ? await pathItem(ref.driveId, loc, signal) : await gj(`/drives/${ref.driveId}/items/${ref.itemId}`, { signal }) }; }
        catch (e) { if (e.status !== 404) throw e; store.del(key); at = null; }   // gone from there: resolve the text again
      }
      if (!at) at = loc.kind === 'share' ? await shareItem(loc, signal) : await pathAt(loc, src, signal);
    } catch (e) {
      if (e.label) throw e;
      if (e.status === 400) throw problem("can't read the location", `SharePoint didn't accept the link in config.js for ${src.name}. Copy it again (Share, then Copy link), or use a /sites/… path.`);
      if (e.status === 404) throw problem('not found', "It isn't where config.js says it is (moved, renamed or deleted).", { status: 404, missing: !!e.missing });
      throw e;
    }
    const item = at.item;
    if (src.kind === 'file' && !item.file) throw problem("can't read the location", `config.js points ${src.name} at a folder (${item.name}); it needs the file itself.`);
    if (src.kind === 'folder' && !item.folder) throw problem("can't read the location", `config.js points ${src.name} at a file (${item.name}); it needs the folder the reports land in.`);
    store.set(key, JSON.stringify({ driveId: at.driveId, itemId: loc.kind === 'share' ? item.id : '', name: item.name || '', webUrl: item.webUrl || '' }));
    return at;
  }
  const notSet = s => off(s, 'not set', "Its location in config.js is blank, so it isn't read.", 'location blank in config.js');

  // ---------- source status ----------
  // state: ok (green) · warn (amber) · err (red) · off (grey: not set, waiting, or no report yet) · load
  // short: the chip's second line when there is no saved date to show (the full message is in its tooltip and card).
  // stale: the time it was last read, when a later refresh couldn't read it.
  function fresh(k) {
    return { k, state: 'off', label: '', message: '', short: '', headers: '', file: '', webUrl: '', savedAt: '', savedBy: '',
      read: null, counted: null, rows: [], issues: [], notes: [], mailProblem: '', mail: null, blocked: false, status: 0, stale: null };
  }
  const ok = s => { s.state = 'ok'; s.label = 'ok'; };
  const warn = (s, label, message) => { s.state = 'warn'; s.label = label; s.message = message; };
  const off = (s, label, message, short) => { s.state = 'off'; s.label = label; s.message = message || ''; s.short = short || ''; };
  function problem(label, message, extra) { const e = new Error(message); e.label = label; return Object.assign(e, extra || {}); }
  function describe(e) {
    if (e && e.label) return { label: e.label, message: e.message, headers: e.headers || '' };
    const s = e && e.status;
    if (s === 401 || s === 403) return { label: 'no access', message: "Your account can't open it on SharePoint." };
    if (s === 404) return { label: 'not found', message: "It wasn't found on SharePoint (moved, renamed or deleted)." };
    if (/failed to fetch|networkerror|load failed/i.test(String((e && e.message) || ''))) return { label: "can't reach", message: "Couldn't reach SharePoint. Check the connection; the page tries again in a few minutes." };
    return { label: "can't read", message: String((e && e.message) || e || 'Unknown error.') };
  }
  function labelFor(text) {
    const t = String(text || '');
    if (/column/i.test(t)) return 'missing column';
    if (/\bno "[^"]*" (tab|sheet)\b|\bno sheets\b/i.test(t)) return 'missing tab';
    if (/\bempty\b|\bno (lines|shipments|container rows|PO lines|items|records)\b|\bno lines were read\b|\bno container rows were read\b/i.test(t)) return 'empty';
    return "can't read the file";
  }
  function fail(s, e) {
    const d = describe(e);
    s.state = 'err'; s.label = d.label; s.message = d.message; s.headers = d.headers || ''; s.status = (e && e.status) || 0;
    const q = /"([^"]+)"/.exec(d.message || '');
    s.mailProblem = (d.label === 'missing column' || d.label === 'missing tab') && q ? `${d.label} "${q[1]}"` : d.label;
  }
  // countTracker / countMaster results: { missing, seen, rowsRead, counts, labels }.
  function missingCols(x, sheet) {
    const names = x.missing.map(n => `"${n}"`).join(', ');
    return problem('missing column', `Missing column${x.missing.length > 1 ? 's' : ''} ${names} on the ${sheet} sheet.`, { headers: (x.seen || []).join(' | ') });
  }
  const cap = t => String(t).replace(/^./, c => c.toUpperCase());
  // Parser errors start with the source's name, which the chip, card and email already show.
  const unprefix = (name, msg) => cap(String(msg || '').replace(new RegExp('^\\s*' + name.replace(/[.*+?^${}()|[\]\\]/g, '\\$&') + ':\\s*', 'i'), ''));
  const resultRows = x => [['Rows read', x.rowsRead],
    ...Object.entries(x.counts || {}).map(([k, v]) => [cap((x.labels && x.labels[k]) || human(k)), v])];
  // The PO Book's copied code throws `<sheet>: missing column "<name>"`.
  function coreProblem(e, rows, sheet) {
    const m = /missing column "([^"]+)"/i.exec((e && e.message) || '');
    if (m) return problem('missing column', `Missing column "${m[1]}" on the ${sheet} sheet.`, { headers: headerList(rows) });
    return problem("can't read the file", `Reading the ${sheet} sheet failed: ${(e && e.message) || e}`);
  }
  function normIssue(source, i) {
    i = i || {};
    return { source: i.source || source, code: String(i.code || ''), sheet: String(i.sheet || ''), row: i.row == null ? '' : String(i.row),
      po: String(i.po || i.poRaw || ''), style: String(i.sty || i.style || i.styleRaw || ''), message: String(i.message || i.issue || '') };
  }
  const itemOf = c => c.item || (c.file && c.file.name ? c.file : c);

  // ---------- state ----------
  let st = {};
  const cache = {};
  let loading = null, timer = null, due = false, again = false, lastLoaded = null, fatal = '', tab = 'inv';
  let fatalAuth = false;   // the fatal message asks to sign in again
  let cycleDone = false;   // a load cycle has ended (finished, cancelled or failed)
  let scopeFatal = '';     // the sign-in isn't read-only, or can't read SharePoint: nothing loads until it's fixed
  let copyCheck = null;   // null: not known · true: same rules as the PO Book · false: the PO Book's rules changed
  let mailScopeMissing = false;
  const mailMemo = new Map();   // source|day -> what happened to its email, in this page session
  // Parsed results, kept for step 2.
  const data = { po: null, ct: null, forwarder: null, drayage: null, upc: null };

  // Once the page has stopped over its sign-in permissions, a reader finishing late shows nothing.
  function showLoading(k) { if (scopeFatal) return; st[k] = { ...(st[k] || fresh(k)), state: 'load' }; renderSources(); }
  function commit(k, s) { if (scopeFatal) return; st[k] = s; renderSources(); }
  // Errors that end the whole cycle rather than one source: Cancel, the sign-in, or a token that isn't read-only.
  const cycleError = e => isAbort(e) || !!(e && (e.auth || e.stop));
  async function attempt(s, fn) {
    try { return await fn(); }
    catch (e) { if (cycleError(e)) throw e; console.error(e); fail(s, e); return null; }
  }

  // noCancel: a step Cancel can't stop cleanly (the error email), so the Cancel buttons are hidden while it runs.
  function setBusy(msg, noCancel) {
    if (msg != null && !loading) return;   // a step finishing after Cancel doesn't bring the spinner back
    $('btnCancelBusy').hidden = $('btnCancelLoad').hidden = msg != null && !!noCancel;
    if (msg == null) {
      $('busy').hidden = true; $('invLoad').hidden = true; $('app').setAttribute('aria-busy', 'false');
      return;
    }
    $('busyMsg').textContent = msg; $('invLoadMsg').textContent = msg;
    $('busy').hidden = false; $('invLoad').hidden = !!lastLoaded;
    $('app').setAttribute('aria-busy', 'true');
  }

  // ---------- reading a file source ----------
  async function fileMeta(k, s, signal) {
    const at = await locate(k, s, signal);
    if (!at) { notSet(s); return null; }
    const meta = at.item;
    s.file = meta.name || s.file; s.webUrl = meta.webUrl || s.webUrl;
    s.savedAt = meta.lastModifiedDateTime || ''; s.savedBy = who(meta.lastModifiedBy);
    return meta;
  }
  const unchanged = (c, meta) => !!c && c.id === meta.id && c.lm === meta.lastModifiedDateTime;
  async function fetchBytes(meta, label, signal) {
    setBusy(`Downloading ${label}…`);
    const buf = await download(meta, signal);
    if (!buf || !buf.byteLength) throw problem('empty', `${meta.name} is empty (0 bytes).`);
    setBusy(`Reading ${label}…`); await yieldUi();
    return buf;
  }
  function readBook(buf, name, extra) {
    try { return XLSX.read(new Uint8Array(buf), { ...XOPT, ...(extra || {}) }); }
    catch (e) { throw problem("can't read the file", `Couldn't open ${name} as an Excel workbook (${e.message}).`); }
  }
  // One sheet's rows, counted from A1. Only that sheet is parsed; a name that differs only in case is still found.
  function readSheet(buf, sheet, fileName) {
    const IC = window.InvCommon;
    let wb = readBook(buf, fileName, { sheets: [sheet] });
    let hit = IC.findSheet(wb, sheet);
    if (!hit || !hit.ws) {
      const names = readBook(buf, fileName, { bookSheets: true }).SheetNames || [];
      const real = names.find(n => IC.S(n).toLowerCase() === sheet.toLowerCase());
      if (!real) throw problem('missing tab', `No sheet named "${sheet}" in ${fileName}. Sheets in the file: ${names.join(', ') || 'none'}.`);
      wb = readBook(buf, fileName, { sheets: [real] });
      hit = IC.findSheet(wb, sheet);
      if (!hit || !hit.ws) throw problem("can't read the file", `The "${sheet}" sheet in ${fileName} couldn't be read.`);
    }
    return IC.sheetRows(XLSX, hit.ws);
  }
  async function sheetSource(k, s, force, signal) {
    const meta = await fileMeta(k, s, signal);
    if (!meta) return null;
    if (!force && unchanged(cache[k], meta)) return cache[k].rows;
    const src = BY_K[k];
    const rows = readSheet(await fetchBytes(meta, src.name, signal), src.sheet, meta.name);
    if (!dataRowCount(rows)) throw problem('empty', `The "${src.sheet}" sheet in ${meta.name} has no rows under its header.`);
    cache[k] = { id: meta.id, lm: meta.lastModifiedDateTime, rows };
    return rows;
  }
  async function bookSource(k, s, force, signal) {
    const meta = await fileMeta(k, s, signal);
    if (!meta) return null;
    if (!force && unchanged(cache[k], meta)) return cache[k].wb;
    const wb = readBook(await fetchBytes(meta, BY_K[k].name, signal), meta.name);
    cache[k] = { id: meta.id, lm: meta.lastModifiedDateTime, wb };
    return wb;
  }
  async function listSource(k, s, fields, useCache, force, signal) {
    const name = BY_K[k].list;
    s.file = name + ' (list)';
    await ensureSite(signal);
    let list;
    try { list = await gj(`/sites/${siteId}/lists/${encodeURIComponent(name)}`, { signal }); }
    catch (e) {
      if (e.status === 404) throw problem('not found', `The ${name} list wasn't found on the Production site (renamed or deleted).`);
      throw e;
    }
    s.file = (list.displayName || name) + ' (list)'; s.webUrl = list.webUrl || '';
    s.savedAt = list.lastModifiedDateTime || ''; s.savedBy = who(list.lastModifiedBy);
    const c = cache[k];
    if (useCache && !force && c && c.id === list.id && c.lm === list.lastModifiedDateTime) return c.items;
    setBusy(`Reading ${name}…`);
    const items = await gAll(`/sites/${siteId}/lists/${list.id}/items?$expand=fields($select=${fields})&$top=999`, signal,
      n => setBusy(`Reading ${name}: ${fmtN(n)} items…`));
    cache[k] = { id: list.id, lm: list.lastModifiedDateTime, items };
    return items;
  }

  // ---------- the PO Book group: tracker, Master Sheet Live and PO_Book_Lines, read as the PO Book reads them ----------
  async function loadPoGroup(force, signal) {
    const PB = window.PoBookCore;
    const S = { tracker: fresh('tracker'), msl: fresh('msl'), pobook: fresh('pobook') };
    for (const k of Object.keys(S)) showLoading(k);
    const [trows, mrows, items] = await Promise.all([
      attempt(S.tracker, () => sheetSource('tracker', S.tracker, force, signal)),
      attempt(S.msl, () => sheetSource('msl', S.msl, force, signal)),
      attempt(S.pobook, () => listSource('pobook', S.pobook, POB_FIELDS, false, force, signal)),
    ]);
    setBusy('Reading the PO Book group…'); await yieldUi();

    // countTracker / countMaster give every row read a result and name every missing column at once; the copied
    // parseTracker / parseMaster then read the lines exactly as the PO Book does.
    let trk = null;
    if (trows) {
      try {
        const tc = PB.countTracker(trows);
        if (tc.missing && tc.missing.length) fail(S.tracker, missingCols(tc, 'PO Status Tracker'));
        else {
          try { trk = PB.parseTracker(trows); } catch (e) { fail(S.tracker, coreProblem(e, trows, 'PO Status Tracker')); }
          S.tracker.read = tc.rowsRead; S.tracker.counted = tc.counts.tracked;
          S.tracker.rows = resultRows(tc);
          if (trk && Number.isFinite(trk.minPO)) S.tracker.rows.push(['Lowest PO on the tracker', String(trk.minPO)]);
          S.tracker.issues = (tc.issues || []).map(i => normIssue('PO Status Tracker', i));
          if (trk && !tc.ok) fail(S.tracker, problem('empty', unprefix('PO Status Tracker', tc.error) || 'No PO lines on the PO Status Tracker sheet.'));
          else if (trk) ok(S.tracker);
        }
      } catch (e) {
        console.error(e); trk = null;
        fail(S.tracker, problem("can't read the file", 'Reading the PO Status Tracker sheet failed: ' + e.message));
      }
    }

    let ov = null, cfg = null;
    if (items) {
      try {
        // Every item read gets a result: a line, the settings item, an older copy of a line (the newest wins, as in
        // the PO Book), or no LineKey (listed). No item with a LineKey at all means the column is gone.
        const keyOf = it => String((it.fields && it.fields.LineKey) || '');
        const noKey = items.filter(it => !keyOf(it)), keys = new Set(items.map(keyOf).filter(Boolean));
        if (items.length > 1 && !keys.size) {
          const seen = Object.keys((items[0] && items[0].fields) || {}).filter(f => !/^@odata/.test(f));
          throw problem('missing column', 'No item in the PO_Book_Lines list has a "LineKey", so no PO Book mark can be read.', { headers: seen.join(' | ') });
        }
        ov = PB.overlayFromItems(items); cfg = PB.cfgFrom(items);
        const older = items.length - noKey.length - keys.size, settingsN = keys.has('__settings') ? 1 : 0;
        S.pobook.read = items.length; S.pobook.counted = ov.size;
        S.pobook.rows = [['List items read', items.length], ['Lines marked in the PO Book', ov.size], ['Settings item', settingsN],
          ['Older copies of a line (the newest is used)', older], ['Items with no LineKey (listed on Problems)', noKey.length]];
        S.pobook.issues = noKey.map(it => normIssue('PO Book lines', { code: 'no-linekey', row: 'Item ' + it.id,
          message: `List item ${it.id}${it.fields && it.fields.Title ? ` ("${it.fields.Title}")` : ''} has no LineKey, so the PO Book can't tie it to a line and its mark is ignored.` }));
        if (!items.length) warn(S.pobook, 'empty', 'The PO_Book_Lines list has no items: nothing has been marked shipped or split in the PO Book yet.');
        else ok(S.pobook);
      } catch (e) { ov = null; cfg = null; fail(S.pobook, e.label ? e : problem("can't read the list", 'Reading the PO Book lines failed: ' + e.message)); }
    }

    if (mrows && (!trk || S.tracker.state === 'err')) {
      const why = S.tracker.label === 'not set' ? 'whose location is blank in config.js' : "which didn't load";
      off(S.msl, 'waiting', `Not read, because it's read together with the PO Status Tracker, ${why}.`, 'read with the PO Status Tracker');
      S.msl.blocked = true;
    } else if (mrows) {
      try {
        readMaster(PB, S, mrows, trk, ov, cfg);
      } catch (e) {
        console.error(e);
        fail(S.msl, problem("can't read the file", 'Reading the Master Sheet Live sheet failed: ' + e.message));
      }
    }
    for (const k of Object.keys(S)) commit(k, S[k]);
  }

  function readMaster(PB, S, mrows, trk, ov, cfg) {
    const cm = PB.countMaster(mrows, trk);
    if (cm.missing && cm.missing.length) { fail(S.msl, missingCols(cm, 'Master Sheet Live')); return; }
    let src;
    try { src = PB.parseMaster(mrows, trk); } catch (e) { fail(S.msl, coreProblem(e, mrows, 'Master Sheet Live')); return; }
    const lines = ov && cfg ? PB.applyOverlay(src, ov, cfg) : src.lines;
    const isOpen = l => (l.status || (l.src === 'round' ? 'shipped' : 'open')) === 'open';
    const open = lines.filter(isOpen).length;
    S.msl.read = cm.rowsRead; S.msl.counted = cm.counts.line;
    S.msl.rows = [...resultRows(cm), { h: 'Lines as the PO Book shows them' }, ['Lines', lines.length], ['Open', open], ['Shipped', lines.length - open]];
    if (!ov) S.msl.notes.push("Open and shipped are counted without the PO Book's marks, because the PO Book lines didn't load.");
    S.msl.issues = (cm.issues || []).map(i => normIssue('Master Sheet Live', i));
    // parseMaster's 'nostyle' issues repeat countMaster's (which carry row numbers).
    for (const i of src.issues || []) {
      if (i.code === 'nostyle') continue;
      if (i.code === 'trkonly') S.tracker.issues.push(normIssue('PO Status Tracker', i));
      else S.msl.issues.push(normIssue('Master Sheet Live', i));
    }
    if (!cm.ok || !src.lines.length) fail(S.msl, problem('empty', unprefix('Master Sheet Live', cm.error) || 'No PO lines found on Master Sheet Live.'));
    else ok(S.msl);
    // Step 2 joins these lines to Container Tracking on C.normPo(l.po) and C.normStyle(l.sty), never on l.key or
    // l.base: the PO Book's keys fold inner spaces and -F, which this tool keeps apart.
    data.po = { trk, src, lines, ov, cfg, counts: cm.counts };
  }

  // ---------- Container Tracking ----------
  async function loadCT(today, force, signal) {
    const s = fresh('ct');
    showLoading('ct');
    const wb = await attempt(s, () => bookSource('ct', s, force, signal));
    if (wb) {
      setBusy('Reading Container Tracking…'); await yieldUi();
      let r = null;
      try { r = window.InvCT.parseContainerTracking(XLSX, wb, today); }
      catch (e) { console.error(e); fail(s, problem("can't read the file", 'Reading Container Tracking failed: ' + e.message)); }
      if (r) {
        const c = r.counts || {};
        s.issues = (r.issues || []).map(i => normIssue('Container Tracking', i));
        s.rows = countRows(c);
        // Rows read on all four tabs against rows used: style lines on Booking and Loaded, container rows on Open and
        // Complete. The rest are blank, header and title rows, or listed on Problems.
        s.read = sumOf(['Booking', 'Loaded', 'Open', 'Complete'].map(t => c[t] && c[t].rowsRead));
        s.counted = sumOf([c.Booking && c.Booking.counted, c.Loaded && c.Loaded.counted, c.Open && c.Open.roundRows, c.Complete && c.Complete.roundRows]);
        s.rows.unshift(['Rows read (all four tabs)', s.read], ['Rows used (style lines + container rows)', s.counted]);
        if (!r.ok) fail(s, problem(labelFor(r.error), unprefix('Container Tracking', r.error) || "Container Tracking couldn't be read."));
        else if (!(r.blocks || []).length) fail(s, problem('empty', 'No lines on the Booking or Loaded tabs.'));
        else { ok(s); data.ct = r; }
      }
    }
    commit('ct', s);
  }

  // ---------- forwarder and drayage: the newest usable report in a folder ----------
  async function parseReport(k, meta, today, force, signal) {
    const ck = 'file:' + k;
    let c = cache[ck];
    if (force || !unchanged(c, meta)) {
      const buf = await fetchBytes(meta, meta.name, signal);
      c = { id: meta.id, lm: meta.lastModifiedDateTime,
        data: k === 'drayage' ? new TextDecoder('utf-8').decode(buf) : readBook(buf, meta.name) };
      cache[ck] = c;
    }
    return k === 'drayage' ? window.InvShipping.parseDrayage(c.data, today) : window.InvShipping.parseForwarder(XLSX, c.data, today);
  }
  const ageText = c => `${plural(c.ageDays, 'day')} (${plural(c.businessDaysOld || 0, 'business day')})`;

  async function loadReport(k, today, force, signal) {
    const src = BY_K[k], s = fresh(k);
    showLoading(k);
    await attempt(s, async () => {
      const at = await locate(k, s, signal);
      if (!at) return notSet(s);
      const drive = at.driveId, folder = at.item;
      s.file = folder.name || s.file; s.webUrl = folder.webUrl || s.webUrl;
      setBusy(`Looking in the ${src.name} folder…`);
      const kids = await gAll(`/drives/${drive}/items/${folder.id}/children?$select=id,name,lastModifiedDateTime,lastModifiedBy,file,folder,webUrl,size&$top=999`, signal);
      const SH = window.InvShipping;
      const maxDays = SH.MAX_REPORT_DAYS ?? MAX_REPORT_DAYS, tries = SH.TRY_FILES ?? 3;
      const cands = SH.pickNewest(kids, src.ext, today) || [];
      if (!cands.length) throw problem('empty', `No ${src.ext} file in the ${folder.name} folder.`);
      const failed = [];
      let used = null, stale = null, lastBad = null;
      for (const c of cands.slice(0, tries)) {
        const f = itemOf(c);
        if (c.tooOld) { stale = { f, c }; break; }   // the rest are older still
        try {
          const meta = await gj(`/drives/${drive}/items/${f.id}`, { signal });
          const r = await parseReport(k, meta, today, force, signal);
          if (r && r.ok) { used = { meta, c, r }; break; }
          lastBad = r;
          failed.push(`${f.name}: ${(r && r.error) || 'nothing usable in it'}`);
        } catch (e) {
          if (cycleError(e)) throw e;
          failed.push(`${f.name}: ${describe(e).message}`);
        }
      }
      if (used) {
        const { meta, c, r } = used;
        s.file = meta.name; s.webUrl = meta.webUrl || s.webUrl; s.savedAt = meta.lastModifiedDateTime || ''; s.savedBy = who(meta.lastModifiedBy);
        s.read = firstNum(r.counts, ['rowsRead']);
        s.counted = firstNum(r.counts, ['records']);
        s.rows = [...countRows(r.counts, SRC_LABELS[k]), ['Age', ageText(c)]];
        s.issues = (r.issues || []).map(i => normIssue(src.name, i));
        for (const x of failed) s.issues.push(normIssue(src.name, { code: 'newer-file-not-read', message: x }));
        if (failed.length) warn(s, 'older file used', `Used ${meta.name}: the newer file couldn't be read (${failed.join('; ')}).`);
        else if (c.amber || (c.businessDaysOld || 0) >= AMBER_BDAYS) warn(s, `${plural(c.businessDaysOld, 'business day')} old`, `${meta.name} was saved ${plural(c.businessDaysOld, 'business day')} ago, so its dates may be behind.`);
        else ok(s);
        data[k] = r;
        return;
      }
      if (failed.length) {
        const seen = lastBad && (lastBad.seen || lastBad.headers);
        throw problem(labelFor((lastBad && lastBad.error) || failed[0]), failed.join('; '), { headers: Array.isArray(seen) ? seen.join(' | ') : '' });
      }
      const { f, c } = stale;
      s.file = f.name; s.webUrl = f.webUrl || s.webUrl; s.savedAt = f.lastModifiedDateTime || ''; s.savedBy = who(f.lastModifiedBy);
      if (c.ageDays == null) {
        throw problem("can't read", `The newest file, ${f.name}, has no saved date SharePoint could give, so it isn't used.`);
      }
      s.rows = [['Newest file age', ageText(c)]];
      warn(s, 'not used', `The newest file, ${f.name}, is ${plural(c.ageDays, 'day')} old (more than ${maxDays}), so it isn't used: stages come from Container Tracking dates only.`);
      s.mailProblem = `not used, the newest file is ${plural(c.ageDays, 'day')} old`;
    });
    commit(k, s);
  }

  // ---------- ATS: step 1 only looks for the newest report (it is read in step 3; nothing is ever uploaded) ----------
  // Where to save the report, in plain words: "Inventory/ATS on the Production site" for /sites/Production/<library>/Inventory/ATS.
  function atsPlace() {
    const parts = locText('ats').split('/').filter(Boolean);
    return /^\/sites\//i.test(locText('ats')) && parts.length > 3 ? `${parts.slice(3).join('/')} on the ${parts[1]} site` : 'the ATS folder set in config.js';
  }
  async function loadAts(today, signal) {
    const s = fresh('ats');
    showLoading('ats');
    const noReport = () => off(s, 'no report yet', `No ATS report yet. Save the weekly ATS report into ${atsPlace()}.`, `save it into ${atsPlace()}`);
    await attempt(s, async () => {
      let at;
      try { at = await locate('ats', s, signal); }
      catch (e) { if (e.missing) return noReport(); throw e; }   // only the folder itself missing; a wrong site or library stays red
      if (!at) return notSet(s);
      const folder = at.item;
      s.file = folder.name || s.file; s.webUrl = folder.webUrl || s.webUrl;
      const kids = (await gAll(`/drives/${at.driveId}/items/${folder.id}/children?$select=id,name,lastModifiedDateTime,lastModifiedBy,file,webUrl&$top=999`, signal))
        .filter(x => x.file && !/^~\$/.test(x.name || ''));
      s.rows = [['Files in the folder', kids.length]];
      if (!kids.length) return noReport();
      const f = kids.slice().sort((a, b) => String(b.lastModifiedDateTime).localeCompare(String(a.lastModifiedDateTime)))[0];
      s.file = f.name; s.webUrl = f.webUrl || s.webUrl; s.savedAt = f.lastModifiedDateTime || ''; s.savedBy = who(f.lastModifiedBy);
      const ms = Date.parse(f.lastModifiedDateTime);
      const day = Number.isFinite(ms) ? window.InvCommon.todayNY(ms) : null;
      const age = day == null ? null : today - day;
      s.rows.push(['Newest file', f.name], ['Saved', age == null ? '' : `${plural(age, 'day')} ago`]);
      if (age != null && age > ATS_MAIL_DAYS) {
        warn(s, `${plural(age, 'day')} old`, `Found ${f.name}, saved ${plural(age, 'day')} ago: the weekly ATS report is late. It's read in step 3.`);
        s.mailProblem = `the newest report is ${plural(age, 'day')} old`;
      } else if (age != null && age > ATS_AMBER_DAYS) {
        warn(s, `${plural(age, 'day')} old`, `Found ${f.name}, saved ${plural(age, 'day')} ago. It's read in step 3.`);
      } else off(s, 'read in step 3', `Found ${f.name}. It's read in step 3.`);
    });
    commit('ats', s);
  }

  // ---------- UPC_MASTER ----------
  async function loadUpc(force, signal) {
    const s = fresh('upc');
    showLoading('upc');
    const items = await attempt(s, () => listSource('upc', s, UPC_FIELDS, true, force, signal));
    if (items) {
      try {
        const r = window.InvUpc.fromListItems(items);
        const c = r.counts || {};
        // A list has no row numbers: its issues carry the item's ID.
        s.issues = (r.issues || []).map(i => normIssue('UPC_MASTER', { ...i, row: i.row == null || i.row === '' ? '' : 'Item ' + i.row }));
        s.rows = countRows(c);
        s.read = firstNum(c, ['rows']) ?? items.length; s.counted = firstNum(c, ['used']);
        if (!items.length || r.ok === false) fail(s, problem('empty', unprefix('UPC_MASTER', r.error) || 'The UPC_MASTER list has no items.'));
        else { ok(s); data.upc = r; }
      } catch (e) { console.error(e); fail(s, problem("can't read the list", 'Reading UPC_MASTER failed: ' + e.message)); }
    }
    commit('upc', s);
  }

  // ---------- is the copied PO Book code still the PO Book's? ----------
  async function sha256Hex(text) {
    const b = await crypto.subtle.digest('SHA-256', new TextEncoder().encode(text));
    return Array.from(new Uint8Array(b), x => x.toString(16).padStart(2, '0')).join('');
  }
  async function checkPoBookCopy(signal) {
    try {
      const r = await fetch(location.origin + '/po-book/index.html', { cache: 'no-store', signal: timed(signal, 15000) });
      if (!r.ok) return;
      const text = window.PoBookCore.coreText(await r.text());   // line ends already read as LF
      if (text == null) { copyCheck = false; return; }
      if (!window.crypto || !crypto.subtle) return;
      copyCheck = (await sha256Hex(text)) === window.PoBookCore.CORE_SHA256;
    } catch (e) {
      if (signal && signal.aborted) throw abortErr();
      /* can't fetch it, or it took too long: say nothing */
    }
  }
  function applyCopyCheck() {
    const s = st.pobook;
    if (!s || copyCheck !== false || s.state === 'load') return;
    const note = "PO Book rules changed since this copy: the PO Book has been updated, and this page still reads Master Sheet Live with the older rules until it's updated too.";
    if (!s.notes.includes(note)) s.notes.push(note);
    if (s.state === 'ok') warn(s, 'PO Book rules changed', note);
  }

  // ---------- the error email: once per source per day in each browser ----------
  function recipientName(addr) {
    const named = String(CFG.errorEmailName || '').trim();
    if (named) return named;
    const local = addr.split('@')[0];
    return /^[a-z]{2,}$/i.test(local) ? local[0].toUpperCase() + local.slice(1).toLowerCase() : addr;
  }
  // What fixes it, and what it holds up, in the email.
  const NEXT = {
    'missing column': 'Put the column back with the name shown, or have the page updated to the new name.',
    'missing tab': 'Put the tab back with the name shown, or have the page updated to the new name.',
    'not found': 'Put it back, or update its location in config.js: it was moved, renamed or deleted.',
    "can't read the location": 'Fix its location in config.js.',
    empty: 'Open the file and check its rows are still there (it may have been cleared or saved over).',
    "can't read the file": 'Open the file and check it opens in Excel; save it again if it is damaged.',
  };
  function nextStep(src, s, user) {
    if (s.label === 'link refused') return `Check ${user} can open the file. If they can, the read-only sign-in can't use sharing links: replace the link in config.js with the file's /sites/<site>/<library>/<path> location.`;
    // The lists' names are in the page itself, not in config.js.
    if (src.kind === 'list' && s.label === 'no access') return `Give ${user} access to the Production site.`;
    if (src.kind === 'list' && s.label === 'not found') return `Check the Production site and its ${src.list} list are still there under those names, or have the page updated to the new names.`;
    if (s.label === 'no access') return `Give ${user} access to it, or check its location in config.js.`;
    if (src.k === 'forwarder' && s.state === 'warn') return "Check the forwarder's report is still arriving in the folder.";
    if (src.k === 'ats' && s.state === 'warn') return 'Check the weekly ATS report is still being saved to the folder.';
    return NEXT[s.label] || 'Open it and check it; the page tries again every 5 minutes while it is open.';
  }
  function effectOf(src) {
    if (CORE_KEYS.includes(src.k)) return "The stage numbers stay off until it's fixed.";
    if (REPORT_KEYS.includes(src.k)) return "Stages come from Container Tracking dates only until it's fixed.";
    if (src.k === 'ats') return 'On hand comes from this report, so it would be out of date.';
    if (src.k === 'upc') return "Styles can't be matched to their UPCs until it's fixed.";
    return '';
  }
  function mailBody(src, s, to) {
    const when = new Date().toLocaleString('en-US', { timeZone: NY, month: 'numeric', day: 'numeric', year: 'numeric', hour: 'numeric', minute: '2-digit' }) + ' New York time';
    const user = (account && (account.name || account.username)) || 'The person who had the page open';
    // Written for Saul as the reader: "your account" on the screen is the account of whoever had the page open.
    const what = s.label === 'no access' ? `${user}'s account can't open ${s.file || 'it'} on SharePoint${s.status ? ` (${s.status})` : ''}.`
      : s.label === 'link refused' ? `SharePoint refused the sharing link in config.js when ${user} opened the page (${s.status || 403}): either that account can't open the file, or Shalam Tool's read-only sign-in isn't allowed to open sharing links.`
        : `${s.label}. ${s.message}`;
    const lines = [
      `Shalam Tool has a problem with ${src.name}.`,
      '',
      `Source: ${src.name}`,
      `File: ${s.file || 'not known'}`,
      `Last saved: ${s.savedAt ? fmtLong(s.savedAt) + (s.savedBy ? ' by ' + s.savedBy : '') : 'not known'}`,
      `Problem: ${what}`,
    ];
    if (s.headers) lines.push(`Headers seen: ${s.headers}`);
    lines.push(`What fixes it: ${nextStep(src, s, user)}`);
    const eff = effectOf(src);
    if (eff) lines.push(`Until then: ${eff}`);
    lines.push(
      `Page open by: ${(account && account.name) || ''} (${(account && account.username) || ''})`,
      `When: ${when}`,
      `Link to the file: ${s.webUrl || 'none'}`,
      `Shalam Tool: ${REDIRECT}`,
      '',
      'This email is sent once per source per day from each browser that has the page open, from the mailbox of whoever is signed in.');
    return {
      message: {
        subject: `Shalam Tool: ${src.name} — ${s.mailProblem}`,
        body: { contentType: 'Text', content: lines.join('\n') },
        toRecipients: [{ emailAddress: { address: to } }],
      },
      saveToSentItems: true,
    };
  }
  // Why an email didn't go, about the mailbox rather than SharePoint.
  function mailFailText(e) {
    if (mailScopeMissing) return 'this sign-in has no permission to send mail (Mail.Send).';
    const code = e && e.status, d = e && e.detail ? ': ' + e.detail : '';
    if (code === 401 || code === 403) return `this account's mailbox didn't allow sending (${code}${d}).`;
    if (code) return `the mail server answered ${code}${d}.`;
    if (e && (e.name === 'TimeoutError' || isAbort(e))) return 'the mail server took too long to answer.';
    return /failed to fetch|networkerror|load failed/i.test(String((e && e.message) || '')) ? "couldn't reach Microsoft 365." : String((e && e.message) || e);
  }
  // Once per source per day (New York date) in this browser: the mark is checked and set before sending, in browser
  // storage and in memory, so neither a reload nor the 5-minute refresh sends it again. With storage blocked, the
  // in-memory mark alone allows one per page session. Old marks are removed.
  const mailedNow = new Set();
  function claimMail(day, k) {
    const key = `mailed:${day}:${k}`;
    store.prune('mailed:', rest => rest.startsWith(day + ':'));
    if (mailedNow.has(key)) return false;
    mailedNow.add(key);
    if (store.get(key)) return false;
    store.set(key, new Date().toISOString());
    return true;
  }
  // The email step isn't stopped by Cancel (stopping between the mark and the send would mean no email today), so
  // Cancel is hidden while it runs; the send has a time limit instead, so a stalled request can't hold the page.
  async function emailProblems(today) {
    const to = String(CFG.errorEmailTo || '').trim();
    const name = to ? recipientName(to) : '';
    const day = window.InvCommon.isoDay(today);
    for (const src of SOURCES) {
      const s = st[src.k];
      if (!s || !s.mailProblem || (s.state !== 'err' && s.state !== 'warn')) continue;
      if (!to) { s.mail = { state: 'none', text: 'No error email address is set in config.js, so nobody is emailed about this.' }; continue; }
      const memo = src.k + '|' + day;
      if (mailMemo.has(memo)) { s.mail = mailMemo.get(memo); continue; }
      let mail;
      if (!claimMail(day, src.k)) {
        mail = { state: 'already', text: `${name} was already emailed about ${src.name} today.` };
      } else {
        setBusy(`Emailing ${name} about ${src.name}… (finishes in a moment)`, true);
        try {
          await g('/me/sendMail', { method: 'POST', signal: timed(null, 30000), body: JSON.stringify(mailBody(src, s, to)) });
          mail = { state: 'sent', text: `${name} has been emailed about this.` };
          toast(`${name} has been emailed about ${src.name}.`, 6000);
        } catch (e) {
          if (cycleError(e)) { store.del(`mailed:${day}:${src.k}`); mailedNow.delete(`mailed:${day}:${src.k}`); throw e; }
          mail = { state: 'failed', text: `Emailing ${name} failed: ${mailFailText(e)} Tell ${name} about the problem above.` };
          // A later visit can try again; this page session doesn't (the memo below), so a failing mailbox isn't retried every 5 minutes.
          store.del(`mailed:${day}:${src.k}`);
        }
      }
      mailMemo.set(memo, mail); s.mail = mail;
    }
  }

  // ---------- the load cycle ----------
  // After a cancelled or failed cycle: a source read in this cycle keeps its new state; any other keeps what was read
  // before, marked with when that was, or, if it was never read, says so (never "not set"). Returns how many were
  // read in this cycle.
  function keepUnread(before, label, message, short) {
    let read = 0;
    for (const x of SOURCES) {
      const cur = st[x.k], prev = before[x.k];
      if (cur && cur !== prev && cur.state !== 'load') { read++; continue; }
      if (prev && prev.state !== 'load' && lastLoaded) st[x.k] = { ...prev, stale: prev.stale || lastLoaded };
      else { const s = fresh(x.k); off(s, label, message, short); st[x.k] = s; }
    }
    return read;
  }

  // The sign-in isn't read-only, or can't read SharePoint: nothing is shown as read, and nothing more is requested.
  function stopAll() {
    fatal = scopeFatal; fatalAuth = true;
    st = {};
    for (const x of SOURCES) { const s = fresh(x.k); off(s, 'not read', 'Not read: Shalam Tool stopped because of its sign-in permissions (see the red banner).', 'see the red banner'); st[x.k] = s; }
    for (const k of Object.keys(data)) data[k] = null;
    lastLoaded = null;
  }

  async function loadAll(force) {
    if (loading) { if (force) again = true; return; }
    if (!account) return;
    if (scopeFatal) { stopAll(); cycleDone = true; renderAll(); return; }   // only fixing the permission and signing in again helps
    const ctl = new AbortController(), signal = ctl.signal;
    loading = ctl;
    const before = { ...st };
    const today = window.InvCommon.todayNY(Date.now());
    fatal = ''; fatalAuth = false;
    setBusy('Connecting to SharePoint…');
    try {
      // Master Sheet Live, the tracker and Container Tracking first; UPC_MASTER last.
      await Promise.all([loadPoGroup(force, signal), loadCT(today, force, signal)]);
      await Promise.all([loadReport('forwarder', today, force, signal), loadReport('drayage', today, force, signal), loadAts(today, signal)]);
      await loadUpc(force, signal);
      if (force || copyCheck == null) await checkPoBookCopy(signal);
      if (signal.aborted) throw abortErr();
      applyCopyCheck();
      await emailProblems(today);
      lastLoaded = Date.now();
    } catch (e) {
      if (e && e.stop) {
        ctl.abort();   // stop the other sources still reading
        stopAll();
      } else if (isAbort(e)) {
        again = false;   // a Refresh clicked during the load doesn't restart it after Cancel
        const read = keepUnread(before, 'not read (cancelled)', 'Loading was cancelled before this was read. Click Refresh to read it.', 'click Refresh');
        toast(lastLoaded ? 'Refresh cancelled. Sources not read this time show what was read before, with its time.'
          : read ? `Loading cancelled: ${plural(read, 'source')} read, the rest not. Click Refresh to read them all.`
            : 'Loading cancelled; nothing was read. Click Refresh to try again.', 6000);
      } else {
        console.error(e);
        ctl.abort();   // stop the other sources still reading: this cycle's results are kept or marked below
        if (e && e.auth) {
          fatalAuth = true;
          fatal = `Your sign-in needs renewing (${(e.errorCode || e.message || String(e)).slice(0, 80)}). Click Sign in again; nothing is lost.`;
        } else {
          fatal = describe(e).message;
        }
        keepUnread(before, 'not read', `Not read: ${fatalAuth ? 'the sign-in needs renewing' : "the sources couldn't be loaded"} (see the red banner).`, 'see the red banner');
      }
    } finally {
      loading = null; cycleDone = true;
      setBusy(null);
      renderAll();
      schedule();
      if (again) { again = false; loadAll(true); }
    }
  }

  function schedule() {
    clearTimeout(timer);
    if (!account || scopeFatal) return;   // stopped over the sign-in's permissions: only Sign in again helps
    timer = setTimeout(() => {
      if (document.hidden || loading) { due = true; schedule(); }
      else loadAll(false);
    }, REFRESH_MS);
  }

  // ---------- drawing ----------
  // A source with no state yet is waiting for its turn, never "not set" (only a blank location in config.js says that).
  const baseWord = s => (s.state === 'load' ? 'loading' : s.state === 'ok' ? 'ok' : s.label || (s.state === 'off' ? 'waiting' : s.state));
  const staleWord = s => (s.stale ? 'as read ' + fmtShort(new Date(s.stale).toISOString()) : '');
  const statusWord = s => (s.stale ? (s.state === 'ok' ? '' : baseWord(s) + ', ') + staleWord(s) : baseWord(s));
  const readCounted = s => (s.read == null && s.counted == null ? '' : `${fmtN(s.read) || '–'} read / ${fmtN(s.counted) || '–'} counted`);

  function chipHtml(src) {
    const s = st[src.k] || fresh(src.k);
    const label = s.state === 'ok' && !s.stale ? '' : statusWord(s);
    const mailBad = s.mail && s.mail.state === 'failed';
    const line2 = [s.savedAt ? 'saved ' + fmtShort(s.savedAt) : '', s.savedBy, readCounted(s)].filter(Boolean).join(' · ') || s.short || '';
    const tip = [src.name, s.file,
      s.savedAt ? 'Saved ' + fmtLong(s.savedAt) + (s.savedBy ? ' by ' + s.savedBy : '') : '',
      readCounted(s), s.state === 'ok' ? '' : s.message, s.stale ? `Not refreshed: this is what was read ${fmtShort(new Date(s.stale).toISOString())}.` : '',
      ...s.notes, s.mail ? s.mail.text : '',
      s.webUrl ? 'Click to open it on SharePoint.' : ''].filter(Boolean).join('\n');
    const inner = `<span class="dot"></span><span class="ct"><span class="l1"><b>${esc(src.name)}</b>${label ? `<em>${esc(label)}</em>` : ''}${mailBad ? '<em>· email failed</em>' : ''}</span>${line2 ? `<small>${esc(line2)}</small>` : ''}</span>`;
    const attrs = `class="chip src ${s.state}" id="chip-${src.k}" data-state="${s.state}" title="${esc(tip)}"`;
    return s.webUrl ? `<a ${attrs} href="${esc(s.webUrl)}" target="_blank" rel="noopener">${inner}</a>` : `<span ${attrs}>${inner}</span>`;
  }
  function renderChips() {
    const el = $('chips');
    el.innerHTML = `<div class="grp" title="Read together, as the PO Book reads them: if one fails, the stage numbers stop."><span class="gl">PO Book</span>${SOURCES.filter(x => x.group).map(chipHtml).join('')}</div>${SOURCES.filter(x => !x.group).map(chipHtml).join('')}`;
  }

  function cardHtml(src) {
    const s = st[src.k] || fresh(src.k);
    const file = s.file ? (s.webUrl ? `<a href="${esc(s.webUrl)}" target="_blank" rel="noopener">${esc(s.file)}</a>` : esc(s.file)) : '';
    const sub = [file, s.savedAt ? 'saved ' + esc(fmtLong(s.savedAt)) : '', esc(s.savedBy)].filter(Boolean).join(' · ');
    const msgs = [s.state === 'ok' ? '' : s.message, ...s.notes].filter(Boolean).map(m => `<div class="msg">${esc(m)}</div>`).join('');
    const rows = (s.rows || []).map(r => (Array.isArray(r) ? `<dt>${esc(r[0])}</dt><dd>${esc(fmtV(r[1]))}</dd>` : `<dt class="h">${esc(r.h)}</dt>`)).join('');
    const iss = s.issues.length ? `<dt>Issues listed on Problems</dt><dd>${fmtN(s.issues.length)}</dd>` : '';
    return `<div class="scard ${s.state}" id="card-${src.k}"><div class="sh"><span class="dot"></span><b>${esc(src.name)}</b><span class="sp"></span><span class="pill ${s.state}">${esc(statusWord(s))}</span></div>` +
      `${sub ? `<div class="sub">${sub}</div>` : ''}${msgs}${s.mail ? `<div class="sub">${esc(s.mail.text)}</div>` : ''}` +
      `${rows || iss ? `<dl class="cnt">${rows}${iss}</dl>` : ''}</div>`;
  }
  function renderCards() { $('srcCards').innerHTML = SOURCES.map(cardHtml).join(''); }

  function renderSources() { renderChips(); renderCards(); }

  function renderHealth() {
    const el = $('health');
    if (scopeFatal) { el.innerHTML = `<span class="r"><b>Stopped: sign-in permissions</b></span>`; el.title = scopeFatal; return; }
    if (fatal) { el.innerHTML = `<span class="r"><b>Couldn't load</b></span>`; el.title = ''; return; }
    if (!cycleDone) { el.innerHTML = ''; return; }
    const bad = SOURCES.filter(x => st[x.k] && (st[x.k].state === 'err' || st[x.k].state === 'warn'));
    if (bad.length) {
      el.innerHTML = bad.map(x => `<span class="${st[x.k].state === 'err' ? 'r' : 'a'}"><b>${esc(x.name)}</b> ${esc(statusWord(st[x.k]))}</span>`).join(' · ');
      el.title = bad.map(x => `${x.name}: ${st[x.k].message}`).join('\n');
      return;
    }
    const unset = SOURCES.filter(x => st[x.k] && st[x.k].state === 'off' && st[x.k].label === 'not set').length;
    const notRead = SOURCES.filter(x => !st[x.k] || st[x.k].stale || /^not read/.test(st[x.k].label || ''));
    if (notRead.length) {
      el.innerHTML = `<span class="a"><b>Not read this time:</b> ${esc(notRead.map(x => x.name).join(', '))}</span>`;
      el.title = 'Click Refresh to read them.';
      return;
    }
    el.innerHTML = `<span class="g">${unset ? 'All sources with a location loaded' : 'All sources loaded'}</span> at ${esc(new Date(lastLoaded || Date.now()).toLocaleTimeString('en-US', { hour: 'numeric', minute: '2-digit' }))}${unset ? ` · ${plural(unset, 'source')} not set in config.js` : ''}`;
    el.title = '';
  }

  const banner = (kind, html) => `<div class="banner ${kind}"><div class="bt">${html}</div></div>`;
  // A message after "Source:" reads on from the colon ("PO Status Tracker: your account can't open it…").
  const onFromColon = m => String(m || '').replace(/^(Your|Not|It|The|No|Missing|Couldn't|Reading|Used|There|Nothing|Found)\b/, w => w.toLowerCase());
  function renderBanners() {
    const out = [];
    if (scopeFatal) out.push(`<div class="banner err" id="stopBanner"><div class="bt"><b>Stopped: nothing is loaded.</b> ${esc(scopeFatal)}</div><button type="button" data-act="signin">Sign in again</button></div>`);
    else if (fatal) out.push(`<div class="banner err"><div class="bt"><b>Couldn't load the sources.</b> ${esc(fatal)}</div>${fatalAuth ? '<button type="button" data-act="signin">Sign in again</button>' : ''}</div>`);
    else {
      const red = CORE_KEYS.filter(k => st[k] && (st[k].state === 'err' || st[k].blocked));
      const unset = CORE_KEYS.filter(k => st[k] && st[k].state === 'off' && st[k].label === 'not set');
      if (red.length) {
        // A blank location is part of the cause too (Master Sheet Live waits on a blank PO Status Tracker), so it's listed here.
        const items = red.map(k => {
          const s = st[k];
          return `<li><b>${esc(BY_K[k].name)}</b>: ${esc(onFromColon(s.message || s.label))}${s.mail ? ` <i>${esc(s.mail.text)}</i>` : ''}</li>`;
        }).concat(unset.map(k => `<li><b>${esc(BY_K[k].name)}</b>: its location is blank in config.js.</li>`)).join('');
        out.push(banner('err', `<b>The stage numbers are not loaded.</b> They come back once this is fixed:<ul>${items}</ul>`));
      } else if (unset.length) {
        out.push(banner('info', `<b>The stage numbers are not loaded:</b> the location of ${esc(unset.map(k => BY_K[k].name).join(', '))} is blank in config.js.`));
      }
      const amb = REPORT_KEYS.filter(k => st[k] && (st[k].state === 'warn' || st[k].state === 'err'));
      if (amb.length) {
        const items = amb.map(k => {
          const s = st[k];
          return `<li><b>${esc(BY_K[k].name)} report</b>: ${esc(onFromColon(s.message))}${s.state === 'err' ? ' Stages come from Container Tracking dates only.' : ''}${s.mail ? ` <i>${esc(s.mail.text)}</i>` : ''}</li>`;
        }).join('');
        out.push(banner('warn', `<b>Shipping reports:</b><ul>${items}</ul>`));
      }
    }
    $('banners').innerHTML = out.join('');
  }

  function renderSrcTable() {
    let h = '<thead><tr><th class="fz">Source</th><th>File</th><th>Saved</th><th>Saved by</th><th class="num">Read</th><th class="num">Counted</th><th>Status</th><th>Message</th></tr></thead><tbody>';
    for (const src of SOURCES) {
      const s = st[src.k] || fresh(src.k);
      const file = s.file ? (s.webUrl ? `<a href="${esc(s.webUrl)}" target="_blank" rel="noopener">${esc(s.file)}</a>` : esc(s.file)) : '';
      const msg = [s.state === 'ok' ? '' : s.message, ...s.notes].filter(Boolean).join(' ');
      h += `<tr id="srow-${src.k}"><td class="nw fz"><b>${esc(src.name)}</b></td><td class="nw">${file}</td><td class="nw">${esc(fmtLong(s.savedAt))}</td>` +
        `<td class="nw">${esc(s.savedBy)}</td><td class="num">${fmtN(s.read)}</td><td class="num">${fmtN(s.counted)}</td>` +
        `<td class="nw"><span class="pill ${s.state}">${esc(statusWord(s))}</span></td><td class="msg">${esc(msg)}` +
        `${s.headers ? `<div class="hd">Headers seen: ${esc(s.headers)}</div>` : ''}` +
        `${s.mail ? `<div class="ml ${s.mail.state === 'failed' ? 'bad' : ''}">${esc(s.mail.text)}</div>` : ''}</td></tr>`;
    }
    $('srcTable').innerHTML = h + '</tbody>';
  }

  let pq = '';
  const expanded = new Set();
  const allIssues = () => SOURCES.flatMap(x => (st[x.k] && st[x.k].state !== 'load' ? st[x.k].issues : []));
  function renderIssues() {
    const all = allIssues(), q = pq.toLowerCase();
    const groups = new Map();
    for (const i of all) {
      if (q && ![i.source, i.code, codeLabel(i.code), i.sheet, i.row, i.po, i.style, i.message].join(' ').toLowerCase().includes(q)) continue;
      if (!groups.has(i.source)) groups.set(i.source, []);
      groups.get(i.source).push(i);
    }
    const order = SOURCES.map(x => x.name);
    const names = [...groups.keys()].sort((a, b) => ((order.indexOf(a) + 1) || 99) - ((order.indexOf(b) + 1) || 99) || a.localeCompare(b));
    let h = '<thead><tr><th class="fz">What</th><th>Sheet</th><th class="num">Row / item</th><th>PO</th><th>Style</th><th>Message</th></tr></thead><tbody>';
    let matched = 0;
    for (const name of names) {
      const list = groups.get(name);
      matched += list.length;
      const byCode = new Map();
      for (const i of list) byCode.set(i.code, (byCode.get(i.code) || 0) + 1);
      const codes = [...byCode].sort((a, b) => b[1] - a[1]).map(([c, n]) => `<span title="${esc(c || '')}">${esc(codeLabel(c))} ${fmtN(n)}</span>`).join(' · ');
      h += `<tr class="gh"><td colspan="6"><b>${esc(name)}</b> · ${plural(list.length, 'issue')}<span class="codes">${codes}</span></td></tr>`;
      const show = expanded.has(name) ? list : list.slice(0, ISSUE_PAGE);
      for (const i of show) {
        h += `<tr><td class="nw fz" title="${esc(i.code)}">${esc(codeLabel(i.code))}</td><td class="nw">${esc(i.sheet)}</td><td class="num">${esc(i.row)}</td><td class="nw">${esc(i.po)}</td><td class="nw">${esc(i.style)}</td><td class="msg">${esc(i.message)}</td></tr>`;
      }
      if (show.length < list.length) {
        h += `<tr><td colspan="6" class="more">Showing ${fmtN(show.length)} of ${fmtN(list.length)}. <button type="button" class="ghost" data-more="${esc(name)}">Show all ${fmtN(list.length)}</button></td></tr>`;
      }
    }
    if (!matched) {
      const why = all.length ? 'No issues match the filter.' : lastLoaded ? 'No issues found in the files that loaded.' : '';
      h += `<tr><td colspan="6" class="empty">${esc(why)}</td></tr>`;
    }
    $('issTable').innerHTML = h + '</tbody>';
    $('pqInfo').textContent = q ? `${fmtN(matched)} of ${plural(all.length, 'issue')} match` : plural(all.length, 'issue');
    $('cProb').textContent = fmtN(all.length);
    document.querySelector('#tabs .tab.iss').classList.toggle('bad', SOURCES.some(x => st[x.k] && st[x.k].state === 'err'));
  }
  function renderProblems() { renderSrcTable(); renderIssues(); }

  function renderAll() {
    renderSources(); renderHealth(); renderBanners(); renderProblems();
  }

  function showTab(t) {
    if (!['inv', 'pull', 'problems'].includes(t)) t = 'inv';
    tab = t;
    for (const b of document.querySelectorAll('#tabs .tab')) {
      b.classList.toggle('on', b.dataset.tab === t);
      b.setAttribute('aria-selected', String(b.dataset.tab === t));
    }
    for (const x of ['inv', 'pull', 'problems']) $('tab-' + x).hidden = x !== t;
    store.set('tab', t);
  }

  // ---------- toast ----------
  function toast(t, ms = 3500) {
    const el = $('toast');
    el.textContent = t; el.style.display = 'block';
    clearTimeout(el._t); el._t = setTimeout(() => { el.style.display = 'none'; }, ms);
  }

  // ---------- start ----------
  async function start() {
    $('signin').hidden = true; $('app').hidden = false; $('chips').hidden = false;
    $('btnRefresh').hidden = false; $('who').hidden = false;
    $('who').innerHTML = `<b>${esc(account.name || account.username || '')}</b><button type="button" class="ghost" id="btnSignout" title="Sign out of Shalam Tool">Sign out</button>`;
    $('btnSignout').onclick = signOut;
    showTab(store.get('tab') || 'inv');
    store.prune('loc:', text => SOURCES.some(x => locText(x.k) === text));   // locations no longer in config.js
    scopeFatal = '';
    renderAll();
    // Every request checks its token's permissions before it is sent (g), so one that allows changes stops the load
    // before anything is read.
    loadAll(true);
  }

  // A new token by signing in again (by redirect the page leaves and comes back; framed, a popup).
  async function reSignIn() {
    try {
      if (framed) {
        const r = await pca.acquireTokenPopup({ scopes: SCOPES, account });
        account = r.account || account;
        start();
      } else await pca.acquireTokenRedirect({ scopes: SCOPES, account, redirectStartPage: location.href });
    } catch (e) { toast("Sign-in didn't finish: " + (e.errorMessage || e.message || e), 6000); }
  }

  function bindUi() {
    $('btnSignin').onclick = signIn;
    $('btnRefresh').onclick = () => loadAll(true);
    $('btnCancelBusy').onclick = $('btnCancelLoad').onclick = () => { if (loading) loading.abort(); };
    $('banners').addEventListener('click', e => { if (e.target.closest('[data-act="signin"]')) reSignIn(); });
    $('tabs').addEventListener('click', e => { const b = e.target.closest('.tab'); if (b) showTab(b.dataset.tab); });
    let qt = null;
    $('pq').addEventListener('input', e => { clearTimeout(qt); qt = setTimeout(() => { pq = e.target.value.trim(); renderIssues(); }, 150); });
    $('issTable').addEventListener('click', e => { const b = e.target.closest('button[data-more]'); if (b) { expanded.add(b.dataset.more); renderIssues(); } });
    document.addEventListener('visibilitychange', () => {
      if (!document.hidden && due && account) { due = false; loadAll(false); }
    });
  }

  // Drifting dots behind the glass, as on the PO Book; still under prefers-reduced-motion.
  function startBg() {
    const c = $('bg'), x = c && c.getContext && c.getContext('2d');
    if (!x) return;
    let seed = 20261028;
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
    const parts = [['InvCommon', 'core/common.js'], ['PoBookCore', 'core/pobook.js'], ['InvUpc', 'core/upc.js'], ['InvCT', 'core/ct.js'],
      ['InvShipping', 'core/shipping.js'], ['XLSX', 'vendor/xlsx-js-style/xlsx.bundle.js']];
    const missing = parts.filter(([name]) => !window[name]).map(p => p[1]);
    if (missing.length) { showSignin(`Part of the page didn't load (${missing.join(', ')}). Reload the page; if it keeps happening, tell Saul.`, true); return; }
    initAuth();
  }

  boot();
})();
