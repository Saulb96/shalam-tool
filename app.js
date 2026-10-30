// Shalam Tool: sign-in, reading every source from SharePoint, the source chips, the failure rules, the error email,
// what can be sold (Available to sell: stock after customer orders, and goods from China with one date each), where
// the goods are (Where it is and Pull from China tabs, the side panel) and the Problems tab. Parsing, the stage rules
// and the selling rules live in core/*.js; this file fetches, calls them and draws.
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
  // What the stages can't see without each report (the amber banner and the error email say it).
  const REPORT_EFFECT = {
    forwarder: "sailings and rolled vessels can't be seen, so a PO entry counts On the water once Container Tracking's ETD passes and At port once its ETA passes",
    drayage: "pickups and deliveries can't be seen, so a container on an Open round stays At port until its round moves to Complete",
  };
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
  const SRC_LABELS = { forwarder: { records: 'Shipments' }, drayage: { records: 'Containers' },
    ats: { styleRows: 'Style rows', grandTotalRows: 'Grand Total rows', noStyleRows: 'Rows with no style', multiWarehouseStyles: 'Styles on more than one row (summed)',
      notANumber: "Rows with a figure that isn't a number", blankNumbers: 'Rows with a blank In WH or customer orders', unknownStyles: "Styles whose stock can't be read",
      inWh: 'In WH, all rows (units)', orders: 'Customer orders, all rows (units)', negativeInWh: 'Styles with In WH below zero', erpAtsDiffers: "Styles where the ERP's ATS differs" } };

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
    // Incoming: matching Master Sheet Live to Container Tracking, and the stages (core/stages.js)
    'near-miss': 'Near-miss: PO or style written differently', unmatched: 'On Container Tracking, not on Master Sheet Live',
    'ct-more': 'Container Tracking holds more than Master Sheet Live', short: 'Short-shipped',
    'shipped-not-on-ct': 'Shipped, not on Container Tracking', 'msl-behind-ct': 'Master Sheet behind Container Tracking',
    'ct-behind-msl': 'Container Tracking behind Master Sheet', 'both-tabs': 'On Booking and Loaded, same round (counted once)',
    'no-etd': 'Booked or loaded, no ETD yet', 'etd-passed-booked': 'ETD passed, still on Booking',
    'unknown-type': 'PO type unknown (not on the PO Status Tracker)', 'no-case-pack': 'No case pack (Master) for the style',
    'part-case': 'Quantity not in whole cases', balance: "Units don't add up (a bug in Shalam Tool)",
    'page-sum': "The page's sums differ from the stage engine's (a bug in Shalam Tool)",
    'missing-column': 'Column missing on Master Sheet Live', 'negative-qty': 'Quantity below zero',
    'other-trip': 'Container number reused (another trip)', 'round-open-and-complete': 'Round on both Open and Complete',
    'complete-in-transit': 'On Complete, but a report still says in transit', 'dup-round': 'Repeat of a line inside a Round',
    // Container Tracking sections with no round (config.js's sections), and POs below the tracker's lowest PO
    'below-tracker': 'Older than the PO Status Tracker', 'ignored-section': 'Ignored section', 'section-not-listed': 'Section not in the list',
    'pull-section': 'Straight to the customer from China', 'planning-section': 'Planned, not booked yet',
    // The ATS report (core/ats.js) and Available to sell (core/sell.js)
    'negative-in-wh': 'In warehouse below zero', 'erp-ats-differs': "The ERP's ATS isn't In WH minus customer orders",
    'not-a-number': 'Not a number', 'grand-total': "Grand Total row missing or repeated (the report can't be checked)",
    'negative-orders': 'Customer orders below zero',
    'dates-disagree': 'Ready dates disagree (more than 7 days apart)', 'no-date': 'No date to sell from',
    oversold: 'Oversold: more customer orders than stock and goods coming', 'orders-unknown': 'Customer orders not known (no ATS report, or not on it)',
    'erp-only': 'On the ATS report only', estimate: 'Date estimated from the order date',
    'erp-received': 'Sailed, but probably received already (not In Trans on the ATS report)',
    'erp-in-transit': 'Orders may be covered by goods In Trans on the ATS report (landed, not received?)',
    'orders-may-include-sold': 'Orders may be the goods sold to ship from China',
    // flags on a style or lot of Available to sell
    now: 'Pull-from date passed: available now', 'part-taken': 'Partly taken by customer orders', taken: 'All taken by customer orders',
    'mixed-case-pack': 'Lots in different case packs', 'stock-unknown': "Stock on the ATS report can't be read", 'not-on-report': 'Not on the ATS report',
  };
  // The same code from another file, where its words differ (the ATS report's rows aren't product lines).
  const SOURCE_LABELS = { 'ATS report': { 'no-style': 'Row with no style' } };
  const codeLabel = (c, src) => (SOURCE_LABELS[src] || {})[c] || CODE_LABELS[c] || (c ? String(c).replace(/-/g, ' ') : '(no code)');

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
  // data: what the parser gave, set only when this state's file was read and used: the stage numbers take a source's
  // data from its state, so a chip and the numbers can never disagree about which file they come from.
  function fresh(k) {
    return { k, state: 'off', label: '', message: '', short: '', headers: '', file: '', webUrl: '', savedAt: '', savedBy: '',
      read: null, counted: null, rows: [], issues: [], notes: [], mailProblem: '', mail: null, blocked: false, status: 0, stale: null, data: null };
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
  let loading = null, timer = null, due = false, again = false, lastLoaded = null, fatal = '', tab = 'sell';
  let fatalAuth = false;   // the fatal message asks to sign in again
  let cycleDone = false;   // a load cycle has ended (finished, cancelled or failed)
  let scopeFatal = '';     // the sign-in isn't read-only, or can't read SharePoint: nothing loads until it's fixed
  let copyCheck = null;   // null: not known · true: same rules as the PO Book · false: the PO Book's rules changed
  let mailScopeMissing = false;
  const mailMemo = new Map();   // source|day -> what happened to its email, in this page session
  // Parsed results, kept for the stage and selling numbers.
  const data = { po: null, ct: null, forwarder: null, drayage: null, upc: null, ats: null };

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
    // The stage engine joins these lines to Container Tracking on normPo / normStyle, never on l.key or l.base: the
    // PO Book's keys fold inner spaces and -F, which this tool keeps apart. It reads each line's Barcode, Master,
    // Category and Sub from the raw rows (mrows) by the line's row number.
    data.po = { trk, src, lines, ov, cfg, counts: cm.counts, mrows };
    if (!cm.ok || !src.lines.length) fail(S.msl, problem('empty', unprefix('Master Sheet Live', cm.error) || 'No PO lines found on Master Sheet Live.'));
    else { ok(S.msl); S.msl.data = data.po; }
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
        else { ok(s); data.ct = r; s.data = r; }
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
        data[k] = r; s.data = r;
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

  // ---------- ATS: the newest weekly ERP report in the ATS folder, read with core/ats.js (nothing is ever uploaded) ----------
  // Where to save the report, in plain words: "Inventory/ATS on the Production site" for /sites/Production/<library>/Inventory/ATS.
  function atsPlace() {
    const parts = locText('ats').split('/').filter(Boolean);
    return /^\/sites\//i.test(locText('ats')) && parts.length > 3 ? `${parts.slice(3).join('/')} on the ${parts[1]} site` : 'the ATS folder set in config.js';
  }
  const isBook = n => /\.(xlsx|xlsm|xls)$/i.test(String(n || ''));
  // The report's date: from its file name ("… 10-29-2026 12-58 PM.xlsx"), else the day it was saved. Its age decides
  // amber (after 8 days) and the error email (after 10).
  async function loadAts(today, force, signal) {
    const s = fresh('ats');
    showLoading('ats');
    const noReport = why => off(s, 'no report yet', `${why || 'No ATS report yet.'} Save the weekly ATS report into ${atsPlace()}.`, `save it into ${atsPlace()}`);
    await attempt(s, async () => {
      let at;
      try { at = await locate('ats', s, signal); }
      catch (e) { if (e.missing) return noReport(); throw e; }   // only the folder itself missing; a wrong site or library stays red
      if (!at) return notSet(s);
      const folder = at.item;
      s.file = folder.name || s.file; s.webUrl = folder.webUrl || s.webUrl;
      setBusy('Looking in the ATS folder…');
      const files = (await gAll(`/drives/${at.driveId}/items/${folder.id}/children?$select=id,name,lastModifiedDateTime,lastModifiedBy,file,webUrl&$top=999`, signal))
        .filter(x => x.file && !/^~\$/.test(x.name || ''));
      const kids = files.filter(x => isBook(x.name));
      s.rows = [['Files in the folder', files.length]];
      if (!kids.length) return noReport(files.length ? `No Excel file in the ATS folder (${plural(files.length, 'other file')}).` : '');
      // The newest report: by the date in its file name (a report opened and saved again keeps its own date), else the
      // day it was saved; then the latest saved.
      const dayOfKid = x => { const d = window.InvAts.dayFromName(x.name); if (d != null) return d; const ms = Date.parse(x.lastModifiedDateTime); return Number.isFinite(ms) ? window.InvCommon.todayNY(ms) : -Infinity; };
      const f = kids.slice().sort((a, b) => (dayOfKid(b) - dayOfKid(a)) || String(b.lastModifiedDateTime).localeCompare(String(a.lastModifiedDateTime)))[0];
      s.file = f.name; s.webUrl = f.webUrl || s.webUrl; s.savedAt = f.lastModifiedDateTime || ''; s.savedBy = who(f.lastModifiedBy);
      const meta = await gj(`/drives/${at.driveId}/items/${f.id}`, { signal });   // with its download link
      s.file = meta.name || f.name; s.webUrl = meta.webUrl || s.webUrl;
      s.savedAt = meta.lastModifiedDateTime || s.savedAt; s.savedBy = who(meta.lastModifiedBy) || s.savedBy;
      let c = cache['file:ats'];
      if (force || !unchanged(c, meta)) {
        const buf = await fetchBytes(meta, 'the ATS report', signal);
        c = { id: meta.id, lm: meta.lastModifiedDateTime, wb: readBook(buf, meta.name) };
        cache['file:ats'] = c;
      }
      setBusy('Reading the ATS report…'); await yieldUi();
      let r;
      try { r = window.InvAts.parseAts(XLSX, c.wb, meta.name, today); }
      catch (e) { console.error(e); throw problem("can't read the file", 'Reading the ATS report failed: ' + ((e && e.message) || e)); }
      if (!r || typeof r !== 'object') throw problem("can't read the file", "Reading the ATS report gave nothing back.");
      const ms = Date.parse(s.savedAt);
      const savedDay = Number.isFinite(ms) ? window.InvCommon.todayNY(ms) : null;
      const fromName = typeof r.reportDay === 'number';
      const day = fromName ? r.reportDay : savedDay;
      const age = day == null ? null : today - day;
      s.reportDay = day; s.reportFrom = fromName ? 'file name' : day == null ? '' : 'saved';
      const n = r.styles instanceof Map ? r.styles.size : null;
      // Rows read (style rows, blank rows and the Grand Total) against style rows counted.
      s.read = firstNum(r.counts, ['rowsRead', 'rows']); s.counted = firstNum(r.counts, ['styleRows']) ?? n;
      s.rows.push(['Newest file', meta.name],
        ['Report date', day == null ? 'not known' : `${window.InvCommon.fmtDay(day)} (${fromName ? 'from the file name' : 'the day it was saved: no date in the file name'})`],
        ['Age', age == null ? '' : plural(age, 'day')], ...countRows(r.counts, SRC_LABELS.ats), ...(r.counts && typeof r.counts.styles === 'number' ? [] : [['Styles', n]]));
      s.issues = (r.issues || []).map(i => normIssue('ATS report', { ...i, source: 'ATS report' }));
      if (!r.ok) {
        const msg = unprefix('ATS report', String(r.error || '')) || "The ATS report couldn't be read.";
        const seen = r.seen || r.headers;
        throw problem(/add up/i.test(msg) ? "doesn't add up" : labelFor(msg), msg, { headers: Array.isArray(seen) ? seen.map(x => window.InvCommon.S(x)).filter(Boolean).join(' | ') : '' });
      }
      if (!n) throw problem('empty', `${meta.name} has no style rows.`);
      const asOf = day == null ? '' : ` (report date ${window.InvCommon.fmtDay(day)})`;
      if (age != null && age > ATS_MAIL_DAYS) {
        warn(s, `${plural(age, 'day')} old`, `${meta.name} is ${plural(age, 'day')} old${asOf}: the weekly ATS report is late, so stock and customer orders may be out of date.`);
        s.mailProblem = `the newest report is ${plural(age, 'day')} old`;
      } else if (age != null && age > ATS_AMBER_DAYS) {
        warn(s, `${plural(age, 'day')} old`, `${meta.name} is ${plural(age, 'day')} old${asOf}, so stock and customer orders may be out of date.`);
      } else if (age != null && age < 0) {
        warn(s, 'date in the future', `The date in the file name of ${meta.name} (${window.InvCommon.fmtDay(day)}) is after today: check the file name's date, the report may not be the newest.`);
      } else ok(s);
      data.ats = r; s.data = r;
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
        else { ok(s); data.upc = r; s.data = r; }
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
    "doesn't add up": "Open the report and check its rows against its Grand Total row (it may have been cut short or edited); save a fresh copy from the ERP into the folder.",
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
  function effectOf(src, s) {
    if (CORE_KEYS.includes(src.k)) return "The stage numbers stay off until it's fixed, and Available to sell with them.";
    if (REPORT_KEYS.includes(src.k)) return `Until it's fixed, ${REPORT_EFFECT[src.k]}.`;
    if (src.k === 'ats') {
      return s && s.state === 'warn' ? 'Stock and customer orders on Available to sell come from this report, so they may be out of date.'
        : 'Stock and customer orders on Available to sell come from this report: until it\'s fixed, Available now says "not loaded" and customer orders aren\'t taken off the goods from China.';
    }
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
    const eff = effectOf(src, s);
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
    stages = null; stageErr = ''; idx = null;
    sell = null; sellErr = ''; sx = null;
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
      await Promise.all([loadReport('forwarder', today, force, signal), loadReport('drayage', today, force, signal), loadAts(today, force, signal)]);
      // The stage numbers as soon as their sources are in; UPC_MASTER (descriptions, categories) fills in after.
      if (!signal.aborted) { buildStages(today); renderStageViews(); }
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
      schedule();   // first, so the 5-minute refresh re-arms whatever the drawing below does
      // After every cycle, finished or not: built from what each chip shows as read (a stopped page has nothing).
      if (scopeFatal) { stages = null; stageErr = ''; idx = null; sell = null; sellErr = ''; sx = null; } else buildStages(today);
      try { markNew(); } catch (e) { console.error(e); }
      renderAll();
      if (again) { again = false; loadAll(true); }
    }
  }

  function schedule() {
    clearTimeout(timer);
    if (!account || scopeFatal) return;   // stopped over the sign-in's permissions: only Sign in again helps
    // Never while the side panel is open: it waits until the panel closes.
    timer = setTimeout(() => {
      if (document.hidden || loading || panel) { due = true; schedule(); }
      else loadAll(false);
    }, REFRESH_MS);
  }
  // A refresh that came due while the page was hidden or the panel open runs once both are past.
  function runDue() {
    if (due && account && !document.hidden && !panel && !loading) { due = false; loadAll(false); }
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
    const [xh, xt] = stageHealth();
    const bad = SOURCES.filter(x => st[x.k] && (st[x.k].state === 'err' || st[x.k].state === 'warn'));
    if (bad.length) {
      el.innerHTML = bad.map(x => `<span class="${st[x.k].state === 'err' ? 'r' : 'a'}"><b>${esc(x.name)}</b> ${esc(statusWord(st[x.k]))}</span>`).join(' · ') + xh;
      el.title = bad.map(x => `${x.name}: ${st[x.k].message}`).concat(xt).join('\n');
      return;
    }
    const unset = SOURCES.filter(x => st[x.k] && st[x.k].state === 'off' && st[x.k].label === 'not set').length;
    const notRead = SOURCES.filter(x => !st[x.k] || st[x.k].stale || /^not read/.test(st[x.k].label || ''));
    if (notRead.length) {
      el.innerHTML = `<span class="a"><b>Not read this time:</b> ${esc(notRead.map(x => x.name).join(', '))}</span>` + xh;
      el.title = ['Click Refresh to read them.', ...xt].join('\n');
      return;
    }
    el.innerHTML = `<span class="g">${unset ? 'All sources with a location loaded' : 'All sources loaded'}</span> at ${esc(new Date(lastLoaded || Date.now()).toLocaleTimeString('en-US', { hour: 'numeric', minute: '2-digit' }))}${unset ? ` · ${plural(unset, 'source')} not set in config.js` : ''}` + xh;
    el.title = xt.join('\n');
  }
  // The health line's stage part: Unmatched units, and the balance when it doesn't add up. [html, tooltip lines]
  function stageHealth() {
    if (stageErr) return [' · <span class="r"><b>Stage numbers not worked out</b></span>', ['The stage numbers couldn\'t be worked out (see the red banner).']];
    if (!stages || !stages.ok) return ['', []];
    const um = (stages.totals.units && stages.totals.units.unmatched) || 0, o = balanceOff();
    const h = [`<span class="${um ? 'a' : 'g'}">Unmatched <b>${fmtN(um)}</b> units</span>`];
    const t = [`Unmatched: ${fmtN(um)} units on Container Tracking with no Master Sheet Live line (not landed), never in the stage totals.`];
    if (balanceBad()) {
      h.push(`<span class="r"><b>${o.msl ? `Balance off by ${fmtN(Math.abs(o.msl))} units` : o.bad ? `Balance off in ${plural(o.bad, 'group')}` : `Container Tracking balance off by ${fmtN(Math.abs(o.ct))} units`}</b></span>`);
      t.push(`Balance: ${fmtN(stages.balance.mslUnits)} units on Master Sheet Live, ${fmtN(stages.balance.outcomeUnits)} in the outcomes (a bug in Shalam Tool; see Problems).`);
    } else t.push(`Balance: every one of the ${fmtN(stages.balance.mslUnits)} units on Master Sheet Live is in exactly one outcome.`);
    return [' · ' + h.join(' · '), t];
  }
  // How far each balance is off: Master Sheet Live units against their outcomes (in total and per PO + style), and
  // Container Tracking units against where they were placed. Anything but zero is a bug, and the page says so.
  function balanceOff() {
    if (!stages || !stages.ok) return null;
    const b = stages.balance;
    return { msl: units(b.diff), bad: Array.isArray(b.bad) ? b.bad.length : 0, ct: units(b.ct && b.ct.diff) };
  }
  const balanceBad = () => { const o = balanceOff(); return !!o && !!(o.msl || o.bad || o.ct); };

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
      if (stageErr) {
        out.push(`<div class="banner err" id="stageBanner"><div class="bt"><b>The stage numbers couldn't be worked out:</b> ${esc(stageErr)}. This is a bug in Shalam Tool, not a problem in the files: tell Saul. Every stage number says "not loaded" until it's fixed.</div></div>`);
      } else if (balanceBad()) {
        const b = stages.balance, o = balanceOff();
        const what = [o.msl || o.bad ? `${fmtN(b.mslUnits)} units on Master Sheet Live, ${fmtN(b.outcomeUnits)} in the outcomes${o.msl ? `, ${fmtN(Math.abs(o.msl))} apart` : ''}${o.bad ? ` (${plural(o.bad, 'PO + style group')} off)` : ''}` : '',
          o.ct ? `Container Tracking has ${fmtN(b.ct.units)} units on its lines, ${fmtN(b.ct.outcomeUnits)} were placed` : ''].filter(Boolean).join('; ');
        out.push(`<div class="banner err" id="balanceBanner"><div class="bt"><b>Units don't add up:</b> ${what}. This is a bug in Shalam Tool, not a problem in the files: tell Saul. The stage numbers may be wrong until it's fixed; Problems has the details.</div></div>`);
      }
      if (sellErr && !stageErr) {
        out.push(`<div class="banner err" id="sellErrBanner"><div class="bt"><b>Available to sell couldn't be worked out:</b> ${esc(sellErr)}. This is a bug in Shalam Tool, not a problem in the files: tell Saul. Available to sell says "not loaded" until it's fixed; the other tabs still work.</div></div>`);
      }
      // A report missing or not used: the stages it would have given come from Container Tracking's own dates.
      const amb = REPORT_KEYS.filter(k => st[k] && (st[k].state === 'warn' || st[k].state === 'err'));
      const unused = stages && stages.ok ? REPORT_KEYS.filter(k => !got(k)) : [];
      if (amb.length || unused.length) {
        const items = amb.map(k => {
          const s = st[k];
          return `<li><b>${esc(BY_K[k].name)} report</b>: ${esc(onFromColon(s.message))}${s.mail ? ` <i>${esc(s.mail.text)}</i>` : ''}</li>`;
        }).concat(unused.filter(k => !amb.includes(k)).map(k => {
          const s = st[k];
          const why = !s || s.state === 'load' ? 'not read yet.' : s.label === 'not set' ? 'its location is blank in config.js.' : onFromColon(s.message || statusWord(s));
          return `<li><b>${esc(BY_K[k].name)} report</b>: ${esc(why)}</li>`;
        })).join('');
        let head = '<b>Shipping reports:</b>';
        if (unused.length) {
          // Said by which report is missing: with one still used, only the blocks no report covers use dates alone.
          const n = stages.counts && stages.counts.trackerOnly, one = n === 1, num = typeof n === 'number';
          const entries = num ? `${fmtN(n)} PO ${one ? 'entry' : 'entries'} on Booking and Loaded` : '';
          const takes = `take${one ? 's its' : ' their'} stage from Container Tracking's ETD and ETA alone`;
          if (unused.length > 1) {
            head = `<b>Stages from Container Tracking dates only:</b> the Forwarder and Drayage reports aren't used, so ${num ? `${entries} ${takes}` : "stages come from Container Tracking's ETD and ETA alone"}.`;
          } else {
            const k = unused[0];
            head = `<b>${esc(BY_K[k].name)} report not used:</b> ${esc(REPORT_EFFECT[k])}.` +
              (num ? ` ${entries} ${one ? 'has' : 'have'} no report covering ${one ? 'it' : 'them'} and ${takes}.` : '');
          }
        }
        out.push(`<div class="banner warn" id="reportsBanner"><div class="bt">${head}<ul>${items}</ul></div></div>`);
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

  let pq = '', showChecked = false;
  const expanded = new Set();
  // The stage engine's issues come first, under "Incoming"; then the selling engine's, under "Available to sell"; then
  // every parser's, by source (the ATS report's under "ATS report").
  const stageIssues = () => (stages && stages.ok ? [...(stages.issues || []).map(i => normIssue('Incoming', i)), ...pageIssues] : []);
  const sellIssues = () => (sell && sell.ok ? [...(sell.issues || []).map(i => normIssue(SELL_SRC, { ...i, source: SELL_SRC })), ...sellPageIssues] : []);
  const allIssues = () => [...stageIssues(), ...sellIssues(), ...SOURCES.flatMap(x => (st[x.k] && st[x.k].state !== 'load' ? st[x.k].issues : []))];

  // A problem marked checked is kept in this browser only, under inv:checked:<source>:<code>:<po>:<style>:<row>, with
  // its message: it comes back when the message changes. With storage blocked the mark lasts for this page session.
  const issKey = i => [i.source, i.code, i.po, i.style, i.row].join(':');
  const checkedMem = new Map();
  function isChecked(i) {
    const k = 'checked:' + issKey(i);
    const v = checkedMem.has(k) ? checkedMem.get(k) : store.get(k);
    return v != null && v === i.message;
  }
  let checkedTold = false;
  function setChecked(i, on) {
    const k = 'checked:' + issKey(i);
    if (on) {
      checkedMem.set(k, i.message);
      const kept = store.set(k, i.message);
      if (!checkedTold || !kept) toast(kept ? 'Marked checked: it\'s hidden until its message changes. "Show checked" brings it back.' : "Marked checked for now: this browser doesn't keep marks, so it comes back next time.", 6000);
      checkedTold = true;
    } else { checkedMem.set(k, null); store.del(k); }
  }
  // "New since you last looked": each issue's key and message, by source, as of the last time this person had the
  // Problems tab open, kept in this browser (inv:seen). Refreshes only work out what is new against it; leaving
  // Problems (or the page while on it) marks what it showed as seen. A source that wasn't read this time keeps its
  // earlier list, so its issues aren't new when it's back; a source seen for the first time has nothing new.
  let seenBy = null, curBy = new Map(), newSet = new Set();
  function hash(t) { let h = 0x811c9dc5; for (let i = 0; i < t.length; i++) { h ^= t.charCodeAt(i); h = Math.imul(h, 0x01000193); } return (h >>> 0).toString(36); }
  const issHash = i => hash(issKey(i) + '\n' + i.message);
  const saveSeen = () => store.set('seen', JSON.stringify(Object.fromEntries([...seenBy].map(([k, v]) => [k, [...v]]))));
  function markNew() {
    if (!seenBy) {
      try { const j = JSON.parse(store.get('seen') || 'null'); seenBy = j && typeof j === 'object' ? new Map(Object.entries(j).map(([k, v]) => [k, new Set(v)])) : null; } catch (e) { seenBy = null; }
      if (!seenBy) seenBy = new Map();
    }
    const cur = new Map(), keys = new Map();
    for (const i of allIssues()) {
      if (!cur.has(i.source)) { cur.set(i.source, new Set()); keys.set(i.source, new Set()); }
      cur.get(i.source).add(issHash(i)); keys.get(i.source).add(issKey(i));
    }
    const read = SOURCES.filter(x => st[x.k] && !st[x.k].stale && (st[x.k].state === 'ok' || st[x.k].state === 'warn')).map(x => x.name);
    if (stages && stages.ok) read.push('Incoming');
    if (sell && sell.ok) read.push(SELL_SRC);
    if (st.ats && !st.ats.stale && st.ats.data) read.push('ATS report');
    newSet = new Set(); curBy = new Map();
    let first = false;
    for (const name of new Set([...cur.keys(), ...read])) {
      const set = cur.get(name) || new Set(), was = seenBy.get(name);
      curBy.set(name, set);
      if (was) { for (const h of set) if (!was.has(h)) newSet.add(h); } else { seenBy.set(name, set); first = true; }
    }
    if (first) saveSeen();
    // Checked marks of a source read this time whose problem is gone: dropped, so old marks don't pile up.
    const readSet = new Set(read);
    store.prune('checked:', rest => { const src = rest.slice(0, rest.indexOf(':')); return !readSet.has(src) || (keys.has(src) && keys.get(src).has(rest)); });
  }
  // The person has had the Problems tab open: what it showed is now seen (the "new" marks go on the next load).
  function commitSeen() {
    if (!seenBy || !curBy.size) return;
    for (const [name, set] of curBy) seenBy.set(name, set);
    saveSeen();
  }

  let shownIss = [];
  function renderIssues() {
    const all = allIssues(), q = pq.toLowerCase();
    const groups = new Map();
    let nChecked = 0, open = 0;
    for (const i of all) {
      const ck = isChecked(i);
      if (ck) nChecked++; else open++;
      if (ck && !showChecked) continue;
      if (q && ![i.source, i.code, codeLabel(i.code, i.source), i.sheet, i.row, i.po, i.style, i.message].join(' ').toLowerCase().includes(q)) continue;
      if (!groups.has(i.source)) groups.set(i.source, []);
      groups.get(i.source).push({ i, ck, nw: newSet.has(issHash(i)) });
    }
    const order = ['Incoming', SELL_SRC, 'ATS report', ...SOURCES.map(x => x.name)];
    const names = [...groups.keys()].sort((a, b) => ((order.indexOf(a) + 1) || 99) - ((order.indexOf(b) + 1) || 99) || a.localeCompare(b));
    let h = '<thead><tr><th class="fz">What</th><th>Sheet</th><th class="num">Row / item</th><th>PO</th><th>Style</th><th>Message</th></tr></thead><tbody>';
    let matched = 0;
    shownIss = [];
    for (const name of names) {
      // New since the last load first, each part in the order the files gave.
      const list = groups.get(name).map((x, n) => ({ ...x, n })).sort((a, b) => (b.nw - a.nw) || (a.n - b.n));
      matched += list.length;
      const byCode = new Map();
      for (const { i } of list) byCode.set(i.code, (byCode.get(i.code) || 0) + 1);
      const codes = [...byCode].sort((a, b) => b[1] - a[1]).map(([c, n]) => `<span title="${esc(c || '')}">${esc(codeLabel(c, name))} ${fmtN(n)}</span>`).join(' · ');
      const nNew = list.filter(x => x.nw).length;
      h += `<tr class="gh"><td colspan="6"><b>${esc(name)}</b> · ${plural(list.length, 'issue')}${nNew ? ` · <b>${fmtN(nNew)} new</b>` : ''}<span class="codes">${codes}</span></td></tr>`;
      const show = expanded.has(name) ? list : list.slice(0, ISSUE_PAGE);
      for (const { i, ck, nw } of show) {
        const ix = shownIss.push(i) - 1;
        h += `<tr class="${nw ? 'isnew' : ''}${ck ? ' isck' : ''}"><td class="nw fz" title="${esc(i.code)}"><input type="checkbox" class="ck" data-ix="${ix}"${ck ? ' checked' : ''} title="Checked: hidden from this list until its message changes" aria-label="Checked">${nw ? '<span class="pill new">new</span>' : ''}${esc(codeLabel(i.code, i.source))}</td><td class="nw">${esc(i.sheet)}</td><td class="num">${esc(i.row)}</td><td class="nw">${esc(i.po)}</td><td class="nw">${esc(i.style)}</td><td class="msg">${esc(i.message)}</td></tr>`;
      }
      if (show.length < list.length) {
        h += `<tr><td colspan="6" class="more">Showing ${fmtN(show.length)} of ${fmtN(list.length)}. <button type="button" class="ghost" data-more="${esc(name)}">Show all ${fmtN(list.length)}</button></td></tr>`;
      }
    }
    if (!matched) {
      const why = all.length ? (open || showChecked ? 'No issues match the filter.' : 'Every issue is marked checked.') : lastLoaded ? 'No issues found in the files that loaded.' : '';
      h += `<tr><td colspan="6" class="empty">${esc(why)}</td></tr>`;
    }
    $('issTable').innerHTML = h + '</tbody>';
    const ckText = nChecked ? ` · ${fmtN(nChecked)} checked${showChecked ? '' : ', hidden'}` : '';
    $('pqInfo').textContent = (q ? `${fmtN(matched)} of ${plural(all.length, 'issue')} match` : plural(all.length, 'issue')) + ckText;
    $('pqCheckedWrap').hidden = !nChecked && !showChecked;
    $('nChecked').textContent = fmtN(nChecked);
    $('cProb').textContent = fmtN(open);
    document.querySelector('#tabs .tab.iss').classList.toggle('bad', SOURCES.some(x => st[x.k] && st[x.k].state === 'err') || !!stageErr || !!sellErr || balanceBad());
  }
  function renderProblems() { renderSrcTable(); renderBalance(); renderIssues(); }

  function renderAll() {
    renderSources(); renderHealth(); renderBanners(); renderProblems(); fitPanes(); renderStageViews();
  }
  // Each tab's pane fills the window below where it starts (banners and tiles above it vary), so its own scroll bars
  // stay on screen and the page itself needs no second scroll to reach them.
  function fitPanes() {
    for (const id of ['sellPane', 'invPane', 'pullPane', 'issPane']) {
      const p = $(id);
      if (!p || !p.getClientRects().length) continue;   // on a hidden tab: sized when it's shown
      const top = p.getBoundingClientRect().top + scrollY;
      p.style.height = Math.max(id === 'issPane' ? 320 : 360, Math.floor(innerHeight - top - 14)) + 'px';
    }
  }

  const TABS = ['sell', 'inv', 'pull', 'problems'];
  function showTab(t) {
    if (!TABS.includes(t)) t = 'sell';
    if (tab === 'problems' && t !== 'problems') commitSeen();
    tab = t;
    for (const b of document.querySelectorAll('#tabs .tab')) {
      b.classList.toggle('on', b.dataset.tab === t);
      b.setAttribute('aria-selected', String(b.dataset.tab === t));
    }
    for (const x of TABS) $('tab-' + x).hidden = x !== t;
    // One set of filters for Available to sell, Where it is and Pull from China: the bar moves to the tab shown. Stage and
    // PO type are about where goods are (not on Available to sell); "only styles with something available" is its own.
    const fbar = $('fbar');
    if (t === 'sell') $('tab-sell').insertBefore(fbar, $('sellPane'));
    else if (t === 'inv') $('tab-inv').insertBefore(fbar, $('invPane'));
    else if (t === 'pull') $('tab-pull').insertBefore(fbar, $('pullPane'));
    fbar.hidden = t === 'problems';
    $('fStage').hidden = $('fType').hidden = t === 'sell';
    $('fAvailWrap').hidden = t !== 'sell';
    $('fq').placeholder = t === 'sell' ? 'Search style, UPC, description…' : 'Search style, UPC, description, PO…';
    $('fq').setAttribute('aria-label', t === 'sell' ? 'Search style, UPC or description' : 'Search style, UPC, description or PO');
    $('btnClearF').hidden = !filtersOn();
    // The loading picture and the sources box go with Where it is when it's shown, else with Available to sell.
    const lead = t === 'inv' ? $('invState') : $('sellState');
    if ($('invLoad').parentNode !== lead) lead.appendChild($('invLoad'));
    const home = t === 'inv' ? $('tab-inv') : $('tab-sell');
    if ($('srcBox').parentNode !== home) home.appendChild($('srcBox'));
    fitPanes();
    if (t === 'sell' && vtSell) vtSell.show();
    if (t === 'inv' && vtInv) vtInv.show();
    if (t === 'pull' && vtPull) vtPull.show();
    if (vtInv) setInfo();
    store.set('tab2', t);
  }

  // ---------- step 2: the stage numbers (core/stages.js) ----------
  // Every unit on Master Sheet Live ends in one allocation with one bucket. The page only sums allocations: the rows,
  // the tiles and the side panel are all sums of the same allocations, after the same filters.
  // pullFromChina: China Pickup POs and the direct-to-customer sections. Those goods are sold (Saul, Oct 30), so they're
  // "Sold, ships from China", never available.
  const BUCKET_LABEL = { noCrd: 'No CRD', inProd: 'In production', ready: 'Ready', booked: 'Booked', loaded: 'Loaded', inChina: 'In China',
    water: 'On the water', port: 'At port', incoming: 'Incoming', landed: 'Landed', pullFromChina: 'Sold, ships from China', short: 'Short-shipped',
    shippedNotOnCt: 'Shipped, not on Container Tracking', duplicate: 'Repeat of a shipped line', ctMore: 'Container Tracking holds more',
    unmatched: 'Unmatched', planning: 'Planned, not booked yet', ignored: 'Ignored section', belowTracker: 'Older than the PO Status Tracker',
    casesIncoming: 'Incoming' };
  // The outcomes outside every total, in running text (proper names keep their capitals).
  const OUT_WORD = { landed: 'landed', short: 'short-shipped', shippedNotOnCt: 'shipped, not on Container Tracking', duplicate: 'repeats of a shipped line', ctMore: 'beyond Master Sheet Live',
    planning: 'planned, not booked yet', ignored: 'in an ignored section', belowTracker: 'older than the PO Status Tracker' };
  const IN_CHINA = ['noCrd', 'inProd', 'ready', 'booked', 'loaded'];
  const INCOMING = [...IN_CHINA, 'water', 'port'];
  // In the balance and the side panel, never in a total. The last three are Container Tracking units only, like ctMore.
  const OUTSIDE = ['landed', 'short', 'shippedNotOnCt', 'duplicate', 'ctMore', 'planning', 'ignored', 'belowTracker'];
  const CT_ONLY = ['ctMore', 'unmatched', 'planning', 'ignored', 'belowTracker'];
  // config.js's direct-to-customer sections, named where Pull from China is explained.
  const PULL_SECTIONS = (CFG.sections && Array.isArray(CFG.sections.pullFromChina) ? CFG.sections.pullFromChina : []).filter(x => typeof x === 'string' && x.trim());
  const PULL_TIP = `Already sold: China Pickup POs${PULL_SECTIONS.length ? `, and Container Tracking's ${PULL_SECTIONS.join(', ')} sections (straight to the customer)` : ''}, shipped to the customer from China. Never available to sell, never in In China or Incoming`;
  const TOT = ['noCrd', 'inProd', 'ready', 'booked', 'loaded', 'inChina', 'water', 'port', 'incoming', 'pullFromChina', 'unmatched'];
  const NUM_COLS = ['noCrd', 'inProd', 'ready', 'booked', 'loaded', 'inChina', 'water', 'port', 'incoming', 'pullFromChina'];
  const bucketsOf = k => (k === 'inChina' ? IN_CHINA : k === 'incoming' || k === 'casesIncoming' ? INCOMING : [k]);
  // The totals an allocation adds to: its bucket, plus In China and Incoming for the stages they are sums of.
  const totalsOf = b => (IN_CHINA.includes(b) ? [b, 'inChina', 'incoming'] : b === 'water' || b === 'port' ? [b, 'incoming']
    : b === 'pullFromChina' || b === 'unmatched' ? [b] : []);
  const isLandedStage = v => /^landed$/i.test(String(v || '').trim());
  // Saul: once Container Tracking-only goods land they show nowhere (they come back through the ATS report).
  const hiddenAlloc = a => (a.bucket === 'unmatched' || a.bucket === 'belowTracker' || a.bucket === 'ignored') && isLandedStage(a.stage);
  const units = v => (typeof v === 'number' && Number.isFinite(v) ? v : 0);
  const nz = n => (n ? fmtN(n) : '');
  const fmtD = d => (typeof d === 'number' ? window.InvCommon.fmtDay(d) : d ? String(d) : '');
  const fmtList = v => (Array.isArray(v) ? v.filter(x => x != null && x !== '').join(', ') : v == null ? '' : String(v));
  const stageText = v => BUCKET_LABEL[v] || String(v || '');
  // A stage with the Container Tracking section it sits under, when that section has no round ("Loaded · Direct A").
  const stageSec = a => stageText(a.stage) + (a.section ? ' · ' + a.section : '');
  const flagText = f => (Array.isArray(f) ? f : f ? [f] : []).map(codeLabel).join(' · ');
  const lc = v => String(v ?? '').toLowerCase();
  function ptypeOf(x) {
    const t = String((x && x.ptype) || '');
    if (/china/i.test(t)) return 'China Pickup';
    if (/unknown/i.test(t) || (x && x.tracked === false) || (Array.isArray(x && x.flags) && x.flags.includes('unknown-type'))) return 'Unknown';
    return t;
  }
  function casesText(c, np, u, loose) {
    if (!u) return '';
    const p = [];
    if (c) p.push(plural(c, 'case'));
    if (loose) p.push(`${fmtN(loose)} loose unit${loose === 1 ? '' : 's'}`);
    if (np) p.push(`${fmtN(np)} units with no case pack`);
    return p.join(' + ');
  }

  let stages = null;     // the last InvStages.build result, or { ok:false, block:[…] }; null before the first build
  let stageErr = '';     // the engine stopped or gave an answer the page can't read: a bug, said so
  let idx = null;        // per build: allocations by style, the filter lists
  let pageIssues = [];   // the page's own check that its sums equal the engine's
  // Available to sell (core/sell.js), built on the stages and the ATS report.
  const SELL_SRC = 'Available to sell';
  let sell = null;       // the last InvSell.buildSell result; null while the stages aren't loaded
  let sellErr = '';      // the selling engine stopped or gave an answer the page can't read: a bug, said so
  let sx = null;         // per build: the selling rows by style, lots by allocation
  let sellPageIssues = [];
  const got = k => (st[k] && st[k].data) || null;

  // The PO Book group and Container Tracking, each read and used, or why not (one line each).
  function stageBlockers() {
    const out = [];
    for (const k of CORE_KEYS) {
      const s = st[k];
      const used = s && !s.blocked && (s.state === 'ok' || s.state === 'warn') && ((k !== 'msl' && k !== 'ct') || s.data);
      if (used) continue;
      out.push({ k, name: BY_K[k].name, why: !s || (s.state === 'load' && !s.data) ? 'not read yet' : s.state === 'load' ? 'loading' : statusWord(s) });
    }
    return out;
  }
  // An error in the page's own stage code shows as the red "couldn't be worked out" banner, never as a load failure.
  const stageCodeErr = e => { console.error(e); stageErr = `the page couldn't draw them (${(e && e.message) || e})`; stages = { ok: false, notLoaded: [], block: [] }; idx = null; sell = null; sx = null; };
  function buildStages(today) {
    try { buildStagesNow(today); } catch (e) { stageCodeErr(e); }
    buildSell(today);
  }
  // Available to sell: only on loaded stages (the not-loaded rule: no partial numbers). The ATS report is optional to
  // the engine: without it (missing or red) customer orders aren't taken off and the page says why.
  function buildSell(today) {
    sell = null; sellErr = ''; sx = null; sellPageIssues = [];
    if (!stages || !stages.ok || stageErr) return;
    let r;
    try { r = window.InvSell.buildSell({ stages, ats: got('ats'), today, upc: got('upc'), reportDay: numOr(st.ats && st.ats.reportDay) }); }
    catch (e) { console.error(e); sellErr = `the selling engine stopped (${(e && e.message) || e})`; return; }
    if (!r || r.ok === false) { sellErr = `the selling engine gave no numbers${r && r.error ? ` (${r.error})` : ''}`; return; }
    if (!Array.isArray(r.styles) || !Array.isArray(r.lots) || !r.totals) { sellErr = "the selling engine gave an answer the page can't read"; return; }
    sell = r;
    try { indexSell(); } catch (e) { console.error(e); sell = null; sx = null; sellErr = `the page couldn't draw it (${(e && e.message) || e})`; }
  }
  function buildStagesNow(today) {
    stageErr = ''; pageIssues = [];
    const block = stageBlockers();
    if (block.length) { stages = { ok: false, notLoaded: block.map(b => b.name), block }; indexStages(); return; }
    const po = got('msl');
    let r;
    try {
      r = window.InvStages.build({ po, mrows: po.mrows, ct: got('ct'), forwarder: got('forwarder'), drayage: got('drayage'), upc: got('upc'), today, sections: CFG.sections });
    } catch (e) {
      console.error(e);
      stageErr = `the stage engine stopped (${(e && e.message) || e})`; stages = { ok: false, notLoaded: [], block: [] }; indexStages(); return;
    }
    if (r && r.ok === false) {
      const names = Array.isArray(r.notLoaded) ? r.notLoaded : [];
      stages = { ok: false, notLoaded: names, block: names.map(n => ({ name: n, why: 'not loaded' })) };
      if (!names.length) stageErr = "the stage engine said the numbers aren't loaded without saying which source";
    } else if (!r || !Array.isArray(r.styles) || !Array.isArray(r.allocs) || !Array.isArray(r.pull) || !r.totals || !r.totals.units || !r.balance) {
      stageErr = "the stage engine gave an answer the page can't read"; stages = { ok: false, notLoaded: [], block: [] };
    } else stages = r;
    indexStages();
  }

  function indexStages() {
    idx = null;
    if (!stages || !stages.ok) return;
    const byStyle = new Map();
    for (const s of stages.styles) if (s && !byStyle.has(s.style)) byStyle.set(s.style, { s, allocs: [] });
    for (const a of stages.allocs) {
      if (!a || hiddenAlloc(a)) continue;
      let e = byStyle.get(a.style);
      if (!e) { e = { s: { style: a.style, styleRaw: a.style, desc: '', category: '', sub: '', factories: [], factory: '', upcs: [] }, allocs: [] }; byStyle.set(a.style, e); }
      e.allocs.push(a);
    }
    const facs = new Set(), cats = new Set();
    for (const { s, allocs } of byStyle.values()) { cats.add(s.category || ''); for (const a of allocs) facs.add(a.factory || ''); }
    for (const r of stages.pull) facs.add(r.factory || '');
    const byId = new Map(), byLine = new Map();
    for (const a of stages.allocs) {
      if (a && a.id != null) byId.set(a.id, a);
      if (a && a.lineKey && a.status === 'open' && !hiddenAlloc(a)) (byLine.get(a.lineKey) || byLine.set(a.lineKey, []).get(a.lineKey)).push(a);
    }
    idx = { byStyle, byId, byLine, list: [...byStyle.values()], facs, cats };
    // The page's sums of every allocation must equal the engine's totals, unit for unit.
    const all = agg0();
    for (const e of idx.list) for (const a of e.allocs) addTo(all, a);
    const off = TOT.filter(k => units(stages.totals.units[k]) !== all.units[k]);
    if (off.length) {
      const msg = `The page adds the allocations up to different totals than the stage engine: ${off.map(k => `${BUCKET_LABEL[k]} ${fmtN(all.units[k])} here, ${fmtN(units(stages.totals.units[k]))} in the engine`).join('; ')}. This is a bug in Shalam Tool: tell Saul.`;
      console.error(msg);
      pageIssues = [{ source: 'Incoming', code: 'page-sum', sheet: '', row: '', po: '', style: '', message: msg }];
    }
  }
  const catOf = style => { const e = idx && idx.byStyle.get(window.InvCommon.normStyle(style)); return (e && e.s.category) || ''; };

  function agg0() { const z = () => Object.fromEntries(TOT.map(k => [k, 0])); return { units: z(), cases: z(), noPack: z() }; }
  function addTo(g, a) {
    const u = units(a.units);
    for (const k of totalsOf(a.bucket)) {
      g.units[k] += u;
      if (typeof a.cases === 'number') g.cases[k] += a.cases; else g.noPack[k] += u;
    }
  }

  // ---------- filters: one set for the Available to sell, Where it is and Pull from China tabs ----------
  // avail: Available to sell only ("only styles with something available"); stage and type: the other two only.
  const F = { q: '', fac: '', cat: '', stage: '', type: '', avail: false };
  const NONE = '__none__';   // the choice for a blank factory or category
  const fOn = () => !!(F.q || F.fac || F.cat || F.stage || F.type);
  const fOnSell = () => !!(F.q || F.fac || F.cat || F.avail);
  const filtersOn = () => fOn() || F.avail;   // any filter at all, a hidden one too: Clear shows and clears it
  const facOk = v => !F.fac || (F.fac === NONE ? !v : v === F.fac);
  const catOk = v => !F.cat || (F.cat === NONE ? !v : v === F.cat);
  const styleHit = s => lc(s.style).includes(F.q) || lc(s.styleRaw).includes(F.q) || lc(s.desc).includes(F.q);
  const allocHit = a => lc(a.upc).includes(F.q) || lc(a.po).includes(F.q) || lc(a.poShown).includes(F.q);
  // sellView: opened from Available to sell, where the stage and PO type filters are hidden (and not applied).
  const passes = (a, sHit, sellView) => facOk(a.factory || '') && (sellView || !F.type || ptypeOf(a) === F.type) &&
    (sellView || !F.stage || bucketsOf(F.stage).includes(a.bucket)) && (sHit || !F.q || allocHit(a));
  const STAGE_OPTS = [['inChina', 'In China (any)'], ...IN_CHINA.map(k => [k, '· ' + BUCKET_LABEL[k]]), ['water', 'On the water'], ['port', 'At port'],
    ['incoming', 'Incoming (any)'], ['pullFromChina', BUCKET_LABEL.pullFromChina], ['unmatched', 'Unmatched']];
  const STAGE_OUT = OUTSIDE.map(k => [k, BUCKET_LABEL[k]]);
  const TYPE_OPTS = [['LA', 'LA'], ['China Pickup', 'China Pickup'], ['Unknown', 'Unknown']];
  function fillSelect(el, all, opts, cur, extra) {
    const o = list => list.map(([v, t]) => `<option value="${esc(v)}"${v === cur ? ' selected' : ''}>${esc(t)}</option>`).join('');
    let h = `<option value="">${esc(all)}</option>${o(opts)}${extra || ''}`;
    if (cur && !opts.some(x => x[0] === cur) && !(extra && extra.includes(`value="${esc(cur)}"`))) h += `<option value="${esc(cur)}" selected>${esc(cur === NONE ? '(blank)' : cur)} (none now)</option>`;
    el.innerHTML = h;
  }
  function renderFilterLists() {
    const named = set => [...set].filter(Boolean).sort((a, b) => a.localeCompare(b, undefined, { numeric: true }));
    // The styles only on the ATS report add their categories (from UPC_MASTER); a style with no factory is "(no factory)".
    const facs = new Set([...(idx ? idx.facs : []), ...(sx ? sx.facs : [])]), cats = new Set([...(idx ? idx.cats : []), ...(sx ? sx.cats : [])]);
    fillSelect($('fFac'), 'All factories', [...named(facs).map(v => [v, v]), ...(facs.has('') ? [[NONE, '(no factory)']] : [])], F.fac);
    fillSelect($('fCat'), 'All categories', [...named(cats).map(v => [v, v]), ...(cats.has('') ? [[NONE, '(no category)']] : [])], F.cat);
    fillSelect($('fStage'), 'All stages', STAGE_OPTS, F.stage,
      `<optgroup label="Not counted (side panel and Problems)">${STAGE_OUT.map(([v, t]) => `<option value="${v}"${v === F.stage ? ' selected' : ''}>${esc(t)}</option>`).join('')}</optgroup>`);
    fillSelect($('fType'), 'All PO types', TYPE_OPTS, F.type);
    $('fAvail').checked = F.avail;
    $('btnClearF').hidden = !filtersOn();
  }

  // ---------- tables that draw only the rows in view ----------
  // Every row has one height (measured from the first one drawn); a spacer row above and below stands in for the
  // rest, so the scroll bar, the sticky header and the frozen columns work as in the full table.
  function vtable(paneId, tableId, o) {
    const pane = $(paneId), table = $(tableId), OVER = 14;
    const v = { rows: [], rowH: o.rowH || 29, a: -1, b: -1, top: 0, raf: 0 };
    const body = () => table.tBodies[0] || table.appendChild(document.createElement('tbody'));
    function draw(force, depth) {
      const headH = (table.tHead && table.tHead.offsetHeight) || 0;
      const top = Math.max(0, pane.scrollTop - headH), h = pane.clientHeight || 640;
      const a = Math.max(0, Math.floor(top / v.rowH) - OVER), b = Math.min(v.rows.length, Math.ceil((top + h) / v.rowH) + OVER);
      if (!force && a === v.a && b === v.b) return;
      v.a = a; v.b = b;
      const cols = o.cols();
      if (!v.rows.length) { body().innerHTML = `<tr><td colspan="${cols}" class="empty">${o.empty()}</td></tr>`; return; }
      const spc = px => `<tr class="spc" aria-hidden="true"><td colspan="${cols}" style="height:${px}px"></td></tr>`;
      let html = a ? spc(a * v.rowH) : '';
      for (let i = a; i < b; i++) html += o.row(v.rows[i], i);
      if (b < v.rows.length) html += spc((v.rows.length - b) * v.rowH);
      body().innerHTML = html;
      // Measured on a main row (a UPC line under a style is drawn to the same height by app.css, but a style row is the reference).
      const tr = body().querySelector('tr.r:not(.upc)') || body().querySelector('tr.r');
      const rh = tr ? tr.getBoundingClientRect().height : 0;
      if (rh > 0 && Math.abs(rh - v.rowH) > 0.5 && (depth || 0) < 2) { v.rowH = rh; draw(true, (depth || 0) + 1); }
    }
    pane.addEventListener('scroll', () => {
      v.top = pane.scrollTop;
      if (!v.raf) v.raf = requestAnimationFrame(() => { v.raf = 0; draw(false); });
    }, { passive: true });
    return {
      get rows() { return v.rows; },
      // New rows. toTop: a filter or sort change; a refresh keeps the place.
      set(rows, toTop) {
        v.rows = rows || [];
        if (table.tHead) table.tHead.innerHTML = o.head(); else table.createTHead().innerHTML = o.head();
        if (toTop) { v.top = 0; pane.scrollTop = 0; }
        draw(true);
      },
      // A tab shown again: back to where it was (a hidden pane can lose its place), then the rows for it.
      show() { if (Math.abs(pane.scrollTop - v.top) > 1) pane.scrollTop = v.top; draw(true); },
      redraw() { draw(true); },
    };
  }
  let vtSell = null, vtInv = null, vtPull = null, vtPanel = null;
  const sortHead = (cols, srt) => `<tr>${cols.map(c => {
    const on = srt && srt.k === c.k, cls = [c.cls, c.n || c.b || c.off ? 'num' : '', c.sum ? 'sum' : '', c.off ? 'off' : ''].filter(Boolean).join(' ');
    const sortable = !c.off && !c.nosort;
    return `<th class="${cls}"${sortable ? ` data-s="${c.k}" aria-sort="${on ? (srt.dir > 0 ? 'ascending' : 'descending') : 'none'}"` : ''}${c.tip ? ` title="${esc(c.tip)}"` : ''}>${esc(c.t)}${on ? `<span class="ar">${srt.dir > 0 ? '▲' : '▼'}</span>` : ''}</th>`;
  }).join('')}</tr>`;
  // Blanks always last; numbers big-first on the first click, text A to Z.
  function cmpBy(val, dir) {
    return (x, y) => {
      const a = val(x), b = val(y), ea = a == null || a === '' || a === 0, eb = b == null || b === '' || b === 0;
      if (ea || eb) return ea === eb ? 0 : ea ? 1 : -1;
      return (typeof a === 'number' && typeof b === 'number' ? a - b : String(a).localeCompare(String(b), undefined, { numeric: true })) * dir;
    };
  }
  function nextSort(srt, cols, k) {
    const c = cols.find(x => x.k === k);
    if (!c) return srt;
    if (srt && srt.k === k) return { k, dir: -srt.dir };
    return { k, dir: c.n || c.b || c.date ? (c.date ? 1 : -1) : 1 };
  }

  // ---------- Where it is tab (step 2's Inventory tab) ----------
  const INV_COLS = [
    { k: 'style', t: 'Style', cls: 'fz f-sty' }, { k: 'desc', t: 'Description' }, { k: 'category', t: 'Category' }, { k: 'factory', t: 'Factory' },
    { k: 'onHand', t: 'On hand', n: 1, tip: 'In warehouse on the ATS report: physical stock, before customer orders (Available to sell takes them off)' },
    ...NUM_COLS.map(k => ({ k, t: BUCKET_LABEL[k], b: 1, sum: k === 'inChina' || k === 'incoming',
      tip: k === 'inChina' ? 'No CRD + In production + Ready + Booked + Loaded' : k === 'incoming' ? 'In China + On the water + At port'
        : k === 'pullFromChina' ? PULL_TIP : '' })),
    { k: 'casesIncoming', t: 'Cases incoming', b: 1, tip: 'Incoming in cases: units ÷ the Master column on Master Sheet Live, rounded up' },
    { k: 'nextEta', t: 'Next ETA', date: 1, tip: 'The earliest ETA of the units on the water or at port' },
  ];
  let sortInv = { k: 'style', dir: 1 };
  let invShown = { styles: 0, empty: 0 };

  // One style's numbers from its allocations (already filtered). Tags: units outside the columns worth a look.
  function sumRow(allocs) {
    const g = agg0(), tags = {};
    let eta = null;
    for (const a of allocs) {
      addTo(g, a);
      if ((a.bucket === 'water' || a.bucket === 'port') && typeof a.eta === 'number' && (eta == null || a.eta < eta)) eta = a.eta;
      if (OUTSIDE.includes(a.bucket) && F.stage === a.bucket) tags[a.bucket] = (tags[a.bucket] || 0) + units(a.units);
    }
    if (g.units.unmatched) tags.unmatched = g.units.unmatched;
    const any = NUM_COLS.some(k => g.units[k]) || Object.values(tags).some(Boolean);
    return { units: g.units, cases: g.cases, noPack: g.noPack, nextEta: eta, tags, any };
  }
  function invRows() {
    const rows = [];
    let empty = 0;
    for (const e of idx.list) {
      const s = e.s;
      if (!catOk(s.category || '')) continue;
      const sHit = !F.q || styleHit(s);
      const allocs = e.allocs.filter(a => passes(a, sHit));
      if (!allocs.length) { if (sHit && !F.fac && !F.type && !F.stage) empty++; continue; }
      const sum = sumRow(allocs);
      if (!sum.any) { empty++; continue; }
      const factories = Array.isArray(s.factories) ? s.factories : [];
      rows.push({ kind: 'style', key: s.style, s, styleRaw: s.styleRaw || s.style, desc: s.desc || '', category: s.category || '', sub: s.sub || '',
        factory: F.fac && F.fac !== NONE ? F.fac : s.factory || factories[0] || '', factories, allocs, ...sum });
    }
    const k = sortInv.k;
    const val = k === 'style' ? r => r.styleRaw : ['desc', 'category', 'factory', 'nextEta'].includes(k) ? r => r[k]
      : k === 'casesIncoming' ? r => r.cases.incoming : k === 'onHand' ? r => inWhOf(r.key) : r => r.units[k];
    rows.sort(cmpBy(val, sortInv.dir));
    invShown = { styles: rows.length, empty };
    // UPC lines under a style with more than one UPC (with the filters on, only the UPCs they leave).
    const out = [];
    for (const r of rows) {
      out.push(r);
      const by = new Map();
      if (!fOn()) for (const u of r.s.upcs || []) by.set((u && (u.upc || u.upcRaw)) || '', []);
      for (const a of r.allocs) if (totalsOf(a.bucket).length || r.tags[a.bucket]) { const k = a.upc || ''; if (!by.has(k)) by.set(k, []); by.get(k).push(a); }
      r.upc1 = by.size === 1 ? [...by.keys()][0] : '';
      if (by.size < 2) continue;
      const subs = [...by].sort((x, y) => (x[0] ? 0 : 1) - (y[0] ? 0 : 1) || x[0].localeCompare(y[0]));
      for (const [upc, list] of subs) out.push({ kind: 'upc', key: r.key, upc, styleRaw: r.styleRaw, allocs: list, ...sumRow(list) });
    }
    return out;
  }
  function invRowHtml(r, i) {
    const nb = k => (r.units[k] ? `<button type="button" class="nb" data-b="${k}">${fmtN(r.units[k])}</button>` : '');
    const up = r.kind === 'upc';
    let h = `<tr class="r${up ? ' upc' : ''}" data-i="${i}">`;
    if (up) h += `<td class="fz f-sty" title="UPC ${esc(r.upc || 'none')} of ${esc(r.styleRaw)}">${esc(r.upc || 'no UPC')}</td><td></td><td></td><td></td>`;
    else {
      const tagWords = (b, n) => (b === 'unmatched' ? `${fmtN(n)} units unmatched: on Container Tracking with no Master Sheet Live line, in the Unmatched tile, never in the stage columns`
        : `${fmtN(n)} units ${OUT_WORD[b] || stageText(b).toLowerCase()}: not counted in any column`);
      const tags = Object.entries(r.tags).map(([b, n]) => `<button type="button" class="tag${b === 'unmatched' ? ' um' : ''}" data-b="${b}" title="${esc(tagWords(b, n))}">${fmtN(n)} ${esc(b === 'unmatched' ? 'unmatched' : OUT_WORD[b] || stageText(b).toLowerCase())}</button>`).join('');
      const tip = [r.styleRaw, r.upc1 ? 'UPC ' + r.upc1 : '', ...Object.entries(r.tags).map(([b, n]) => tagWords(b, n)), 'Click for every line behind this style'].filter(Boolean).join('\n');
      // The UPC gives way first, so a tag always shows whatever the UPC's length.
      h += `<td class="fz f-sty" title="${esc(tip)}"><div class="sc"><button type="button" class="sb" data-b="all">${esc(r.styleRaw)}</button>${r.upc1 ? `<span class="u">${esc(r.upc1)}</span>` : ''}${tags}</div></td>` +
        `<td class="desc" title="${esc(r.desc)}">${esc(r.desc)}</td><td class="cut" title="${esc([r.category, r.sub].filter(Boolean).join(' · '))}">${esc(r.category)}</td>` +
        `<td class="cut" title="${esc(r.factories.join(', '))}">${esc(r.factory)}</td>`;
    }
    h += up ? '<td class="num"></td>' : onHandCell(r.key);
    for (const k of NUM_COLS) h += `<td class="num${k === 'inChina' || k === 'incoming' ? ' sum' : ''}">${nb(k)}</td>`;
    const c = r.cases.incoming, np = r.noPack.incoming;
    const ci = !r.units.incoming ? '' : `<button type="button" class="nb" data-b="casesIncoming" title="${esc(casesText(c, np, r.units.incoming))}">${c ? fmtN(c) : ''}${np ? `<span class="np">${c ? '+ no case pack' : 'no case pack'}</span>` : ''}</button>`;
    return h + `<td class="num">${ci}</td><td class="nw">${esc(fmtD(r.nextEta))}</td></tr>`;
  }

  // ---------- the tiles ----------
  const TILES = [['inChina', 'In China', 'No CRD + In production + Ready + Booked + Loaded'], ['water', 'On the water', 'Sailed, not arrived'],
    ['port', 'At port', 'Arrived or picked up, not delivered'], ['incoming', 'Incoming', 'In China + On the water + At port'],
    ['pullFromChina', BUCKET_LABEL.pullFromChina, PULL_TIP],
    ['unmatched', 'Unmatched', 'On Container Tracking with no Master Sheet Live line (not landed): never in the stage numbers']];
  function renderTiles(g) {
    const lead = stageLead();
    const live = TILES.map(([k, t, tip]) => {
      if (!g) return `<div class="glass tile nl" data-t="${k}"><div class="lbl">${esc(t)}</div><div class="n">${esc(lead.word)}</div><div class="cs"></div></div>`;
      return `<button type="button" class="glass tile${panel && panel.kind === 'tile' && panel.b === k ? ' on' : ''}" data-t="${k}" title="${esc(tip + '. Click for the lines behind it.')}" aria-label="${esc(`${t}: ${fmtN(g.units[k])} units`)}">` +
        `<div class="lbl">${esc(t)}${fOn() ? '<i>filtered</i>' : ''}</div><div class="n">${fmtN(g.units[k])}<small>units</small></div>` +
        `<div class="cs">${esc(casesText(g.cases[k], g.noPack[k], g.units[k]))}</div></button>`;
    });
    // Stock after customer orders, and what can be sold from China, are on their own tab.
    const later = '<button type="button" class="glass tile off go" data-go="sell" title="Open Available to sell"><div class="lbl">On hand and available</div><div class="n">Available to sell →</div><div class="cs">stock after customer orders, and goods from China</div></button>';
    $('tiles').innerHTML = live.join('') + later;
  }

  // What the Inventory and Pull from China tabs say while there are no numbers: loading, not loaded, or a bug.
  // short: the line in an empty table.
  function stageLead() {
    if (stages && stages.ok) return null;
    if (scopeFatal || fatal) return { word: 'not loaded', h: 'Stage numbers not loaded', p: 'Nothing could be read: see the red banner.', short: 'Not loaded: nothing could be read (see the red banner).' };
    if (stageErr) return { word: 'not loaded', h: "The stage numbers couldn't be worked out", p: `${cap(stageErr)}. This is a bug in Shalam Tool, not a problem in the files: tell Saul.`, short: "Not loaded: the stage numbers couldn't be worked out (see the red banner)." };
    if (!stages) {
      const busy = !!loading || !cycleDone;
      return { word: busy ? 'loading…' : 'not loaded', h: busy ? 'Loading the sources…' : 'Stage numbers not loaded',
        p: busy ? 'Master Sheet Live, the PO Status Tracker and Container Tracking first; the numbers appear as soon as they are in.' : '', short: busy ? 'Loading…' : 'Not loaded.' };
    }
    const items = (stages.block || []).map(b => `${b.name} (${b.why})`);
    // A red or blank source has a banner saying what to fix; one that wasn't read (Cancel) only needs a Refresh.
    const banner = (stages.block || []).some(b => { const s = b.k && st[b.k]; return !!s && (s.state === 'err' || s.blocked || s.label === 'not set'); });
    return { word: 'not loaded', h: 'Stage numbers not loaded',
      p: `Nothing is shown until the PO Book group (Master Sheet Live, PO Status Tracker, PO Book lines) and Container Tracking are all read.${items.length ? ` Not read now: ${items.join(', ')}.` : ''} ${banner ? 'The banner above says what to fix.' : 'Click Refresh to read them.'}`,
      short: `Not loaded${items.length ? ': ' + items.join(', ') : ''}.` };
  }
  function renderLead() {
    const L = stageLead();
    $('invState').hidden = !L;
    $('pullState').hidden = !L;
    if (!L) return;
    $('invStateH').textContent = L.h; $('invStateP').textContent = L.p;
    $('pullStateH').textContent = L.h === 'Loading the sources…' ? L.h : 'Not loaded'; $('pullStateP').textContent = L.p;
  }
  const emptyText = what => { const L = stageLead(); return esc(L ? L.short : what); };

  function renderInv(toTop) {
    let g = null, rows = [];
    if (idx) {
      rows = invRows();
      g = agg0();
      for (const r of rows) if (r.kind === 'style') for (const a of r.allocs) addTo(g, a);
    }
    renderTiles(g);
    vtInv.set(rows, toTop);
    setInfo();
  }
  // The line at the end of the filter bar, for the tab it sits on.
  function setInfo() {
    $('fInfo').textContent = tab === 'sell' ? (sx ? sellInfo() : '') : !idx ? '' : tab === 'pull' ? pullInfo
      : `${plural(invShown.styles, 'style')}${invShown.empty ? ` · ${fmtN(invShown.empty)} more ${fOn() ? (invShown.empty === 1 ? 'matches but has' : 'match but have') : (invShown.empty === 1 ? 'has' : 'have')} nothing in these columns (shipped, landed or not counted)` : ''}`;
  }

  // ---------- Pull from China tab ----------
  const PULL_COLS = [
    { k: 'po', t: 'PO', cls: 'fz f-po' }, { k: 'style', t: 'Style', cls: 'fz f-sty2' }, { k: 'desc', t: 'Description' }, { k: 'customer', t: 'Customer' },
    { k: 'ptype', t: 'PO type' }, { k: 'factory', t: 'Factory' }, { k: 'crd', t: 'CRD', date: 1 }, { k: 'units', t: 'Units', n: 1 }, { k: 'cases', t: 'Cases', n: 1 },
    { k: 'ctStatus', t: 'Container Tracking status' }, { k: 'round', t: 'Round' }, { k: 'containers', t: 'Container' },
    { k: 'etd', t: 'ETD', date: 1 }, { k: 'eta', t: 'ETA', date: 1 },
    { k: 'pullFrom', t: 'Pull-from date', date: 1, tip: "When a buyer can have it: in China, 10 days before the ETD, else 5 days before the CRD, else about 75 days after the order date; sailed, its ETA. \"Sold\": China Pickup or straight to the customer" },
    { k: 'availAfter', t: 'Available after orders', n: 1, tip: 'Units left to sell after customer orders beyond the stock are taken off, earliest goods first' },
    { k: 'flags', t: 'Flags' },
  ];
  let sortPull = null, pullInfo = '';
  // The stage filter on a Pull from China line: by its Container Tracking status, and its CRD while it's not on it.
  function pullStageOk(r) {
    // By the status of each part (a section name such as "Direct Loaded" is not a stage); "Planned" is at its CRD stage.
    const st = Array.isArray(r.ctParts) ? r.ctParts.map(p => String(p.status || '')) : String(r.ctStatus || '').split(' · ').map(x => x.replace(/\s[\d,]+$/, ''));
    const is = w => st.includes(w), not = is('Not on tracker') || is('Planned'), today = stages.today;
    const flags = Array.isArray(r.flags) ? r.flags : [];
    const has = { booked: is('Booked'), loaded: is('Loaded'), water: is('On the water'), port: is('At port'), landed: is('Landed'),
      noCrd: not && r.crd == null, inProd: not && typeof r.crd === 'number' && r.crd >= today, ready: not && typeof r.crd === 'number' && r.crd < today,
      pullFromChina: ptypeOf(r) === 'China Pickup' || flags.includes('pull-section'), ctMore: flags.includes('ct-more'),
      planning: flags.includes('planning-section') };
    return bucketsOf(F.stage).some(b => has[b]);
  }
  function pullRows() {
    const rows = stages.pull.filter(r => r && facOk(r.factory || '') && (!F.type || ptypeOf(r) === F.type) && (!F.cat || catOk(catOf(r.style))) &&
      (!F.q || [r.style, r.styleRaw, r.desc, r.upc, r.po, r.poShown].some(v => lc(v).includes(F.q))) && (!F.stage || pullStageOk(r)));
    if (sortPull) {
      const k = sortPull.k;
      rows.sort(cmpBy(r => (k === 'flags' ? flagText(r.flags) : k === 'containers' || k === 'round' ? fmtList(r[k]) : k === 'ptype' ? ptypeOf(r)
        : k === 'style' ? r.styleRaw || r.style : k === 'pullFrom' ? pullSellVal(r, 'day') : k === 'availAfter' ? pullSellVal(r, 'avail') : r[k]), sortPull.dir));
    }
    const u = rows.reduce((p, r) => p + units(r.units), 0), cp = rows.filter(r => ptypeOf(r) === 'China Pickup').reduce((p, r) => p + units(r.units), 0);
    pullInfo = `${plural(rows.length, 'open line')} · ${fmtN(u)} units${cp ? ` (China Pickup ${fmtN(cp)})` : ''}`;
    return rows;
  }
  function pullRowHtml(r, i) {
    const st2 = String(r.ctStatus || ''), sty = r.styleRaw || r.style;
    const poTip = r.poShown ? `PO ${r.po} on Master Sheet Live; the PO Book shows it as ${r.poShown}` : `PO ${r.po}`;
    return `<tr class="r" data-i="${i}"><td class="fz f-po" title="${esc(poTip)}">${esc(r.po)}<span class="u">${r.poShown ? esc('as ' + r.poShown) : '&nbsp;'}</span></td>` +
      `<td class="fz f-sty2" title="${esc([sty, r.upc ? 'UPC ' + r.upc : ''].filter(Boolean).join('\n'))}">${esc(sty)}<span class="u">${r.upc ? esc(r.upc) : '&nbsp;'}</span></td>` +
      `<td class="desc" title="${esc(r.desc)}">${esc(r.desc)}</td><td class="cut" title="${esc(r.customer)}">${esc(r.customer)}</td><td class="nw">${esc(ptypeOf(r))}</td>` +
      `<td class="cut" title="${esc(r.factory)}">${esc(r.factory)}</td><td class="nw">${esc(fmtD(r.crd))}</td>` +
      `<td class="num">${r.units ? `<button type="button" class="nb" data-b="line" title="Click for this line's stages and where each fact came from">${fmtN(r.units)}</button>` : ''}</td>` +
      `<td class="num">${typeof r.cases === 'number' ? nz(r.cases) : r.units ? '<span class="np" title="No case pack (Master) for this line">no case pack</span>' : ''}</td>` +
      // A mix of stages takes two lines: the first part, then the rest.
      `<td class="cut st${/^(not on tracker$|planned · )/i.test(st2) ? ' nt' : ''}" title="${esc(st2)}">${esc(st2.split(' · ')[0])}<span class="u">${st2.includes(' · ') ? esc(st2.split(' · ').slice(1).join(' · ')) : '&nbsp;'}</span></td><td class="nw">${esc(fmtList(r.round))}</td>` +
      `<td class="cut" title="${esc(fmtList(r.containers))}">${esc(fmtList(r.containers))}</td><td class="nw">${esc(fmtD(r.etd))}</td><td class="nw">${esc(fmtD(r.eta))}</td>` +
      pullSellCells(r) +
      `<td class="cut flag" title="${esc(flagText(r.flags))}">${esc(flagText(r.flags))}</td></tr>`;
  }
  function renderPull(toTop) {
    const rows = idx ? pullRows() : [];
    vtPull.set(rows, toTop);
    const c = $('cPull');
    c.hidden = !idx; c.textContent = idx ? fmtN(rows.length) : '';
    setInfo();
  }

  // ---------- Available to sell (core/sell.js) ----------
  // Per style: the stock left after customer orders (the weekly ATS report), and the lots in China or on the water, each
  // with one date a buyer can have it. Customer orders beyond the stock come off the earliest lots. Sold goods (China
  // Pickup POs, the direct-to-customer sections) are shown, never available. The page only draws the engine's numbers;
  // the tiles are sums of the rows shown, and with no filter they must equal the engine's totals.
  const numOr = v => (typeof v === 'number' && Number.isFinite(v) ? v : null);
  const hasFlag = (x, f) => !!x && Array.isArray(x.flags) && x.flags.includes(f);
  const upcText = u => (typeof u === 'string' ? u : (u && (u.upc || u.upcRaw)) || '');
  const sellToday = () => numOr(sell && sell.today) ?? numOr(stages && stages.today) ?? window.InvCommon.todayNY(Date.now());
  // m/d in running text ("10 days before ETD 11/14"); with the year when it isn't this year's.
  const md = d => { const t = fmtD(d); return d != null && window.InvCommon.dayYear(d) === window.InvCommon.dayYear(sellToday()) ? t.replace(/\/\d{4}$/, '') : t; };

  // What the stock columns can show: the ATS report's numbers, or why not.
  function atsMode() {
    const s = st.ats;
    if (got('ats')) return 'ok';
    if (!s || s.state === 'load') return 'wait';
    if (s.state === 'err') return 'bad';
    if (s.label === 'no report yet') return 'none';
    if (s.label === 'not set') return 'unset';
    return 'unread';
  }
  const ATS_WORD = { bad: 'not loaded', none: 'no report yet', unset: 'not set', unread: 'not read', wait: 'loading…' };
  // The engine took customer orders off (it had the report, and the page has it too).
  const stockOk = () => !!sell && !sell.atsMissing && atsMode() === 'ok';
  const stockWord = () => { const m = atsMode(); return m === 'ok' ? 'not loaded' : ATS_WORD[m] || 'not loaded'; };
  const atsDay = () => numOr(st.ats && st.ats.reportDay) ?? numOr(sell && sell.reportDay);
  const asOfText = () => { const d = atsDay(); return d == null ? '' : ` of ${fmtD(d)}`; };
  function atsWhy() {
    const m = atsMode(), s = st.ats || fresh('ats');
    if (m === 'bad') return `The ATS report couldn't be used: ${onFromColon(s.message)}`;
    if (m === 'none') return `No ATS report yet: save the weekly ATS report into ${atsPlace()}.`;
    if (m === 'unset') return "The ATS folder's location is blank in config.js.";
    if (m === 'unread') return "The ATS report wasn't read this time: click Refresh.";
    if (m === 'wait') return 'The ATS report is still loading.';
    return "The ATS report loaded, but customer orders weren't taken off (a bug in Shalam Tool: tell Saul).";
  }
  // The style's row on the ATS report (Where it is shows its In warehouse as On hand).
  function atsRowOf(style) { const a = got('ats'); return a && a.styles instanceof Map ? a.styles.get(style) || null : null; }
  const inWhOf = style => { const r = atsRowOf(style); return r ? numOr(r.inWh) : null; };
  function onHandCell(style) {
    const m = atsMode();
    if (m !== 'ok') return `<td class="num nr" title="${esc(atsWhy())}">${esc(ATS_WORD[m])}</td>`;
    const r = atsRowOf(style);
    if (!r) return `<td class="num nr" title="${esc(`This style has no row on the ATS report${asOfText()}`)}">not on report</td>`;
    const v = numOr(r.inWh);
    // A blank or a figure that isn't a number is never 0 (core/ats.js lists it).
    if (v == null) return `<td class="num nr" title="${esc(`In WH on the ATS report${asOfText()} is blank or not a number (listed on Problems)`)}">can't be read</td>`;
    return `<td class="num${v < 0 ? ' neg' : ''}" title="${esc(`In warehouse on the ATS report${asOfText()}, before customer orders`)}">${v ? fmtN(v) : ''}</td>`;
  }

  // A lot's date: its pull-from date in China, its arrival (ETA) once sailed.
  const lotDay = l => numOr(l.availableDate) ?? numOr(l.inChina ? l.pullFrom : l.arrival);
  const lotOrder = (a, b) => { const x = lotDay(a), y = lotDay(b); return (x == null) - (y == null) || (x ?? 0) - (y ?? 0) || String(a.po || '').localeCompare(String(b.po || ''), undefined, { numeric: true }); };
  const allocOf = l => (idx && l && l.allocId != null ? idx.byId.get(l.allocId) || null : null);
  const packNum = s => (typeof s.casePack === 'number' && s.casePack > 0 ? s.casePack : null);
  const lotPack = (l, s) => { const a = allocOf(l), m = numOr(l.master) ?? (a ? numOr(a.master) : null); return m && m > 0 ? m : packNum(s); };
  // When a buyer can have a lot, in words (Saul's rule): in China, 10 days before the ETD, else 5 days before the CRD
  // (ready date), else about 75 days after the order date (an estimate); a date already past is "now". Sailed: its ETA
  // (one already past says so: due, or at port).
  function lotWhen(l) {
    const today = sellToday(), d = lotDay(l), dates = l.dates && typeof l.dates === 'object' ? l.dates : {};
    if (!l.inChina) {
      const eta = md(numOr(l.eta) ?? d);
      return d == null ? { day: null, main: 'no ETA yet', why: 'sailed, with no ETA', none: true }
        : d < today ? { day: d, main: fmtD(d), why: `ETA ${eta} passed: due now`, passed: true } : { day: d, main: fmtD(d), why: `arrives: ETA ${eta}` };
    }
    const rule = String(l.pullRule || (d == null ? 'none' : '')), od = numOr(dates.orderDate);
    const why = rule === 'etd' ? `10 days before ETD ${md(numOr(l.etd))}` : rule === 'crd' ? `5 days before CRD ${md(numOr(l.readyDate))}`
      : rule === 'order' ? `estimated, 75 days after the order date${od != null ? ' ' + md(od) : ''}` : d == null ? 'no ETD, CRD or order date' : '';
    if (d == null) return { day: null, main: 'no date', why, none: true };
    const est = rule === 'order' || !!l.estimate, dateText = (est ? 'about ' : '') + fmtD(d);
    return d < today ? { day: d, main: 'now', why: `from ${dateText}${why ? ': ' + why : ''}`, est } : { day: d, main: dateText, why, est };
  }
  // The ready date (CRD) and every source's date: Container Tracking's Ready Date, the PO Book's own date, Factory CRD,
  // CRD Expected, Ship by (core/sell.js's order; the first with a date is used). Short names for the line under the date.
  const READY_SHORT = { ctReady: 'CT ready', poBook: 'PO Book', factoryCrd: 'Factory CRD', crdExpected: 'CRD Expected', shipBy: 'Ship by', crd: 'CRD' };
  const READY_KEYS = ['ctReady', 'poBook', 'factoryCrd', 'crdExpected', 'shipBy'];
  function readyInfo(l) {
    const IS = window.InvSell || {}, LBL = IS.READY_LABEL || {}, FILE = IS.READY_FILE || {};
    const d = l.dates && typeof l.dates === 'object' ? l.dates : {}, won = String(l.readyFrom || '');
    const have = (Array.isArray(l.readySources) ? l.readySources.map(x => ({ k: String(x && x.key), day: numOr(x && x.day), far: !!(x && x.far) }))
      : READY_KEYS.map(k => ({ k, day: numOr(d[k]) }))).filter(x => x.day != null);
    const disagree = hasFlag(l, 'dates-disagree');
    const long = k => `${cap(LBL[k] || READY_SHORT[k] || k)}${FILE[k] ? ` (${FILE[k]})` : ''}`;
    return {
      disagree,
      line: have.map(x => `${READY_SHORT[x.k] || x.k} ${md(x.day)}${x.k === won && have.length > 1 ? ' (used)' : x.far ? ' (wrong year?)' : ''}`).join(' · '),
      tip: ["Ready date (CRD): the first of Container Tracking's Ready Date, the PO Book's own date, Factory CRD, CRD Expected and Ship by; a date more than a year ahead or before 2015 is skipped.",
        ...have.map(x => `${long(x.k)}: ${fmtD(x.day)}${x.k === won ? ' (used)' : x.far ? ' (not used: more than a year ahead or before 2015, a wrong year?)' : ''}`), !have.length ? 'No source has a date.' : '',
        disagree ? 'These dates are more than 7 days apart: check which one is right.' : ''].filter(Boolean).join('\n'),
    };
  }
  // The first date a style's goods in China can be pulled: the engine's (a day, or { date, rule }), or, if it gives none
  // at all, the earliest of its lots still in China with units left.
  function firstPullOf(s, lots) {
    const f = s.firstPull, open = lots.filter(l => l.inChina && units(l.available) > 0);
    let day = null, rule = '', lot = null;
    if (f && typeof f === 'object') { day = numOr(f.date ?? f.day); rule = String(f.rule || ''); }
    else if (typeof f === 'number') { day = f; rule = String(s.firstPullRule || ''); }
    if (f === undefined) { lot = open.find(l => numOr(l.pullFrom) != null) || null; day = lot ? numOr(lot.pullFrom) : null; }
    else if (day != null) lot = open.find(l => numOr(l.pullFrom) === day) || null;
    if (!rule && lot) rule = String(lot.pullRule || '');
    return { day, rule, lot, est: rule === 'order' || !!s.firstPullEstimate || !!(lot && lot.estimate), open: open.length };
  }
  function firstArrivalOf(s, lots) {
    const f = s.firstArrival;
    if (typeof f === 'number') return f;
    if (f && typeof f === 'object') return numOr(f.date ?? f.day);
    if (f !== undefined) return null;
    const ds = lots.filter(l => !l.inChina && units(l.available) > 0).map(l => numOr(l.arrival) ?? numOr(l.eta)).filter(d => d != null);
    return ds.length ? Math.min(...ds) : null;
  }

  function indexSell() {
    const byAlloc = new Map(), lotsByStyle = new Map();
    for (const l of sell.lots) {
      if (!l) continue;
      if (l.allocId != null) byAlloc.set(l.allocId, l);
      (lotsByStyle.get(l.style) || lotsByStyle.set(l.style, []).get(l.style)).push(l);
    }
    const upc = got('upc'), upcMap = upc && upc.byStyle instanceof Map ? upc.byStyle : null;
    const list = [], byStyle = new Map(), cats = new Set(), facs = new Set();
    for (const s of sell.styles) {
      if (!s || byStyle.has(s.style)) continue;
      const lots = (Array.isArray(s.lots) ? s.lots.filter(Boolean) : lotsByStyle.get(s.style) || []).slice().sort(lotOrder);
      const e = idx && idx.byStyle.get(s.style), um = (upcMap && upcMap.get(s.style)) || [];
      const pick = k => { const x = um.find(y => y && y[k]); return x ? x[k] : ''; };
      const category = s.category || (e && e.s.category) || pick('category');
      const factories = [...new Set([...(Array.isArray(s.factories) ? s.factories : []), ...(e && Array.isArray(e.s.factories) ? e.s.factories : []),
        ...lots.map(l => { const a = allocOf(l); return l.factory || (a ? a.factory : ''); })].filter(Boolean))].sort();
      const upcs = [...new Set([...(Array.isArray(s.upcs) ? s.upcs : []).map(upcText), ...lots.map(l => l.upc || '')].filter(Boolean))];
      const desc = s.desc || (e && e.s.desc) || pick('desc');
      const r = { s, key: s.style, styleRaw: s.styleRaw || s.style, desc, category, factories, upcs, lots,
        fp: firstPullOf(s, lots), fa: firstArrivalOf(s, lots) };
      r.hay = [r.key, r.styleRaw, desc, ...upcs].map(lc).join('\n');
      list.push(r); byStyle.set(s.style, r);
      cats.add(category || '');
      if (factories.length) for (const f of factories) facs.add(f); else facs.add('');
    }
    sx = { list, byStyle, byAlloc, cats, facs, memo: new WeakMap() };
    // The rows must add up to the engine's totals, unit for unit.
    const KEYS = ['now', 'fromChina', 'arriving', 'total', 'soldChina', 'oversold'];
    const off = KEYS.filter(k => typeof sell.totals[k] === 'number' && Math.abs(sell.totals[k] - list.reduce((p, r) => p + units(r.s[k]), 0)) > 1e-6);
    if (off.length) {
      const msg = `The page adds the styles up to different totals than the selling engine: ${off.map(k => `${SELL_TILE_LABEL[k] || k} ${fmtN(list.reduce((p, r) => p + units(r.s[k]), 0))} here, ${fmtN(sell.totals[k])} in the engine`).join('; ')}. This is a bug in Shalam Tool: tell Saul.`;
      console.error(msg);
      sellPageIssues = [{ source: SELL_SRC, code: 'page-sum', sheet: '', row: '', po: '', style: '', message: msg }];
    }
  }

  // ---------- the Available to sell table ----------
  const SELL_COLS = [
    { k: 'style', t: 'Style', cls: 'fz f-sty' }, { k: 'desc', t: 'Description' },
    { k: 'casePack', t: 'Case pack', n: 1, tip: 'Units in a master carton (Master on Master Sheet Live); "mixed" when its goods come in different ones' },
    { k: 'inWh', t: 'In warehouse', n: 1, tip: 'In WH on the weekly ATS report: physical stock' },
    { k: 'orders', t: 'Customer orders', n: 1, tip: 'Total Customer Orders on the ATS report: sold, not shipped yet' },
    { k: 'now', t: 'Available now', b: 1, tip: 'In warehouse minus customer orders, when above zero. Orders beyond the stock come off the earliest goods from China' },
    { k: 'fromChina', t: 'Available from China', b: 1, tip: 'Still in China (in production, ready, booked, loaded), after customer orders' },
    { k: 'firstPull', t: 'First pull date', date: 1, tip: 'The first date goods in China can be pulled: 10 days before the ETD, else 5 days before the CRD, else about 75 days after the order date. "now": that date has passed' },
    { k: 'arriving', t: 'Arriving', b: 1, tip: 'On the water or at port, after customer orders' },
    { k: 'firstArrival', t: 'First arrival', date: 1, tip: 'The earliest ETA of the goods on the water or at port' },
    { k: 'total', t: 'Total available', b: 1, sum: 1, tip: 'Available now + from China + arriving' },
    { k: 'soldChina', t: 'Sold (ships from China)', b: 1, tip: PULL_TIP },
  ];
  const SELL_STOCK = ['inWh', 'orders', 'now', 'total'];
  let sortSell = { k: 'style', dir: 1 }, sellShown = { styles: 0, all: 0 };
  // Something a buyer can have: stock after orders (when known), goods from China or arriving.
  const sellsAny = r => (stockOk() && units(r.s.now) > 0) || units(r.s.fromChina) > 0 || units(r.s.arriving) > 0;
  function sellFiltered() {
    return sx.list.filter(r => (!F.q || r.hay.includes(F.q)) && catOk(r.category) &&
      (!F.fac || (F.fac === NONE ? !r.factories.length : r.factories.includes(F.fac))) && (!F.avail || sellsAny(r)));
  }
  function sellRows() {
    const ok = stockOk(), rows = sellFiltered(), k = sortSell.k;
    const val = k === 'style' ? r => r.styleRaw : k === 'desc' ? r => r.desc : k === 'casePack' ? r => r.s.casePack
      : k === 'firstPull' ? r => r.fp.day : k === 'firstArrival' ? r => r.fa
        : SELL_STOCK.includes(k) ? r => (ok ? numOr(r.s[k]) : null) : r => numOr(r.s[k]);
    rows.sort(cmpBy(val, sortSell.dir));
    sellShown = { styles: rows.length, all: sx.list.length };
    return rows;
  }
  const sellInfo = () => `${plural(sellShown.styles, 'style')}${fOnSell() ? ` of ${fmtN(sellShown.all)}` : ''}`;
  function stockTip(s) {
    const net = numOr(s.stockNet) ?? (units(s.inWh) - units(s.orders));
    return `In warehouse ${fmtN(units(s.inWh))} − customer orders ${fmtN(units(s.orders))} = ${fmtN(net)}${net < 0 ? `: the ${fmtN(-net)} ordered beyond the stock come off the earliest goods from China` : ''}. Click for its lots.`;
  }
  function firstPullHtml(r) {
    const fp = r.fp;
    if (fp.day == null) return fp.open ? '<span class="nt" title="Goods in China with no ETD, CRD or order date (listed on Problems)">no date</span>' : '';
    // The rule in words, without the "from <date>:" a past date carries (the tooltip starts with the date).
    const w = fp.lot ? lotWhen(fp.lot) : null, why = w ? w.why.replace(/^from [^:]*:\s*/, '').replace(/^from .*$/, '') : '';
    const today = sellToday(), now = fp.day < today;
    const word = now ? 'now' : (fp.est ? 'about ' : '') + fmtD(fp.day);
    const tip = `Pull-from date ${fp.est ? 'about ' : ''}${fmtD(fp.day)}${why ? ': ' + why : ''}${now ? ` (${plural(today - fp.day, 'day')} ago)` : ''}`;
    // "now" says since when, so a date passed months ago (a CRD nobody updated) shows.
    return `<span class="${fp.est ? 'est' : ''}" title="${esc(tip)}">${esc(word)}</span>${now ? `<span class="u">since ${esc(md(fp.day))}</span>` : ''}`;
  }
  function firstArrivalHtml(r) {
    if (r.fa == null) return '';
    const passed = r.fa < sellToday();
    return `<span${passed ? ' title="This ETA has passed: due now, or at port"' : ''}>${esc(fmtD(r.fa))}</span>${passed ? '<span class="u">ETA passed</span>' : ''}`;
  }
  // A style the ATS report has no row for, or whose In WH or customer orders on it can't be read: its stock isn't known
  // (never 0), so its lots show in full.
  const notOnReport = s => hasFlag(s, 'not-on-report') || s.onReport === false;
  const stockUnknown = s => hasFlag(s, 'stock-unknown');
  const cantRead = '<span class="nt" title="Blank or not a number on the ATS report (listed on Problems)">can\'t be read</span>';
  function sellRowHtml(r, i) {
    const s = r.s, ok = stockOk(), partial = ok && (notOnReport(s) || stockUnknown(s));
    const nb = (k, v) => (v ? `<button type="button" class="nb" data-b="${k}">${fmtN(v)}</button>` : '');
    const tags = [];
    const maybe = [hasFlag(s, 'erp-in-transit') ? `the ATS report shows ${fmtN(s.erpInTrans)} In Trans, more than we have on the water, at port or booked (landed, not received yet?)` : '',
      hasFlag(s, 'orders-may-include-sold') ? `${fmtN(s.soldChina)} units are sold to ship from China: the orders may be those` : ''].filter(Boolean);
    if (units(s.oversold) > 0) tags.push(`<button type="button" class="tag ov" data-b="all" title="${esc(`Oversold by ${fmtN(s.oversold)} units: customer orders are more than the stock and every lot coming${maybe.length ? '. Check first: ' + maybe.join('; ') : ''}`)}">oversold ${fmtN(s.oversold)}</button>`);
    if (hasFlag(s, 'erp-received')) tags.push(`<button type="button" class="tag um" data-b="all" title="${esc(`${fmtN(s.erpReceived)} units on the water or at port have an ETA on or before the ATS report's date but aren't In Trans on it: the ERP has probably received them, so they may be in Available now and in Arriving (listed on Problems)`)}">received?</button>`);
    if (ok && notOnReport(s)) tags.push(`<span class="tag um" title="${esc(`No row on the ATS report${asOfText()}, so its customer orders aren't known: its goods from China show in full`)}">not on ATS report</span>`);
    else if (ok && stockUnknown(s)) tags.push(`<span class="tag um" title="${esc(`Its In WH or customer orders on the ATS report${asOfText()} can't be read, so its customer orders aren't taken off: its goods from China show in full`)}">stock can't be read</span>`);
    const upcLine = r.upcs.length ? r.upcs[0] + (r.upcs.length > 1 ? ` +${r.upcs.length - 1} more` : '') : '';
    const tip = [r.styleRaw, r.upcs.length ? 'UPC ' + r.upcs.join(', ') : '', r.category, r.factories.length ? 'Factory ' + r.factories.join(', ') : '',
      'Click for its stock and every lot, with their dates'].filter(Boolean).join('\n');
    let h = `<tr class="r" data-i="${i}"><td class="fz f-sty" title="${esc(tip)}"><div class="sc"><button type="button" class="sb" data-b="all">${esc(r.styleRaw)}</button>${tags.join('')}</div><span class="u">${esc(upcLine) || '&nbsp;'}</span></td>` +
      `<td class="desc" title="${esc(r.desc)}">${esc(r.desc)}</td>` +
      `<td class="num"${s.casePack === 'mixed' ? ' title="Its goods come in different case packs"' : ''}>${typeof s.casePack === 'number' ? fmtN(s.casePack) : s.casePack === 'mixed' ? 'mixed' : ''}</td>`;
    const inWh = numOr(s.inWh), orders = numOr(s.orders);
    if (!ok) h += `<td class="nr" colspan="3" title="${esc(atsWhy())}">${esc(stockWord())}</td>`;
    else if (notOnReport(s)) h += `<td class="nr" colspan="3" title="${esc(`This style has no row on the ATS report${asOfText()}`)}">not on report</td>`;
    else if (stockUnknown(s)) {
      h += `<td class="num${inWh < 0 ? ' neg' : ''}">${inWh == null ? cantRead : inWh ? fmtN(inWh) : ''}</td><td class="num">${orders == null ? cantRead : nz(orders)}</td>` +
        '<td class="num nr" title="In WH or customer orders can\'t be read on the ATS report, so stock after orders isn\'t known (never 0)">not known</td>';
    } else {
      h += `<td class="num${inWh < 0 ? ' neg' : ''}">${inWh ? fmtN(inWh) : ''}</td><td class="num">${nz(orders)}</td>` +
        `<td class="num" title="${esc(stockTip(s))}">${nb('now', numOr(s.now))}</td>`;
    }
    h += `<td class="num">${nb('fromChina', numOr(s.fromChina))}</td><td class="nw">${firstPullHtml(r)}</td>` +
      `<td class="num">${nb('arriving', numOr(s.arriving))}</td><td class="nw">${firstArrivalHtml(r)}</td>`;
    // Without the stock (no report, or not on it) Total available would be a part only: the column says so.
    h += !ok ? `<td class="num sum nr" title="${esc(atsWhy())}">${esc(stockWord())}</td>`
      : partial ? `<td class="num sum pt" title="Stock not known: this is from China and arriving only">${nb('total', numOr(s.total))}</td>`
        : `<td class="num sum">${nb('total', numOr(s.total))}</td>`;
    return h + `<td class="num">${nb('soldChina', numOr(s.soldChina))}</td></tr>`;
  }

  // ---------- its tiles: sums of the rows shown, in units and cases ----------
  const SELL_TILE_LABEL = { now: 'Available now', fromChina: 'Available from China', arriving: 'Arriving', total: 'Total available', oversold: 'Oversold styles',
    soldChina: 'Sold (ships from China)' };
  const SELL_TILES = [
    ['now', 'In warehouse minus customer orders (the ATS report), where above zero', false],
    ['fromChina', 'Still in China (in production, ready, booked, loaded), after customer orders. Click for the lots and their pull-from dates', true],
    ['arriving', 'On the water or at port, after customer orders. Click for the lots and their ETAs', true],
    ['total', 'Available now + from China + arriving. Click for the lots', true],
    ['oversold', 'Styles with more customer orders than their stock and every lot coming (each one is on Problems)', false],
  ];
  function sellAgg(rows) {
    const z = () => ({ u: 0, c: 0, np: 0, lo: 0 });
    const g = { now: z(), fromChina: z(), arriving: z(), total: z(), oversold: z(), oversoldStyles: 0 };
    // Whole cases only (a part case is never offered as a case); the rest are loose units.
    const add = (o, u, p) => { if (!u) return; o.u += u; if (p) { const c = Math.floor(u / p + 1e-9); o.c += c; o.lo += u - c * p; } else o.np += u; };
    for (const r of rows) {
      const s = r.s, p = packNum(s);
      if (s.cases && typeof s.cases === 'object') {
        // The engine's cases (each lot by its own Master; stock by the style's), and the units with no case pack.
        const np = s.noCasePackUnits && typeof s.noCasePackUnits === 'object' ? s.noCasePackUnits : {};
        const lo = s.looseUnits && typeof s.looseUnits === 'object' ? s.looseUnits : {};
        for (const k of ['now', 'fromChina', 'arriving']) { g[k].u += units(s[k]); g[k].c += units(s.cases[k]); g[k].np += units(np[k]); g[k].lo += units(lo[k]); }
      } else {
        add(g.now, units(s.now), p);
        // Cases lot by lot (each with its own Master) when the lots add up to the style's number, else by the style's.
        for (const [k, inChina] of [['fromChina', true], ['arriving', false]]) {
          const ls = r.lots.filter(l => !!l.inChina === inChina), want = units(s[k]);
          if (ls.reduce((q, l) => q + units(l.available), 0) === want) for (const l of ls) add(g[k], units(l.available), lotPack(l, s));
          else add(g[k], want, p);
        }
      }
      g.total.u += units(s.total);
      if (units(s.oversold) > 0) { g.oversoldStyles++; add(g.oversold, units(s.oversold), p); }
    }
    g.total.c = g.now.c + g.fromChina.c + g.arriving.c; g.total.np = g.now.np + g.fromChina.np + g.arriving.np; g.total.lo = g.now.lo + g.fromChina.lo + g.arriving.lo;
    return g;
  }
  function renderSellTiles(g, L) {
    const ok = stockOk(), why = atsWhy(), filt = fOnSell() ? '<i>filtered</i>' : '';
    $('sellTiles').innerHTML = SELL_TILES.map(([k, tip, click]) => {
      const t = SELL_TILE_LABEL[k];
      if (!g) return `<div class="glass tile nl" data-t="${k}"><div class="lbl">${esc(t)}</div><div class="n">${esc(L ? L.word : 'not loaded')}</div><div class="cs"></div></div>`;
      const needsStock = k === 'now' || k === 'total' || k === 'oversold';
      if (needsStock && !ok) return `<div class="glass tile nl" data-t="${k}" title="${esc(why)}"><div class="lbl">${esc(t)}</div><div class="n">${esc(stockWord())}</div><div class="cs">needs the ATS report</div></div>`;
      const o = g[k];
      const n = k === 'oversold' ? `${fmtN(g.oversoldStyles)}<small>style${g.oversoldStyles === 1 ? '' : 's'}</small>` : `${fmtN(o.u)}<small>units</small>`;
      const cs = k === 'oversold' ? (o.u ? `${fmtN(o.u)} units oversold` : 'none') : casesText(o.c, o.np, o.u, o.lo);
      const warnLine = !ok && (k === 'fromChina' || k === 'arriving') ? `<div class="cs wn">customer orders not taken off</div>` : '';
      const inner = `<div class="lbl">${esc(t)}${filt}</div><div class="n">${n}</div><div class="cs">${esc(cs)}</div>${warnLine}`;
      const on = panel && panel.kind === 'sellTile' && panel.b === k ? ' on' : '';
      return click ? `<button type="button" class="glass tile${on}" data-t="${k}" title="${esc(tip)}">${inner}</button>`
        : `<div class="glass tile flat${k === 'oversold' && g.oversoldStyles ? ' bad' : ''}" data-t="${k}" title="${esc(tip)}">${inner}</div>`;
    }).join('');
  }

  // ---------- what the tab says when it has no numbers, and the ATS report's banner ----------
  function sellLead() {
    const L = stageLead();
    if (L) return { word: L.word, h: L.h === 'Loading the sources…' ? L.h : 'Available to sell is not loaded', p: L.p, short: L.short };
    if (sellErr) {
      return { word: 'not loaded', h: "Available to sell couldn't be worked out", p: `${cap(sellErr)}. This is a bug in Shalam Tool, not a problem in the files: tell Saul. Where it is and Pull from China still work.`,
        short: "Not loaded: Available to sell couldn't be worked out (see the red banner)." };
    }
    if (!sx) return { word: 'not loaded', h: 'Available to sell is not loaded', p: '', short: 'Not loaded.' };
    return null;
  }
  function renderSellBanner() {
    const el = $('sellNote');
    if (!sx) { el.innerHTML = ''; return; }
    const m = atsMode(), s = st.ats || fresh('ats'), d = atsDay();
    const mail = s.mail ? ` <i>${esc(s.mail.text)}</i>` : '';
    let h = '';
    if (m === 'ok' && sell.atsMissing) h = banner('err', `<b>Customer orders aren't taken off:</b> ${esc(atsWhy())}`);
    else if (m === 'ok' && s.state === 'warn') h = banner('warn', `<b>The ATS report is ${esc(s.label)}</b>${d != null ? ` (report date ${esc(fmtD(d))})` : ''}: stock and customer orders may be out of date.${mail}`);
    else if (m === 'bad') h = banner('err', `<b>The ATS report couldn't be used:</b> ${esc(onFromColon(s.message))} Available now says "not loaded", and customer orders aren't taken off the goods from China.${mail}`);
    else if (m === 'none' || m === 'unset') h = banner('warn', `<b>Customer orders aren't taken off</b> until the weekly ATS report is in ${esc(atsPlace())}${m === 'unset' ? " (the ATS folder's location is blank in config.js)" : ''}: the goods from China show in full, and Available now waits for the report.`);
    else if (m === 'unread') h = banner('warn', "<b>The ATS report wasn't read this time</b>, so customer orders aren't taken off: click Refresh.");
    el.innerHTML = h;
  }
  function renderSell(toTop) {
    if (!vtSell) return;
    const L = sellLead();
    $('sellState').hidden = !L;
    if (L) { $('sellStateH').textContent = L.h; $('sellStateP').textContent = L.p; }
    renderSellBanner();
    const rows = sx ? sellRows() : [];
    renderSellTiles(sx ? sellAgg(rows) : null, L);
    vtSell.set(rows, toTop);
    setInfo();
  }

  // ---------- its side panel: a style's lots in date order, then its stock ----------
  const SELL_PANEL_COLS = [
    { k: 'style', t: 'Style', many: 1 }, { k: 'status', t: 'Status' }, { k: 'po', t: 'PO' }, { k: 'units', t: 'Units', n: 1 },
    { k: 'taken', t: 'Taken by orders', n: 1 }, { k: 'available', t: 'Available', n: 1 },
    { k: 'when', t: 'Available from', date: 1, tip: 'In China: the pull-from date and its rule. Sailed: the ETA' },
    { k: 'readyDate', t: 'Ready date (CRD)', date: 1, tip: "The first of Container Tracking's Ready Date, the PO Book's own date, Factory CRD, CRD Expected and Ship by, with every source's date under it" },
    { k: 'etd', t: 'ETD', date: 1 }, { k: 'eta', t: 'ETA', date: 1 }, { k: 'flags', t: 'Flags' },
  ];
  const LOT_STATUS = ['No CRD', 'In production', 'Ready', 'Booked', 'Loaded', 'On the water', 'At port'];
  let sellPanelCols = SELL_PANEL_COLS;
  // A lot's flags, without the ones its other columns already say (now, about, taken, dates disagree, no date).
  const LOT_SAID = ['now', 'estimate', 'part-taken', 'taken', 'dates-disagree', 'no-date'];
  // Without the ATS report every lot's orders are unknown: the stock line says it once instead.
  const lotFlags = l => flagText((Array.isArray(l.flags) ? l.flags : []).filter(f => !LOT_SAID.includes(f) && (f !== 'orders-unknown' || stockOk())));
  function sellPanelRowHtml(l, i) {
    const w = lotWhen(l), rd = readyInfo(l), a = allocOf(l), fr = (a && a.from && typeof a.from === 'object') ? a.from : {};
    const shown = l.poShown && l.poShown !== l.po ? l.poShown : '';
    const sr = sx && sx.byStyle.get(l.style), p = lotPack(l, sr ? sr.s : {});
    const u = units(l.units), cases = p && u ? plural(Math.ceil(u / p - 1e-9), 'case') : '';
    const cust = l.customer || (a && a.customer) || '';
    const cells = {
      style: `<td class="nw">${esc((sr && sr.styleRaw) || l.style)}<span class="u">${esc(l.upc || '') || '&nbsp;'}</span></td>`,
      status: `<td class="nw">${esc(l.status)}<span class="u">${esc(l.inChina ? 'in China' : 'sailed')}</span></td>`,
      po: `<td class="nw"><b>${esc(l.po)}</b><span class="u">${shown ? esc('as ' + shown) : cust ? esc(cust) : '&nbsp;'}</span></td>`,
      units: `<td class="num">${fmtN(u)}<span class="u">${esc(cases) || '&nbsp;'}</span></td>`,
      taken: `<td class="num">${nz(units(l.taken))}<span class="u">&nbsp;</span></td>`,
      available: `<td class="num"><b>${nz(units(l.available))}</b><span class="u">${units(l.available) || !units(l.taken) ? '&nbsp;' : 'all taken'}</span></td>`,
      when: `<td class="nw${w.none ? ' nt' : ''}${w.est ? ' est' : ''}" title="${esc([w.main, w.why].filter(Boolean).join(': '))}">${esc(w.main)}<span class="u">${esc(w.why) || '&nbsp;'}</span></td>`,
      readyDate: `<td class="nw rd${rd.disagree ? ' flag' : ''}" title="${esc(rd.tip)}">${esc(fmtD(numOr(l.readyDate))) || '<span class="nt">none</span>'}${rd.disagree ? ' <b>dates disagree</b>' : ''}<span class="u">${esc(rd.line) || '&nbsp;'}</span></td>`,
      etd: `<td class="nw">${esc(fmtD(numOr(l.etd)))}<span class="u">${numOr(l.etd) != null && fr.etd ? esc(fr.etd) : '&nbsp;'}</span></td>`,
      eta: `<td class="nw">${esc(fmtD(numOr(l.eta)))}<span class="u">${numOr(l.eta) != null && fr.eta ? esc(fr.eta) : '&nbsp;'}</span></td>`,
      flags: `<td class="cut flag" title="${esc(lotFlags(l))}">${esc(lotFlags(l))}<span class="u">&nbsp;</span></td>`,
    };
    return `<tr class="r" data-i="${i}">${sellPanelCols.map(c => cells[c.k]).join('')}</tr>`;
  }
  function sellFilterWords() {
    const w = [];
    if (F.q) w.push(`search "${F.q}"`);
    if (F.fac) w.push(`factory ${F.fac === NONE ? '(none)' : F.fac}`);
    if (F.cat) w.push(`category ${F.cat === NONE ? '(none)' : F.cat}`);
    if (F.avail) w.push('only styles with something available');
    return w.join(', ');
  }
  // The stock line under a style's lots.
  function stockLine(r) {
    const s = r.s, of = asOfText(), has = r.lots.length > 0;
    const full = has ? 'the lots above show in full' : 'and it has no open lot from China';
    if (!stockOk()) return `Stock: ${stockWord()}. ${atsWhy()} Customer orders aren't taken off${has ? ': the lots above show in full' : ''}.`;
    if (notOnReport(s)) return `Stock: not on the ATS report${of}, so its customer orders aren't known${has ? ': the lots above show in full' : ''}.`;
    if (stockUnknown(s) || numOr(s.inWh) == null || numOr(s.orders) == null) {
      return `Stock (ATS report${of}): in warehouse ${numOr(s.inWh) == null ? "can't be read" : fmtN(s.inWh)}, customer orders ${numOr(s.orders) == null ? "can't be read" : fmtN(s.orders)}: stock after orders isn't known (never 0), so customer orders aren't taken off${has ? ' and ' + full : ''}.`;
    }
    const net = numOr(s.stockNet) ?? (s.inWh - units(s.orders));
    let t = `Stock (ATS report${of}): ${fmtN(s.inWh)} in warehouse − ${fmtN(units(s.orders))} customer orders = ${fmtN(net)} net`;
    // Cross-checks with the ATS report (listed on Problems; never counted).
    const notes = [hasFlag(s, 'erp-received') ? ` Check: ${fmtN(s.erpReceived)} units on the water or at port have an ETA on or before the report's date but aren't In Trans on it, so the ERP has probably received them already: they may be counted twice (in stock and arriving).` : '',
      hasFlag(s, 'erp-in-transit') ? ` Check: the ATS report shows ${fmtN(s.erpInTrans)} In Trans, ${fmtN(s.erpInTransBeyond)} more than we have on the water, at port or booked with an ETD: probably goods that have landed but aren't received yet, which may cover these orders (not counted here).` : '',
      hasFlag(s, 'orders-may-include-sold') ? ` Check: ${fmtN(s.soldChina)} units of this style are sold to ship from China; if those orders are in the ERP's customer orders, they're taken off twice.` : ''].join('');
    if (net >= 0) return t + (net ? `: ${fmtN(net)} available now.` : ': nothing available now.') + notes;
    if (!has) return t + `: ${fmtN(-net)} more ordered than in stock, ${full}: oversold.` + notes;
    const taken = r.lots.reduce((p, l) => p + units(l.taken), 0);
    t += `: the ${fmtN(-net)} ordered beyond the stock come off the earliest lots above (${fmtN(taken)} taken)`;
    return t + (units(s.oversold) ? `, and ${fmtN(s.oversold)} more than every lot holds: oversold.` : '.') + notes;
  }
  function renderSellPanel(toTop) {
    const p = panel;
    sellPanelCols = SELL_PANEL_COLS.filter(c => !c.many || p.kind === 'sellTile');
    const stockEl = $('panelStock');
    if (!sx) {
      $('panelSub').textContent = 'Available to sell is not loaded, so there is nothing to show now.';
      $('panelNote').textContent = ''; stockEl.hidden = true;
      vtPanel.set([], toTop);
      return;
    }
    let lots = [], r = null, rowsShown = null;
    if (p.kind === 'sell') { r = sx.byStyle.get(p.key) || null; lots = r ? r.lots.slice() : []; }
    else {
      rowsShown = sellFiltered();
      lots = rowsShown.flatMap(x => x.lots).filter(l => (p.b === 'fromChina' ? l.inChina : p.b === 'arriving' ? !l.inChina : true));
    }
    let rows = lots.slice().sort((a, b) => lotOrder(a, b) || String(a.style).localeCompare(String(b.style)));
    if (sortPanel) {
      const k = sortPanel.k;
      rows = rows.sort(cmpBy(l => (k === 'when' ? lotDay(l) : k === 'status' ? LOT_STATUS.indexOf(l.status) + 1 || null : k === 'flags' ? lotFlags(l)
        : k === 'style' ? (sx.byStyle.get(l.style) || {}).styleRaw || l.style : ['units', 'taken', 'available', 'readyDate', 'etd', 'eta'].includes(k) ? numOr(l[k]) : l[k]), sortPanel.dir));
    }
    const sum = (list, f) => list.reduce((q, l) => q + units(f(l)), 0);
    const china = lots.filter(l => l.inChina), sailed = lots.filter(l => !l.inChina);
    const left = sum(lots, l => l.available), aChina = sum(china, l => l.available), aSailed = sum(sailed, l => l.available);
    const parts = [plural(lots.length, 'lot'), aChina ? `${fmtN(aChina)} available from China` : '', aSailed ? `${fmtN(aSailed)} arriving` : '',
      sum(lots, l => l.taken) ? `${fmtN(sum(lots, l => l.taken))} taken by customer orders` : '', lots.length && !left ? 'nothing left to sell in them' : ''];
    if (r) {
      if (units(r.s.soldChina)) parts.push(`${fmtN(r.s.soldChina)} sold (ships from China), not available`);
      if (units(r.s.unmatched)) parts.push(`${fmtN(r.s.unmatched)} on Container Tracking with no Master Sheet Live line, not available`);
    }
    $('panelSub').textContent = parts.filter(Boolean).join(' · ');
    $('panelNote').textContent = p.kind === 'sell' ? (r ? '' : "This style isn't on Available to sell now (it may have shipped, or its style changed).")
      : fOnSell() ? `Filters on: ${sellFilterWords()}.` : '';
    let stock = '';
    if (r) stock = stockLine(r);
    else if (p.kind === 'sellTile' && p.b === 'total' && rowsShown) stock = stockOk() ? `Plus ${fmtN(rowsShown.reduce((q, x) => q + units(x.s.now), 0))} available now from stock (not lots, so not listed).` : `Stock: ${stockWord()}. ${atsWhy()}`;
    stockEl.textContent = stock; stockEl.hidden = !stock;
    vtPanel.set(rows, toTop);
  }

  // ---------- on the Pull from China tab: each open line's pull-from date and what's left after orders ----------
  function pullSell(r) {
    if (!sx || !r) return null;
    let v = sx.memo.get(r);
    if (v) return v;
    const ids = Array.isArray(r.allocs) ? r.allocs : [];
    const lots = ids.map(id => sx.byAlloc.get(id)).filter(Boolean).sort(lotOrder);
    let sold = 0;
    for (const id of ids) { const a = idx && idx.byId.get(id); if (a && a.bucket === 'pullFromChina') sold += units(a.units); }
    const soldLine = ptypeOf(r) === 'China Pickup' || hasFlag(r, 'pull-section');
    const first = lots.find(l => lotDay(l) != null) || lots[0] || null;
    v = { lots, sold, soldOnly: !lots.length && (sold > 0 || soldLine), avail: lots.length ? lots.reduce((q, l) => q + units(l.available), 0) : null,
      taken: lots.reduce((q, l) => q + units(l.taken), 0), when: first ? lotWhen(first) : null, dates: new Set(lots.map(lotDay).filter(d => d != null)).size };
    sx.memo.set(r, v);
    return v;
  }
  const pullSellVal = (r, k) => { const v = pullSell(r); if (!v || v.soldOnly) return null; return k === 'day' ? (v.when ? v.when.day : null) : v.avail; };
  function pullSellCells(r) {
    if (!sx) { const w = sellErr ? 'not loaded' : ''; return `<td class="nw nt">${w}<span class="u">&nbsp;</span></td><td class="num nt">${w}<span class="u">&nbsp;</span></td>`; }
    const v = pullSell(r);
    if (v.soldOnly) {
      return `<td class="nw nt" title="${esc(PULL_TIP)}">Sold<span class="u">ships from China</span></td><td class="num nt" title="Sold: never available to sell">Sold<span class="u">&nbsp;</span></td>`;
    }
    if (!v.lots.length) return '<td class="nw"><span class="u">&nbsp;</span></td><td class="num"><span class="u">&nbsp;</span></td>';
    const w = v.when, more = v.dates > 1 ? ` (+${v.dates - 1} later date${v.dates > 2 ? 's' : ''}: the side panel has them)` : '';
    const ok = stockOk();
    const under = [v.taken ? `${fmtN(v.taken)} taken` : '', v.sold ? `${fmtN(v.sold)} sold` : '', !ok ? 'orders not taken off' : ''].filter(Boolean).join(' · ');
    return `<td class="nw${w.none ? ' nt' : ''}${w.est ? ' est' : ''}" title="${esc([w.main, w.why].filter(Boolean).join(': ') + more)}">${esc(w.main)}<span class="u">${esc(w.why) || '&nbsp;'}</span></td>` +
      `<td class="num" title="${esc(`${fmtN(units(r.units))} units on the line${v.taken ? `; ${fmtN(v.taken)} taken by customer orders beyond the stock` : ''}${v.sold ? `; ${fmtN(v.sold)} sold, ships from China` : ''}${!ok ? `. ${atsWhy()} Customer orders aren't taken off.` : ''}`)}">` +
      `${v.avail ? fmtN(v.avail) : v.taken ? '<span class="nt">all taken</span>' : ''}<span class="u">${esc(under) || '&nbsp;'}</span></td>`;
  }

  // ---------- the side panel: the allocations behind a number ----------
  // panel: { kind: 'cell' | 'style' | 'tile' | 'pull', key (style), upc (a UPC line, '' for no UPC), b (bucket), po }
  let panel = null, panelFrom = null;
  const FACT_LABEL = { units: 'Units', qty: 'Quantity', status: 'Open or shipped', po: 'PO', poShown: 'PO in the PO Book', style: 'Style', upc: 'UPC', customer: 'Customer', ptype: 'PO type',
    factory: 'Factory', crd: 'CRD', round: 'Round', containers: 'Containers', etd: 'ETD', eta: 'ETA', stage: 'Stage', master: 'Master (case pack)',
    cases: 'Cases', desc: 'Description', category: 'Category' };
  const PANEL_COLS = [
    { k: 'style', t: 'Style', many: 1 }, { k: 'upc', t: 'UPC' }, { k: 'po', t: 'PO' }, { k: 'poShown', t: 'PO in the PO Book' }, { k: 'customer', t: 'Customer' },
    { k: 'bucket', t: 'Counts in' }, { k: 'units', t: 'Units', n: 1 }, { k: 'cases', t: 'Cases', n: 1 }, { k: 'crd', t: 'CRD', date: 1 },
    { k: 'round', t: 'Round' }, { k: 'containers', t: 'Containers' }, { k: 'etd', t: 'ETD', date: 1 }, { k: 'eta', t: 'ETA', date: 1 },
    { k: 'stage', t: 'Stage' }, { k: 'flags', t: 'Flags' }, { k: 'from', t: 'From' },
  ];
  let sortPanel = null, panelCols = PANEL_COLS;
  const BUCKET_ORDER = [...INCOMING, 'pullFromChina', 'unmatched', ...OUTSIDE];
  // Which file each fact came from. The facts with a column of their own show it under their value (FROM_IN_CELL);
  // the From column names the rest fact by fact, with the sheet rows on its second line.
  const FROM_IN_CELL = ['upc', 'po', 'poShown', 'customer', 'units', 'cases', 'crd', 'round', 'containers', 'etd', 'eta', 'stage'];
  const factName = k => FACT_LABEL[k] || cap(String(k).replace(/([a-z])([A-Z])/g, '$1 $2').toLowerCase());
  function fromOf(a) {
    const f = a.from && typeof a.from === 'object' ? a.from : {};
    const all = Object.entries(f).filter(([, v]) => v).map(([k, v]) => `${factName(k)}: ${v}`);
    const rest = Object.entries(f).filter(([k, v]) => v && !FROM_IN_CELL.includes(k)).map(([k, v]) => `${factName(k)}: ${v}`);
    const rows = [];
    if (a.sheetRow != null && a.sheetRow !== '') rows.push(`Master Sheet Live row ${a.sheetRow}`);
    for (const r of [].concat(a.ctRows || [])) if (r != null) rows.push('Container Tracking ' + (typeof r === 'object' ? `${r.sheet || ''} row ${r.row ?? ''}`.trim() : `row ${r}`));
    return { text: rest.join(' · '), rows: rows.join(' · '), tip: [...all, ...rows].join('\n') };
  }
  // A near-miss says what Master Sheet Live has instead.
  const flagsOf = a => flagText(a.flags) + (a.nearMiss ? ` (Master Sheet Live has PO ${a.nearMiss.po}, style ${a.nearMiss.style})` : '');
  function panelAllocs(p) {
    if (!idx) return [];
    if (p.kind === 'pull') {
      // The line's own allocations, by its PO Book key (an allocation's id changes when its stage moves or rows move);
      // with no key, by id. If the line is gone, the style's allocations on that PO, and the panel says so.
      const mine = p.lineKey ? idx.byLine.get(p.lineKey) || [] : Array.isArray(p.ids) ? p.ids.map(id => idx.byId.get(id)).filter(a => a && !hiddenAlloc(a)) : null;
      p.gone = !!p.lineKey && !mine.length;
      if (mine && (mine.length || !p.lineKey)) return mine.slice();
      const e = idx.byStyle.get(window.InvCommon.normStyle(p.key));
      const po = window.InvCommon.normPo(p.po);
      return e ? e.allocs.filter(a => window.InvCommon.normPo(a.po) === po || window.InvCommon.normPo(a.poShown) === po) : [];
    }
    const pick = p.kind === 'style' ? null : new Set(bucketsOf(p.b));
    const list = p.kind === 'tile' ? idx.list : [idx.byStyle.get(p.key)].filter(Boolean);
    const out = [];
    for (const e of list) {
      if (!catOk(e.s.category || '')) continue;
      const sHit = !F.q || styleHit(e.s);
      for (const a of e.allocs) {
        if (!passes(a, sHit, !!p.sellView) || (pick && !pick.has(a.bucket)) || (p.upc != null && (a.upc || '') !== p.upc)) continue;
        out.push(a);
      }
    }
    return out;
  }
  function panelRowHtml(a, i) {
    const f = fromOf(a), shown = a.poShown && a.poShown !== a.po ? a.poShown : '';
    const fr = a.from && typeof a.from === 'object' ? a.from : {};
    // The file a fact came from, on the line under it (every row two lines, so the windowed rows keep one height).
    const by = (k, has = true) => `<span class="u">${has && fr[k] ? esc(fr[k]) : '&nbsp;'}</span>`;
    const cells = {
      style: `<td class="nw">${esc(a.styleRaw || a.style)}<span class="u">&nbsp;</span></td>`, upc: `<td class="nw">${esc(a.upc || '')}${by('upc', !!a.upc)}</td>`,
      po: `<td class="nw"><b>${esc(a.po)}</b>${by('po')}</td>`,
      poShown: `<td class="nw">${esc(shown)}${by('poShown', !!shown)}</td>`, customer: `<td class="cut" title="${esc(a.customer)}">${esc(a.customer)}${by('customer', !!a.customer)}</td>`,
      bucket: `<td class="nw${OUTSIDE.includes(a.bucket) ? ' nt' : ''}">${esc(stageText(a.bucket))}<span class="u">${OUTSIDE.includes(a.bucket) ? 'not counted' : '&nbsp;'}</span></td>`,
      units: `<td class="num">${fmtN(units(a.units))}${by('units')}</td>`,
      cases: `<td class="num">${typeof a.cases === 'number' ? fmtN(a.cases) : '<span class="np" title="No case pack (Master) for this line">no case pack</span>'}${by('cases', typeof a.cases === 'number')}</td>`,
      crd: `<td class="nw">${esc(fmtD(a.crd))}${by('crd', a.crd != null)}</td>`, round: `<td class="nw">${esc(fmtList(a.round))}${by('round', !!fmtList(a.round))}</td>`,
      containers: `<td class="cut" title="${esc(fmtList(a.containers))}">${esc(fmtList(a.containers))}${by('containers', !!fmtList(a.containers))}</td>`,
      etd: `<td class="nw">${esc(fmtD(a.etd))}${by('etd', a.etd != null)}</td>`, eta: `<td class="nw">${esc(fmtD(a.eta))}${by('eta', a.eta != null)}</td>`,
      // A section with no round goes with the stage ("Loaded · Direct A").
      stage: `<td class="nw" title="${esc(stageSec(a))}">${esc(stageSec(a))}${by('stage')}</td>`,
      flags: `<td class="cut flag" title="${esc(flagsOf(a))}">${esc(flagsOf(a))}<span class="u">&nbsp;</span></td>`,
      from: `<td class="from" title="${esc(f.tip)}">${esc(f.text) || '&nbsp;'}<span class="u">${esc(f.rows) || '&nbsp;'}</span></td>`,
    };
    return `<tr class="r" data-i="${i}">${panelCols.map(c => cells[c.k]).join('')}</tr>`;
  }
  const markTile = () => {
    for (const t of $('tiles').querySelectorAll('.tile[data-t]')) t.classList.toggle('on', !!panel && panel.kind === 'tile' && panel.b === t.dataset.t);
    for (const t of $('sellTiles').querySelectorAll('.tile[data-t]')) t.classList.toggle('on', !!panel && panel.kind === 'sellTile' && panel.b === t.dataset.t);
  };
  const isSellPanel = () => !!panel && (panel.kind === 'sell' || panel.kind === 'sellTile');
  const curPanelCols = () => (isSellPanel() ? sellPanelCols : panelCols);
  function openPanel(p, from) {
    panel = p; panelFrom = from || null; sortPanel = null;
    $('panel').hidden = false;
    renderPanel(true);
    markTile();
    $('btnPanelX').focus();
  }
  function closePanel() {
    if (!panel) return;
    panel = null; $('panel').hidden = true;
    markTile();
    if (panelFrom && document.contains(panelFrom)) panelFrom.focus();
    panelFrom = null;
    runDue();
  }
  function panelTitle(p) {
    const e = idx && idx.byStyle.get(p.key), name = (e && (e.s.styleRaw || e.s.style)) || p.key || '';
    if (p.kind === 'sell') { const r = sx && sx.byStyle.get(p.key); return `${(r && r.styleRaw) || name}: available to sell`; }
    if (p.kind === 'sellTile') return `${SELL_TILE_LABEL[p.b] || p.b}: every style`;
    if (p.kind === 'tile') return `${BUCKET_LABEL[p.b]}: every style`;
    if (p.kind === 'pull') return `PO ${p.po}${p.poShown ? ` (${p.poShown})` : ''} · ${p.name || p.key}${p.upc ? ' · UPC ' + p.upc : ''}`;
    const upc = p.upc != null ? ` · UPC ${p.upc || 'none'}` : '';
    if (p.kind === 'style') return `${name}${upc}: every line`;
    return `${name}${upc} · ${p.b === 'casesIncoming' ? 'Incoming (cases)' : BUCKET_LABEL[p.b] || p.b}`;
  }
  function filterWords() {
    const w = [];
    if (F.q) w.push(`search "${F.q}"`);
    if (F.fac) w.push(`factory ${F.fac === NONE ? '(none)' : F.fac}`);
    if (F.cat) w.push(`category ${F.cat === NONE ? '(none)' : F.cat}`);
    if (F.stage) w.push(`stage ${BUCKET_LABEL[F.stage] || F.stage}`);
    if (F.type) w.push(`PO type ${F.type}`);
    return w.join(', ');
  }
  function renderPanel(toTop) {
    if (!panel) return;
    $('panelTitle').textContent = panelTitle(panel);
    $('panelLbl').textContent = isSellPanel() ? 'What can be sold' : 'Behind this number';
    if (isSellPanel()) { renderSellPanel(toTop); return; }
    $('panelStock').hidden = true;
    panelCols = PANEL_COLS.filter(c => !c.many || panel.kind === 'tile');
    if (!idx) {
      $('panelSub').textContent = 'The stage numbers are not loaded, so there is nothing behind them now.';
      $('panelNote').textContent = '';
      vtPanel.set([], toTop);
      return;
    }
    const list = panelAllocs(panel);
    let rows = list.slice().sort((x, y) => (BUCKET_ORDER.indexOf(x.bucket) - BUCKET_ORDER.indexOf(y.bucket)) || String(x.style).localeCompare(String(y.style)) ||
      String(x.po).localeCompare(String(y.po), undefined, { numeric: true }) || ((x.crd ?? 1e9) - (y.crd ?? 1e9)));
    if (sortPanel) {
      const k = sortPanel.k;
      rows = rows.sort(cmpBy(a => (k === 'from' ? fromOf(a).text : k === 'flags' ? flagText(a.flags) : k === 'round' || k === 'containers' ? fmtList(a[k]) : k === 'bucket' ? BUCKET_ORDER.indexOf(a.bucket) + 1 : a[k]), sortPanel.dir));
    }
    const g = agg0(); for (const a of list) addTo(g, a);
    const key = panel.kind === 'style' || panel.kind === 'pull' ? null : panel.b === 'casesIncoming' ? 'incoming' : panel.b;
    const sumOf = pred => list.filter(pred).reduce((p, a) => p + units(a.units), 0);
    // Unmatched units are never "counted" in the stage numbers: said apart. Units outside every total by why.
    const counted = key ? (key === 'unmatched' ? 0 : g.units[key]) : sumOf(a => totalsOf(a.bucket).length && a.bucket !== 'unmatched');
    const um = key ? (key === 'unmatched' ? g.units.unmatched : 0) : sumOf(a => a.bucket === 'unmatched');
    const outBy = {};
    for (const a of list) if (!totalsOf(a.bucket).length) outBy[a.bucket] = (outBy[a.bucket] || 0) + units(a.units);
    const outside = Object.values(outBy).reduce((p, n) => p + n, 0);
    const cs = key ? casesText(g.cases[key], g.noPack[key], g.units[key]) : '';
    $('panelSub').textContent = [plural(list.length, 'line'), key !== 'unmatched' ? `${fmtN(counted)} units counted` : '',
      um ? `${fmtN(um)} units unmatched (never in the stage numbers)` : '', cs,
      outside ? `${fmtN(outside)} units not counted (${Object.entries(outBy).map(([b, n]) => `${fmtN(n)} ${OUT_WORD[b] || stageText(b).toLowerCase()}`).join(', ')})` : ''].filter(Boolean).join(' · ');
    $('panelNote').textContent = panel.kind === 'pull' ? (panel.gone ? "This line isn't on Master Sheet Live as it was (its PO or style changed, or it shipped): showing this style's lines on the PO instead." : '')
      : panel.sellView ? (fOnSell() ? `Filters on: ${sellFilterWords()}.` : '') : fOn() ? `Filters on: ${filterWords()}.` : '';
    vtPanel.set(rows, toTop);
  }

  // ---------- Problems: the balance line ----------
  function renderBalance() {
    const el = $('balance');
    if (!stages || !stages.ok) {
      el.className = 'balance off';
      el.innerHTML = `<b>Balance check:</b> ${stageErr ? "not worked out, because the stage numbers couldn't be (see the red banner)." : 'not loaded, because the stage numbers are not loaded.'}`;
      return;
    }
    const b = stages.balance, by = {};
    for (const a of stages.allocs) if (a && !CT_ONLY.includes(a.bucket)) by[a.bucket] = (by[a.bucket] || 0) + units(a.units);
    const inChina = IN_CHINA.reduce((p, k) => p + (by[k] || 0), 0);
    const parts = [[inChina, 'in China'], [by.water, 'on the water'], [by.port, 'at port'], [by.landed, 'landed'], [by.pullFromChina, 'sold, shipping from China'],
      [by.short, 'short-shipped'], [by.shippedNotOnCt, 'shipped, not on Container Tracking (treated as landed)'],
      [by.duplicate, 'repeats of a line already inside a Round (shipped once)']].filter(([n]) => n);
    const sum = parts.reduce((p, [n]) => p + n, 0);
    const eq = parts.map(([n, w]) => `${fmtN(n)} ${w}`).join(' + ') || '0';
    // Container Tracking-only units (landed Unmatched ones show nowhere); the planning, ignored and older ones in full.
    const co = {};
    for (const a of stages.allocs) if (a && CT_ONLY.includes(a.bucket) && (a.bucket !== 'unmatched' || !hiddenAlloc(a))) co[a.bucket] = (co[a.bucket] || 0) + units(a.units);
    // Planned units have no Problems line (they're planned, not wrong): said where they show instead.
    const CO_WORD = { ctMore: 'beyond Master Sheet Live', unmatched: 'unmatched', planning: "planned, not booked yet (config.js; not listed below: they show in each style's side panel)",
      ignored: 'in an ignored section (config.js)', belowTracker: "on POs below the PO Status Tracker's lowest PO" };
    const coText = CT_ONLY.filter(k => co[k]).map(k => `${fmtN(co[k])} units ${CO_WORD[k]}`).join(', ');
    const listedBelow = CT_ONLY.some(k => co[k] && k !== 'planning');
    const extra = coText ? `<div class="bp">Not part of it (on Container Tracking only${listedBelow ? ', listed below' : ''}): ${esc(coText)}.</div>` : '';
    const groups = typeof b.groups === 'number' ? ` across ${plural(b.groups, 'PO + style group')}` : '';
    const o = balanceOff();
    if (!balanceBad() && sum === units(b.outcomeUnits)) {
      el.className = 'balance ok';
      el.innerHTML = `<b>Every unit adds up:</b> ${fmtN(b.mslUnits)} units on Master Sheet Live = ${eq}${groups}.${extra}`;
      return;
    }
    el.className = 'balance bad';
    const lines = [];
    if (o.msl || o.bad) {
      const bad = (b.bad || []).slice(0, 8).map(x => `PO ${x.po} ${x.style}: ${fmtN(x.mslUnits)} on Master Sheet Live, ${fmtN(x.outcomeUnits)} in outcomes`).join('; ');
      lines.push(`${fmtN(b.mslUnits)} units on Master Sheet Live, ${fmtN(b.outcomeUnits)} in the outcomes (${eq})${o.msl ? `, ${fmtN(Math.abs(o.msl))} apart` : ''}` +
        `${o.bad ? `; ${plural(o.bad, 'PO + style group')} off${bad ? ` (${bad}${o.bad > 8 ? ', …' : ''})` : ''}` : ''}`);
    }
    if (o.ct) lines.push(`Container Tracking has ${fmtN(b.ct.units)} units on its lines (not counting lines on both Booking and Loaded), but ${fmtN(b.ct.outcomeUnits)} were placed`);
    // The engine's balance holds, but its allocations add up to something else here.
    if (!lines.length) lines.push(`the stage engine counts ${fmtN(b.outcomeUnits)} units in the outcomes, but its lines add up to ${fmtN(sum)} here (${eq})`);
    el.innerHTML = `<b>Units don't add up:</b> ${esc(lines.join('. '))}. This is a bug in Shalam Tool, not a problem in the files: tell Saul.${extra}`;
  }

  // ---------- drawing the stage views ----------
  function renderStageViews(toTop) {
    if (!vtInv) return;
    try {
      renderLead();
      renderFilterLists();
      renderSell(toTop);
      renderInv(toTop);
      renderPull(toTop);
      renderPanel(false);
    } catch (e) {
      if (stageErr) throw e;   // drawing "not loaded" itself failed: nothing more to fall back to
      stageCodeErr(e);
      renderBanners(); renderProblems(); renderStageViews(toTop);
    }
  }

  function bindStageUi() {
    const cellOf = (vt, e) => { const tr = e.target.closest('tr.r'); return tr ? vt.rows[+tr.dataset.i] : null; };
    vtInv = vtable('invPane', 'invTable', { cols: () => INV_COLS.length, head: () => sortHead(INV_COLS, sortInv), row: invRowHtml,
      empty: () => emptyText(fOn() ? 'Nothing matches these filters.' : 'No styles with goods in China, on the way or sold from China.') });
    vtPull = vtable('pullPane', 'pullTable', { rowH: 40, cols: () => PULL_COLS.length, head: () => sortHead(PULL_COLS, sortPull), row: pullRowHtml,
      empty: () => emptyText(fOn() ? 'No open line matches these filters.' : 'No open lines on Master Sheet Live.') });
    vtSell = vtable('sellPane', 'sellTable', { rowH: 40, cols: () => SELL_COLS.length, head: () => sortHead(SELL_COLS, sortSell), row: sellRowHtml,
      empty: () => esc(sellLead() ? sellLead().short : F.avail || fOnSell() ? 'No style matches these filters.' : 'No styles on the ATS report or with goods coming.') });
    // The side panel draws either the lines behind a stage number or a style's lots (Available to sell).
    vtPanel = vtable('panelPane', 'panelTable', { cols: () => curPanelCols().length, head: () => sortHead(curPanelCols(), sortPanel),
      row: (r, i) => (isSellPanel() ? sellPanelRowHtml(r, i) : panelRowHtml(r, i)),
      empty: () => (isSellPanel() ? (sx ? 'Nothing coming from China or on the water for this.' : 'Not loaded.') : idx ? 'No lines behind this number now.' : 'Not loaded.') });
    $('sellTable').addEventListener('click', e => {
      const th = e.target.closest('th[data-s]');
      if (th) { sortSell = nextSort(sortSell, SELL_COLS, th.dataset.s); renderSell(true); return; }
      const btn = e.target.closest('button[data-b]'), r = btn && cellOf(vtSell, e);
      if (!r) return;
      // Sold (ships from China) isn't a lot: its lines are the stage panel's.
      if (btn.dataset.b === 'soldChina') openPanel({ kind: 'cell', key: r.key, b: 'pullFromChina', sellView: true }, btn);
      else openPanel({ kind: 'sell', key: r.key, b: btn.dataset.b }, btn);
    });
    $('sellTiles').addEventListener('click', e => {
      const t = e.target.closest('button.tile[data-t]');
      if (t) openPanel({ kind: 'sellTile', b: t.dataset.t }, t);
    });
    $('invTable').addEventListener('click', e => {
      const th = e.target.closest('th[data-s]');
      if (th) { sortInv = nextSort(sortInv, INV_COLS, th.dataset.s); renderInv(true); return; }
      const btn = e.target.closest('button[data-b]'), r = btn && cellOf(vtInv, e);
      if (!r) return;
      const b = btn.dataset.b;
      openPanel({ kind: b === 'all' ? 'style' : 'cell', key: r.key, upc: r.kind === 'upc' ? r.upc : undefined, b }, btn);
    });
    $('pullTable').addEventListener('click', e => {
      const th = e.target.closest('th[data-s]');
      if (th) { sortPull = nextSort(sortPull, PULL_COLS, th.dataset.s); renderPull(true); return; }
      const btn = e.target.closest('button[data-b]'), r = btn && cellOf(vtPull, e);
      if (r) openPanel({ kind: 'pull', key: r.style, name: r.styleRaw || r.style, po: r.po, poShown: r.poShown || '', upc: r.upc || '', lineKey: r.lineKey || '', ids: Array.isArray(r.allocs) ? r.allocs.slice() : null }, btn);
    });
    $('panelTable').addEventListener('click', e => {
      const th = e.target.closest('th[data-s]');
      if (th) { sortPanel = nextSort(sortPanel, curPanelCols(), th.dataset.s); renderPanel(true); }
    });
    $('tiles').addEventListener('click', e => {
      const go = e.target.closest('[data-go]');
      if (go) { showTab(go.dataset.go); return; }
      const t = e.target.closest('button.tile[data-t]');
      if (t) openPanel({ kind: 'tile', b: t.dataset.t }, t);
    });
    $('btnPanelX').onclick = $('btnPanelClose').onclick = closePanel;
    document.addEventListener('keydown', e => { if (e.key === 'Escape' && panel) { e.preventDefault(); closePanel(); } });
    // Filters: the rows, the tiles and an open panel follow them.
    let ft = null;
    const apply = () => { $('btnClearF').hidden = !filtersOn(); renderSell(true); renderInv(true); renderPull(true); renderPanel(true); };
    $('fq').addEventListener('input', e => { clearTimeout(ft); ft = setTimeout(() => { F.q = e.target.value.trim().toLowerCase(); apply(); }, 150); });
    for (const [id, k] of [['fFac', 'fac'], ['fCat', 'cat'], ['fStage', 'stage'], ['fType', 'type']]) $(id).addEventListener('change', e => { F[k] = e.target.value; apply(); });
    $('fAvail').addEventListener('change', e => { F.avail = e.target.checked; apply(); });
    // Clear clears every filter, the ones hidden on this tab too, so none stays on unseen.
    $('btnClearF').onclick = () => {
      Object.assign(F, { q: '', fac: '', cat: '', stage: '', type: '', avail: false });
      $('fq').value = ''; renderFilterLists(); apply();
    };
    addEventListener('resize', () => { fitPanes(); if (tab === 'sell') vtSell.redraw(); else if (tab === 'inv') vtInv.redraw(); else if (tab === 'pull') vtPull.redraw(); if (panel) vtPanel.redraw(); });
    // Problems: checked marks and the switch that shows them.
    $('issTable').addEventListener('change', e => {
      const ck = e.target.closest('input.ck');
      if (!ck) return;
      const i = shownIss[+ck.dataset.ix];
      if (i) { setChecked(i, ck.checked); renderIssues(); }
    });
    $('pqChecked').addEventListener('change', e => { showChecked = e.target.checked; renderIssues(); });
    // The sources box stays as each person left it.
    const box = $('srcBox');
    box.open = store.get('srcOpen') === '1';
    box.addEventListener('toggle', () => store.set('srcOpen', box.open ? '1' : '0'));
    renderStageViews(true);
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
    showTab(store.get('tab2') || 'sell');
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
    document.addEventListener('visibilitychange', runDue);
    addEventListener('pagehide', () => { if (tab === 'problems') commitSeen(); });
    bindStageUi();
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
      ['InvShipping', 'core/shipping.js'], ['InvStages', 'core/stages.js'], ['InvAts', 'core/ats.js'], ['InvSell', 'core/sell.js'],
      ['XLSX', 'vendor/xlsx-js-style/xlsx.bundle.js']];
    const missing = parts.filter(([name]) => !window[name]).map(p => p[1]);
    if (missing.length) { showSignin(`Part of the page didn't load (${missing.join(', ')}). Reload the page; if it keeps happening, tell Saul.`, true); return; }
    initAuth();
  }

  boot();
})();
