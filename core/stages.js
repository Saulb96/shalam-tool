// Incoming: every Master Sheet Live unit gets one stage (No CRD, In production, Ready, Booked, Loaded, On the water,
// At port, Pull from China) or one outcome outside the totals (landed, short-shipped, shipped but not on Container
// Tracking), by matching each line to Container Tracking on PO and style and asking the forwarder and drayage reports
// where each container is. Container Tracking units with no line are Unmatched. A balance proves nothing is lost.
// Container Tracking sections with no round follow config.js's `sections` (Saul, Oct 29): pullFromChina sections go
// straight to the customer (Pull from China), planning sections aren't booked yet (the Master Sheet stage stands, their
// units are 'planning'), ignore sections are 'ignored'. Lines on POs below the PO Status Tracker's lowest PO are
// 'belowTracker'. Those three are Container Tracking-only outcomes, outside every total.
// Each allocation carries `dates`: every source's date for it (Container Tracking's Ready Date, the PO Book's own date,
// Factory CRD, CRD Expected, Ship by, order date), for core/sell.js to sync the ready date and list disagreements.
// Pure: no network, no DOM; today (a day number) is passed in. Loads in the page (window.InvStages, after
// core/common.js) and in Node (module.exports).
(function (root, factory) {
  const node = typeof module === 'object' && module.exports;
  const api = factory(node ? require('./common.js') : root.InvCommon);
  if (node) module.exports = api;
  root.InvStages = api;
})(typeof globalThis !== 'undefined' ? globalThis : this, function (C) {
  'use strict';
  if (!C) throw new Error('core/common.js must load before core/stages.js');
  const { S } = C;
  const SOURCE = 'Incoming';
  const F = { msl: 'Master Sheet Live', trk: 'PO Status Tracker', pob: 'PO Book', ct: 'Container Tracking',
    fwd: 'Forwarder', dray: 'Drayage', upc: 'UPC_MASTER' };

  // Where a container is, least advanced first.
  const NOT_SAILED = 0, WATER = 1, PORT = 2, LANDED = 3;
  const LEVEL_KEY = ['', 'water', 'port', 'landed'];
  const STAGE = { noCrd: 'No CRD', inProd: 'In production', ready: 'Ready', booked: 'Booked', loaded: 'Loaded',
    water: 'On the water', port: 'At port', landed: 'Landed', shipped: 'Shipped', planned: 'Planned' };
  // Container Tracking stages: when units compete for a cap, the most advanced keeps its units first.
  const RANK = { booked: 1, loaded: 2, water: 3, port: 4, landed: 5 };
  // 'duplicate': a row the PO Book counts as shipped because the same PO, style and quantity is already inside a Round
  // (its 'dupround' issue), when Container Tracking has no units left for it. Outside every total, like 'short'.
  // 'planning', 'ignored', 'belowTracker': Container Tracking units in a planning or ignored section, or on a PO below
  // the PO Status Tracker's lowest PO. Never matched, outside every total, kept so the Container Tracking balance holds.
  const BUCKETS = ['noCrd', 'inProd', 'ready', 'booked', 'loaded', 'water', 'port', 'landed', 'pullFromChina', 'short',
    'shippedNotOnCt', 'duplicate', 'ctMore', 'unmatched', 'planning', 'ignored', 'belowTracker'];
  const CT_ONLY = { planning: 'planning', ignore: 'ignored', below: 'belowTracker' };   // block kind -> its outcome
  // config.js's section lists, by the kind of block they make. A name matches trimmed, in any case, spaces as one.
  const SECTION_LISTS = [['pullFromChina', 'pull'], ['planning', 'planning'], ['ignore', 'ignore']];
  const secKey = v => S(v).replace(/\s+/g, ' ').toLowerCase();
  // Section names once each, however they're spaced or cased (the first spelling is kept).
  const uniqSec = a => { const seen = new Set(); return a.map(S).filter(s => { const k = secKey(s); return k && !seen.has(k) && seen.add(k); }); };
  function sectionRules(sec) {
    if (!sec || typeof sec !== 'object') return null;
    const m = new Map();
    for (const [list, kind] of SECTION_LISTS) for (const name of Array.isArray(sec[list]) ? sec[list] : []) {
      const k = secKey(name);
      if (k && !m.has(k)) m.set(k, kind);
    }
    return m;
  }
  const TOTAL_KEYS = ['noCrd', 'inProd', 'ready', 'booked', 'loaded', 'inChina', 'water', 'port', 'incoming',
    'pullFromChina', 'unmatched'];
  const IN_CHINA = ['noCrd', 'inProd', 'ready', 'booked', 'loaded'];
  // The buckets a China Pickup unit goes to Pull from China instead of.
  const STOCK = new Set([...IN_CHINA, 'water', 'port']);
  const SAILED_RANK = RANK.water;
  // A report's dates this far from the tracker row's are another trip of a reused container number.
  const TRIP_DAYS = 45;
  // A shipped line not on Container Tracking is listed when its CRD or order date is this recent.
  const RECENT_DAYS = 90;

  const MSL_SPEC = {
    upc: { names: ['Barcode', 'UPC'] },
    master: { names: ['Master', 'Master Pack', 'Master Carton'] },
    category: { names: ['Category'] },
    sub: { names: ['Sub', 'Sub Category', 'SubCategory'] },
  };
  const MSL_WHAT = {
    upc: ['Barcode', 'the UPCs come from Container Tracking only'],
    master: ['Master', 'no cases can be worked out from it (units only)'],
    category: ['Category', 'categories come from UPC_MASTER only'],
    sub: ['Sub', 'sub-categories come from UPC_MASTER only'],
  };

  const fmtN = n => Number(n).toLocaleString('en-US');
  const unitsText = n => `${fmtN(n)} unit${n === 1 ? '' : 's'}`;
  const poPart = k => String(k ?? '').split('|')[0].replace(/#\d+$/, '');
  const base6 = po => { const m = String(po).match(/^(\d{6})/); return m ? m[1] : ''; };
  const looseStyle = s => s.replace(/\s+/g, '').replace(/-F$/, '');
  const zero = keys => Object.fromEntries(keys.map(k => [k, 0]));
  const uniq = a => [...new Set(a.filter(v => v != null && v !== ''))];
  const goodMaster = v => { const n = C.num(v); return typeof n === 'number' && Number.isFinite(n) && n > 0 ? n : null; };
  const whole = (u, m) => Math.abs(u / m - Math.round(u / m)) < 1e-9;
  const isDay = v => typeof v === 'number' && Number.isFinite(v);
  const roundText = rounds => (rounds || []).join(', ');
  // The PO Book's own style key, used only to tell which file a fact came from, never to match anything.
  const pbSty = v => String(v ?? '').replace(/\s+/g, '').toUpperCase();
  const counted = a => STOCK.has(a.bucket) || a.bucket === 'pullFromChina' || (a.bucket === 'unmatched' && a.stage !== STAGE.landed);
  const totalKey = a => (counted(a) ? a.bucket : null);

  function build(input) {
    const { po, ct, forwarder, drayage, upc, today, sections } = input || {};
    if (!isDay(today)) throw new TypeError('InvStages.build: today must be a day number');
    const notLoaded = [];
    if (!po || !Array.isArray(po.lines)) notLoaded.push(F.msl);
    if (!ct || ct.ok === false || !Array.isArray(ct.blocks) || !ct.rounds) notLoaded.push(F.ct);
    if (notLoaded.length) return { ok: false, notLoaded };

    const issues = [];
    const issue = (code, poV, style, row, message, extra) =>
      issues.push(Object.assign({ source: SOURCE, code, po: poV || '', style: style || '', row: row ?? '', message }, extra));
    const counts = {
      lines: 0, openLines: 0, shippedLines: 0, groups: 0, mslUnits: 0, negativeQty: 0,
      ctBlocks: 0, ctLines: 0, ctUnits: 0, matchedBlocks: 0, unmatchedBlocks: 0, nearMisses: 0, ambiguousPairs: 0,
      bothTabs: 0, bothTabsUnits: 0, noEtd: 0, etdPassedBooked: 0, trackerOnly: 0, otherTrip: 0, completeInTransit: 0,
      ctMoreUnits: 0, ctMoreLandedUnits: 0, unmatchedUnits: 0, unmatchedLandedUnits: 0,
      pullSectionUnits: 0, planningUnits: 0, ignoredUnits: 0, belowTrackerLines: 0, belowTrackerUnits: 0, belowTrackerNotLandedUnits: 0,
      sectionsNotListed: 0, sectionNotListedUnits: 0,
      duplicateUnits: 0, partCase: 0, noCasePackStyles: 0, stylesNotInUpcMaster: 0, reportsMissing: [],
    };

    // ---------- Master Sheet Live facts by row: Barcode, Master, Category, Sub (none required)
    const mrows = input.mrows || po.mrows || null;
    const mcols = mrows && mrows.length ? C.findColumns(mrows[0] || [], MSL_SPEC).cols : {};
    for (const k of Object.keys(MSL_SPEC)) {
      if (mcols[k] != null) continue;
      const [name, what] = MSL_WHAT[k];
      issue('missing-column', '', '', 1, mrows && mrows.length
        ? `Master Sheet Live has no "${name}" column, so ${what}.`
        : `Master Sheet Live's rows weren't read, so there is no "${name}" column: ${what}.`, { sheet: F.msl });
    }
    const factCache = new Map();
    function factOf(row) {
      if (factCache.has(row)) return factCache.get(row);
      const r = (mrows && row >= 1 && mrows[row - 1]) || [];
      const at = k => (mcols[k] == null ? null : r[mcols[k]] ?? null);
      const f = { upc: C.normUpc(at('upc')), upcRaw: S(at('upc')), master: goodMaster(at('master')),
        category: S(at('category')), sub: S(at('sub')) };
      factCache.set(row, f);
      return f;
    }

    // ---------- the lines, their groups (original Master Sheet PO + style) and candidate (PO, style) pairs
    const trk = po.trk || null, ov = po.ov || null;
    const trackerOf = l => {
      if (!trk || !trk.T || typeof trk.T.get !== 'function') return undefined;
      const sty = pbSty(l.sty), b = l.base || base6(C.normPo(l.po));
      return trk.T.get(b + '|' + sty) || trk.T.get(b + '|' + (sty.endsWith('-F') ? sty.slice(0, -2) : sty + '-F')) || null;
    };
    const groups = new Map(), lines = [];
    const styleMaster = new Map();   // style -> { row, master } of its latest Master Sheet line with a Master
    const styleLatest = new Map();   // style -> latest line (highest row)
    for (const l of po.lines) {
      const style = C.normStyle(l.sty);
      if (!style) continue;
      const status = l.status || (l.src === 'round' ? 'shipped' : 'open');
      const origRaw = l.renamed && l.srcPo ? l.srcPo : l.split && l.parentKey ? poPart(l.parentKey) : poPart(l.key) || l.po;
      const orig = C.normPo(origRaw), shown = C.normPo(l.po);
      let units = typeof l.qty === 'number' && Number.isFinite(l.qty) ? l.qty : 0;
      if (units < 0) {
        counts.negativeQty++;
        issue('negative-qty', orig, style, l.row, `PO ${l.po}, style ${S(l.sty)} (Master Sheet Live row ${l.row}): the PO Book shows ${fmtN(units)} units, less than zero (a split larger than the line), so the line counts no units.`, { sheet: F.msl });
        units = 0;
      }
      const tr = trackerOf(l);
      const fact = factOf(l.row);
      const x = { l, style, styleRaw: S(l.sty), status, orig, origRaw: S(origRaw), shown, units, left: units, fact, tr,
        ptype: l.tracked === false ? 'Unknown' : l.ptype === 'China Pickup' ? 'China Pickup' : 'LA', };
      lines.push(x);
      counts.lines++; counts.mslUnits += units;
      if (status === 'open') counts.openLines++; else counts.shippedLines++;
      const gk = orig + '|' + style;
      let g = groups.get(gk);
      if (!g) groups.set(gk, g = { key: gk, po: orig, poRaw: S(origRaw), style, lines: [], cands: new Set(), chunks: [], planning: [] });
      g.lines.push(x);
      g.cands.add(orig + '|' + style); g.cands.add(shown + '|' + style);
      if (fact.master != null) { const p = styleMaster.get(style); if (!p || l.row > p.row) styleMaster.set(style, { row: l.row, master: fact.master }); }
      const lp = styleLatest.get(style); if (!lp || l.row > lp.l.row) styleLatest.set(style, x);
    }
    counts.groups = groups.size;
    // A pair belongs to the group whose original PO it is; otherwise to the first group that names it.
    const pairIndex = new Map();
    for (const g of groups.values()) pairIndex.set(g.key, g);
    for (const g of groups.values()) for (const p of g.cands) {
      const at = pairIndex.get(p);
      if (!at) pairIndex.set(p, g);
      else if (at !== g && at.key !== p) counts.ambiguousPairs++;
    }
    // Near-miss lookups: same 6-digit PO and same style; same PO and a style that differs only by spaces or -F.
    const byBaseStyle = new Map(), byPoLoose = new Map();
    for (const g of groups.values()) for (const p of g.cands) {
      const [cpo, sty] = [p.slice(0, p.indexOf('|')), p.slice(p.indexOf('|') + 1)];
      const a = base6(cpo) + '|' + sty, b = cpo + '|' + looseStyle(sty);
      if (base6(cpo) && !byBaseStyle.has(a)) byBaseStyle.set(a, { g, po: cpo, style: sty });
      if (!byPoLoose.has(b)) byPoLoose.set(b, { g, po: cpo, style: sty });
    }

    // ---------- the reports, by container
    const fwd = forwarder && forwarder.ok !== false && forwarder.byContainer instanceof Map ? forwarder.byContainer : null;
    let dray = null;
    if (drayage && drayage.ok !== false) {
      if (drayage.byContainer instanceof Map) dray = drayage.byContainer;
      else if (Array.isArray(drayage.records)) dray = new Map(drayage.records.filter(r => r && r.container).map(r => [r.container, r]));
    }
    if (!fwd) counts.reportsMissing.push(F.fwd);
    if (!dray) counts.reportsMissing.push(F.dray);
    // Two dates of one trip: within TRIP_DAYS of each other (when both are known). A date typed without a year can be
    // read a year off (core/ct.js only ever moves a yearless ETD back), so with one, a year apart is the same trip too.
    const near = (a, b, yearless) => a == null || b == null || Math.abs(a - b) <= TRIP_DAYS ||
      (!!yearless && Math.abs(Math.abs(a - b) - 365) <= TRIP_DAYS);
    const yearlessRow = row => !!(row.etaYearless || row.etdYearless);
    const fwdTrip = (f, row) => (f.eta != null && row.eta != null ? near(f.eta, row.eta, yearlessRow(row)) : near(f.etd, row.etd, yearlessRow(row)));
    // Drayage has no ETD. With no ETA on the row yet, its port date must fall between the row's ETD and 2 × TRIP_DAYS
    // after it (an old delivery of a reused container number comes before the ETD).
    function drayTrip(d, row) {
      const x = d.eta ?? d.actualDelivery ?? null;
      if (x == null) return true;
      if (row.eta != null) return near(x, row.eta, yearlessRow(row));
      if (row.etd == null) return true;
      return (yearlessRow(row) ? [row.etd, row.etd - 365, row.etd + 365] : [row.etd]).some(e => x >= e && x - e <= 2 * TRIP_DAYS);
    }
    const otherTripSeen = new Set();
    // quiet: a block whose units aren't staged (below the tracker's lowest PO) still gets its stage, but lists nothing.
    function otherTrip(c, file, row, quiet) {
      if (quiet) return;
      const k = file + '|' + c + '|' + row.sheet + row.row;
      if (otherTripSeen.has(k)) return;
      otherTripSeen.add(k);
      counts.otherTrip++;
      issue('other-trip', '', '', row.row, `Container ${c} (${row.sheet} row ${row.row}, round ${roundText(row.rounds)}): the ${file} report's dates for it are more than ${TRIP_DAYS} days from Container Tracking's, so it's read as another trip of a reused container number and not used for this round.`, { sheet: F.ct, round: (row.rounds || [])[0] });
    }

    // A round on Complete whose container a report still shows in transit: listed once per Open/Complete row.
    const completeBehind = new Map();   // sheet+row -> { row, items: Map(container -> text) }
    // One container on one Open/Complete row: the most advanced evidence, in the contract's order. Reports first;
    // Container Tracking's own dates only when no report covers it. A round on Complete is Landed whatever the reports
    // say (plan: "Landed: its round is on the Complete tab"); the forwarder's dates are still the ones shown.
    function evidence(c, row, complete, quiet) {
      let d = dray ? dray.get(c) : null, f = fwd ? fwd.get(c) : null;
      if (d && !drayTrip(d, row)) { otherTrip(c, F.dray, row, quiet); d = null; }
      if (f && !fwdTrip(f, row)) { otherTrip(c, F.fwd, row, quiet); f = null; }
      const ev = {
        etd: f && f.etd != null ? f.etd : row.etd, eta: f && f.eta != null ? f.eta : row.eta,
        etdFrom: f && f.etd != null ? F.fwd : row.etd != null ? F.ct : '', etaFrom: f && f.eta != null ? F.fwd : row.eta != null ? F.ct : '',
      };
      let rep = null;   // what the reports say, if anything: { level, by }
      const st = f ? S(f.status) : '';
      if (d && (d.actualDelivery != null || d.availability === 'DELIVERED')) rep = { level: LANDED, by: F.dray };
      else if (d && (d.actualPickup != null || d.availability === 'AVAILABLE')) rep = { level: PORT, by: F.dray };
      else if (st === 'Arrived') rep = { level: PORT, by: F.fwd };
      else if (st === 'Sailed' || st === 'Soon To Arrive') rep = { level: WATER, by: F.fwd };
      else if (st === 'Pending Departure') rep = { level: NOT_SAILED, by: F.fwd };
      ev.rep = rep ? rep.level : null;
      if (complete) {
        // Listed when a report says it hasn't arrived, or drayage covers it without a delivery. The forwarder keeps
        // Arrived shipments for days after delivery, so "Arrived" alone is normal for a Complete round.
        const notArrived = st === 'Sailed' || st === 'Soon To Arrive' || st === 'Pending Departure';
        const drayOpen = !!d && !(d.actualDelivery != null || d.availability === 'DELIVERED');
        if (!quiet && rep && rep.level < LANDED && (notArrived || drayOpen)) {
          const k = row.sheet + '|' + row.row;
          const e = completeBehind.get(k) || completeBehind.set(k, { row, items: new Map() }).get(k);
          const says = [st ? `the forwarder says ${st}${f.eta != null ? ` (ETA ${C.fmtDay(f.eta)})` : ''}` : '',
            d && rep.by === F.dray ? `drayage says ${d.actualPickup != null ? 'picked up' : 'available'}` : d ? `drayage says ${S(d.availability).toLowerCase() || 'no delivery'}` : ''].filter(Boolean);
          e.items.set(c, says.join(' and '));
          ev.behind = true;
        }
        return Object.assign(ev, { level: LANDED, by: rep && rep.level === LANDED ? rep.by : F.ct });
      }
      if (rep) return Object.assign(ev, { level: rep.level, by: rep.by });
      return Object.assign(ev, { level: datesLevel(row, false), by: F.ct });
    }
    function datesLevel(row, complete) {
      if (complete) return LANDED;
      if (row.eta != null && row.eta < today) return PORT;
      if (row.etd != null && row.etd < today) return WATER;
      return NOT_SAILED;
    }

    // ---------- what kind of block it is: below the tracker's lowest PO (first, whatever its section), in a round, before
    // any section, or a section with no round: one of config.js's lists, or not on them ('unlisted')
    // The PO Book reads no Master Sheet Live line on a PO below the PO Status Tracker's lowest PO, so those never match.
    const minPO = po.trk && Number.isFinite(po.trk.minPO) ? po.trk.minPO : null;
    const belowTracker = p => minPO != null && !!base6(p) && +base6(p) < minPO;
    const secRules = sectionRules(sections);
    const kindCache = new Map();
    function kindOf(b) {
      let k = kindCache.get(b);
      if (k) return k;
      const name = S(b.section);
      // A header that reads like a round with no number ("Round TBD") is a section like any other here (Container
      // Tracking's own check also lists it as round-no-number).
      k = belowTracker(b.po) ? 'below' : b.rounds && b.rounds.length ? 'round'
        : !name || !secRules ? 'none' : secRules.get(secKey(name)) || 'unlisted';
      kindCache.set(b, k);
      return k;
    }
    const sectionOf = b => (b.rounds && b.rounds.length ? '' : S(b.section));

    // ---------- each Booking/Loaded block's stage, from its round(s): the least advanced of everything it covers
    const blockCache = new Map();
    function blockInfo(b) {
      let bi = blockCache.get(b);
      if (bi) return bi;
      const kind = kindOf(b), quiet = kind === 'below';
      const evs = [], ctRows = [], containers = [], flags = [];
      let level = null, noEtd = !b.rounds || !b.rounds.length;
      for (const k of b.rounds || []) {
        const open = (ct.rounds.open && ct.rounds.open[k]) || [], comp = (ct.rounds.complete && ct.rounds.complete[k]) || [];
        const rows = open.length ? open : comp, complete = !open.length && comp.length > 0;
        if (open.length && comp.length && !flags.includes('round-open-and-complete')) flags.push('round-open-and-complete');
        if (!rows.length) { noEtd = true; level = NOT_SAILED; continue; }
        let rl = null;
        for (const row of rows) {
          ctRows.push({ sheet: row.sheet, row: row.row });
          const here = row.containers && row.containers.length
            ? row.containers.map(c => { containers.push(c); return evidence(c, row, complete, quiet); })
            : [{ level: datesLevel(row, complete), by: F.ct, etd: row.etd, eta: row.eta, etdFrom: row.etd != null ? F.ct : '', etaFrom: row.eta != null ? F.ct : '' }];
          for (const ev of here) { evs.push(ev); rl = rl == null ? ev.level : Math.min(rl, ev.level); }
        }
        level = level == null ? rl : Math.min(level, rl);
      }
      if (level == null) level = NOT_SAILED;
      const stage = level === NOT_SAILED ? (b.sheet === 'Loaded' ? 'loaded' : 'booked') : LEVEL_KEY[level];
      const latest = key => evs.reduce((m, e) => (e[key] != null && (m == null || e[key] > m[key]) ? e : m), null);
      const etdE = latest('etd'), etaE = latest('eta');
      const by = uniq(evs.map(e => e.by));
      const reportSailed = evs.some(e => e.rep != null && e.rep >= WATER);
      const etdPassed = evs.some(e => e.etd != null && e.etd < today);
      const etdPassedBooked = b.sheet === 'Booking' && level < LANDED && etdPassed && !reportSailed;
      // Not sailed because the forwarder says Pending Departure (not because of Container Tracking's own dates).
      const pend = level === NOT_SAILED ? evs.filter(e => e.by === F.fwd && e.level === NOT_SAILED) : [];
      // A listed section has no round by design: its flag says what it is, never "no ETD". An unlisted one says so instead.
      if (kind === 'pull') flags.push('pull-section');
      else if (kind === 'unlisted') flags.push('section-not-listed');
      else if (noEtd && (kind === 'round' || kind === 'none')) flags.push('no-etd');
      if (etdPassedBooked && !quiet) flags.push('etd-passed-booked');
      if (evs.some(e => e.behind)) flags.push('complete-in-transit');
      bi = { kind, stage, rounds: (b.rounds || []).slice(), containers: uniq(containers), ctRows, flags,
        etd: etdE ? etdE.etd : null, eta: etaE ? etaE.eta : null, etdFrom: etdE ? etdE.etdFrom : '', etaFrom: etaE ? etaE.etaFrom : '',
        stageFrom: noEtd && !evs.length ? F.ct : (by.length ? by.join(', ') : F.ct),
        pendingFwd: pend.length ? { etd: pend.reduce((m, e) => (e.etd != null && (m == null || e.etd > m) ? e.etd : m), null) } : null,
        trackerOnly: evs.length > 0 && level < LANDED && evs.every(e => e.by === F.ct), noEtd, etdPassedBooked, used: false };
      blockCache.set(b, bi);
      return bi;
    }

    // ---------- Container Tracking lines: Booking lines also on Loaded (same PO + style + round, or the same section
    // however it's spaced or cased) count once, from Loaded
    const tabKeys = b => (b.rounds && b.rounds.length ? b.rounds : ['none:' + secKey(b.section)]);
    const loadedKeys = new Set();
    for (const b of ct.blocks) if (b.sheet === 'Loaded')
      for (const ln of b.lines || []) for (const k of tabKeys(b)) loadedKeys.add(b.po + '|' + ln.style + '|' + k);
    const unmatchedPairs = new Map();   // po|style -> { po, style, chunks: [] }
    const ctOnly = [];                  // chunks of planning, ignored and below-tracker blocks: never staged
    const unlisted = new Map();         // section key -> { name, tabs, pos, units, row, sheet }
    for (const b of ct.blocks) {
      const kind = kindOf(b);
      // Planning, ignored and below-tracker lines are never staged: their per-line checks stay quiet (one summary line each,
      // or none for planning). A Booking copy of one on Loaded is still left out, so it's placed once.
      const quietLine = !!CT_ONLY[kind];
      for (const ln of b.lines || []) {
        if (typeof ln.qty === 'number' && ln.qty < 0) {
          counts.negativeQty++;
          if (kind !== 'below' && kind !== 'ignore') issue('negative-qty', b.po, ln.style, ln.row, `PO ${b.po}, style ${ln.styleRaw || ln.style} (${b.sheet} row ${ln.row}): Container Tracking shows ${fmtN(ln.qty)} units, less than zero, so the line isn't counted. If it corrects another line, change that line's quantity instead.`, { sheet: b.sheet });
          continue;
        }
        const units = typeof ln.qty === 'number' && Number.isFinite(ln.qty) && ln.qty > 0 ? ln.qty : 0;
        if (!units) continue;
        if (b.sheet === 'Booking') {
          const hit = tabKeys(b).find(k => loadedKeys.has(b.po + '|' + ln.style + '|' + k));
          if (hit != null) {
            counts.bothTabs++; counts.bothTabsUnits += units;
            const rk = typeof hit === 'number' ? `round ${hit}` : `"${S(b.section) || '(no section)'}"`;
            if (!quietLine) issue('both-tabs', b.po, ln.style, ln.row, `PO ${b.po}, style ${ln.styleRaw || ln.style}, ${rk} is on both Booking (row ${ln.row}) and Loaded: its ${unitsText(units)} on Booking aren't counted; Loaded's are.`, { sheet: 'Booking' });
            continue;
          }
        }
        counts.ctLines++; counts.ctUnits += units;
        const bi = blockInfo(b);
        const chunk = { b, ln, bi, units, left: units, stage: bi.stage, pull: kind === 'pull', section: sectionOf(b) };
        const g = pairIndex.get(b.po + '|' + ln.style);
        // Planning, ignored and below-tracker units are never matched for staging: their own outcome (after the groups).
        if (CT_ONLY[kind]) {
          ctOnly.push(Object.assign(chunk, { g: g || null, outcome: CT_ONLY[kind] }));
          if (kind === 'planning' && g) g.planning.push(chunk);
          continue;
        }
        bi.used = true;
        if (kind === 'pull') counts.pullSectionUnits += units;
        if (kind === 'unlisted') {
          const k = secKey(b.section);
          const e = unlisted.get(k) || unlisted.set(k, { name: S(b.section), tabs: new Set(), pos: new Set(), units: 0, row: b.row, sheet: b.sheet }).get(k);
          e.tabs.add(b.sheet); e.pos.add(b.po); e.units += units;
          counts.sectionNotListedUnits += units;
        }
        if (g) { g.chunks.push(chunk); counts.matchedBlocks++; continue; }
        counts.unmatchedBlocks++;
        const pk = b.po + '|' + ln.style;
        let u = unmatchedPairs.get(pk);
        if (!u) unmatchedPairs.set(pk, u = { po: b.po, style: ln.style, styleRaw: ln.styleRaw || ln.style, chunks: [] });
        u.chunks.push(chunk);
      }
    }
    // Blocks staged (not planning, ignored or below the tracker); no-etd counts the blocks listed as such.
    for (const bi of blockCache.values()) {
      if (!bi.used) continue;
      counts.ctBlocks++;
      if (bi.trackerOnly) counts.trackerOnly++;
      if (bi.flags.includes('no-etd')) counts.noEtd++;
      if (bi.etdPassedBooked) counts.etdPassedBooked++;
    }
    for (const { row, items } of completeBehind.values()) {
      counts.completeInTransit++;
      const list = [...items].map(([c, says]) => `${says} for container ${c}`).join('; ');
      issue('complete-in-transit', '', '', row.row, `Round ${roundText(row.rounds)} is on Container Tracking's Complete tab (row ${row.row}), so it counts as Landed, but ${list}. Check that it was delivered, or whether the container number was used again.`, { sheet: row.sheet, round: (row.rounds || [])[0] });
    }
    // Container Tracking pairs with no exact match, by 6-digit PO + style and by PO + loose style, to say where a group's
    // goods may sit on Container Tracking under another spelling.
    const unByBase = new Map(), unByLoose = new Map();
    const push = (m, k, v) => (m.get(k) || m.set(k, []).get(k)).push(v);
    for (const u of unmatchedPairs.values()) {
      if (base6(u.po)) push(unByBase, base6(u.po) + '|' + u.style, u);
      push(unByLoose, u.po + '|' + looseStyle(u.style), u);
    }
    function ctElsewhere(g) {
      const out = new Set();
      for (const p of g.cands) {
        const cpo = p.slice(0, p.indexOf('|')), sty = p.slice(p.indexOf('|') + 1);
        for (const u of unByBase.get(base6(cpo) + '|' + sty) || []) out.add(u);
        for (const u of unByLoose.get(cpo + '|' + looseStyle(sty)) || []) out.add(u);
      }
      return [...out].sort((a, b) => (a.po + '|' + a.style).localeCompare(b.po + '|' + b.style));
    }
    const elsewhereWords = us => us.slice(0, 2).map(u => {
      const n = u.chunks.reduce((s, c) => s + c.units, 0);
      return `${unitsText(n)} of style ${u.styleRaw} on PO ${u.po} (${uniq(u.chunks.map(c => STAGE[c.stage])).join(', ')})`;
    }).join(' and ') + (us.length > 2 ? ` and ${us.length - 2} more` : '');

    // ---------- allocations
    const allocs = [], ids = new Map();
    // Which file a line's CRD came from, following the PO Book's order: a date set in the PO Book; a backorder PO
    // typed on Master Sheet Live uses its Ship by; else Factory CRD / CRD Expected (tracker), else Ship by.
    const crdFrom = x => {
      const l = x.l;
      if (l.crd == null) return '';
      if (l.split) return F.pob;
      const o = ov && typeof ov.get === 'function' ? ov.get(l.key) : null;
      if (o && l.bo && S(o.CRDOverride)) return F.pob;
      if (/-\d+$/.test(x.origRaw)) return F.msl;
      if (x.tr === undefined) return l.tracked ? F.trk : F.msl;
      return x.tr && (x.tr.crd ?? x.tr.exp) != null && (x.tr.crd ?? x.tr.exp) === l.crd ? F.trk : F.msl;
    };
    const factoryFrom = x => (x.tr === undefined ? (x.l.tracked ? F.trk : F.msl) : x.tr && x.tr.sup ? F.trk : F.msl);

    // ---------- every source's dates for an allocation (core/sell.js picks the ready date and lists disagreements)
    // Ship by sits on a Master Sheet Live PO block's first row, as the PO Book reads it; each line takes its block's.
    const MSL_BLOCK = { po: { names: ['Po Number', 'PO Number'] }, shipBy: { names: ['Ship by'] }, name: { names: ['Product name'] } };
    let shipByRows = null;
    function shipByOf(row) {
      if (!shipByRows) {
        shipByRows = new Map();
        const bc = mrows && mrows.length ? C.findColumns(mrows[0] || [], MSL_BLOCK).cols : {};
        if (bc.po != null && bc.shipBy != null) {
          let cur;   // undefined: no PO block yet (or a Round row ended it)
          for (let i = 1; i < mrows.length; i++) {
            const r = mrows[i];
            if (!r) continue;
            if (bc.name != null && /^\s*Round/i.test(S(r[bc.name]))) { cur = undefined; continue; }
            if (S(r[bc.po])) cur = C.dayOf(r[bc.shipBy]);
            if (cur !== undefined) shipByRows.set(i + 1, cur);
          }
        }
      }
      return shipByRows.get(row) ?? null;
    }
    // Container Tracking's Ready Date (on a Booking/Loaded block's PO row): the latest of the blocks given, with its row.
    function ctReadyOf(cs) {
      let best = null;
      for (const c of cs) if (c && c.b && isDay(c.b.readyDate) && (!best || c.b.readyDate > best.day)) best = { day: c.b.readyDate, sheet: c.b.sheet, row: c.b.row };
      return best;
    }
    const noDates = () => ({ ctReady: null, ctRow: null, poBook: null, factoryCrd: null, crdExpected: null, shipBy: null, orderDate: null, original: null });
    // A Master Sheet line's dates. A date the PO Book set itself (a split, or a backorder's CRD override) is `poBook`; a
    // backorder PO typed on Master Sheet Live takes its own Ship by (the PO Book's rule). In both cases the tracker's
    // Factory CRD and CRD Expected (and, for a PO Book date, Ship by) belong to the original PO: kept under `original`,
    // never used as this line's ready date.
    function lineDates(x, cs) {
      const l = x.l, tr = x.tr || null, ct = ctReadyOf(cs);
      const own = { factoryCrd: tr && isDay(tr.crd) ? tr.crd : null, crdExpected: tr && isDay(tr.exp) ? tr.exp : null, shipBy: shipByOf(l.row) };
      const d = Object.assign(noDates(), { ctReady: ct ? ct.day : null, ctRow: ct ? { sheet: ct.sheet, row: ct.row } : null,
        orderDate: isDay(l.od) ? l.od : null });
      const keep = o => (Object.values(o).some(v => v != null) ? o : null);
      if (l.crd != null && crdFrom(x) === F.pob) { d.poBook = l.crd; d.original = keep(own); }
      else if (/-\d+$/.test(x.origRaw)) { d.shipBy = own.shipBy; d.original = keep({ factoryCrd: own.factoryCrd, crdExpected: own.crdExpected }); }
      else Object.assign(d, own);
      return d;
    }
    // Container Tracking units with no Master Sheet line: the block's own Ready Date and Order Date only.
    function ctDates(c) {
      const ct = ctReadyOf([c]);
      return Object.assign(noDates(), { ctReady: ct ? ct.day : null, ctRow: ct ? { sheet: ct.sheet, row: ct.row } : null,
        orderDate: isDay(c.b.orderDate) ? c.b.orderDate : null });
    }
    function addAlloc(a) {
      const base = [a.bucket, a.po, a.style, a.sheetRow ?? '', a.ctRows.length ? a.ctRows[0].sheet + a.ctRows[0].row : ''].join('|');
      const n = (ids.get(base) || 0) + 1;
      ids.set(base, n);
      a.id = n > 1 ? base + '#' + n : base;
      if (a.cases != null && a.master && !whole(a.units, a.master)) a.flags.push('part-case');
      allocs.push(a);
      return a;
    }
    const withCases = (units, master) => (master ? Math.ceil(units / master - 1e-9) : null);
    // An allocation from a Master Sheet line, optionally carrying a Container Tracking chunk's stage and dates.
    function lineAlloc(g, x, c, n, bucket, stageName, flags, stageFrom) {
      const l = x.l, master = x.fact.master;
      const upcV = x.fact.upc || (c && c.ln.upc) || null;
      const from = {
        units: F.msl, status: l.manual ? F.pob : F.msl, po: F.msl,
        poShown: l.renamed || l.split || l.derived ? F.pob : F.msl,
        customer: S(l.cust) ? F.trk : '', ptype: l.tracked === false ? '' : F.trk,
        factory: factoryFrom(x), crd: crdFrom(x), stage: stageFrom,
        upc: x.fact.upc ? F.msl : upcV ? F.ct : '', cases: master ? F.msl : '',
      };
      if (c) Object.assign(from, { round: F.ct, containers: F.ct, etd: c.bi.etdFrom, eta: c.bi.etaFrom });
      return addAlloc({
        bucket, units: n, cases: withCases(n, master), master, style: g.style, styleRaw: x.styleRaw, upc: upcV,
        po: g.po, poShown: x.shown !== g.po ? S(l.po) : '', ptype: x.ptype, customer: S(l.cust), factory: S(l.sup),
        crd: l.crd ?? null, od: l.od ?? null, round: c ? roundText(c.bi.rounds) : '', rounds: c ? c.bi.rounds : [],
        containers: c ? c.bi.containers : [], etd: c ? c.bi.etd : null, eta: c ? c.bi.eta : null, stage: stageName,
        ctStage: c ? STAGE[c.stage] : '', section: c ? c.section : '', sheetRow: l.row, lineKey: l.key, status: x.status,
        ctRows: c ? [{ sheet: c.b.sheet, row: c.ln.row }, ...c.bi.ctRows] : [], flags: uniq([...(flags || []), ...(c ? c.bi.flags : [])]),
        from, group: g.key, dates: lineDates(x, c ? [c] : []),
      });
    }
    // An allocation of Container Tracking units no Master Sheet line takes (more than the sheet, or unmatched).
    function ctAlloc(g, c, n, bucket, flags) {
      const style = c.ln.style, sm = styleMaster.get(style);
      const master = goodMaster(c.ln.master) ?? (sm ? sm.master : null);
      return addAlloc({
        bucket, units: n, cases: withCases(n, master), master, style, styleRaw: c.ln.styleRaw || style, upc: c.ln.upc || null,
        po: g ? g.po : c.b.po, poShown: g && g.po !== c.b.po ? c.b.po : '', ptype: '', customer: '', factory: S(c.b.factory),
        crd: null, od: null, round: roundText(c.bi.rounds), rounds: c.bi.rounds, containers: c.bi.containers,
        etd: c.bi.etd, eta: c.bi.eta, stage: STAGE[c.stage], ctStage: STAGE[c.stage], section: c.section, sheetRow: null, lineKey: '', status: '',
        ctRows: [{ sheet: c.b.sheet, row: c.ln.row }, ...c.bi.ctRows], flags: uniq([...(flags || []), ...c.bi.flags]),
        from: { units: F.ct, po: F.ct, factory: F.ct, stage: c.bi.stageFrom, round: F.ct, containers: F.ct,
          etd: c.bi.etdFrom, eta: c.bi.etaFrom, upc: c.ln.upc ? F.ct : '',
          cases: goodMaster(c.ln.master) ? F.ct : master ? F.msl : '' },
        group: g ? g.key : '', dates: ctDates(c),
      });
    }
    // Lines (in order) take units from chunks (in order) until either runs out.
    function take(xs, chunks, fn) {
      for (const x of xs) for (const c of chunks) {
        if (x.left <= 0) break;
        if (c.left <= 0) continue;
        const n = Math.min(x.left, c.left);
        x.left -= n; c.left -= n;
        fn(x, c, n);
      }
    }
    const byRow = (a, b) => a.l.row - b.l.row || String(a.l.key).localeCompare(String(b.l.key));
    const byCrd = (a, b) => (a.l.crd == null) - (b.l.crd == null) || (a.l.crd ?? 0) - (b.l.crd ?? 0) || byRow(a, b);
    const byRank = (a, b) => RANK[b.stage] - RANK[a.stage] || (a.b.sheet === b.b.sheet ? 0 : a.b.sheet === 'Loaded' ? -1 : 1) || a.ln.row - b.ln.row;
    const groupNote = [];   // per-group facts for the issues below

    for (const g of [...groups.values()].sort((a, b) => a.key.localeCompare(b.key))) {
      const cp = g.lines.some(x => x.ptype === 'China Pickup');
      const bucketOf = st => (cp && STOCK.has(st) ? 'pullFromChina' : st);
      // A direct-to-customer section's units go to Pull from China, like a China Pickup PO's, whatever the PO type.
      const chunkBucket = c => (c.pull ? 'pullFromChina' : bucketOf(c.stage));
      // On Container Tracking only as planned: the line keeps its own stage, and says so.
      const plan = g.planning.length ? ['planning-section'] : [];
      const shipped = g.lines.filter(x => x.status !== 'open').sort(byRow);
      const open = g.lines.filter(x => x.status === 'open').sort(byCrd);
      const chunks = g.chunks.slice().sort(byRank);
      // Unsailed units: a round's (or Booked/Loaded with no round) first, then a direct-to-customer section's (it never
      // moves into a round), whatever their rows' order when both want the same Master Sheet units.
      const sailed = chunks.filter(c => RANK[c.stage] >= SAILED_RANK), unsailed = chunks.filter(c => RANK[c.stage] < SAILED_RANK && !c.pull);
      const pullC = chunks.filter(c => RANK[c.stage] < SAILED_RANK && c.pull);
      const note = { g, cp, ctUnits: chunks.reduce((s, c) => s + c.units, 0), mslBehind: [], ctBehind: [], ctMore: [], short: [], notOnCt: [], dup: [] };
      groupNote.push(note);
      // 1. shipped lines take sailed units
      take(shipped, sailed, (x, c, n) => lineAlloc(g, x, c, n, bucketOf(c.stage), STAGE[c.stage], [], c.bi.stageFrom));
      // 2. sailed units beyond shipped belong to open lines, at their Container Tracking stage
      take(open, sailed, (x, c, n) => { note.mslBehind.push({ c, n }); lineAlloc(g, x, c, n, bucketOf(c.stage), STAGE[c.stage], ['msl-behind-ct'], c.bi.stageFrom); });
      for (const c of sailed) if (c.left > 0) { note.ctMore.push({ c, n: c.left }); ctAlloc(g, c, c.left, 'ctMore', ['ct-more']); c.left = 0; }
      // 3. shipped units not covered take unsailed round units: On the water (the Round row says it sailed)
      take(shipped, unsailed, (x, c, n) => { note.ctBehind.push({ c, n, x }); lineAlloc(g, x, c, n, bucketOf('water'), STAGE.water, ['ct-behind-msl'], F.msl); });
      // 5. open units take the remaining unsailed units, Loaded first, earliest CRD first; then a direct-to-customer
      // section's units (Pull from China), before any shipped line does
      take(open, [...unsailed, ...pullC], (x, c, n) => lineAlloc(g, x, c, n, chunkBucket(c), STAGE[c.stage], [], c.bi.stageFrom));
      // 3b. shipped units still not covered take a direct-to-customer section's units left: they stay Pull from China,
      // at Container Tracking's stage (capped by the group, as rule 1 says; listed as ct-behind-msl)
      take(shipped, pullC, (x, c, n) => { note.ctBehind.push({ c, n, x }); lineAlloc(g, x, c, n, 'pullFromChina', STAGE[c.stage], ['ct-behind-msl'], c.bi.stageFrom); });
      for (const c of [...unsailed, ...pullC]) if (c.left > 0) { note.ctMore.push({ c, n: c.left }); ctAlloc(g, c, c.left, 'ctMore', ['ct-more']); c.left = 0; }
      // 4. shipped units still left: short-shipped, or shipped and not on Container Tracking (treated as landed). A row
      // the PO Book counts as shipped only because it repeats a line inside a Round is its own outcome: Container
      // Tracking having no units for it agrees it shipped once (step 1 already lists it as 'dupround'). A shipped line
      // whose PO and style are only in a planning section is said so by its Problems line, not by a flag.
      for (const x of shipped) if (x.left > 0) {
        const n = x.left; x.left = 0;
        if (x.l.dupRound) { note.dup.push({ x, n }); counts.duplicateUnits += n; lineAlloc(g, x, null, n, 'duplicate', STAGE.shipped, ['dup-round'], F.pob); }
        else if (note.ctUnits > 0) { note.short.push({ x, n }); lineAlloc(g, x, null, n, 'short', STAGE.shipped, ['short'], F.msl); }
        else { note.notOnCt.push({ x, n }); lineAlloc(g, x, null, n, 'shippedNotOnCt', STAGE.shipped, ['shipped-not-on-ct'], F.msl); }
      }
      // 6. open units still left keep their Master Sheet stage (with the planning section they're under, if any)
      const planSec = uniqSec(g.planning.map(c => c.section)).join(', ');
      for (const x of open) if (x.left > 0) {
        const n = x.left; x.left = 0;
        const st = x.l.crd == null ? 'noCrd' : x.l.crd >= today ? 'inProd' : 'ready';
        const a = lineAlloc(g, x, null, n, bucketOf(st), STAGE[st], plan, x.l.crd == null ? F.msl : crdFrom(x) || F.msl);
        if (planSec) a.section = planSec;
        // Container Tracking has this PO and style only in a planning section: its Ready Date there is this line's.
        if (g.planning.length) {
          const ct = ctReadyOf(g.planning);
          if (ct) Object.assign(a.dates, { ctReady: ct.day, ctRow: { sheet: ct.sheet, row: ct.row } });
        }
      }
    }
    // Container Tracking units with no Master Sheet line: Unmatched, at their stage (Landed ones are outside everything).
    const nearOf = u => {
      const a = byBaseStyle.get(base6(u.po) + '|' + u.style);
      if (a && a.po !== u.po) return { ...a, why: `the same style on PO ${a.po}` };
      const b = byPoLoose.get(u.po + '|' + looseStyle(u.style));
      if (b && b.style !== u.style) return { ...b, why: `style ${b.style} on the same PO (the two differ only by spaces or a -F ending, which makes them different items)` };
      return null;
    };
    const stageWords = c => STAGE[c.stage] + (c.section ? ` under "${c.section}"` : '');
    for (const u of [...unmatchedPairs.values()].sort((a, b) => (a.po + '|' + a.style).localeCompare(b.po + '|' + b.style))) {
      const near = nearOf(u);
      for (const c of u.chunks.sort(byRank)) {
        const a = ctAlloc(null, c, c.units, 'unmatched', near ? ['near-miss'] : ['unmatched']);
        if (near) a.nearMiss = { po: near.po, style: near.style };
        if (c.stage === 'landed') counts.unmatchedLandedUnits += c.units;
        else counts.unmatchedUnits += c.units;
      }
      const live = u.chunks.filter(c => c.stage !== 'landed');
      if (!live.length) continue;
      const n = live.reduce((s, c) => s + c.units, 0), first = live[0];
      const where = `${first.b.sheet} row ${first.ln.row}${live.length > 1 ? ` and ${live.length - 1} more` : ''}`;
      const stages = uniq(live.map(stageWords)).join(', ');
      if (near) {
        counts.nearMisses++;
        issue('near-miss', u.po, u.style, first.ln.row, `Container Tracking has ${unitsText(n)} of style ${u.styleRaw} on PO ${u.po} (${where}, ${stages}), but Master Sheet Live has no such line. It does have ${near.why}: if they're the same order, correct one sheet so they match. Until then these units are Unmatched, not in the stage totals.`, { sheet: first.b.sheet, near: { po: near.po, style: near.style } });
      } else {
        issue('unmatched', u.po, u.style, first.ln.row, `Container Tracking has ${unitsText(n)} of style ${u.styleRaw} on PO ${u.po} (${where}, ${stages}) with no line on Master Sheet Live, so they're counted as Unmatched, not in the stage totals.`, { sheet: first.b.sheet });
      }
    }

    // ---------- Container Tracking-only outcomes: planning and ignored sections, and POs below the tracker's lowest PO.
    // Each line is kept (the Container Tracking balance holds) and shown with its section; none is in any total.
    const quote = s => `"${s}"`;
    const sOf = (n, w) => `${fmtN(n)} ${w}${n === 1 ? '' : 's'}`;
    const listPos = pos => { const a = [...pos].sort(); return a.slice(0, 8).join(', ') + (a.length > 8 ? ` and ${fmtN(a.length - 8)} more` : ''); };
    const OUT_FLAG = { planning: 'planning-section', ignored: 'ignored-section', belowTracker: 'below-tracker' };
    const outSum = {};
    for (const c of ctOnly) {
      const a = ctAlloc(c.g, c, c.units, c.outcome, [OUT_FLAG[c.outcome]]);
      if (c.outcome === 'planning') { a.stage = a.ctStage = STAGE.planned; a.from.stage = F.ct; }
      const e = outSum[c.outcome] || (outSum[c.outcome] = { lines: 0, units: 0, notLanded: 0, pos: new Set(), sections: [], first: c });
      e.lines++; e.units += c.units; e.pos.add(c.b.po);
      if (c.stage !== 'landed') e.notLanded += c.units;
      if (c.section) e.sections.push(c.section);
    }
    if (outSum.planning) counts.planningUnits = outSum.planning.units;
    if (outSum.ignored) {
      const e = outSum.ignored, names = uniqSec(e.sections), one = names.length === 1;
      counts.ignoredUnits = e.units;
      issue('ignored-section', '', '', e.first.b.row, `Container Tracking's ${names.map(quote).join(' and ')} section${one ? ' is' : 's are'} ignored, as config.js says: ${sOf(e.lines, 'line')} on ${sOf(e.pos.size, 'PO')} (${listPos(e.pos)}), ${unitsText(e.units)}, not counted anywhere.`, { sheet: e.first.b.sheet });
    }
    if (outSum.belowTracker) {
      const e = outSum.belowTracker;
      counts.belowTrackerLines = e.lines; counts.belowTrackerUnits = e.units; counts.belowTrackerNotLandedUnits = e.notLanded;
      issue('below-tracker', '', '', '', `Container Tracking has ${sOf(e.lines, 'line')} on ${sOf(e.pos.size, 'PO')} below the PO Status Tracker's lowest PO (${minPO}): ${unitsText(e.units)}, ${fmtN(e.notLanded)} of them not landed. The PO Book (so this tool) reads no Master Sheet Live line below it, so they can't be matched and aren't counted anywhere.`);
    }
    // A section with no round that isn't on config.js's lists: counted as Booked/Loaded, listed once per section.
    for (const e of unlisted.values()) {
      counts.sectionsNotListed++;
      const tabs = ['Booking', 'Loaded'].filter(t => e.tabs.has(t)).join(' and ');
      issue('section-not-listed', '', '', e.row, `Section '${e.name}' on ${tabs} has no round and isn't in config.js's section list: counted as Booked/Loaded until it's added (${sOf(e.pos.size, 'PO')}, ${unitsText(e.units)}).`, { sheet: e.sheet });
    }

    // ---------- issues per group
    const sumN = a => a.reduce((s, x) => s + x.n, 0);
    const roundsOf = a => uniq(a.flatMap(x => x.c.bi.rounds)).sort((p, q) => p - q);
    const roundWords = r => (r.length ? ` (round${r.length > 1 ? 's' : ''} ${r.join(', ')})` : '');
    const shippedWords = xs => (xs.every(y => y.x.l.manual) ? 'marked Shipped in the PO Book'
      : xs.some(y => y.x.l.manual) ? 'shipped (inside a Round on Master Sheet Live, or marked Shipped in the PO Book)' : 'shipped (inside a Round on Master Sheet Live)');
    for (const note of groupNote) {
      const g = note.g, first = g.lines.slice().sort(byRow)[0], row = first.l.row;
      const name = `PO ${g.poRaw || g.po}, style ${first.styleRaw}`;
      if (note.mslBehind.length) {
        const n = sumN(note.mslBehind);
        issue('msl-behind-ct', g.po, g.style, row, `${name}: open on Master Sheet Live, but Container Tracking says ${unitsText(n)} sailed${roundWords(roundsOf(note.mslBehind))}. They're counted at Container Tracking's stage; Master Sheet Live probably needs its Round row.`, { sheet: F.msl });
      }
      if (note.ctBehind.length) {
        // What kept them unsailed: the forwarder's Pending Departure (a rolled vessel, or a Round row added too early),
        // or Container Tracking's own tabs and dates.
        // A direct-to-customer section's units are said apart: they stay Pull from China.
        const pull = note.ctBehind.filter(x => x.c.pull), rest = note.ctBehind.filter(x => !x.c.pull);
        const pend = rest.filter(x => x.c.bi.pendingFwd), ctSide = rest.filter(x => !x.c.bi.pendingFwd);
        const etdOf = a => a.reduce((m, x) => (x.c.bi.pendingFwd.etd != null && (m == null || x.c.bi.pendingFwd.etd > m) ? x.c.bi.pendingFwd.etd : m), null);
        const parts = [];
        if (ctSide.length) parts.push(`Container Tracking still shows ${unitsText(sumN(ctSide))} as ${uniq(ctSide.map(x => STAGE[x.c.stage])).join(' / ')}${roundWords(roundsOf(ctSide))}`);
        if (pend.length) { const e = etdOf(pend); parts.push(`the forwarder still says Pending Departure${e != null ? ` (ETD ${C.fmtDay(e)})` : ''} for ${unitsText(sumN(pend))}${roundWords(roundsOf(pend))} although Master Sheet Live has its Round row: check whether the vessel was rolled`); }
        if (pull.length) parts.push(`Container Tracking has ${unitsText(sumN(pull))} under ${uniqSec(pull.map(x => x.c.section)).map(quote).join(' and ')}, which config.js lists as going straight to the customer from China`);
        const counted = [rest.length ? `${pull.length ? unitsText(sumN(rest)) + ' are' : "They're"} counted as ${note.cp ? 'Pull from China (a China Pickup PO)' : 'On the water'}, since the line has shipped` : '',
          pull.length ? `${rest.length ? unitsText(sumN(pull)) + ' are' : "They're"} counted as Pull from China: check whether the customer has them already` : ''].filter(Boolean);
        issue('ct-behind-msl', g.po, g.style, row, `${name}: ${shippedWords(note.ctBehind)}, but ${parts.join('; and ')}. ${counted.join('; ')}.`, { sheet: F.msl });
      }
      const more = note.ctMore.filter(x => x.c.stage !== 'landed'), moreLanded = note.ctMore.filter(x => x.c.stage === 'landed');
      counts.ctMoreUnits += sumN(more); counts.ctMoreLandedUnits += sumN(moreLanded);
      if (more.length) {
        const n = sumN(more), secs = uniqSec(more.map(x => x.c.section));
        issue('ct-more', g.po, g.style, row, `${name}: Container Tracking holds ${unitsText(n)} more than Master Sheet Live${roundWords(roundsOf(more))}${secs.length ? ` (under ${secs.map(quote).join(', ')})` : ''}. Master Sheet Live's quantity is counted; the extra isn't.`, { sheet: F.msl });
      }
      // Container Tracking may hold this PO and style written another way (a near-miss seen from this side).
      const other = note.short.length || note.notOnCt.length ? ctElsewhere(g) : [];
      if (note.short.length) {
        const n = note.short.reduce((s, x) => s + x.n, 0);
        const also = other.length ? ` Container Tracking also has ${elsewhereWords(other)}, matching no Master Sheet Live line: if that's this order, correct one sheet so they match.` : '';
        issue('short', g.po, g.style, row, `${name}: ${shippedWords(note.short)}, but Container Tracking has ${unitsText(n)} fewer for it (short-shipped, or Container Tracking is missing a line). Those ${unitsText(n)} aren't counted.${also}`, { sheet: F.msl });
      }
      if (note.notOnCt.length && !note.cp) {
        const recent = note.notOnCt.filter(x => [x.x.l.crd, x.x.l.od].some(d => d != null && d >= today - RECENT_DAYS));
        if (recent.length) {
          const n = note.notOnCt.reduce((s, x) => s + x.n, 0);
          // Container Tracking lines in a planning section don't count as being on it: said so.
          const planned = uniqSec(g.planning.map(c => c.section));
          const where = other.length
            ? `not on Container Tracking under this PO and style. Container Tracking does have ${elsewhereWords(other)}, matching no Master Sheet Live line: if that's this order, correct one sheet so they match`
            : planned.length ? `on Container Tracking only under ${planned.map(quote).join(' and ')}, which config.js lists as planned, not booked yet`
            : 'nowhere on Container Tracking (air freight, or a missing line?)';
          issue('shipped-not-on-ct', g.po, g.style, row, `${name}: ${shippedWords(note.notOnCt)}, but ${where}. Its ${unitsText(n)} are treated as landed, so they aren't counted.`, { sheet: F.msl });
        }
      }
    }
    // PO type unknown (the PO + style isn't on the PO Status Tracker): counted as stock, listed once per PO.
    const unknownPos = new Map();
    for (const a of allocs) {
      if (a.ptype !== 'Unknown') continue;
      if (!a.flags.includes('unknown-type')) a.flags.push('unknown-type');
      // Units in Pull from China (a direct-to-customer section) aren't stock whatever the type: not part of this line.
      if (!counted(a) || a.bucket === 'pullFromChina') continue;
      const u = unknownPos.get(a.po) || unknownPos.set(a.po, { row: a.sheetRow, styles: new Set(), units: 0 }).get(a.po);
      u.styles.add(a.styleRaw); u.units += a.units;
      if (a.sheetRow != null && (u.row == null || a.sheetRow < u.row)) u.row = a.sheetRow;
    }
    for (const [k, u] of unknownPos) {
      const st = [...u.styles].sort();
      issue('unknown-type', k, st.length === 1 ? st[0] : '', u.row, `PO ${k} (${st.length === 1 ? 'style ' + st[0] : st.length + ' styles'}) isn't on the PO Status Tracker, so whether it's LA stock or a China Pickup isn't known. Its ${unitsText(u.units)} are counted as stock, like LA, until it's added.`, { sheet: F.msl });
    }

    // Blocks with no Open/Complete row or no round, and Booking blocks whose ETD passed with no report saying they sailed.
    for (const b of ct.blocks) {
      const bi = blockCache.get(b);
      if (!bi || !bi.used) continue;
      const ln = (b.lines || []).filter(l => l.qty > 0), first = ln[0];
      const what = `PO ${b.po} (${b.sheet} row ${b.row}${b.section ? ', under "' + b.section + '"' : ''})`;
      const styleV = ln.length === 1 ? first.style : '';
      // A section with no round on config.js's lists has none by design; one not on them has its own line per section.
      if (bi.flags.includes('no-etd')) {
        const missing = (b.rounds || []).filter(k => !((ct.rounds.open || {})[k] || (ct.rounds.complete || {})[k]));
        issue('no-etd', b.po, styleV, b.row, `${what}: ${b.rounds && b.rounds.length ? `round ${missing.join(', ')} has no row on Open or Complete` : 'not in a round'}, so it has no ETD and counts as ${STAGE[bi.stage]}.`, { sheet: b.sheet });
      }
      if (bi.etdPassedBooked) {
        issue('etd-passed-booked', b.po, styleV, b.row, `${what}: still on Booking, but its ETD (${C.fmtDay(bi.etd)}) has passed and no report says it sailed. Counted as ${STAGE[bi.stage]}; check whether it sailed or the vessel was rolled.`, { sheet: b.sheet });
      }
    }

    // ---------- totals, styles, Pull from China
    const totals = { units: zero(TOTAL_KEYS), cases: zero(TOTAL_KEYS), noCasePackUnits: zero(TOTAL_KEYS) };
    const addTo = (t, k, a) => {
      t.units[k] += a.units;
      if (a.cases != null) t.cases[k] += a.cases; else if (t.noCasePackUnits) t.noCasePackUnits[k] += a.units;
    };
    const derive = o => {
      for (const t of Object.values(o)) {
        if (!t) continue;
        t.inChina = IN_CHINA.reduce((s, k) => s + t[k], 0);
        t.incoming = t.inChina + t.water + t.port;
      }
    };
    const styleMap = new Map();
    const upcMap = upc && upc.byStyle instanceof Map ? upc.byStyle : null;
    let partCaseAllocs = 0;
    const partLines = new Set(), partCt = new Set();
    for (const a of allocs) {
      const k = totalKey(a);
      if (!k) continue;
      addTo(totals, k, a);
      if (a.flags.includes('part-case')) {
        partCaseAllocs++;
        if (a.sheetRow != null) partLines.add(a.lineKey); else partCt.add(a.ctRows[0].sheet + a.ctRows[0].row);
      }
      let s = styleMap.get(a.style);
      if (!s) styleMap.set(a.style, s = { style: a.style, units: zero(TOTAL_KEYS), cases: zero(TOTAL_KEYS), noCasePackUnits: zero(TOTAL_KEYS), upcs: new Map(), counted: [], noCasePack: false });
      addTo(s, k, a);
      if (a.cases == null) s.noCasePack = true;
      s.counted.push(a);
      const uk = a.upc || '';
      let u = s.upcs.get(uk);
      if (!u) s.upcs.set(uk, u = { upc: a.upc || null, units: zero(TOTAL_KEYS), cases: zero(TOTAL_KEYS) });
      addTo(u, k, a);
    }
    derive(totals);
    counts.partCase = partCaseAllocs;
    if (partCaseAllocs) {
      issue('part-case', '', '', '', `${fmtN(partCaseAllocs)} quantit${partCaseAllocs === 1 ? 'y' : 'ies'} in the stage totals ${partCaseAllocs === 1 ? "isn't" : "aren't"} whole cases (units ÷ Master has a remainder; ${fmtN(partLines.size)} from Master Sheet Live lines, ${fmtN(partCt.size)} from Container Tracking lines), so ${partCaseAllocs === 1 ? 'its' : 'their'} cases are rounded up.`);
    }
    const allocsByStyle = new Map();
    for (const a of allocs) (allocsByStyle.get(a.style) || allocsByStyle.set(a.style, []).get(a.style)).push(a);
    const styles = [];
    for (const s of [...styleMap.values()].sort((a, b) => (a.style < b.style ? -1 : a.style > b.style ? 1 : 0))) {
      derive({ u: s.units, c: s.cases, n: s.noCasePackUnits });
      for (const u of s.upcs.values()) derive({ u: u.units, c: u.cases });
      const um = upcMap ? upcMap.get(s.style) : null;
      if (upcMap && !(um && um.length)) counts.stylesNotInUpcMaster++;
      const latest = styleLatest.get(s.style);
      const umDesc = um && um.find(e => e.desc), umCat = um && um.find(e => e.category), umSub = um && um.find(e => e.sub);
      const from = {};
      let desc = '', category = '', sub = '';
      if (umDesc) { desc = umDesc.desc; from.desc = F.upc; } else if (latest && S(latest.l.desc)) { desc = S(latest.l.desc); from.desc = latest.tr ? F.trk : F.msl; }
      if (umCat) { category = umCat.category; from.category = F.upc; } else if (latest && latest.fact.category) { category = latest.fact.category; from.category = F.msl; }
      if (umSub) { sub = umSub.sub; from.sub = F.upc; } else if (latest && latest.fact.sub) { sub = latest.fact.sub; from.sub = F.msl; }
      const factories = uniq(s.counted.map(a => a.factory)).sort();
      const customers = uniq(s.counted.map(a => a.customer)).sort();
      const etas = s.counted.filter(a => (a.bucket === 'water' || a.bucket === 'port') && a.eta != null).map(a => a.eta);
      const upcs = [...s.upcs.values()].sort((a, b) => (a.upc == null) - (b.upc == null) || String(a.upc).localeCompare(String(b.upc)));
      // With no Master column at all, its one missing-column issue says so; one per style would bury the rest.
      if (s.noCasePack) counts.noCasePackStyles++;
      if (s.noCasePack && mcols.master != null) {
        issue('no-case-pack', '', s.style, latest ? latest.l.row : '', `Style ${latest ? latest.styleRaw : s.style}: ${unitsText(s.noCasePackUnits.incoming + s.noCasePackUnits.pullFromChina + s.noCasePackUnits.unmatched)} in the stage totals have no Master (case pack) on Master Sheet Live or Container Tracking, so they show as units only, not cases.`, { sheet: F.msl });
      }
      styles.push({
        style: s.style, styleRaw: latest ? latest.styleRaw : s.counted[0].styleRaw, desc, category, sub, from,
        factories, factory: factories.length > 1 ? 'Several' : factories[0] || '', customers,
        upcs, units: s.units, cases: s.cases, noCasePackUnits: s.noCasePackUnits, noCasePack: s.noCasePack,
        nextEta: etas.length ? Math.min(...etas) : null, allocs: allocsByStyle.get(s.style) || [],
      });
    }
    const styleOf = new Map(styles.map(s => [s.style, s]));

    // Pull from China: one row per open Master Sheet line, China Pickup included, with its Container Tracking status.
    // A section with no round goes with its status ("Loaded · Direct A"; in a mix "Loaded (Direct A) 3,000"); units at the Master
    // Sheet stage whose PO and style Container Tracking has only in a planning section are "Planned · <section>".
    const ORDER = ['Landed', 'At port', 'On the water', 'Loaded', 'Booked', STAGE.planned, 'Not on tracker'];
    const byLine = new Map();
    for (const a of allocs) if (a.lineKey && a.status === 'open') (byLine.get(a.lineKey) || byLine.set(a.lineKey, []).get(a.lineKey)).push(a);
    const pull = [];
    const partWords = p => (p.section ? `${p.status} · ${p.section}` : p.status);
    for (const x of lines.filter(y => y.status === 'open').sort((a, b) => a.orig.localeCompare(b.orig) || a.style.localeCompare(b.style) || byRow(a, b))) {
      const as = byLine.get(x.l.key) || [];
      const gp = groups.get(x.orig + '|' + x.style), planned = gp ? uniqSec(gp.planning.map(c => c.section)).join(', ') : '';
      const parts = new Map();
      for (const a of as) {
        const status = a.ctStage || (planned ? STAGE.planned : 'Not on tracker'), section = a.ctStage ? a.section || '' : planned;
        const k = status + '|' + secKey(section);
        (parts.get(k) || parts.set(k, { status, section, units: 0 }).get(k)).units += a.units;
      }
      const ctParts = [...parts.values()].sort((p, q) => ORDER.indexOf(p.status) - ORDER.indexOf(q.status) || p.section.localeCompare(q.section));
      const sd = styleOf.get(x.style);
      const ct2 = as.filter(a => a.ctStage);
      pull.push({
        po: x.orig, poShown: x.shown !== x.orig ? S(x.l.po) : '', style: x.style, styleRaw: x.styleRaw, upc: x.fact.upc || (ct2[0] && ct2[0].upc) || null,
        desc: sd ? sd.desc : S(x.l.desc), customer: S(x.l.cust), ptype: x.ptype, factory: S(x.l.sup), crd: x.l.crd ?? null,
        units: x.units, cases: withCases(x.units, x.fact.master), sheetRow: x.l.row, lineKey: x.l.key,
        ctStatus: ctParts.length === 1 ? partWords(ctParts[0])
          : ctParts.map(p => `${p.section ? `${p.status} (${p.section})` : p.status} ${fmtN(p.units)}`).join(' · ') || 'Not on tracker',
        ctParts,
        round: uniq(ct2.flatMap(a => a.rounds)).sort((p, q) => p - q).join(', '), containers: uniq(ct2.flatMap(a => a.containers)),
        etd: ct2.reduce((m, a) => (a.etd != null && (m == null || a.etd > m) ? a.etd : m), null),
        eta: ct2.reduce((m, a) => (a.eta != null && (m == null || a.eta > m) ? a.eta : m), null),
        flags: uniq(as.flatMap(a => a.flags)), allocs: as.map(a => a.id),
      });
    }

    // ---------- balance: every Master Sheet unit in exactly one outcome; every Container Tracking unit used once
    // Both sums are taken from the allocations themselves (what the page shows), not from the steps' bookkeeping.
    const outByGroup = new Map();
    let ctPlaced = 0;
    for (const a of allocs) {
      if (a.sheetRow != null) outByGroup.set(a.group, (outByGroup.get(a.group) || 0) + a.units);
      if (a.ctRows.length) ctPlaced += a.units;
    }
    const bad = [];
    let outcomeUnits = 0;
    for (const g of groups.values()) {
      const m = g.lines.reduce((s, x) => s + x.units, 0), o = outByGroup.get(g.key) || 0;
      outcomeUnits += o;
      if (Math.abs(m - o) > 1e-6) bad.push({ po: g.po, style: g.style, mslUnits: m, outcomeUnits: o });
    }
    const balance = {
      groups: groups.size, mslUnits: counts.mslUnits, outcomeUnits, diff: counts.mslUnits - outcomeUnits, bad,
      ct: { units: counts.ctUnits, outcomeUnits: ctPlaced, diff: counts.ctUnits - ctPlaced },
    };
    if (bad.length || Math.abs(balance.diff) > 1e-6) {
      issue('balance', '', '', '', `The balance check failed: Master Sheet Live has ${unitsText(counts.mslUnits)} on its lines, but ${unitsText(outcomeUnits)} ended in a stage or outcome (${fmtN(bad.length)} PO + style pair${bad.length === 1 ? '' : 's'} don't add up). This is a bug in the tool, not in the sheets: the stage numbers can't be trusted until it's fixed.`);
    }
    if (Math.abs(balance.ct.diff) > 1e-6) {
      issue('balance', '', '', '', `The balance check failed: Container Tracking has ${unitsText(counts.ctUnits)} on its lines (not counting lines on both Booking and Loaded), but ${unitsText(ctPlaced)} were placed. This is a bug in the tool, not in the sheets: the stage numbers can't be trusted until it's fixed.`);
    }

    return { ok: true, today, totals, styles, allocs, pull, balance, issues, counts };
  }

  return { build, BUCKETS, TOTAL_KEYS, IN_CHINA, STAGE, RANK, TRIP_DAYS, RECENT_DAYS };
});
