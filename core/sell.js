// Available to sell: per style, what a buyer can have now (warehouse stock after customer orders, from the ERP's ATS
// report) and from China (each open lot with one date), after everything already sold (Saul, Oct 30).
// - A lot is one allocation of core/stages.js still in China (No CRD, In production, Ready, Booked, Loaded) or sailed (On
//   the water, At port) that isn't sold: China Pickup POs and direct-to-customer sections are sold, never available.
// - Its ready date is synced: Container Tracking's Ready Date, else the PO Book's own date, else Factory CRD, else CRD
//   Expected, else Ship by (a date more than a year ahead or before 2015, a wrong year, is skipped). Sources more than
//   7 days apart are listed.
// - Pull-from date (Saul's rule) in China: 10 days before the ETD; no ETD, 5 days before the ready date (CRD); no ready
//   date, about 75 days after the order date (an estimate). Sailed lots are available on their ETA.
// - Arriving comes from the ERP's ATS report (Saul, Oct 30): for a style on it, each In Trans quantity with its ETA is a
//   lot (the same snapshot as the stock). Our own on-the-water / at-port lots are matched to it: when the ERP has as much
//   or more, ours are counted there (not again), and the rest of its In Trans is goods we still see on a vessel in China
//   (Booked / Loaded with an ETD, earliest ETD first, taken off those lots), then landed but not received yet (on its
//   shipments due on or before the report's day), then In Trans due after it that our shipping data doesn't show sailed.
//   When the ERP has less, a shipment and one of ours of the same quantity are matched first; of the rest, ours with an
//   ETA on or before the report's day are received (in its In WH, dropped; never more than the style's In WH), and what's
//   left is kept as Arriving from our data (not on the ERP's In Trans yet). A style not on the report, or no report: our
//   own sailed lots, as before. Every unit is accounted for: each lot's units = matched + received + in the ERP's
//   transit + taken by orders + available.
// - Cross-checks, listed and never counted (Saul decides): In Trans our data doesn't show sailed while Container Tracking
//   has the style sailed under a PO with no Master Sheet line and the Master Sheet still has goods in China (maybe offered
//   twice); units taken into the ERP's transit that its Import PO Backlog says are still in China.
// - Customer orders come off stock first, then the lots in date order; what's left is oversold. Cases are whole cases.
// Pure: no network, no DOM; today (a day number) is passed in. Loads in the page (window.InvSell, after core/common.js)
// and in Node (module.exports).
(function (root, factory) {
  const node = typeof module === 'object' && module.exports;
  const api = factory(node ? require('./common.js') : root.InvCommon);
  if (node) module.exports = api;
  root.InvSell = api;
})(typeof globalThis !== 'undefined' ? globalThis : this, function (C) {
  'use strict';
  if (!C) throw new Error('core/common.js must load before core/sell.js');
  const { S } = C;
  const SOURCE = 'Available to sell';

  const IN_CHINA = ['noCrd', 'inProd', 'ready', 'booked', 'loaded'];
  const SAILED = ['water', 'port'];
  const STATUS = { noCrd: 'No CRD', inProd: 'In production', ready: 'Ready', booked: 'Booked', loaded: 'Loaded',
    water: 'On the water', port: 'At port', erp: 'In transit (ERP)' };
  // In China, on a vessel already (the ERP may count it in transit): Booked or Loaded, with an ETD.
  const ON_VESSEL = ['booked', 'loaded'];
  const DAYS_BEFORE_ETD = 10, DAYS_BEFORE_CRD = 5, DAYS_AFTER_ORDER = 75, DISAGREE_DAYS = 7;
  // The ready date's sources, first one with a date wins.
  const READY_ORDER = ['ctReady', 'poBook', 'factoryCrd', 'crdExpected', 'shipBy'];
  const READY_LABEL = { ctReady: "Container Tracking's Ready Date", poBook: "the PO Book's date", factoryCrd: 'Factory CRD',
    crdExpected: 'CRD Expected', shipBy: 'Ship by', crd: 'CRD' };
  const READY_FILE = { ctReady: 'Container Tracking', poBook: 'PO Book', factoryCrd: 'PO Status Tracker',
    crdExpected: 'PO Status Tracker', shipBy: 'Master Sheet Live', crd: 'PO Book' };
  const RULES = ['etd', 'crd', 'order', 'none'];
  // The totals counted in cases too, then (units only) Arriving by where it comes from (after customer orders) and how
  // our own lots meet the ERP's In Trans.
  const CASE_KEYS = ['now', 'fromChina', 'arriving', 'total', 'soldChina', 'oversold', 'ordersTaken'];
  const TOTAL_KEYS = [...CASE_KEYS, 'arrivingErp', 'arrivingOurs', 'erpInTrans', 'ourIncoming', 'matchedOurs', 'receivedDropped',
    'inErpTransitTaken', 'landedNotReceived', 'erpUnmatched', 'keptExtra', 'receivedCapped', 'maybeTwice', 'takenInBacklog'];
  // Our sheets' stages for goods that have sailed or arrived (a sold line at another stage is still in China).
  const GONE_STAGES = ['On the water', 'At port', 'Landed', 'Shipped'];
  // Units of a lot the ERP's report already counts (not available here): matched to its In Trans, received, or in transit.
  const replaced = l => (l.matched || 0) + (l.received || 0) + (l.inErpTransit || 0);

  const fmtN = n => Number(n).toLocaleString('en-US');
  const unitsText = n => `${fmtN(n)} unit${n === 1 ? '' : 's'}`;
  const isDay = v => typeof v === 'number' && Number.isFinite(v);
  const fmt = C.fmtDay;
  const zero = (keys = TOTAL_KEYS) => Object.fromEntries(keys.map(k => [k, 0]));
  const uniq = a => [...new Set(a.filter(v => v != null && v !== ''))];
  // Cases a buyer can have: whole cases only (a part case is never offered as a case); the rest are loose units.
  const casesOf = (units, master) => (isDay(master) && master > 0 ? Math.max(0, Math.floor(units / master + 1e-9)) : null);
  const looseOf = (units, master) => { const c = casesOf(units, master); return c == null ? 0 : Math.max(0, units - c * master); };
  // A source date more than a year ahead, or before 2015 (core/ct.js's "date-far" rule), isn't used as the ready date.
  const FLOOR = C.ymd(2015, 1, 1);
  const farDay = (d, today) => isDay(today) && (d > today + 366 || d < FLOOR);
  const byPo = (a, b) => String(a).localeCompare(String(b), undefined, { numeric: true });

  // The synced ready date of one allocation's dates: { readyDate, readyFrom, sources: [{ key, day, far }], spread, disagree }.
  // The first source with a date wins, skipping a date more than a year ahead or before 2015 (a wrong year: kept in
  // sources, marked far, and still counted for dates-disagree); if every source is far, the first is used.
  // Allocations from an older engine (no `dates`) use their CRD.
  function readyOf(dates, crd, today) {
    if (!dates) return { readyDate: isDay(crd) ? crd : null, readyFrom: isDay(crd) ? 'crd' : null,
      sources: isDay(crd) ? [{ key: 'crd', day: crd, far: false }] : [], spread: 0, disagree: false };
    const sources = READY_ORDER.filter(k => isDay(dates[k])).map(k => ({ key: k, day: dates[k], far: farDay(dates[k], today) }));
    const days = sources.map(s => s.day);
    const spread = days.length > 1 ? Math.max(...days) - Math.min(...days) : 0;
    const use = sources.find(s => !s.far) || sources[0] || null;
    return { readyDate: use ? use.day : null, readyFrom: use ? use.key : null,
      sources, spread, disagree: sources.length > 1 && spread > DISAGREE_DAYS };
  }

  // A lot still in China: its pull-from date and rule. { pullFrom, pullRule, estimate, pullRange }.
  function pullOf(etd, readyDate, orderDate) {
    if (isDay(etd)) return { pullFrom: etd - DAYS_BEFORE_ETD, pullRule: 'etd', estimate: false, pullRange: null };
    if (isDay(readyDate)) return { pullFrom: readyDate - DAYS_BEFORE_CRD, pullRule: 'crd', estimate: false, pullRange: null };
    if (isDay(orderDate)) {
      const d = orderDate + DAYS_AFTER_ORDER;
      return { pullFrom: d, pullRule: 'order', estimate: true, pullRange: { from: orderDate, to: d } };
    }
    return { pullFrom: null, pullRule: 'none', estimate: false, pullRange: null };
  }

  // A lot's date in words, for the side panel: "10 days before ETD 11/14/2026", "5 days before CRD 12/1/2026",
  // "about 1/13/2027: 75 days after the order date 10/30/2026", "ETA 12/3/2026".
  function ruleWords(lot) {
    if (!lot) return '';
    switch (lot.pullRule || lot.arrivalRule) {
      case 'etd': return `${DAYS_BEFORE_ETD} days before ETD ${fmt(lot.etd)}`;
      case 'crd': return `${DAYS_BEFORE_CRD} days before CRD ${fmt(lot.readyDate)}`;
      case 'order': return `about ${fmt(lot.pullFrom)}: ${DAYS_AFTER_ORDER} days after the order date ${fmt(lot.dates.orderDate)}`;
      case 'eta': return `ETA ${fmt(lot.eta)}`;
      default: return lot.inChina ? 'no ETD, CRD or order date' : 'no ETA';
    }
  }

  // input: { stages (InvStages.build's result), ats (InvAts.parseAts's result, or null), today (a day number),
  //          upc (InvUpc's result, optional: UPCs and descriptions for styles with stock only),
  //          reportDay (optional: the ATS report's day when its file name has none, e.g. the day it was saved) }.
  function buildSell(input) {
    const { stages, ats, today, upc } = input || {};
    if (!isDay(today)) throw new TypeError('InvSell.buildSell: today must be a day number');
    if (!stages || stages.ok === false || !Array.isArray(stages.allocs)) {
      return { ok: false, today, notLoaded: (stages && stages.notLoaded) || ['Master Sheet Live', 'Container Tracking'] };
    }
    const atsOk = !!(ats && ats.ok !== false && ats.styles instanceof Map);
    const atsMissing = !atsOk;
    const repDay = !atsOk ? null : isDay(input.reportDay) ? input.reportDay : isDay(ats.reportDay) ? ats.reportDay : null;
    // The report's In Trans Qty (with its ETA Date) is Arriving for every style on it; without that column, our own
    // sailed lots are, for every style (one line on Problems).
    const transKnown = atsOk && Array.isArray(ats.columns) && ats.columns.includes('inTrans');
    const issues = [];
    const issue = (code, po, style, message, extra) => issues.push(Object.assign({ source: SOURCE, code, po: po || '', style: style || '', row: '', message }, extra));
    const counts = {
      styles: 0, stylesOnReport: 0, stylesNotOnReport: 0, stylesStockUnknown: 0, stylesWithStock: 0, oversoldStyles: 0,
      erpOnlyStyles: 0, ordersUnknownStyles: 0, ordersUnknownUnits: 0,
      lots: 0, lotsInChina: 0, lotsSailed: 0, lotsTaken: 0, lotsFullyTaken: 0, pullNow: 0,
      byRule: Object.fromEntries(RULES.map(k => [k, 0])), byRuleUnits: Object.fromEntries(RULES.map(k => [k, 0])),
      sailedByRule: { eta: 0, none: 0 }, readyFrom: Object.fromEntries([...READY_ORDER, 'crd', 'none'].map(k => [k, 0])),
      datesDisagreeLots: 0, datesDisagreePairs: 0, datesDisagreeSailedLots: 0, noDateLots: 0,
      readyNotCrd: 0, readyFarSkipped: 0, soldAllocs: 0, unmatchedAllocs: 0,
      // The ERP's In Trans against our own lots (styles / lots / units):
      erpStyles: 0, erpLots: 0, erpLotsNoEta: 0, erpEtaMismatchStyles: 0, erpInTransUnknownStyles: 0, noInTransColumn: atsOk && !transKnown,
      erpMoreStyles: 0, erpLessStyles: 0, matchedLots: 0, receivedStyles: 0, receivedLots: 0, inErpTransitStyles: 0, inErpTransitLots: 0,
      landedNotReceivedStyles: 0, keptExtraStyles: 0, keptExtraLots: 0, noReportDayStyles: 0,
      erpUnmatchedStyles: 0, sameQtyMatchedLots: 0, receivedCappedStyles: 0, maybeTwiceStyles: 0, takenInBacklogStyles: 0,
      soldInOrdersStyles: 0, soldInOrdersUnits: 0, soldInOrdersOversold: 0,
    };

    // ---------- lots, sold and unmatched units, from the allocations
    const lots = [], soldBy = new Map(), soldCasesBy = new Map(), soldNoPackBy = new Map(), unmatchedBy = new Map();
    const latestMaster = new Map();   // style -> { row, master } of its latest Master Sheet line with a Master
    const knownStyles = new Set();    // every style on any Master Sheet Live or Container Tracking line
    // For the cross-checks with the ERP's report: Container Tracking units on the water or at port with no Master Sheet
    // line for them (unmatched, or more than the sheet), and sold units still in China (its backlog counts those too).
    const ctSailedNoLineBy = new Map(), soldInChinaBy = new Map();
    const add = (m, k, n) => m.set(k, (m.get(k) || 0) + n);
    for (const a of stages.allocs) {
      knownStyles.add(a.style);
      if (a.sheetRow != null && isDay(a.master) && a.master > 0) {
        const p = latestMaster.get(a.style);
        if (!p || a.sheetRow > p.row) latestMaster.set(a.style, { row: a.sheetRow, master: a.master });
      }
      if ((a.bucket === 'unmatched' || a.bucket === 'ctMore') && (a.stage === 'On the water' || a.stage === 'At port') && a.units > 0) add(ctSailedNoLineBy, a.style, a.units);
      const lotBucket = IN_CHINA.includes(a.bucket) || SAILED.includes(a.bucket);
      const sold = a.bucket === 'pullFromChina' || (lotBucket && (a.ptype === 'China Pickup' || (a.flags || []).includes('pull-section')));
      if (sold) {
        counts.soldAllocs++;
        add(soldBy, a.style, a.units);
        if (IN_CHINA.includes(a.bucket) || (a.bucket === 'pullFromChina' && !GONE_STAGES.includes(a.stage))) add(soldInChinaBy, a.style, a.units);
        if (a.cases != null) add(soldCasesBy, a.style, a.cases); else add(soldNoPackBy, a.style, a.units);
        continue;
      }
      if (a.bucket === 'unmatched' && a.stage !== 'Landed') { counts.unmatchedAllocs++; add(unmatchedBy, a.style, a.units); continue; }
      if (!lotBucket || !(a.units > 0)) continue;
      const inChina = IN_CHINA.includes(a.bucket);
      const dates = a.dates || null;
      const r = readyOf(dates, a.crd, today);
      const orderDate = dates ? dates.orderDate : isDay(a.od) ? a.od : null;
      const lot = {
        id: a.id, allocId: a.id, style: a.style, styleRaw: a.styleRaw, upc: a.upc || null, po: a.po, poShown: a.poShown || '',
        status: STATUS[a.bucket], bucket: a.bucket, source: 'ours', inChina, units: a.units,
        matched: 0, received: 0, inErpTransit: 0, notOnErp: 0, taken: 0, available: a.units,
        master: isDay(a.master) && a.master > 0 ? a.master : null, cases: null,
        readyDate: r.readyDate, readyFrom: r.readyFrom, readySources: r.sources, readySpread: r.spread,
        dates: Object.assign({ ctReady: null, ctRow: null, poBook: null, factoryCrd: null, crdExpected: null, shipBy: null, orderDate: null, original: null },
          dates || {}, { orderDate }),
        crd: isDay(a.crd) ? a.crd : null, etd: isDay(a.etd) ? a.etd : null, eta: isDay(a.eta) ? a.eta : null,
        pullFrom: null, pullRule: null, pullRange: null, pullNow: false, arrival: null, arrivalRule: null,
        availableDate: null, estimate: false, flags: [], allocFlags: (a.flags || []).slice(),
        factory: a.factory || '', customer: a.customer || '', ptype: a.ptype || '', round: a.round || '', containers: a.containers || [],
        section: a.section || '', sheetRow: a.sheetRow ?? null,
      };
      if (inChina) {
        Object.assign(lot, pullOf(lot.etd, lot.readyDate, orderDate));
        lot.availableDate = lot.pullFrom;
        lot.pullNow = lot.pullFrom != null && lot.pullFrom < today;
        counts.byRule[lot.pullRule]++; counts.byRuleUnits[lot.pullRule] += lot.units;
        if (lot.pullNow) { lot.flags.push('now'); counts.pullNow++; }
        if (lot.estimate) lot.flags.push('estimate');
        counts.lotsInChina++;
      } else {
        lot.arrival = lot.eta; lot.arrivalRule = lot.eta != null ? 'eta' : 'none';
        lot.availableDate = lot.arrival;
        counts.sailedByRule[lot.arrivalRule]++;
        counts.lotsSailed++;
      }
      // No date to sell it from: flagged once the ERP's In Trans has been matched (a sailed lot it counts needs none).
      if (r.disagree) {
        if (inChina) { lot.flags.push('dates-disagree'); counts.datesDisagreeLots++; }
        else counts.datesDisagreeSailedLots++;
      }
      counts.readyFrom[r.readyFrom || 'none']++;
      if (r.sources.some(s => s.far && s.key !== r.readyFrom) && r.readyFrom) counts.readyFarSkipped++;
      // Without Container Tracking's Ready Date, the synced date is the PO Book's CRD (a check on the engine's dates).
      if (dates && !isDay(dates.ctReady) && (lot.readyDate ?? null) !== lot.crd) counts.readyNotCrd++;
      lots.push(lot);
    }
    const lotOrder = (a, b) => (a.availableDate == null) - (b.availableDate == null) || (a.availableDate ?? 0) - (b.availableDate ?? 0) ||
      byPo(a.po, b.po) || String(a.id).localeCompare(String(b.id));
    lots.sort(lotOrder);
    const lotsBy = new Map();
    for (const l of lots) (lotsBy.get(l.style) || lotsBy.set(l.style, []).get(l.style)).push(l);

    // ---------- the ERP's In Trans, for a style on the ATS report (Saul, Oct 30)
    // One shipment of the report's In Trans (its i-th quantity on report row x, ETA from the ETA Date list lined up
    // with it): an Arriving lot, dated by its ETA.
    const NO_DATES = { ctReady: null, ctRow: null, poBook: null, factoryCrd: null, crdExpected: null, shipBy: null, orderDate: null, original: null };
    function erpLot(style, styleRaw, x, i, n, eta) {
      return { id: `erp|${style}|${x.row}|${i + 1}`, allocId: null, style, styleRaw, upc: null, po: '', poShown: '',
        status: STATUS.erp, bucket: 'erp', source: 'erp', inChina: false, units: n,
        matched: 0, received: 0, inErpTransit: 0, notOnErp: 0, taken: 0, available: n, master: null, cases: null,
        readyDate: null, readyFrom: null, readySources: [], readySpread: 0, dates: Object.assign({}, NO_DATES),
        crd: null, etd: null, eta, pullFrom: null, pullRule: null, pullRange: null, pullNow: false,
        arrival: eta, arrivalRule: eta != null ? 'eta' : 'none', availableDate: eta, estimate: false,
        flags: eta == null ? ['erp-eta-mismatch'] : [], allocFlags: [],
        factory: '', customer: '', ptype: '', round: '', containers: [], section: '', sheetRow: null,
        warehouse: S(x.warehouse), reportRow: x.row, shipment: i + 1 };
    }
    const etaProblems = [];   // styles whose ETA Date list doesn't line up with In Trans: { style, styleRaw, rows, units }
    const backlogKnown = atsOk && Array.isArray(ats.columns) && ats.columns.includes('backlog');
    const sumU = (a, f = l => l.units) => a.reduce((t, l) => t + f(l), 0);
    const etaOrder = (a, b) => (a.eta == null) - (b.eta == null) || (a.eta ?? 0) - (b.eta ?? 0) || byPo(a.po, b.po) || String(a.id).localeCompare(String(b.id));
    // Our lots of one style (ls, in date order) against its row(s) on the report (e). Sets each lot's matched / received
    // / inErpTransit / notOnErp and available; answers { mode ('erp' | 'ours'), erpLots, erpInTrans, ourSailed,
    // matched, received, taken (into the ERP's transit), landed (In Trans left on shipments due by the report's day:
    // landed, not received), unmatched (In Trans left, due after it: not sailed on our data), kept, capped (received
    // beyond In WH, kept instead), and the cross-checks maybeTwice and takenInBacklog (units, listed only), flags }.
    function erpStep(style, styleRaw, e, ls) {
      const sailed = ls.filter(l => !l.inChina);
      const res = { mode: 'ours', erpLots: [], erpInTrans: null, ourSailed: sailed.reduce((t, l) => t + l.units, 0),
        matched: 0, received: 0, taken: 0, landed: 0, unmatched: 0, kept: 0, capped: 0, maybeTwice: 0, takenInBacklog: 0, flags: [] };
      if (!transKnown || !e) return res;
      // An In Trans quantity that can't be read (text that isn't a number, or below zero; core/ats.js lists it): the
      // ERP's In Trans isn't known for the style, so Arriving stays our own sailed lots.
      const rows = Array.isArray(e.rows) ? e.rows : [];
      if (rows.some(x => (x.inTransBad || []).length || (x.inTransList || []).some(n => !(typeof n === 'number' && n >= 0)))) {
        res.flags.push('erp-in-trans-unknown'); counts.erpInTransUnknownStyles++;
        return res;
      }
      res.mode = 'erp'; counts.erpStyles++;
      const bad = [];
      for (const x of rows) {
        const q = x.inTransList || [];
        if (!q.length) continue;
        // The ETA Date list, part for part (an unreadable date stays in its place as null).
        const parts = Array.isArray(x.etaParts) ? x.etaParts : Array.isArray(x.eta) ? x.eta : [];
        const lined = parts.length === q.length;
        const unreadable = lined ? q.filter((n, i) => n > 0 && !isDay(parts[i])).length : 0;
        let noEta = 0;
        q.forEach((n, i) => {
          if (!(n > 0)) return;
          const eta = lined && isDay(parts[i]) ? parts[i] : null;
          if (eta == null) noEta += n;
          res.erpLots.push(erpLot(style, styleRaw, x, i, n, eta));
        });
        if (!lined || unreadable) bad.push({ row: x.row, q: q.length, d: parts.length, unreadable, units: noEta });
      }
      if (bad.length) { etaProblems.push({ style, styleRaw, rows: bad, units: bad.reduce((t, b) => t + b.units, 0) }); res.flags.push('erp-eta-mismatch'); counts.erpEtaMismatchStyles++; }
      counts.erpLots += res.erpLots.length; counts.erpLotsNoEta += res.erpLots.filter(l => l.eta == null).length;
      const E = res.erpInTrans = res.erpLots.reduce((t, l) => t + l.units, 0);
      if (E >= res.ourSailed) {
        // The ERP has all our sailed units In Trans: they're counted there (matched), not again. What it has beyond them
        // is goods we still see on a vessel in China (Booked / Loaded with an ETD; earliest ETD first, taken off those
        // lots), then landed but not received yet, then due later with nothing sailed on our data (nothing taken off).
        if (sailed.length) counts.erpMoreStyles++;
        for (const l of sailed) { l.matched = l.units; res.matched += l.units; l.flags.push('matched-erp'); counts.matchedLots++; }
        let x = E - res.ourSailed;
        const onVessel = ls.filter(l => l.inChina && ON_VESSEL.includes(l.bucket) && l.etd != null)
          .sort((a, b) => a.etd - b.etd || byPo(a.po, b.po) || String(a.id).localeCompare(String(b.id)));
        for (const l of onVessel) {
          if (x <= 0) break;
          const n = Math.min(l.units, x);
          l.inErpTransit = n; x -= n; res.taken += n;
          l.flags.push('in-erp-transit'); counts.inErpTransitLots++;
        }
        // What's left: landed, not received yet, as far as the ERP's shipments due on or before the report's day go
        // (beyond our own sailed lots due by then); the rest is due after it (or has no ETA) and nothing on our shipping
        // data has sailed for it.
        const due = list => (repDay == null ? 0 : sumU(list.filter(l => l.eta != null && l.eta <= repDay)));
        res.landed = Math.min(x, Math.max(0, due(res.erpLots) - due(sailed)));
        res.unmatched = x - res.landed;
        if (res.taken) { res.flags.push('in-erp-transit'); counts.inErpTransitStyles++; }
        if (res.landed) { res.flags.push('landed-not-received'); counts.landedNotReceivedStyles++; }
        if (res.unmatched) { res.flags.push('erp-in-trans-unmatched'); counts.erpUnmatchedStyles++; }
        // Cross-check (listed, never counted; Saul decides): that In Trans may be goods our Master Sheet still has in
        // China, when Container Tracking shows the style sailed under a PO with no Master Sheet line. Then the same
        // units are offered twice: as Arriving and from China.
        if (res.unmatched) {
          const held = sumU(ls.filter(l => l.inChina), l => l.units - l.inErpTransit);
          res.maybeTwice = Math.min(res.unmatched, ctSailedNoLineBy.get(style) || 0, held);
          if (res.maybeTwice > 0) { res.flags.push('maybe-twice'); counts.maybeTwiceStyles++; }
        }
        // Cross-check (listed, never counted; Saul decides): the report's Import PO Backlog is the style's goods the ERP
        // still has in China; taking more into its transit than our China goods (sold ones too) beyond that backlog
        // contradicts it.
        if (res.taken && backlogKnown) {
          const backlog = sumU(rows, r => sumU((r.backlogList || []).filter(n => typeof n === 'number' && n > 0), n => n));
          const china = sumU(ls.filter(l => l.inChina)) + (soldInChinaBy.get(style) || 0);
          res.takenInBacklog = Math.max(0, res.taken - Math.max(0, china - backlog));
          if (res.takenInBacklog > 0) { res.flags.push('taken-in-backlog'); counts.takenInBacklogStyles++; }
        }
      } else {
        // The ERP has less In Trans than we have sailed. A shipment and one of our sailed lots of the same quantity are the
        // same goods (the closest ETA first): matched. Of the rest, the difference, earliest ETA first, is either received
        // already (an ETA on or before the report's day: in its In WH, so dropped here, never more than the style's In
        // WH) or not on its In Trans yet (kept as Arriving from our data). The rest of our sailed units are its In Trans.
        counts.erpLessStyles++;
        let d = res.ourSailed - E;
        if (repDay == null) { res.flags.push('no-report-day'); counts.noReportDayStyles++; }
        const free = sailed.slice().sort(etaOrder);
        for (const s of res.erpLots.slice().sort(etaOrder)) {
          const gap = l => (s.eta == null || l.eta == null ? Infinity : Math.abs(l.eta - s.eta));
          const same = free.filter(l => l.units === s.units).sort((a, b) => gap(a) - gap(b) || etaOrder(a, b))[0];
          if (!same) continue;
          free.splice(free.indexOf(same), 1);
          same.matched = same.units; res.matched += same.units;
          same.flags.push('matched-erp'); counts.matchedLots++; counts.sameQtyMatchedLots++;
        }
        // Received units are in the style's In WH: never more than it (unknown In WH: no limit).
        let room = e && isDay(e.inWh) ? Math.max(0, e.inWh) : Infinity;
        for (const l of free) {
          const n = Math.min(l.units, d);
          d -= n;
          const passed = n > 0 && repDay != null && l.eta != null && l.eta <= repDay;
          const rec = passed ? Math.min(n, room) : 0;
          if (rec > 0) { l.received = rec; room -= rec; res.received += rec; l.flags.push('received'); counts.receivedLots++; }
          if (n - rec > 0) { l.notOnErp = n - rec; res.kept += n - rec; l.flags.push('not-on-erp-in-trans'); counts.keptExtraLots++; }
          if (passed && n - rec > 0) res.capped += n - rec;
          l.matched = l.units - n; res.matched += l.matched;
          if (l.matched) { l.flags.push('matched-erp'); counts.matchedLots++; }
        }
        if (res.received) { res.flags.push('received'); counts.receivedStyles++; }
        if (res.kept) { res.flags.push('not-on-erp-in-trans'); counts.keptExtraStyles++; }
        if (res.capped) { res.flags.push('received-over-in-wh'); counts.receivedCappedStyles++; }
      }
      for (const l of ls) l.available = l.units - replaced(l);
      return res;
    }

    // ---------- every style: on the stage engine's list, with a lot, or on the ATS report
    const stageStyle = new Map((stages.styles || []).map(s => [s.style, s]));
    const keys = new Set([...stageStyle.keys(), ...lotsBy.keys(), ...(atsOk ? ats.styles.keys() : [])]);
    const upcBy = upc && upc.byStyle instanceof Map ? upc.byStyle : null;
    const styles = [];
    const notOnReport = [], unknownStock = [], erpOnly = [];
    for (const style of [...keys].sort((a, b) => (a < b ? -1 : a > b ? 1 : 0))) {
      const st = stageStyle.get(style) || null, ours = lotsBy.get(style) || [], e = atsOk ? ats.styles.get(style) || null : null;
      const um = upcBy ? upcBy.get(style) || [] : [];
      const styleRaw = (st && st.styleRaw) || (ours[0] && ours[0].styleRaw) || (e && e.styleRaw) || style;
      // Arriving: the ERP's In Trans for a style on the report (our own sailed lots matched to it), else our sailed lots.
      const x = erpStep(style, styleRaw, e, ours);
      const ls = x.erpLots.length ? [...ours, ...x.erpLots].sort(lotOrder) : ours;
      for (const l of x.erpLots) lots.push(l);
      const flags = x.flags.slice();
      // No date to sell it from: only what's left of a lot (a sailed lot the ERP counts needs none).
      for (const l of ls) if (l.availableDate == null && l.available > 0) { l.flags.push('no-date'); if (l.source === 'ours') counts.noDateLots++; }
      const known = !!(e && isDay(e.inWh) && isDay(e.orders));
      let inWh = null, orders = null, stockNet = null, now = null, shortfall = 0;
      if (known) {
        inWh = e.inWh; orders = e.orders; stockNet = inWh - orders;
        now = Math.max(0, stockNet); shortfall = Math.max(0, -stockNet);
        counts.stylesOnReport++;
      } else if (e) {
        inWh = isDay(e.inWh) ? e.inWh : null; orders = isDay(e.orders) ? e.orders : null;
        flags.push('stock-unknown'); counts.stylesOnReport++; counts.stylesStockUnknown++;
        if (ls.length) unknownStock.push(style);
      } else if (atsOk) {
        flags.push('not-on-report'); counts.stylesNotOnReport++;
        if (ls.length) notOnReport.push(style);
      }
      // Customer orders beyond stock come off the lots (the ERP's In Trans by ETA, ours by their date), earliest available
      // date first, from what each lot has left.
      let left = shortfall, ordersTaken = 0;
      for (const l of ls) {
        if (left <= 0) break;
        if (!(l.available > 0)) continue;
        const n = Math.min(l.available, left);
        l.taken = n; l.available -= n; left -= n; ordersTaken += n;
        l.flags.push(l.available ? 'part-taken' : 'taken');
        counts.lotsTaken++; if (!l.available) counts.lotsFullyTaken++;
      }
      const oversold = left;
      const soldChina = soldBy.get(style) || 0, unmatched = unmatchedBy.get(style) || 0;
      // Orders beyond stock no more than the units sold to ship from China: those orders may be the same goods (listed,
      // never counted: an open question to Saul).
      let soldInOrders = 0;
      if (known && shortfall > 0 && soldChina > 0 && shortfall <= soldChina) {
        soldInOrders = shortfall; flags.push('orders-may-include-sold');
        counts.soldInOrdersStyles++; counts.soldInOrdersUnits += shortfall; if (oversold > 0) counts.soldInOrdersOversold++;
      }
      const open = ls.filter(l => l.available > 0);
      if (!known && open.length) { flags.push('orders-unknown'); for (const l of open) l.flags.push('orders-unknown'); counts.ordersUnknownStyles++; counts.ordersUnknownUnits += open.reduce((s, l) => s + l.available, 0); }
      if (oversold > 0) { flags.push('oversold'); counts.oversoldStyles++; }
      if (now > 0) counts.stylesWithStock++;
      if (e && !knownStyles.has(style) && !(um.length)) { flags.push('erp-only'); erpOnly.push(style); }
      if (ls.some(l => l.flags.includes('dates-disagree'))) flags.push('dates-disagree');
      if (ls.some(l => l.flags.includes('no-date'))) flags.push('no-date');

      // Case pack: our lots' Master, 'mixed' when they differ; with no lot, the latest Master Sheet line's. The ERP's
      // shipments carry none: they take the style's (as stock does).
      const masters = uniq(ours.map(l => l.master));
      const lm = latestMaster.get(style);
      const casePack = masters.length > 1 ? 'mixed' : masters.length ? masters[0] : lm ? lm.master : null;
      if (masters.length > 1) flags.push('mixed-case-pack');
      const stockPack = isDay(casePack) ? casePack : lm ? lm.master : null;
      for (const l of x.erpLots) l.master = stockPack;

      const sum = (pred, f) => ls.filter(pred).reduce((s, l) => s + f(l), 0);
      const fromChina = sum(l => l.inChina, l => l.available), arriving = sum(l => !l.inChina, l => l.available);
      const arrivingErp = sum(l => l.source === 'erp', l => l.available);
      for (const l of ls) { l.cases = casesOf(l.available, l.master); l.looseUnits = looseOf(l.available, l.master); }
      const caseSum = pred => sum(l => pred(l) && l.cases != null, l => l.cases);
      const noPackSum = pred => sum(l => pred(l) && l.cases == null, l => l.available);
      const nowCases = now == null ? 0 : casesOf(now, stockPack);
      const cases = { now: nowCases ?? 0, fromChina: caseSum(l => l.inChina), arriving: caseSum(l => !l.inChina), total: 0,
        soldChina: soldCasesBy.get(style) || 0, oversold: 0, ordersTaken: 0 };
      const noCasePackUnits = { now: now != null && nowCases == null ? now : 0, fromChina: noPackSum(l => l.inChina),
        arriving: noPackSum(l => !l.inChina), total: 0, soldChina: soldNoPackBy.get(style) || 0, oversold: 0, ordersTaken: 0 };
      cases.total = cases.now + cases.fromChina + cases.arriving;
      noCasePackUnits.total = noCasePackUnits.now + noCasePackUnits.fromChina + noCasePackUnits.arriving;
      // Units with a case pack that don't make a whole case (after customer orders): never offered as a case.
      const looseUnits = { now: now == null ? 0 : looseOf(now, stockPack), fromChina: sum(l => l.inChina, l => l.looseUnits),
        arriving: sum(l => !l.inChina, l => l.looseUnits), total: 0 };
      looseUnits.total = looseUnits.now + looseUnits.fromChina + looseUnits.arriving;

      const pulls = ls.filter(l => l.inChina && l.available > 0 && l.pullFrom != null);
      const first = pulls.reduce((m, l) => (!m || l.pullFrom < m.pullFrom ? l : m), null);
      const arr = ls.filter(l => !l.inChina && l.available > 0 && l.arrival != null);
      const umDesc = um.find(u => u.desc), umCat = um.find(u => u.category);
      const upcs = uniq([...(st ? st.upcs.map(u => u.upc) : []), ...ls.map(l => l.upc), ...um.map(u => u.upc)]);
      styles.push({
        style, styleRaw,
        desc: (st && st.desc) || (umDesc && umDesc.desc) || (e && e.desc) || '',
        category: (st && st.category) || (umCat && umCat.category) || (e && e.category) || '',
        factory: (st && st.factory) || '', factories: st ? st.factories.slice() : [],
        upcs, casePack, onReport: atsOk ? !!e : null,
        inWh, orders, stockNet, now, fromChina, arriving, total: (now ?? 0) + fromChina + arriving,
        soldChina, unmatched, oversold, ordersTaken, cases, noCasePackUnits, looseUnits, soldInOrders,
        // Where Arriving comes from ('erp': the report's In Trans; 'ours': our sailed lots), split after customer orders,
        // and how our own lots meet the ERP's In Trans (units, before orders): every one of ours (ourIncoming, the stage
        // engine's Incoming for the style) is in a lot, matched to the ERP's In Trans, received, or in its transit.
        arrivingFrom: x.mode, arrivingErp, arrivingOurs: arriving - arrivingErp,
        erpInTrans: x.erpInTrans, erpShipments: x.erpLots.length, ourIncoming: ours.reduce((t, l) => t + l.units, 0), ourSailed: x.ourSailed,
        matchedOurs: x.matched, receivedDropped: x.received, inErpTransitTaken: x.taken, landedNotReceived: x.landed, keptExtra: x.kept,
        erpUnmatched: x.unmatched, receivedCapped: x.capped,
        // Cross-checks with the ERP's report (units; listed on Problems, never counted): maybe offered twice (In Trans
        // our data doesn't show sailed, while Container Tracking has the style sailed with no Master Sheet line and the
        // sheet still has it in China), and taken into the ERP's transit though its Import PO Backlog has it in China.
        maybeTwice: x.maybeTwice, takenInBacklog: x.takenInBacklog,
        firstPull: first ? first.pullFrom : null, firstPullRule: first ? first.pullRule : null,
        firstPullEstimate: !!(first && first.estimate), firstPullNow: !!(first && first.pullNow),
        firstArrival: arr.length ? Math.min(...arr.map(l => l.arrival)) : null,
        lots: ls, flags,
      });
    }
    counts.styles = styles.length;
    // Every lot, ours and the ERP's In Trans, in date order (lotsInChina / lotsSailed count ours, erpLots the ERP's).
    lots.sort(lotOrder);
    counts.lots = lots.length;
    // The cross-check notes on a style with orders beyond its stock, in words.
    function checkNotes(s) {
      const n = [];
      if (s.soldInOrders > 0) n.push(`${unitsText(s.soldChina)} of this style ${s.soldChina === 1 ? 'is' : 'are'} sold to ship from China: if those orders are in the ERP's Total Customer Orders, they're taken off twice.`);
      return n;
    }

    // ---------- issues
    // Ready dates that disagree: one line per PO + style still in China, every source's date named.
    const pairs = new Map();
    for (const l of lots) if (l.flags.includes('dates-disagree')) {
      const k = l.po + '|' + l.style;
      const p = pairs.get(k) || pairs.set(k, { po: l.po, style: l.style, styleRaw: l.styleRaw, units: 0, src: new Map(), used: new Set(), spread: 0 }).get(k);
      p.units += l.units; p.spread = Math.max(p.spread, l.readySpread);
      for (const s of l.readySources) p.src.set(s.key + '|' + s.day, s);
      p.used.add(l.readyFrom + '|' + l.readyDate);
    }
    counts.datesDisagreePairs = pairs.size;
    const byPair = (a, b) => byPo(a.po, b.po) || (a.style < b.style ? -1 : a.style > b.style ? 1 : 0);
    for (const p of [...pairs.values()].sort(byPair)) {
      const srcs = [...p.src.values()].sort((a, b) => READY_ORDER.indexOf(a.key) - READY_ORDER.indexOf(b.key) || a.day - b.day);
      const used = [...p.used].map(x => { const [k, d] = x.split('|'); return `${fmt(+d)} (${READY_LABEL[k]})`; });
      const far = srcs.filter(s => s.far && !p.used.has(s.key + '|' + s.day));
      const farTxt = far.length ? ` ${far.map(s => `${READY_LABEL[s.key]} ${fmt(s.day)}`).join(' and ')} ${far.length > 1 ? 'are' : 'is'} more than a year ahead or before 2015 (a wrong year?), so ${far.length > 1 ? "they aren't" : "it isn't"} used.` : '';
      issue('dates-disagree', p.po, p.style, `PO ${p.po}, style ${p.styleRaw}: the ready dates are up to ${fmtN(p.spread)} days apart: ${srcs.map(s => `${READY_LABEL[s.key]} ${fmt(s.day)} (${READY_FILE[s.key]})`).join(', ')}.${farTxt} The ready date used is ${used.join(' and ')}; correct the sheet that's out of date.`,
        { sources: srcs.map(s => ({ key: s.key, label: READY_LABEL[s.key], file: READY_FILE[s.key], day: s.day })) });
    }
    // Our lots with no date to sell them from: one line per PO + style (the ERP's shipments with no ETA are listed
    // below, by the report's row).
    const noDate = new Map();
    for (const l of lots) if (l.source === 'ours' && l.flags.includes('no-date')) {
      const k = l.po + '|' + l.style + '|' + (l.inChina ? 'c' : 's');
      const p = noDate.get(k) || noDate.set(k, { po: l.po, style: l.style, styleRaw: l.styleRaw, inChina: l.inChina, units: 0, statuses: new Set() }).get(k);
      p.units += l.units - replaced(l); p.statuses.add(l.status);
    }
    for (const p of [...noDate.values()].sort(byPair)) {
      issue('no-date', p.po, p.style, p.inChina
        ? `PO ${p.po}, style ${p.styleRaw}: ${unitsText(p.units)} still in China (${[...p.statuses].join(', ')}) have no ETD, no ready date (CRD) and no order date, so there's no date to sell them from. Add a CRD or order date.`
        : `PO ${p.po}, style ${p.styleRaw}: ${unitsText(p.units)} ${[...p.statuses].join(', ').toLowerCase()} have no ETA, so there's no arrival date to sell them from.`);
    }
    for (const s of styles) if (s.oversold > 0) {
      const notes = checkNotes(s);
      issue('oversold', '', s.style, `Style ${s.styleRaw}: customer orders (${fmtN(s.orders)}) are ${unitsText(s.oversold)} more than the warehouse stock (${fmtN(s.inWh)})${s.lots.length ? ` and every open lot from China (${unitsText(s.ordersTaken)})` : ' with no open lot from China'}.${notes.length ? ` Check before turning a buyer away: ${notes.join(' ')}` : ' Nothing is available to sell.'}`,
        { units: s.oversold });
    }
    // The ERP's In Trans quantities whose ETA can't be told (the ETA Date list doesn't line up with In Trans, or a date
    // can't be read): one line per style, naming the report's rows.
    for (const p of etaProblems) {
      const rows = p.rows.map(b => (b.q !== b.d ? `row ${b.row}: ${fmtN(b.q)} In Trans quantit${b.q === 1 ? 'y' : 'ies'}, ${fmtN(b.d)} ETA date${b.d === 1 ? '' : 's'}`
        : `row ${b.row}: ${fmtN(b.unreadable)} ETA date${b.unreadable === 1 ? '' : 's'} that can't be read`));
      issue('erp-eta-mismatch', '', p.style, `Style ${p.styleRaw}: the ATS report's ETA Date list doesn't line up with its In Trans Qty list (${rows.join('; ')}), so ${unitsText(p.units)} In Trans ${p.units === 1 ? 'has' : 'have'} no arrival date. They're still counted as Arriving, with no date; check the ERP's ETA dates.`,
        { units: p.units, rows: p.rows.map(b => b.row) });
    }
    // Without the report's In Trans Qty column, Arriving is our own sailed lots for every style.
    if (counts.noInTransColumn) {
      issue('no-in-trans-column', '', '', `The ATS report has no "In Trans Qty" column, so Arriving comes from our shipping data (on the water, at port) for every style, and goods the ERP has already received or counts in transit can't be told apart.`);
    }
    // Styles whose orders beyond stock may be covered by goods not counted here: one summary line each.
    const summary = (code, unitsOf, text) => {
      const list = styles.filter(s => s.flags.includes(code));
      if (!list.length) return;
      const u = list.reduce((t, s) => t + unitsOf(s), 0), ov = list.filter(s => s.oversold > 0).length;
      issue(code, '', '', `${fmtN(list.length)} style${list.length === 1 ? '' : 's'} (${fmtN(ov)} of them oversold): ${text(u)} For example ${list.slice().sort((a, b) => unitsOf(b) - unitsOf(a)).slice(0, 5).map(s => `${s.styleRaw} ${fmtN(unitsOf(s))}`).join(', ')}${list.length > 5 ? '…' : ''}.`,
        { count: list.length, units: u });
    };
    summary('orders-may-include-sold', s => s.soldInOrders,
      u => `the customer orders beyond the warehouse stock (${unitsText(u)}) are no more than the style's goods sold to ship from China (China Pickup or straight to the customer). If the ERP's Total Customer Orders include those orders, they are taken off twice.`);
    // How the ERP's In Trans met our own data: what it can't be matched to, and the cross-checks (never counted).
    summary('maybe-twice', s => s.maybeTwice,
      u => `${unitsText(u)} may be offered twice. The ATS report has them In Trans, due after its date, beyond what our shipping data shows on the water, at port, booked or loaded; Container Tracking shows the style on the water or at port under a PO with no Master Sheet Live line (a PO written differently?), and the Master Sheet still has goods of the style in China. Both are shown (Arriving, and from China or taken by orders) until Saul decides how these count: check before selling them.`);
    summary('erp-in-trans-unmatched', s => s.erpUnmatched,
      u => `${unitsText(u)} In Trans on the ATS report, due after its date (or with no ETA), are more than our shipping data has on the water, at port, booked or loaded for them. They're offered as Arriving on the report's word; check Container Tracking and the Master Sheet for them.`);
    summary('taken-in-backlog', s => s.takenInBacklog,
      u => `${unitsText(u)} on Booked or Loaded lots were taken off Available from China because the ATS report has more In Trans than our sailed goods, but its Import PO Backlog still holds the style's goods in China, so they may still be there (From China may be short by that much, and a style called oversold). Saul decides how these count.`);
    summary('received-over-in-wh', s => s.receivedCapped,
      u => `${unitsText(u)} on the water or at port have an ETA on or before the ATS report's date and aren't on its In Trans, but the style's In WH is less, so they can't all be in its warehouse number: the rest are kept as Arriving from our shipping data. Check whether they were received.`);
    summary('no-report-day', s => s.ourSailed - (s.erpInTrans || 0),
      u => `the ATS report's date isn't known, so none of the ${unitsText(u)} on the water or at port beyond its In Trans is called received: they're all kept as Arriving.`);
    if (atsMissing) {
      const n = counts.ordersUnknownStyles;
      if (n) issue('orders-unknown', '', '', `${ats && ats.ok === false ? "The ATS report couldn't be used" : "There's no ATS report"}, so customer orders aren't taken off: the ${fmtN(n)} style${n === 1 ? '' : 's'} with lots from China show${n === 1 ? 's' : ''} ${unitsText(counts.ordersUnknownUnits)} in full, and stock now isn't known.`,
        { count: n });
    } else if (notOnReport.length || unknownStock.length) {
      const ex = [...notOnReport, ...unknownStock].slice(0, 5).map(k => (stageStyle.get(k) || {}).styleRaw || k).join(', ');
      const parts = [notOnReport.length ? `${fmtN(notOnReport.length)} style${notOnReport.length === 1 ? ' is' : 's are'} not on the ATS report` : '',
        unknownStock.length ? `${fmtN(unknownStock.length)} ${unknownStock.length === 1 ? 'has' : 'have'} a stock figure on it that can't be read` : ''].filter(Boolean);
      issue('orders-unknown', '', '', `${parts.join(' and ')}, so customer orders aren't taken off their lots from China (${unitsText(counts.ordersUnknownUnits)} shown in full): for example ${ex}${notOnReport.length + unknownStock.length > 5 ? '…' : ''}.`,
        { count: notOnReport.length + unknownStock.length });
    }
    counts.erpOnlyStyles = erpOnly.length;
    if (erpOnly.length) {
      const withStock = styles.filter(s => s.flags.includes('erp-only') && (s.inWh || s.orders)).length;
      issue('erp-only', '', '', `${fmtN(erpOnly.length)} style${erpOnly.length === 1 ? ' is' : 's are'} on the ATS report but on no Master Sheet Live or Container Tracking line and not in UPC_MASTER (${fmtN(withStock)} with stock or orders), so they show with the report's numbers only (stock, and any In Trans): for example ${erpOnly.slice(0, 5).map(k => ats.styles.get(k).styleRaw).join(', ')}${erpOnly.length > 5 ? '…' : ''}.`,
        { count: erpOnly.length });
    }

    // ---------- every unit accounted for, style by style (a failure is a bug in this file, said so on Problems):
    // each lot's units = matched + received + in the ERP's transit + taken + available, none below 0; our lots hold the
    // stage engine's Incoming for the style and the ERP's lots its In Trans; the ERP's In Trans = our matched units +
    // those in its transit + landed, not received + due later, not sailed on our data; received ≤ In WH; so from China +
    // arriving + taken by orders + matched + received + in the ERP's transit = Incoming + the ERP's In Trans.
    const offBy = [];
    for (const s of styles) {
      const bad = s.lots.filter(l => [l.matched, l.received, l.inErpTransit, l.taken, l.available].some(v => !(v >= 0)) ||
        Math.abs(l.units - replaced(l) - l.taken - l.available) > 1e-9);
      const st = stageStyle.get(s.style), u = (st && st.units) || null;
      const inc = u && isDay(u.incoming) ? u.incoming : null;
      const lotU = s.lots.reduce((t, l) => t + l.units, 0);
      const why = [bad.length ? `${bad.length} lot${bad.length === 1 ? '' : 's'} don't add up` : '',
        inc != null && Math.abs(inc - s.ourIncoming) > 1e-9 ? `its lots hold ${fmtN(s.ourIncoming)}, Incoming is ${fmtN(inc)}` : '',
        s.arrivingFrom === 'erp' && Math.abs(s.erpInTrans - s.matchedOurs - s.inErpTransitTaken - s.landedNotReceived - s.erpUnmatched) > 1e-9 ? "the ERP's In Trans isn't matched + in its transit + landed + not sailed on our data" : '',
        isDay(s.inWh) && s.receivedDropped > Math.max(0, s.inWh) ? `received ${fmtN(s.receivedDropped)} is more than In WH ${fmtN(s.inWh)}` : '',
        Math.abs(lotU - s.ourIncoming - (s.erpInTrans || 0)) > 1e-9 ? "the lots aren't Incoming + the ERP's In Trans" : '',
        Math.abs(lotU - s.lots.reduce((t, l) => t + replaced(l), 0) - s.fromChina - s.arriving - s.ordersTaken) > 1e-9 ? "the lots aren't from China + arriving + taken" : ''].filter(Boolean);
      if (why.length) offBy.push(`${s.styleRaw} (${why.join('; ')})`);
    }
    if (offBy.length) {
      issue('balance', '', '', `Available to sell doesn't add up on ${fmtN(offBy.length)} style${offBy.length === 1 ? '' : 's'}: ${offBy.slice(0, 5).join(', ')}${offBy.length > 5 ? '…' : ''}. This is a bug in Shalam Tool, not in the files: its numbers can't be trusted until it's fixed.`);
    }

    // ---------- totals (units, and cases where the Master is known)
    const totals = zero(), tCases = zero(CASE_KEYS), tNoPack = zero(CASE_KEYS), tLoose = { now: 0, fromChina: 0, arriving: 0, total: 0 };
    for (const s of styles) {
      totals.now += s.now || 0; totals.fromChina += s.fromChina; totals.arriving += s.arriving; totals.soldChina += s.soldChina;
      totals.oversold += s.oversold; totals.ordersTaken += s.ordersTaken;
      for (const k of ['arrivingErp', 'arrivingOurs', 'erpInTrans', 'ourIncoming', 'matchedOurs', 'receivedDropped', 'inErpTransitTaken',
        'landedNotReceived', 'erpUnmatched', 'keptExtra', 'receivedCapped', 'maybeTwice', 'takenInBacklog']) totals[k] += s[k] || 0;
      for (const k of ['now', 'fromChina', 'arriving', 'soldChina']) { tCases[k] += s.cases[k]; tNoPack[k] += s.noCasePackUnits[k]; }
      for (const k of ['now', 'fromChina', 'arriving', 'total']) tLoose[k] += s.looseUnits[k];
    }
    totals.total = totals.now + totals.fromChina + totals.arriving;
    tCases.total = tCases.now + tCases.fromChina + tCases.arriving;
    tNoPack.total = tNoPack.now + tNoPack.fromChina + tNoPack.arriving;
    totals.cases = tCases; totals.noCasePackUnits = tNoPack; totals.looseUnits = tLoose;

    return { ok: true, today, atsMissing, atsError: ats && ats.ok === false ? ats.error || '' : '',
      reportDay: repDay, totals, styles, lots, issues, counts };
  }

  return { buildSell, readyOf, pullOf, ruleWords, SOURCE, STATUS, IN_CHINA, SAILED, READY_ORDER, READY_LABEL, READY_FILE,
    RULES, TOTAL_KEYS, DAYS_BEFORE_ETD, DAYS_BEFORE_CRD, DAYS_AFTER_ORDER, DISAGREE_DAYS };
});
