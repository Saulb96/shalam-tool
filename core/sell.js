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
//   lot (the same snapshot as the stock). Its In Trans is the truth for what has left China and not been received:
//   nothing is offered twice (as Arriving and again from China) and no unit vanishes. Our own on-the-water / at-port
//   lots are matched to its shipments first (one of the same quantity, then the closest ETA). When the ERP has as much or
//   more, ours are counted there (not again), and the rest of its In Trans is placed in turn:
//   1. landed, not received yet: our Landed goods of the style with an ETA (else ETD) from 60 days before the report's
//      day up to today take it up to their units (on its shipments due by that day, or due no more than 14 days after
//      their own date: one due later has left China per the ERP), then its shipments due on or before the report's day
//      (nothing comes off China for these). With an Import PO Backlog, only what's beyond our China goods the backlog no
//      longer has (they have left China per the ERP: step 2 takes them first) is called landed;
//   2. left China per the ERP (shipments due after the report's day, or with no ETA, and what step 1 left): taken off our
//      lots still in China, Booked / Loaded with an ETD first (earliest ETD first), then lots on a PO whose 6-digit base
//      is one Container Tracking shows sailed or landed for the style (or the shipment's own PO, when the report has
//      one), then any other (earliest pull-from date first); never taking our China goods below the report's Import PO
//      Backlog for the style (sold goods count until they're loaded: its backlog has sold goods booked or ready, never
//      those already loaded);
//   3. what's still unplaced stays Arriving and is listed, one line per style.
//   When the ERP has less, a shipment and one of ours of the same quantity are matched first; of the rest, ours with an
//   ETA on or before the report's day are received (in its In WH, dropped; never more than the style's In WH), and what's
//   left is kept as Arriving from our data (not on the ERP's In Trans yet). Then, as far as our China goods are beyond
//   its backlog, its In Trans due after the report's day matched to ours with an ETA on or before that day is taken off
//   China instead and those of ours received. A style not on the report, or no report: our own sailed lots, as before.
//   An Import PO Backlog part below zero or that isn't a number makes the style's backlog unknown (no limit, listed).
//   Every unit is accounted for: each lot's units = matched + received + in the ERP's transit + taken by orders +
//   available; the ERP's In Trans = matched + landed + taken off China + unplaced.
// - Guards (a failure is a bug, said so on Problems): nothing is In Trans landed or unplaced while our goods in China
//   beyond the report's Import PO Backlog are still offered (offered twice; with no backlog, unplaced while any are), and
//   nothing comes off China below that backlog.
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
  // The order our lots still in China come off when the ERP counts goods of the style in transit beyond ours: on a vessel,
  // on a PO that has sailed, any other.
  const TIERS = ['vessel', 'po', 'other'];
  const DAYS_BEFORE_ETD = 10, DAYS_BEFORE_CRD = 5, DAYS_AFTER_ORDER = 75, DISAGREE_DAYS = 7;
  // Our Landed goods this recent (ETA, else ETD, from this many days before the report's day up to today) may be what
  // the ERP still has In Trans: landed, not received yet.
  const LANDED_DAYS = 60;
  // ... on the ERP's shipments due by the report's day, or due after it but no more than this many days after their date.
  const LANDED_LATE_DAYS = 14;
  // Container Tracking's stages for goods that have left China.
  const CT_GONE = ['On the water', 'At port', 'Landed'];
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
  // The ERP's In Trans against our lots, per style and in total (units): matched, landed (ours, or due by the report's
  // day; landedLater: ours on shipments due after it, within 14 days of their date), taken off China (by tier;
  // takenDueBy: on shipments due by the report's day, as its backlog no longer has them in China), unplaced
  // (unplacedInChina: while our sheets still show the style in China); received / kept when it has less; offeredTwice is
  // the guard (always 0).
  const ERP_KEYS = ['erpInTrans', 'ourIncoming', 'matchedOurs', 'receivedDropped', 'inErpTransitTaken', 'landedNotReceived',
    'landedOurs', 'landedDue', 'landedLater', 'erpUnplaced', 'unplacedInChina', 'keptExtra', 'receivedCapped',
    'takenOnVessel', 'takenSamePo', 'takenOther', 'takenDueBy', 'offeredTwice'];
  const TOTAL_KEYS = [...CASE_KEYS, 'arrivingErp', 'arrivingOurs', ...ERP_KEYS];
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
  // A PO's 6-digit base ("900601" of "900601A" or "900437-BCF"), '' when it has none.
  const base6 = po => { const m = String(po ?? '').match(/^(\d{6})/); return m ? m[1] : ''; };

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
      inErpTransitTierLots: Object.fromEntries(TIERS.map(k => [k, 0])), landedNotReceivedStyles: 0, landedOursStyles: 0, landedDueStyles: 0,
      landedLaterStyles: 0, keptExtraStyles: 0, keptExtraLots: 0, noReportDayStyles: 0, erpUnplacedStyles: 0, unplacedInChinaStyles: 0,
      backlogGuardStyles: 0, sameQtyMatchedLots: 0, receivedCappedStyles: 0, offeredTwiceStyles: 0,
      backlogUnknownStyles: 0, takenDueByStyles: 0, lessTakenStyles: 0,
      soldInOrdersStyles: 0, soldInOrdersUnits: 0, soldInOrdersOversold: 0,
    };

    // ---------- lots, sold and unmatched units, from the allocations
    const lots = [], soldBy = new Map(), soldCasesBy = new Map(), soldNoPackBy = new Map(), unmatchedBy = new Map();
    const latestMaster = new Map();   // style -> { row, master } of its latest Master Sheet line with a Master
    const knownStyles = new Set();    // every style on any Master Sheet Live or Container Tracking line
    // For placing the ERP's In Trans beyond our sailed goods: our Landed goods (not sold: { po, units, day, dayFrom }),
    // the PO bases Container Tracking shows sailed or landed, and sold units still in China before they're loaded (the
    // report's Import PO Backlog has sold goods booked or ready, never those already loaded).
    const landedBy = new Map(), goneBasesBy = new Map(), soldInChinaBy = new Map();
    const add = (m, k, n) => m.set(k, (m.get(k) || 0) + n);
    const soldA = a => a.ptype === 'China Pickup' || (a.flags || []).includes('pull-section');
    for (const a of stages.allocs) {
      knownStyles.add(a.style);
      if (a.sheetRow != null && isDay(a.master) && a.master > 0) {
        const p = latestMaster.get(a.style);
        if (!p || a.sheetRow > p.row) latestMaster.set(a.style, { row: a.sheetRow, master: a.master });
      }
      // Container Tracking's own stage when the allocation has one (a line with no Container Tracking units has '').
      if (a.units > 0 && CT_GONE.includes(a.ctStage || a.stage) && base6(a.po)) (goneBasesBy.get(a.style) || goneBasesBy.set(a.style, new Set()).get(a.style)).add(base6(a.po));
      if (a.bucket === 'landed' && a.units > 0 && !soldA(a)) {
        const day = isDay(a.eta) ? a.eta : isDay(a.delivered) ? a.delivered : isDay(a.etd) ? a.etd : null;
        (landedBy.get(a.style) || landedBy.set(a.style, []).get(a.style)).push({ po: a.po, units: a.units, day,
          dayFrom: isDay(a.eta) ? 'eta' : isDay(a.delivered) ? 'delivered' : isDay(a.etd) ? 'etd' : null });
      }
      const lotBucket = IN_CHINA.includes(a.bucket) || SAILED.includes(a.bucket);
      const sold = a.bucket === 'pullFromChina' || (lotBucket && soldA(a));
      if (sold) {
        counts.soldAllocs++;
        add(soldBy, a.style, a.units);
        if ((IN_CHINA.includes(a.bucket) || (a.bucket === 'pullFromChina' && !GONE_STAGES.includes(a.stage))) && a.bucket !== 'loaded' && a.stage !== 'Loaded') add(soldInChinaBy, a.style, a.units);
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
        matched: 0, received: 0, inErpTransit: 0, erpTier: null, notOnErp: 0, taken: 0, available: a.units,
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
    // Its `place` says what the shipment is to our data (units): our goods on the water or at port (matched), landed on
    // our sheets (landedOurs) or due by the report's day (landedDue), taken off our lots in China (offChina), unplaced.
    function erpLot(style, styleRaw, x, i, n, eta) {
      const po = Array.isArray(x.poParts) && x.poParts.length === (x.inTransList || []).length ? S(x.poParts[i]) : '';
      return { id: `erp|${style}|${x.row}|${i + 1}`, allocId: null, style, styleRaw, upc: null, po: '', poShown: '', erpPo: po,
        status: STATUS.erp, bucket: 'erp', source: 'erp', inChina: false, units: n,
        matched: 0, received: 0, inErpTransit: 0, erpTier: null, notOnErp: 0, taken: 0, available: n, master: null, cases: null,
        place: { matched: 0, landedOurs: 0, landedDue: 0, offChina: 0, unplaced: 0 },
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
    // / inErpTransit (and erpTier) / notOnErp and available, and each of the ERP's shipments' place; answers { mode
    // ('erp' | 'ours'), erpLots, erpInTrans, ourSailed, matched, received, kept, capped (received beyond In WH, kept
    // instead), landed (= landedOurs + landedDue; landedLots: our Landed goods it took; landedLater: of those, on
    // shipments due after the report's day), taken (off our lots in China; takenBy tier; takenDueBy: on shipments due by
    // the report's day), unplaced (unplacedInChina of it while our sheets still show the style in China), backlog (the
    // report's Import PO Backlog, null without the column or with a part below zero or not a number: backlogOff lists
    // those parts), takeable (what may come off China under it, null with no backlog), chinaAll, offeredTwice (the
    // guard: 0), flags }.
    function erpStep(style, styleRaw, e, ls) {
      const sailed = ls.filter(l => !l.inChina);
      const res = { mode: 'ours', erpLots: [], erpInTrans: null, ourSailed: sailed.reduce((t, l) => t + l.units, 0),
        matched: 0, received: 0, kept: 0, capped: 0, landed: 0, landedOurs: 0, landedDue: 0, landedLater: 0, landedLots: [],
        taken: 0, takenBy: Object.fromEntries(TIERS.map(k => [k, 0])), takenDueBy: 0, unplaced: 0, unplacedInChina: 0,
        backlog: null, backlogOff: null, takeable: null, chinaAll: null, offeredTwice: 0, flags: [] };
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
      // The report's Import PO Backlog: the style's goods the ERP still has in China. A part below zero (a correction?) or
      // one that isn't a number (core/ats.js lists both) makes it unknown for the style: no limit from it, never 0.
      if (backlogKnown) {
        const parts = rows.flatMap(r => r.backlogList || []);
        const off = [...parts.filter(n => !(typeof n === 'number' && n >= 0)), ...rows.flatMap(r => r.backlogBad || [])];
        if (off.length) { res.backlogOff = off; res.flags.push('erp-backlog-unknown'); counts.backlogUnknownStyles++; }
        else res.backlog = sumU(parts, n => n);
      }
      // Our goods still in China, and how many of them the backlog no longer has there (takeable: they've left China per
      // the ERP; with no backlog, no limit). Sold goods count until they're loaded: the report's backlog has sold goods
      // booked or ready, never those already loaded (China Pickup or straight to the customer).
      const china = ls.filter(l => l.inChina);
      res.chinaAll = sumU(china) + (soldInChinaBy.get(style) || 0);
      if (res.backlog != null) res.takeable = Math.max(0, res.chinaAll - res.backlog);
      const offMax = Math.min(sumU(china), res.takeable == null ? Infinity : res.takeable);
      const dueBy = s => repDay != null && s.eta != null && s.eta <= repDay;
      // n units off our lots still in China: on a vessel (Booked / Loaded with an ETD, earliest ETD first), then on a PO
      // whose 6-digit base is in `bases` (Container Tracking shows it sailed or landed for the style, or it's one of the
      // ERP's shipments' own PO), then any other (earliest pull-from date first); put(units, lot) places each part on the
      // ERP's shipments. Answers the units taken. Called once per style at most.
      const takeOffChina = (n, bases, put) => {
        const pullOrder = (a, b) => (a.pullFrom == null) - (b.pullFrom == null) || (a.pullFrom ?? 0) - (b.pullFrom ?? 0) || byPo(a.po, b.po) || String(a.id).localeCompare(String(b.id));
        const vessel = china.filter(l => ON_VESSEL.includes(l.bucket) && l.etd != null).sort((a, b) => a.etd - b.etd || byPo(a.po, b.po) || String(a.id).localeCompare(String(b.id)));
        const samePo = china.filter(l => !vessel.includes(l) && bases.has(base6(l.po))).sort(pullOrder);
        const other = china.filter(l => !vessel.includes(l) && !samePo.includes(l)).sort(pullOrder);
        let k = n;
        for (const [tier, list] of [['vessel', vessel], ['po', samePo], ['other', other]]) {
          for (const l of list) {
            if (k <= 0) break;
            const m = Math.min(l.units, k);
            l.inErpTransit = m; l.erpTier = tier; k -= m; res.taken += m; res.takenBy[tier] += m;
            l.flags.push('in-erp-transit'); counts.inErpTransitLots++; counts.inErpTransitTierLots[tier]++;
            put(m, l);
          }
        }
        if (res.taken) { res.flags.push('in-erp-transit'); counts.inErpTransitStyles++; }
        return n - k;
      };
      if (E >= res.ourSailed) {
        // The ERP has all our sailed units In Trans: they're counted there (matched), not again. The rest of its In Trans
        // (x) is placed on its shipments in turn: landed, not received yet; left China (off our lots there); unplaced.
        if (sailed.length) counts.erpMoreStyles++;
        const ships = res.erpLots, left = new Map(ships.map(s => [s, s.units]));
        // A shipment our Landed goods dated `day` may be: due by the report's day, or due after it but within 14 days of
        // their date (the same trip, the ERP's ETA behind). One due later than that has left China per the ERP (its
        // Import PO Backlog agrees on the real files), so it's never called landed.
        const landedFits = day => s => dueBy(s) || (s.eta != null && day != null && s.eta <= day + LANDED_LATE_DAYS);
        // n units of our goods dated `day` placed on the shipments still unplaced (only those `fits` allows; with dueFirst,
        // those due by the report's day before any other): one of the same quantity first, then the closest ETA, then the
        // earliest.
        const place = (n, day, key, dueFirst, fits = () => true) => {
          const gap = s => (s.eta == null || day == null ? Infinity : Math.abs(s.eta - day));
          const rank = (a, b) => gap(a) - gap(b) || etaOrder(a, b);
          const open = ships.filter(s => left.get(s) > 0 && fits(s));
          let k = n;
          for (const group of dueFirst ? [open.filter(dueBy), open.filter(s => !dueBy(s))] : [open]) {
            const same = group.filter(s => s.units === k && left.get(s) === k).sort(rank)[0];
            for (const s of [...(same ? [same] : []), ...group.filter(s => s !== same).sort(rank)]) {
              if (k <= 0) break;
              const m = Math.min(left.get(s), k);
              left.set(s, left.get(s) - m); s.place[key] += m; k -= m;
              if (key === 'landedOurs' && !dueBy(s)) res.landedLater += m;
            }
          }
        };
        for (const l of sailed.slice().sort(etaOrder)) {
          l.matched = l.units; res.matched += l.units; l.flags.push('matched-erp'); counts.matchedLots++;
          place(l.units, l.eta, 'matched', false);
        }
        let x = E - res.ourSailed;
        // 1. Landed, not received yet: our Landed goods of the style dated from 60 days before the report's day up to
        // today (the latest first) take it, up to their units, on the shipments they can be (above); then the shipments
        // due on or before the report's day. Nothing comes off China for these. With no report day, nothing is landed.
        // With a backlog, the In Trans it says has left China (our China goods beyond it, as far as they go) comes off
        // China in step 2 first: only what's beyond that may be called landed (room), so nothing is offered twice.
        let room = res.backlog == null ? Infinity : Math.max(0, x - offMax);
        if (repDay == null) { if (x > 0) { res.flags.push('no-report-day'); counts.noReportDayStyles++; } }
        else {
          const from = repDay - LANDED_DAYS, to = Math.max(repDay, today);
          const recent = (landedBy.get(style) || []).filter(a => a.day != null && a.day >= from && a.day <= to)
            .sort((a, b) => b.day - a.day || byPo(a.po, b.po));
          for (const a of recent) {
            const fits = landedFits(a.day), n = Math.min(a.units, room, sumU(ships.filter(s => fits(s)), s => left.get(s)));
            if (n <= 0) continue;
            place(n, a.day, 'landedOurs', true, fits); x -= n; room -= n; res.landedOurs += n;
            res.landedLots.push({ po: a.po, units: n, day: a.day, dayFrom: a.dayFrom });
          }
          for (const s of ships) if (dueBy(s) && left.get(s) > 0 && room > 0) { const m = Math.min(left.get(s), room); s.place.landedDue += m; left.set(s, left.get(s) - m); x -= m; room -= m; res.landedDue += m; }
        }
        res.landed = res.landedOurs + res.landedDue;
        // 2. Left China per the ERP (due after the report's day, or no ETA, and what step 1 left): off our lots still in
        // China (above). Never below the report's Import PO Backlog: only our China goods beyond it may come off.
        const bases = new Set(goneBasesBy.get(style) || []);
        for (const s of ships) if (left.get(s) > 0 && base6(s.erpPo)) bases.add(base6(s.erpPo));
        x -= takeOffChina(Math.min(x, offMax), bases, (m, l) => place(m, l.eta, 'offChina', false));
        res.takenDueBy = sumU(ships.filter(dueBy), s => s.place.offChina);
        // 3. Still unplaced: Arriving on the ERP's word, and listed (some our sheets may still show in China).
        res.unplaced = x;
        for (const s of ships) s.place.unplaced = left.get(s);
        const held = sumU(china, l => l.units - l.inErpTransit), sold = soldInChinaBy.get(style) || 0;
        res.unplacedInChina = Math.min(res.unplaced, held);
        // The guard: In Trans called landed or left unplaced while our China goods the backlog no longer has are still
        // offered would be offered twice (with no backlog: unplaced while any of ours are). The rule never leaves any; the
        // balance check says so if it ever does.
        res.offeredTwice = res.backlog == null ? Math.min(res.unplaced, held)
          : Math.min(res.landed + res.unplaced, held, Math.max(0, held + sold - res.backlog));
        if (res.landed) { res.flags.push('landed-not-received'); counts.landedNotReceivedStyles++; }
        if (res.landedOurs) counts.landedOursStyles++;
        if (res.landedDue) counts.landedDueStyles++;
        if (res.landedLater) counts.landedLaterStyles++;
        if (res.takenDueBy) counts.takenDueByStyles++;
        if (res.unplaced) { res.flags.push('erp-transit-unplaced'); counts.erpUnplacedStyles++; }
        if (res.unplacedInChina) { counts.unplacedInChinaStyles++; if (res.backlog != null) counts.backlogGuardStyles++; }
        if (res.offeredTwice) counts.offeredTwiceStyles++;
        for (const s of ships) {
          const p = s.place;
          if (p.matched) s.flags.push('erp-ours-sailed');
          if (p.landedOurs || p.landedDue) s.flags.push('landed-not-received');
          if (p.offChina) s.flags.push('erp-off-china');
          if (p.unplaced) s.flags.push('erp-transit-unplaced');
        }
      } else {
        // The ERP has less In Trans than we have sailed. A shipment and one of our sailed lots of the same quantity are the
        // same goods (the closest ETA first): matched. Of the rest, the difference, earliest ETA first, is either received
        // already (an ETA on or before the report's day: in its In WH, so dropped here, never more than the style's In
        // WH) or not on its In Trans yet (kept as Arriving from our data). The rest of our sailed units are its In Trans.
        counts.erpLessStyles++;
        let d = res.ourSailed - E;
        if (repDay == null) { res.flags.push('no-report-day'); counts.noReportDayStyles++; }
        const free = sailed.slice().sort(etaOrder), sameShips = new Set();
        for (const s of res.erpLots.slice().sort(etaOrder)) {
          const gap = l => (s.eta == null || l.eta == null ? Infinity : Math.abs(l.eta - s.eta));
          const same = free.filter(l => l.units === s.units).sort((a, b) => gap(a) - gap(b) || etaOrder(a, b))[0];
          if (!same) continue;
          free.splice(free.indexOf(same), 1); sameShips.add(s);
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
          l.received = rec; room -= rec;
          l.notOnErp = n - rec; res.kept += n - rec;
          if (passed && n - rec > 0) res.capped += n - rec;
          l.matched = l.units - n;
        }
        // Our China goods the backlog no longer has have left China per the ERP: as far as its In Trans due after the
        // report's day (not a shipment of the same quantity as one of ours) was matched above to ours with an ETA on or
        // before that day, those of ours were received instead (in its In WH, never more than it) and that In Trans comes
        // off our lots still in China. In Trans matched to ours still on their way is left as it is (it can't be told).
        const later = res.erpLots.filter(s => !sameShips.has(s) && !dueBy(s)).sort(etaOrder);
        const recv = free.filter(l => l.matched > 0 && repDay != null && l.eta != null && l.eta <= repDay);
        const k = res.backlog == null ? 0 : Math.min(offMax, sumU(later), sumU(recv, l => l.matched), room);
        if (k > 0) {
          let r = k;
          for (const l of recv) { if (r <= 0) break; const m = Math.min(l.matched, r); l.matched -= m; l.received += m; room -= m; r -= m; }
          const bases = new Set(goneBasesBy.get(style) || []);
          for (const s of later) if (base6(s.erpPo)) bases.add(base6(s.erpPo));
          const left = new Map(later.map(s => [s, s.units]));
          takeOffChina(k, bases, m0 => {
            let m = m0;
            for (const s of later) { if (m <= 0) break; const q = Math.min(left.get(s), m); left.set(s, left.get(s) - q); s.place.offChina += q; m -= q; }
          });
          counts.lessTakenStyles++;
        }
        for (const l of free) {
          res.received += l.received; res.matched += l.matched;
          if (l.received) { l.flags.push('received'); counts.receivedLots++; }
          if (l.notOnErp) { l.flags.push('not-on-erp-in-trans'); counts.keptExtraLots++; }
          if (l.matched) { l.flags.push('matched-erp'); counts.matchedLots++; }
        }
        if (res.received) { res.flags.push('received'); counts.receivedStyles++; }
        if (res.kept) { res.flags.push('not-on-erp-in-trans'); counts.keptExtraStyles++; }
        if (res.capped) { res.flags.push('received-over-in-wh'); counts.receivedCappedStyles++; }
        // Every one of the ERP's other shipments is our goods on the water or at port.
        for (const s of res.erpLots) {
          s.place.matched = s.units - s.place.offChina;
          if (s.place.matched) s.flags.push('erp-ours-sailed');
          if (s.place.offChina) s.flags.push('erp-off-china');
        }
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
        matchedOurs: x.matched, receivedDropped: x.received, keptExtra: x.kept, receivedCapped: x.capped,
        // The ERP's In Trans beyond our sailed goods: 1. landed, not received yet (ours on our sheets: landedLots, or due
        // by the report's day); 2. taken off our lots in China (by tier: on a vessel, on a PO that sailed, other;
        // takenDueBy: on its shipments due by the report's day), never below its Import PO Backlog (erpBacklog, null when
        // unknown: erpBacklogOff lists a part below zero or not a number; takeable: our China goods beyond it, sold ones
        // not loaded yet too); 3. unplaced (unplacedInChina of it while our sheets still show the style in China).
        // offeredTwice: the guard, always 0.
        landedNotReceived: x.landed, landedOurs: x.landedOurs, landedDue: x.landedDue, landedLater: x.landedLater, landedLots: x.landedLots,
        inErpTransitTaken: x.taken, takenOnVessel: x.takenBy.vessel, takenSamePo: x.takenBy.po, takenOther: x.takenBy.other, takenDueBy: x.takenDueBy,
        erpUnplaced: x.unplaced, unplacedInChina: x.unplacedInChina, erpBacklog: x.backlog, erpBacklogOff: x.backlogOff, takeable: x.takeable, chinaAll: x.chinaAll,
        offeredTwice: x.offeredTwice,
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
    // The ERP's In Trans that our data can't place (not sailed, landed or taken off China): one line per style. It stays
    // Arriving (the ERP's word); our goods in China are kept as far as the report's Import PO Backlog has them there.
    for (const s of styles) if (s.erpUnplaced > 0) {
      const u = s.erpUnplaced, inChina = s.unplacedInChina;
      const off = Array.isArray(s.erpBacklogOff) && s.erpBacklogOff.length
        ? ` The ATS report's Import PO Backlog for the style (${s.erpBacklogOff.map(v => (typeof v === 'number' ? fmtN(v) : `"${String(v).slice(0, 20)}"`)).join(', ')}) is below zero or not a number, so it isn't used; it may be the ERP's own correction of this In Trans.` : '';
      const why = inChina > 0
        ? `the ERP counts ${unitsText(u)} in transit that our sheets still show in China; check before offering from China. They're offered as Arriving, and our goods still in China stay available from China because the ATS report's Import PO Backlog still has ${fmtN(s.erpBacklog)} of the style there (our sheets have ${fmtN(s.chinaAll - s.inErpTransitTaken)}, sold goods not loaded yet included${s.inErpTransitTaken ? `, after ${fmtN(s.inErpTransitTaken)} came off` : ''}).`
        : `the ERP counts ${unitsText(u)} in transit that our sheets don't have on the water, at port, landed or still to sell in China. They're offered as Arriving on the ERP's word; check Container Tracking and the Master Sheet for them.${off}`;
      issue('erp-transit-unplaced', '', s.style, `Style ${s.styleRaw}: ${why}`, { units: u, inChina });
    }
    summary('received-over-in-wh', s => s.receivedCapped,
      u => `${unitsText(u)} on the water or at port have an ETA on or before the ATS report's date and aren't on its In Trans, but the style's In WH is less, so they can't all be in its warehouse number: the rest are kept as Arriving from our shipping data. Check whether they were received.`);
    summary('no-report-day', s => Math.max(0, s.ourSailed - (s.erpInTrans || 0)),
      u => `the ATS report's date isn't known, so ${u ? `none of the ${unitsText(u)} on the water or at port beyond its In Trans is called received (they're all kept as Arriving), and ` : ''}none of its In Trans beyond our sailed goods is called landed, not received yet: it comes off our goods still in China instead, as far as its Import PO Backlog allows.`);
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
    // landed, not received + taken off China + unplaced, and each of its shipments is placed in full; received ≤ In WH;
    // so from China + arriving + taken by orders + matched + received + in the ERP's transit = Incoming + the ERP's In
    // Trans. And the guards: nothing offered twice, nothing taken off China below the report's Import PO Backlog.
    const offBy = [];
    const PLACE = ['matched', 'landedOurs', 'landedDue', 'offChina', 'unplaced'];
    for (const s of styles) {
      const bad = s.lots.filter(l => [l.matched, l.received, l.inErpTransit, l.taken, l.available].some(v => !(v >= 0)) ||
        Math.abs(l.units - replaced(l) - l.taken - l.available) > 1e-9 ||
        (l.source === 'erp' && (!l.place || PLACE.some(k => !(l.place[k] >= 0)) || Math.abs(PLACE.reduce((t, k) => t + l.place[k], 0) - l.units) > 1e-9)));
      const st = stageStyle.get(s.style), u = (st && st.units) || null;
      const inc = u && isDay(u.incoming) ? u.incoming : null;
      const lotU = s.lots.reduce((t, l) => t + l.units, 0);
      const erpL = s.lots.filter(l => l.source === 'erp'), placed = k => erpL.reduce((t, l) => t + ((l.place && l.place[k]) || 0), 0);
      const why = [bad.length ? `${bad.length} lot${bad.length === 1 ? '' : 's'} don't add up` : '',
        inc != null && Math.abs(inc - s.ourIncoming) > 1e-9 ? `its lots hold ${fmtN(s.ourIncoming)}, Incoming is ${fmtN(inc)}` : '',
        s.arrivingFrom === 'erp' && (Math.abs(s.erpInTrans - s.matchedOurs - s.landedNotReceived - s.inErpTransitTaken - s.erpUnplaced) > 1e-9 ||
          s.landedNotReceived !== s.landedOurs + s.landedDue || s.inErpTransitTaken !== s.takenOnVessel + s.takenSamePo + s.takenOther ||
          placed('landedOurs') !== s.landedOurs || placed('landedDue') !== s.landedDue || placed('offChina') !== s.inErpTransitTaken || placed('unplaced') !== s.erpUnplaced)
          ? "the ERP's In Trans isn't matched + landed + taken off China + unplaced" : '',
        s.offeredTwice > 0 ? `${fmtN(s.offeredTwice)} In Trans on the ERP are still offered from China too` : '',
        s.takeable != null && s.inErpTransitTaken > s.takeable ? `${fmtN(s.inErpTransitTaken)} came off China, below the ERP's Import PO Backlog (${fmtN(s.takeable)} could)` : '',
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
      for (const k of ['arrivingErp', 'arrivingOurs', ...ERP_KEYS]) totals[k] += s[k] || 0;
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
    RULES, TOTAL_KEYS, ERP_KEYS, TIERS, LANDED_DAYS, LANDED_LATE_DAYS, DAYS_BEFORE_ETD, DAYS_BEFORE_CRD, DAYS_AFTER_ORDER, DISAGREE_DAYS };
});
