// Available to sell: per style, what a buyer can have now (warehouse stock after customer orders, from the ERP's ATS
// report) and from China (each open lot with one date), after everything already sold (Saul, Oct 30).
// - A lot is one allocation of core/stages.js still in China (No CRD, In production, Ready, Booked, Loaded) or sailed (On
//   the water, At port) that isn't sold: China Pickup POs and direct-to-customer sections are sold, never available.
// - Its ready date is synced: Container Tracking's Ready Date, else the PO Book's own date, else Factory CRD, else CRD
//   Expected, else Ship by (a date more than a year ahead or before 2015, a wrong year, is skipped). Sources more than
//   7 days apart are listed.
// - Pull-from date (Saul's rule) in China: 10 days before the ETD; no ETD, 5 days before the ready date (CRD); no ready
//   date, about 75 days after the order date (an estimate). Sailed lots are available on their ETA.
// - Customer orders come off stock first, then the lots in date order; what's left is oversold. Cases are whole cases.
// - The ATS report's In Trans is a cross-check only, never counted: sailed lots it has probably received already, and
//   orders beyond stock that goods it still has In Trans (landed, not received?) or goods sold to ship from China may
//   cover, are flagged and listed for Saul.
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
    water: 'On the water', port: 'At port' };
  const DAYS_BEFORE_ETD = 10, DAYS_BEFORE_CRD = 5, DAYS_AFTER_ORDER = 75, DISAGREE_DAYS = 7;
  // The ready date's sources, first one with a date wins.
  const READY_ORDER = ['ctReady', 'poBook', 'factoryCrd', 'crdExpected', 'shipBy'];
  const READY_LABEL = { ctReady: "Container Tracking's Ready Date", poBook: "the PO Book's date", factoryCrd: 'Factory CRD',
    crdExpected: 'CRD Expected', shipBy: 'Ship by', crd: 'CRD' };
  const READY_FILE = { ctReady: 'Container Tracking', poBook: 'PO Book', factoryCrd: 'PO Status Tracker',
    crdExpected: 'PO Status Tracker', shipBy: 'Master Sheet Live', crd: 'PO Book' };
  const RULES = ['etd', 'crd', 'order', 'none'];
  const TOTAL_KEYS = ['now', 'fromChina', 'arriving', 'total', 'soldChina', 'oversold', 'ordersTaken'];

  const fmtN = n => Number(n).toLocaleString('en-US');
  const unitsText = n => `${fmtN(n)} unit${n === 1 ? '' : 's'}`;
  const isDay = v => typeof v === 'number' && Number.isFinite(v);
  const fmt = C.fmtDay;
  const zero = () => Object.fromEntries(TOTAL_KEYS.map(k => [k, 0]));
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
    // The report's In Trans Qty is a cross-check only, never counted (Saul decides how these goods should count).
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
      erpReceivedStyles: 0, erpReceivedUnits: 0, erpReceivedAvailable: 0, erpInTransStyles: 0, erpInTransUnits: 0, erpInTransOversold: 0,
      soldInOrdersStyles: 0, soldInOrdersUnits: 0, soldInOrdersOversold: 0,
    };

    // ---------- lots, sold and unmatched units, from the allocations
    const lots = [], soldBy = new Map(), soldCasesBy = new Map(), soldNoPackBy = new Map(), unmatchedBy = new Map();
    const latestMaster = new Map();   // style -> { row, master } of its latest Master Sheet line with a Master
    const knownStyles = new Set();    // every style on any Master Sheet Live or Container Tracking line
    const add = (m, k, n) => m.set(k, (m.get(k) || 0) + n);
    for (const a of stages.allocs) {
      knownStyles.add(a.style);
      if (a.sheetRow != null && isDay(a.master) && a.master > 0) {
        const p = latestMaster.get(a.style);
        if (!p || a.sheetRow > p.row) latestMaster.set(a.style, { row: a.sheetRow, master: a.master });
      }
      const lotBucket = IN_CHINA.includes(a.bucket) || SAILED.includes(a.bucket);
      const sold = a.bucket === 'pullFromChina' || (lotBucket && (a.ptype === 'China Pickup' || (a.flags || []).includes('pull-section')));
      if (sold) {
        counts.soldAllocs++;
        add(soldBy, a.style, a.units);
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
        status: STATUS[a.bucket], bucket: a.bucket, inChina, units: a.units, taken: 0, available: a.units,
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
      if (lot.availableDate == null) { lot.flags.push('no-date'); counts.noDateLots++; }
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
    counts.lots = lots.length;
    const lotOrder = (a, b) => (a.availableDate == null) - (b.availableDate == null) || (a.availableDate ?? 0) - (b.availableDate ?? 0) ||
      byPo(a.po, b.po) || String(a.id).localeCompare(String(b.id));
    lots.sort(lotOrder);
    const lotsBy = new Map();
    for (const l of lots) (lotsBy.get(l.style) || lotsBy.set(l.style, []).get(l.style)).push(l);

    // ---------- every style: on the stage engine's list, with a lot, or on the ATS report
    const stageStyle = new Map((stages.styles || []).map(s => [s.style, s]));
    const keys = new Set([...stageStyle.keys(), ...lotsBy.keys(), ...(atsOk ? ats.styles.keys() : [])]);
    const upcBy = upc && upc.byStyle instanceof Map ? upc.byStyle : null;
    const styles = [];
    const notOnReport = [], unknownStock = [], erpOnly = [];
    for (const style of [...keys].sort((a, b) => (a < b ? -1 : a > b ? 1 : 0))) {
      const st = stageStyle.get(style) || null, ls = lotsBy.get(style) || [], e = atsOk ? ats.styles.get(style) || null : null;
      const um = upcBy ? upcBy.get(style) || [] : [];
      const flags = [];
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
      // Customer orders beyond stock come off the lots, earliest available date first.
      let left = shortfall, ordersTaken = 0;
      for (const l of ls) {
        if (left <= 0) break;
        const n = Math.min(l.units, left);
        l.taken = n; l.available = l.units - n; left -= n; ordersTaken += n;
        l.flags.push(l.available ? 'part-taken' : 'taken');
        counts.lotsTaken++; if (!l.available) counts.lotsFullyTaken++;
      }
      const oversold = left;
      const soldChina = soldBy.get(style) || 0, unmatched = unmatchedBy.get(style) || 0;
      // Cross-checks with the ATS report, never counted (listed, so the numbers aren't trusted blindly):
      // - sailed lots with an ETA on or before the report's day, beyond its In Trans: the ERP has probably received them
      //   into In WH already, so they may be in Available now and in Arriving;
      // - its In Trans beyond what we have on the water, at port or booked with an ETD, while orders are beyond stock:
      //   probably goods that landed but aren't received yet, which may cover those orders;
      // - orders beyond stock no more than the units sold to ship from China: those orders may be the same goods.
      let erpInTrans = null, erpReceived = 0, erpInTransBeyond = 0, soldInOrders = 0;
      if (known && transKnown) {
        erpInTrans = isDay(e.inTrans) ? e.inTrans : 0;
        const uSum = list => list.reduce((t, l) => t + l.units, 0);
        const sailedU = uSum(ls.filter(l => !l.inChina));
        if (repDay != null) {
          const early = ls.filter(l => !l.inChina && l.eta != null && l.eta <= repDay);
          erpReceived = Math.min(uSum(early), Math.max(0, sailedU - erpInTrans));
          if (erpReceived > 0) {
            flags.push('erp-received'); for (const l of early) l.flags.push('erp-received');
            counts.erpReceivedStyles++; counts.erpReceivedUnits += erpReceived;
            counts.erpReceivedAvailable += Math.min(erpReceived, early.reduce((t, l) => t + l.available, 0));
          }
        }
        if (shortfall > 0) {
          erpInTransBeyond = Math.max(0, erpInTrans - sailedU - uSum(ls.filter(l => l.inChina && l.etd != null)));
          if (erpInTransBeyond > 0) {
            flags.push('erp-in-transit'); counts.erpInTransStyles++; counts.erpInTransUnits += Math.min(erpInTransBeyond, shortfall);
            if (oversold > 0) counts.erpInTransOversold++;
          }
        }
      }
      if (known && shortfall > 0 && soldChina > 0 && shortfall <= soldChina) {
        soldInOrders = shortfall; flags.push('orders-may-include-sold');
        counts.soldInOrdersStyles++; counts.soldInOrdersUnits += shortfall; if (oversold > 0) counts.soldInOrdersOversold++;
      }
      if (!known && ls.length) { flags.push('orders-unknown'); for (const l of ls) l.flags.push('orders-unknown'); counts.ordersUnknownStyles++; counts.ordersUnknownUnits += ls.reduce((s, l) => s + l.units, 0); }
      if (oversold > 0) { flags.push('oversold'); counts.oversoldStyles++; }
      if (now > 0) counts.stylesWithStock++;
      if (e && !knownStyles.has(style) && !(um.length)) { flags.push('erp-only'); erpOnly.push(style); }
      if (ls.some(l => l.flags.includes('dates-disagree'))) flags.push('dates-disagree');
      if (ls.some(l => l.flags.includes('no-date'))) flags.push('no-date');

      // Case pack: the lots' Master, 'mixed' when they differ; with no lot, the latest Master Sheet line's.
      const masters = uniq(ls.map(l => l.master));
      const lm = latestMaster.get(style);
      const casePack = masters.length > 1 ? 'mixed' : masters.length ? masters[0] : lm ? lm.master : null;
      if (masters.length > 1) flags.push('mixed-case-pack');
      const stockPack = isDay(casePack) ? casePack : lm ? lm.master : null;

      const sum = (pred, f) => ls.filter(pred).reduce((s, l) => s + f(l), 0);
      const fromChina = sum(l => l.inChina, l => l.available), arriving = sum(l => !l.inChina, l => l.available);
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
      const umDesc = um.find(x => x.desc), umCat = um.find(x => x.category);
      const upcs = uniq([...(st ? st.upcs.map(u => u.upc) : []), ...ls.map(l => l.upc), ...um.map(x => x.upc)]);
      styles.push({
        style, styleRaw: (st && st.styleRaw) || (ls[0] && ls[0].styleRaw) || (e && e.styleRaw) || style,
        desc: (st && st.desc) || (umDesc && umDesc.desc) || (e && e.desc) || '',
        category: (st && st.category) || (umCat && umCat.category) || (e && e.category) || '',
        factory: (st && st.factory) || '', factories: st ? st.factories.slice() : [],
        upcs, casePack, onReport: atsOk ? !!e : null,
        inWh, orders, stockNet, now, fromChina, arriving, total: (now ?? 0) + fromChina + arriving,
        soldChina, unmatched, oversold, ordersTaken, cases, noCasePackUnits, looseUnits,
        erpInTrans, erpReceived, erpInTransBeyond, soldInOrders,
        firstPull: first ? first.pullFrom : null, firstPullRule: first ? first.pullRule : null,
        firstPullEstimate: !!(first && first.estimate), firstPullNow: !!(first && first.pullNow),
        firstArrival: arr.length ? Math.min(...arr.map(l => l.arrival)) : null,
        lots: ls, flags,
      });
    }
    counts.styles = styles.length;
    // The cross-check notes on a style with orders beyond its stock, in words.
    function checkNotes(s) {
      const n = [];
      if (s.erpInTransBeyond > 0) n.push(`the ATS report shows ${fmtN(s.erpInTrans)} In Trans, ${unitsText(s.erpInTransBeyond)} more than we have on the water, at port or booked with an ETD: probably goods that have landed but aren't received yet (not counted here).`);
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
    // Lots with no date to sell them from: one line per PO + style.
    const noDate = new Map();
    for (const l of lots) if (l.availableDate == null) {
      const k = l.po + '|' + l.style + '|' + (l.inChina ? 'c' : 's');
      const p = noDate.get(k) || noDate.set(k, { po: l.po, style: l.style, styleRaw: l.styleRaw, inChina: l.inChina, units: 0, statuses: new Set() }).get(k);
      p.units += l.units; p.statuses.add(l.status);
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
    // Sailed lots the ERP has probably received already: one line per style.
    for (const s of styles) if (s.erpReceived > 0) {
      const early = s.lots.filter(l => l.flags.includes('erp-received'));
      issue('erp-received', early.length === 1 ? early[0].po : '', s.style, `Style ${s.styleRaw}: ${unitsText(s.erpReceived)} on the water or at port (PO ${uniq(early.map(l => l.po)).join(', ')}, ETA ${uniq(early.map(l => fmt(l.eta))).join(', ')}) have an ETA on or before the ATS report's date (${fmt(repDay)}) but aren't In Trans on it (In Trans ${fmtN(s.erpInTrans)}), so the ERP has probably received them into In WH already: they may be counted twice, in Available now and in Arriving. Check the ERP and Container Tracking.`,
        { units: s.erpReceived });
    }
    // Styles whose orders beyond stock may be covered by goods not counted here: one summary line each.
    const summary = (code, unitsOf, text) => {
      const list = styles.filter(s => s.flags.includes(code));
      if (!list.length) return;
      const u = list.reduce((t, s) => t + unitsOf(s), 0), ov = list.filter(s => s.oversold > 0).length;
      issue(code, '', '', `${fmtN(list.length)} style${list.length === 1 ? '' : 's'} (${fmtN(ov)} of them oversold): ${text(u)} For example ${list.slice().sort((a, b) => unitsOf(b) - unitsOf(a)).slice(0, 5).map(s => `${s.styleRaw} ${fmtN(unitsOf(s))}`).join(', ')}${list.length > 5 ? '…' : ''}.`,
        { count: list.length, units: u });
    };
    summary('erp-in-transit', s => Math.min(s.erpInTransBeyond, Math.max(0, -s.stockNet)),
      u => `customer orders are beyond the warehouse stock while the ATS report shows more In Trans than we have on the water, at port or booked with an ETD, probably goods that have landed but aren't received yet. Up to ${unitsText(u)} of those orders may be covered by them, so lots from China may be taken, or a style called oversold, when they aren't.`);
    summary('orders-may-include-sold', s => s.soldInOrders,
      u => `the customer orders beyond the warehouse stock (${unitsText(u)}) are no more than the style's goods sold to ship from China (China Pickup or straight to the customer). If the ERP's Total Customer Orders include those orders, they are taken off twice.`);
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
      issue('erp-only', '', '', `${fmtN(erpOnly.length)} style${erpOnly.length === 1 ? ' is' : 's are'} on the ATS report but on no Master Sheet Live or Container Tracking line and not in UPC_MASTER (${fmtN(withStock)} with stock or orders), so they show with their stock only: for example ${erpOnly.slice(0, 5).map(k => ats.styles.get(k).styleRaw).join(', ')}${erpOnly.length > 5 ? '…' : ''}.`,
        { count: erpOnly.length });
    }

    // ---------- totals (units, and cases where the Master is known)
    const totals = zero(), tCases = zero(), tNoPack = zero(), tLoose = { now: 0, fromChina: 0, arriving: 0, total: 0 };
    for (const s of styles) {
      totals.now += s.now || 0; totals.fromChina += s.fromChina; totals.arriving += s.arriving; totals.soldChina += s.soldChina;
      totals.oversold += s.oversold; totals.ordersTaken += s.ordersTaken;
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
