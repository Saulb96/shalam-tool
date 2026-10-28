// The PO Book's own reading of Master Sheet Live, the PO Status Tracker and PO_Book_Lines, copied unchanged from
// the PO Book's index.html (everything between the CORE markers), so a line is open or shipped exactly as the PO Book
// shows it. The helpers after the block add what the PO Book doesn't do: a result for every row read.
// Loads in the page (window.PoBookCore) and in Node (module.exports). Pure: no network, no DOM.
// No 'use strict' here: the copied block runs as it does in the PO Book, in a plain script.
(function (root, factory) {
  const api = factory();
  if (typeof module === 'object' && module.exports) module.exports = api;
  root.PoBookCore = api;
})(typeof globalThis !== 'undefined' ? globalThis : this, function () {
/*CORE-START*/
const DAY=864e5;
const S=v=>String(v??'').trim();
const normSty=v=>String(v??'').replace(/\s+/g,'').toUpperCase();
const num=v=>(v==null||v===''||isNaN(+v))?null:+v;
function ymd(y,m,d){if(y<2000||y>2200||m<1||m>12||d<1)return null;const t=Date.UTC(y,m-1,d);return new Date(t).getUTCDate()===d?t/DAY:null;}
function toDay(v){
  if(v==null||v==='')return null;
  if(typeof v==='number')return v>=36526?Math.floor(v)-25569:null;
  if(v instanceof Date)return isNaN(v)?null:ymd(v.getFullYear(),v.getMonth()+1,v.getDate());
  const s=S(v);let m=s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{2,4})$/);
  if(m){let y=+m[3];if(y<100)y+=2000;return ymd(y,+m[1],+m[2]);}
  m=s.match(/^(\d{4})-(\d{2})-(\d{2})/);return m?ymd(+m[1],+m[2],+m[3]):null;
}
const dayYear=d=>new Date(d*DAY).getUTCFullYear();
const fmtDay=d=>{if(d==null)return'';const o=new Date(d*DAY);return(o.getUTCMonth()+1)+'/'+o.getUTCDate()+'/'+o.getUTCFullYear();};
const isoDay=d=>d==null?'':new Date(d*DAY).toISOString().slice(0,10);
const todayDay=()=>{const n=new Date();return Date.UTC(n.getFullYear(),n.getMonth(),n.getDate())/DAY;};
const DT='\\d{1,2}/\\d{1,2}(?:/\\d{2,4})?';
function mkdate(s,anchor,after){
  const p=s.split('/').map(Number);let y;
  if(p.length===3)y=p[2]<100?p[2]+2000:p[2];
  else{const b=after??anchor;if(b==null)return null;y=dayYear(b);}
  let d=ymd(y,p[0],p[1]);if(d==null)return null;
  if(p.length===2){
    if(after!=null&&d<after)d=ymd(y+1,p[0],p[1]);
    else if(after==null&&anchor!=null&&d<anchor-45)d=ymd(y+1,p[0],p[1]);
  }
  return d;
}
function parseNotes(raw,anchor){
  const n=String(raw??'').replace(/\([^)]*\)/g,'');
  let cpo=[];const mm=/\bPO\b[ \t]*#?[ \t]*:?[ \t]*/i.exec(n);
  if(mm){
    const lines=n.slice(mm.index+mm[0].length).split('\n');
    for(let k=0;k<lines.length;k++){
      const l=lines[k].trim(),nums=l.match(/\d[\d-]{4,}/g)||[];
      if(k===0&&nums.length){cpo.push(nums[0]);break;}
      if(k>0&&nums.length&&/^[\d\s,&-]+$/.test(l)){cpo.push(...nums);continue;}
      if(k>0)break;
    }
  }
  cpo=cpo.filter(c=>!/^10[0-2]\d{3}$/.test(c));
  let cs=null,cc=null;
  let w=new RegExp('\\bSW\\b\\s*:?\\s*('+DT+')\\s*-\\s*/?\\s*('+DT+')','i').exec(n);
  if(w){cs=mkdate(w[1],anchor);cc=cs!=null?mkdate(w[2],anchor,cs):null;}
  else{w=new RegExp('SHIP\\s*(?:DATE)?\\s*:?\\s*('+DT+')','i').exec(n);if(w)cs=mkdate(w[1],anchor);}
  return{cpo:cpo.join(', '),cs,cc};
}
function cols(h,want,sheet){
  const o={};for(const[k,name]of Object.entries(want)){const i=h.findIndex(x=>x===name);if(i<0)throw new Error(`${sheet}: missing column "${name}"`);o[k]=i;}return o;
}
function parseTracker(rows){
  const h=(rows[0]||[]).map(S);
  const c=cols(h,{po:'PO#',sup:'Supplier',od:'Order date',desc:'Item Description',sty:'Style #',type:'Order Type',cust:'Customer',first:'Customer 1st Ship Date',exp:'CRD Expected',crd:'Factory CRD'},'PO Status Tracker');
  const T=new Map(),all=[];let minPO=Infinity;
  for(let i=1;i<rows.length;i++){
    const r=rows[i];if(!r)continue;
    const po=S(r[c.po]),b=po.replace(/\s/g,'').match(/^(\d{6})/);if(!b)continue;
    minPO=Math.min(minPO,+b[1]);
    const sty=normSty(r[c.sty]);if(!sty)continue;
    const t={po,base:b[1],sty,styRaw:S(r[c.sty]),sup:S(r[c.sup]),od:toDay(r[c.od]),desc:S(r[c.desc]),type:S(r[c.type]),cust:S(r[c.cust]),first:toDay(r[c.first]),exp:toDay(r[c.exp]),crd:toDay(r[c.crd])};
    T.set(t.base+'|'+sty,t);all.push(t);
  }
  for(const t of all)if(t.sty.endsWith('-F')&&!T.has(t.base+'|'+t.sty.slice(0,-2)))T.set(t.base+'|'+t.sty.slice(0,-2),t);
  return{T,all,minPO};
}
function parseMaster(rows,trk){
  const h=(rows[0]||[]).map(S);
  const c=cols(h,{name:'Product name',sty:'Style Number',price:'Price',qty:'Quantity',po:'Po Number',fac:'Factory',od:'Order Date',sb:'Ship by',notes:'Notes'},'Master Sheet Live');
  const isRound=r=>r&&/^\s*Round/i.test(S(r[c.name]));
  let last=-1;for(let i=1;i<rows.length;i++)if(isRound(rows[i]))last=i;
  const lines=[],issues=[],seen=new Set(),used=new Map();let blk=null;
  for(let i=1;i<rows.length;i++){
    const r=rows[i];if(!r)continue;
    if(isRound(r)){blk=null;continue;}
    const pc=S(r[c.po]);
    if(pc)blk={po:pc,fac:S(r[c.fac]),od:toDay(r[c.od]),sb:toDay(r[c.sb]),notes:r[c.notes]};
    if(!blk)continue;
    const b=blk.po.replace(/\s/g,'').match(/^(\d{6})/);if(!b||+b[1]<trk.minPO)continue;
    const base=b[1],sty=normSty(r[c.sty]);
    if(!sty){if(S(r[c.name]))issues.push({code:'nostyle',po:blk.po,sty:'',desc:S(r[c.name]),issue:'No style # on the Master Sheet Live line — left out'});continue;}
    const tr=trk.T.get(base+'|'+sty)||trk.T.get(base+'|'+(sty.endsWith('-F')?sty.slice(0,-2):sty+'-F'));
    seen.add(base+'|'+(tr?tr.sty:sty));
    const od=(tr&&tr.od)??blk.od,n=parseNotes(blk.notes,od);
    const bo=/-\d+$/.test(blk.po);
    const crd=bo?blk.sb:(tr?(tr.crd??tr.exp??blk.sb):blk.sb);
    const cust=tr?tr.cust:'';
    let key=(blk.po+'|'+sty).toUpperCase();const u=(used.get(key)||0)+1;used.set(key,u);if(u>1)key+='#'+u;
    if(!tr)issues.push({code:'mslonly',po:blk.po,sty:S(r[c.sty]),desc:S(r[c.name]),issue:'Not on the PO Status Tracker — shown with Master Sheet data only'});
    lines.push({key,po:blk.po,base,sty:S(r[c.sty]),desc:(tr&&tr.desc)||S(r[c.name]),sup:(tr&&tr.sup)||blk.fac,od,qty:num(r[c.qty]),price:num(r[c.price]),
      crd,ptype:tr&&tr.type==='China Pickup'?'China Pickup':'LA',cust,cpo:cust==='M1S Stock'?'':n.cpo,cs:n.cs??(tr?tr.first:null),cc:n.cc,
      src:i>last?'open':'round',bo,tracked:!!tr,row:i+1});
  }
  const inRound=new Map();for(const l of lines)if(l.src==='round'){const k=l.key.split('#')[0];(inRound.get(k)||inRound.set(k,[]).get(k)).push(l.qty);}
  for(const l of lines)if(l.src==='open'&&(inRound.get(l.key.split('#')[0])||[]).includes(l.qty)){
    l.src='round';l.dupRound=true;
    issues.push({code:'dupround',po:l.po,sty:l.sty,desc:l.desc,issue:`Master Sheet row ${l.row} is below the last Round but the same line/qty is already inside a Round — counted as shipped`});
  }
  for(const t of trk.all){
    const k=t.base+'|'+t.sty;
    if(!seen.has(k)&&!seen.has(k+'-F')&&!(t.sty.endsWith('-F')&&seen.has(k.slice(0,-2))))
      issues.push({code:'trkonly',po:t.po,sty:t.styRaw,desc:t.desc,issue:'On the PO Status Tracker but not on Master Sheet Live — no qty/price, not shown'});
  }
  lines.sort(cmpPO);
  return{lines,issues};
}
const sfx=po=>{const m=String(po).match(/^(.*)-(\d+)$/);return m?{root:m[1],n:+m[2]}:{root:String(po),n:0};};
function cmpPO(a,b){return(+a.base-+b.base)||a.po.localeCompare(b.po,undefined,{numeric:true})||String(a.sty??'').localeCompare(String(b.sty??''),undefined,{numeric:true});}
function applyOverlay(src,ov,cfg){
  const lines=src.lines.map(l=>({...l,splitOut:0})),byKey=new Map(lines.map(l=>[l.key,l]));
  for(const o of ov.values()){
    if(o.Kind!=='split'||byKey.has(o.LineKey))continue;
    const p=byKey.get(o.ParentKey);if(!p)continue;
    const q=+o.Qty||0;p.qty=(p.qty??0)-q;p.splitOut+=q;
    const nl={...p,key:o.LineKey,po:o.PONumber,qty:q,crd:toDay(o.CRDOverride)??p.crd,bo:true,split:true,parentKey:p.key,splitOut:0,src:'open'};
    lines.push(nl);byKey.set(nl.key,nl);
  }
  for(const l of lines){
    const o=ov.get(l.key);
    if(o&&!l.split&&S(o.PONumber)&&S(o.PONumber)!==l.po){l.srcPo=l.po;l.po=S(o.PONumber);l.bo=true;l.renamed=true;}
    l.manual=!!(o&&o.Shipped===true);l.shipDate=l.manual?S(o.ShippedDate):'';l.shipBy=l.manual?S(o.ShippedBy):'';
    l.status=(l.src==='round'||l.manual)?'shipped':'open';
  }
  /* a PO that is partly shipped: its open styles show as the next -N backorder PO */
  const groups=new Map();for(const l of lines)(groups.get(l.po)||groups.set(l.po,[]).get(l.po)).push(l);
  const rootMax=new Map();for(const po of groups.keys()){const s=sfx(po);rootMax.set(s.root,Math.max(rootMax.get(s.root)||0,s.n));}
  const made=new Map();
  for(const[po,ls]of[...groups]){
    const open=ls.filter(l=>l.status==='open');if(!open.length||open.length===ls.length)continue;
    const root=sfx(po).root,mx=rootMax.get(root)||0,top=groups.get(root+'-'+mx);
    let target=made.get(root);
    if(!target){target=(mx>0&&top&&top!==ls&&top.every(l=>l.status==='open'))?root+'-'+mx:root+'-'+(mx+1);made.set(root,target);rootMax.set(root,Math.max(mx,sfx(target).n));}
    for(const l of open){l.po=target;l.bo=true;l.derived=true;}
  }
  for(const l of lines){
    const o=ov.get(l.key);
    if(o&&l.bo&&!l.split&&toDay(o.CRDOverride)!=null)l.crd=toDay(o.CRDOverride);
    l.buf=(l.cs!=null&&l.crd!=null)?l.cs-l.crd-(l.ptype==='LA'?cfg.la:0):null;
    l.val=(l.qty??0)*(l.price??0);
  }
  lines.sort(cmpPO);
  return lines;
}
/*CORE-END*/

  // ---------- below the copied block: helpers for this tool ----------

  // sha256 of the text strictly between the two markers in the PO Book as GitHub serves it (LF line ends), taken from
  // saulb96/po-book commit 21ef6e6 (Oct 2, 2026). The page and tools/check-real.mjs compare it with the live file.
  const CORE_SHA256 = '3ce21b049e7ca66f910077725bab7c2f675f6b7d425c0697d098be0b6edb080a';
  // Built from pieces so this file holds each marker only once, around the copied block.
  const MARK_A = '/*CORE' + '-START*/', MARK_B = '/*CORE' + '-END*/';

  // The text strictly between the markers, Windows line ends read as LF (a local clone has CRLF, GitHub serves LF),
  // or null when either marker is missing.
  function coreText(html) {
    if (html == null) return null;
    const s = String(html).replace(/\r\n/g, '\n');
    const a = s.indexOf(MARK_A);
    if (a < 0) return null;
    const b = s.indexOf(MARK_B, a + MARK_A.length);
    return b < 0 ? null : s.slice(a + MARK_A.length, b);
  }

  // PO_Book_Lines items ({ id, lastModifiedDateTime, fields }) by LineKey, as the PO Book's loadOverlay reads them:
  // an item without a LineKey is skipped, and a LineKey saved twice keeps its newest item.
  const DEF_CFG = { la: 30, amber: 14, refresh: 5 };
  function byLineKey(items) {
    const m = new Map();
    for (const it of items || []) {
      const f = it.fields || {};
      if (!f.LineKey) continue;
      const prev = m.get(f.LineKey);
      if (!prev || it.lastModifiedDateTime > prev._lm) m.set(f.LineKey, { ...f, _id: it.id, _lm: it.lastModifiedDateTime });
    }
    return m;
  }
  // The overlay applyOverlay takes: every line's newest item, without the settings item.
  function overlayFromItems(items) {
    const m = byLineKey(items);
    m.delete('__settings');
    return m;
  }
  // The PO Book's settings (the '__settings' item's Data over the defaults). Unreadable JSON gives the defaults, as
  // in the PO Book.
  function cfgFrom(items) {
    const it = byLineKey(items).get('__settings');
    try { return { ...DEF_CFG, ...(it ? JSON.parse(it.Data || '{}') : {}) }; } catch (e) { return { ...DEF_CFG }; }
  }

  // The columns the copied block asks for, found the same way (exact header text, ends trimmed), but every missing
  // one is named at once with the headers seen.
  const MASTER_COLS = { name: 'Product name', sty: 'Style Number', price: 'Price', qty: 'Quantity', po: 'Po Number',
    fac: 'Factory', od: 'Order Date', sb: 'Ship by', notes: 'Notes' };
  const TRACKER_COLS = { po: 'PO#', sup: 'Supplier', od: 'Order date', desc: 'Item Description', sty: 'Style #',
    type: 'Order Type', cust: 'Customer', first: 'Customer 1st Ship Date', exp: 'CRD Expected', crd: 'Factory CRD' };
  function headerCheck(rows, want, sheet) {
    const h = ((rows && rows[0]) || []).map(S), seen = h.filter(Boolean), c = {}, missing = [];
    for (const [k, name] of Object.entries(want)) { const i = h.indexOf(name); if (i < 0) missing.push(name); else c[k] = i; }
    const error = missing.length ? `${sheet}: missing column${missing.length > 1 ? 's' : ''} ` +
      missing.map(n => `"${n}"`).join(', ') + `. Headers seen: ${seen.join(', ') || 'none'}` : '';
    return { c, missing, seen, error };
  }
  const blankRow = r => !r || r.every(v => S(v) === '');
  const sixDigits = po => po.replace(/\s/g, '').match(/^(\d{6})/);
  // One PO: 6 digits, up to 3 letters, maybe a -suffix (900437A, 900437-BCF, 900437-1). Anything else, such as two POs
  // in one cell, is read by its first 6 digits.
  const PO_SHAPE = /^\d{6}[A-Z]{0,3}(?:\s*-\s*[A-Z0-9]+)?$/i;
  // A style as this tool compares them (common.js normStyle): inner spaces and -F kept.
  const exactSty = v => String(v ?? '').replace(/^[\s﻿ ]+|[\s﻿ ]+$/g, '').toUpperCase();

  const MASTER_RESULTS = {
    line: 'counted', round: 'Round row', blank: 'blank', noPo: 'no PO number above it',
    badPo: "PO number doesn't start with 6 digits", belowMin: 'older than the PO Status Tracker\'s lowest PO, by design',
    noStyle: 'no style number' };

  // A second pass over the rows parseMaster reads, walking them the same way, so that every row gets a result: a line
  // (parseMaster turns exactly these rows into lines), a Round row, blank, no PO number above it, a PO number that
  // doesn't start with 6 digits, a PO below the tracker's lowest (left out by design), or no style. The copied block
  // drops most of these without a word. Issues cover every row left out other than by design (blank, Round, below the
  // lowest PO), with row numbers, plus lines whose quantity is blank or not a number (the PO Book keeps them).
  // parseMaster's own 'nostyle' issues are a subset of the 'nostyle' issues here, without row numbers.
  function countMaster(rows, trk) {
    const sheet = 'Master Sheet Live';
    const hc = headerCheck(rows, MASTER_COLS, sheet);
    const counts = { line: 0, round: 0, blank: 0, noPo: 0, badPo: 0, belowMin: 0, noStyle: 0 };
    const out = { ok: false, error: hc.error, missing: hc.missing, seen: hc.seen, rowsRead: 0, counts, results: [],
      issues: [], minPO: trk ? trk.minPO : null, labels: MASTER_RESULTS };
    if (hc.error) return out;
    if (!trk) { out.error = `${sheet}: the PO Status Tracker wasn't read, so no line can be counted`; return out; }
    const c = hc.c, isRound = r => r && /^\s*Round/i.test(S(r[c.name]));
    const issue = (row, code, message) => out.issues.push({ source: sheet, code, sheet, row, message });
    let blk = null;
    for (let i = 1; i < rows.length; i++) {
      const r = rows[i], row = i + 1;
      let res;
      if (!r) res = 'blank';
      else if (isRound(r)) { blk = null; res = 'round'; }
      else {
        const pc = S(r[c.po]);
        if (pc) blk = { po: pc, row, odd: !PO_SHAPE.test(pc) };
        if (blankRow(r)) res = 'blank';
        else if (!blk) {
          res = 'noPo';
          issue(row, 'nopo', 'No PO number on this row or above it since the last Round row: the PO Book leaves it out');
        } else {
          const b = sixDigits(blk.po);
          if (!b) {
            res = 'badPo';
            issue(row, 'badpo', `PO number "${blk.po}" doesn't start with 6 digits: the PO Book leaves the row out`);
          } else if (+b[1] < trk.minPO) res = 'belowMin';
          else if (!normSty(r[c.sty])) {
            res = 'noStyle';
            issue(row, 'nostyle', 'No style number: the PO Book leaves the row out');
          } else {
            res = 'line';
            const q = r[c.qty];
            if (S(q) === '') issue(row, 'qty', 'Quantity is blank: the line counts no units');
            else if (num(q) == null) issue(row, 'qty', `Quantity "${S(q)}" isn't a number: the line counts no units`);
            if (blk.odd && !blk.listed) {
              blk.listed = true;
              issue(blk.row, 'po-text', `PO number "${blk.po}" isn't one PO: the PO Book reads it as PO ${b[1]}, and it can't match a PO on Container Tracking`);
            }
            // The PO Book takes the tracker's details for a style written another way (spaces, -F): listed, since
            // this tool keeps them apart.
            const sty = normSty(r[c.sty]);
            const tr = trk.T && (trk.T.get(b[1] + '|' + sty) || trk.T.get(b[1] + '|' + (sty.endsWith('-F') ? sty.slice(0, -2) : sty + '-F')));
            if (tr && exactSty(tr.styRaw) !== exactSty(r[c.sty]))
              issue(row, 'style-alias', `Style "${exactSty(r[c.sty])}" takes its customer, CRD and China Pickup / LA from "${exactSty(tr.styRaw)}" on the PO Status Tracker (PO ${tr.po}): the PO Book reads them as one item`);
          }
        }
      }
      counts[res]++;
      out.results.push({ row, result: res });
    }
    out.rowsRead = out.results.length;
    if (counts.line) out.ok = true;
    else out.error = trk.minPO === Infinity
      ? `${sheet}: no lines counted, because the PO Status Tracker gave no PO numbers`
      : `${sheet}: no PO lines found in ${out.rowsRead} rows`;
    return out;
  }

  const TRACKER_RESULTS = { tracked: 'counted', blank: 'blank', badPo: "no PO number starting with 6 digits",
    noStyle: 'no style number' };

  // The same for the PO Status Tracker rows parseTracker reads. Two rows with the same PO (first 6 digits) and style
  // are both counted, but the PO Book takes the lower row's details for that pair, so the pair is listed.
  function countTracker(rows) {
    const sheet = 'PO Status Tracker';
    const hc = headerCheck(rows, TRACKER_COLS, sheet);
    const counts = { tracked: 0, blank: 0, badPo: 0, noStyle: 0 };
    const out = { ok: false, error: hc.error, missing: hc.missing, seen: hc.seen, rowsRead: 0, counts, results: [],
      issues: [], labels: TRACKER_RESULTS };
    if (hc.error) return out;
    const c = hc.c, at = new Map();
    const issue = (row, code, message) => out.issues.push({ source: sheet, code, sheet, row, message });
    for (let i = 1; i < rows.length; i++) {
      const r = rows[i], row = i + 1;
      let res;
      if (blankRow(r)) res = 'blank';
      else {
        const po = S(r[c.po]), b = sixDigits(po), sty = normSty(r[c.sty]);
        if (!b) {
          res = 'badPo';
          issue(row, 'badpo', po ? `PO# "${po}" doesn't start with 6 digits: the PO Book leaves the row out`
            : 'No PO#: the PO Book leaves the row out');
        } else if (!sty) {
          res = 'noStyle';
          issue(row, 'nostyle', 'No style #: the PO Book leaves the row out');
        } else {
          res = 'tracked';
          if (!PO_SHAPE.test(po)) issue(row, 'po-text', `PO# "${po}" isn't one PO: the PO Book reads it as PO ${b[1]}`);
          const k = b[1] + '|' + sty, ex = exactSty(r[c.sty]), prev = at.get(k);
          if (prev && prev.ex === ex) issue(row, 'dup', `Same PO and style as row ${prev.row}: the PO Book uses this row's details for both`);
          else if (prev) issue(row, 'style-alias', `Same PO as row ${prev.row}, and the same style once spaces are taken out ("${prev.ex}" and "${ex}"): the PO Book reads them as one item and uses this row's details for both`);
          at.set(k, { row, ex });
        }
      }
      counts[res]++;
      out.results.push({ row, result: res });
    }
    out.rowsRead = out.results.length;
    if (counts.tracked) out.ok = true;
    else out.error = `${sheet}: no PO lines found in ${out.rowsRead} rows`;
    return out;
  }

  return { parseTracker, parseMaster, applyOverlay, toDay, fmtDay, normSty, CORE_SHA256, coreText, overlayFromItems,
    cfgFrom, countMaster, countTracker };
});
