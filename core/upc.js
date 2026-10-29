// UPC_MASTER (the style list on the Production site) read from its Graph list items into lookups by style and by
// UPC. A bad UPC is listed, never reshaped; a UPC on two styles is listed. Pure: items in, maps and issues out.
// Loads in the page (window.InvUpc) and in Node (module.exports); needs core/common.js first in the page.
(function (root, factory) {
  const api = factory(typeof module === 'object' && module.exports ? require('./common.js') : root.InvCommon);
  if (typeof module === 'object' && module.exports) module.exports = api;
  root.InvUpc = api;
})(typeof globalThis !== 'undefined' ? globalThis : this, function (C) {
  'use strict';
  if (!C) throw new Error('core/common.js must load before core/upc.js');
  const { S, normStyle, normUpc } = C;
  const SOURCE = 'UPC_MASTER';

  // Digits only, 11 to 14 long: a UPC (good or bad), never a style, so such a Title is not used as one.
  const upcLike = s => /^\d{11,14}$/.test(s);

  // items: [{ id, fields: { Title, UPC, Style, Description, Category, SubCategory, Brand } }].
  // Style is Style, else Title when Title isn't a UPC. UPC is UPC, else Title when UPC is blank and Title looks like a
  // UPC (some items carry it in Title only); a malformed one is listed with its text as it is.
  // byStyle: normStyle -> [entry] (one per UPC; an item with no good UPC still gives the style its description).
  // byUpc: upc -> entry of the first item with that UPC. entry = { upc (or null), upcRaw, style, styleRaw, desc,
  // category, sub, brand, itemId }.
  // A good UPC with no style is one summary line, not one issue per item: UPC_MASTER holds thousands of them (8,788 on
  // Oct 29), and listing each buried the problems someone can fix (Saul, Oct 29).
  function fromListItems(items) {
    const byStyle = new Map(), byUpc = new Map(), issues = [], upcOnly = [];
    const counts = { rows: 0, used: 0, styles: 0, upcs: 0, badUpc: 0, noUpc: 0, noStyle: 0, upcOnly: 0, dupUpc: 0, repeats: 0 };
    const issue = (row, code, message) => issues.push({ source: SOURCE, code, row, message });

    for (const it of items || []) {
      counts.rows++;
      const f = (it && it.fields) || {}, itemId = it && it.id != null ? String(it.id) : '';
      const title = S(f.Title);
      let upcRaw = S(f.UPC);
      if (!upcRaw && upcLike(title)) upcRaw = title;
      const upc = normUpc(upcRaw);
      let styleRaw = S(f.Style);
      if (!styleRaw && title && title !== upcRaw && !upcLike(title)) styleRaw = title;
      const style = normStyle(styleRaw);
      const e = { upc, upcRaw, style, styleRaw, desc: S(f.Description), category: S(f.Category), sub: S(f.SubCategory),
        brand: S(f.Brand), itemId };
      const where = style ? `style ${styleRaw}` : 'an item with no style';

      if (!upcRaw) { counts.noUpc++; issue(itemId, 'noupc', `No UPC on ${where}`); }
      else if (!upc) { counts.badUpc++; issue(itemId, 'badupc', `UPC "${upcRaw}" on ${where} isn't 12 or 13 digits: left as it is, not matched`); }
      if (!style) {
        counts.noStyle++;
        if (upc) { counts.upcOnly++; upcOnly.push(upc); } else issue(itemId, 'nostyle', 'No style and no usable UPC: not used');
      }

      let used = false;
      if (upc) {
        const first = byUpc.get(upc);
        if (!first) { byUpc.set(upc, e); used = true; }
        else if (first.style !== style) {
          counts.dupUpc++;
          issue(itemId, 'dupupc', `UPC ${upc} is on ${style ? 'style ' + styleRaw : 'an item with no style'} and on ` +
            `${first.style ? 'style ' + first.styleRaw : 'an item with no style'} (item ${first.itemId}): looked up by UPC, the first one is used`);
        } else {
          counts.repeats++;
          issue(itemId, 'repeat', `UPC ${upc} on style ${styleRaw || '(none)'} is listed twice (also item ${first.itemId})`);
        }
      }
      if (style) {
        const list = byStyle.get(style) || [];
        if (!upc || !list.some(x => x.upc === upc)) { list.push(e); used = true; }
        byStyle.set(style, list);
      }
      if (used) counts.used++;
    }
    if (upcOnly.length) {
      const n = upcOnly.length.toLocaleString('en-US');
      issues.push({ source: SOURCE, code: 'upconly', row: '', count: upcOnly.length,
        message: `${n} ${upcOnly.length === 1 ? 'row has' : 'rows have'} only a UPC, with no style, so ${upcOnly.length === 1 ? 'it' : 'they'} can't be tied to a style. ` +
          `For example UPC ${upcOnly.slice(0, 3).join(', ')}.` });
    }
    counts.styles = byStyle.size;
    counts.upcs = byUpc.size;
    const ok = counts.rows > 0 && byStyle.size > 0;
    const error = ok ? '' : counts.rows ? `${SOURCE}: none of the ${counts.rows} items has a style` : `${SOURCE}: no items read`;
    return { ok, error, byStyle, byUpc, issues, counts };
  }

  return { fromListItems };
});
