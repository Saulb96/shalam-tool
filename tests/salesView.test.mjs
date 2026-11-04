import test from 'node:test';
import assert from 'node:assert/strict';
import { core } from './load.mjs';

const C = core('common');
const V = core('salesView');
const D = (m, d, y = 2026) => C.ymd(y, m, d);

// ---------- invented sheets, as InvSalesSheet.readWorkbook gives them back (styles SV…, brands Brand A / Brand B)
const row = o => ({ style: 'SV100', upc: '000111222333', desc: 'Item SV100', brand: 'Brand A', category: 'Lip', sub: 'Gloss', pack: '24/6', note: '',
  onHand: 0, inTransit: 0, transitFrom: null, total: 0, photoKey: o.style || 'SV100', ...o });
const SALES = {
  kind: 'sales',
  rows: [
    row({ onHand: 100, inTransit: 300, transitFrom: 'now', total: 400 }),
    row({ pack: '12/3', note: 'Reserved -ZZQ', onHand: 50, total: 50 }),
    row({ style: 'SV200', upc: '000000000200', desc: 'Bath gloves', brand: 'Brand B', category: 'Bath', sub: 'Gloves', pack: '12/–', inTransit: 80, transitFrom: D(12, 8), total: 80 }),
    row({ style: 'SV300', desc: 'Shower cap', brand: '', category: 'Bath', sub: '', onHand: 10, inTransit: 0, total: 10 }),
  ],
  byDate: [
    { style: 'SV100', desc: 'Item SV100', category: 'Lip', pack: '24/6', note: '', when: 'now', where: 'On hand', units: 100 },
    { style: 'SV100', desc: 'Item SV100', category: 'Lip', pack: '24/6', note: '', when: 'now', where: 'In transit', units: 120 },
    { style: 'SV100', desc: 'Item SV100', category: 'Lip', pack: '24/6', note: '', when: D(11, 27), where: 'In transit', units: 100 },
    { style: 'SV100', desc: 'Item SV100', category: 'Lip', pack: '24/6', note: '', when: D(12, 20), where: 'In transit', units: 80 },
    { style: 'SV100', desc: 'Item SV100', category: 'Lip', pack: '12/3', note: 'Reserved -ZZQ', when: 'now', where: 'On hand', units: 50 },
    { style: 'SV200', desc: 'Bath gloves', category: 'Bath', pack: '12/–', note: '', when: D(12, 8), where: 'In transit', units: 80 },
    { style: 'SV300', desc: 'Shower cap', category: 'Bath', pack: '24/6', note: '', when: 'now', where: 'On hand', units: 10 },
  ],
  photos: [{ data: new Uint8Array([1]) }, { data: new Uint8Array([1]) }, null, { data: new Uint8Array([2]) }],
};
const CHINA = {
  kind: 'china',
  rows: [
    { style: 'SV100', upc: '', desc: 'Item SV100', brand: 'Brand A', category: 'Lip', sub: 'Gloss', pack: '24/6', note: '', units: 480, pullFrom: D(12, 1), estimate: false, photoKey: 'SV100' },
    { style: 'SV100', upc: '', desc: 'Item SV100', brand: 'Brand A', category: 'Lip', sub: 'Gloss', pack: '24/6', note: '', units: 100, pullFrom: D(12, 15), estimate: true, photoKey: 'SV100' },
    { style: 'SV200', upc: '', desc: 'Bath gloves', brand: 'Brand B', category: 'Bath', sub: 'Gloves', pack: '12/–', note: '', units: 60, pullFrom: null, estimate: false, photoKey: 'SV200' },
  ],
  byDate: [],
  photos: [null, null, null],
};
const sum = (v, k) => v.rows.reduce((t, r) => t + r[k], 0);

test('no filter: every line as it is, with its photo and schedule', () => {
  const v = V.view(SALES, {});
  assert.deepEqual(v.rows, SALES.rows);
  assert.deepEqual(v.byDate, SALES.byDate);
  assert.deepEqual(v.at, [0, 1, 2, 3]);
  assert.equal(v.photos[3], SALES.photos[3]);
  assert.deepEqual(v.totals, { onHand: 160, inTransit: 380, total: 540 });
  assert.deepEqual(v.counts, { rows: 4, styles: 3 });
  assert.equal(V.filtered({}), false);
  assert.equal(V.filtered({ brand: V.ALL, category: V.ALL, sub: V.ALL, where: 'all', q: ' ', by: null }), false);
});

test('the choices: brands and categories A to Z with none last; sub categories by category', () => {
  const ch = V.choices(SALES);
  assert.deepEqual(ch.brands, ['Brand A', 'Brand B', '']);
  assert.deepEqual(ch.categories, ['Bath', 'Lip']);
  assert.deepEqual(ch.subs, { [V.ALL]: ['Gloss', 'Gloves', ''], Bath: ['Gloves', ''], Lip: ['Gloss'] });
});

test('brand, category, sub category and the search: each one value; every word must match the style, UPC or description', () => {
  assert.deepEqual(V.view(SALES, { brand: 'Brand B' }).rows.map(r => r.style), ['SV200']);
  assert.deepEqual(V.view(SALES, { brand: '' }).rows.map(r => r.style), ['SV300'], 'no brand');
  assert.deepEqual(V.view(SALES, { category: 'Bath' }).rows.map(r => r.style), ['SV200', 'SV300']);
  assert.deepEqual(V.view(SALES, { category: 'Bath', sub: '' }).rows.map(r => r.style), ['SV300']);
  assert.deepEqual(V.view(SALES, { q: 'bath GLOVES' }).rows.map(r => r.style), ['SV200']);
  assert.deepEqual(V.view(SALES, { q: '000000000200' }).rows.map(r => r.style), ['SV200']);
  assert.deepEqual(V.view(SALES, { q: 'sv100 nothing' }).rows, []);
  const b = V.view(SALES, { category: 'Bath' });
  assert.deepEqual(b.byDate.map(x => x.style), ['SV200', 'SV300'], 'the schedule of the lines shown only');
  assert.deepEqual(b.photos, [null, SALES.photos[3]]);
  assert.equal(V.filtered({ category: 'Bath' }), true);
  assert.equal(V.filtered({ q: 'x' }), true);
});

test('on hand only or in transit only: the other units left out, a line with none left off', () => {
  const oh = V.view(SALES, { where: 'onHand' });
  assert.deepEqual(oh.rows.map(r => [r.style, r.pack, r.onHand, r.inTransit, r.transitFrom, r.total]),
    [['SV100', '24/6', 100, 0, null, 100], ['SV100', '12/3', 50, 0, null, 50], ['SV300', '24/6', 10, 0, null, 10]]);
  assert.ok(oh.byDate.every(b => b.where === 'On hand'));
  const it = V.view(SALES, { where: 'transit' });
  assert.deepEqual(it.rows.map(r => [r.style, r.onHand, r.inTransit, r.total]), [['SV100', 0, 300, 300], ['SV200', 0, 80, 80]]);
  assert.ok(it.byDate.every(b => b.where === 'In transit'));
  assert.deepEqual(it.totals, { onHand: 0, inTransit: 380, total: 380 });
});

test('available by a date: on hand at once, in transit from its date on or before it; the first date kept', () => {
  const v = V.view(SALES, { by: D(12, 1) });
  assert.deepEqual(v.rows.map(r => [r.style, r.pack, r.onHand, r.inTransit, r.transitFrom, r.total]),
    [['SV100', '24/6', 100, 220, 'now', 320], ['SV100', '12/3', 50, 0, null, 50], ['SV300', '24/6', 10, 0, null, 10]]);
  assert.deepEqual(v.byDate.filter(b => b.style === 'SV100' && b.pack === '24/6').map(b => [b.when, b.units]), [['now', 100], ['now', 120], [D(11, 27), 100]]);
  assert.equal(sum(v, 'total'), v.totals.total);
  // Only what's in transit by 12/8: SV200 from that day.
  assert.deepEqual(V.view(SALES, { by: D(12, 8), where: 'transit' }).rows.map(r => [r.style, r.inTransit, r.transitFrom]), [['SV100', 220, 'now'], ['SV200', 80, D(12, 8)]]);
  assert.equal(V.filtered({ by: D(12, 1) }), true);
});

test('Pull from China: brand, category and search the same; available by a date from its pull-from date (no date: never)', () => {
  assert.deepEqual(V.view(CHINA, {}).totals, { units: 640 });
  assert.deepEqual(V.view(CHINA, { by: D(12, 10) }).rows.map(r => [r.style, r.units]), [['SV100', 480]]);
  assert.deepEqual(V.view(CHINA, { by: D(12, 31) }).rows.map(r => r.units), [480, 100]);
  assert.deepEqual(V.view(CHINA, { brand: 'Brand B' }).rows.map(r => r.style), ['SV200']);
  // On hand or in transit means nothing here.
  assert.equal(V.view(CHINA, { where: 'onHand' }).rows.length, 3);
  assert.deepEqual(V.view(CHINA, { q: 'gloves' }).counts, { rows: 1, styles: 1 });
});

test('the sheet is never changed by a view', () => {
  const before = JSON.stringify(SALES);
  V.view(SALES, { where: 'transit', by: D(12, 1), category: 'Lip' });
  assert.equal(JSON.stringify(SALES), before);
});
