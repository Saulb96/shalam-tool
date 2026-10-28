import test from 'node:test';
import assert from 'node:assert/strict';
import { core } from './load.mjs';

const U = core('upc');
let n = 0;
const it = fields => ({ id: String(++n), fields });

test('UPC_MASTER items become lookups by style and by UPC; problems are listed, never fixed', () => {
  n = 0;
  const items = [
    it({ Title: 'AB100', UPC: '012345678905', Style: 'AB100', Description: 'Lip gloss', Category: 'Cosmetics', SubCategory: 'Lips', Brand: 'Brand 1' }), // 1
    it({ Title: 'AB100', UPC: '012345678912', Style: ' ab100 ', Description: 'Lip gloss, shade 2' }), // 2 same style, second UPC
    it({ Title: '012345678929', UPC: '012345678929', Style: '', Description: 'Loose item' }), // 3 Title is the UPC: no style
    it({ Title: 'CD 200', UPC: '', Style: null, Description: 'Brush' }), // 4 style from Title, no UPC
    it({ Title: 'EF300', UPC: '12345-678', Style: 'EF300' }), // 5 bad UPC
    it({ Title: 'GH400', UPC: '012345678905', Style: 'GH400' }), // 6 UPC already on AB100
    it({ Title: 'AB100', UPC: '012345678905', Style: 'AB100' }), // 7 the same UPC and style again
    it({ Title: '012345678936', UPC: null, Style: '' }), // 8 UPC only in Title
    it({ Title: '99000123456', UPC: '', Style: '' }), // 9 11 digits in Title: a bad UPC, not a style
    it({ Title: 'Other', UPC: 840000000017, Style: 'IJ500' }), // 10 UPC as a number; Style wins over Title
    it({ Title: 'IJ501', UPC: 12345678905, Style: 'IJ501' }), // 11 a number that lost its leading zero
    it({ Title: 'RM 2201', UPC: '840000000024', Style: '' }), // 12
    it({ Title: 'RM2201', UPC: '840000000031', Style: '' }), // 13 a different style from RM 2201
  ];
  const r = U.fromListItems(items);
  assert.equal(r.ok, true);
  assert.deepEqual(r.counts, { rows: 13, used: 11, styles: 8, upcs: 7, badUpc: 3, noUpc: 1, noStyle: 3, dupUpc: 1, repeats: 1 });

  const ab = r.byStyle.get('AB100');
  assert.deepEqual(ab.map(e => e.upc), ['012345678905', '012345678912']);
  assert.deepEqual(ab[0], { upc: '012345678905', upcRaw: '012345678905', style: 'AB100', styleRaw: 'AB100', desc: 'Lip gloss',
    category: 'Cosmetics', sub: 'Lips', brand: 'Brand 1', itemId: '1' });
  assert.equal(r.byUpc.get('012345678905').itemId, '1'); // the first item keeps the UPC
  assert.deepEqual(r.byStyle.get('GH400').map(e => e.upc), ['012345678905']);
  assert.deepEqual(r.byStyle.get('CD 200').map(e => [e.upc, e.desc]), [[null, 'Brush']]);
  assert.deepEqual(r.byStyle.get('EF300').map(e => [e.upc, e.upcRaw]), [[null, '12345-678']]);
  assert.equal(r.byUpc.get('012345678929').style, '');
  assert.equal(r.byUpc.get('012345678936').upcRaw, '012345678936');
  assert.equal(r.byUpc.get('840000000017').style, 'IJ500');
  assert.ok(!r.byStyle.has('OTHER') && !r.byStyle.has('012345678929') && !r.byStyle.has('99000123456'));
  assert.equal(r.byStyle.get('IJ501')[0].upcRaw, '12345678905');
  assert.ok(r.byStyle.has('RM 2201') && r.byStyle.has('RM2201'));

  assert.deepEqual(r.issues.map(i => `${i.row} ${i.code}`),
    ['3 nostyle', '4 noupc', '5 badupc', '6 dupupc', '7 repeat', '8 nostyle', '9 badupc', '9 nostyle', '11 badupc']);
  assert.ok(r.issues.every(i => i.source === 'UPC_MASTER' && i.message));
  assert.match(r.issues.find(i => i.code === 'dupupc').message, /GH400.*AB100/);
});

test('every item gets a result: used, or an issue', () => {
  n = 0;
  const items = [it({ Style: 'A1', UPC: '012345678905' }), it({ Style: 'A1', UPC: '012345678905' }), it({}), { id: 'x' }];
  const r = U.fromListItems(items);
  const withIssue = new Set(r.issues.map(i => i.row));
  assert.equal(r.counts.rows, 4);
  assert.equal(r.counts.used + withIssue.size, 4);
});

test('no items, or no styles, is never success', () => {
  const none = U.fromListItems([]);
  assert.equal(none.ok, false);
  assert.match(none.error, /no items/);
  assert.equal(U.fromListItems(null).ok, false);
  n = 0;
  const nostyle = U.fromListItems([it({ Title: '012345678905', UPC: '012345678905' })]);
  assert.equal(nostyle.ok, false);
  assert.match(nostyle.error, /none of the 1 items has a style/);
});
