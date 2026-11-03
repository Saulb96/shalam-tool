import test from 'node:test';
import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import { core, root, loadXLSX, loadExcelJS } from './load.mjs';

// The pictures worker (picturesWorker.js) run here as a plain function, with stand-ins for what a worker has: its own
// files (importScripts gives it the ones Node loaded), postMessage (the answers, kept), and the browser's drawing
// (createImageBitmap and OffscreenCanvas: counted; a canvas with no 2D context draws nothing, as when the browser's
// drawing fails). Invented data only.
const C = core('common');
const SRC = readFileSync(root('picturesWorker.js'), 'utf8');
const NAME = 'Example_Co_Inventory_Report_-_On_Hand_In_Transit_In_Production_11-02-2026_04-59_PM.xlsx';
const PNG1 = Uint8Array.from(Buffer.from('iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==', 'base64'));
const PNG2 = Uint8Array.from(Buffer.from('iVBORw0KGgoAAAANSUhEUgAAAAIAAAABCAYAAAD0In+KAAAAEUlEQVR42mP8z8DAwMDAwAAAD/4CAVv6FX4AAAAASUVORK5CYII=', 'base64'));
const HEAD = ['Image', 'Color Code', 'Style', 'Style Description', 'Category', 'Sub Category', 'Warehouse', 'In WH', 'Total Customer Orders', 'ATS from in WH'];
const row = (style, inWh) => [null, 'XX', style, 'Item ' + style, 'Lip', 'Gloss', 'WH A', inWh, 0, String(inWh)];
const ROWS = [['Inventory Report (invented)'], HEAD, row('PC201', 5), row('PC202', 6), row('PC203', 7), [null, null, 'Grand Total', null, null, null, null, 18, 0]];

// pics: [[0-based sheet row, image name]]; the same name placed twice is one picture part.
async function report(pics, images, edit) {
  const E = loadExcelJS(), wb = new E.Workbook(), ws = wb.addWorksheet('Report');
  for (const r of ROWS) ws.addRow(r);
  if (edit) edit(wb);
  const ids = {};
  for (const [r, k] of pics) {
    if (!(k in ids)) ids[k] = wb.addImage({ buffer: images[k], extension: 'png' });
    ws.addImage(ids[k], { tl: { col: 0, row: r }, br: { col: 1, row: r + 1 } });
  }
  return new Uint8Array(await wb.xlsx.writeBuffer());
}

// One read in the worker: { msgs (everything it said), done (its last answer), bitmaps (pictures drawn) }.
async function runWorker(bytes, o = {}) {
  const msgs = [], self = { crypto: globalThis.crypto };
  let bitmaps = 0, closed = false;
  const importScripts = (...files) => {
    assert.deepEqual(files, ['vendor/xlsx-js-style/xlsx.bundle.js', 'core/common.js', 'core/ats.js', 'core/erpPictures.js']);
    Object.assign(self, { XLSX: loadXLSX(), InvCommon: C, InvAts: core('ats'), InvErpPictures: core('erpPictures') });
  };
  const createImageBitmap = async () => { bitmaps++; return { width: 900, height: 300, close() {} }; };
  class OffscreenCanvas {
    constructor(w, h) { this.width = w; this.height = h; }
    getContext() { return o.noContext ? null : { fillRect() {}, drawImage() {} }; }
    async convertToBlob(opt) { return new Blob([`invented small picture ${this.width}x${this.height}`], { type: opt.type }); }
  }
  new Function('self', 'importScripts', 'postMessage', 'close', 'createImageBitmap', 'OffscreenCanvas', SRC)(
    self, importScripts, m => msgs.push(m), () => { closed = true; }, createImageBitmap, OffscreenCanvas);
  const buf = bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength);
  self.onmessage({ data: { type: 'read', bytes: buf, fileName: NAME, today: C.ymd(2026, 11, 2), maxSide: 160, quality: 0.8 } });
  for (const end = Date.now() + 20000; !closed && Date.now() < end;) await new Promise(r => setTimeout(r, 5));
  assert.ok(closed, 'the worker closed itself after answering');
  return { msgs, done: msgs.find(m => m.type === 'done'), bitmaps };
}

test('the worker: a picture shown for several styles is made small and hashed once; each style gets it', async () => {
  // PC201 and PC202 share one picture part; PC203 has its own.
  const { msgs, done, bitmaps } = await runWorker(await report([[2, 'a'], [3, 'a'], [4, 'b']], { a: PNG1, b: PNG2 }));
  assert.equal(msgs[0].type, 'progress', 'the first thing it says');
  assert.equal(done.ok, true, done.error);
  assert.equal(bitmaps, 2, 'two picture parts drawn, not three');
  const by = Object.fromEntries(done.pictures.map(p => [p[0], p]));
  assert.deepEqual(Object.keys(by).sort(), ['PC201', 'PC202', 'PC203']);
  assert.equal(by.PC201[5], by.PC202[5], 'the same hash');
  assert.notEqual(by.PC201[5], by.PC203[5]);
  assert.match(by.PC201[5], /^[0-9a-f]{64}$/);
  assert.ok(by.PC201[2] instanceof Blob && by.PC201[2].type === 'image/jpeg');
  assert.deepEqual([done.counts.pictures, done.counts.undrawn, done.scaled], [3, 0, true]);
});

test('the worker: the numbers come first, with the numbers\' part on its own (for the page to keep, no person\'s name in it)', async () => {
  const bytes = await report([[2, 'a']], { a: PNG1 }, wb => { wb.creator = 'Invented Person'; });
  const { msgs, done } = await runWorker(bytes);
  const n = msgs.findIndex(m => m.type === 'numbers'), d = msgs.findIndex(m => m.type === 'done');
  assert.ok(n > 0 && n < d, msgs.map(m => m.type).join(', '));
  const slim = msgs[n].slim;
  assert.ok(slim instanceof Uint8Array && slim.length > 0);
  assert.ok(!Buffer.from(slim).toString('latin1').includes('Invented Person'));
  assert.equal(done.ok, true, done.error);
  assert.deepEqual(msgs[n].numbers, done.numbers);
});

test('the worker: when the browser can draw none of the pictures, it says so in its counts (undrawn), with none sent back', async () => {
  const { done, bitmaps } = await runWorker(await report([[2, 'a'], [4, 'b']], { a: PNG1, b: PNG2 }), { noContext: true });
  // ok, as the file was read; the page must not take this for a report with no pictures (app.js: picturesStep).
  assert.deepEqual([done.ok, done.pictures.length, done.counts.undrawn, done.counts.anchors, bitmaps], [true, 0, 2, 2, 2]);
});
