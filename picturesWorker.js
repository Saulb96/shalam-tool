// The ERP report's pictures, read off the page's main thread (Saul, Nov 2): a Web Worker the page starts for one big ATS
// report (InvErpPictures.BIG_REPORT, 2 MB or more: a report with pictures) and ends with terminate(), on Cancel or once
// it has answered. It loads only the page's own files and makes no network request: the report's bytes come in from
// the page, its numbers and small pictures go back. The numbers (whole on their own) come first; the pictures only at
// the very end, so a worker that is ended or fails leaves no half set of pictures behind.
// The page sends one message:
//   { type: 'read', bytes: ArrayBuffer (transferred), fileName, today (a day number), maxSide (160), quality (0.8),
//     numbersOnly (true: the numbers alone, when the page already has this report's pictures) }
// and gets, in order:
//   { type: 'progress', step: 'opening', done: 0, total: 0 }   once, when the bytes are in (the first thing it says)
//   { type: 'numbers', fileName, numbers, slim }               once the numbers are read, before the pictures (whole on
//                                                              their own: the page uses them while the pictures are read);
//                                                              slim: the numbers' part of the file on its own (readPictures'),
//                                                              for the page to keep, or null
//   { type: 'progress', step: 'reading', done, total }         the report's picture parts unpacked (every 50 and the last)
//   { type: 'progress', step: 'scaling', done, total }         the pictures made small (every 50 and the last)
//   { type: 'done', ok, error, fileName, numbers, pictures, scaled, counts, differs }
// - numbers: InvAts.parseAts on the report (the same as reading the whole file; its own ok and error), or null when the
//   file can't be opened at all.
// - pictures: [[style key, styleRaw, Blob, width, height, hash]], one per style; scaled true: a JPEG (quality 0.8), its
//   longest side at most maxSide px, as app.js makes Quote Images photos small. scaled false (this browser can't draw in
//   a worker): the pictures as they are in the report (Blob type image/png, image/jpeg, …), width and height null, for
//   the page to make small itself. hash: of the picture's bytes in the report (SHA-256, hex; FNV-1a and the size where
//   there's no SubtleCrypto), for the sales sheet's fingerprint: the same picture gives the same hash in any browser,
//   where the small copy's bytes can differ from one browser or graphics card to another. A picture shown for several
//   styles is made small (and hashed) once.
// - counts: readPictures' counts, plus undrawn (pictures the browser couldn't draw, left out; pictures is then the
//   styles with a picture sent back); differs: readPictures' ([{ style, styleRaw, rows }]).
// - ok false: error in plain words; pictures empty; numbers kept when they were read. noReader true: the worker's own
//   parts didn't load (the file itself wasn't looked at, so the page reads it as it reads a small report).
'use strict';
let loadError = '';
try {
  importScripts('vendor/xlsx-js-style/xlsx.bundle.js', 'core/common.js', 'core/ats.js', 'core/erpPictures.js');
} catch (e) { loadError = (e && e.message) || String(e); }

// Pictures made small at the same time.
const SCALE_LANES = 4;

// A picture's bytes in the report, hashed (see pictures above).
async function hashOf(bytes) {
  if (self.crypto && self.crypto.subtle) {
    return Array.from(new Uint8Array(await self.crypto.subtle.digest('SHA-256', bytes)), x => x.toString(16).padStart(2, '0')).join('');
  }
  let h = 0x811c9dc5;
  for (let i = 0; i < bytes.length; i++) { h ^= bytes[i]; h = Math.imul(h, 0x01000193); }
  return `${(h >>> 0).toString(16)}-${bytes.length}`;
}

// One picture: its longest side maxSide px (never larger than it was), on white (a JPEG has no transparency), a JPEG.
async function small(bytes, ext, maxSide, quality) {
  const bm = await createImageBitmap(new Blob([bytes], { type: 'image/' + ext }));
  try {
    const sc = Math.min(1, maxSide / Math.max(bm.width, bm.height, 1));
    const w = Math.max(1, Math.round(bm.width * sc)), h = Math.max(1, Math.round(bm.height * sc));
    const cv = new OffscreenCanvas(w, h), cx = cv.getContext('2d');
    cx.fillStyle = '#fff'; cx.fillRect(0, 0, w, h);
    cx.imageSmoothingQuality = 'high';
    cx.drawImage(bm, 0, 0, w, h);
    const blob = await cv.convertToBlob({ type: 'image/jpeg', quality });
    if (!blob || !blob.size) throw new Error("the picture couldn't be made smaller");
    return { blob, w, h };
  } finally { if (bm.close) bm.close(); }
}

async function read(msg) {
  const fileName = String(msg.fileName || 'the ATS report');
  const say = (step, done, total) => postMessage({ type: 'progress', step, done, total });
  const answer = (o) => postMessage(Object.assign({ type: 'done', ok: false, error: '', fileName, numbers: null, pictures: [], scaled: false, counts: {}, differs: [] }, o));
  if (loadError || !self.XLSX || !self.InvAts || !self.InvErpPictures) {
    return answer({ noReader: true, error: `Part of the picture reader didn't load (${loadError || 'the spreadsheet library or core files'}). Reload the page; if it keeps happening, tell Saul.` });
  }
  if (!(msg.bytes instanceof ArrayBuffer) || !msg.bytes.byteLength) return answer({ error: `${fileName} came to the picture reader empty (0 bytes).` });
  say('opening', 0, 0);
  const r = await self.InvErpPictures.readPictures({ XLSX: self.XLSX, InvAts: self.InvAts, bytes: new Uint8Array(msg.bytes), fileName,
    today: msg.today, numbersOnly: !!msg.numbersOnly, onProgress: (done, total) => say('reading', done, total),
    onNumbers: (numbers, slim) => postMessage({ type: 'numbers', fileName, numbers, slim: slim || null }) });
  const counts = Object.assign({}, r.counts, { undrawn: 0 });
  if (!r.ok) return answer({ error: r.error, numbers: r.numbers, counts });

  const list = [...r.pictures];
  const maxSide = Number(msg.maxSide) > 0 ? Number(msg.maxSide) : 160, quality = Number(msg.quality) > 0 ? Number(msg.quality) : 0.8;
  // A picture part shown for several styles has the same bytes for each (readPictures unpacks a part once): made small
  // and hashed once, and let go after its last style.
  const uses = new Map(), made = new Map();
  for (const [, p] of list) uses.set(p.bytes, (uses.get(p.bytes) || 0) + 1);
  const once = (p, fn) => {
    const b = p.bytes;
    if (!made.has(b)) made.set(b, fn(b));
    const got = made.get(b);
    if (uses.set(b, uses.get(b) - 1).get(b) <= 0) made.delete(b);
    p.bytes = null;
    return got;
  };
  // Without drawing in a worker, the pictures go back as they are, for the page to make small.
  if (typeof OffscreenCanvas !== 'function' || typeof createImageBitmap !== 'function') {
    const pictures = [];
    for (const [key, p] of list) {
      const ext = p.ext, s = await once(p, async b => ({ blob: new Blob([b], { type: 'image/' + ext }), hash: await hashOf(b) }));
      pictures.push([key, p.styleRaw, s.blob, null, null, s.hash]);
    }
    return answer({ ok: true, numbers: r.numbers, pictures, scaled: false, counts, differs: r.differs });
  }
  const out = new Array(list.length);
  let next = 0, done = 0;
  const lane = async () => {
    while (next < list.length) {
      const i = next++, [key, p] = list[i], ext = p.ext;
      const s = await once(p, b => small(b, ext, maxSide, quality).then(async x => ({ ...x, hash: await hashOf(b) })).catch(() => null));
      if (s) out[i] = [key, p.styleRaw, s.blob, s.w, s.h, s.hash];
      else counts.undrawn++;
      if (++done % 50 === 0 || done === list.length) say('scaling', done, list.length);
    }
  };
  await Promise.all(Array.from({ length: Math.min(SCALE_LANES, list.length) }, lane));
  const pictures = out.filter(Boolean);
  counts.pictures = pictures.length;
  answer({ ok: true, numbers: r.numbers, pictures, scaled: true, counts, differs: r.differs });
}

self.onmessage = e => {
  const msg = e.data || {};
  if (msg.type !== 'read') return;
  read(msg).catch(err => postMessage({ type: 'done', ok: false, error: `Reading the pictures failed: ${(err && err.message) || err}`,
    fileName: String(msg.fileName || ''), numbers: null, pictures: [], scaled: false, counts: {}, differs: [] }))
    .finally(() => close());
};
