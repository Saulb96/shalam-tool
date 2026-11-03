// The pictures in the ERP's weekly ATS report (Saul, Nov 2). The ERP can export the report with a picture per style in
// its first column ("Image"); that file is ~40 MB against ~160 KB without them, and unpacking all of it in JavaScript
// takes ~12 s. So the file is opened as the zip it is: only its XML parts are unpacked (with the platform's own
// DecompressionStream, not in JavaScript), the workbook is read from them alone (the same numbers as reading the whole
// file: pictures hold no cells), and then only the pictures that sit on a style's row are unpacked.
// - zipEntries(bytes): the zip's directory (name, method, sizes, offset), nothing unpacked.
// - inflateEntry(bytes, entry, { CFB }): one part's bytes (a Promise). Without DecompressionStream('deflate-raw'),
//   SheetJS's zip reader (XLSX.CFB) reads the whole file once instead (slow on a big report).
// - readPictures({ XLSX, InvAts, bytes, fileName, today, readOpts, onProgress }): the report's numbers (parseAts on the
//   same sheet) and its pictures, one per style; and the numbers' part of the file on its own (slim), for the page to
//   keep, so a report it has read once isn't downloaded again for its numbers.
// - pictureSize(bytes, ext): a picture's width and height from its first bytes.
// - BIG_REPORT (2 MB): a report this size or bigger is read in the pictures worker (picturesWorker.js); READ_OPTS: the
//   page's reading options; storedZip(parts) and crc32(bytes): the zip helpers this uses (tests use them too).
// A picture never comes from outside the file: a picture whose link points elsewhere (the ERP's picture server) is
// skipped and counted, and its address is never kept, shown or opened. A damaged or made-up file can't make it unpack
// or draw more than an ERP report could need (the limits below): what is past them is skipped, or the file not read.
// Pure: no network, no DOM. Loads in the page and in the pictures worker (window.InvErpPictures / self.InvErpPictures)
// and in Node (module.exports).
(function (root, factory) {
  const api = factory();
  if (typeof module === 'object' && module.exports) module.exports = api;
  root.InvErpPictures = api;
})(typeof globalThis !== 'undefined' ? globalThis : this, function () {
  'use strict';
  // A report this size or bigger is read off the page's main thread, for its pictures (a report without them is ~160 KB).
  const BIG_REPORT = 2 * 1024 * 1024;
  // As app.js reads every workbook (XOPT there), so the numbers are the page's.
  const READ_OPTS = { cellFormula: false, cellHTML: false, cellStyles: false };
  // Kinds of picture the browser can draw, told by their first bytes (never by the name alone).
  const KINDS = [['png', [0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]], ['jpeg', [0xff, 0xd8, 0xff]], ['gif', [0x47, 0x49, 0x46, 0x38]], ['bmp', [0x42, 0x4d]]];
  // Names that are never a picture the browser can draw, so they aren't even unpacked.
  const UNDRAWABLE = /\.(emf|wmf|tiff?|svg|pdf|bin|wdp|jxr)$/i;
  // Picture parts unpacked at the same time (one after another took ~0.8 s for ~2,000 on the real report; 16: ~0.4 s).
  const UNPACK_LANES = 16;
  // The limits, far beyond the real report's (each in brackets): one part unpacked (its drawing, ~2 MB); the XML parts
  // read for the numbers, all together (1.4 MB); one picture (67 KB); all the pictures (41 MB); a picture's pixels (a
  // few hundred across: drawing one of 25 million takes ~0.5 s and ~100 MB).
  const MAX_PART = 64 * 1024 * 1024, MAX_XML = 200 * 1024 * 1024;
  const MAX_PICTURE = 8 * 1024 * 1024, MAX_PICTURES = 300 * 1024 * 1024, MAX_PIXELS = 25 * 1000 * 1000;
  // Packing never shrinks data more than ~1,032 to 1: a part that says it unpacks to more is damaged or made up.
  const MAX_RATIO = 1100;
  const mb = n => `${Math.round(n / 1048576)} MB`;

  const utf8 = new TextDecoder('utf-8');
  const bad = msg => new Error(msg);

  // ---------- the zip

  const u8Of = b => (b instanceof Uint8Array ? b : b instanceof ArrayBuffer ? new Uint8Array(b) : ArrayBuffer.isView(b) ? new Uint8Array(b.buffer, b.byteOffset, b.byteLength) : null);
  const u64 = (v, at) => v.getUint32(at, true) + v.getUint32(at + 4, true) * 2 ** 32;

  // The zip's central directory: [{ name, method, flags, crc32, compressedSize, size, offset }] in its own order.
  // Throws (plain words) when the bytes aren't a zip or its directory doesn't fit in the file.
  function zipEntries(bytes) {
    const u = u8Of(bytes);
    if (!u) throw bad("the file's bytes are missing");
    const v = new DataView(u.buffer, u.byteOffset, u.byteLength), len = u.byteLength;
    // The end record: in the last 22 bytes, or before a comment of up to 64 KB.
    let end = -1;
    for (let i = len - 22; i >= 0 && i >= len - 22 - 0xffff; i--) {
      if (v.getUint32(i, true) === 0x06054b50) { end = i; break; }
    }
    if (end < 0) throw bad("it isn't a zip file (an .xlsx is one)");
    let count = v.getUint16(end + 10, true), cdSize = v.getUint32(end + 12, true), cdAt = v.getUint32(end + 16, true);
    // Zip64 (very big files): the real figures are in another record, pointed to just before the end record.
    if (count === 0xffff || cdSize === 0xffffffff || cdAt === 0xffffffff) {
      const loc = end - 20;
      if (loc < 0 || v.getUint32(loc, true) !== 0x07064b50) throw bad("its zip directory is damaged (no zip64 record)");
      const z = u64(v, loc + 8);
      if (z + 56 > len || v.getUint32(z, true) !== 0x06064b50) throw bad("its zip directory is damaged (bad zip64 record)");
      count = u64(v, z + 32); cdSize = u64(v, z + 40); cdAt = u64(v, z + 48);
    }
    if (cdAt + cdSize > len) throw bad('its zip directory is past the end of the file (is it cut short?)');
    const out = [];
    let p = cdAt;
    for (let k = 0; k < count; k++) {
      if (p + 46 > len || v.getUint32(p, true) !== 0x02014b50) throw bad(`its zip directory is damaged (entry ${k + 1} of ${count})`);
      const flags = v.getUint16(p + 8, true), method = v.getUint16(p + 10, true);
      const crc32 = v.getUint32(p + 16, true);
      let compressedSize = v.getUint32(p + 20, true), size = v.getUint32(p + 24, true);
      const nameLen = v.getUint16(p + 28, true), extraLen = v.getUint16(p + 30, true), commentLen = v.getUint16(p + 32, true);
      let offset = v.getUint32(p + 42, true);
      if (p + 46 + nameLen + extraLen > len) throw bad(`its zip directory is damaged (entry ${k + 1} of ${count})`);
      const name = utf8.decode(u.subarray(p + 46, p + 46 + nameLen)).replace(/^\/+/, '');
      // Zip64 sizes and offset, in that order, for the ones marked 0xffffffff.
      if (size === 0xffffffff || compressedSize === 0xffffffff || offset === 0xffffffff) {
        for (let e = p + 46 + nameLen, stop = e + extraLen; e + 4 <= stop;) {
          const id = v.getUint16(e, true), n = v.getUint16(e + 2, true);
          if (id === 0x0001) {
            let q = e + 4;
            if (size === 0xffffffff && q + 8 <= e + 4 + n) { size = u64(v, q); q += 8; }
            if (compressedSize === 0xffffffff && q + 8 <= e + 4 + n) { compressedSize = u64(v, q); q += 8; }
            if (offset === 0xffffffff && q + 8 <= e + 4 + n) offset = u64(v, q);
            break;
          }
          e += 4 + n;
        }
      }
      if (offset + 30 > len) throw bad(`"${name}" starts past the end of the file (is it cut short?)`);
      out.push({ name, method, flags, crc32, compressedSize, size, offset });
      p += 46 + nameLen + extraLen + commentLen;
    }
    return out;
  }

  // CRC-32 (the zip's own check on every part).
  let CRC = null;
  function crc32(u) {
    if (!CRC) {
      CRC = new Int32Array(256);
      for (let n = 0; n < 256; n++) { let c = n; for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1; CRC[n] = c; }
    }
    let c = -1;
    for (let i = 0, n = u.length; i < n; i++) c = CRC[(c ^ u[i]) & 0xff] ^ (c >>> 8);
    return (c ^ -1) >>> 0;
  }

  // Whether this platform unpacks raw deflate itself (browsers since 2022, Node 18+). Asked again only if it changes.
  let dsFor, dsOk = false;
  function canInflate() {
    const DS = typeof DecompressionStream === 'function' ? DecompressionStream : null;
    if (DS !== dsFor) {
      dsFor = DS; dsOk = false;
      if (DS) try { new DS('deflate-raw'); dsOk = true; } catch (e) { /* this platform only knows gzip and deflate */ }
    }
    return dsOk;
  }
  // Unpacked into the size the zip records; null as soon as it comes out bigger (a damaged or hostile part is never
  // let grow).
  async function inflateRaw(data, size) {
    const ds = new DecompressionStream('deflate-raw'), w = ds.writable.getWriter();
    // A broken stream shows up when the result is read; these two only stop it being reported twice.
    w.write(data).catch(() => {});
    w.close().catch(() => {});
    const out = new Uint8Array(size), rd = ds.readable.getReader();
    let at = 0;
    for (;;) {
      const { done, value } = await rd.read();
      if (done) return out.subarray(0, at);
      if (at + value.length > size) { rd.cancel().catch(() => {}); return null; }
      out.set(value, at); at += value.length;
    }
  }
  // Without DecompressionStream: SheetJS's zip reader unpacks the whole file once (every picture too) and parts come
  // from it.
  const whole = new WeakMap();
  function cfbPart(CFB, u, name) {
    let parts = whole.get(u);
    if (!parts) {
      const z = CFB.read(u, { type: 'array' });
      parts = new Map();
      z.FileIndex.forEach((f, i) => { if (f.type === 2) parts.set(z.FullPaths[i].replace(/^Root Entry\//, '').toLowerCase(), f.content); });
      whole.set(u, parts);
    }
    const c = parts.get(name.toLowerCase());
    if (c == null) throw bad(`"${name}" isn't in the file`);
    return Uint8Array.from(c);
  }

  // One part's bytes, unpacked: a copy of its own (never a view of the file's bytes). Checked against the size and the
  // CRC-32 the zip records; a part packed some other way, locked with a password, or damaged → an error saying so. The
  // size it says it unpacks to is checked before anything is set aside for it: never more than MAX_PART, nor more than
  // packing can give from its packed size.
  async function inflateEntry(bytes, entry, opts) {
    const u = u8Of(bytes);
    if (!u || !entry) throw bad("the file's bytes or the part are missing");
    const name = entry.name;
    if (entry.flags & 1) throw bad(`"${name}" is locked with a password`);
    if (entry.method !== 0 && entry.method !== 8) throw bad(`"${name}" is packed in a way this page can't unpack (zip method ${entry.method})`);
    if (entry.size > MAX_PART) throw bad(`"${name}" would be ${entry.size} bytes unpacked: too big for a part of a report`);
    if (entry.method === 8 && entry.size > entry.compressedSize * MAX_RATIO + 4096) {
      throw bad(`"${name}" says it unpacks to ${entry.size} bytes from ${entry.compressedSize}, more than packing can give: the file is damaged`);
    }
    const v = new DataView(u.buffer, u.byteOffset, u.byteLength);
    if (entry.offset + 30 > u.length || v.getUint32(entry.offset, true) !== 0x04034b50) throw bad(`"${name}" isn't where the zip directory says it is (the file may be damaged)`);
    const start = entry.offset + 30 + v.getUint16(entry.offset + 26, true) + v.getUint16(entry.offset + 28, true);
    if (start + entry.compressedSize > u.length) throw bad(`"${name}" runs past the end of the file (is it cut short?)`);
    const data = u.subarray(start, start + entry.compressedSize);
    let out;
    if (entry.method === 0) out = data.slice();
    else if (canInflate()) {
      try { out = await inflateRaw(data, entry.size); } catch (e) { throw bad(`"${name}" can't be unpacked: the file may be damaged (${(e && e.message) || e})`); }
      if (!out) throw bad(`"${name}" comes out bigger than the ${entry.size} bytes the zip records: the file may be damaged`);
    } else if (opts && opts.CFB && typeof opts.CFB.read === 'function') out = cfbPart(opts.CFB, u, name);
    else throw bad("this browser can't unpack the file (it has no DecompressionStream)");
    if (out.length !== entry.size) throw bad(`"${name}" came out ${out.length} bytes, not the ${entry.size} the zip records: the file may be damaged`);
    if (crc32(out) !== entry.crc32) throw bad(`"${name}" is damaged (its CRC-32 check doesn't match)`);
    return out;
  }

  // A zip of the given parts, stored (not packed): what SheetJS reads instead of the whole report.
  function storedZip(parts) {
    const enc = new TextEncoder(), heads = [];
    let size = 22;
    for (const [name, data] of parts) {
      const nb = enc.encode(name);
      heads.push({ nb, data, crc: crc32(data), ascii: nb.length === name.length });
      size += 30 + 46 + 2 * nb.length + data.length;
    }
    const out = new Uint8Array(size), v = new DataView(out.buffer);
    let p = 0;
    for (const h of heads) {
      h.at = p;
      v.setUint32(p, 0x04034b50, true); v.setUint16(p + 4, 20, true); v.setUint16(p + 6, h.ascii ? 0 : 0x800, true);
      v.setUint32(p + 14, h.crc, true); v.setUint32(p + 18, h.data.length, true); v.setUint32(p + 22, h.data.length, true);
      v.setUint16(p + 26, h.nb.length, true);
      out.set(h.nb, p + 30); out.set(h.data, p + 30 + h.nb.length);
      p += 30 + h.nb.length + h.data.length;
    }
    const cd = p;
    for (const h of heads) {
      v.setUint32(p, 0x02014b50, true); v.setUint16(p + 4, 20, true); v.setUint16(p + 6, 20, true); v.setUint16(p + 8, h.ascii ? 0 : 0x800, true);
      v.setUint32(p + 16, h.crc, true); v.setUint32(p + 20, h.data.length, true); v.setUint32(p + 24, h.data.length, true);
      v.setUint16(p + 28, h.nb.length, true); v.setUint32(p + 42, h.at, true);
      out.set(h.nb, p + 46);
      p += 46 + h.nb.length;
    }
    v.setUint32(p, 0x06054b50, true); v.setUint16(p + 8, heads.length, true); v.setUint16(p + 10, heads.length, true);
    v.setUint32(p + 12, p - cd, true); v.setUint32(p + 16, cd, true);
    return out;
  }

  // ---------- the parts' XML (read as text: a worker has no DOMParser)

  const unXml = s => String(s).replace(/&(?:#x([0-9a-f]+)|#(\d+)|(amp|lt|gt|quot|apos));/gi, (m, h, d, n) =>
    h ? String.fromCodePoint(parseInt(h, 16)) : d ? String.fromCodePoint(+d) : { amp: '&', lt: '<', gt: '>', quot: '"', apos: "'" }[n.toLowerCase()]);
  // An element's attributes as [[qualified name, value]] (prefixes vary between writers, so callers match local names).
  const attrList = s => [...String(s).matchAll(/([\w.:-]+)\s*=\s*(?:"([^"]*)"|'([^']*)')/g)].map(m => [m[1], unXml(m[2] ?? m[3])]);
  const local = q => q.slice(q.indexOf(':') + 1);
  const attr = (list, name) => { const a = list.find(([q]) => q === name || local(q) === name); return a ? a[1] : null; };
  // The relationship id attribute (r:id, r:embed, r:link): a prefixed name, so "sheetId" or a bare "id" never match.
  const relAttr = (list, name) => { const a = list.find(([q]) => q.includes(':') && local(q) === name); return a ? a[1] : null; };

  const dirOf = p => p.slice(0, p.lastIndexOf('/') + 1);
  const relsOf = p => `${dirOf(p)}_rels/${p.slice(p.lastIndexOf('/') + 1)}.rels`;
  // A relationship's target as a part name: absolute ("/xl/…") or relative to the part that points to it.
  function resolve(from, target) {
    let t = String(target).replace(/\\/g, '/');
    try { t = decodeURIComponent(t); } catch (e) { /* kept as written */ }
    const parts = (t.startsWith('/') ? t.slice(1) : dirOf(from) + t).split('/'), out = [];
    for (const x of parts) { if (x === '..') out.pop(); else if (x !== '.' && x !== '') out.push(x); }
    return out.join('/');
  }
  // A .rels part: Map(id → { type, target (a part name), external }). An external target is never kept: only that it
  // is one.
  function readRels(xml, from) {
    const m = new Map();
    for (const r of String(xml || '').matchAll(/<(?:[\w.-]+:)?Relationship\b([^>]*)>/g)) {
      const a = attrList(r[1]), id = attr(a, 'Id'), type = attr(a, 'Type') || '', target = attr(a, 'Target');
      if (id == null) continue;
      const external = /^external$/i.test(attr(a, 'TargetMode') || '');
      m.set(id, { type, external, target: external || target == null ? null : resolve(from, target) });
    }
    return m;
  }
  const relOfType = (rels, kind) => [...rels.values()].find(r => !r.external && r.type.endsWith('/' + kind));

  const kindOf = u => {
    for (const [k, sig] of KINDS) if (u.length >= sig.length && sig.every((b, i) => u[i] === b)) return k;
    if (u.length >= 12 && u[0] === 0x52 && u[1] === 0x49 && u[2] === 0x46 && u[3] === 0x46 && u[8] === 0x57 && u[9] === 0x45 && u[10] === 0x42 && u[11] === 0x50) return 'webp';
    return null;
  };
  const sameBytes = (a, b) => { if (a.length !== b.length) return false; for (let i = 0; i < a.length; i++) if (a[i] !== b[i]) return false; return true; };
  // A picture's width and height in pixels from its first bytes ({ w, h }), or null when they can't be told (the
  // browser then finds out for itself whether it can draw it). ext: as kindOf tells it.
  function pictureSize(u, ext) {
    const n = u ? u.length : 0, be16 = i => (u[i] << 8) | u[i + 1], le16 = i => u[i] | (u[i + 1] << 8);
    const be32 = i => ((u[i] << 24) | (u[i + 1] << 16) | (u[i + 2] << 8) | u[i + 3]) >>> 0, le32 = i => u[i] | (u[i + 1] << 8) | (u[i + 2] << 16) | (u[i + 3] << 24);
    const tag = (i, s) => n >= i + s.length && [...s].every((c, k) => u[i + k] === c.charCodeAt(0));
    if (ext === 'png') return tag(12, 'IHDR') && n >= 24 ? { w: be32(16), h: be32(20) } : null;
    if (ext === 'gif') return n >= 10 ? { w: le16(6), h: le16(8) } : null;
    if (ext === 'bmp') return n >= 26 ? (le32(14) === 12 ? { w: le16(18), h: le16(20) } : { w: Math.abs(le32(18)), h: Math.abs(le32(22)) }) : null;
    if (ext === 'webp') {
      if (tag(12, 'VP8X') && n >= 30) return { w: 1 + (u[24] | (u[25] << 8) | (u[26] << 16)), h: 1 + (u[27] | (u[28] << 8) | (u[29] << 16)) };
      if (tag(12, 'VP8 ') && n >= 30) return { w: le16(26) & 0x3fff, h: le16(28) & 0x3fff };
      if (tag(12, 'VP8L') && n >= 25) return { w: 1 + (((u[22] & 0x3f) << 8) | u[21]), h: 1 + (((u[24] & 0x0f) << 10) | (u[23] << 2) | (u[22] >> 6)) };
      return null;
    }
    if (ext === 'jpeg') {
      // Its markers in order, to the frame header (SOF0–SOF15 but DHT, JPG and DAC), which holds the height then the width.
      for (let i = 2; i + 9 <= n;) {
        if (u[i] !== 0xff) return null;
        const m = u[i + 1];
        if (m === 0xff) { i++; continue; }
        if (m === 0x01 || (m >= 0xd0 && m <= 0xd8)) { i += 2; continue; }
        if (m === 0xd9 || m === 0xda) return null;
        if (m >= 0xc0 && m <= 0xcf && m !== 0xc4 && m !== 0xc8 && m !== 0xcc) return { w: be16(i + 7), h: be16(i + 5) };
        i += 2 + be16(i + 2);
      }
    }
    return null;
  }
  // A .rels part with its links to outside the file taken out (the cells never need one), so the numbers' part the page
  // keeps holds no address (a sheet's links to the ERP's picture server among them).
  function noOutside(b) {
    const t = utf8.decode(b), out = t.replace(/<(?:[\w.-]+:)?Relationship\b[^>]*\bTargetMode\s*=\s*["']External["'][^>]*?(?:\/>|>\s*<\/(?:[\w.-]+:)?Relationship\s*>)/gi, '');
    return out === t ? b : new TextEncoder().encode(out);
  }
  // Parts the numbers never need and that may name a person (who saved the file, who wrote a comment): never unpacked
  // for them, so never kept either. The drawings are read on their own, for the pictures.
  const NOT_FOR_NUMBERS = /(^|\/)drawings\/|^docProps\/|^customXml\/|^xl\/(persons|threadedComments)\/|^xl\/comments[^/]*\.xml$/i;
  // An address left anywhere in the numbers' part (the XML's own namespace names aside): it isn't kept then.
  const hasAddress = b => /\b[a-z][a-z0-9+.-]*:\/\//i.test(utf8.decode(b).replace(/\s(?:xmlns(?::[\w.-]+)?|Type|Requires|uri)\s*=\s*(?:"[^"]*"|'[^']*')/g, ''));

  // ---------- the report's pictures

  // o: { XLSX, InvAts, bytes (Uint8Array or ArrayBuffer), fileName, today (a day number), readOpts (as app.js reads
  // workbooks; READ_OPTS by default), onProgress(done, total) (every 50 pictures looked at, optional), onNumbers(numbers,
  // slim) (once the numbers are read, before any picture is: the page shows them without waiting for the pictures,
  // optional), numbersOnly (the numbers alone: no picture is looked for, as when the page already has this report's
  // pictures) }.
  // Answers { ok, error, numbers, slim, pictures, differs, counts }:
  // - numbers: InvAts.parseAts on the report, exactly as reading the whole file gives it (its own ok and error), or null
  //   when the file can't be opened at all.
  // - slim: the numbers' part of the file on its own (a zip of its XML but the drawings, ~1.5 MB of the real ~40 MB),
  //   for the page to keep: read again with readPictures, it gives the same numbers. Nothing in it names a person (who
  //   saved the file, who wrote a comment) or holds an address (links to outside the file are taken out); null when the
  //   file isn't an .xlsx, or when an address is left in it anyway (it isn't kept then).
  // - pictures: Map(style key, as parseAts keys styles → { styleRaw, bytes (Uint8Array), ext ('png'|'jpeg'|'gif'|'bmp'|
  //   'webp'), row (the sheet row, from 1) }); one per style: the same picture on several rows of a style is one; when
  //   they differ, the first by row is kept and the style is in differs ([{ style, styleRaw, rows }]).
  // - counts: anchors (every drawing anchor on the sheet), pictures (styles with one), external (linked to a picture
  //   outside the file: skipped), unsupported (a kind the browser can't draw: EMF, WMF, TIFF, SVG, …), noStyleRow (on
  //   the title, header, a blank or Grand Total row, or above the header), differs, absolute (anchored to no row),
  //   notPicture (a shape or chart), missing (its picture isn't in the file), damaged (its picture can't be unpacked),
  //   tooBig (its picture is bigger than any ERP report's: over 8 MB, or 25 million pixels; skipped), bytesIn (the kept
  //   pictures' bytes).
  // ok false (error in plain words) only when the file can't be opened, its rows can't be matched to styles, or its XML
  // or its pictures are far bigger than an ERP report's (200 MB, 300 MB): a report with no pictures is ok with 0.
  async function readPictures(o) {
    const { XLSX, InvAts } = o || {};
    const fileName = String((o && o.fileName) || 'the ATS report');
    const counts = { anchors: 0, pictures: 0, external: 0, unsupported: 0, noStyleRow: 0, differs: 0, absolute: 0, notPicture: 0, missing: 0,
      damaged: 0, tooBig: 0, bytesIn: 0 };
    const res = { ok: false, error: '', numbers: null, slim: null, pictures: new Map(), differs: [], counts };
    const fail = msg => Object.assign(res, { ok: false, error: msg, pictures: new Map(), differs: [] });
    if (!XLSX || !InvAts) return fail("The picture reader's parts didn't load (the spreadsheet library or core/ats.js).");
    const u = u8Of(o.bytes);
    if (!u || !u.length) return fail(`${fileName} is empty (0 bytes).`);
    const readOpts = Object.assign({}, READ_OPTS, (o && o.readOpts) || {}, { type: 'array' });
    const inflateOpts = { CFB: XLSX.CFB };
    const progress = typeof o.onProgress === 'function' ? o.onProgress : null;
    const book = b => {
      try { return XLSX.read(b, readOpts); }
      catch (e) { throw bad(`Couldn't open ${fileName} as an Excel workbook (${(e && e.message) || e}).`); }
    };
    const numbersOf = wb => {
      try { return InvAts.parseAts(XLSX, wb, fileName, o.today); }
      catch (e) { throw bad(`Reading ${fileName} failed: ${(e && e.message) || e}`); }
    };
    // Not a zip (an old .xls): read whole as the page reads it, no pictures looked for.
    let entries;
    try { entries = zipEntries(u); }
    catch (e) {
      try { res.numbers = numbersOf(book(u)); } catch (e2) { return fail(e2.message); }
      return Object.assign(res, { ok: true });
    }
    const byName = new Map(entries.filter(e => !e.name.endsWith('/')).map(e => [e.name.toLowerCase(), e]));
    const part = name => byName.get(String(name).toLowerCase()) || null;
    const unpacked = new Map();   // the XML parts already unpacked for SheetJS, so none is unpacked twice
    const textOf = async name => {
      const e = part(name);
      return e ? utf8.decode(unpacked.get(e.name) || await inflateEntry(u, e, inflateOpts)) : null;
    };

    // The numbers: the workbook part (from the package's own rels); a binary workbook (.xlsb) is read whole, with no
    // pictures looked for. Otherwise every XML part but the drawings and the ones that may name a person (the cells
    // never need them), unpacked, links to outside the file taken out of their rels, in a stored zip SheetJS reads
    // quickly: the numbers' part, kept by the page. Their sizes are checked first: never more than MAX_XML in all.
    let wbPart;
    try {
      wbPart = (relOfType(readRels(await textOf('_rels/.rels'), ''), 'officeDocument') || {}).target || 'xl/workbook.xml';
      if (!/\.xml$/i.test(wbPart) || !part(wbPart)) {
        res.numbers = numbersOf(book(u));
        return Object.assign(res, { ok: true });
      }
      const xml = entries.filter(e => !e.name.endsWith('/') && /\.(xml|rels|vml)$/i.test(e.name) && !NOT_FOR_NUMBERS.test(e.name));
      const total = xml.reduce((t, e) => t + e.size, 0);
      if (total > MAX_XML) return fail(`${fileName} can't be read: its cells would be ${mb(total)} unpacked, far more than an ATS report's (a few MB), so the file is damaged or isn't an ATS report.`);
      const slim = [];
      let address = false;
      for (const e of xml) {
        let b = await inflateEntry(u, e, inflateOpts);
        unpacked.set(e.name, b);
        if (/\.rels$/i.test(e.name)) b = noOutside(b);
        address = address || hasAddress(b);
        slim.push([e.name, b]);
      }
      const z = storedZip(slim);
      res.numbers = numbersOf(book(z));
      res.slim = address ? null : z;
    } catch (e) {
      return fail(/^Couldn't open |^Reading /.test(e.message) ? e.message : `Couldn't open ${fileName} as an Excel workbook (${e.message}).`);
    }
    const numbers = res.numbers;
    if (!numbers || typeof numbers !== 'object') return fail(`Reading ${fileName} gave nothing back.`);
    if (typeof o.onNumbers === 'function') o.onNumbers(numbers, res.slim);
    if (o.numbersOnly) return Object.assign(res, { ok: true });

    // The pictures.
    try {
      // Its rows, as parseAts read them: sheet row (from 1) → the style on it. Without a header there's none.
      if (numbers.headerRow == null) return fail(`The pictures in ${fileName} can't be matched to styles: ${numbers.error || 'its header row wasn\'t found'}`);
      const rowStyle = new Map();
      if (numbers.styles instanceof Map) for (const s of numbers.styles.values()) for (const r of s.rows || []) rowStyle.set(r.row, s);

      // The sheet parseAts read → its drawing → the anchors.
      const wbText = await textOf(wbPart);
      const wbRels = readRels(await textOf(relsOf(wbPart)), wbPart);
      let sheetPart = null;
      for (const m of String(wbText || '').matchAll(/<(?:[\w.-]+:)?sheet\b([^>]*)>/g)) {
        const a = attrList(m[1]);
        if (attr(a, 'name') === numbers.sheet) { const r = wbRels.get(relAttr(a, 'id')); sheetPart = r && !r.external ? r.target : null; break; }
      }
      if (!sheetPart || !part(sheetPart)) return fail(`The pictures in ${fileName} can't be found: its sheet "${numbers.sheet}" isn't where the workbook says.`);
      const sheetRels = readRels(await textOf(relsOf(sheetPart)), sheetPart);
      // The sheet's own <drawing r:id> names its drawing; without one the sheet has no pictures.
      const sheetXml = await textOf(sheetPart);
      const dm = /<(?:[\w.-]+:)?drawing\b([^>]*)>/.exec(sheetXml || '');
      const dRel = dm ? sheetRels.get(relAttr(attrList(dm[1]), 'id')) : null;
      if (!dRel || dRel.external || !part(dRel.target)) return Object.assign(res, { ok: true });
      const drawPart = dRel.target;
      const drawRels = readRels(await textOf(relsOf(drawPart)), drawPart);
      // A writer may give a second copy of an anchor for old Excel (mc:Fallback): only the first counts.
      const drawXml = (await textOf(drawPart)).replace(/<([\w.-]+:)?Fallback\b[^>]*>[\s\S]*?<\/\1Fallback>/g, '');

      // Each anchor: its row and the part its picture is in, before anything is unpacked.
      const found = [];
      for (const m of drawXml.matchAll(/<([\w.-]+:)?(twoCellAnchor|oneCellAnchor|absoluteAnchor)\b[^>]*>([\s\S]*?)<\/\1\2>/g)) {
        counts.anchors++;
        const body = m[3], blip = /<(?:[\w.-]+:)?blip\b([^>]*)>/.exec(body);
        if (!blip) { counts.notPicture++; continue; }
        if (m[2] === 'absoluteAnchor') { counts.absolute++; continue; }
        // A picture linked (r:link), or embedded from a relationship that points outside the file: never opened.
        const a = attrList(blip[1]), embed = relAttr(a, 'embed');
        const rel = embed != null ? drawRels.get(embed) : null;
        if ((rel && rel.external) || (embed == null && relAttr(a, 'link') != null)) { counts.external++; continue; }
        const from = /<(?:[\w.-]+:)?from\b[^>]*>([\s\S]*?)<\/(?:[\w.-]+:)?from>/.exec(body);
        const rowM = from && /<(?:[\w.-]+:)?row>\s*(\d+)\s*</.exec(from[1]);
        const row = rowM ? +rowM[1] + 1 : null;
        const s = row == null ? null : rowStyle.get(row);
        if (!s) { counts.noStyleRow++; continue; }
        if (!rel || !rel.target || !part(rel.target)) { counts.missing++; continue; }
        if (UNDRAWABLE.test(rel.target)) { counts.unsupported++; continue; }
        // Bigger than any ERP picture, by the size the zip records: never unpacked.
        if (part(rel.target).size > MAX_PICTURE) { counts.tooBig++; continue; }
        found.push({ row, s, target: rel.target });
      }

      // Each picture part unpacked once, however many rows show it, a few at a time (the platform unpacks them in
      // parallel): { bytes, ext }, null when the browser can't draw it, 'damaged' when it can't be unpacked, 'tooBig' when
      // its first bytes say it has more than MAX_PIXELS (that picture only is skipped and counted; the rest are still
      // read). All of them together, by the sizes the zip records, never more than MAX_PICTURES.
      const targets = [...new Set(found.map(f => f.target))], media = new Map();
      const all = targets.reduce((t, x) => t + part(x).size, 0);
      if (all > MAX_PICTURES) return fail(`The pictures in ${fileName} would be ${mb(all)} unpacked, far more than an ERP report's (about 40 MB), so they aren't read.`);
      let next = 0, done = 0;
      const lane = async () => {
        while (next < targets.length) {
          const target = targets[next++];
          let b = null;
          try { b = await inflateEntry(u, part(target), inflateOpts); } catch (e) { media.set(target, 'damaged'); }
          if (b) {
            const ext = kindOf(b), d = ext && pictureSize(b, ext);
            media.set(target, !ext ? null : d && d.w * d.h > MAX_PIXELS ? 'tooBig' : { bytes: b, ext });
          }
          if (progress && (++done % 50 === 0 || done === targets.length)) progress(done, targets.length);
        }
      };
      await Promise.all(Array.from({ length: Math.min(UNPACK_LANES, targets.length) }, lane));
      // By row, so a style's first picture is the one kept; another one on the style is the same picture or differs.
      found.sort((a, b) => a.row - b.row);
      const differs = new Map();
      for (const f of found) {
        const key = f.s.style, kept = res.pictures.get(key), m = media.get(f.target);
        if (kept && kept.target === f.target) continue;
        if (m === 'damaged') counts.damaged++;
        else if (m === 'tooBig') counts.tooBig++;
        else if (!m) counts.unsupported++;
        else if (!kept) res.pictures.set(key, { styleRaw: f.s.styleRaw, bytes: m.bytes, ext: m.ext, row: f.row, target: f.target });
        else if (!sameBytes(kept.bytes, m.bytes)) {
          const d = differs.get(key) || differs.set(key, { style: key, styleRaw: f.s.styleRaw, rows: [kept.row] }).get(key);
          d.rows.push(f.row);
        }
      }
      for (const p of res.pictures.values()) { delete p.target; counts.bytesIn += p.bytes.length; }
      res.differs = [...differs.values()];
      counts.differs = res.differs.length;
      counts.pictures = res.pictures.size;
      return Object.assign(res, { ok: true });
    } catch (e) {
      return fail(`The pictures in ${fileName} couldn't be read: ${(e && e.message) || e}.`);
    }
  }

  return { zipEntries, inflateEntry, readPictures, pictureSize, storedZip, crc32, BIG_REPORT, READ_OPTS };
});
