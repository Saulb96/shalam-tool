import test from 'node:test';
import assert from 'node:assert/strict';
import { deflateRawSync } from 'node:zlib';
import { core, loadXLSX, loadExcelJS } from './load.mjs';

const C = core('common');
const A = core('ats');
const EP = core('erpPictures');
const X = loadXLSX();
const TODAY = C.ymd(2026, 11, 2);
// The ERP's file name as SharePoint saves it (underscores for spaces), with an invented company.
const NAME = 'Example_Co_Inventory_Report_-_On_Hand_In_Transit_In_Production_11-02-2026_04-59_PM.xlsx';
// As app.js reads files (a new object each time: SheetJS adds its defaults to the one it is given).
const xopt = () => ({ type: 'array', cellFormula: false, cellHTML: false, cellStyles: false });
const whole = (bytes, name = NAME) => A.parseAts(X, X.read(bytes, xopt()), name, TODAY);
const read = (bytes, o = {}) => EP.readPictures({ XLSX: X, InvAts: A, bytes, fileName: NAME, today: TODAY, ...o });

// ---------- invented data (styles PC1xx, warehouses WH A / WH B, invented pictures)

const cat = (...parts) => Uint8Array.from(Buffer.concat(parts.map(p => Buffer.from(p))));
// A 1×1 PNG and a 2×1 one; the rest are only the first bytes of their kind (enough to tell what they are; a GIF's size,
// 1×1, is in them) and a tag.
const PNG1 = Uint8Array.from(Buffer.from('iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==', 'base64'));
const PNG2 = Uint8Array.from(Buffer.from('iVBORw0KGgoAAAANSUhEUgAAAAIAAAABCAYAAAD0In+KAAAAEUlEQVR42mP8z8DAwMDAwAAAD/4CAVv6FX4AAAAASUVORK5CYII=', 'base64'));
const PNG_E = cat([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a], 'invented picture E');
const GIF_B = cat('GIF89a', [1, 0, 1, 0], 'invented picture B');
const JPG_C = cat([0xff, 0xd8, 0xff, 0xe0], 'invented picture C');
const TIFF = cat([0x49, 0x49, 0x2a, 0x00], 'invented TIFF saved as .png');
const EMF = cat([0x01, 0x00, 0x00, 0x00, 0x6c, 0x00, 0x00, 0x00], 'invented EMF');
// The external picture's link, as the ERP writes it: a picture server address with a signature in it.
const LINK = 'https://pictures.example.invalid/bucket/PC104.png?X-Amz-Credential=INVENTEDKEY123&amp;X-Amz-Signature=deadbeef';
const LEAKS = ['pictures.example', 'X-Amz', 'INVENTEDKEY123', 'deadbeef'];

const TITLE = ['Inventory Report - On Hand_In Transit_In Production'];
const HEAD = ['Image', 'Color Code', 'Style', 'Style Description', 'Category', 'Sub Category', 'Warehouse', 'In WH',
  'Total Customer Orders', 'ATS from in WH', 'In Trans Qty', 'ETA Date'];
const row = (style, wh, inWh, orders, o = {}) => [null, 'XX', style, 'Item ' + style, 'Lip', 'Gloss', wh, inWh, orders, String(inWh - orders), o.inTrans ?? null, o.eta ?? null];
// The title on row 1, the header on row 2, styles on rows 3–13 (0-based 2–12), row 9 blank, the Grand Total on row 14.
const ROWS = [TITLE, HEAD,
  row('PC101', 'WH A', 120, 20), row('PC102', 'WH A', 60, 10), row('PC102', 'WH B', 40, 0), row('PC103', 'WH A', 10, 0), row('pc103 ', 'WH B', 6, 1),
  row('PC104', 'WH A', 5, 5, { inTrans: '1,200, 300', eta: '11/20/2026, 12/1/2026' }), [],
  row('PC105', 'WH A', 30, 0), row('PC105', 'WH B', 20, 0), row('PC106', 'WH A', 7, 0), row('PC107', 'WH A', 8, 0),
  [null, null, 'Grand Total', null, null, null, null, 306, 36, null, null, null],
];

// An invented ERP report (sheet "Report", in the ERP's layout: "Image" first in the header), pictures placed by ExcelJS
// in column A: pics [[0-based sheet row, image name, 'one' for a one-cell anchor]]; each image name is its own picture
// part (two names with the same bytes are two parts). Images are added in the order they're first placed: ExcelJS
// links a picture placed twice to the wrong part otherwise.
async function report(pics = [], images = {}, rows = ROWS, edit = null) {
  const E = loadExcelJS(), wb = new E.Workbook(), ws = wb.addWorksheet('Report');
  for (const r of rows) ws.addRow(r);
  if (edit) edit(wb, ws);
  const ids = {};
  for (const k of new Set([...pics.map(p => p[1]), ...Object.keys(images)])) {
    ids[k] = wb.addImage({ buffer: images[k], extension: k.startsWith('gif') ? 'gif' : k.startsWith('jpg') ? 'jpeg' : 'png' });
  }
  for (const [r, k, one] of pics) {
    ws.addImage(ids[k], one ? { tl: { col: 0, row: r }, ext: { width: 40, height: 40 } } : { tl: { col: 0, row: r }, br: { col: 1, row: r + 1 } });
  }
  return new Uint8Array(await wb.xlsx.writeBuffer());
}
const text = b => Buffer.from(b).toString('utf8');
// Every part of a zip, unpacked: Map(name → bytes).
async function partsOf(bytes) {
  const m = new Map();
  for (const e of EP.zipEntries(bytes)) if (!e.name.endsWith('/')) m.set(e.name, await EP.inflateEntry(bytes, e));
  return m;
}
// The parts written back packed, by SheetJS's own zip writer (another deflate than ExcelJS's).
function zipOf(parts) {
  const z = X.CFB.utils.cfb_new();
  for (const [n, b] of parts) X.CFB.utils.cfb_add(z, n, b);
  return new Uint8Array(X.CFB.write(z, { fileType: 'zip', type: 'array', compression: true }));
}
const mediaNamed = (parts, bytes) => [...parts.keys()].find(n => n.startsWith('xl/media/') && Buffer.compare(Buffer.from(parts.get(n)), Buffer.from(bytes)) === 0);

// The main invented report: every case at once.
// - PC101: a picture (two-cell anchor), and the same one again in an mc:AlternateContent with a Fallback copy.
// - PC102: one picture part on both of its rows (one-cell anchors). PC103: two parts with the same bytes ("pc103 " on
//   its second row is the same style to parseAts).
// - PC104: its picture is linked to the ERP's picture server (External): skipped, its address never kept.
// - PC105: two different pictures, the second row's placed first in the drawing: the first by row is kept (differs).
// - PC106: an EMF; PC107: TIFF bytes in a .png part: the browser can't draw either.
// - On the header, the blank row and the Grand Total: skipped. One picture anchored to no row; one shape (no picture).
async function mainReport() {
  const images = { pngA: PNG1, gifB: GIF_B, jpgC: JPG_C, jpgC2: JPG_C, pngD: PNG2, pngE: PNG_E, pngExt: cat('to be linked'), pngEmf: cat('to be EMF'), pngTiff: TIFF, pngHead: PNG_E };
  const pics = [[1, 'pngHead'], [2, 'pngA'], [3, 'gifB', 'one'], [4, 'gifB', 'one'], [5, 'jpgC'], [6, 'jpgC2'], [7, 'pngExt'], [8, 'pngA'],
    [10, 'pngE'], [9, 'pngD'], [11, 'pngEmf'], [12, 'pngTiff'], [13, 'pngA']];
  const parts = await partsOf(await report(pics, images));
  // PC104's picture: linked to an address outside the file, and its part gone (as in the ERP's report).
  const ext = mediaNamed(parts, images.pngExt), emf = mediaNamed(parts, images.pngEmf), relsName = 'xl/drawings/_rels/drawing1.xml.rels';
  let rels = text(parts.get(relsName));
  rels = rels.replace(`Target="../media/${ext.slice(9)}"`, `Target="${LINK}" TargetMode="External"`);
  parts.delete(ext);
  // PC106's: an EMF part.
  rels = rels.replace(`Target="../media/${emf.slice(9)}"`, 'Target="../media/image90.emf"');
  parts.delete(emf); parts.set('xl/media/image90.emf', EMF);
  parts.set(relsName, Buffer.from(rels));
  parts.set('[Content_Types].xml', Buffer.from(text(parts.get('[Content_Types].xml')).replace('<Default ', '<Default Extension="emf" ContentType="image/x-emf"/><Default ')));
  // A picture anchored to no row, a shape on PC101's row, and PC101's picture again with a Fallback copy for old Excel.
  const R = 'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"';
  const rid = /Id="(rId\d+)"[^>]*Target="\.\.\/media\/([^"]+)"/g, ridOf = name => [...rels.matchAll(rid)].find(m => m[2] === name.slice(9))[1];
  const pic = (id, r) => `<xdr:pic><xdr:nvPicPr><xdr:cNvPr id="${id}" name="Picture ${id}"/><xdr:cNvPicPr/></xdr:nvPicPr><xdr:blipFill><a:blip ${R} r:embed="${r}"/></xdr:blipFill><xdr:spPr/></xdr:pic>`;
  const cell = (tag, r, c = 0) => `<xdr:${tag}><xdr:col>${c}</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>${r}</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:${tag}>`;
  const a = ridOf(mediaNamed(parts, PNG1));
  const two = (body, r) => `<xdr:twoCellAnchor>${cell('from', r)}${cell('to', r + 1, 1)}${body}<xdr:clientData/></xdr:twoCellAnchor>`;
  const extra = `<xdr:absoluteAnchor><xdr:pos x="0" y="0"/><xdr:ext cx="9525" cy="9525"/>${pic(90, a)}<xdr:clientData/></xdr:absoluteAnchor>` +
    two('<xdr:sp><xdr:nvSpPr><xdr:cNvPr id="91" name="Shape 91"/><xdr:cNvSpPr/></xdr:nvSpPr><xdr:spPr/></xdr:sp>', 2) +
    `<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"><mc:Choice Requires="a14">${two(pic(92, a), 2)}</mc:Choice>` +
    `<mc:Fallback>${two(pic(93, a), 2)}</mc:Fallback></mc:AlternateContent>`;
  parts.set('xl/drawings/drawing1.xml', Buffer.from(text(parts.get('xl/drawings/drawing1.xml')).replace('</xdr:wsDr>', extra + '</xdr:wsDr>')));
  return zipOf(parts);
}
// A copy of a zip with one part's packed bytes overwritten (0xff: never valid deflate).
function spoil(bytes, name) {
  const e = EP.zipEntries(bytes).find(x => x.name === name), out = Uint8Array.from(bytes), v = new DataView(out.buffer);
  const start = e.offset + 30 + v.getUint16(e.offset + 26, true) + v.getUint16(e.offset + 28, true);
  out.fill(0xff, start, start + e.compressedSize);
  return out;
}
// A copy of a zip whose directory (and the part's own header) says one part unpacks to size bytes: what a damaged or
// made-up file can say.
function claim(bytes, name, size) {
  const out = Uint8Array.from(bytes), v = new DataView(out.buffer), len = out.length;
  let end = len - 22;
  while (v.getUint32(end, true) !== 0x06054b50) end--;
  for (let p = v.getUint32(end + 16, true), k = v.getUint16(end + 10, true); k > 0; k--) {
    const nl = v.getUint16(p + 28, true), el = v.getUint16(p + 30, true), cl = v.getUint16(p + 32, true);
    if (Buffer.from(out.subarray(p + 46, p + 46 + nl)).toString('utf8') === name) {
      v.setUint32(p + 24, size, true);
      v.setUint32(v.getUint32(p + 42, true) + 22, size, true);
      return out;
    }
    p += 46 + nl + el + cl;
  }
  throw new Error(`no ${name} in the zip`);
}
// A zip of the given parts, each packed with zlib's deflate at its tightest (a part of 70 MB of spaces packs to ~70 KB).
function deflatedZip(parts) {
  const locals = [], dir = [];
  let at = 0;
  for (const [name, data] of parts) {
    const nb = Buffer.from(name), packed = deflateRawSync(data, { level: 9 }), crc = EP.crc32(data);
    const head = Buffer.alloc(30), cd = Buffer.alloc(46);
    head.writeUInt32LE(0x04034b50, 0); head.writeUInt16LE(20, 4); head.writeUInt16LE(8, 8); head.writeUInt32LE(crc, 14);
    head.writeUInt32LE(packed.length, 18); head.writeUInt32LE(data.length, 22); head.writeUInt16LE(nb.length, 26);
    cd.writeUInt32LE(0x02014b50, 0); cd.writeUInt16LE(20, 4); cd.writeUInt16LE(20, 6); cd.writeUInt16LE(8, 10); cd.writeUInt32LE(crc, 16);
    cd.writeUInt32LE(packed.length, 20); cd.writeUInt32LE(data.length, 24); cd.writeUInt16LE(nb.length, 28); cd.writeUInt32LE(at, 42);
    locals.push(head, nb, packed); dir.push(cd, nb);
    at += 30 + nb.length + packed.length;
  }
  const cdBytes = Buffer.concat(dir), eocd = Buffer.alloc(22);
  eocd.writeUInt32LE(0x06054b50, 0); eocd.writeUInt16LE(parts.length, 8); eocd.writeUInt16LE(parts.length, 10);
  eocd.writeUInt32LE(cdBytes.length, 12); eocd.writeUInt32LE(at, 16);
  return new Uint8Array(Buffer.concat([...locals, cdBytes, eocd]));
}
// A PNG's first bytes saying it is w × h (no picture after them).
function pngHead(w, h) {
  const ihdr = Buffer.alloc(25);
  ihdr.writeUInt32BE(13, 0); ihdr.write('IHDR', 4); ihdr.writeUInt32BE(w, 8); ihdr.writeUInt32BE(h, 12); ihdr[16] = 8; ihdr[17] = 2;
  return cat([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a], ihdr, 'invented picture, nothing after its header');
}
// Everything an answer holds, as text (Maps, Sets and bytes included), to scan for what must never be in it.
const dump = v => JSON.stringify(v, (k, x) => (x instanceof Map ? [...x] : x instanceof Set ? [...x] : x instanceof Uint8Array ? Buffer.from(x).toString('latin1') : x));

// ---------- the zip

test('zipEntries lists the parts without unpacking them; inflateEntry unpacks packed (deflate) and stored parts', async () => {
  const bytes = await report([[2, 'a']], { a: PNG1 });   // ExcelJS packs every part
  const es = EP.zipEntries(bytes);
  const names = es.map(e => e.name);
  for (const n of ['[Content_Types].xml', '_rels/.rels', 'xl/workbook.xml', 'xl/worksheets/sheet1.xml', 'xl/drawings/drawing1.xml', 'xl/media/image1.png']) assert.ok(names.includes(n), n);
  const png = es.find(e => e.name === 'xl/media/image1.png');
  assert.equal(png.method, 8);
  assert.equal(png.size, PNG1.length);
  assert.ok(png.offset >= 0 && png.compressedSize > 0 && typeof png.crc32 === 'number');
  assert.deepEqual(await EP.inflateEntry(bytes, png), PNG1);
  assert.match(text(await EP.inflateEntry(bytes.buffer, es.find(e => e.name === 'xl/workbook.xml'))), /<sheet [^>]*name="Report"/);
  // Its own copy, never a view of the file's bytes.
  const got = await EP.inflateEntry(bytes, es.find(e => e.name === 'xl/workbook.xml'));
  assert.notEqual(got.buffer, bytes.buffer);

  // Stored parts (method 0), a name that isn't plain ASCII, and an empty part.
  const parts = [['a.txt', Buffer.from('stored text')], ['dossier/é.xml', Buffer.from('<x>é</x>')], ['empty.bin', new Uint8Array(0)]];
  const z = EP.storedZip(parts), ze = EP.zipEntries(z);
  assert.deepEqual(ze.map(e => [e.name, e.method, e.size]), [['a.txt', 0, 11], ['dossier/é.xml', 0, 9], ['empty.bin', 0, 0]]);
  for (const [i, e] of ze.entries()) assert.deepEqual(Buffer.from(await EP.inflateEntry(z, e)), Buffer.from(parts[i][1]));
  // And SheetJS reads such a stored zip as a workbook.
  const slim = await partsOf(bytes);
  for (const n of [...slim.keys()]) if (!/\.(xml|rels)$/.test(n)) slim.delete(n);
  assert.equal(X.read(EP.storedZip([...slim]), xopt()).Sheets.Report.C3.v, 'PC101');
  // SheetJS's own deflate too.
  const re = zipOf(await partsOf(bytes)), rp = EP.zipEntries(re).find(e => e.name === 'xl/media/image1.png');
  assert.equal(rp.method, 8);
  assert.deepEqual(await EP.inflateEntry(re, rp), PNG1);
});

test('a file that is not a zip, cut short, packed another way, locked or damaged: an error that says so', async () => {
  assert.throws(() => EP.zipEntries(Buffer.from('Style,In WH\nPC101,5\n')), /isn't a zip file/);
  assert.throws(() => EP.zipEntries(new Uint8Array(10)), /isn't a zip file/);
  const bytes = await report([[2, 'a']], { a: PNG1 });
  // Cut short: its directory (at the end) is gone.
  assert.throws(() => EP.zipEntries(bytes.subarray(0, bytes.length - 30)), /isn't a zip file/);
  const es = EP.zipEntries(bytes), png = es.find(e => e.name === 'xl/media/image1.png');
  await assert.rejects(EP.inflateEntry(bytes, { ...png, method: 12 }), /"xl\/media\/image1\.png" is packed in a way this page can't unpack \(zip method 12\)/);
  await assert.rejects(EP.inflateEntry(bytes, { ...png, flags: 1 }), /is locked with a password/);
  await assert.rejects(EP.inflateEntry(bytes, { ...png, crc32: (png.crc32 + 1) >>> 0 }), /is damaged \(its CRC-32 check doesn't match\)/);
  await assert.rejects(EP.inflateEntry(bytes, { ...png, size: png.size + 1 }), /came out 70 bytes, not the 71 the zip records/);
  await assert.rejects(EP.inflateEntry(bytes, { ...png, size: png.size - 1 }), /"xl\/media\/image1\.png" comes out bigger than the 69 bytes the zip records: the file may be damaged/);
  await assert.rejects(EP.inflateEntry(bytes, { ...png, size: 2 ** 30 }), /would be 1073741824 bytes unpacked: too big for a part of a report/);
  await assert.rejects(EP.inflateEntry(bytes, { ...png, offset: png.offset + 1 }), /isn't where the zip directory says it is/);
  await assert.rejects(EP.inflateEntry(bytes.subarray(0, png.offset + 40), png), /runs past the end of the file/);
  // Packed bytes that aren't deflate.
  await assert.rejects(EP.inflateEntry(spoil(bytes, png.name), png), /^Error: "xl\/media\/image1\.png" can't be unpacked: the file may be damaged \(/);
});

// ---------- the report's pictures

test('readPictures: one picture per style from the rows they sit on; links, undrawable kinds and rows with no style skipped and counted', async () => {
  const bytes = await mainReport();
  let calls = 0;
  const r = await read(bytes, { onProgress: () => calls++ });
  assert.equal(r.ok, true, r.error);
  assert.equal(r.error, '');
  assert.deepEqual([...r.pictures.keys()].sort(), ['PC101', 'PC102', 'PC103', 'PC105']);
  const p = k => { const x = r.pictures.get(k); return [x.styleRaw, x.ext, x.row, Buffer.from(x.bytes).equals(Buffer.from({ PC101: PNG1, PC102: GIF_B, PC103: JPG_C, PC105: PNG2 }[k]))]; };
  assert.deepEqual(p('PC101'), ['PC101', 'png', 3, true]);
  assert.deepEqual(p('PC102'), ['PC102', 'gif', 4, true]);     // one part on both rows: one picture
  assert.deepEqual(p('PC103'), ['PC103', 'jpeg', 6, true]);    // two parts with the same bytes: one picture, not "differs"
  assert.deepEqual(p('PC105'), ['PC105', 'png', 10, true]);    // the first by row, though the other was placed first
  assert.ok(r.pictures.get('PC101').bytes instanceof Uint8Array);
  assert.deepEqual(r.differs, [{ style: 'PC105', styleRaw: 'PC105', rows: [10, 11] }]);
  assert.deepEqual(r.counts, { anchors: 16, pictures: 4, external: 1, unsupported: 2, noStyleRow: 3, differs: 1, absolute: 1, notPicture: 1,
    missing: 0, damaged: 0, tooBig: 0, bytesIn: PNG1.length + GIF_B.length + JPG_C.length + PNG2.length });
  assert.ok(calls >= 1, 'progress was reported');
  // The numbers are parseAts's on the same sheet, as reading the whole file gives them (and the date from the name).
  assert.equal(r.numbers.ok, true, r.numbers.error);
  assert.deepEqual(r.numbers, whole(bytes));
  assert.equal(r.numbers.reportDay, C.ymd(2026, 11, 2));
  assert.equal(r.numbers.styles.get('PC104').inTrans, 1500);
  // The picture server's address is nowhere in the answer.
  const all = dump(r);
  assert.ok(all.includes('PC105') && all.includes('IHDR'), 'the scan reads the answer, pictures included');
  for (const s of LEAKS) assert.ok(!all.includes(s), `${s} is in the answer`);
});

test('readPictures: the numbers are handed over before any picture is unpacked; numbersOnly reads them alone', async () => {
  const bytes = await mainReport(), seen = [];
  const r = await read(bytes, { onNumbers: n => seen.push(['numbers', n]), onProgress: (d, t) => seen.push(['progress', d, t]) });
  assert.equal(r.ok, true, r.error);
  assert.deepEqual(seen.map(x => x[0]), ['numbers', 'progress'], 'the numbers once, first');
  assert.equal(seen[0][1], r.numbers);
  assert.deepEqual(seen[0][1], whole(bytes));
  // The numbers alone: the same numbers, no picture looked for or counted, no progress.
  const got = [];
  const only = await read(bytes, { numbersOnly: true, onNumbers: n => got.push(n), onProgress: () => got.push('progress') });
  assert.equal(only.ok, true, only.error);
  assert.deepEqual(only.numbers, whole(bytes));
  assert.deepEqual([only.pictures.size, only.differs, Object.values(only.counts).filter(Boolean)], [0, [], []]);
  assert.equal(got.length, 1, 'onNumbers only');
  // A file whose numbers can't be read hands nothing over.
  let called = 0;
  const bad = await read(EP.storedZip([['notes.txt', Buffer.from('not a workbook')]]), { onNumbers: () => called++ });
  assert.deepEqual([bad.ok, called], [false, 0]);
});

test('readPictures: SheetJS is given only the XML the cells need (no drawings, no pictures, no links)', async () => {
  const bytes = await mainReport(), seen = [];
  const spy = Object.assign(Object.create(X), { read: (b, o) => { seen.push([b, { ...o }]); return X.read(b, o); } });
  const r = await EP.readPictures({ XLSX: spy, InvAts: A, bytes, fileName: NAME, today: TODAY });
  assert.equal(r.ok, true, r.error);
  assert.equal(seen.length, 1);
  assert.deepEqual(seen[0][1], xopt());
  const names = EP.zipEntries(seen[0][0]).map(e => e.name);
  assert.ok(names.includes('xl/worksheets/sheet1.xml') && names.includes('xl/sharedStrings.xml'), names.join(', '));
  assert.deepEqual(names.filter(n => /drawing|media/.test(n)), []);
  const sent = Buffer.from(seen[0][0]).toString('latin1');
  for (const s of LEAKS) assert.ok(!sent.includes(s), s);
  // The page's own reading options are passed through (type stays 'array').
  seen.length = 0;
  await EP.readPictures({ XLSX: spy, InvAts: A, bytes, fileName: NAME, today: TODAY, readOpts: { cellDates: false, type: 'binary' } });
  assert.deepEqual(seen[0][1], { ...xopt(), cellDates: false });
});

test('readPictures: a report with no pictures is ok with none; its numbers equal the whole-file read', async () => {
  const bytes = await report();
  const r = await read(bytes);
  assert.equal(r.ok, true, r.error);
  assert.equal(r.pictures.size, 0);
  assert.deepEqual(r.differs, []);
  assert.deepEqual(Object.values(r.counts).filter(Boolean), []);
  assert.deepEqual(r.numbers, whole(bytes));
  assert.equal(r.numbers.styles.size, 7);
  // Pictures only on rows with no style: still ok, none kept.
  const off = await read(await report([[0, 'a'], [1, 'a'], [8, 'a']], { a: PNG1 }));
  assert.deepEqual([off.ok, off.pictures.size, off.counts.noStyleRow, off.counts.anchors], [true, 0, 3, 3]);
});

test('readPictures: the numbers equal the whole-file read when the report has problems of its own', async () => {
  // Doesn't add up (Grand Total off): the numbers say so, the pictures are still read.
  const rows = ROWS.map((r, i) => (i === ROWS.length - 1 ? [null, null, 'Grand Total', null, null, null, null, 999, 36] : r));
  const bytes = await report([[2, 'a']], { a: PNG1 }, rows);
  const r = await read(bytes);
  assert.equal(r.numbers.ok, false);
  assert.match(r.numbers.error, /doesn't add up/);
  assert.deepEqual(r.numbers, whole(bytes));
  assert.deepEqual([r.ok, [...r.pictures.keys()]], [true, ['PC101']]);
  // No header: no numbers, and the pictures can't be matched to styles.
  const nohead = await report([[2, 'a']], { a: PNG1 }, [TITLE, ['just', 'text'], ['more', 'text']]);
  const n = await read(nohead);
  assert.equal(n.ok, false);
  assert.match(n.error, /^The pictures in Example_Co_.*\.xlsx can't be matched to styles: ATS report: no "Style", "In WH", "Total Customer Orders" column/);
  assert.deepEqual(n.numbers, whole(nohead));
  assert.equal(n.pictures.size, 0);
});

test('readPictures: a file that is not a zip is read whole, as the page reads it; one that can\'t be read at all is not ok', async () => {
  const csv = Buffer.from('Style,In WH,Total Customer Orders\nPC101,5,1\nPC102,7,0\nGrand Total,12,1\n');
  const r = await read(new Uint8Array(csv), { fileName: 'Inventory Report 11-02-2026.csv' });
  assert.equal(r.ok, true, r.error);
  assert.equal(r.pictures.size, 0);
  assert.deepEqual(r.numbers, whole(new Uint8Array(csv), 'Inventory Report 11-02-2026.csv'));
  assert.equal(r.numbers.styles.size, 2);
  // A zip that isn't a workbook.
  const z = await read(EP.storedZip([['notes.txt', Buffer.from('not a workbook')]]));
  assert.equal(z.ok, false);
  assert.match(z.error, /^Couldn't open Example_Co_.*\.xlsx as an Excel workbook \(/);
  assert.equal(z.numbers, null);
  // Damaged XML: not ok, and said in plain words.
  const bytes = await report([[2, 'a']], { a: PNG1 });
  const d = await read(spoil(bytes, 'xl/worksheets/sheet1.xml'));
  assert.equal(d.ok, false);
  assert.match(d.error, /^Couldn't open Example_Co_.*\.xlsx as an Excel workbook \("xl\/worksheets\/sheet1\.xml" /);
  assert.equal((await read(new Uint8Array(0))).error, `${NAME} is empty (0 bytes).`);
  assert.match((await EP.readPictures({ bytes })).error, /parts didn't load/);
});

test('readPictures: a damaged picture is skipped and counted; the others are still read', async () => {
  const bytes = await report([[2, 'a'], [3, 'b']], { a: PNG1, b: PNG2 });
  const r = await read(spoil(bytes, 'xl/media/image2.png'));
  assert.equal(r.ok, true, r.error);
  assert.deepEqual([[...r.pictures.keys()], r.counts.damaged, r.counts.pictures], [['PC101'], 1, 1]);
});

test('without DecompressionStream, SheetJS\'s zip reader unpacks the file instead: the same answer', async () => {
  const bytes = await mainReport(), fast = await read(bytes);
  const pngName = mediaNamed(await partsOf(bytes), PNG1), png = EP.zipEntries(bytes).find(e => e.name === pngName);
  const d = Object.getOwnPropertyDescriptor(globalThis, 'DecompressionStream');
  Object.defineProperty(globalThis, 'DecompressionStream', { value: undefined, configurable: true, writable: true });
  try {
    const slow = await read(bytes);
    assert.equal(slow.ok, true, slow.error);
    assert.deepEqual(slow, fast);
    await assert.rejects(EP.inflateEntry(bytes, png), /this browser can't unpack the file \(it has no DecompressionStream\)/);
    assert.deepEqual(await EP.inflateEntry(bytes, png, { CFB: X.CFB }), PNG1);
  } finally { Object.defineProperty(globalThis, 'DecompressionStream', d); }
  assert.deepEqual(await EP.inflateEntry(bytes, png), PNG1);
});

// ---------- a damaged or made-up file: nothing past what an ERP report could need is unpacked or drawn

test('a part that says it unpacks to more than packing can give, or to more than 64 MB, is refused before anything is set aside for it', async () => {
  const bytes = await report([[2, 'a']], { a: PNG1 }), png = EP.zipEntries(bytes).find(e => e.name === 'xl/media/image1.png');
  await assert.rejects(EP.inflateEntry(bytes, { ...png, size: png.compressedSize * 1100 + 4097 }),
    /^Error: "xl\/media\/image1\.png" says it unpacks to \d+ bytes from \d+, more than packing can give: the file is damaged$/);
  await assert.rejects(EP.inflateEntry(bytes, { ...png, size: 65 * 1024 * 1024, compressedSize: 1024 * 1024 }), /"xl\/media\/image1\.png" would be 68157440 bytes unpacked: too big for a part of a report/);
  // Within both: unpacked as before.
  assert.deepEqual(await EP.inflateEntry(bytes, png), PNG1);
});

test('readPictures: a report whose parts really unpack to far more than an ERP report\'s is not read (a 70 MB part packed to ~70 KB)', async () => {
  const parts = [...(await partsOf(await report([[2, 'a']], { a: PNG1 })))];
  // An extra XML part nothing points to, 70 MB of spaces: refused before it is unpacked.
  const big = deflatedZip([...parts, ['xl/invented.xml', Buffer.alloc(70 * 1024 * 1024, 0x20)]]);
  assert.ok(big.length < 200 * 1024, `${big.length} bytes`);
  const r = await read(big);
  assert.equal(r.ok, false);
  assert.match(r.error, /^Couldn't open Example_Co_.*\.xlsx as an Excel workbook \("xl\/invented\.xml" would be 73400320 bytes unpacked: too big for a part of a report\)\.$/);
  // XML parts that say they unpack to more than 200 MB all together: refused before any is unpacked.
  let said = await report([[2, 'a']], { a: PNG1 });
  for (const n of ['xl/worksheets/sheet1.xml', 'xl/sharedStrings.xml', 'xl/styles.xml', 'xl/theme/theme1.xml']) said = claim(said, n, 60 * 1024 * 1024);
  const x = await read(said);
  assert.equal(x.ok, false);
  assert.equal(x.error, `${NAME} can't be read: its cells would be 240 MB unpacked, far more than an ATS report's (a few MB), so the file is damaged or isn't an ATS report.`);
  assert.equal(x.numbers, null);
});

test('readPictures: a picture bigger than any ERP picture (over 8 MB, or 25 million pixels) is skipped and counted; far too many MB of them in all: none read', async () => {
  // PC101: a 9 MB part (packed to a few KB); PC102: a PNG whose header says 6,000 × 5,000; PC103: a real 1×1 PNG.
  const parts = await partsOf(await report([[2, 'a'], [3, 'b'], [5, 'c']], { a: PNG2, b: PNG_E, c: PNG1 }));
  const nameOf = b => mediaNamed(parts, b);
  parts.set(nameOf(PNG2), cat([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a], Buffer.alloc(9 * 1024 * 1024)));
  parts.set(nameOf(PNG_E), pngHead(6000, 5000));
  const r = await read(deflatedZip([...parts]));
  assert.equal(r.ok, true, r.error);
  assert.deepEqual([[...r.pictures.keys()], r.counts.tooBig, r.counts.pictures, r.counts.bytesIn], [['PC103'], 2, 1, PNG1.length]);
  assert.deepEqual(r.numbers, whole(await report()), 'the numbers as ever');
  // 40 picture parts of 8 MB each (by what the zip says): 320 MB, more than 300 MB, so none is read; the numbers are.
  const imgs = {}, pics = [];
  for (let i = 0; i < 40; i++) { imgs['p' + i] = cat(PNG1, `invented ${i}`); pics.push([[2, 3, 4, 5, 6, 7, 9, 10, 11, 12][i % 10], 'p' + i]); }
  let many = await report(pics, imgs);
  for (const e of EP.zipEntries(many)) if (e.name.startsWith('xl/media/')) many = claim(many, e.name, 8 * 1024 * 1024);
  const m = await read(many);
  assert.equal(m.ok, false);
  assert.equal(m.error, `The pictures in ${NAME} would be 320 MB unpacked, far more than an ERP report's (about 40 MB), so they aren't read.`);
  assert.equal(m.pictures.size, 0);
  assert.deepEqual(m.numbers, whole(many));
});

test('pictureSize: a picture\'s width and height from its first bytes, or null when they can\'t be told', () => {
  const jpeg = cat([0xff, 0xd8, 0xff, 0xe0, 0x00, 0x10], Buffer.alloc(14), [0xff, 0xc0, 0x00, 0x11, 0x08, 0x01, 0x2c, 0x03, 0x84, 0x03], Buffer.alloc(9));
  const bmp = Buffer.alloc(30); bmp.write('BM', 0); bmp.writeUInt32LE(40, 14); bmp.writeInt32LE(640, 18); bmp.writeInt32LE(-480, 22);
  const webp = Buffer.alloc(30); webp.write('RIFF', 0); webp.write('WEBP', 8); webp.write('VP8X', 12); webp.writeUIntLE(1199, 24, 3); webp.writeUIntLE(799, 27, 3);
  assert.deepEqual([PNG1, PNG2, pngHead(6000, 5000)].map(b => EP.pictureSize(b, 'png')), [{ w: 1, h: 1 }, { w: 2, h: 1 }, { w: 6000, h: 5000 }]);
  assert.deepEqual([EP.pictureSize(GIF_B, 'gif'), EP.pictureSize(jpeg, 'jpeg'), EP.pictureSize(bmp, 'bmp'), EP.pictureSize(webp, 'webp')],
    [{ w: 1, h: 1 }, { w: 900, h: 300 }, { w: 640, h: 480 }, { w: 1200, h: 800 }]);
  // No IHDR where a PNG has it, a JPEG with no frame header: not told (the browser finds out).
  assert.deepEqual([EP.pictureSize(PNG_E, 'png'), EP.pictureSize(JPG_C, 'jpeg'), EP.pictureSize(new Uint8Array(3), 'gif')], [null, null, null]);
});

// ---------- the numbers' part, kept by the page

test('readPictures: the numbers\' part on its own (slim) gives the same numbers, and holds no person\'s name, no address, no picture', async () => {
  // Saved by an invented person, its title linked to the invented picture server.
  const link = LINK.replace(/&amp;/g, '&');
  const edit = (wb, ws) => { wb.creator = 'Invented Person'; wb.lastModifiedBy = 'Invented Person'; ws.getCell('A1').value = { text: TITLE[0], hyperlink: link }; };
  const bytes = await report([[2, 'a'], [3, 'b']], { a: PNG1, b: PNG2 }, ROWS, edit);
  const everything = dump([...(await partsOf(bytes))]);
  assert.ok(everything.includes('Invented Person') && everything.includes('INVENTEDKEY123'), 'the whole file has them, so the test means something');
  let handed = null;
  const r = await read(bytes, { onNumbers: (n, s) => { handed = s; } });
  assert.equal(r.ok, true, r.error);
  assert.ok(r.slim instanceof Uint8Array && r.slim.length > 1000, 'the numbers\' part');
  assert.equal(handed, r.slim, 'handed over with the numbers');
  const names = EP.zipEntries(r.slim).map(e => e.name);
  assert.ok(names.includes('xl/worksheets/sheet1.xml') && names.includes('xl/sharedStrings.xml'), names.join(', '));
  assert.deepEqual(names.filter(n => /^docProps\/|drawing|media/.test(n)), []);
  const text = Buffer.from(r.slim).toString('latin1');
  for (const s of ['Invented Person', ...LEAKS]) assert.ok(!text.includes(s), `${s} is in the numbers' part`);
  // Read again (as the page reads its copy): the very numbers of the whole file.
  const again = await read(r.slim, { numbersOnly: true });
  assert.equal(again.ok, true, again.error);
  assert.deepEqual(again.numbers, whole(bytes));
  assert.deepEqual(r.numbers, whole(bytes));
  // An address left in a cell: the numbers are read, but the part isn't offered to keep.
  const inCell = await report([], {}, ROWS.map((x, i) => (i === 0 ? [`See ${link}`] : x)));
  const c = await read(inCell);
  assert.equal(c.ok, true, c.error);
  assert.deepEqual([c.slim, c.numbers], [null, whole(inCell)]);
});

test('a big report is 2 MB or more; the reading options are the page\'s', () => {
  assert.equal(EP.BIG_REPORT, 2 * 1024 * 1024);
  assert.deepEqual({ type: 'array', ...EP.READ_OPTS }, xopt());
  assert.equal(EP.crc32(Buffer.from('123456789')), 0xcbf43926);
});
