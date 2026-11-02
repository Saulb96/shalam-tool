// Loads the page's browser files into Node for tests: the vendored spreadsheet libraries and the core parsers.
import { readFileSync } from 'node:fs';
import { createRequire } from 'node:module';
import { fileURLToPath } from 'node:url';
import { dirname, join } from 'node:path';

const ROOT = join(dirname(fileURLToPath(import.meta.url)), '..');
const require = createRequire(import.meta.url);

let xlsx = null;
// The bundle is a browser script. Run as a plain function it builds its own XLSX object, as it does in the page.
export function loadXLSX() {
  if (!xlsx) xlsx = new Function(readFileSync(join(ROOT, 'vendor/xlsx-js-style/xlsx.bundle.js'), 'utf8') + '\n;return XLSX;')();
  return xlsx;
}

// ExcelJS, the sales sheet's writer (loaded only when a sheet is built in the page): its bundle is UMD, so in Node it
// answers module.exports.
export const loadExcelJS = () => require(join(ROOT, 'vendor/exceljs/exceljs.min.js'));

export const core = name => require(join(ROOT, 'core', name + '.js'));
export const root = (...p) => join(ROOT, ...p);

// Every part of a .xlsx (a zip) by its path, as bytes: SheetJS's zip reader (XLSX.CFB), as the page uses it. The
// reader's own marker entry (its name starts with \x01) isn't in the file.
export function xlsxParts(bytes) {
  const X = loadXLSX(), z = X.CFB.read(bytes, { type: 'array' }), out = new Map();
  z.FileIndex.forEach((f, i) => {
    const p = z.FullPaths[i].replace(/^Root Entry\//, '');
    if (f.type === 2 && !p.startsWith('\x01') && f.content && f.content.length) out.set(p, new Uint8Array(f.content));
  });
  return out;
}

// A workbook built from { sheetName: rows (array of arrays) }, written to .xlsx bytes and read back the way the
// page reads a downloaded file, so tests go through the real file format.
export function xlsxBytes(sheets) {
  const X = loadXLSX();
  const wb = X.utils.book_new();
  for (const [name, rows] of Object.entries(sheets)) X.utils.book_append_sheet(wb, X.utils.aoa_to_sheet(rows), name);
  return new Uint8Array(X.write(wb, { type: 'array', bookType: 'xlsx' }));
}
export const readBook = bytes => loadXLSX().read(bytes, { type: 'array' });
