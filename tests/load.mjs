// Loads the page's browser files into Node for tests: the vendored spreadsheet library and the core parsers.
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

export const core = name => require(join(ROOT, 'core', name + '.js'));
export const root = (...p) => join(ROOT, ...p);

// A workbook built from { sheetName: rows (array of arrays) }, written to .xlsx bytes and read back the way the
// page reads a downloaded file, so tests go through the real file format.
export function xlsxBytes(sheets) {
  const X = loadXLSX();
  const wb = X.utils.book_new();
  for (const [name, rows] of Object.entries(sheets)) X.utils.book_append_sheet(wb, X.utils.aoa_to_sheet(rows), name);
  return new Uint8Array(X.write(wb, { type: 'array', bookType: 'xlsx' }));
}
export const readBook = bytes => loadXLSX().read(bytes, { type: 'array' });
