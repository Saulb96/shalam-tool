import test from 'node:test';
import assert from 'node:assert/strict';
import { core, loadXLSX, xlsxBytes, readBook } from './load.mjs';

const C = core('common');

test('numbers: blank is null, text is NaN, never 0', () => {
  assert.equal(C.num(''), null);
  assert.equal(C.num(null), null);
  assert.equal(C.num(' 1,200 '), 1200);
  assert.ok(Number.isNaN(C.num('n/a')));
  assert.equal(C.num(48), 48);
});

test('styles keep inner spaces and -F; POs drop spaces and keep suffixes', () => {
  assert.notEqual(C.normStyle('RM 2201'), C.normStyle('RM2201'));
  assert.notEqual(C.normStyle('BB-310-F'), C.normStyle('BB-310'));
  assert.equal(C.normStyle(' x90001 '), 'X90001');
  assert.equal(C.normPo(' 900437-bcf'), '900437-BCF');
  assert.equal(C.normPo('900437 -1'), '900437-1');
});

test('UPCs are 12-13 digit text, never reshaped', () => {
  assert.equal(C.normUpc(990001234567), '990001234567');
  assert.equal(C.normUpc('0990001234567'), '0990001234567');
  assert.equal(C.normUpc('99000123456'), null);
  assert.equal(C.normUpc('9900-0123-4567'), null);
});

test('headers are found by name, ignoring case, spaces and punctuation', () => {
  const r = C.findColumns(['ID#', 'Container #', 'Round #', 'ETD'], {
    id: { names: ['ID'], required: true }, cont: { names: ['Container'], required: true },
    eta: { names: ['ETA'], required: true }, etd: { names: ['ETD'], required: false } });
  assert.deepEqual(r.cols, { id: 0, cont: 1, etd: 3 });
  assert.deepEqual(r.missing, ['ETA']);
});

test('dates', () => {
  const d = C.dayOf('10/28/2026');
  assert.equal(C.fmtDay(d), '10/28/2026');
  assert.equal(C.dayOf(46323), C.dayOf('10/28/2026'));
  assert.equal(C.dayOf('2026-10-28T00:00:00'), d);
  assert.equal(C.dayOf('10.28.26'), d);
  assert.deepEqual(C.toDay('10/28', 2026), { day: d, yearless: true });
  assert.equal(C.toDay('10/28'), null);
  assert.equal(C.dayOf('2/30/2026'), null);
  assert.equal(C.todayNY(Date.UTC(2026, 9, 29, 3, 0)), d); // 11 pm on 10/28 in New York
  assert.equal(C.businessDaysBetween(C.dayOf('10/23/2026'), C.dayOf('10/26/2026')), 1); // Fri -> Mon
});

test('sheet rows always start at column A', () => {
  const X = loadXLSX();
  const wb = readBook(xlsxBytes({ S: [[null, 'a', 'b'], [null, 1, 2]] }));
  const ws = wb.Sheets.S;
  ws['!ref'] = 'B1:C2'; // a sheet whose recorded range starts at B
  assert.deepEqual(C.sheetRows(X, ws)[0], [null, 'a', 'b']);
  assert.equal(C.findSheet(wb, ' s ').name, 'S');
});
