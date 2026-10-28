import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import { at, enter, rows, setup, setupTable } from './helpers';
import type { FakeSheet } from './mocks';

const CALC_COLUMN = 7;

// distinct timestamps, so that an update always changes `_updatedAt`
beforeEach(() => {
  vi.useFakeTimers();
  vi.setSystemTime(new Date('2026-01-01T00:00:00Z'));
});
afterEach(() => vi.useRealTimers());

/** Formula of the `calc` column on a sheet row. */
const formulaAt = (sheet: FakeSheet, row: number): string =>
  sheet.getRange(row, CALC_COLUMN).getFormulas()[0]![0]!;

/**
 * Ann, Bob, Cy and Dee inserted through the library, followed by a `calc`
 * formula column typed in by hand (`=R<row>`), as a person would add it.
 */
function withFormulas() {
  const context = setupTable('people', {}, { evaluate: (f) => `value${f}` });
  const entries = context.table.insertMany([
    { name: 'Ann', age: 30 },
    { name: 'Bob', age: 40 },
    { name: 'Cy', age: 50 },
    { name: 'Dee', age: 60 },
  ]);
  const { sheet } = context;
  sheet.getRange(1, CALC_COLUMN).setValues([['calc']]);
  for (let row = 2; row <= 5; row++) {
    sheet.getRange(row, CALC_COLUMN).setValues([[`=R${row}`]]);
  }
  sheet.calls.setValues.length = 0;
  vi.advanceTimersByTime(1000);
  return { ...context, entries };
}

const calc = (sheet: FakeSheet) => rows(sheet).map((r) => r.calc);

const FORMULAS = ['=R2', '=R3', '=R4', '=R5'];
const COMPUTED = FORMULAS.map((f) => `value${f}`);

describe('formula preservation (H4)', () => {
  it('keeps the formula of an updated row when the patch does not set it', () => {
    const { table, sheet, entries } = withFormulas();

    table.update({ _id: at(entries, 1)._id!, age: 41 });

    expect(formulaAt(sheet, 3)).toBe('=R3');
    expect(calc(sheet)).toEqual(COMPUTED);
  });

  it('keeps the formulas of rows it did not update', () => {
    const { table, sheet, entries } = withFormulas();

    table.update({ _id: at(entries, 1)._id!, age: 41 });

    expect([2, 4, 5].map((row) => formulaAt(sheet, row))).toEqual([
      '=R2',
      '=R4',
      '=R5',
    ]);
  });

  it('writes the value of a formula cell the entry sets', () => {
    const { table, sheet, entries } = withFormulas();

    table.update({ _id: at(entries, 1)._id!, calc: 'static' });

    expect(formulaAt(sheet, 3)).toBe('');
    expect(calc(sheet)[1]).toBe('static');
  });

  it('returns the computed value of a preserved formula cell', () => {
    const { table, entries } = withFormulas();

    const updated = table.update({ _id: at(entries, 1)._id!, age: 41 });

    expect(updated.calc).toBe('value=R3');
  });

  it('survives soft delete and restore', () => {
    const { table, sheet, entries } = withFormulas();

    table.softDeleteMany([at(entries, 0), at(entries, 2)]);
    expect(FORMULAS.map((_, i) => formulaAt(sheet, i + 2))).toEqual(FORMULAS);

    table.restoreMany([at(entries, 0), at(entries, 2)]);
    expect(FORMULAS.map((_, i) => formulaAt(sheet, i + 2))).toEqual(FORMULAS);
  });

  it('survives updateWhere', () => {
    const { table, sheet } = withFormulas();

    table.updateWhere((e) => (e.age as number) > 35, { age: 1 });

    expect(FORMULAS.map((_, i) => formulaAt(sheet, i + 2))).toEqual(FORMULAS);
    expect(rows(sheet).map((r) => r.age)).toEqual([30, 1, 1, 1]);
  });

  it('survives deleting other rows', () => {
    const { table, sheet, entries } = withFormulas();

    table.deleteMany([at(entries, 0), at(entries, 2)]);

    expect([formulaAt(sheet, 2), formulaAt(sheet, 3)]).toEqual(['=R3', '=R5']);
    expect(rows(sheet).map((r) => r.name)).toEqual(['Bob', 'Dee']);
  });

  it('survives the backfill of a row typed into the sheet', () => {
    const { db, mocks } = setup({}, { evaluate: (f) => `value${f}` });
    const sheet = mocks.activeSpreadsheet.insertSheet('people');
    enter(sheet, [
      ['name', 'calc'],
      ['Ann', '=B9'],
      ['Bob', '=B8'],
    ]);
    const table = db.table({ sheetName: 'people' });

    const found = table.find();

    expect(found.map((e) => e.calc)).toEqual(['value=B9', 'value=B8']);
    const calcColumn = sheet
      .getRange(1, 1, 3, 2)
      .getFormulas()
      .map((r) => r[1]);
    expect(calcColumn).toEqual(['', '=B9', '=B8']);
    expect(found.every((e) => e._id)).toBe(true);
  });
});

describe('cell-scoped writes', () => {
  const AGE = 2;
  const UPDATED_AT = 5;

  /** `[row, col, numRows, numCols]` of every `setValues` call. */
  const written = (sheet: FakeSheet) =>
    sheet.calls.setValues.map((c) => [c.row, c.col, c.numRows, c.numCols]);

  it('writes only the changed cell and _updatedAt of the touched row', () => {
    const { table, sheet, entries } = withFormulas();

    table.update({ _id: at(entries, 2)._id!, age: 51 });

    expect(written(sheet)).toEqual([
      [4, AGE, 1, 1],
      [4, UPDATED_AT, 1, 1],
    ]);
  });

  it('preserves formulas when a full entry from find() is updated', () => {
    const { table, sheet } = withFormulas();
    const entry = at(table.find(), 1);

    entry.age = 41;
    table.update(entry);

    expect(written(sheet)).toEqual([
      [3, AGE, 1, 1],
      [3, UPDATED_AT, 1, 1],
    ]);
    expect(FORMULAS.map((_, i) => formulaAt(sheet, i + 2))).toEqual(FORMULAS);
  });

  it('writes only _updatedAt for an update that changes nothing else', () => {
    const { table, sheet } = withFormulas();

    table.update(at(table.find(), 0));

    expect(written(sheet)).toEqual([[2, UPDATED_AT, 1, 1]]);
  });

  it('never writes a spilled array-formula range it does not change', () => {
    const { table, sheet } = withFormulas();
    // an ARRAYFORMULA anchored at G2 spills values into G3:G5, which hold no
    // formula of their own
    sheet.getRange(2, CALC_COLUMN).setValues([['=ARRAYFORMULA(A2:A5)']]);
    sheet.getRange(3, CALC_COLUMN, 3, 1).clearContent();
    for (const [i, name] of ['Bob', 'Cy', 'Dee'].entries()) {
      sheet.getRange(3 + i, CALC_COLUMN).setValues([[name]]);
    }
    sheet.calls.setValues.length = 0;

    table.updateWhere(() => true, { age: 0 });
    table.softDeleteWhere(() => true);

    const touchesSpill = sheet.calls.setValues.some(
      (c) => c.col <= CALC_COLUMN && c.col + c.numCols > CALC_COLUMN,
    );
    expect(touchesSpill).toBe(false);
    expect(formulaAt(sheet, 2)).toBe('=ARRAYFORMULA(A2:A5)');
  });

  it('overwrites a formula cell that the entry sets to a new value', () => {
    const { table, sheet, entries } = withFormulas();

    table.update({ _id: at(entries, 1)._id!, calc: 'static' });

    expect(formulaAt(sheet, 3)).toBe('');
    expect(calc(sheet)[1]).toBe('static');
  });

  it('does not write a cell set to the value it already has', () => {
    const { table, sheet, entries } = withFormulas();

    table.update({ _id: at(entries, 1)._id!, age: 40, name: 'Bob' });

    expect(written(sheet)).toEqual([[3, UPDATED_AT, 1, 1]]);
  });

  it('merges adjacent rows that change the same columns', () => {
    const { table, sheet, entries } = withFormulas();

    table.softDeleteMany([at(entries, 2), at(entries, 0), at(entries, 1)]);

    expect(written(sheet)).toEqual([[2, UPDATED_AT, 3, 2]]);
  });

  it('writes one call per column run when columns are not adjacent', () => {
    const { table, sheet, entries } = withFormulas();

    table.updateMany(entries.slice(0, 3).map((e) => ({ _id: e._id!, age: 1 })));

    expect(written(sheet)).toEqual([
      [2, AGE, 3, 1],
      [2, UPDATED_AT, 3, 1],
    ]);
  });

  it('writes rows whose changed columns differ separately', () => {
    const { table, sheet, entries } = withFormulas();

    table.updateMany([
      { _id: at(entries, 0)._id!, age: 1 },
      { _id: at(entries, 1)._id!, name: 'Bo' },
      { _id: at(entries, 3)._id!, age: 1 },
    ]);

    expect(written(sheet)).toEqual([
      [2, AGE, 1, 1],
      [2, UPDATED_AT, 1, 1],
      [3, 1, 1, 1],
      [3, UPDATED_AT, 1, 1],
      [5, AGE, 1, 1],
      [5, UPDATED_AT, 1, 1],
    ]);
  });

  it('writes only the repaired system cells when backfilling', () => {
    const { table, sheet } = withFormulas();
    sheet.getRange(6, 1).setValues([['Eve']]);
    sheet.calls.setValues.length = 0;

    table.find();

    expect(written(sheet)).toEqual([[6, 3, 1, 4]]);
  });

  it('does not overwrite a manual edit made after the rows were read', () => {
    const { table, sheet, entries } = withFormulas();
    const readBody = sheet.getDataRange;
    sheet.getDataRange = () => {
      const range = readBody();
      const values = range.getValues();
      sheet.getRange(2, 2).setValues([[999]]);
      sheet.getRange(5, 1).setValues([['Manual']]);
      range.getValues = () => values;
      return range;
    };

    table.update({ _id: at(entries, 2)._id!, age: 51 });

    expect(rows(sheet).map((r) => [r.name, r.age])).toEqual([
      ['Ann', 999],
      ['Bob', 40],
      ['Cy', 51],
      ['Manual', 60],
    ]);
  });
});

describe('deleting rows in place', () => {
  /** `n` entries; `trimTo` shrinks the sheet to that many rows first. */
  const fill = (n: number, frozen?: number, trimTo?: number) => {
    const context = setupTable();
    const entries = context.table.insertMany(
      Array.from({ length: n }, (_, i) => ({ name: `r${i}` })),
    );
    if (trimTo !== undefined) {
      context.sheet.deleteRows(trimTo + 1, context.sheet.getMaxRows() - trimTo);
    }
    if (frozen !== undefined) context.sheet.setFrozenRows(frozen);
    context.sheet.deleteRows.mockClear();
    return { ...context, entries };
  };

  const deleted = (sheet: FakeSheet) =>
    sheet.deleteRows.mock.calls.map((c) => [...c]);

  it('deletes a single row', () => {
    const { table, sheet, entries } = fill(3);

    table.deleteOne(at(entries, 1));

    expect(deleted(sheet)).toEqual([[3, 1]]);
    expect(table.find().map((e) => e.name)).toEqual(['r0', 'r2']);
  });

  it('deletes non-contiguous rows bottom-up', () => {
    const { table, sheet, entries } = fill(5);

    table.deleteMany([at(entries, 0), at(entries, 4), at(entries, 2)]);

    expect(deleted(sheet)).toEqual([
      [6, 1],
      [4, 1],
      [2, 1],
    ]);
    expect(table.find().map((e) => e.name)).toEqual(['r1', 'r3']);
  });

  it('deletes a contiguous run with one call', () => {
    const { table, sheet, entries } = fill(5);

    table.deleteMany([at(entries, 3), at(entries, 1), at(entries, 2)]);

    expect(deleted(sheet)).toEqual([[3, 3]]);
    expect(table.find().map((e) => e.name)).toEqual(['r0', 'r4']);
  });

  it('writes nothing when deleting', () => {
    const { table, sheet, entries } = fill(3);
    sheet.calls.setValues.length = 0;

    table.deleteOne(at(entries, 0));

    expect(sheet.calls.setValues).toEqual([]);
  });

  it('deletes a duplicated entry once', () => {
    const { table, sheet, entries } = fill(3);

    table.deleteMany([at(entries, 0), at(entries, 0)]);

    expect(deleted(sheet)).toEqual([[2, 1]]);
  });

  it('does nothing for an empty batch', () => {
    const { table, sheet } = fill(3);

    table.deleteMany([]);

    expect(deleted(sheet)).toEqual([]);
    expect(table.find()).toHaveLength(3);
  });

  it('returns the pre-deletion entries from deleteWhere', () => {
    const { table } = fill(3);

    const removed = table.deleteWhere((e) => e.name !== 'r1');

    expect(removed.map((e) => e.name)).toEqual(['r0', 'r2']);
    expect(table.find().map((e) => e.name)).toEqual(['r1']);
  });

  it('keeps a row when the body fills the sheet', () => {
    const { table, sheet, entries } = fill(3, undefined, 4);

    expect(() => table.deleteMany(entries)).not.toThrow();

    expect(sheet.getMaxRows()).toBe(2);
    expect(table.find()).toEqual([]);
    expect(sheet.getLastRow()).toBe(1);
    table.insert({ name: 'again' });
    expect(table.find().map((e) => e.name)).toEqual(['again']);
  });

  it.each([
    ['no frozen rows', 0],
    ['two frozen rows', 2],
  ])('deletes every entry with %s', (_name, frozen) => {
    const { table, sheet, entries } = fill(3, frozen, frozen ? 6 : 4);

    expect(() => table.deleteMany(entries)).not.toThrow();

    expect(sheet.getMaxRows()).toBeGreaterThan(sheet.getFrozenRows());
    expect(table.find()).toEqual([]);
    table.insert({ name: 'again' });
    expect(table.find().map((e) => e.name)).toEqual(['again']);
  });

  it('keeps a row when deleting part of a body held by frozen rows', () => {
    const { table, sheet, entries } = fill(2, 2, 3);

    expect(() => table.deleteOne(at(entries, 1))).not.toThrow();

    expect(sheet.getMaxRows()).toBe(3);
    expect(table.find().map((e) => e.name)).toEqual(['r0']);
  });
});
