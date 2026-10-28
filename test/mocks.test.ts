import { describe, expect, it } from 'vitest';
import {
  CELL_LIMIT_MESSAGE,
  COLUMNS_OUT_OF_BOUNDS_MESSAGE,
  DELETE_ALL_ROWS_MESSAGE,
  ROWS_OUT_OF_BOUNDS_MESSAGE,
  FakeSheet,
  FakeSpreadsheet,
  coerceUserEntered,
  createMocks,
} from './mocks';

describe('coerceUserEntered', () => {
  it.each([
    ["'0012", { value: '0012', formula: '' }],
    ['=A1+1', { value: '', formula: '=A1+1' }],
    ['0012', { value: 12, formula: '' }],
    ['-3.5', { value: -3.5, formula: '' }],
    ['1e3', { value: 1000, formula: '' }],
    ['true', { value: true, formula: '' }],
    ['FALSE', { value: false, formula: '' }],
    ['plain text', { value: 'plain text', formula: '' }],
    ['', { value: '', formula: '' }],
    ['2026-02-30', { value: '2026-02-30', formula: '' }],
  ])('interprets %j', (input, expected) => {
    expect(coerceUserEntered(input)).toEqual(expected);
  });

  it('turns an ISO date into a Date at UTC midnight', () => {
    expect(coerceUserEntered('2026-10-27').value).toEqual(
      new Date('2026-10-27T00:00:00Z'),
    );
  });

  it('stores non-string values as they are', () => {
    const date = new Date();
    expect(coerceUserEntered(7).value).toBe(7);
    expect(coerceUserEntered(false).value).toBe(false);
    expect(coerceUserEntered(date).value).toBe(date);
  });
});

describe('FakeSheet', () => {
  it('starts as an empty 1000x26 grid', () => {
    const sheet = new FakeSheet('s');

    expect(sheet.getMaxRows()).toBe(1000);
    expect(sheet.getMaxColumns()).toBe(26);
    expect(sheet.getLastRow()).toBe(0);
    expect(sheet.getLastColumn()).toBe(0);
    expect(sheet.getDataRange().getValues()).toEqual([['']]);
  });

  it('reports the last row and column holding any content', () => {
    const sheet = new FakeSheet('s');
    sheet.getRange(3, 2).setValues([[0]]);
    sheet.getRange(5, 1).setValues([['=1']]);

    expect(sheet.getLastRow()).toBe(5);
    expect(sheet.getLastColumn()).toBe(2);

    sheet.getRange(5, 1).clearContent();
    expect(sheet.getLastRow()).toBe(3);
  });

  it('throws when a range exceeds the grid', () => {
    const sheet = new FakeSheet('s');

    expect(() => sheet.getRange(1000, 1, 2, 1)).toThrow(
      ROWS_OUT_OF_BOUNDS_MESSAGE,
    );
    expect(() => sheet.getRange(1, 26, 1, 2)).toThrow(
      COLUMNS_OUT_OF_BOUNDS_MESSAGE,
    );
    expect(() => sheet.getRange(0, 1)).toThrow('invalid');
  });

  it('rejects data whose dimensions differ from the range', () => {
    const range = new FakeSheet('s').getRange(1, 1, 2, 2);

    expect(() => range.setValues([['a', 'b']])).toThrow(
      'The number of rows in the data does not match the number of rows in the range.',
    );
    expect(() => range.setValues([['a'], ['b']])).toThrow(
      'The number of columns in the data does not match the number of columns in the range.',
    );
  });

  it('rejects a string input over 50,000 characters, apostrophe included', () => {
    const sheet = new FakeSheet('s');
    const range = sheet.getRange(1, 1, 1, 2);

    expect(() => range.setValues([['a', 'x'.repeat(50_001)]])).toThrow(
      CELL_LIMIT_MESSAGE,
    );
    expect(() => range.setValues([['a', `'${'x'.repeat(50_000)}`]])).toThrow(
      'Your input contains more than the maximum of 50000 characters in a single cell.',
    );
    expect(sheet.getLastRow()).toBe(0);
    expect(sheet.calls.setValues).toHaveLength(0);

    range.setValues([['a', `'${'x'.repeat(49_999)}`]]);
    expect(sheet.getRange(1, 2).getValues()[0]![0]).toHaveLength(49_999);
  });

  describe('spill', () => {
    const FORMULA = '=ARRAYFORMULA(IF(LEN(B2:B), ROW(B2:B), ""))';

    /** A spill in C2 that outputs `values` down the column. */
    const spilled = (values: unknown[]) => {
      const sheet = new FakeSheet('s');
      sheet.spill(2, 3, values, FORMULA);
      return sheet;
    };

    it('counts blank output as content for the last row and data range', () => {
      const sheet = spilled(Array(999).fill(''));

      expect(sheet.getLastRow()).toBe(1000);
      expect(sheet.getLastColumn()).toBe(3);
      expect(sheet.getDataRange().getValues()).toHaveLength(1000);
    });

    it('holds the formula in the anchor only', () => {
      const sheet = spilled([2, 3]);

      expect(sheet.getRange(1, 3, 4, 1).getFormulas()).toEqual([
        [''],
        [FORMULA],
        [''],
        [''],
      ]);
      expect(sheet.getRange(2, 3, 2, 1).getValues()).toEqual([[2], [3]]);
    });

    it('breaks when a value is written into a spilled cell', () => {
      const sheet = spilled([2, 3, 4]);

      sheet.getRange(3, 3).setValues([['typed']]);

      expect(sheet.brokenSpills).toEqual([[2, 3]]);
      expect(sheet.getRange(2, 3, 4, 1).getValues()).toEqual([
        ['#REF!'],
        ['typed'],
        [''],
        [''],
      ]);
      expect(sheet.getLastRow()).toBe(3);
    });

    it('does not break when a neighbouring cell is written', () => {
      const sheet = spilled([2, 3]);

      sheet.getRange(2, 1, 2, 2).setValues([
        ['a', 'b'],
        ['c', 'd'],
      ]);

      expect(sheet.brokenSpills).toEqual([]);
      expect(sheet.getRange(2, 3, 2, 1).getValues()).toEqual([[2], [3]]);
    });
  });

  it('returns formulas and the evaluator result for formula cells', () => {
    const sheet = new FakeSheet('s', { evaluate: (formula) => formula.length });
    sheet.getRange(1, 1, 1, 2).setValues([['=SUM(B1)', 'text']]);

    const range = sheet.getRange(1, 1, 1, 2);
    expect(range.getFormulas()).toEqual([['=SUM(B1)', '']]);
    expect(range.getValues()).toEqual([[8, 'text']]);
  });

  it('shifts content when inserting rows and columns', () => {
    const sheet = new FakeSheet('s');
    sheet.getRange(1, 1, 2, 2).setValues([
      ['a', 'b'],
      ['c', 'd'],
    ]);

    sheet.insertRowsAfter(1, 2);
    sheet.insertColumnsAfter(1, 1);

    expect(sheet.getMaxRows()).toBe(1002);
    expect(sheet.getMaxColumns()).toBe(27);
    expect(sheet.getRange(4, 3).getValues()).toEqual([['d']]);
    expect(sheet.getRange(2, 2).getValues()).toEqual([['']]);
  });

  it('deletes rows but never every non-frozen row', () => {
    const sheet = new FakeSheet('s');
    sheet.setFrozenRows(1);
    sheet.getRange(3, 1).setValues([['x']]);

    sheet.deleteRows(3, 998);
    expect(sheet.getMaxRows()).toBe(2);

    expect(() => sheet.deleteRows(2, 1)).toThrow(DELETE_ALL_ROWS_MESSAGE);
    expect(() => sheet.deleteRows(2, 5)).toThrow(ROWS_OUT_OF_BOUNDS_MESSAGE);
    expect(sheet.getFrozenRows()).toBe(1);
  });
});

describe('FakeSpreadsheet', () => {
  it('rejects a duplicate sheet name', () => {
    const spreadsheet = new FakeSpreadsheet();
    spreadsheet.insertSheet('X');

    expect(() => spreadsheet.insertSheet('X')).toThrow(
      'A sheet with the name "X" already exists. Please enter another name.',
    );
    expect(spreadsheet.getSheetByName('X')).not.toBeNull();
    expect(spreadsheet.getSheetByName('Y')).toBeNull();
    expect(spreadsheet.getSheets()).toHaveLength(1);
  });
});

describe('createMocks', () => {
  it('shares one held flag per scope across handles', () => {
    const mocks = createMocks();
    const outer = mocks.LockService.getScriptLock()!;
    const nested = mocks.LockService.getScriptLock()!;

    outer.waitLock(1);
    expect(mocks.isLocked()).toBe(true);
    expect(nested.tryLock(1)).toBe(true);

    nested.releaseLock();
    expect(mocks.isLocked()).toBe(false);
    expect(outer.hasLock()).toBe(false);
  });

  it('returns no lock for an unavailable scope', () => {
    const mocks = createMocks({ unavailable: ['document'] });

    expect(mocks.LockService.getDocumentLock()).toBeNull();
    expect(mocks.LockService.getUserLock()).not.toBeNull();
  });

  it('records whether a lock was held at each flush', () => {
    const mocks = createMocks();

    mocks.SpreadsheetApp.flush();
    mocks.LockService.getUserLock()!.waitLock(1);
    mocks.SpreadsheetApp.flush();

    expect(mocks.flushLockStates).toEqual([false, true]);
  });

  it('generates deterministic UUIDs', () => {
    const { Utilities } = createMocks();

    expect(Utilities.getUuid()).toBe('00000000-0000-4000-8000-000000000001');
    expect(Utilities.getUuid()).toBe('00000000-0000-4000-8000-000000000002');
  });

  it('opens one stable spreadsheet per id and url', () => {
    const { SpreadsheetApp } = createMocks();

    expect(SpreadsheetApp.openById('a')).toBe(SpreadsheetApp.openById('a'));
    expect(SpreadsheetApp.openById('a')).not.toBe(SpreadsheetApp.openById('b'));
    expect(SpreadsheetApp.openByUrl('a')).not.toBe(
      SpreadsheetApp.openById('a'),
    );
  });
});
