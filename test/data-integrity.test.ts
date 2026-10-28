import { describe, expect, it } from 'vitest';
import {
  at,
  createLockService,
  enter,
  header,
  rows,
  setup,
  setupTable,
  sheetOf,
} from './helpers';
import type { FakeSheet } from './mocks';

const NOW = new Date('2024-01-01T00:00:00Z');

/** A system-column header followed by `extra` columns. */
const HEADER = ['_id', '_createdAt', '_updatedAt', '_isDeleted'];

/** A fully populated body row for `id`, with `rest` after the system cells. */
const row = (id: string, ...rest: unknown[]): unknown[] => [
  id,
  NOW,
  NOW,
  false,
  ...rest,
];

/** A `people` sheet typed by hand, and a table over it. */
function typed(grid: unknown[][]) {
  const context = setup({ logger: { info: () => {}, warn: () => {} } });
  const sheet = context.mocks.activeSpreadsheet.insertSheet('people');
  enter(sheet, grid);
  return {
    ...context,
    sheet,
    table: context.db.table({ sheetName: 'people' }),
  };
}

describe('rowNumbers validation', () => {
  it.each([
    [{ columnKeys: 0 }, 'columnKeys'],
    [{ columnKeys: 1.5 }, 'columnKeys'],
    [{ firstData: -1 }, 'firstData'],
    [{ firstData: NaN }, 'firstData'],
    [{ firstData: undefined }, 'firstData'],
  ])('the constructor rejects %j', (rowNumbers, key) => {
    const { GasSheetDb } = setup();

    expect(
      () =>
        new GasSheetDb({
          useActiveSpreadsheet: true,
          rowNumbers: rowNumbers as Partial<GasSheetDbRowsReference>,
        }),
    ).toThrow(`[GasSheetDb] Invalid "rowNumbers.${key}"`);
  });

  it.each([
    [{ columnKeys: 2, firstData: 2 }],
    [{ columnKeys: 3, firstData: 2 }],
  ])('the constructor rejects firstData not after columnKeys: %j', (rows) => {
    const { GasSheetDb } = setup();

    expect(
      () => new GasSheetDb({ useActiveSpreadsheet: true, rowNumbers: rows }),
    ).toThrow('[GasSheetDb] Invalid "rowNumbers": "firstData" (');
  });

  it('table() validates the merged result', () => {
    const { db } = setup();

    expect(() =>
      db.table({ sheetName: 't', rowNumbers: { columnKeys: 3 } }),
    ).toThrow('[GasSheetDb] Invalid "rowNumbers": "firstData" (2)');
    expect(() =>
      db.table({ sheetName: 't', rowNumbers: { columnKeys: 0 } }),
    ).toThrow('[GasSheetDb] Invalid "rowNumbers.columnKeys"');
  });

  it('table() accepts an override that is valid once merged', () => {
    const { db } = setup({ rowNumbers: { firstData: 5 } });

    expect(
      db.table({ sheetName: 't', rowNumbers: { columnKeys: 3 } }).rowNumbers,
    ).toEqual({ columnKeys: 3, firstData: 5 });
  });

  it('appends below a spacer row between the header and firstData', () => {
    const { db, mocks } = setup({ rowNumbers: { firstData: 3 } });
    const table = db.table({ sheetName: 'people' });

    table.insert({ name: 'Ann' });
    table.insert({ name: 'Bob' });

    const sheet = sheetOf(mocks.activeSpreadsheet, 'people');
    const nameCol = header(sheet).indexOf('name') + 1;
    expect(sheet.getRange(2, 1, 1, 5).getValues()[0]).toEqual(
      Array(5).fill(''),
    );
    expect(sheet.getRange(3, nameCol).getValues()[0]![0]).toBe('Ann');
    expect(sheet.getRange(4, nameCol).getValues()[0]![0]).toBe('Bob');
    expect(table.find().map((e) => e.name)).toEqual(['Ann', 'Bob']);
  });
});

describe('creating a sheet with a header below row 1', () => {
  it('bands only the rows from the header to the end of the grid', () => {
    const { db, mocks } = setup({
      rowNumbers: { columnKeys: 3, firstData: 4 },
    });

    db.table({ sheetName: 'offset' });

    const sheet = sheetOf(mocks.activeSpreadsheet, 'offset');
    expect(sheet.calls.applyRowBanding[0]).toMatchObject({
      row: 3,
      numRows: sheet.getMaxRows() - 2,
    });
    expect(sheet.getFrozenRows()).toBe(3);
  });
});

describe('blank rows inside the data range', () => {
  const grid = (): unknown[][] => [
    [...HEADER, 'name'],
    row('a', 'Ann'),
    Array(5).fill(''),
    row('b', 'Bob'),
  ];

  it('are not returned as entries', () => {
    const { table } = typed(grid());

    expect(table.find().map((e) => e.name)).toEqual(['Ann', 'Bob']);
  });

  it('are not assigned metadata, and do not force the locked path', () => {
    const { table, sheet, mocks } = typed(grid());
    const before = rows(sheet);

    table.find();

    expect(rows(sheet)).toEqual(before);
    expect(mocks.isLocked()).toBe(false);
    expect(sheet.calls.setValues).toHaveLength(1);
  });

  it('are skipped when the repair runs because another row needs it', () => {
    const { table, sheet } = typed([
      [...HEADER, 'name'],
      ['', '', '', '', 'Ann'],
      Array(5).fill(''),
      ['', '', '', '', 'Bob'],
    ]);

    expect(table.find().map((e) => e.name)).toEqual(['Ann', 'Bob']);
    expect(at(rows(sheet), 1)).toEqual({
      _id: '',
      _createdAt: '',
      _updatedAt: '',
      _isDeleted: '',
      name: '',
    });
  });

  it('are not matched by the predicate writers', () => {
    const { table } = typed(grid());

    expect(table.deleteWhere(() => true)).toHaveLength(2);
  });
});

describe('header keys', () => {
  it('matches a numeric header to the string key of an entry', () => {
    const { table, sheet } = typed([[...HEADER, 2024], row('a', 'x')]);

    expect(table.find()[0]).toMatchObject({ '2024': 'x' });

    table.update({ _id: 'a', '2024': 'y' });

    expect(header(sheet)).toEqual([...HEADER, 2024]);
    expect(rows(sheet)[0]).toMatchObject({ '2024': 'y' });
  });

  it('does not append a column for a key that matches a numeric header', () => {
    const { table, sheet } = typed([[...HEADER, 2024], row('a', 'x')]);

    table.insert({ '2024': 'z' });
    table.insert({ '2024': 'w' });

    expect(sheet.getLastColumn()).toBe(5);
  });

  it('ignores blank header cells', () => {
    const { table, sheet } = typed([
      [...HEADER, '', 'name'],
      row('a', 'stray', 'Ann'),
    ]);

    expect(Object.keys(at(table.find(), 0))).not.toContain('');

    table.update({ _id: 'a', name: 'Anne', '': 'overwrite' });

    expect(at(sheet.getDataRange().getValues(), 1).slice(4)).toEqual([
      'stray',
      'Anne',
    ]);
  });

  it('appends new columns after the last used column, past blank headers', () => {
    const { table, sheet } = typed([
      [...HEADER, '', 'name', ''],
      row('a', '', 'Ann', ''),
    ]);

    table.update({ _id: 'a', extra: 1 });

    expect(header(sheet)).toEqual([...HEADER, '', 'name', 'extra']);
  });

  it('keeps the first of duplicate header names and warns once per operation', () => {
    const { db, mocks } = setup();
    const warn = {
      info: () => {},
      warn: (() => {}) as GasSheetDbLogger['warn'],
    };
    const calls: unknown[][] = [];
    warn.warn = (...args) => void calls.push(args);
    const sheet = mocks.activeSpreadsheet.insertSheet('people');
    enter(sheet, [[...HEADER, 'name', 'name'], row('a', 'first', 'second')]);
    const table = db.table({ sheetName: 'people', logger: warn });

    expect(table.find()[0]).toMatchObject({ name: 'first' });
    expect(calls).toEqual([
      ['Duplicate column keys', { sheet: 'people', columns: ['name'] }],
    ]);

    table.update({ _id: 'a', name: 'changed' });

    expect(at(sheet.getDataRange().getValues(), 1).slice(4)).toEqual([
      'changed',
      'second',
    ]);
    expect(calls).toHaveLength(2);
  });

  it('warns once when a read falls back to the locked path', () => {
    const { table, logger } = setupTable();
    table.insert({ name: 'a' });
    const sheet = sheetOf(typedSpreadsheet(table), 'people');
    sheet.getRange(1, 7, 1, 1).setValues([['name']]);
    sheet.getRange(2, 1, 1, 1).setValues([['']]);
    logger.warn.mockClear();

    table.find();

    expect(
      logger.warn.mock.calls.filter(([msg]) => msg === 'Duplicate column keys'),
    ).toHaveLength(1);
  });
});

/** The spreadsheet a table lives in. */
function typedSpreadsheet(table: {
  spreadsheet: GoogleAppsScript.Spreadsheet.Spreadsheet;
}) {
  return table.spreadsheet as unknown as Parameters<typeof sheetOf>[0];
}

describe('a stale cached schema', () => {
  it('does not duplicate system columns when the header is typed after construction', () => {
    const { db, mocks } = setup({ logger: { info: () => {}, warn: () => {} } });
    const table = db.table({ sheetName: 'people' });
    const sheet = sheetOf(mocks.activeSpreadsheet, 'people');
    enter(sheet, [
      [...HEADER, 'name'],
      ['', '', '', '', 'Ann'],
    ]);

    table.updateWhere(() => true, { name: 'Anne' });

    expect(header(sheet)).toEqual([...HEADER, 'name']);
  });
});

describe('empty batches', () => {
  it('return an empty array without taking the lock', () => {
    const { lockService, withLock } = createLockService();
    const { db } = setup({ lockService });
    const table = db.table({ sheetName: 'people' });
    withLock.mockClear();

    expect(table.insertMany([])).toEqual([]);
    expect(table.updateMany([])).toEqual([]);
    expect(table.softDeleteMany([])).toEqual([]);
    expect(table.restoreMany([])).toEqual([]);
    expect(table.deleteMany([])).toBeUndefined();
    expect(withLock).not.toHaveBeenCalled();
  });

  it('do not touch the sheet', () => {
    const { table, sheet } = setupTable();
    table.insert({ name: 'a' });
    sheet.calls.setValues.length = 0;

    table.deleteMany([]);

    expect(sheet.calls.setValues).toHaveLength(0);
    expect(sheet.deleteRows).not.toHaveBeenCalled();
  });
});

describe('batch atomicity', () => {
  const seeded = () => {
    const context = setupTable();
    const entries = context.table.insertMany([
      { name: 'a' },
      { name: 'b' },
      { name: 'c' },
    ]);
    return { ...context, entries };
  };

  it.each([
    [
      'updateMany',
      (t: ReturnType<typeof seeded>['table'], e: GasSheetDbEntry[]) =>
        t.updateMany(e),
    ],
    [
      'softDeleteMany',
      (t: ReturnType<typeof seeded>['table'], e: GasSheetDbEntry[]) =>
        t.softDeleteMany(e),
    ],
    [
      'restoreMany',
      (t: ReturnType<typeof seeded>['table'], e: GasSheetDbEntry[]) =>
        t.restoreMany(e),
    ],
  ])(
    '%s mutates and writes nothing when a later entry is not found',
    (_, run) => {
      const { table, sheet, entries } = seeded();
      const batch = [
        { ...at(entries, 0), name: 'changed' },
        { ...at(entries, 1), _isDeleted: false },
        { _id: 'missing' },
      ];
      const snapshot = batch.map((e) => ({ ...e }));
      const stored = rows(sheet);
      sheet.calls.setValues.length = 0;

      expect(() => run(table, batch)).toThrow('Entry not found: missing');

      expect(batch).toEqual(snapshot);
      expect(rows(sheet)).toEqual(stored);
      expect(sheet.calls.setValues).toHaveLength(0);
    },
  );

  it('does not add columns when the batch is rejected', () => {
    const { table, sheet, entries } = seeded();

    expect(() =>
      table.updateMany([{ ...at(entries, 0), extra: 1 }, { _id: 'missing' }]),
    ).toThrow();

    expect(header(sheet)).not.toContain('extra');
  });

  it('rejects a missing _id before mutating earlier entries', () => {
    const { table, entries } = seeded();
    const first = { ...at(entries, 0) };

    expect(() => table.softDeleteMany([first, {}])).toThrow('without "_id"');

    expect(first._isDeleted).toBe(false);
  });

  it('still mutates the passed objects on success', () => {
    const { table, entries } = seeded();
    const target = at(entries, 0);

    table.softDeleteMany([target]);
    expect(target._isDeleted).toBe(true);

    table.restoreMany([target]);
    expect(target._isDeleted).toBe(false);
  });
});

describe('_isDeleted backfill', () => {
  const read = (cell: unknown) => {
    const { table, sheet } = typed([
      [...HEADER, 'name'],
      ['a', NOW, NOW, cell, 'Ann'],
    ]);
    const logged: unknown[][] = [];
    table.logger.info = (...args: unknown[]) => void logged.push(args);
    return { entries: table.find({ withTrashed: true }), sheet, logged };
  };

  it.each([
    ['', false],
    [1, true],
    [0, false],
    ['yes', true],
    [true, true],
    [false, false],
  ])('turns %j into %j', (cell, expected) => {
    const { entries, sheet } = read(cell);

    expect(at(entries, 0)._isDeleted).toBe(expected);
    expect(at(rows(sheet), 0)._isDeleted).toBe(expected);
  });

  it('keeps a deleted row out of find()', () => {
    const { table } = typed([[...HEADER], ['a', NOW, NOW, 1]]);

    expect(table.find()).toEqual([]);
    expect(table.findTrashed()).toHaveLength(1);
  });

  it('counts the normalised row in the backfill log', () => {
    const { table } = typed([[...HEADER], ['a', NOW, NOW, 1]]);
    const info: unknown[][] = [];
    table.logger.info = (...args: unknown[]) => void info.push(args);

    table.find();

    expect(info).toContainEqual([
      'Backfilled system metadata',
      { sheet: 'people', rows: 1 },
    ]);
  });
});

describe('duplicate _id values', () => {
  const dups = () =>
    typed([
      [...HEADER, 'name'],
      row('x', 'first'),
      row('y', 'other'),
      row('x', 'second'),
      row('x', 'third'),
    ]);

  it('reassigns later rows through the locked repair path', () => {
    const { table, sheet } = dups();
    const logged: unknown[][] = [];
    table.logger.info = (...args: unknown[]) => void logged.push(args);

    const entries = table.find();

    const ids = entries.map((e) => e._id);
    expect(ids[0]).toBe('x');
    expect(ids[1]).toBe('y');
    expect(new Set(ids).size).toBe(4);
    expect(entries.map((e) => e.name)).toEqual([
      'first',
      'other',
      'second',
      'third',
    ]);
    expect(rows(sheet).map((r) => r._id)).toEqual(ids);
    expect(logged).toEqual([
      ['Reassigned duplicate ids', { sheet: 'people', rows: 2 }],
    ]);
  });

  it('does not log a backfill for rows that only had a duplicate id', () => {
    const { table } = dups();
    const logged: string[] = [];
    table.logger.info = (msg: string) => void logged.push(msg);

    table.find();

    expect(logged).not.toContain('Backfilled system metadata');
  });

  it('is stable on the next read', () => {
    const { table, sheet } = dups();
    table.find();
    sheet.calls.setValues.length = 0;

    table.find();

    expect(sheet.calls.setValues).toHaveLength(0);
  });
});

describe('insert conflicts', () => {
  const seeded = () => {
    const context = setupTable();
    const [ann] = context.table.insertMany([{ name: 'Ann' }, { name: 'Bob' }]);
    return { ...context, ann: ann! };
  };

  it('rejects an _id that already exists, writing and mutating nothing', () => {
    const { table, sheet, ann } = seeded();
    const fresh: GasSheetDbEntry = { name: 'Zed' };
    const clash: GasSheetDbEntry = { _id: ann._id!, name: 'Dup' };
    const stored = rows(sheet);
    sheet.calls.setValues.length = 0;

    let err: Error | undefined;
    try {
      table.insertMany([fresh, clash]);
    } catch (e) {
      err = e as Error;
    }

    expect(err).toMatchObject({
      message: `[GasSheetDb] Duplicate "_id": ${ann._id}`,
      statusCode: 409,
      code: 'CONFLICT',
    });
    expect(fresh).toEqual({ name: 'Zed' });
    expect(clash).toEqual({ _id: ann._id, name: 'Dup' });
    expect(rows(sheet)).toEqual(stored);
    expect(sheet.calls.setValues).toHaveLength(0);
  });

  it('rejects an _id repeated within the batch', () => {
    const { table, sheet } = seeded();
    const stored = rows(sheet);

    expect(() =>
      table.insertMany([
        { _id: 'new', name: 'a' },
        { _id: 'new', name: 'b' },
      ]),
    ).toThrow('[GasSheetDb] Duplicate "_id": new');

    expect(rows(sheet)).toEqual(stored);
  });

  it('rejects inserting an entry obtained from find()', () => {
    const { table } = seeded();

    expect(() => table.insert(at(table.find(), 0))).toThrow('Duplicate "_id"');
  });

  it('accepts a caller-chosen _id that is free', () => {
    const { table } = seeded();

    expect(table.insert({ _id: 'mine', name: 'Cy' })._id).toBe('mine');
  });

  it('reads the sheet once per insert, with or without an _id', () => {
    const { table, sheet } = seeded();
    const read = sheet.getRange;
    let reads = 0;
    sheet.getRange = ((...args: Parameters<typeof read>) => {
      const range = read(...args);
      const values = range.getValues.bind(range);
      range.getValues = () => (reads++, values());
      return range;
    }) as typeof read;

    table.insert({ name: 'no id' });
    const withoutId = reads;
    reads = 0;
    table.insert({ _id: 'mine', name: 'with id' });

    expect(withoutId).toBeGreaterThan(0);
    expect(reads).toBe(withoutId);
  });
});

describe('a stale cached schema on delete', () => {
  it('deletes the right rows after a column was inserted before _id', () => {
    const { table, sheet, logger } = setupTable();
    const [a, b, c] = table.insertMany([
      { name: 'a' },
      { name: 'b' },
      { name: 'c' },
    ]);
    const shifted = sheet
      .getDataRange()
      .getValues()
      .map((values, i) => [i === 0 ? 'name' : 'x', ...values]);
    enter(sheet, shifted);
    logger.warn.mockClear();

    table.deleteMany([b!]);

    expect(rows(sheet).map((r) => r._id)).toEqual([a!._id, c!._id]);
    expect(logger.warn).toHaveBeenCalledTimes(1);
    expect(logger.warn).toHaveBeenCalledWith('Duplicate column keys', {
      sheet: 'people',
      columns: ['name'],
    });
  });
});

describe('the insert position', () => {
  const SPILL_COLUMN = 6;
  const NAME_COLUMN = 1;
  const OPEN_ENDED = '=ARRAYFORMULA(IF(LEN(B2:B), ROW(B2:B), ""))';

  /** Ann and Bob inserted through the library, with a `rank` column after them. */
  const seeded = (context = setupTable()) => {
    context.table.insertMany([{ name: 'Ann' }, { name: 'Bob' }]);
    context.sheet.getRange(1, SPILL_COLUMN).setValues([['rank']]);
    context.sheet.calls.setValues.length = 0;
    return context;
  };

  /** Rows 2 to the last row of the sheet, `''` below the two entries. */
  const openEnded = (sheet: FakeSheet) =>
    sheet.spill(
      2,
      SPILL_COLUMN,
      Array.from({ length: sheet.getMaxRows() - 1 }, (_, i) =>
        i < 2 ? i + 2 : '',
      ),
      OPEN_ENDED,
    );

  /** Whether any `setValues` call wrote to column `col`. */
  const wroteColumn = (sheet: FakeSheet, col: number) =>
    sheet.calls.setValues.some((c) => c.col <= col && col < c.col + c.numCols);

  const names = (sheet: FakeSheet, col = NAME_COLUMN) =>
    sheet
      .getRange(2, col, 12, 1)
      .getValues()
      .map(([name]) => name);

  it('lands below an open-ended spill, leaving it intact', () => {
    const { table, sheet } = seeded();
    openEnded(sheet);
    expect(sheet.getLastRow()).toBe(1000);

    table.insert({ name: 'Cy' });

    expect(names(sheet).slice(0, 4)).toEqual(['Ann', 'Bob', 'Cy', '']);
    expect(sheet.brokenSpills).toEqual([]);
    expect(sheet.getRange(2, SPILL_COLUMN).getFormulas()).toEqual([
      [OPEN_ENDED],
    ]);
    expect(wroteColumn(sheet, SPILL_COLUMN)).toBe(false);
    expect(sheet.insertRowsAfter).not.toHaveBeenCalled();
    expect(table.find().map((e) => e.name)).toEqual(['Ann', 'Bob', 'Cy']);
  });

  it('keeps landing in the next row across inserts', () => {
    const { table, sheet } = seeded();
    openEnded(sheet);

    table.insert({ name: 'Cy' });
    table.insertMany([{ name: 'Dee' }, { name: 'Eve' }]);
    table.insert({ name: 'Fay' });

    expect(names(sheet).slice(0, 6)).toEqual([
      'Ann',
      'Bob',
      'Cy',
      'Dee',
      'Eve',
      'Fay',
    ]);
    expect(sheet.brokenSpills).toEqual([]);
  });

  it('writes a batch into blank rows with one call per run of columns', () => {
    const { table, sheet } = seeded();
    openEnded(sheet);

    table.insertMany([{ name: 'Cy' }, { name: 'Dee' }, { name: 'Eve' }]);

    expect(
      sheet.calls.setValues.map((c) => [c.row, c.col, c.numRows, c.numCols]),
    ).toEqual([[4, 1, 3, 5]]);
  });

  it('does not write the cells of an entry it leaves unset', () => {
    const { table, sheet } = seeded();
    table.insert({ name: 'Cy', note: 'x' });
    sheet.calls.setValues.length = 0;
    openEnded(sheet);

    table.insert({ name: 'Dee' });

    // `note` sits after `rank`; the entry does not set it, nor the spill
    expect(wroteColumn(sheet, SPILL_COLUMN)).toBe(false);
    expect(wroteColumn(sheet, SPILL_COLUMN + 1)).toBe(false);
  });

  it('lands in the first row of an empty table below a spill', () => {
    const { table, sheet } = setupTable();
    table.find();
    sheet.getRange(1, SPILL_COLUMN).setValues([['rank']]);
    sheet.spill(2, SPILL_COLUMN, Array(999).fill(''), OPEN_ENDED);

    table.insert({ name: 'Ann' });

    // `name` is appended after the last used column, the spill's
    expect(names(sheet, SPILL_COLUMN + 1)[0]).toBe('Ann');
    expect(sheet.brokenSpills).toEqual([]);
  });

  it('does not write a closed spill', () => {
    const { table, sheet } = seeded();
    sheet.spill(2, SPILL_COLUMN, [2, 3], '=ARRAYFORMULA(ROW(B2:B3))');

    table.insert({ name: 'Cy' });

    expect(names(sheet).slice(0, 3)).toEqual(['Ann', 'Bob', 'Cy']);
    expect(sheet.brokenSpills).toEqual([]);
    expect(wroteColumn(sheet, SPILL_COLUMN)).toBe(false);
  });

  it('keeps a per-row formula column that was filled below the entries', () => {
    const { table, sheet } = seeded();
    sheet
      .getRange(2, SPILL_COLUMN, 3, 1)
      .setValues([['=R2'], ['=R3'], ['=R4']]);
    sheet.calls.setValues.length = 0;

    table.insert({ name: 'Cy' });

    expect(sheet.getRange(2, SPILL_COLUMN, 3, 1).getFormulas()).toEqual([
      ['=R2'],
      ['=R3'],
      ['=R4'],
    ]);
    expect(names(sheet).slice(0, 3)).toEqual(['Ann', 'Bob', 'Cy']);
    expect(wroteColumn(sheet, SPILL_COLUMN)).toBe(false);
  });

  it('treats a row whose formula returns a value as taken', () => {
    const { table, sheet } = seeded();
    sheet.getRange(4, SPILL_COLUMN).setValues([['=R4']]);
    sheet.evaluate = () => 'filled';

    table.insert({ name: 'Cy' });

    expect(names(sheet).slice(0, 4)).toEqual(['Ann', 'Bob', '', 'Cy']);
    expect(sheet.getRange(4, SPILL_COLUMN).getFormulas()).toEqual([['=R4']]);
  });

  it('does not overwrite a row a person typed without metadata', () => {
    const { table, sheet } = typed([
      [...HEADER, 'name'],
      row('a', 'Ann'),
      ['', '', '', '', 'Typed'],
    ]);

    table.insert({ name: 'Cy' });

    expect(sheet.getRange(3, 5).getValues()).toEqual([['Typed']]);
    expect(sheet.getRange(4, 5).getValues()).toEqual([['Cy']]);
  });

  it('does not fill a blank row between entries', () => {
    const { table, sheet } = typed([
      [...HEADER, 'name'],
      row('a', 'Ann'),
      Array(5).fill(''),
      row('b', 'Bob'),
    ]);

    table.insert({ name: 'Cy' });

    expect(sheet.getRange(3, 1, 1, 5).getValues()).toEqual([Array(5).fill('')]);
    expect(sheet.getRange(5, 5).getValues()).toEqual([['Cy']]);
  });

  it('reuses blank rows below the entries before adding any', () => {
    const { table, sheet } = typed([
      [...HEADER, 'name'],
      row('a', 'Ann'),
      Array(5).fill(''),
      Array(5).fill(''),
    ]);

    table.insertMany([{ name: 'Bob' }, { name: 'Cy' }, { name: 'Dee' }]);

    expect(names(sheet, 5).slice(0, 4)).toEqual(['Ann', 'Bob', 'Cy', 'Dee']);
    expect(sheet.insertRowsAfter).not.toHaveBeenCalled();
  });

  it('adds rows only when the sheet runs out', () => {
    const { table, sheet } = setupTable();
    table.insertMany(
      Array.from({ length: 999 }, (_, i) => ({ name: `n${i}` })),
    );
    expect(sheet.getLastRow()).toBe(1000);

    table.insert({ name: 'last' });

    expect(sheet.insertRowsAfter).toHaveBeenCalledWith(1000, 1);
    expect(sheet.getRange(1001, NAME_COLUMN).getValues()).toEqual([['last']]);
  });
});
