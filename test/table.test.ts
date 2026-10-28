import { describe, expect, it } from 'vitest';
import {
  at,
  createLockService,
  createLogger,
  enter,
  header,
  rows,
  setup,
  setupProbed,
  setupTable,
  sheetOf,
} from './helpers';

const SYSTEM_COLUMNS = ['_id', '_createdAt', '_updatedAt', '_isDeleted'];

/** A table pre-filled with Ann, Bob and Cy through the library. */
function seeded(options: GasSheetDbOptions = {}) {
  const context = setupTable('people', options);
  const entries = context.table.insertMany([
    { name: 'Ann', age: 30 },
    { name: 'Bob', age: 40 },
    { name: 'Cy', age: 50 },
  ]);
  context.logger.info.mockClear();
  return { ...context, entries };
}

const names = (entries: GasSheetDbEntry[]): unknown[] =>
  entries.map((entry) => entry.name);

describe('insert', () => {
  it('adds system metadata to the entry and to the stored row', () => {
    const { table, sheet } = setupTable();
    const entry: GasSheetDbEntry = { name: 'Ann' };

    const inserted = table.insert(entry);

    expect(entry).toMatchObject({
      _id: '00000000-0000-4000-8000-000000000001',
      _isDeleted: false,
    });
    expect(entry._createdAt).toBeInstanceOf(Date);
    expect(entry._updatedAt).toBeInstanceOf(Date);
    expect(inserted).toEqual(entry);
    expect(rows(sheet)).toEqual([
      {
        _id: entry._id,
        _createdAt: entry._createdAt,
        _updatedAt: entry._updatedAt,
        _isDeleted: false,
        name: 'Ann',
      },
    ]);
  });

  it('keeps a supplied _id and _createdAt', () => {
    const { table } = setupTable();
    const createdAt = new Date('2020-01-01T00:00:00Z');

    const inserted = table.insert({ _id: 'custom', _createdAt: createdAt });

    expect(inserted._id).toBe('custom');
    expect(inserted._createdAt).toEqual(createdAt);
  });

  it('inserts many entries in order and returns them decoded', () => {
    const { table, sheet } = setupTable();

    const inserted = table.insertMany([{ name: 'Ann' }, { name: 'Bob' }]);

    expect(names(inserted)).toEqual(['Ann', 'Bob']);
    expect(names(rows(sheet))).toEqual(['Ann', 'Bob']);
    expect(new Set(inserted.map((e) => e._id)).size).toBe(2);
  });

  it('creates columns for new keys and logs them once', () => {
    const { table, sheet, logger } = setupTable();

    table.insert({ name: 'Ann' });
    table.insertMany([{ name: 'Bob', age: 40 }, { city: 'Lima' }]);

    expect(header(sheet).slice().sort()).toEqual(
      [...SYSTEM_COLUMNS, 'age', 'city', 'name'].sort(),
    );
    expect(rows(sheet).map((r) => [r.name, r.age, r.city])).toEqual([
      ['Ann', '', ''],
      ['Bob', 40, ''],
      ['', '', 'Lima'],
    ]);
    expect(logger.info).toHaveBeenCalledWith('Added columns', {
      sheet: 'people',
      columns: ['age', 'city'],
    });
  });

  it('grows the sheet when rows or columns run out', () => {
    const { table, sheet } = setupTable();
    const wide = Object.fromEntries(
      Array.from({ length: 30 }, (_, i) => [`col${i}`, i]),
    );

    table.insert(wide);
    table.insertMany(Array.from({ length: 1000 }, (_, i) => ({ n: i })));

    expect(sheet.insertColumnsAfter).toHaveBeenCalled();
    expect(sheet.insertRowsAfter).toHaveBeenCalled();
    expect(table.find()).toHaveLength(1001);
  });

  it('round-trips a Date and nullish values', () => {
    const { table } = setupTable();
    const when = new Date('2024-05-06T07:08:09Z');

    const [stored] = table.insertMany([
      { when, nothing: null, gap: undefined },
    ]);

    expect(stored?.when).toEqual(when);
    expect(stored?.nothing).toBe('');
    expect(table.find()[0]?.when).toEqual(when);
  });
});

describe('JSON codec', () => {
  it('stores objects and arrays as prefixed JSON and restores them', () => {
    const { table, sheet } = setupTable();
    const entry = { tags: ['a', 'b'], meta: { nested: { ok: true } } };

    table.insert(entry);

    expect(rows(sheet)[0]?.tags).toBe('__JSON__["a","b"]');
    expect(table.find()[0]).toMatchObject(entry);
  });

  it('falls back to the raw string and warns once per read', () => {
    const { table, sheet, logger } = setupTable();
    table.insertMany([{ data: 'x' }, { data: 'y' }, { data: 'z' }]);
    const col = header(sheet).indexOf('data') + 1;
    sheet
      .getRange(2, col, 3, 1)
      .setValues([['__JSON__{bad'], ['__JSON__[1'], ['__JSON__']]);

    const found = table.find();

    expect(found.map((e) => e.data)).toEqual([
      '__JSON__{bad',
      '__JSON__[1',
      '__JSON__',
    ]);
    expect(logger.warn).toHaveBeenCalledExactlyOnceWith(
      'Failed to decode JSON cells',
      { sheet: 'people', cells: 3, error: expect.any(String) },
    );
  });
});

describe('find', () => {
  it('returns an empty array for a fresh table and adds the system columns', () => {
    const { table, sheet } = setupTable();

    expect(table.find()).toEqual([]);
    expect(header(sheet)).toEqual(SYSTEM_COLUMNS);
  });

  it('does not write when there is nothing to repair', () => {
    const { table, sheet } = seeded();
    const writes = sheet.calls.setValues.length;

    table.find();
    table.find();

    expect(sheet.calls.setValues).toHaveLength(writes);
  });

  it('filters with findWhere and findOneWhere', () => {
    const { table } = seeded();

    expect(names(table.findWhere((e) => (e.age as number) >= 40))).toEqual([
      'Bob',
      'Cy',
    ]);
    expect(table.findWhere(() => false)).toEqual([]);
    expect(table.findOneWhere((e) => e.name === 'Bob')).toMatchObject({
      age: 40,
    });
    expect(table.findOneWhere(() => false)).toBeNull();
  });

  it('handles trashed entries through find options and findTrashed', () => {
    const { table, entries } = seeded();
    table.softDelete(at(entries, 1));

    expect(names(table.find())).toEqual(['Ann', 'Cy']);
    expect(names(table.find({ withTrashed: true }))).toEqual([
      'Ann',
      'Bob',
      'Cy',
    ]);
    expect(names(table.find({ onlyTrashed: true }))).toEqual(['Bob']);
    expect(names(table.findTrashed())).toEqual(['Bob']);
    expect(table.findTrashed((e) => e.name === 'Ann')).toEqual([]);
    expect(names(table.findTrashed((e) => e.name === 'Bob'))).toEqual(['Bob']);
  });
});

describe('backfill of manually entered rows', () => {
  /** A table over a sheet a user filled in before the library saw it. */
  function manual(grid: unknown[][]) {
    const logger = createLogger();
    const { db, mocks } = setup({ logger });
    const sheet = mocks.activeSpreadsheet.insertSheet('people');
    enter(sheet, grid);
    return { table: db.table({ sheetName: 'people' }), sheet, logger };
  }

  it('assigns metadata to rows typed into the sheet and logs it once', () => {
    const { table, logger } = manual([['name'], ['Ann'], ['Bob']]);

    const found = table.find();

    expect(names(found)).toEqual(['Ann', 'Bob']);
    for (const entry of found) {
      expect(entry._id).toEqual(expect.any(String));
      expect(entry._createdAt).toBeInstanceOf(Date);
      expect(entry._updatedAt).toBeInstanceOf(Date);
      expect(entry._isDeleted).toBe(false);
    }
    expect(new Set(found.map((e) => e._id)).size).toBe(2);
    expect(logger.info).toHaveBeenCalledWith('Backfilled system metadata', {
      sheet: 'people',
      rows: 2,
    });

    logger.info.mockClear();
    table.find();
    expect(logger.info).not.toHaveBeenCalled();
  });

  it('keeps ids that were already present', () => {
    const { table } = manual([
      ['_id', 'name'],
      ['keep-me', 'Ann'],
      ['', 'Bob'],
    ]);

    const found = table.find();

    expect(found.map((e) => e._id)).toEqual([
      'keep-me',
      expect.stringMatching(/^0{8}-/),
    ]);
  });
});

describe('update', () => {
  it('applies a partial update and preserves other fields', () => {
    const { table, sheet, entries } = seeded();
    const before = at(entries, 0);

    const updated = table.update({ _id: before._id!, age: 31 });

    expect(updated).toMatchObject({ name: 'Ann', age: 31, _id: before._id });
    expect(updated._updatedAt).toBeInstanceOf(Date);
    expect(rows(sheet).map((r) => r.age)).toEqual([31, 40, 50]);
  });

  it('adds columns for new keys', () => {
    const { table, sheet, entries } = seeded();

    table.update({ _id: at(entries, 0)._id!, nick: 'A' });

    expect(header(sheet)).toContain('nick');
    expect(rows(sheet).map((r) => r.nick)).toEqual(['A', '', '']);
  });

  it('updates many entries at once', () => {
    const { table, entries } = seeded();

    const updated = table.updateMany([
      { _id: at(entries, 2)._id!, age: 51 },
      { _id: at(entries, 0)._id!, age: 31 },
    ]);

    expect(updated.map((e) => e.age)).toEqual([51, 31]);
    expect(table.find().map((e) => e.age)).toEqual([31, 40, 51]);
  });

  it('throws without an _id', () => {
    const { table } = seeded();

    expect(() => table.update({ name: 'x' })).toThrow(
      '[GasSheetDb] Cannot update entry without "_id"',
    );
  });

  it('throws for an unknown _id', () => {
    const { table } = seeded();

    expect(() => table.update({ _id: 'nope' })).toThrow(
      '[GasSheetDb] Entry not found: nope',
    );
  });
});

describe('updateWhere / updateOneWhere', () => {
  it('applies a patch object to every match', () => {
    const { table } = seeded();

    const updated = table.updateWhere((e) => (e.age as number) >= 40, {
      status: 'senior',
    });

    expect(names(updated)).toEqual(['Bob', 'Cy']);
    expect(table.find().map((e) => e.status)).toEqual(['', 'senior', 'senior']);
  });

  it('applies the patch returned by an updater function', () => {
    const { table } = seeded();

    table.updateWhere(
      () => true,
      (e) => ({ age: (e.age as number) + 1 }),
    );

    expect(table.find().map((e) => e.age)).toEqual([31, 41, 51]);
  });

  it('honours an updater that mutates the entry and returns nothing', () => {
    const { table } = seeded();

    table.updateWhere(
      (e) => e.name === 'Ann',
      (e) => {
        e.age = 99;
      },
    );

    expect(table.find().map((e) => e.age)).toEqual([99, 40, 50]);
  });

  it('cannot be redirected to another row through _id', () => {
    const { table, entries } = seeded();

    table.updateWhere(
      (e) => e.name === 'Ann',
      () => ({ _id: at(entries, 2)._id!, name: 'Hijacked' }),
    );

    expect(names(table.find())).toEqual(['Hijacked', 'Bob', 'Cy']);
  });

  it('excludes trashed entries unless asked', () => {
    const { table, entries } = seeded();
    table.softDelete(at(entries, 0));

    expect(table.updateWhere(() => true, { flag: 1 })).toHaveLength(2);
    expect(
      names(table.updateWhere(() => true, { flag: 2 }, { onlyTrashed: true })),
    ).toEqual(['Ann']);
  });

  it('acquires the lock once across read and write', () => {
    const { lockService, withLock } = createLockService();
    const { table } = seeded({ lockService });
    withLock.mockClear();

    table.updateWhere(() => true, { flag: 1 });
    expect(withLock).toHaveBeenCalledTimes(1);

    table.updateOneWhere(() => true, { flag: 2 });
    expect(withLock).toHaveBeenCalledTimes(2);
  });

  it('updates only the first match with updateOneWhere', () => {
    const { table } = seeded();

    const updated = table.updateOneWhere((e) => (e.age as number) > 30, {
      age: 0,
    });

    expect(updated).toMatchObject({ name: 'Bob', age: 0 });
    expect(table.find().map((e) => e.age)).toEqual([30, 0, 50]);
  });

  it('returns an empty array or null when nothing matches', () => {
    const { table, sheet } = seeded();
    const writes = sheet.calls.setValues.length;

    expect(table.updateWhere(() => false, { x: 1 })).toEqual([]);
    expect(table.updateOneWhere(() => false, { x: 1 })).toBeNull();
    expect(sheet.calls.setValues).toHaveLength(writes);
  });
});

describe('soft delete and restore', () => {
  it('trashes and restores a single entry', () => {
    const { table, entries } = seeded();
    const ann = at(entries, 0);

    const trashed = table.softDelete(ann);

    expect(trashed._isDeleted).toBe(true);
    expect(ann._isDeleted).toBe(true);
    expect(names(table.find())).toEqual(['Bob', 'Cy']);

    const restored = table.restore(ann);

    expect(restored._isDeleted).toBe(false);
    expect(names(table.find())).toEqual(['Ann', 'Bob', 'Cy']);
  });

  it('trashes and restores many entries', () => {
    const { table, entries } = seeded();
    const targets = [at(entries, 0), at(entries, 2)];

    table.softDeleteMany(targets);
    expect(names(table.find())).toEqual(['Bob']);
    expect(names(table.findTrashed())).toEqual(['Ann', 'Cy']);

    table.restoreMany(targets);
    expect(names(table.find())).toEqual(['Ann', 'Bob', 'Cy']);
  });

  it('trashes matches with softDeleteWhere, skipping already trashed ones', () => {
    const { table, entries } = seeded();
    table.softDelete(at(entries, 0));

    const trashed = table.softDeleteWhere(() => true);

    expect(names(trashed)).toEqual(['Bob', 'Cy']);
    expect(table.find()).toEqual([]);
    expect(table.softDeleteWhere(() => true)).toEqual([]);
  });

  it('throws when the entry has no _id or does not exist', () => {
    const { table } = seeded();

    expect(() => table.softDelete({ name: 'x' })).toThrow(
      '[GasSheetDb] Cannot update entry without "_id"',
    );
    expect(() => table.restore({ _id: 'nope' })).toThrow(
      '[GasSheetDb] Entry not found: nope',
    );
  });
});

describe('delete', () => {
  it('permanently deletes one entry', () => {
    const { table, sheet, entries } = seeded();

    table.deleteOne(at(entries, 1));

    expect(names(table.find({ withTrashed: true }))).toEqual(['Ann', 'Cy']);
    expect(rows(sheet)).toHaveLength(2);
  });

  it('permanently deletes many entries, trashed or not', () => {
    const { table, entries } = seeded();
    table.softDelete(at(entries, 2));

    table.deleteMany([at(entries, 0), at(entries, 2)]);

    expect(names(table.find({ withTrashed: true }))).toEqual(['Bob']);
  });

  it('deletes matches with deleteWhere and returns them as they were', () => {
    const { table } = seeded();

    const deleted = table.deleteWhere((e) => (e.age as number) < 45);

    expect(names(deleted)).toEqual(['Ann', 'Bob']);
    expect(names(table.find())).toEqual(['Cy']);
    expect(table.deleteWhere(() => false)).toEqual([]);
  });

  it('leaves the trash alone unless onlyTrashed is set', () => {
    const { table, entries } = seeded();
    table.softDelete(at(entries, 0));

    expect(names(table.deleteWhere((e) => e.name === 'Ann'))).toEqual([]);
    expect(names(table.deleteWhere(() => true, { onlyTrashed: true }))).toEqual(
      ['Ann'],
    );
    expect(names(table.find({ withTrashed: true }))).toEqual(['Bob', 'Cy']);
  });

  it('throws without an _id or for an unknown _id', () => {
    const { table } = seeded();

    expect(() => table.deleteOne({ name: 'x' })).toThrow(
      '[GasSheetDb] Cannot delete entry without "_id"',
    );
    expect(() => table.deleteMany([{ _id: 'nope' }])).toThrow(
      '[GasSheetDb] Entry not found: nope',
    );
  });
});

describe('read path', () => {
  /** Counts of the sheet mutations a read must not perform. */
  const mutations = (sheet: ReturnType<typeof setupTable>['sheet']) => ({
    setValues: sheet.calls.setValues.length,
    clearContent: sheet.calls.clearContent.length,
    deleteRows: sheet.deleteRows.mock.calls.length,
    insertRows: sheet.insertRowsAfter.mock.calls.length,
    insertColumns: sheet.insertColumnsAfter.mock.calls.length,
  });

  it.each([
    ['find', (t: ReturnType<typeof seeded>['table']) => t.find()],
    [
      'findWhere',
      (t: ReturnType<typeof seeded>['table']) => t.findWhere(() => true),
    ],
    [
      'findOneWhere',
      (t: ReturnType<typeof seeded>['table']) => t.findOneWhere(() => true),
    ],
    ['findTrashed', (t: ReturnType<typeof seeded>['table']) => t.findTrashed()],
  ])(
    '%s of a healthy table takes no lock and writes nothing',
    (_name, read) => {
      const { lockService, withLock } = createLockService();
      const { table, sheet, mocks } = seeded({ lockService });
      withLock.mockClear();
      mocks.SpreadsheetApp.flush.mockClear();
      const before = mutations(sheet);

      read(table);

      expect(withLock).not.toHaveBeenCalled();
      expect(mocks.SpreadsheetApp.flush).not.toHaveBeenCalled();
      expect(mutations(sheet)).toEqual(before);
    },
  );

  it('returns the same entries as it did under the lock', () => {
    const { table, entries } = seeded();
    table.softDelete(at(entries, 1));

    expect(names(table.find())).toEqual(['Ann', 'Cy']);
    expect(names(table.find({ withTrashed: true }))).toEqual([
      'Ann',
      'Bob',
      'Cy',
    ]);
    expect(names(table.findTrashed())).toEqual(['Bob']);
    expect(names(table.findWhere((e) => (e.age as number) > 30))).toEqual([
      'Cy',
    ]);
    expect(table.findOneWhere((e) => e.name === 'Cy')).toMatchObject({
      age: 50,
    });
  });

  it('reads correctly after rows were deleted, without trimming the sheet', () => {
    const { table, sheet, entries } = seeded();
    table.deleteMany([at(entries, 0), at(entries, 2)]);
    sheet.insertRowsAfter(sheet.getMaxRows(), 5);
    sheet.deleteRows.mockClear();

    expect(names(table.find())).toEqual(['Bob']);
    expect(sheet.deleteRows).not.toHaveBeenCalled();
  });

  it('reads an emptied table', () => {
    const { table, entries } = seeded();
    table.deleteMany(entries);

    expect(table.find()).toEqual([]);
    expect(table.findTrashed()).toEqual([]);
  });

  it('reads a table whose header is not on the first row', () => {
    const { lockService, withLock } = createLockService();
    const { db, mocks } = setup({
      lockService,
      rowNumbers: { columnKeys: 2, firstData: 3 },
    });
    enter(mocks.activeSpreadsheet.insertSheet('offset'), [
      [''],
      ['name'],
      ['Ann'],
    ]);
    const table = db.table({ sheetName: 'offset' });

    expect(names(table.find())).toEqual(['Ann']);
    withLock.mockClear();
    expect(names(table.find())).toEqual(['Ann']);
    expect(withLock).not.toHaveBeenCalled();
  });

  it('takes exactly one lock to backfill rows typed into the sheet', () => {
    const { lockService, withLock } = createLockService();
    const { db, mocks } = setup({ lockService });
    enter(mocks.activeSpreadsheet.insertSheet('people'), [
      ['name'],
      ['Ann'],
      ['Bob'],
    ]);
    const table = db.table({ sheetName: 'people' });

    expect(names(table.find())).toEqual(['Ann', 'Bob']);
    expect(withLock).toHaveBeenCalledTimes(1);

    withLock.mockClear();
    expect(names(table.find())).toEqual(['Ann', 'Bob']);
    expect(withLock).not.toHaveBeenCalled();
  });

  it('takes the lock to add missing system columns', () => {
    const { lockService, withLock } = createLockService();
    const { table } = setupTable('people', { lockService });
    withLock.mockClear();

    expect(table.find()).toEqual([]);

    expect(withLock).toHaveBeenCalledTimes(1);
  });

  it('takes the lock to backfill a single row that lacks an id', () => {
    const { lockService, withLock } = createLockService();
    const { table, sheet } = seeded({ lockService });
    sheet.getRange(3, header(sheet).indexOf('_id') + 1).setValues([['']]);
    withLock.mockClear();

    const found = table.find();

    expect(withLock).toHaveBeenCalledTimes(1);
    expect(found.every((entry) => entry._id)).toBe(true);
  });
});

describe('logging outside the lock', () => {
  it('emits "Created table sheet" after the lock is released', () => {
    const { db, logger, lockedAtLog } = setupProbed();

    db.table({ sheetName: 'people' });

    expect(logger.info).toHaveBeenCalledWith('Created table sheet', {
      sheet: 'people',
    });
    expect(lockedAtLog).toEqual([false]);
  });

  it('emits "Added columns" after the lock is released', () => {
    const { db, mocks, logger, lockedAtLog } = setupProbed();
    const table = db.table({ sheetName: 'people' });
    logger.info.mockClear();
    lockedAtLog.length = 0;
    mocks.LockService.getScriptLock.mockClear();

    table.insert({ name: 'Ann' });

    expect(mocks.LockService.getScriptLock).toHaveBeenCalledOnce();
    expect(logger.info).toHaveBeenCalledWith('Added columns', {
      sheet: 'people',
      columns: expect.arrayContaining(['name']),
    });
    expect(lockedAtLog.length).toBeGreaterThan(0);
    expect(lockedAtLog.every((locked) => !locked)).toBe(true);
  });

  it('emits "Backfilled system metadata" after the lock is released', () => {
    const { db, mocks, logger, lockedAtLog } = setupProbed();
    enter(mocks.activeSpreadsheet.insertSheet('people'), [['name'], ['Ann']]);
    const table = db.table({ sheetName: 'people' });

    table.find();

    expect(logger.info).toHaveBeenCalledWith('Backfilled system metadata', {
      sheet: 'people',
      rows: 1,
    });
    expect(lockedAtLog.every((locked) => !locked)).toBe(true);
  });

  it('emits "Failed to decode JSON cells" after the lock is released', () => {
    const { db, mocks, logger, lockedAtLog } = setupProbed();
    const table = db.table({ sheetName: 'people' });
    table.insert({ data: 'x' });
    const sheet = sheetOf(mocks.activeSpreadsheet, 'people');
    const col = header(sheet).indexOf('data') + 1;
    sheet.getRange(2, col).setValues([['__JSON__{bad']]);
    logger.warn.mockClear();
    lockedAtLog.length = 0;
    mocks.LockService.getScriptLock.mockClear();

    table.updateWhere(() => true, { flag: 1 });

    expect(mocks.LockService.getScriptLock).toHaveBeenCalledOnce();
    expect(logger.warn).toHaveBeenCalledWith(
      'Failed to decode JSON cells',
      expect.objectContaining({ sheet: 'people', cells: 1 }),
    );
    expect(lockedAtLog.every((locked) => !locked)).toBe(true);
  });

  it('emits queued lines in order', () => {
    const { db, logger } = setupProbed();
    const table = db.table({ sheetName: 'people' });
    logger.info.mockClear();

    table.insert({ name: 'Ann' });
    table.insert({ nick: 'A' });

    expect(logger.info.mock.calls.map(([msg]) => msg)).toEqual([
      'Added columns',
      'Added columns',
    ]);
  });

  it('still emits the lines when the operation throws', () => {
    const { db, logger, lockedAtLog } = setupProbed();
    const table = db.table({ sheetName: 'people' });
    table.insert({ name: 'Ann' });
    logger.info.mockClear();
    lockedAtLog.length = 0;

    expect(() => table.update({ _id: 'missing', extra: 1 })).toThrow(
      '[GasSheetDb] Entry not found: missing',
    );

    expect(logger.info).toHaveBeenCalledWith('Added columns', {
      sheet: 'people',
      columns: ['extra'],
    });
    expect(lockedAtLog).toEqual([false]);
  });

  it('does not mask an operation error with a throwing logger', () => {
    const { db, logger } = setupProbed();
    const table = db.table({ sheetName: 'people' });
    table.insert({ name: 'Ann' });
    logger.info.mockImplementation(() => {
      throw new Error('logger down');
    });

    expect(() => table.update({ _id: 'missing', extra: 1 })).toThrow(
      '[GasSheetDb] Entry not found: missing',
    );
  });

  it('does not fail a completed operation over a throwing logger', () => {
    const { db, mocks, logger } = setupProbed();
    const table = db.table({ sheetName: 'people' });
    logger.info.mockImplementation(() => {
      throw new Error('logger down');
    });

    const inserted = table.insert({ name: 'Ann' });

    expect(inserted.name).toBe('Ann');
    expect(mocks.console.error).toHaveBeenCalledWith(
      expect.stringContaining('[GasSheetDb] Logger failed'),
      expect.anything(),
    );
    expect(table.find()).toHaveLength(1);
  });
});
