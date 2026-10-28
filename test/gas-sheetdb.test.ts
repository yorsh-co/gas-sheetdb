import { describe, expect, it, vi } from 'vitest';
import { FakeSpreadsheet } from './mocks';
import {
  asSpreadsheet,
  createLockService,
  createLogger,
  header,
  setup,
  setupTable,
  sheetOf,
} from './helpers';

describe('constructor: spreadsheet source', () => {
  it('uses the spreadsheet it is given', () => {
    const { GasSheetDb, mocks } = setup();
    mocks.SpreadsheetApp.getActiveSpreadsheet.mockClear();
    const spreadsheet = new FakeSpreadsheet();

    const db = new GasSheetDb({ spreadsheet: asSpreadsheet(spreadsheet) });

    expect(db.spreadsheet).toBe(spreadsheet);
    expect(mocks.SpreadsheetApp.getActiveSpreadsheet).not.toHaveBeenCalled();
  });

  it('opens by url', () => {
    const { GasSheetDb, mocks } = setup();

    const db = new GasSheetDb({ spreadsheetUrl: 'https://example.test/s' });

    expect(mocks.SpreadsheetApp.openByUrl).toHaveBeenCalledExactlyOnceWith(
      'https://example.test/s',
    );
    expect(db.spreadsheet).toBe(
      mocks.SpreadsheetApp.openByUrl.mock.results[0]?.value,
    );
  });

  it('opens by id', () => {
    const { GasSheetDb, mocks } = setup();

    const db = new GasSheetDb({ spreadsheetId: 'abc' });

    expect(mocks.SpreadsheetApp.openById).toHaveBeenCalledExactlyOnceWith(
      'abc',
    );
    expect(db.spreadsheet).toBe(
      mocks.SpreadsheetApp.openById.mock.results[0]?.value,
    );
  });

  it.each([{ useActiveSpreadsheet: true }, {}])(
    'uses the active spreadsheet for %j',
    (options) => {
      const { GasSheetDb, mocks } = setup();

      const db = new GasSheetDb(options);

      expect(db.spreadsheet).toBe(mocks.activeSpreadsheet);
    },
  );

  it('throws when there is no active spreadsheet', () => {
    const { GasSheetDb, mocks } = setup();
    mocks.SpreadsheetApp.getActiveSpreadsheet.mockReturnValue(null);

    expect(() => new GasSheetDb()).toThrow(
      '[GasSheetDb] No active spreadsheet is accessible by this script.',
    );
  });

  it('throws naming every source when more than one is passed', () => {
    const { GasSheetDb } = setup();

    expect(
      () =>
        new GasSheetDb({
          spreadsheetUrl: 'https://example.test/s',
          spreadsheetId: 'abc',
          useActiveSpreadsheet: true,
        }),
    ).toThrow(
      'Received more than one spreadsheet source option: spreadsheetUrl, spreadsheetId, useActiveSpreadsheet.',
    );
  });
});

describe('constructor: rowNumbers', () => {
  it('defaults to a header on row 1 and data from row 2', () => {
    const { db } = setup();

    expect(db.rowNumbers).toEqual({ columnKeys: 1, firstData: 2 });
  });

  it('merges a partial override onto the defaults', () => {
    const { db } = setup({ rowNumbers: { firstData: 4 } });

    expect(db.rowNumbers).toEqual({ columnKeys: 1, firstData: 4 });
  });

  it('merges a table override onto the instance values', () => {
    const { db } = setup({ rowNumbers: { firstData: 4 } });

    const table = db.table({ sheetName: 't', rowNumbers: { firstData: 5 } });

    expect(table.rowNumbers).toEqual({ columnKeys: 1, firstData: 5 });
    expect(table.rowIndexes).toEqual({ columnKeys: 0, firstData: 4 });
    expect(db.rowNumbers).toEqual({ columnKeys: 1, firstData: 4 });
  });
});

describe('constructor: lockScope', () => {
  it('defaults to script', () => {
    const { db } = setup();

    expect(db.lockScope).toBe('script');
  });

  it('rejects an unknown scope', () => {
    const { GasSheetDb } = setup();

    expect(
      () => new GasSheetDb({ lockScope: 'global' as GasSheetDbLockScope }),
    ).toThrow(
      '[GasSheetDb] Invalid "lockScope": "global". Expected one of: script, document, user.',
    );
  });

  it('rejects an unknown scope on table()', () => {
    const { db } = setup();

    expect(() =>
      db.table({ lockScope: 'global' as GasSheetDbLockScope }),
    ).toThrow('Invalid "lockScope": "global"');
  });

  it('lets a table override the instance scope', () => {
    const { db } = setup({ lockScope: 'user' });

    expect(db.table({ sheetName: 'a' }).lockScope).toBe('user');
    expect(db.table({ sheetName: 'b', lockScope: 'document' }).lockScope).toBe(
      'document',
    );
  });
});

describe('constructor: logger', () => {
  it.each([
    ['an empty object', {}],
    ['null', null],
    ['an object without warn', { info: () => undefined }],
    ['a class rather than an instance', class Logger {}],
  ])('rejects %s', (_label, logger) => {
    const { GasSheetDb } = setup();

    expect(
      () => new GasSheetDb({ logger: logger as unknown as GasSheetDbLogger }),
    ).toThrow('[GasSheetDb] Invalid "logger"');
  });

  it('rejects an invalid logger on table()', () => {
    const { db } = setup();

    expect(() =>
      db.table({ logger: {} as unknown as GasSheetDbLogger }),
    ).toThrow('Invalid "logger"');
  });

  it('accepts console', () => {
    const { GasSheetDb, mocks } = setup();

    expect(
      () =>
        new GasSheetDb({ logger: mocks.console, useActiveSpreadsheet: true }),
    ).not.toThrow();
  });

  it('falls back to the execution log with the [GasSheetDb] prefix', () => {
    const { db, mocks } = setup();

    db.table({ sheetName: 'people' });

    expect(mocks.console.log).toHaveBeenCalledExactlyOnceWith(
      '[GasSheetDb] Created table sheet',
      { sheet: 'people' },
    );
  });

  it('sends log lines to an injected child-logger-like object', () => {
    const child = { ...createLogger(), child: vi.fn() };
    const { db, mocks } = setup({ logger: child });

    db.table({ sheetName: 'people' }).insert({ name: 'Ann' });

    expect(child.info).toHaveBeenCalledWith('Created table sheet', {
      sheet: 'people',
    });
    expect(child.info).toHaveBeenCalledWith('Added columns', {
      sheet: 'people',
      columns: expect.arrayContaining(['_id', 'name']),
    });
    expect(mocks.console.log).not.toHaveBeenCalled();
  });

  it('lets a table logger take precedence over the instance logger', () => {
    const instanceLogger = createLogger();
    const tableLogger = createLogger();
    const { db } = setup({ logger: instanceLogger });

    db.table({ sheetName: 'a', logger: tableLogger });
    db.table({ sheetName: 'b' });

    expect(tableLogger.info).toHaveBeenCalledExactlyOnceWith(
      'Created table sheet',
      { sheet: 'a' },
    );
    expect(instanceLogger.info).toHaveBeenCalledExactlyOnceWith(
      'Created table sheet',
      { sheet: 'b' },
    );
  });
});

describe('table()', () => {
  it('creates a named sheet with banding, a bold frozen header and a log line', () => {
    const { db, mocks } = setup({ logger: createLogger() });

    const table = db.table({ sheetName: 'people' });
    const sheet = sheetOf(mocks.activeSpreadsheet, 'people');

    expect(table.sheet).toBe(sheet);
    expect(table.sheetName).toBe('people');
    expect(sheet.getFrozenRows()).toBe(1);
    expect(sheet.calls.applyRowBanding).toEqual([
      { row: 1, col: 1, numRows: 1000, numCols: 26 },
    ]);
    expect(sheet.calls.setFontWeight).toEqual([
      { row: 1, col: 1, numRows: 1, numCols: 26 },
    ]);
  });

  it('reuses an existing sheet and reads its header', () => {
    const { table, sheet, db, logger } = setupTable();
    table.insert({ name: 'Ann' });
    logger.info.mockClear();

    const again = db.table({ sheetName: 'people' });

    expect(again.sheet).toBe(sheet);
    expect(again.schema.columnKeys).toEqual(header(sheet));
    expect(logger.info).not.toHaveBeenCalled();
  });

  it('names unnamed tables gsd-table-N', () => {
    const { db, mocks } = setup({ logger: createLogger() });

    const names = [db.table(), db.table()].map((t) => t.sheetName);

    expect(names).toEqual(['gsd-table-0', 'gsd-table-1']);
    expect(mocks.activeSpreadsheet.getSheets()).toHaveLength(2);
  });

  it('skips a default name that is already taken', () => {
    const { db, mocks } = setup({ logger: createLogger() });
    mocks.activeSpreadsheet.insertSheet('gsd-table-0');
    mocks.activeSpreadsheet.insertSheet('gsd-table-1');

    expect(db.table().sheetName).toBe('gsd-table-2');
  });

  it.todo('creates a sheet when rowNumbers.columnKeys is not 1');
});

describe('locking', () => {
  it.each([
    ['script', 'getScriptLock'],
    ['document', 'getDocumentLock'],
    ['user', 'getUserLock'],
  ] as const)('acquires the %s lock with a 30000ms wait', (scope, getter) => {
    const { table, mocks } = setupTable('people', { lockScope: scope });

    table.find();

    const lock = mocks.LockService[getter].mock.results[0]?.value;
    expect(lock.waitLock).toHaveBeenCalledExactlyOnceWith(30000);
    expect(lock.releaseLock).toHaveBeenCalledOnce();
  });

  it('throws when the document lock is unavailable', () => {
    const { table } = setupTable(
      'people',
      { lockScope: 'document' },
      { unavailable: ['document'] },
    );

    expect(() => table.find()).toThrow(
      'LockService returned no lock for scope "document"',
    );
  });

  it('releases the lock when the operation throws', () => {
    const { table, mocks } = setupTable();
    table.insert({ name: 'Ann' });

    expect(() => table.update({ _id: 'missing' })).toThrow(
      'Entry not found: missing',
    );

    expect(mocks.isLocked()).toBe(false);
  });

  it('passes the table scope to an injected lockService', () => {
    const { lockService, withLock } = createLockService();
    const { table, mocks } = setupTable('people', {
      lockScope: 'document',
      lockService,
    });

    table.insert({ name: 'Ann' });

    expect(withLock).toHaveBeenCalledExactlyOnceWith(
      'document',
      expect.any(Function),
    );
    expect(mocks.LockService.getDocumentLock).not.toHaveBeenCalled();
  });

  it('flushes before and after the callback, inside the lock', () => {
    const events: string[] = [];
    const lockService = {
      withLock: <T>(_scope: string, callback: () => T): T => {
        events.push('enter');
        const result = callback();
        events.push('exit');
        return result;
      },
    };
    const { table, mocks } = setupTable('people', { lockService });
    mocks.SpreadsheetApp.flush.mockImplementation(() => {
      events.push('flush');
    });

    table.find();

    expect(events).toEqual(['enter', 'flush', 'flush', 'exit']);
  });

  it('flushes inside the default lock', () => {
    const { table, mocks } = setupTable();

    table.insert({ name: 'Ann' });

    expect(mocks.flushLockStates).toEqual([true, true]);
  });
});
