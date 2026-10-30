import { mkdtempSync, rmSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { createGasLoader } from 'gas-dev-config/test';
import { describe, expect, it, vi } from 'vitest';
import { createMocks, FakeSpreadsheet } from './mocks';
import {
  asSpreadsheet,
  createLockService,
  createLogger,
  header,
  setup,
  setupProbed,
  setupTable,
  sheetOf,
} from './helpers';
import { loadGasSheetDb, ROOT, SOURCE_FILES } from './harness';

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

  it('throws when useActiveSpreadsheet is false and no other source is passed', () => {
    const { GasSheetDb, mocks } = setup();
    mocks.SpreadsheetApp.getActiveSpreadsheet.mockClear();

    expect(() => new GasSheetDb({ useActiveSpreadsheet: false })).toThrow(
      '[GasSheetDb] "useActiveSpreadsheet" is false, so a spreadsheet source is required.',
    );
    expect(mocks.SpreadsheetApp.getActiveSpreadsheet).not.toHaveBeenCalled();
  });

  it('accepts useActiveSpreadsheet: false alongside another source', () => {
    const { GasSheetDb, mocks } = setup();

    const db = new GasSheetDb({
      useActiveSpreadsheet: false,
      spreadsheetId: 'abc',
    });

    expect(db.spreadsheet).toBe(
      mocks.SpreadsheetApp.openById.mock.results[0]?.value,
    );
  });

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
      db.table({
        sheetName: 'people',
        lockScope: 'global' as GasSheetDbLockScope,
      }),
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

describe('constructor: lockTimeoutMs', () => {
  it('defaults to 30000', () => {
    const { db } = setup();

    expect(db.lockTimeoutMs).toBe(30000);
    expect(db.table({ sheetName: 'a' }).lockTimeoutMs).toBe(30000);
  });

  it.each([
    ['zero', 0],
    ['negative', -1],
    ['NaN', Number.NaN],
    ['Infinity', Number.POSITIVE_INFINITY],
    ['a string', '5000'],
    ['null', null],
  ])('rejects %s', (_label, value) => {
    const { GasSheetDb } = setup();

    expect(
      () => new GasSheetDb({ lockTimeoutMs: value as unknown as number }),
    ).toThrow('[GasSheetDb] Invalid "lockTimeoutMs"');
  });

  it('rejects an invalid value on table()', () => {
    const { db } = setup();

    expect(() => db.table({ sheetName: 'people', lockTimeoutMs: -5 })).toThrow(
      '[GasSheetDb] Invalid "lockTimeoutMs": -5',
    );
  });

  it('lets a table override the instance value', () => {
    const { db } = setup({ lockTimeoutMs: 5000 });

    expect(db.table({ sheetName: 'a' }).lockTimeoutMs).toBe(5000);
    expect(db.table({ sheetName: 'b', lockTimeoutMs: 100 }).lockTimeoutMs).toBe(
      100,
    );
  });
});

describe('constructor: lockService', () => {
  it.each([
    ['null', null],
    ['an empty object', {}],
    ['a string', 'GasLock'],
    [
      'the LockService global',
      { getScriptLock: () => null, getDocumentLock: () => null },
    ],
  ])('rejects %s', (_label, lockService) => {
    const { GasSheetDb } = setup();

    expect(
      () =>
        new GasSheetDb({
          lockService: lockService as unknown as GasSheetDbLockService,
        }),
    ).toThrow(/^\[GasSheetDb\] Invalid "lockService".*LockService/);
  });

  it('accepts an object exposing withLock', () => {
    const { GasSheetDb } = setup();
    const { lockService } = createLockService();

    expect(() => new GasSheetDb({ lockService })).not.toThrow();
  });
});

describe('default lock service', () => {
  it('uses the GasLock global when present, with scope and timeout', () => {
    const gasLock = createLockService();
    const { db, mocks } = setupProbed(
      { lockScope: 'user', lockTimeoutMs: 1234 },
      { GasLock: gasLock.lockService },
    );
    const table = db.table({ sheetName: 'people' });
    gasLock.withLock.mockClear();

    table.insert({ name: 'Ann' });

    expect(gasLock.withLock).toHaveBeenCalledExactlyOnceWith(
      'user',
      expect.any(Function),
      { timeoutMs: 1234 },
    );
    expect(mocks.LockService.getUserLock).not.toHaveBeenCalled();
  });

  it('uses the fallback when GasLock is absent', () => {
    const { table, mocks } = setupTable('people', { lockTimeoutMs: 777 });

    table.insert({ name: 'Ann' });

    const lock = mocks.LockService.getScriptLock.mock.results.at(-1)?.value;
    expect(lock.waitLock).toHaveBeenCalledWith(777);
  });

  it('resolves GasLock at each acquisition, not at construction', () => {
    const { GasSheetDb, mocks } = loadGasSheetDb();
    const db = new GasSheetDb({ useActiveSpreadsheet: true });
    const table = db.table({ sheetName: 'people' });
    const gasLock = createLockService();
    // the library's own realm, where a later-evaluated gas-lock would define it
    const context = (GasSheetDb.constructor as FunctionConstructor)(
      'return globalThis',
    )() as Record<string, unknown>;
    mocks.LockService.getScriptLock.mockClear();

    context.GasLock = gasLock.lockService;
    table.insert({ name: 'Ann' });

    expect(gasLock.withLock).toHaveBeenCalledOnce();
    expect(mocks.LockService.getScriptLock).not.toHaveBeenCalled();
  });

  it('falls back when GasLock is not initialised yet', () => {
    const dir = mkdtempSync(join(tmpdir(), 'gas-sheetdb-'));
    const trailing = join(dir, 'late-gas-lock.ts');
    // GasLock is declared after the library already used it, as when
    // gas-lock's file is evaluated later than a top-level GasSheetDb: reading
    // the binding then throws a ReferenceError.
    writeFileSync(
      trailing,
      `const earlyDb = new GasSheetDb({ useActiveSpreadsheet: true });
       earlyDb.table({ sheetName: 'early' }).insert({ name: 'Ann' });
       const GasLock = { withLock: () => { throw new Error('unexpected'); } };`,
    );
    try {
      const mocks = createMocks();
      const load = createGasLoader<{ GasSheetDb: typeof GasSheetDb }>({
        files: [...SOURCE_FILES, trailing],
        exports: ['GasSheetDb'],
        root: ROOT,
      });

      expect(() =>
        load({
          SpreadsheetApp: mocks.SpreadsheetApp,
          LockService: mocks.LockService,
          Utilities: mocks.Utilities,
          console: mocks.console,
          Error,
          Date,
        }),
      ).not.toThrow();
      expect(mocks.LockService.getScriptLock).toHaveBeenCalled();
    } finally {
      rmSync(dir, { recursive: true, force: true });
    }
  });
});

describe('sheet creation race', () => {
  it('reuses a sheet another execution created before the lock was taken', () => {
    const { db, mocks } = setup({ logger: createLogger() });
    const existing = mocks.activeSpreadsheet.insertSheet('people');
    mocks.activeSpreadsheet.insertSheet.mockClear();
    mocks.activeSpreadsheet.getSheetByName.mockImplementationOnce(() => null);

    const table = db.table({ sheetName: 'people' });

    expect(table.sheet).toBe(existing);
    expect(mocks.activeSpreadsheet.insertSheet).not.toHaveBeenCalled();
  });

  it('does not log a creation it did not perform', () => {
    const logger = createLogger();
    const { db, mocks } = setup({ logger });
    mocks.activeSpreadsheet.insertSheet('people');
    mocks.activeSpreadsheet.getSheetByName.mockImplementationOnce(() => null);

    db.table({ sheetName: 'people' });

    expect(logger.info).not.toHaveBeenCalled();
  });

  it('creates the missing sheet while holding the lock', () => {
    const { db, mocks } = setup();
    let lockedOnInsert = false;
    const insert = mocks.activeSpreadsheet.insertSheet.getMockImplementation()!;
    mocks.activeSpreadsheet.insertSheet.mockImplementation((name: string) => {
      lockedOnInsert = mocks.isLocked();
      return insert(name);
    });

    db.table({ sheetName: 'people' });

    expect(lockedOnInsert).toBe(true);
  });

  it('does not take the lock when the sheet exists', () => {
    const { db, mocks } = setup();
    mocks.activeSpreadsheet.insertSheet('people');

    db.table({ sheetName: 'people' });

    expect(mocks.LockService.getScriptLock).not.toHaveBeenCalled();
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
      db.table({
        sheetName: 'people',
        logger: {} as unknown as GasSheetDbLogger,
      }),
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

  it.each([undefined, '', null, 7])(
    'rejects a missing or non-string sheetName (%j) without creating a sheet',
    (sheetName) => {
      const { db, mocks } = setup({ logger: createLogger() });
      const sheetsBefore = mocks.activeSpreadsheet.getSheets().length;

      expect(() =>
        db.table({ sheetName: sheetName as unknown as string }),
      ).toThrow('[GasSheetDb] Invalid "sheetName"');
      expect(mocks.activeSpreadsheet.getSheets()).toHaveLength(sheetsBefore);
    },
  );
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
    const { db, mocks } = setup(
      { lockScope: 'document' },
      { unavailable: ['document'] },
    );
    mocks.activeSpreadsheet.insertSheet('people');
    const table = db.table({ sheetName: 'people' });

    expect(() => table.find()).toThrow(
      'LockService returned no lock for scope "document"',
    );
  });

  it('throws on creating a sheet when the document lock is unavailable', () => {
    const { db } = setup(
      { lockScope: 'document' },
      { unavailable: ['document'] },
    );

    expect(() => db.table({ sheetName: 'people' })).toThrow(
      'LockService returned no lock for scope "document"',
    );
  });

  it('releases the lock when the operation throws', () => {
    const { table, mocks } = setupTable();
    table.insert({ name: 'Ann' });

    expect(() => table.update({ _id: 'missing' })).toThrow(
      '[GasSheetDb] Entry not found: missing',
    );

    expect(mocks.isLocked()).toBe(false);
  });

  it('passes the table scope to an injected lockService', () => {
    const { lockService, withLock } = createLockService();
    const { table, mocks } = setupTable('people', {
      lockScope: 'document',
      lockService,
    });

    withLock.mockClear();

    table.insert({ name: 'Ann' });

    expect(withLock).toHaveBeenCalledExactlyOnceWith(
      'document',
      expect.any(Function),
      { timeoutMs: 30000 },
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
    events.length = 0;
    mocks.SpreadsheetApp.flush.mockImplementation(() => {
      events.push('flush');
    });

    table.find();

    expect(events).toEqual(['enter', 'flush', 'flush', 'exit']);
  });

  it('flushes inside the default lock', () => {
    const { table, mocks } = setupTable();
    mocks.flushLockStates.length = 0;

    table.insert({ name: 'Ann' });

    expect(mocks.flushLockStates).toEqual([true, true]);
  });
});
