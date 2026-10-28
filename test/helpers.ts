import { vi } from 'vitest';
import { loadGasSheetDb } from './harness';
import type { FakeSheet, FakeSpreadsheet, MockOptions } from './mocks';

type Row = Record<string, unknown>;

/** Spy logger satisfying `GasSheetDbLogger`. */
export function createLogger() {
  return { info: vi.fn(), warn: vi.fn() };
}

/**
 * Pass-through lock service that records its calls. `lockService` is what to
 * inject; `withLock` is the spy to assert on.
 */
export function createLockService() {
  const withLock = vi.fn(
    (scope: GasSheetDbLockScope, callback: () => unknown): unknown => {
      void scope;
      return callback();
    },
  );
  return { withLock, lockService: { withLock } as GasSheetDbLockService };
}

/** Loads the library and a `GasSheetDb` on the active spreadsheet. */
export function setup(
  options: GasSheetDbOptions = {},
  mockOptions: MockOptions = {},
) {
  const { GasSheetDb, mocks } = loadGasSheetDb(mockOptions);
  const db = new GasSheetDb({ useActiveSpreadsheet: true, ...options });
  return { GasSheetDb, mocks, db };
}

/**
 * Loads the library and a `GasSheetDb` whose spy logger records, at each
 * call, whether the fake `LockService` has a lock held (`lockedAtLog`).
 * `globals` adds context globals, such as a fake `GasLock`.
 */
export function setupProbed(
  options: GasSheetDbOptions = {},
  globals: Record<string, unknown> = {},
) {
  const { GasSheetDb, mocks } = loadGasSheetDb({}, globals);
  const lockedAtLog: boolean[] = [];
  const record = (): void => {
    lockedAtLog.push(mocks.isLocked());
  };
  const logger = {
    info: vi.fn<GasSheetDbLogger['info']>(record),
    warn: vi.fn<GasSheetDbLogger['warn']>(record),
  };
  const db = new GasSheetDb({ useActiveSpreadsheet: true, logger, ...options });
  return { GasSheetDb, mocks, db, logger, lockedAtLog };
}

/** Loads the library and a named table, with a spy logger by default. */
export function setupTable(
  sheetName = 'people',
  options: GasSheetDbOptions = {},
  mockOptions: MockOptions = {},
) {
  const logger = createLogger();
  const { GasSheetDb, mocks, db } = setup({ logger, ...options }, mockOptions);
  const table = db.table({ sheetName });
  return {
    GasSheetDb,
    mocks,
    db,
    table,
    sheet: sheetOf(mocks.activeSpreadsheet, sheetName),
    logger,
  };
}

/** The sheet called `name`; fails the test when it does not exist. */
export function sheetOf(spreadsheet: FakeSpreadsheet, name: string): FakeSheet {
  const sheet = spreadsheet.getSheetByName(name);
  if (!sheet) throw new Error(`No sheet named "${name}"`);
  return sheet;
}

/** Casts a fake to the Apps Script type the library expects. */
export function asSpreadsheet(
  spreadsheet: FakeSpreadsheet,
): GoogleAppsScript.Spreadsheet.Spreadsheet {
  return spreadsheet as unknown as GoogleAppsScript.Spreadsheet.Spreadsheet;
}

/** The header row of a sheet. */
export function header(sheet: FakeSheet): string[] {
  return sheet
    .getRange(1, 1, 1, Math.max(sheet.getLastColumn(), 1))
    .getValues()[0] as string[];
}

/** The rows below the header, keyed by header name. */
export function rows(sheet: FakeSheet): Row[] {
  const [keys = [], ...body] = sheet.getDataRange().getValues();
  return body.map((values) =>
    Object.fromEntries(keys.map((key, i) => [String(key), values[i]])),
  );
}

/**
 * Writes `grid` from A1 as a user typing into the sheet would, bypassing the
 * library.
 */
export function enter(sheet: FakeSheet, grid: unknown[][]): void {
  sheet.getRange(1, 1, grid.length, grid[0]?.length ?? 1).setValues(grid);
}

/** Returns `items[index]`, failing the test when it is absent. */
export function at<T>(items: readonly T[], index: number): T {
  const item = items[index];
  if (item === undefined) throw new Error(`No item at index ${index}`);
  return item;
}
