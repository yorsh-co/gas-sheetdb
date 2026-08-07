/** Row/column position config for a table's header and data rows. */
interface GasSheetDbRowsReference {
  columnKeys: number;
  firstData: number;
}
/** A stored entry: persisted system metadata.  */
interface GasSheetDbSystemFields {
  _id?: string;
  _createdAt?: Date;
  _updatedAt?: Date;
  _isDeleted?: boolean;
}
/** A stored entry: arbitrary user properties plus persisted system metadata. */
type GasSheetDbEntry = GasSheetDbSystemFields & Record<string, unknown>;
/** A single raw sheet cell value, as returned by Range#getValues(). */
type GasSheetDbCellValue = string | number | boolean | Date;
/** Constructor options for `GasSheetDb`. */
interface GasSheetDbOptions {
  spreadsheet?: GoogleAppsScript.Spreadsheet.Spreadsheet;
  spreadsheetUrl?: string;
  spreadsheetId?: string;
  useActiveSpreadsheet?: boolean;
  rowNumbers?: Partial<GasSheetDbRowsReference>;
  /** @default 'script' */
  lockScope?: GasSheetDbLockScope;
  /** @default an internal fallback calling LockService directly, with no reentrancy protection */
  lockService?: GasSheetDbLockService;
  /** @default an internal fallback writing to the execution log via console */
  logger?: GasSheetDbLogger;
}
/** Options for `GasSheetDb.table()`. */
interface GasSheetDbTableOptions {
  sheetName?: string | null;
  rowNumbers?: Partial<GasSheetDbRowsReference>;
  /** Overrides the parent `GasSheetDb` instance's lockScope for this table only. */
  lockScope?: GasSheetDbLockScope;
  /** Overrides the parent `GasSheetDb` instance's logger for this table only. */
  logger?: GasSheetDbLogger;
}
/** Constructor options for the internal `_GasSheetDbTable` class. */
interface GasSheetDbTableConstructorOptions {
  spreadsheet: GoogleAppsScript.Spreadsheet.Spreadsheet;
  sheetName?: string | null;
  rowNumbers: GasSheetDbRowsReference;
  lockScope: GasSheetDbLockScope;
  lockService: GasSheetDbLockService;
  logger: GasSheetDbLogger;
}
/** Matches entries in the predicate-based query, update and delete methods. */
type GasSheetDbPredicate = (entry: GasSheetDbEntry) => boolean;
/**
 * The change `updateWhere()`/`updateOneWhere()` applies to each matched
 * entry: either a patch object applied to every match, or a function
 * returning a patch for a given entry. A function that mutates its entry in
 * place and returns nothing is honoured too.
 */
type GasSheetDbUpdate =
  GasSheetDbEntry | ((entry: GasSheetDbEntry) => GasSheetDbEntry | void);
/** Options for `_GasSheetDbTable#find()`. */
interface GasSheetDbFindOptions {
  withTrashed?: boolean;
  onlyTrashed?: boolean;
}
/**
 * Minimal lock-service shape gas-sheetdb needs to guard its writes.
 * GasLock (github.com/yorsh-co/gas-lock) satisfies this directly — pass
 * it in to get reentrancy-safe nesting with your own LockService calls.
 * Omit it and gas-sheetdb falls back to acquiring Apps Script's
 * LockService directly, with no reentrancy protection.
 */
interface GasSheetDbLockService {
  withLock<T>(
    scope: GasSheetDbLockScope,
    callback: () => T,
    options?: {
      timeoutMs?: number;
    },
  ): T;
}
type GasSheetDbLockScope = 'script' | 'document' | 'user';
/**
 * Minimal logging surface gas-sheetdb needs. GasLogger
 * (github.com/yorsh-co/gas-logger) satisfies it directly — pass an instance,
 * or a `child()` of one, and gas-sheetdb's log lines inherit whatever sinks,
 * levels and bindings that logger is configured with. Plain `console`
 * satisfies it too. Omit it and gas-sheetdb writes to the execution log.
 *
 * Only `info` and `warn` are declared: every failure gas-sheetdb cannot
 * recover from is thrown rather than logged, so there is nothing to report
 * at `error`.
 *
 * Flushing a buffered logger is the caller's responsibility — gas-sheetdb
 * owns no execution boundary to flush on.
 */
interface GasSheetDbLogger {
  info(msg: string, meta?: Record<string, unknown>): void;
  warn(msg: string, meta?: Record<string, unknown>): void;
}
