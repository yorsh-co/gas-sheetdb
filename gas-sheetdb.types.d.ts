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
}
/** Options for `GasSheetDb.table()`. */
interface GasSheetDbTableOptions {
  sheetName?: string | null;
  rowNumbers?: Partial<GasSheetDbRowsReference>;
  /** Overrides the parent `GasSheetDb` instance's lockScope for this table only. */
  lockScope?: GasSheetDbLockScope;
}
/** Constructor options for the internal `_GasSheetDbTable` class. */
interface GasSheetDbTableConstructorOptions {
  spreadsheet: GoogleAppsScript.Spreadsheet.Spreadsheet;
  sheetName?: string | null;
  rowNumbers: GasSheetDbRowsReference;
  lockScope: GasSheetDbLockScope;
  lockService: GasSheetDbLockService;
}
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
