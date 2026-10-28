/**
 * Entry point for database-like table access in sheets.
 */
declare class GasSheetDb {
  spreadsheet: GoogleAppsScript.Spreadsheet.Spreadsheet;
  rowNumbers: GasSheetDbRowsReference;
  lockScope: GasSheetDbLockScope;
  lockService: GasSheetDbLockService;
  lockTimeoutMs: number;
  logger: GasSheetDbLogger;
  constructor(options?: GasSheetDbOptions);
  /**
   * Resolve the spreadsheet from the provided source option.
   * Throws if more than one source is passed in options, or if
   * `useActiveSpreadsheet` is false and no other source is passed.
   */
  private _resolveSpreadsheet;
  /**
   * Merge a partial row configuration onto `base` and validate the result:
   * both rows must be positive integers, with the data after the header.
   * Shared by the constructor and table(), so both reject bad input alike.
   */
  private _resolveRowNumbers;
  /**
   * Validate a lockScope, falling back to the library default when omitted.
   * Shared by the constructor and table(), so both reject bad input alike.
   */
  private _resolveLockScope;
  /**
   * Validate a lockService, falling back to the library default when omitted.
   */
  private _resolveLockService;
  /**
   * Validate a lock timeout in milliseconds, falling back to the library
   * default when omitted. Shared by the constructor and table(), so both
   * reject bad input alike.
   */
  private _resolveLockTimeoutMs;
  /**
   * Validate a logger, falling back to the library default when omitted.
   * Shared by the constructor and table(), so both reject bad input alike.
   */
  private _resolveLogger;
  /**
   * Create a table wrapper for a sheet, created if it does not exist.
   * Throws if `sheetName` is not a non-empty string.
   */
  table(options: GasSheetDbTableOptions): _GasSheetDbTable;
}
