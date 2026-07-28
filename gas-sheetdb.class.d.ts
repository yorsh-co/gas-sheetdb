declare const _GAS_SHEET_DB_DEFAULT_ROW_NUMBERS: GasSheetDbRowsReference;
/**
 * Entry point for database-like table access in sheets.
 */
declare class GasSheetDb {
  spreadsheet: GoogleAppsScript.Spreadsheet.Spreadsheet;
  rowNumbers: GasSheetDbRowsReference;
  lockScope: GasSheetDbLockScope;
  lockService: GasSheetDbLockService;
  constructor(options?: GasSheetDbOptions);
  /**
     * Resolve the spreadsheet from the provided source option.
     * Throws if more than on source is passed in options.
  
     */
  private _resolveSpreadsheet;
  /**
   * Validate a lockScope, falling back to the library default when omitted.
   * Shared by the constructor and table(), so both reject bad input alike.
   */
  private _resolveLockScope;
  /**
   * Create a table wrapper for a sheet.
   */
  table(options?: GasSheetDbTableOptions): _GasSheetDbTable;
}
