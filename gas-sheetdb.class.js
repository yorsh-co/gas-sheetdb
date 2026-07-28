'use strict';
const _GAS_SHEET_DB_DEFAULT_ROW_NUMBERS = Object.freeze({
  columnKeys: 1,
  firstData: 2,
});
/**
 * Entry point for database-like table access in sheets.
 */
class GasSheetDb {
  constructor(options = {}) {
    this.spreadsheet = this._resolveSpreadsheet(options);
    this.rowNumbers = {
      ..._GAS_SHEET_DB_DEFAULT_ROW_NUMBERS,
      ...options.rowNumbers,
    };
    this.lockScope = this._resolveLockScope(options.lockScope);
    this.lockService = options.lockService ?? _GAS_SHEETDB_DEFAULT_LOCK_SERVICE;
  }
  /**
     * Resolve the spreadsheet from the provided source option.
     * Throws if more than on source is passed in options.
  
     */
  _resolveSpreadsheet(options) {
    const { spreadsheet, spreadsheetUrl, spreadsheetId, useActiveSpreadsheet } =
      options;
    // filter arguments
    const sources = {
      spreadsheet: !!spreadsheet,
      spreadsheetUrl: !!spreadsheetUrl,
      spreadsheetId: !!spreadsheetId,
      useActiveSpreadsheet: !!useActiveSpreadsheet,
    };
    const selected = Object.keys(sources).filter((k) => sources[k]);
    if (selected.length > 1) {
      throw new Error(
        `[GasSheetDb] Received more than one spreadsheet source option: ${selected.join(', ')}. Pass exactly one of "spreadsheet", "spreadsheetUrl", "spreadsheetId" or "useActiveSpreadsheet".`,
      );
    }
    // return spreadsheet by the indicated method
    if (spreadsheet) return spreadsheet;
    if (spreadsheetUrl) return SpreadsheetApp.openByUrl(spreadsheetUrl);
    if (spreadsheetId) return SpreadsheetApp.openById(spreadsheetId);
    // default: return active spreadsheet
    const activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    if (!activeSpreadsheet) {
      throw new Error(
        '[GasSheetDb] No active spreadsheet is accessible by this script. Pass "spreadsheet", "spreadsheetUrl" or "spreadsheetId" instead.',
      );
    }
    return activeSpreadsheet;
  }
  /**
   * Validate a lockScope, falling back to the library default when omitted.
   * Shared by the constructor and table(), so both reject bad input alike.
   */
  _resolveLockScope(lockScope) {
    if (lockScope === undefined) return _GAS_SHEETDB_DEFAULT_LOCK_SCOPE;
    if (!_GAS_SHEETDB_LOCK_SCOPES.includes(lockScope)) {
      throw new Error(
        `[GasSheetDb] Invalid "lockScope": "${lockScope}". Expected one of: ${_GAS_SHEETDB_LOCK_SCOPES.join(', ')}.`,
      );
    }
    return lockScope;
  }
  /**
   * Create a table wrapper for a sheet.
   */
  table(options = {}) {
    const { sheetName = null, rowNumbers = {}, lockScope } = options;
    return new _GasSheetDbTable({
      spreadsheet: this.spreadsheet,
      sheetName,
      rowNumbers: {
        ...this.rowNumbers,
        ...rowNumbers,
      },
      lockScope:
        lockScope === undefined
          ? this.lockScope
          : this._resolveLockScope(lockScope),
      lockService: this.lockService,
    });
  }
}
