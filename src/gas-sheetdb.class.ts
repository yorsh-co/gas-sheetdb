const _GAS_SHEET_DB_DEFAULT_ROW_NUMBERS: GasSheetDbRowsReference =
  Object.freeze({
    columnKeys: 1,
    firstData: 2,
  });

/**
 * Entry point for database-like table access in sheets.
 */
class GasSheetDb {
  spreadsheet: GoogleAppsScript.Spreadsheet.Spreadsheet;
  rowNumbers: GasSheetDbRowsReference;
  lockScope: GasSheetDbLockScope;
  lockService: GasSheetDbLockService;
  logger: GasSheetDbLogger;

  constructor(options: GasSheetDbOptions = {}) {
    this.spreadsheet = this._resolveSpreadsheet(options);

    this.rowNumbers = {
      ..._GAS_SHEET_DB_DEFAULT_ROW_NUMBERS,
      ...options.rowNumbers,
    };

    this.lockScope = this._resolveLockScope(options.lockScope);
    this.lockService = options.lockService ?? _GAS_SHEETDB_DEFAULT_LOCK_SERVICE;
    this.logger = this._resolveLogger(options.logger);
  }

  /**
   * Resolve the spreadsheet from the provided source option.
   * Throws if more than on source is passed in options.

   */
  private _resolveSpreadsheet(
    options: GasSheetDbOptions,
  ): GoogleAppsScript.Spreadsheet.Spreadsheet {
    const { spreadsheet, spreadsheetUrl, spreadsheetId, useActiveSpreadsheet } =
      options;

    // filter arguments
    const sources = {
      spreadsheet: !!spreadsheet,
      spreadsheetUrl: !!spreadsheetUrl,
      spreadsheetId: !!spreadsheetId,
      useActiveSpreadsheet: !!useActiveSpreadsheet,
    };

    const selected = Object.keys(sources).filter(
      (k) => sources[k as keyof typeof sources],
    );

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
  private _resolveLockScope(
    lockScope?: GasSheetDbLockScope,
  ): GasSheetDbLockScope {
    if (lockScope === undefined) return _GAS_SHEETDB_DEFAULT_LOCK_SCOPE;
    if (!_GAS_SHEETDB_LOCK_SCOPES.includes(lockScope)) {
      throw new Error(
        `[GasSheetDb] Invalid "lockScope": "${lockScope}". Expected one of: ${_GAS_SHEETDB_LOCK_SCOPES.join(', ')}.`,
      );
    }
    return lockScope;
  }

  /**
   * Validate a logger, falling back to the library default when omitted.
   * Shared by the constructor and table(), so both reject bad input alike.
   */
  private _resolveLogger(logger?: GasSheetDbLogger): GasSheetDbLogger {
    if (logger === undefined) return _GAS_SHEETDB_DEFAULT_LOGGER;

    if (
      !logger ||
      typeof logger.info !== 'function' ||
      typeof logger.warn !== 'function'
    ) {
      throw new Error(
        '[GasSheetDb] Invalid "logger": expected an object exposing info(msg, meta?) and warn(msg, meta?) — e.g. a GasLogger instance, or console. Passing the GasLogger class rather than an instance is the usual cause.',
      );
    }

    return logger;
  }

  /**
   * Create a table wrapper for a sheet.
   */
  table(options: GasSheetDbTableOptions = {}): _GasSheetDbTable {
    const { sheetName = null, rowNumbers = {}, lockScope, logger } = options;

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
      logger: logger === undefined ? this.logger : this._resolveLogger(logger),
    });
  }
}
