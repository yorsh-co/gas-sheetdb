/**
 * Entry point for database-like table access in sheets.
 */
class GasSheetDb {
  spreadsheet: GoogleAppsScript.Spreadsheet.Spreadsheet;
  rowNumbers: GasSheetDbRowsReference;
  lockScope: GasSheetDbLockScope;
  lockService: GasSheetDbLockService;
  lockTimeoutMs: number;
  logger: GasSheetDbLogger;

  constructor(options: GasSheetDbOptions = {}) {
    this.spreadsheet = this._resolveSpreadsheet(options);

    this.rowNumbers = this._resolveRowNumbers(
      _GAS_SHEETDB_DEFAULT_ROW_NUMBERS,
      options.rowNumbers,
    );

    this.lockScope = this._resolveLockScope(options.lockScope);
    this.lockService = this._resolveLockService(options.lockService);
    this.lockTimeoutMs = this._resolveLockTimeoutMs(options.lockTimeoutMs);
    this.logger = this._resolveLogger(options.logger);
  }

  /**
   * Resolve the spreadsheet from the provided source option.
   * Throws if more than one source is passed in options, or if
   * `useActiveSpreadsheet` is false and no other source is passed.
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

    if (useActiveSpreadsheet === false && selected.length === 0) {
      throw new Error(
        '[GasSheetDb] "useActiveSpreadsheet" is false, so a spreadsheet source is required. Pass "spreadsheet", "spreadsheetUrl" or "spreadsheetId".',
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
   * Merge a partial row configuration onto `base` and validate the result:
   * both rows must be positive integers, with the data after the header.
   * Shared by the constructor and table(), so both reject bad input alike.
   */
  private _resolveRowNumbers(
    base: GasSheetDbRowsReference,
    rowNumbers?: Partial<GasSheetDbRowsReference>,
  ): GasSheetDbRowsReference {
    const merged = { ...base, ...rowNumbers };

    for (const key of ['columnKeys', 'firstData'] as const) {
      if (!Number.isInteger(merged[key]) || merged[key] < 1) {
        throw new Error(
          `[GasSheetDb] Invalid "rowNumbers.${key}": ${String(merged[key])}. Expected a positive integer.`,
        );
      }
    }

    if (merged.firstData <= merged.columnKeys) {
      throw new Error(
        `[GasSheetDb] Invalid "rowNumbers": "firstData" (${merged.firstData}) must be greater than "columnKeys" (${merged.columnKeys}).`,
      );
    }

    return merged;
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
   * Validate a lockService, falling back to the library default when omitted.
   */
  private _resolveLockService(
    lockService?: GasSheetDbLockService,
  ): GasSheetDbLockService {
    if (lockService === undefined) return _GAS_SHEETDB_DEFAULT_LOCK_SERVICE;

    if (!lockService || typeof lockService.withLock !== 'function') {
      throw new Error(
        '[GasSheetDb] Invalid "lockService": expected an object exposing withLock(scope, callback, options?) — e.g. GasLock. Passing the LockService global from Apps Script rather than GasLock (or a withLock wrapper) is the usual cause.',
      );
    }

    return lockService;
  }

  /**
   * Validate a lock timeout in milliseconds, falling back to the library
   * default when omitted. Shared by the constructor and table(), so both
   * reject bad input alike.
   */
  private _resolveLockTimeoutMs(lockTimeoutMs?: number): number {
    if (lockTimeoutMs === undefined)
      return _GAS_SHEETDB_DEFAULT_LOCK_TIMEOUT_MS;

    if (
      typeof lockTimeoutMs !== 'number' ||
      !Number.isFinite(lockTimeoutMs) ||
      lockTimeoutMs <= 0
    ) {
      throw new Error(
        `[GasSheetDb] Invalid "lockTimeoutMs": ${String(lockTimeoutMs)}. Expected a positive, finite number of milliseconds.`,
      );
    }

    return lockTimeoutMs;
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
   * Create a table wrapper for a sheet, created if it does not exist.
   * Throws if `sheetName` is not a non-empty string.
   */
  table(options: GasSheetDbTableOptions): _GasSheetDbTable {
    const {
      sheetName,
      rowNumbers = {},
      lockScope,
      lockTimeoutMs,
      logger,
    } = options;

    if (typeof sheetName !== 'string' || !sheetName) {
      throw new Error(
        `[GasSheetDb] Invalid "sheetName": ${String(sheetName)}. Expected a non-empty string.`,
      );
    }

    return new _GasSheetDbTable({
      spreadsheet: this.spreadsheet,
      sheetName,
      rowNumbers: this._resolveRowNumbers(this.rowNumbers, rowNumbers),
      lockScope:
        lockScope === undefined
          ? this.lockScope
          : this._resolveLockScope(lockScope),
      lockService: this.lockService,
      lockTimeoutMs:
        lockTimeoutMs === undefined
          ? this.lockTimeoutMs
          : this._resolveLockTimeoutMs(lockTimeoutMs),
      logger: logger === undefined ? this.logger : this._resolveLogger(logger),
    });
  }
}
