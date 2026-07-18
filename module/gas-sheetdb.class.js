const _GAS_SHEET_DB_DEFAULT_ROW_NUMBERS = Object.freeze({
  columnKeys: 1,
  firstData: 2,
});

/**
 * Entry point for database-like table access in sheets.
 */
class GasSheetDb {
  /**
   * @param {Object} [options]
   * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} [options.spreadsheet]
   * @param {string} [options.spreadsheetUrl]
   * @param {string} [options.spreadsheetId]
   *
   * @param {boolean} [options.useActiveSpreadsheet] // Default when not spreadsheet configuration is provided
   *
   * @param {_GasSheetDbRowsReference} [options.rowNumbers] - instance-wide default row config
   */
  constructor(options = {}) {
    this.spreadsheet = this._resolveSpreadsheet(options);

    this.rowNumbers = {
      ..._GAS_SHEET_DB_DEFAULT_ROW_NUMBERS,
      ...options.rowNumbers,
    };
  }

  /**
   * Resolve the spreadsheet from the provided source option.
   * Throws if more than on source is passed in options.
   *
   * @param {Object} options
   *
   * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} [options.spreadsheet]
   * @param {string} [options.spreadsheetUrl]
   * @param {string} [options.spreadsheetId]
   *
   * @param {boolean} [options.useActiveSpreadsheet]
   *
   * @returns {GoogleAppsScript.Spreadsheet.Spreadsheet}
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
   * Create a table wrapper for a sheet.
   *
   * @param {Object} options
   * @param {string} [options.sheetName = null] // If left blank, the sheet will automatically be named 'gsd-i'. `i` will be replaced by an index.
   * @param {_GasSheetDbRowsReference} [options.rowNumbers = {}] // Optionally set sheet-specific row numbers to overrule the row numbers defined in the GasSheetDb instance.
   *
   * @returns {_GasSheetDbTable}
   */
  table(options = {}) {
    const { sheetName = null, rowNumbers = {} } = options;

    return new _GasSheetDbTable({
      spreadsheet: this.spreadsheet,
      sheetName,
      rowNumbers: {
        ...this.rowNumbers,
        ...rowNumbers,
      },
    });
  }
}
