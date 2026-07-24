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

  constructor(options: GasSheetDbOptions = {}) {
    this.spreadsheet = this._resolveSpreadsheet(options);

    this.rowNumbers = {
      ..._GAS_SHEET_DB_DEFAULT_ROW_NUMBERS,
      ...options.rowNumbers,
    };
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
   * Create a table wrapper for a sheet.
   */
  table(options: GasSheetDbTableOptions = {}): _GasSheetDbTable {
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
