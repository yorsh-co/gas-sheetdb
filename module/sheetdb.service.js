/**
 * Entry point for database-like table access in sheets.
 */
const SheetDb = (() => {
  let _spreadsheet;

  /**
   * Get the configured spreadsheet instance.
   * The spreadsheet is cached after first access.
   *
   * Uses the active spreadsheet when
   * `SHEETDB_USE_ACTIVE_SPREADSHEET` is enabled.
   *
   * @returns {GoogleAppsScript.Spreadsheet.Spreadsheet}
   */
  const getSpreadsheet = () => {
    if (_spreadsheet) {
      return _spreadsheet;
    }

    if (SHEETDB_USE_ACTIVE_SPREADSHEET) {
      _spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

      if (!_spreadsheet) {
        throw new Error('No active spreadsheet is available for this script.');
      }

      return _spreadsheet;
    }

    if (!SHEETDB_SPREADSHEET_URL) {
      throw new Error(
        'Required value "SHEETDB_SPREADSHEET_URL" is missing from "spreadsheet.config.js"',
      );
    }

    _spreadsheet = SpreadsheetApp.openByUrl(SHEETDB_SPREADSHEET_URL);

    return _spreadsheet;
  };

  /**
   * Create a table wrapper for a sheet.
   *
   * @param {string} sheetName
   * @returns {_SheetDbTable}
   */
  const table = (sheetName) => {
    return new _SheetDbTable({
      spreadsheet: getSpreadsheet(),
      sheetName,
    });
  };

  return Object.freeze({
    getSpreadsheet,
    table,
  });
})();
