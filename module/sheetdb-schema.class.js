/**
 * Manages sheet column headers and schema updates.
 */
class _SheetDbTableSchema {
  /**
   * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
   * @param {Object} rowNumbers // Ex: `SHEETDB_ROW_NUMBERS`
   */
  constructor(sheet, rowNumbers) {
    this.sheet = sheet;
    this.rows = rowNumbers;

    this.headers = this.loadHeaders();
  }

  /**
   * Load header values from the sheet.
   *
   * @returns {string[]}
   */
  loadHeaders() {
    const numCols = this.sheet.getLastColumn();

    if (!numCols) {
      return [];
    }

    const values = this.sheet
      .getRange(this.rows.headers, 1, 1, numCols)
      .getValues();

    return values[0];
  }

  /**
   * Reload the header values from the sheet.
   */
  reload() {
    this.headers = this.loadHeaders();
  }

  /**
   * Ensure all headers exist in the sheet.
   * Missing columns are appended automatically.
   *
   * @param {string[]} headers
   */
  ensureColumns(headers) {
    const newHeaders = headers.filter(
      (header) => !this.headers.includes(header),
    );

    if (!newHeaders.length) {
      return;
    }

    const maxCols = this.sheet.getMaxColumns();
    const lastCol = this.sheet.getLastColumn();

    const colsNeeded = newHeaders.length + lastCol - maxCols;

    if (colsNeeded > 0) {
      this.sheet.insertColumnsAfter(Math.max(lastCol, 1), colsNeeded);
    }

    this.sheet
      .getRange(this.rows.headers, lastCol + 1, 1, newHeaders.length)
      .setValues([newHeaders]);

    this.headers.push(...newHeaders);
  }
}
