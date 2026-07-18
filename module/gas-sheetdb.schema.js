/**
 * Manages sheet column keys and schema updates.
 */
class _GasSheetDbTableSchema {
  /**
   * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
   * @param {_GasSheetDbRowsReference} rowNumbers
   */
  constructor(sheet, rowNumbers) {
    this.sheet = sheet;
    this.rowNumbers = rowNumbers;

    this.columnKeys = this.loadColumnKeys();
  }

  /**
   * Load column key values from the sheet.
   *
   * @returns {string[]}
   */
  loadColumnKeys() {
    const numCols = this.sheet.getLastColumn();

    if (!numCols) {
      return [];
    }

    const values = this.sheet
      .getRange(this.rowNumbers.columnKeys, 1, 1, numCols)
      .getValues();

    return values[0];
  }

  /**
   * Reload the column key values from the sheet.
   */
  reload() {
    this.columnKeys = this.loadColumnKeys();
  }

  /**
   * Ensure all column keys exist in the sheet.
   * Missing columns are appended automatically.
   *
   * @param {string[]} columnKeys
   */
  ensureColumns(columnKeys) {
    const newColumnKeys = columnKeys.filter(
      (key) => !this.columnKeys.includes(key),
    );

    if (!newColumnKeys.length) {
      return;
    }

    const maxCols = this.sheet.getMaxColumns();
    const lastCol = this.sheet.getLastColumn();

    const colsNeeded = newColumnKeys.length + lastCol - maxCols;

    if (colsNeeded > 0) {
      this.sheet.insertColumnsAfter(Math.max(lastCol, 1), colsNeeded);
    }

    this.sheet
      .getRange(this.rowNumbers.columnKeys, lastCol + 1, 1, newColumnKeys.length)
      .setValues([newColumnKeys]);

    this.columnKeys.push(...newColumnKeys);
  }
}
