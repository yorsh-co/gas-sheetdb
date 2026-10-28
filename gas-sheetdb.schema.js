'use strict';
/**
 * Manages sheet column keys and schema updates.
 */
class _GasSheetDbTableSchema {
  constructor(sheet, rowNumbers, logger) {
    /**
     * Normalised header keys, one per header cell so indexes match sheet
     * columns. A blank or repeated header is `''`, which no entry can address.
     */
    this.columnKeys = [];
    /** Header names repeated by a later column, from the last header read. */
    this.duplicateKeys = [];
    this.sheet = sheet;
    this.rowNumbers = rowNumbers;
    this.logger = logger;
    this.reload();
  }
  /**
   * Set the column keys from a raw header row: every cell is read as a
   * string, blank cells are ignored, and when a name repeats the first
   * column wins and later ones are ignored like blanks.
   */
  setColumnKeys(headerRow) {
    const seen = new Set();
    const duplicates = new Set();
    this.columnKeys = headerRow.map((cell) => {
      const key = String(cell);
      if (!key) return '';
      if (seen.has(key)) {
        duplicates.add(key);
        return '';
      }
      seen.add(key);
      return key;
    });
    this.duplicateKeys = [...duplicates];
  }
  /**
   * Load column key values from the sheet, normalised by `setColumnKeys`.
   */
  loadColumnKeys() {
    const numCols = this.sheet.getLastColumn();
    this.setColumnKeys(
      numCols
        ? this.sheet
            .getRange(this.rowNumbers.columnKeys, 1, 1, numCols)
            .getValues()[0]
        : [],
    );
    return this.columnKeys;
  }
  /**
   * Reload the column key values from the sheet.
   */
  reload() {
    this.loadColumnKeys();
  }
  /**
   * Ensure all column keys exist in the sheet.
   * Missing columns are appended automatically.
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
      .getRange(
        this.rowNumbers.columnKeys,
        lastCol + 1,
        1,
        newColumnKeys.length,
      )
      .setValues([newColumnKeys.map(_GasSheetDbValueCodec.toSheetInput)]);
    this.columnKeys.push(...newColumnKeys);
    this.logger.info('Added columns', {
      sheet: this.sheet.getSheetName(),
      columns: newColumnKeys,
    });
  }
}
