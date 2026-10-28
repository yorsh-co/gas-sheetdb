/**
 * Manages sheet column keys and schema updates.
 */
declare class _GasSheetDbTableSchema {
  sheet: GoogleAppsScript.Spreadsheet.Sheet;
  rowNumbers: GasSheetDbRowsReference;
  logger: GasSheetDbLogger;
  /**
   * Normalised header keys, one per header cell so indexes match sheet
   * columns. A blank or repeated header is `''`, which no entry can address.
   */
  columnKeys: string[];
  /** Header names repeated by a later column, from the last header read. */
  duplicateKeys: string[];
  constructor(
    sheet: GoogleAppsScript.Spreadsheet.Sheet,
    rowNumbers: GasSheetDbRowsReference,
    logger: GasSheetDbLogger,
  );
  /**
   * Set the column keys from a raw header row: every cell is read as a
   * string, blank cells are ignored, and when a name repeats the first
   * column wins and later ones are ignored like blanks.
   */
  setColumnKeys(headerRow: GasSheetDbCellValue[]): void;
  /**
   * Load column key values from the sheet, normalised by `setColumnKeys`.
   */
  loadColumnKeys(): string[];
  /**
   * Reload the column key values from the sheet.
   */
  reload(): void;
  /**
   * Ensure all column keys exist in the sheet.
   * Missing columns are appended automatically.
   */
  ensureColumns(columnKeys: string[]): void;
}
