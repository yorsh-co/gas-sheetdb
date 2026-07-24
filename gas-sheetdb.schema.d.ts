/**
 * Manages sheet column keys and schema updates.
 */
declare class _GasSheetDbTableSchema {
  sheet: GoogleAppsScript.Spreadsheet.Sheet;
  rowNumbers: GasSheetDbRowsReference;
  columnKeys: string[];
  constructor(
    sheet: GoogleAppsScript.Spreadsheet.Sheet,
    rowNumbers: GasSheetDbRowsReference,
  );
  /**
   * Load column key values from the sheet.
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
