/**
 * Table-style wrapper around a Google Spreadsheet sheet.
 *
 * Provides object-based querying, inserts, updates,
 * automatic column management, and metadata handling.
 */
declare class _GasSheetDbTable {
  spreadsheet: GoogleAppsScript.Spreadsheet.Spreadsheet;
  sheetName: string | null;
  rowNumbers: GasSheetDbRowsReference;
  rowIndexes: GasSheetDbRowsReference;
  sheet: GoogleAppsScript.Spreadsheet.Sheet;
  schema: _GasSheetDbTableSchema;
  constructor({
    spreadsheet,
    sheetName,
    rowNumbers,
  }: GasSheetDbTableConstructorOptions);
  /**
   * Derive base-0 data array row indexes from
   * base-1 spreadsheet row numbers.
   */
  private _deriveRowIndexes;
  /**
   * Ensure the sheet exists, by returning the existing
   * sheet, or inserting a new one.
   */
  private _ensureSheet;
  /**
   * Insert a new sheet to the spreadsheet, using the provided
   * name. Applies basic formatting including alternating
   * row colors and a frozen header row.
   */
  _insertSheet(): GoogleAppsScript.Spreadsheet.Sheet;
  /**
   * Read all rows as objects.
   * Adds runtime metadata to each entry.
   */
  find(options?: GasSheetDbFindOptions): GasSheetDbEntry[];
  /**
   * Filter entries using a predicate.
   */
  findWhere(
    predicateFn: (entry: GasSheetDbEntry) => boolean,
  ): GasSheetDbEntry[] | null;
  /**
   * Find the first matching entry.
   */
  findOneWhere(
    predicateFn: (entry: GasSheetDbEntry) => boolean,
  ): GasSheetDbEntry | null;
  /**
   * Find trashed entries.
   * Optionally, filter trashed entries.
   */
  findTrashed(
    predicateFn?: ((entry: GasSheetDbEntry) => boolean) | null,
  ): GasSheetDbEntry[];
  /**
   * Insert a single entry.
   */
  insert(entry: GasSheetDbEntry): GasSheetDbEntry;
  /**
   * Insert multiple entries.
   * Missing columns are created automatically.
   */
  insertMany(entries: GasSheetDbEntry[]): GasSheetDbEntry[];
  /**
   * Update a single entry.
   * Requires `_id`.
   */
  update(entry: GasSheetDbEntry): GasSheetDbEntry;
  /**
   * Update multiple existing entries.
   * Requires `_id` for each entry.
   */
  updateMany(entries: GasSheetDbEntry[]): GasSheetDbEntry[];
  /**
   * Soft delete a single entry.
   * Requires `_id`.
   */
  softDelete(entry: GasSheetDbEntry): GasSheetDbEntry;
  /**
   * Soft delete multiple existing entries.
   * Requires `_id` for each entry.
   */
  softDeleteMany(entries: GasSheetDbEntry[]): GasSheetDbEntry[];
  /**
   * Restore a single soft-deleted entry.
   * Requires `_id`.
   */
  restore(entry: GasSheetDbEntry): GasSheetDbEntry;
  /**
   * Restore multiple soft-deleted entries.
   * Requires `_id` for each entry.
   */
  restoreMany(entries: GasSheetDbEntry[]): GasSheetDbEntry[];
  /**
   * Permanently delete a single entry.
   * Requires `_id`.
   */
  delete(entry: GasSheetDbEntry): void;
  /**
   * Permanently delete multiple existing entries.
   * Requires `_id` for each entry.
   */
  deleteMany(entries: GasSheetDbEntry[]): void;
  /**
   * Return the table body, without the column keys row.
   */
  private _getTableBodyData;
  /**
   * Write to the entire table body on the sheet.
   */
  private _setTableBodyData;
  /**
   * Append rows to the sheet.
   */
  private _appendTableBodyData;
  /**
   * Ensure the sheet has enough empty rows.
   */
  private _ensureBlankRows;
  /**
   * Remove extra rows from the sheet.
   */
  private _removeExtraRows;
  /**
   * Map the row indexes
   */
  private _mapTableBodyRowIndexesById;
  /**
   * Return the base-0 index for column corresponding to the provided key.
   */
  private _getColumnIndex;
  /**
   * Extract the column keys from an entry's property keys.
   */
  private _extractKeys;
  /**
   * Check if a field should be stored in the sheet.
   */
  private _isPersistedField;
  /**
   * Build a sheet row from an entry.
   * Existing values are preserved when a field is undefined.
   */
  private _buildRow;
  /**
   * Decode a raw sheet row into an entry object.
   */
  private _decodeRow;
  /**
   * Ensure in-sheet insert metadata for all entries.
   * This ensures that entries inserted manually or by processes
   * other than table.insert() contain the required metadata for
   * reading and writing with SheetDB.
   */
  private _ensureRequiredMetadata;
  /**
   * Apply system fields for new entries.
   */
  private _applyInsertMetadata;
  /**
   * Update system timestamps.
   */
  private _applyUpdateMetadata;
  /**
   * Return uuid.
   */
  private _generateId;
  /**
   * Execute a callback inside a document lock.
   */
  private _withLock;
}
