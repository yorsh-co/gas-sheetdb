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
  lockScope: GasSheetDbLockScope;
  lockService: GasSheetDbLockService;
  constructor({
    spreadsheet,
    sheetName,
    rowNumbers,
    lockScope,
    lockService,
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
  private _insertSheet;
  /**
   * Read all rows as objects.
   * Adds runtime metadata to each entry.
   */
  find(options?: GasSheetDbFindOptions): GasSheetDbEntry[];
  /**
   * Read all rows as objects, without acquiring the lock.
   *
   * Callers are responsible for holding the lock. This exists so that a
   * single public method can read and then write within one lock — see
   * `_withLock` for why nesting the public methods is not an option.
   */
  private _findUnlocked;
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
   * Update multiple existing entries, without acquiring the lock.
   *
   * Callers are responsible for holding the lock.
   */
  private _updateManyUnlocked;
  /**
   * Update every entry matching a predicate, holding a single lock across
   * both the read and the write.
   *
   * `update` is either a patch object applied to every match, or a function
   * returning a patch for a given entry. A function that mutates its entry
   * in place and returns nothing is honoured too.
   *
   * Returns the updated entries, or an empty array when nothing matched.
   */
  updateWhere(
    predicateFn: GasSheetDbPredicate,
    update: GasSheetDbUpdate,
    options?: GasSheetDbFindOptions,
  ): GasSheetDbEntry[];
  /**
   * Update the first entry matching a predicate, holding a single lock
   * across both the read and the write.
   *
   * Returns the updated entry, or `null` when nothing matched.
   */
  updateOneWhere(
    predicateFn: GasSheetDbPredicate,
    update: GasSheetDbUpdate,
    options?: GasSheetDbFindOptions,
  ): GasSheetDbEntry | null;
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
   * Soft delete every entry matching a predicate, holding a single lock
   * across both the read and the write.
   *
   * Already-trashed entries are excluded from matching by default.
   *
   * Returns the soft-deleted entries, or an empty array when nothing matched.
   */
  softDeleteWhere(
    predicateFn: GasSheetDbPredicate,
    options?: GasSheetDbFindOptions,
  ): GasSheetDbEntry[];
  /**
   * Permanently delete a single entry.
   * Requires `_id`.
   */
  deleteOne(entry: GasSheetDbEntry): void;
  /**
   * Permanently delete multiple existing entries.
   * Requires `_id` for each entry.
   */
  deleteMany(entries: GasSheetDbEntry[]): void;
  /**
   * Permanently delete multiple existing entries, without acquiring the lock.
   *
   * Callers are responsible for holding the lock.
   */
  private _deleteManyUnlocked;
  /**
   * Permanently delete every entry matching a predicate, holding a single
   * lock across both the read and the write.
   *
   * Trashed entries are excluded from matching by default; pass
   * `{ onlyTrashed: true }` to purge the trash.
   *
   * Returns the entries as they were before deletion, or an empty array
   * when nothing matched.
   */
  deleteWhere(
    predicateFn: GasSheetDbPredicate,
    options?: GasSheetDbFindOptions,
  ): GasSheetDbEntry[];
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
   * Exits early if the table body has no rows to avoid
   * exception 'it is not possible to delete all non-frozen rows'.
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
   * Resolve the patch to apply to a matched entry.
   *
   * An updater that mutates its entry in place and returns nothing falls
   * back to the entry itself, so an in-place edit is never silently
   * discarded. `_id` is taken from the matched entry last, so an updater can
   * neither drop it nor redirect the write to a different row. Matched
   * entries always carry one, since `_ensureRequiredMetadata` backfills it
   * before any row is read.
   */
  private _resolveUpdate;
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
   * Execute a callback inside this table's configured lock, via whichever
   * lockService was injected (or the minimal default fallback).
   *
   * The flush runs inside the lock, resyncing Apps Script's
   * local spreadsheet model.
   *
   * Public methods acquire the lock exactly once and must never call one
   * another: the default lockService is non-reentrant, so a nested call
   * blocks on a lock this execution already holds until it times out.
   * Methods that need to read and then write within one lock compose the
   * private `_*Unlocked` cores instead.
   */
  private _withLock;
}
