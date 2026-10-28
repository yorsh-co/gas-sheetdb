/**
 * Table-style wrapper around a Google Spreadsheet sheet.
 *
 * Provides object-based querying, inserts, updates,
 * automatic column management, and metadata handling.
 */
declare class _GasSheetDbTable {
  spreadsheet: GoogleAppsScript.Spreadsheet.Spreadsheet;
  sheetName: string;
  rowNumbers: GasSheetDbRowsReference;
  rowIndexes: GasSheetDbRowsReference;
  sheet: GoogleAppsScript.Spreadsheet.Sheet;
  schema: _GasSheetDbTableSchema;
  lockScope: GasSheetDbLockScope;
  lockService: GasSheetDbLockService;
  lockTimeoutMs: number;
  logger: GasSheetDbLogger;
  /**
   * Log lines queued while the lock is held, or null outside `_withLock`.
   * See `_withLock` for why nothing is logged under the lock.
   */
  private _deferredLogs;
  /**
   * The logger this table and its schema log through: it queues lines while
   * the lock is held and forwards them to `logger` once it is released.
   * Keeps the schema unaware of locking.
   */
  private readonly _logSink;
  constructor({
    spreadsheet,
    sheetName,
    rowNumbers,
    lockScope,
    lockService,
    lockTimeoutMs,
    logger,
  }: GasSheetDbTableConstructorOptions);
  /**
   * Derive base-0 data array row indexes from
   * base-1 spreadsheet row numbers.
   */
  private _deriveRowIndexes;
  /**
   * Ensure the sheet exists, by returning the existing
   * sheet, or inserting a new one under the lock. The lookup is repeated
   * once the lock is held, as another execution may have created the sheet
   * in the meantime.
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
   *
   * Reads without the lock unless the system columns or any row's metadata
   * need repairing, in which case the repair and the read run under it.
   */
  find(options?: GasSheetDbFindOptions): GasSheetDbEntry[];
  /**
   * Read all rows as objects without acquiring the lock, or return null when
   * the sheet needs repairing first (see `_ensureRequiredMetadata`).
   *
   * Header and body come from a single `getDataRange()` read, so they are
   * consistent with each other even while another execution writes. Nothing
   * is flushed: a read already observes this execution's own writes, and
   * there is nothing of ours to publish.
   */
  private _findWithoutLock;
  /**
   * Read all rows as objects, without acquiring the lock.
   *
   * Callers are responsible for holding the lock. This exists so that a
   * single public method can read and then write within one lock — see
   * `_withLock` for why nesting the public methods is not an option.
   */
  private _findUnlocked;
  /**
   * Decode the rows matching the trashed filters of `options`.
   */
  private _selectEntries;
  /**
   * Filter entries using a predicate. Trashed entries are excluded by
   * default; see `GasSheetDbFindOptions`.
   *
   * Returns an empty array when nothing matches.
   */
  findWhere(
    predicateFn: GasSheetDbPredicate,
    options?: GasSheetDbFindOptions,
  ): GasSheetDbEntry[];
  /**
   * Find the first matching entry. Trashed entries are excluded by default;
   * see `GasSheetDbFindOptions`.
   */
  findOneWhere(
    predicateFn: GasSheetDbPredicate,
    options?: GasSheetDbFindOptions,
  ): GasSheetDbEntry | null;
  /**
   * Find trashed entries.
   * Optionally, filter trashed entries.
   */
  findTrashed(predicateFn?: GasSheetDbPredicate | null): GasSheetDbEntry[];
  /**
   * Insert a single entry.
   */
  insert(entry: GasSheetDbEntry): GasSheetDbEntry;
  /**
   * Insert multiple entries.
   * Missing columns are created automatically.
   *
   * Throws a conflict, before writing anything or changing any entry, when an
   * entry's `_id` is already stored or repeats within the batch. Returns an
   * empty array for an empty batch.
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
   *
   * Every entry is validated before any is changed or written, so a batch
   * with a missing or unknown `_id` leaves the sheet and the passed objects
   * untouched. Returns an empty array for an empty batch.
   */
  updateMany(entries: GasSheetDbEntry[]): GasSheetDbEntry[];
  /**
   * Update multiple existing entries, without acquiring the lock.
   *
   * `patch` is assigned to every entry once the whole batch is validated.
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
   *
   * Validated as a whole like `updateMany`. Returns an empty array for an
   * empty batch.
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
   *
   * Validated as a whole like `updateMany`. Returns an empty array for an
   * empty batch.
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
   * Write only the cells of `data` that differ from the rows as they were
   * read, so formulas, array-formula spills and manual edits elsewhere are
   * left alone.
   *
   * Changed cells are grouped into column runs per row; consecutive rows
   * with the same runs are written together, one `setValues` call per run.
   */
  private _writeChangedCells;
  /**
   * Whether two cell values are equal, treating a missing cell as blank and
   * comparing dates by time.
   */
  private _isSameCell;
  /**
   * Group ascending, unique indexes into contiguous `[start, count]` runs.
   */
  private _groupRuns;
  /**
   * Append rows after the last non-blank row of the table body, `body` being
   * the body as read before the append.
   *
   * The position comes from the body rather than `getLastRow()`: an
   * open-ended array formula spills `''` down to the last row of the sheet,
   * and Sheets counts that as content. Blank rows below the last entry are
   * reused, and rows are added only when the sheet runs out.
   *
   * Only the cells that differ from the target rows are written, so a
   * spilled or formula cell in a column the entries do not set is left alone.
   */
  private _appendTableBodyData;
  /**
   * Ensure the sheet has enough empty rows for `amount` rows from `startRow`.
   */
  private _ensureBlankRows;
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
   * Check if a field should be stored in the sheet. The empty key stands for
   * blank and repeated headers, which are never written.
   */
  private _isPersistedField;
  /**
   * Throw a validation error when a persisted field of any entry holds a
   * value that cannot be stored. Runs before anything is written or changed.
   */
  private _assertStorable;
  /**
   * Build a sheet row from an entry.
   * Existing values are preserved when a field is undefined, and in columns
   * without a key.
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
   * Decode raw sheet rows into entry objects.
   *
   * Cells whose stored JSON no longer parses fall back to their raw string.
   * Those are counted across the whole batch and reported as a single log
   * line, so one corrupt column cannot emit a log line per row — which, with
   * a sheet-backed logger, would mean a sheet write per row inside the lock.
   */
  private _decodeRows;
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
   * Return the base-0 column indexes of the system fields, -1 for any that
   * is not on the sheet.
   */
  private _getSystemColumnIndexes;
  /**
   * Whether a row is entirely empty, as when a person cleared its contents.
   * A row with a formula is blank only while the formula evaluates to ''.
   */
  private _isBlankRow;
  /**
   * Whether a row lacks any system metadata that
   * `_ensureRequiredMetadata` would backfill. A blank row is not an entry,
   * so it never does.
   */
  private _rowNeedsBackfill;
  /**
   * Indexes of the rows whose `_id` repeats that of a row above them. The
   * topmost row keeps the id.
   */
  private _findDuplicateIdRows;
  /**
   * Throw a conflict when a supplied `_id` is already stored in `body` or
   * repeats within `entries`.
   */
  private _assertUniqueIds;
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
   * lockService was injected (or the default: GasLock when loaded, else the
   * minimal fallback), waiting up to `lockTimeoutMs` for it.
   *
   * The flush runs inside the lock, resyncing Apps Script's
   * local spreadsheet model.
   *
   * Public methods acquire the lock exactly once and must never call one
   * another. Apps Script locks are owned per execution, so a nested acquire
   * succeeds at once, but with the default lockService its release then
   * ends the outer critical section early, leaving the rest of the outer
   * method unprotected. Methods that need to read and then write within one
   * lock compose the private `_*Unlocked` cores instead.
   *
   * Log lines raised by the callback are queued and emitted once the lock is
   * released, in order, whether or not it threw. A sheet-backed logger takes
   * the script lock itself, and its release would end this critical section
   * early.
   */
  private _withLock;
  /**
   * Warn about repeated header names in the last header read, once per
   * operation: called as an operation ends, never while the lock is held.
   */
  private _warnDuplicateColumns;
  /**
   * Queue a log line while the lock is held, else emit it at once.
   */
  private _log;
  /**
   * Forward a log line to the logger. A throwing logger never fails the
   * operation, which has already run, nor masks its error: the failure is
   * reported through `console` and the remaining lines are still emitted.
   */
  private _emit;
}
