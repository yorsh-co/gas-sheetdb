'use strict';
/**
 * Table-style wrapper around a Google Spreadsheet sheet.
 *
 * Provides object-based querying, inserts, updates,
 * automatic column management, and metadata handling.
 */
class _GasSheetDbTable {
  constructor({
    spreadsheet,
    sheetName,
    rowNumbers,
    lockScope,
    lockService,
    lockTimeoutMs,
    logger,
  }) {
    /**
     * Log lines queued while the lock is held, or null outside `_withLock`.
     * See `_withLock` for why nothing is logged under the lock.
     */
    this._deferredLogs = null;
    /**
     * The logger this table and its schema log through: it queues lines while
     * the lock is held and forwards them to `logger` once it is released.
     * Keeps the schema unaware of locking.
     */
    this._logSink = {
      info: (msg, meta) => this._log({ level: 'info', msg, meta }),
      warn: (msg, meta) => this._log({ level: 'warn', msg, meta }),
    };
    this.spreadsheet = spreadsheet;
    this.sheetName = sheetName;
    this.rowNumbers = rowNumbers; // Base-1 row numbers
    this.rowIndexes = this._deriveRowIndexes(); // Base-0 row indexes
    // assigned before _ensureSheet(), which locks and logs
    this.lockScope = lockScope;
    this.lockService = lockService;
    this.lockTimeoutMs = lockTimeoutMs;
    this.logger = logger;
    this.sheet = this._ensureSheet();
    this.schema = new _GasSheetDbTableSchema(
      this.sheet,
      this.rowNumbers,
      this._logSink,
    );
  }
  // =========================
  // CONSTRUCTOR HELPERS
  // =========================
  /**
   * Derive base-0 data array row indexes from
   * base-1 spreadsheet row numbers.
   */
  _deriveRowIndexes() {
    const rowIndexes = {};
    Object.entries(this.rowNumbers).forEach(
      ([k, v]) => (rowIndexes[k] = v - 1),
    );
    return rowIndexes;
  }
  /**
   * Ensure the sheet exists, by returning the existing
   * sheet, or inserting a new one under the lock. The lookup is repeated
   * once the lock is held, as another execution may have created the sheet
   * in the meantime.
   */
  _ensureSheet() {
    const sheet = this.spreadsheet.getSheetByName(this.sheetName);
    if (sheet) return sheet;
    return this._withLock(
      () =>
        this.spreadsheet.getSheetByName(this.sheetName) || this._insertSheet(),
    );
  }
  /**
   * Insert a new sheet to the spreadsheet, using the provided
   * name. Applies basic formatting including alternating
   * row colors and a frozen header row.
   */
  _insertSheet() {
    const newSheet = this.spreadsheet.insertSheet(this.sheetName);
    const maxCols = newSheet.getMaxColumns();
    newSheet
      .getRange(
        this.rowNumbers.columnKeys,
        1,
        newSheet.getMaxRows() - this.rowNumbers.columnKeys + 1,
        maxCols,
      )
      .applyRowBanding();
    newSheet
      .getRange(this.rowNumbers.columnKeys, 1, 1, maxCols)
      .setFontWeight('bold');
    newSheet.setFrozenRows(this.rowNumbers.columnKeys);
    this._logSink.info('Created table sheet', { sheet: this.sheetName });
    return newSheet;
  }
  // =========================
  // QUERY
  // =========================
  /**
   * Read all rows as objects.
   *
   * Reads without the lock unless the system columns or any row's metadata
   * need repairing, in which case the repair and the read run under it.
   */
  find(options = {}) {
    const entries = this._findWithoutLock(options);
    if (!entries) return this._withLock(() => this._findUnlocked(options));
    this._warnDuplicateColumns();
    return entries;
  }
  /**
   * Read all rows as objects without acquiring the lock, or return null when
   * the sheet needs repairing first (see `_ensureRequiredMetadata`).
   *
   * Header and body come from a single `getDataRange()` read, so they are
   * consistent with each other even while another execution writes. Nothing
   * is flushed: a read already observes this execution's own writes, and
   * there is nothing of ours to publish.
   */
  _findWithoutLock(options) {
    const data = this.sheet.getDataRange().getValues();
    this.schema.setColumnKeys(data[this.rowIndexes.columnKeys] ?? []);
    if (
      Object.values(_GAS_SHEETDB_SYSTEM_FIELDS).some(
        (key) => !this.schema.columnKeys.includes(key),
      )
    ) {
      return null;
    }
    const body = data.slice(this.rowIndexes.firstData);
    const columns = this._getSystemColumnIndexes();
    if (
      body.some((row) => this._rowNeedsBackfill(row, columns)) ||
      this._findDuplicateIdRows(body, columns.id).size
    ) {
      return null;
    }
    return this._selectEntries(body, options);
  }
  /**
   * Read all rows as objects, without acquiring the lock.
   *
   * Callers are responsible for holding the lock. This exists so that a
   * single public method can read and then write within one lock — see
   * `_withLock` for why nesting the public methods is not an option.
   */
  _findUnlocked(options = {}) {
    this._ensureRequiredMetadata();
    return this._selectEntries(this._getTableBodyData(), options);
  }
  /**
   * Decode the rows matching the trashed filters of `options`.
   */
  _selectEntries(data, options) {
    const { withTrashed = false, onlyTrashed = false } = options;
    const isTrashedColIndex = this._getColumnIndex(
      _GAS_SHEETDB_SYSTEM_FIELDS.IS_DELETED,
    );
    const rows = [];
    for (let dataIndex = 0; dataIndex < data.length; dataIndex++) {
      const row = data[dataIndex];
      if (this._isBlankRow(row ?? [])) continue;
      // handle `onlyTrashed = true`
      if (onlyTrashed && !row?.[isTrashedColIndex]) continue;
      // handle `withTrashed = false`
      if (!onlyTrashed && !withTrashed && row?.[isTrashedColIndex]) continue;
      rows.push(row || []);
    }
    return this._decodeRows(rows);
  }
  /**
   * Filter entries using a predicate. Trashed entries are excluded by
   * default; see `GasSheetDbFindOptions`.
   *
   * Returns an empty array when nothing matches.
   */
  findWhere(predicateFn, options = {}) {
    return this.find(options).filter(predicateFn);
  }
  /**
   * Find the first matching entry. Trashed entries are excluded by default;
   * see `GasSheetDbFindOptions`.
   */
  findOneWhere(predicateFn, options = {}) {
    return this.find(options).find(predicateFn) || null;
  }
  /**
   * Find trashed entries.
   * Optionally, filter trashed entries.
   */
  findTrashed(predicateFn = null) {
    if (predicateFn) {
      return this.find({ onlyTrashed: true }).filter(predicateFn);
    }
    return this.find({ onlyTrashed: true });
  }
  // =========================
  // INSERT
  // =========================
  /**
   * Insert a single entry.
   */
  insert(entry) {
    return this.insertMany([entry])[0];
  }
  /**
   * Insert multiple entries.
   * Missing columns are created automatically.
   *
   * Throws a conflict, before writing anything or changing any entry, when an
   * entry's `_id` is already stored or repeats within the batch. Returns an
   * empty array for an empty batch.
   */
  insertMany(entries) {
    if (!entries.length) return [];
    return this._withLock(() => {
      this.schema.reload();
      // one read serves the uniqueness check and the append position
      const body = this._getTableBodyData();
      this._assertUniqueIds(entries, body);
      this._assertStorable(entries);
      entries.forEach((entry) => this._applyInsertMetadata(entry));
      const columnKeys = this._extractKeys(entries);
      this.schema.ensureColumns(columnKeys);
      this.schema.reload();
      const data = entries.map((entry) => this._buildRow(entry));
      this._appendTableBodyData(data, body);
      const insertedEntries = this._decodeRows(data);
      return insertedEntries;
    });
  }
  // =========================
  // UPDATE
  // =========================
  /**
   * Update a single entry.
   * Requires `_id`.
   */
  update(entry) {
    return this.updateMany([entry])[0];
  }
  /**
   * Update multiple existing entries.
   * Requires `_id` for each entry.
   *
   * Every entry is validated before any is changed or written, so a batch
   * with a missing or unknown `_id` leaves the sheet and the passed objects
   * untouched. Returns an empty array for an empty batch.
   */
  updateMany(entries) {
    if (!entries.length) return [];
    return this._withLock(() => this._updateManyUnlocked(entries));
  }
  /**
   * Update multiple existing entries, without acquiring the lock.
   *
   * `patch` is assigned to every entry once the whole batch is validated.
   * Callers are responsible for holding the lock.
   */
  _updateManyUnlocked(entries, patch = {}) {
    this.schema.reload();
    // load the current data after schema reload
    const data = this._getTableBodyData();
    // map the row indexes; a sheet without `_id` holds no entry to find
    const rowIndexesById =
      this._getColumnIndex(_GAS_SHEETDB_SYSTEM_FIELDS.ID) === -1
        ? new Map()
        : this._mapTableBodyRowIndexesById(data);
    // validate the whole batch before changing anything; the row indexes
    // also let the entries be returned in their post-`_buildRow` state
    const updatedRowIndexes = entries.map((entry) => {
      if (!entry._id) {
        throw _GasSheetDbErrors.validation(
          '[GasSheetDb] Cannot update entry without "_id"',
        );
      }
      const rowIndex = rowIndexesById.get(entry._id);
      if (rowIndex === undefined) {
        throw _GasSheetDbErrors.notFound(
          `[GasSheetDb] Entry not found: ${entry._id}`,
        );
      }
      return rowIndex;
    });
    this._assertStorable(entries.map((entry) => ({ ...entry, ...patch })));
    entries.forEach((entry) => {
      Object.assign(entry, patch);
      this._applyUpdateMetadata(entry);
    });
    // add new property columns
    this.schema.ensureColumns(this._extractKeys(entries));
    this.schema.reload();
    // each touched row as read, to find the cells the update changes
    const originalRows = new Map();
    entries.forEach((entry, i) => {
      const rowIndex = updatedRowIndexes[i];
      if (!originalRows.has(rowIndex))
        originalRows.set(rowIndex, data[rowIndex]);
      data[rowIndex] = this._buildRow(entry, data[rowIndex]);
    });
    this._writeChangedCells(data, originalRows);
    // return the entries in their post-`_buildRow` state
    return this._decodeRows(
      updatedRowIndexes.map((rowIndex) => data[rowIndex]),
    );
  }
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
  updateWhere(predicateFn, update, options = {}) {
    return this._withLock(() => {
      const entries = this._findUnlocked(options).filter(predicateFn);
      if (!entries.length) return [];
      return this._updateManyUnlocked(
        entries.map((entry) => this._resolveUpdate(entry, update)),
      );
    });
  }
  /**
   * Update the first entry matching a predicate, holding a single lock
   * across both the read and the write.
   *
   * Returns the updated entry, or `null` when nothing matched.
   */
  updateOneWhere(predicateFn, update, options = {}) {
    return this._withLock(() => {
      const entry = this._findUnlocked(options).find(predicateFn);
      if (!entry) return null;
      return this._updateManyUnlocked([this._resolveUpdate(entry, update)])[0];
    });
  }
  // =========================
  // SOFT DELETE
  // =========================
  /**
   * Soft delete a single entry.
   * Requires `_id`.
   */
  softDelete(entry) {
    return this.softDeleteMany([entry])[0];
  }
  /**
   * Soft delete multiple existing entries.
   * Requires `_id` for each entry.
   *
   * Validated as a whole like `updateMany`. Returns an empty array for an
   * empty batch.
   */
  softDeleteMany(entries) {
    if (!entries.length) return [];
    return this._withLock(() =>
      this._updateManyUnlocked(entries, {
        [_GAS_SHEETDB_SYSTEM_FIELDS.IS_DELETED]: true,
      }),
    );
  }
  /**
   * Restore a single soft-deleted entry.
   * Requires `_id`.
   */
  restore(entry) {
    return this.restoreMany([entry])[0];
  }
  /**
   * Restore multiple soft-deleted entries.
   * Requires `_id` for each entry.
   *
   * Validated as a whole like `updateMany`. Returns an empty array for an
   * empty batch.
   */
  restoreMany(entries) {
    if (!entries.length) return [];
    return this._withLock(() =>
      this._updateManyUnlocked(entries, {
        [_GAS_SHEETDB_SYSTEM_FIELDS.IS_DELETED]: false,
      }),
    );
  }
  /**
   * Soft delete every entry matching a predicate, holding a single lock
   * across both the read and the write.
   *
   * Already-trashed entries are excluded from matching by default.
   *
   * Returns the soft-deleted entries, or an empty array when nothing matched.
   */
  softDeleteWhere(predicateFn, options = {}) {
    return this.updateWhere(
      predicateFn,
      { [_GAS_SHEETDB_SYSTEM_FIELDS.IS_DELETED]: true },
      options,
    );
  }
  // =========================
  // DELETE
  // =========================
  /**
   * Permanently delete a single entry.
   * Requires `_id`.
   */
  deleteOne(entry) {
    this.deleteMany([entry]);
  }
  /**
   * Permanently delete multiple existing entries.
   * Requires `_id` for each entry.
   */
  deleteMany(entries) {
    if (!entries.length) return;
    this._withLock(() => {
      this.schema.reload();
      this._deleteManyUnlocked(entries);
    });
  }
  /**
   * Permanently delete multiple existing entries, without acquiring the lock.
   *
   * Callers are responsible for holding the lock.
   */
  _deleteManyUnlocked(entries) {
    // load the current data after schema reload
    const data = this._getTableBodyData();
    // map the row indexes
    const rowIndexesById = this._mapTableBodyRowIndexesById(data);
    const entriesRowIndexes = [];
    for (const entry of entries) {
      if (!entry._id) {
        throw _GasSheetDbErrors.validation(
          '[GasSheetDb] Cannot delete entry without "_id"',
        );
      }
      const rowIndex = rowIndexesById.get(entry._id);
      if (rowIndex === undefined) {
        throw _GasSheetDbErrors.notFound(
          `[GasSheetDb] Entry not found: ${entry._id}`,
        );
      }
      entriesRowIndexes.push(rowIndex);
    }
    const uniqueRowIndexes = [...new Set(entriesRowIndexes)].sort(
      (a, b) => a - b,
    );
    if (!uniqueRowIndexes.length) return;
    // Sheets refuses to delete every non-frozen row: spare the last one, and
    // clear its content instead
    if (
      this.sheet.getMaxRows() - uniqueRowIndexes.length <=
      this.sheet.getFrozenRows()
    ) {
      this.sheet
        .getRange(
          this.rowNumbers.firstData + uniqueRowIndexes.pop(),
          1,
          1,
          this.sheet.getMaxColumns(),
        )
        .clearContent();
    }
    // delete from the bottom up, so earlier row numbers stay valid
    this._groupRuns(uniqueRowIndexes)
      .reverse()
      .forEach(([start, count]) =>
        this.sheet.deleteRows(this.rowNumbers.firstData + start, count),
      );
  }
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
  deleteWhere(predicateFn, options = {}) {
    return this._withLock(() => {
      const entries = this._findUnlocked(options).filter(predicateFn);
      if (!entries.length) return [];
      this._deleteManyUnlocked(entries);
      return entries;
    });
  }
  // =========================
  // TABLE BODY HELPERS
  // =========================
  /**
   * Return the table body, without the column keys row.
   */
  _getTableBodyData() {
    const data = this.sheet.getDataRange().getValues();
    if (!data.length) {
      return [];
    }
    const bodyOffset = this.rowIndexes.firstData;
    return data.slice(bodyOffset);
  }
  /**
   * Write only the cells of `data` that differ from the rows as they were
   * read, so formulas, array-formula spills and manual edits elsewhere are
   * left alone.
   *
   * Changed cells are grouped into column runs per row; consecutive rows
   * with the same runs are written together, one `setValues` call per run.
   */
  _writeChangedCells(data, originalRows) {
    const blocks = [];
    for (const rowIndex of [...originalRows.keys()].sort((a, b) => a - b)) {
      const original = originalRows.get(rowIndex);
      const changed = data[rowIndex].flatMap((value, colIndex) =>
        this._isSameCell(value, original[colIndex]) ? [] : [colIndex],
      );
      if (!changed.length) continue;
      const runs = this._groupRuns(changed);
      const last = blocks[blocks.length - 1];
      if (
        last &&
        last.start + last.count === rowIndex &&
        JSON.stringify(last.runs) === JSON.stringify(runs)
      ) {
        last.count++;
      } else {
        blocks.push({ start: rowIndex, count: 1, runs });
      }
    }
    for (const { start, count, runs } of blocks) {
      for (const [colStart, colCount] of runs) {
        this.sheet
          .getRange(
            this.rowNumbers.firstData + start,
            colStart + 1,
            count,
            colCount,
          )
          .setValues(
            data
              .slice(start, start + count)
              .map((row) =>
                row
                  .slice(colStart, colStart + colCount)
                  .map(_GasSheetDbValueCodec.toSheetInput),
              ),
          );
      }
    }
  }
  /**
   * Whether two cell values are equal, treating a missing cell as blank and
   * comparing dates by time.
   */
  _isSameCell(a, b) {
    if (a instanceof Date && b instanceof Date) {
      return a.getTime() === b.getTime();
    }
    return (a ?? '') === (b ?? '');
  }
  /**
   * Group ascending, unique indexes into contiguous `[start, count]` runs.
   */
  _groupRuns(indexes) {
    const runs = [];
    for (const index of indexes) {
      const run = runs[runs.length - 1];
      if (run && run[0] + run[1] === index) run[1]++;
      else runs.push([index, 1]);
    }
    return runs;
  }
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
  _appendTableBodyData(data, body) {
    if (!data.length || !data[0]?.length)
      throw new Error('[GasSheetDb] Empty data');
    let lastIndex = body.length - 1;
    while (lastIndex >= 0 && this._isBlankRow(body[lastIndex])) lastIndex--;
    // the body starts at `firstData`, so a spacer row above it is never used
    const offset = lastIndex + 1;
    this._ensureBlankRows(this.rowNumbers.firstData + offset, data.length);
    // the rows keep their body index, which is what `_writeChangedCells` reads
    const rows = [];
    const originalRows = new Map();
    data.forEach((row, i) => {
      rows[offset + i] = row;
      originalRows.set(offset + i, body[offset + i] ?? []);
    });
    this._writeChangedCells(rows, originalRows);
  }
  /**
   * Ensure the sheet has enough empty rows for `amount` rows from `startRow`.
   */
  _ensureBlankRows(startRow, amount) {
    const maxRows = this.sheet.getMaxRows();
    const rowsNeeded = startRow + amount - 1 - maxRows;
    if (rowsNeeded > 0) {
      this.sheet.insertRowsAfter(maxRows, rowsNeeded);
    }
  }
  /**
   * Map the row indexes
   */
  _mapTableBodyRowIndexesById(tableBody = null) {
    const data = tableBody || this._getTableBodyData();
    const idColIndex = this._getColumnIndex(_GAS_SHEETDB_SYSTEM_FIELDS.ID);
    if (idColIndex === -1) {
      throw new Error(
        `[GasSheetDb] Cannot map rows by "${_GAS_SHEETDB_SYSTEM_FIELDS.ID}": column does not exist on sheet "${this.sheetName}". Call find() at least once before update/delete.`,
      );
    }
    const rowMap = new Map();
    for (const [i, row] of data.entries()) {
      rowMap.set(row[idColIndex], i);
    }
    return rowMap;
  }
  // =========================
  // TABLE COLUMN HELPERS
  // =========================
  /**
   * Return the base-0 index for column corresponding to the provided key.
   */
  _getColumnIndex(key) {
    return this.schema.columnKeys.indexOf(key);
  }
  /**
   * Extract the column keys from an entry's property keys.
   */
  _extractKeys(entries) {
    return [
      ...new Set(
        entries.flatMap((entry) =>
          Object.keys(entry).filter((key) => this._isPersistedField(key)),
        ),
      ),
    ];
  }
  // =========================
  // TABLE ENTRY HELPERS
  // =========================
  /**
   * Check if a field should be stored in the sheet. The empty key stands for
   * blank and repeated headers, which are never written.
   */
  _isPersistedField(key) {
    return key !== '';
  }
  /**
   * Throw a validation error when a persisted field of any entry holds a
   * value that cannot be stored. Runs before anything is written or changed.
   */
  _assertStorable(entries) {
    for (const entry of entries) {
      for (const [column, value] of Object.entries(entry)) {
        if (value === undefined || !this._isPersistedField(column)) continue;
        const reason = _GasSheetDbValueCodec.unsupportedReason(value);
        if (reason) {
          throw _GasSheetDbErrors.validation(
            `[GasSheetDb] Cannot store column "${column}": ${reason}`,
            { column, reason },
          );
        }
      }
    }
  }
  /**
   * Build a sheet row from an entry.
   * Existing values are preserved when a field is undefined, and in columns
   * without a key.
   */
  _buildRow(entry, existingRow = []) {
    return this.schema.columnKeys.map((key, index) => {
      const value = key === '' ? undefined : entry[key];
      if (value === undefined) {
        return existingRow[index] ?? '';
      }
      return _GasSheetDbValueCodec.encode(value);
    });
  }
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
  _resolveUpdate(entry, update) {
    const patch = typeof update === 'function' ? update(entry) : update;
    return { ...(patch || entry), _id: entry._id };
  }
  /**
   * Decode raw sheet rows into entry objects.
   *
   * Cells whose stored JSON no longer parses fall back to their raw string.
   * Those are counted across the whole batch and reported as a single log
   * line, so one corrupt column cannot emit a log line per row — which, with
   * a sheet-backed logger, would mean a sheet write per row inside the lock.
   */
  _decodeRows(rows) {
    let failedCells = 0;
    let firstError;
    const onDecodeError = (err) => {
      if (!failedCells) firstError = err;
      failedCells++;
    };
    const entries = rows.map((row) => this._decodeRow(row, onDecodeError));
    if (failedCells) {
      this._logSink.warn('Failed to decode JSON cells', {
        sheet: this.sheetName,
        cells: failedCells,
        error: firstError?.message ?? String(firstError),
      });
    }
    return entries;
  }
  /**
   * Decode a raw sheet row into an entry object.
   */
  _decodeRow(row, onDecodeError) {
    const entry = {};
    this.schema.columnKeys.forEach((key, colIndex) => {
      if (key === '') return;
      const raw = row[colIndex];
      entry[key] = _GasSheetDbValueCodec.decode(
        raw === undefined ? '' : raw,
        onDecodeError,
      );
    });
    return entry;
  }
  /**
   * Ensure in-sheet insert metadata for all entries.
   * This ensures that entries inserted manually or by processes
   * other than table.insert() contain the required metadata for
   * reading and writing with SheetDB.
   */
  _ensureRequiredMetadata() {
    this.schema.reload();
    this.schema.ensureColumns(Object.values(_GAS_SHEETDB_SYSTEM_FIELDS));
    this.schema.reload();
    const data = this._getTableBodyData();
    if (!data.length) return;
    const columns = this._getSystemColumnIndexes();
    const {
      id: idColIndex,
      createdAt: createdAtColIndex,
      updatedAt: updatedAtColIndex,
      isDeleted: isTrashedColIndex,
    } = columns;
    const now = new Date();
    const duplicateIdRows = this._findDuplicateIdRows(data, idColIndex);
    let backfilledRows = 0;
    // written and logged as one batch rather than per row: a table filled by
    // hand or by another process backfills every row at once
    const originalRows = new Map();
    for (const [rowIndex, row] of data.entries()) {
      const needsBackfill = this._rowNeedsBackfill(row, columns);
      if (!needsBackfill && !duplicateIdRows.has(rowIndex)) continue;
      originalRows.set(rowIndex, [...row]);
      if (needsBackfill) backfilledRows++;
      if (!row[idColIndex] || duplicateIdRows.has(rowIndex)) {
        row[idColIndex] = this._generateId();
      }
      if (!row[createdAtColIndex]) row[createdAtColIndex] = now;
      if (!row[updatedAtColIndex]) row[updatedAtColIndex] = now;
      if (typeof row[isTrashedColIndex] !== 'boolean') {
        row[isTrashedColIndex] = Boolean(row[isTrashedColIndex]);
      }
    }
    if (!originalRows.size) return;
    this._writeChangedCells(data, originalRows);
    if (backfilledRows) {
      this._logSink.info('Backfilled system metadata', {
        sheet: this.sheetName,
        rows: backfilledRows,
      });
    }
    if (duplicateIdRows.size) {
      this._logSink.info('Reassigned duplicate ids', {
        sheet: this.sheetName,
        rows: duplicateIdRows.size,
      });
    }
  }
  // =========================
  // METADATA HELPERS
  // =========================
  /**
   * Return the base-0 column indexes of the system fields, -1 for any that
   * is not on the sheet.
   */
  _getSystemColumnIndexes() {
    return {
      id: this._getColumnIndex(_GAS_SHEETDB_SYSTEM_FIELDS.ID),
      createdAt: this._getColumnIndex(_GAS_SHEETDB_SYSTEM_FIELDS.CREATED_AT),
      updatedAt: this._getColumnIndex(_GAS_SHEETDB_SYSTEM_FIELDS.UPDATED_AT),
      isDeleted: this._getColumnIndex(_GAS_SHEETDB_SYSTEM_FIELDS.IS_DELETED),
    };
  }
  /**
   * Whether a row is entirely empty, as when a person cleared its contents.
   * A row with a formula is blank only while the formula evaluates to ''.
   */
  _isBlankRow(row) {
    return row.every((cell) => cell === '');
  }
  /**
   * Whether a row lacks any system metadata that
   * `_ensureRequiredMetadata` would backfill. A blank row is not an entry,
   * so it never does.
   */
  _rowNeedsBackfill(row, columns) {
    return (
      !this._isBlankRow(row) &&
      (!row[columns.id] ||
        !row[columns.createdAt] ||
        !row[columns.updatedAt] ||
        typeof row[columns.isDeleted] !== 'boolean')
    );
  }
  /**
   * Indexes of the rows whose `_id` repeats that of a row above them. The
   * topmost row keeps the id.
   */
  _findDuplicateIdRows(data, idColIndex) {
    const seen = new Set();
    const duplicates = new Set();
    data.forEach((row, rowIndex) => {
      const id = row[idColIndex];
      if (!id) return;
      if (seen.has(id)) duplicates.add(rowIndex);
      else seen.add(id);
    });
    return duplicates;
  }
  /**
   * Throw a conflict when a supplied `_id` is already stored in `body` or
   * repeats within `entries`.
   */
  _assertUniqueIds(entries, body) {
    const ids = entries.flatMap((entry) => (entry._id ? [entry._id] : []));
    if (!ids.length) return;
    const idColIndex = this._getColumnIndex(_GAS_SHEETDB_SYSTEM_FIELDS.ID);
    const taken = new Set(
      idColIndex === -1 ? [] : body.map((row) => String(row[idColIndex])),
    );
    for (const id of ids) {
      if (taken.has(String(id))) {
        throw _GasSheetDbErrors.conflict(
          `[GasSheetDb] Duplicate "${_GAS_SHEETDB_SYSTEM_FIELDS.ID}": ${id}`,
        );
      }
      taken.add(String(id));
    }
  }
  /**
   * Apply system fields for new entries.
   */
  _applyInsertMetadata(entry) {
    const now = new Date();
    const keys = _GAS_SHEETDB_SYSTEM_FIELDS;
    if (!entry[keys.ID]) {
      entry[keys.ID] = this._generateId();
    }
    if (!entry[keys.CREATED_AT]) {
      entry[keys.CREATED_AT] = now;
    }
    entry[keys.UPDATED_AT] = now;
    if (typeof entry[keys.IS_DELETED] !== 'boolean') {
      entry[keys.IS_DELETED] = false;
    }
  }
  /**
   * Update system timestamps.
   */
  _applyUpdateMetadata(entry) {
    const keys = _GAS_SHEETDB_SYSTEM_FIELDS;
    entry[keys.UPDATED_AT] = new Date();
  }
  /**
   * Return uuid.
   */
  _generateId() {
    return Utilities.getUuid();
  }
  // =========================
  // RUNTIME HELPERS
  // =========================
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
  _withLock(callback) {
    const logs = [];
    this._deferredLogs = logs;
    try {
      return this.lockService.withLock(
        this.lockScope,
        () => {
          SpreadsheetApp.flush();
          const response = callback();
          SpreadsheetApp.flush();
          return response;
        },
        { timeoutMs: this.lockTimeoutMs },
      );
    } finally {
      this._deferredLogs = null;
      logs.forEach((log) => this._emit(log));
      this._warnDuplicateColumns();
    }
  }
  /**
   * Warn about repeated header names in the last header read, once per
   * operation: called as an operation ends, never while the lock is held.
   */
  _warnDuplicateColumns() {
    // the schema is not yet assigned when the constructor creates the sheet
    const columns = this.schema?.duplicateKeys;
    if (columns?.length) {
      this._logSink.warn('Duplicate column keys', {
        sheet: this.sheetName,
        columns,
      });
    }
  }
  /**
   * Queue a log line while the lock is held, else emit it at once.
   */
  _log(log) {
    if (this._deferredLogs) this._deferredLogs.push(log);
    else this._emit(log);
  }
  /**
   * Forward a log line to the logger. A throwing logger never fails the
   * operation, which has already run, nor masks its error: the failure is
   * reported through `console` and the remaining lines are still emitted.
   */
  _emit({ level, msg, meta }) {
    try {
      this.logger[level](msg, meta);
    } catch (error) {
      console.error(`[GasSheetDb] Logger failed to record "${msg}"`, {
        error: error?.message ?? String(error),
      });
    }
  }
}
