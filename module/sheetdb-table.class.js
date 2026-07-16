/**
 * Table-style wrapper around a Google Spreadsheet sheet.
 *
 * Provides object-based querying, inserts, updates,
 * automatic column management, and metadata handling.
 */
class _SheetDbTable {
  /**
   * @param {object} options
   * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} options.spreadsheet
   * @param {string} options.sheetName
   */
  constructor({ spreadsheet, sheetName } = {}) {
    this.spreadsheet = spreadsheet;
    this.sheetName = sheetName;

    this.rowNumbers = SHEETDB_ROW_NUMBERS;
    this.rowIndexes = SHEETDB_ROW_INDEXES;

    this.sheet = spreadsheet.getSheetByName(sheetName);

    if (!this.sheet) {
      this._insertSheet(sheetName);
    }

    this.schema = new _SheetDbTableSchema(this.sheet, this.rowNumbers);
  }

  // =========================
  // QUERY
  // =========================

  /**
   * Read all rows as objects.
   * Adds runtime metadata to each entry.
   *
   * @returns {object[]}
   */
  find() {
    this._withLock(() => {
      this._ensureRequiredMetadata();

      const headers = this.schema.headers;

      const data = this._getTableBodyData();

      const entries = [];

      for (let dataIndex = 0; dataIndex < data.length; dataIndex++) {
        const entry = {};

        const row = data[dataIndex];

        headers.forEach((header, colIndex) => {
          entry[header] = _SheetDbValueCodec.decode(row[colIndex]);
        });

        entries.push(entry);
      }

      return entries;
    });
  }

  /**
   * Filter entries using a predicate.
   *
   * @param {(entry: object) => boolean} predicateFn
   * @returns {object[]}
   */
  findWhere(predicateFn) {
    return this.find().filter(predicateFn);
  }

  /**
   * Find the first matching entry.
   *
   * @param {(entry: object) => boolean} predicateFn
   * @returns {object|null}
   */
  findOneWhere(predicateFn) {
    return this.find().find(predicateFn) || null;
  }

  // =========================
  // INSERT
  // =========================

  /**
   * Insert a single entry.
   *
   * @param {object} entry
   */
  insert(entry) {
    this.insertMany([entry]);
  }

  /**
   * Insert multiple entries.
   * Missing columns are created automatically.
   *
   * @param {object[]} entries - Mutated in place: `_id`, `_createdAt`,
   *   and `_updatedAt` are added if not already present.
   */
  insertMany(entries) {
    this._withLock(() => {
      entries.forEach((entry) => this._applyInsertMetadata(entry));

      this.schema.reload();

      const headers = this._extractHeaders(entries);

      this.schema.ensureColumns(headers);

      this.schema.reload();

      const data = entries.map((entry) => this._buildRow(entry));

      this._appendTableBodyData(data);
    });
  }

  // =========================
  // UPDATE
  // =========================

  /**
   * Update a single entry.
   * Requires `_id`.
   *
   * @param {object} entry
   */
  update(entry) {
    this.updateMany([entry]);
  }

  /**
   * Update multiple existing entries.
   * Requires `_id` for each entry.
   *
   * @param {object[]} entries - Mutated in place: `_updatedAt` is
   *   overwritten with the current time.
   */
  updateMany(entries) {
    this._withLock(() => {
      // add new property columns
      this.schema.reload();

      const headers = this._extractHeaders(entries);

      this.schema.ensureColumns(headers);
      this.schema.reload();

      // load the current data after schema reload
      const data = this._getTableBodyData();

      // map the row indexes
      const rowIndexesById = this._mapTableBodyRowIndexesById(data);

      for (const entry of entries) {
        if (!entry._id) {
          throw new Error('Cannot update entry without "_id"');
        }

        const rowIndex = rowIndexesById.get(entry._id);

        if (rowIndex === undefined) {
          throw new Error(`Entry not found: ${entry._id}`);
        }

        this._applyUpdateMetadata(entry);

        data[rowIndex] = this._buildRow(entry, data[rowIndex]);
      }

      this._setTableBodyData(data);
    });
  }

  // =========================
  // DELETE
  // =========================

  /**
   * Delete a single entry.
   * Requires `_id`.
   *
   * @param {object} entry
   */
  delete(entry) {
    this.deleteMany([entry]);
  }

  /**
   * Delete multiple existing entries.
   * Requires `_id` for each entry.
   *
   * @param {object[]} entries
   */
  deleteMany(entries) {
    this._withLock(() => {
      // load the current data after schema reload
      const data = this._getTableBodyData();

      // map the row indexes
      const rowIndexesById = this._mapTableBodyRowIndexesById(data);

      const entriesRowIndexes = [];
      for (const entry of entries) {
        if (!entry._id) {
          throw new Error('Cannot update entry without "_id"');
        }

        const rowIndex = rowIndexesById.get(entry._id);

        if (rowIndex === undefined) {
          throw new Error(`Entry not found: ${entry._id}`);
        }

        entriesRowIndexes.push(rowIndex);
      }

      // sort in reverse order to delete from the bottom up
      entriesRowIndexes.sort((a, b) => b - a);

      entriesRowIndexes.forEach((rowIndex) => data.splice(rowIndex, 1));

      this._setTableBodyData(data);
    });
  }

  // =========================
  // SHEET HELPERS
  // =========================
  /**
   *
   * @param {string} sheetName
   */
  _insertSheet(sheetName) {
    this.sheet = this.spreadsheet.insertSheet(sheetName);

    const maxCols = this.sheet.getMaxColumns();

    this.sheet
      .getRange(this.rowNumbers.headers, 1, this.sheet.getMaxRows(), maxCols)
      .applyRowBanding();

    this.sheet
      .getRange(this.rowNumbers.headers, 1, 1, maxCols)
      .setFontWeight('bold');

    this.sheet.setFrozenRows(this.rowNumbers.headers);

    Logger.log('[SheetDb] Created missing sheet "%s"', sheetName);
  }

  // =========================
  // TABLE BODY HELPERS
  // =========================

  /**
   * Return the table body, without the headers.
   *
   * @returns {*[][]}
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
   * Write to the entire table body on the sheet.
   *
   * @param {*[][]} data
   */
  _setTableBodyData(data = []) {
    if (data.length) {
      const tableBodyRange = this.sheet.getRange(
        this.rowNumbers.firstData,
        1,
        data.length,
        data[0].length,
      );

      tableBodyRange.setValues(data);
    }

    this._removeExtraRows(data.length);

    SpreadsheetApp.flush();
  }

  /**
   * Append rows to the sheet.
   *
   * @param {*[][]} data
   */
  _appendTableBodyData(data) {
    if (!data.length) throw new Error('Empty data');

    this._ensureBlankRows(data.length);

    this.sheet
      .getRange(this.sheet.getLastRow() + 1, 1, data.length, data[0].length)
      .setValues(data);

    SpreadsheetApp.flush();
  }

  /**
   * Ensure the sheet has enough empty rows.
   *
   * @param {number} amount
   */
  _ensureBlankRows(amount) {
    const maxRows = this.sheet.getMaxRows();

    const lastRow = this.sheet.getLastRow();

    const rowsNeeded = amount + lastRow - maxRows;

    if (rowsNeeded > 0) {
      this.sheet.insertRowsAfter(lastRow, rowsNeeded);
    }
  }

  /**
   * Remove extra rows from the sheet.
   *
   * @param {number} tableBodyLen
   */
  _removeExtraRows(tableBodyLen) {
    const lastTableBodyRow = this.rowNumbers.firstData - 1 + tableBodyLen;

    const maxRows = this.sheet.getMaxRows();

    const numExtraRows = maxRows - lastTableBodyRow;

    if (numExtraRows > 0) {
      this.sheet.deleteRows(lastTableBodyRow + 1, numExtraRows);
    }
  }

  /**
   * Map the row indexes
   *
   * @param {*[][]} [tableBody=null]
   * @returns {Map<string, number>}
   */
  _mapTableBodyRowIndexesById(tableBody = null) {
    const data = tableBody || this._getTableBodyData();

    const idColIndex = this._getColumnIndex(SHEETDB_SYSTEM_FIELDS.ID);

    if (idColIndex === -1) {
      throw new Error(
        `Cannot map rows by "${SHEETDB_SYSTEM_FIELDS.ID}": column does not exist on sheet "${this.sheetName}". Call find() at least once before update/delete.`,
      );
    }

    const rowMap = new Map();

    for (let i = 0; i < data.length; i++) {
      rowMap.set(data[i][idColIndex], i);
    }

    return rowMap;
  }

  // =========================
  // TABLE COLUMN HELPERS
  // =========================

  /**
   * Return the base-0 index for column corresponding to the provided key.
   *
   * @param {string} key
   * @returns {number}
   */
  _getColumnIndex(key) {
    return this.schema.headers.indexOf(key);
  }

  /**
   * Extract the column headers from an entry's property keys.
   *
   * @param {object[]} entries
   * @returns {string[]}
   */
  _extractHeaders(entries) {
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
   * Check if a field should be stored in the sheet.
   *
   * @param {string} key
   * @returns {boolean}
   */
  _isPersistedField(key) {
    return !SHEETDB_NON_PERSISTED_FIELDS.has(key);
  }

  /**
   * Build a sheet row from an entry.
   * Existing values are preserved when a field is undefined.
   *
   * @param {object} entry
   * @param {Array} existingRow
   * @returns {Array}
   */
  _buildRow(entry, existingRow = []) {
    return this.schema.headers.map((header, index) => {
      const value = entry[header];

      if (value === undefined) {
        return existingRow[index] ?? '';
      }

      return _SheetDbValueCodec.encode(value);
    });
  }

  /**
   * Ensure in-sheet insert metadata for all entries.
   * This ensures that entries inserted manually or by processes
   * other than table.insert() contain the required metadata for
   * reading and writing with SheetDB.
   */
  _ensureRequiredMetadata() {
    this.schema.ensureColumns(Object.values(SHEETDB_SYSTEM_FIELDS));
    this.schema.reload();

    const data = this._getTableBodyData();

    if (!data.length) return;

    const idColIndex = this._getColumnIndex(SHEETDB_SYSTEM_FIELDS.ID);
    const createdAtColIndex = this._getColumnIndex(
      SHEETDB_SYSTEM_FIELDS.CREATED_AT,
    );
    const updatedAtColIndex = this._getColumnIndex(
      SHEETDB_SYSTEM_FIELDS.UPDATED_AT,
    );

    const now = new Date();

    for (let i = 0; i < data.length; i++) {
      if (!data[i][idColIndex]) data[i][idColIndex] = this._generateId();

      if (!data[i][createdAtColIndex]) data[i][createdAtColIndex] = now;

      if (!data[i][updatedAtColIndex]) data[i][updatedAtColIndex] = now;
    }

    this._setTableBodyData(data);
  }

  // =========================
  // METADATA HELPERS
  // =========================

  /**
   * Apply system fields for new entries.
   *
   * @param {object} entry
   */
  _applyInsertMetadata(entry) {
    const now = new Date();

    const keys = SHEETDB_SYSTEM_FIELDS;

    if (!entry[keys.ID]) {
      entry[keys.ID] = this._generateId();
    }

    if (!entry[keys.CREATED_AT]) {
      entry[keys.CREATED_AT] = now;
    }

    entry[keys.UPDATED_AT] = now;
  }

  /**
   * Update system timestamps.
   *
   * @param {object} entry
   */
  _applyUpdateMetadata(entry) {
    const keys = SHEETDB_SYSTEM_FIELDS;

    entry[keys.UPDATED_AT] = new Date();
  }

  /**
   * Return uuid.
   *
   * @returns {string}
   */
  _generateId() {
    return Utilities.getUuid();
  }

  // =========================
  // RUNTIME HELPERS
  // =========================

  /**
   * Execute a callback inside a document lock.
   *
   * @param {Function} callback
   * @returns {*}
   */
  _withLock(callback) {
    const lock = LockService.getDocumentLock();

    lock.waitLock(30000);

    try {
      return callback();
    } finally {
      lock.releaseLock();
    }
  }
}
