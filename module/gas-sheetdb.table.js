/**
 * Table-style wrapper around a Google Spreadsheet sheet.
 *
 * Provides object-based querying, inserts, updates,
 * automatic column management, and metadata handling.
 */
class _GasSheetDbTable {
  /**
   * @param {Object} options
   * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} options.spreadsheet
   * @param {string|null} [options.sheetName]
   * @param {_GasSheetDbRowsReference} [options.rowNumbers]
   */
  constructor({ spreadsheet, sheetName, rowNumbers } = {}) {
    /** @type {GoogleAppsScript.Spreadsheet.Spreadsheet} */
    this.spreadsheet = spreadsheet;

    /** @type {string} */
    this.sheetName = sheetName;

    /** @type {_GasSheetDbRowsReference} */
    this.rowNumbers = rowNumbers; // Base-1 row numbers

    /** @type {_GasSheetDbRowsReference} */
    this.rowIndexes = this._deriveRowIndexes(); // Base-0 row indexes

    /** @type {GoogleAppsScript.Spreadsheet.Sheet} */
    this.sheet = this._ensureSheet();

    /** @type {_GasSheetDbTableSchema} */
    this.schema = new _GasSheetDbTableSchema(this.sheet, this.rowNumbers);
  }

  // =========================
  // CONSTRUCTOR HELPERS
  // =========================

  /**
   * Derive base-0 data array row indexes from
   * base-1 spreadsheet row numbers.
   * @returns {_GasSheetDbRowsReference}
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
   * sheet, or inserting a new one.
   *
   * @returns {GoogleAppsScript.SpreadsheetApp.Sheet}
   */
  _ensureSheet() {
    if (this.sheetName) {
      const sheet = this.spreadsheet.getSheetByName(this.sheetName);

      if (sheet) return sheet;
    }

    return this._insertSheet();
  }

  /**
   * Insert a new sheet to the spreadsheet, using the provided
   * name. Applies basic formatting including alternating
   * row colors and a frozen header row.
   *
   * @returns {GoogleAppsScript.SpreadsheetApp.Sheet}
   */
  _insertSheet() {
    let newSheetName = this.sheetName;

    if (!newSheetName) {
      const allSheetNames = this.spreadsheet
        .getSheets()
        .map((s) => s.getSheetName());

      const defaultSheetName = 'gsd-table-{i}';

      let i = 0;
      do {
        newSheetName = defaultSheetName.replace('{i}', i);
      } while (allSheetNames.includes(newSheetName));
    }

    this.sheetName = newSheetName;

    const newSheet = this.spreadsheet.insertSheet(newSheetName);

    const maxCols = newSheet.getMaxColumns();

    newSheet
      .getRange(this.rowNumbers.columnKeys, 1, newSheet.getMaxRows(), maxCols)
      .applyRowBanding();

    newSheet
      .getRange(this.rowNumbers.columnKeys, 1, 1, maxCols)
      .setFontWeight('bold');

    newSheet.setFrozenRows(this.rowNumbers.columnKeys);

    Logger.log('[GasSheetDb] Created new table sheet "%s"', newSheetName);

    return newSheet;
  }

  // =========================
  // QUERY
  // =========================

  /**
   * Read all rows as objects.
   * Adds runtime metadata to each entry.
   *
   * @returns {Object[]}
   */
  find() {
    return this._withLock(() => {
      this._ensureRequiredMetadata();

      const columnKeys = this.schema.columnKeys;

      const data = this._getTableBodyData();

      const entries = [];

      for (let dataIndex = 0; dataIndex < data.length; dataIndex++) {
        const entry = {};

        const row = data[dataIndex];

        columnKeys.forEach((key, colIndex) => {
          entry[key] = _GasSheetDbValueCodec.decode(row[colIndex]);
        });

        entries.push(entry);
      }

      return entries;
    });
  }

  /**
   * Filter entries using a predicate.
   *
   * @param {(entry: Object) => boolean} predicateFn
   * @returns {Object[]}
   */
  findWhere(predicateFn) {
    return this.find().filter(predicateFn);
  }

  /**
   * Find the first matching entry.
   *
   * @param {(entry: Object) => boolean} predicateFn
   * @returns {Object|null}
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
   * @param {Object} entry
   */
  insert(entry) {
    this.insertMany([entry]);
  }

  /**
   * Insert multiple entries.
   * Missing columns are created automatically.
   *
   * @param {Object[]} entries - Mutated in place: `_id`, `_createdAt`,
   *   and `_updatedAt` are added if not already present.
   */
  insertMany(entries) {
    this._withLock(() => {
      entries.forEach((entry) => this._applyInsertMetadata(entry));

      this.schema.reload();

      const columnKeys = this._extractKeys(entries);

      this.schema.ensureColumns(columnKeys);

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
   * @param {Object} entry
   */
  update(entry) {
    this.updateMany([entry]);
  }

  /**
   * Update multiple existing entries.
   * Requires `_id` for each entry.
   *
   * @param {Object[]} entries - Mutated in place: `_updatedAt` is
   *   overwritten with the current time.
   */
  updateMany(entries) {
    this._withLock(() => {
      // add new property columns
      this.schema.reload();

      const columnKeys = this._extractKeys(entries);

      this.schema.ensureColumns(columnKeys);
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
   * @param {Object} entry
   */
  delete(entry) {
    this.deleteMany([entry]);
  }

  /**
   * Delete multiple existing entries.
   * Requires `_id` for each entry.
   *
   * @param {Object[]} entries
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
  // TABLE BODY HELPERS
  // =========================

  /**
   * Return the table body, without the column keys row.
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

    const idColIndex = this._getColumnIndex(_GAS_SHEETDB_SYSTEM_FIELDS.ID);

    if (idColIndex === -1) {
      throw new Error(
        `Cannot map rows by "${_GAS_SHEETDB_SYSTEM_FIELDS.ID}": column does not exist on sheet "${this.sheetName}". Call find() at least once before update/delete.`,
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
    return this.schema.columnKeys.indexOf(key);
  }

  /**
   * Extract the column keys from an entry's property keys.
   *
   * @param {Object[]} entries
   * @returns {string[]}
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
   * Check if a field should be stored in the sheet.
   *
   * @param {string} key
   * @returns {boolean}
   */
  _isPersistedField(key) {
    return !_GAS_SHEETDB_NON_PERSISTED_FIELDS.has(key);
  }

  /**
   * Build a sheet row from an entry.
   * Existing values are preserved when a field is undefined.
   *
   * @param {Object} entry
   * @param {Array} existingRow
   * @returns {Array}
   */
  _buildRow(entry, existingRow = []) {
    return this.schema.columnKeys.map((key, index) => {
      const value = entry[key];

      if (value === undefined) {
        return existingRow[index] ?? '';
      }

      return _GasSheetDbValueCodec.encode(value);
    });
  }

  /**
   * Ensure in-sheet insert metadata for all entries.
   * This ensures that entries inserted manually or by processes
   * other than table.insert() contain the required metadata for
   * reading and writing with SheetDB.
   */
  _ensureRequiredMetadata() {
    this.schema.ensureColumns(Object.values(_GAS_SHEETDB_SYSTEM_FIELDS));
    this.schema.reload();

    const data = this._getTableBodyData();

    if (!data.length) return;

    const idColIndex = this._getColumnIndex(_GAS_SHEETDB_SYSTEM_FIELDS.ID);
    const createdAtColIndex = this._getColumnIndex(
      _GAS_SHEETDB_SYSTEM_FIELDS.CREATED_AT,
    );
    const updatedAtColIndex = this._getColumnIndex(
      _GAS_SHEETDB_SYSTEM_FIELDS.UPDATED_AT,
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
   * @param {Object} entry
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
  }

  /**
   * Update system timestamps.
   *
   * @param {Object} entry
   */
  _applyUpdateMetadata(entry) {
    const keys = _GAS_SHEETDB_SYSTEM_FIELDS;

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
