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
      this.sheet = spreadsheet.insertSheet(sheetName);

      Logger.log('[_SheetDbTable] Created missing sheet "%s"', sheetName);
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
    const data = this.sheet.getDataRange().getValues();

    if (!data.length) {
      return [];
    }

    const headers = this.schema.headers;

    const bodyOffset = this.rowIndexes.firstData;
    const body = data.slice(bodyOffset);

    const entries = [];

    for (let dataIndex = 0; dataIndex < body.length; dataIndex++) {
      const entry = {};

      const row = body[dataIndex];

      headers.forEach((header, colIndex) => {
        entry[header] = _SheetDbValueCodec.decode(row[colIndex]);
      });

      if (!entry._runtime) {
        entry._runtime = {};
      }

      entry._runtime.rowNumber =
        dataIndex + bodyOffset + (this.rowNumbers.firstData - 1);

      entries.push(entry);
    }

    return entries;
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
   * @param {object[]} entries
   */
  insertMany(entries) {
    this._withLock(() => {
      entries.forEach((entry) => this._applyInsertMetadata(entry));

      this.schema.reload();

      const headers = this._extractHeaders(entries);

      this.schema.ensureColumns(headers);

      this.schema.reload();

      const rows = entries.map((entry) => this._buildRow(entry));

      this._ensureBlankRows(rows.length);

      this.sheet
        .getRange(this.sheet.getLastRow() + 1, 1, rows.length, rows[0].length)
        .setValues(rows);
    });
  }

  // =========================
  // UPDATE
  // =========================

  /**
   * Update a single entry.
   * Requires `_id` and `_runtime.rowNumber`.
   *
   * @param {object} entry
   */
  update(entry) {
    this.updateMany([entry]);
  }

  /**
   * Update multiple existing entries.
   * Requires `_id` and `_runtime.rowNumber` for each entry.
   *
   * @param {object[]} entries
   */
  updateMany(entries) {
    this._withLock(() => {
      this.schema.reload();

      for (const entry of entries) {
        if (!entry._id) {
          throw new Error('Cannot update entry without "_id"');
        }

        this._applyUpdateMetadata(entry);

        const rowNumber = entry?._runtime?.rowNumber;

        if (!rowNumber) {
          throw new Error('Cannot update entry without "_runtime.rowNumber"');
        }

        const headers = this._extractHeaders([entry]);

        this.schema.ensureColumns(headers);

        this.schema.reload();

        const existingRow = this.sheet
          .getRange(rowNumber, 1, 1, this.schema.headers.length)
          .getValues()[0];

        const updatedRow = this._buildRow(entry, existingRow);

        this.sheet
          .getRange(rowNumber, 1, 1, updatedRow.length)
          .setValues([updatedRow]);
      }
    });
  }

  // =========================
  // HELPERS
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

  /**
   * Apply system fields for new entries.
   *
   * @param {object} entry
   */
  _applyInsertMetadata(entry) {
    const now = new Date();

    const keys = SHEETDB_SYSTEM_FIELDS;

    if (!entry[keys.ID]) {
      entry[keys.ID] = Utilities.getUuid();
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
}
