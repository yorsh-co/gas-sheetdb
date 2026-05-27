/**
 * Use the spreadsheet bound to the Apps Script project.
 * When false, `SHEETDB_SPREADSHEET_URL` is used instead.
 *
 * See the library README.md "Required Apps Script Scopes" and
 * "Spreadsheet Modes" for mode details.
 */
const SHEETDB_USE_ACTIVE_SPREADSHEET = true;

/**
 * Spreadsheet URL used when `SHEETDB_USE_ACTIVE_SPREADSHEET` is false.
 */
const SHEETDB_SPREADSHEET_URL = '';

/**
 * Known sheet names used by the module.
 * Helps with autocomplete and avoids hardcoded strings.
 *
 * Example:
 * const table = SheetDb.table(SHEETDB_SHEET_NAMES.USERS);
 */
const SHEETDB_SHEET_NAMES = Object.freeze({
  SYSTEM: Object.freeze({
    CONFIG: '.config',
    ERRORS: '.errors',
  }),

  USERS: '👤 Users',
});

/**
 * Sheet row numbers (1-based).
 */
const SHEETDB_ROW_NUMBERS = {
  headers: 1,
  firstData: 2,
};

/**
 * Zero-based row indexes derived from `SHEET_ROW_NUMBERS`.
 */
const SHEETDB_ROW_INDEXES = {
  headers: null,
  firstData: null,
};

Object.entries(SHEETDB_ROW_NUMBERS).forEach(
  ([k, v]) => (SHEETDB_ROW_INDEXES[k] = v - 1),
);

/**
 * Persistent fields managed automatically by the module.
 */
const SHEETDB_SYSTEM_FIELDS = Object.freeze({
  ID: '_id',
  CREATED_AT: '_createdAt',
  UPDATED_AT: '_updatedAt',
});

/**
 * Runtime-only fields added while reading rows but not persisted in the sheet.
 */
const SHEETDB_NON_PERSISTED_FIELDS = new Set(['_runtime']);
