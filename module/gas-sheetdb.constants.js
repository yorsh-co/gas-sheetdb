/**
 * Persistent fields managed automatically by the module.
 */
const _GAS_SHEETDB_SYSTEM_FIELDS = Object.freeze({
  ID: '_id',
  CREATED_AT: '_createdAt',
  UPDATED_AT: '_updatedAt',
});

/**
 * Runtime-only fields added while reading rows but not persisted in the sheet.
 */
const _GAS_SHEETDB_NON_PERSISTED_FIELDS = new Set(['_runtime']);
