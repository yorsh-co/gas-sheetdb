/**
 * Persistent fields managed automatically by the module.
 */
declare const _GAS_SHEETDB_SYSTEM_FIELDS: Readonly<{
    readonly ID: "_id";
    readonly CREATED_AT: "_createdAt";
    readonly UPDATED_AT: "_updatedAt";
    readonly IS_DELETED: "_isDeleted";
}>;
/**
 * Runtime-only fields added while reading rows but not persisted in the sheet.
 */
declare const _GAS_SHEETDB_NON_PERSISTED_FIELDS: ReadonlySet<string>;
