/**
 * Persistent fields managed automatically by the module.
 */
declare const _GAS_SHEETDB_SYSTEM_FIELDS: Readonly<{
  readonly ID: '_id';
  readonly CREATED_AT: '_createdAt';
  readonly UPDATED_AT: '_updatedAt';
  readonly IS_DELETED: '_isDeleted';
}>;
/**
 * Runtime-only fields added while reading rows but not persisted in the sheet.
 */
declare const _GAS_SHEETDB_NON_PERSISTED_FIELDS: ReadonlySet<string>;
/** Scopes accepted by `lockScope`. */
declare const _GAS_SHEETDB_LOCK_SCOPES: readonly GasSheetDbLockScope[];
/**
 * 'document' locks are unavailable outside a containing-document execution
 * context (e.g. a web app's doGet/doPost) — 'script' is the one that always
 * works.
 */
declare const _GAS_SHEETDB_DEFAULT_LOCK_SCOPE: GasSheetDbLockScope;
declare const _GAS_SHEETDB_DEFAULT_LOCK_TIMEOUT_MS = 30000;
/**
 * Used when no `lockService` is injected. Deliberately minimal: it acquires
 * a LockService lock and releases it, nothing more. It does NOT reproduce
 * gas-lock's reentrancy tracking — a nested call for a scope an outer caller
 * already holds will block until timeout, exactly as raw LockService always
 * has. Inject GasLock to get reentrancy safety; that logic lives there and
 * only there.
 */
declare const _GAS_SHEETDB_DEFAULT_LOCK_SERVICE: GasSheetDbLockService;
