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
 * Used when no `logger` is injected. Deliberately minimal: it writes to the
 * execution log, keeping the `[GasSheetDb]` prefix and passing meta through
 * as an object when given. It has no levels, no bindings and no sheet sink —
 * inject a GasLogger for those; that logic lives there and only there.
 */
declare const _GAS_SHEETDB_DEFAULT_LOGGER: GasSheetDbLogger;
/** Header and first data row used when `rowNumbers` is not overridden. */
declare const _GAS_SHEETDB_DEFAULT_ROW_NUMBERS: GasSheetDbRowsReference;
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
 * Used when no `lockService` is injected and the `GasLock` global is absent.
 * Deliberately minimal: it acquires a LockService lock and releases it,
 * nothing more. It does NOT reproduce gas-lock's ownership tracking. Apps
 * Script locks are owned per execution, so acquiring a scope an outer caller
 * already holds succeeds at once, but this releaseLock() then releases the
 * outer caller's lock too, ending its critical section early. Load GasLock
 * to get that safety; that logic lives there and only there.
 */
declare const _GAS_SHEETDB_FALLBACK_LOCK_SERVICE: GasSheetDbLockService;
/**
 * Used when no `lockService` is injected: delegates to the `GasLock` global
 * when loaded, else to the minimal fallback. Resolved at each acquisition
 * rather than at construction, as a GasSheetDb is often constructed at top
 * level, before gas-lock's file has been evaluated.
 */
declare const _GAS_SHEETDB_DEFAULT_LOCK_SERVICE: GasSheetDbLockService;
