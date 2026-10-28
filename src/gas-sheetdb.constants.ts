/**
 * Persistent fields managed automatically by the module.
 */
const _GAS_SHEETDB_SYSTEM_FIELDS = Object.freeze({
  ID: '_id',
  CREATED_AT: '_createdAt',
  UPDATED_AT: '_updatedAt',
  IS_DELETED: '_isDeleted',
} as const);

/**
 * Runtime-only fields added while reading rows but not persisted in the sheet.
 */
const _GAS_SHEETDB_NON_PERSISTED_FIELDS: ReadonlySet<string> = new Set([
  '_runtime',
]);

/**
 * Used when no `logger` is injected. Deliberately minimal: it writes to the
 * execution log, keeping the `[GasSheetDb]` prefix and appending meta as
 * JSON so a bare console still shows it. It has no levels, no bindings and
 * no sheet sink — inject a GasLogger for those; that logic lives there and
 * only there.
 */
const _GAS_SHEETDB_DEFAULT_LOGGER: GasSheetDbLogger = (() => {
  return Object.freeze({
    info: (msg: string, meta?: Record<string, unknown>): void =>
      console.log(`[GasSheetDb] ${msg}`, meta),
    warn: (msg: string, meta?: Record<string, unknown>): void =>
      console.warn(`[GasSheetDb] ${msg}`, meta),
  });
})();

/** Scopes accepted by `lockScope`. */
const _GAS_SHEETDB_LOCK_SCOPES: readonly GasSheetDbLockScope[] = Object.freeze([
  'script',
  'document',
  'user',
]);

/**
 * 'document' locks are unavailable outside a containing-document execution
 * context (e.g. a web app's doGet/doPost) — 'script' is the one that always
 * works.
 */
const _GAS_SHEETDB_DEFAULT_LOCK_SCOPE: GasSheetDbLockScope = 'script';

const _GAS_SHEETDB_DEFAULT_LOCK_TIMEOUT_MS = 30000;

/**
 * Used when no `lockService` is injected and the `GasLock` global is absent.
 * Deliberately minimal: it acquires a LockService lock and releases it,
 * nothing more. It does NOT reproduce gas-lock's ownership tracking. Apps
 * Script locks are owned per execution, so acquiring a scope an outer caller
 * already holds succeeds at once, but this releaseLock() then releases the
 * outer caller's lock too, ending its critical section early. Load GasLock
 * to get that safety; that logic lives there and only there.
 */
const _GAS_SHEETDB_FALLBACK_LOCK_SERVICE: GasSheetDbLockService = Object.freeze(
  {
    withLock<T>(
      scope: GasSheetDbLockScope,
      callback: () => T,
      options: { timeoutMs?: number } = {},
    ): T {
      const lock =
        scope === 'document'
          ? LockService.getDocumentLock()
          : scope === 'user'
            ? LockService.getUserLock()
            : LockService.getScriptLock();

      if (!lock) {
        throw new Error(
          `[GasSheetDb] LockService returned no lock for scope "${scope}" — "document" locks are unavailable outside the context of a containing document (e.g. a web app execution). Use "script" or "user", or inject a lockService such as GasLock.`,
        );
      }

      lock.waitLock(options.timeoutMs ?? _GAS_SHEETDB_DEFAULT_LOCK_TIMEOUT_MS);
      try {
        return callback();
      } finally {
        lock.releaseLock();
      }
    },
  },
);

/**
 * Used when no `lockService` is injected: delegates to the `GasLock` global
 * when loaded, else to the minimal fallback. Resolved at each acquisition
 * rather than at construction, as a GasSheetDb is often constructed at top
 * level, before gas-lock's file has been evaluated.
 */
const _GAS_SHEETDB_DEFAULT_LOCK_SERVICE: GasSheetDbLockService = Object.freeze({
  withLock<T>(
    scope: GasSheetDbLockScope,
    callback: () => T,
    options?: { timeoutMs?: number },
  ): T {
    let service = _GAS_SHEETDB_FALLBACK_LOCK_SERVICE;

    try {
      if (
        typeof GasLock !== 'undefined' &&
        typeof GasLock.withLock === 'function'
      ) {
        service = GasLock;
      }
    } catch {
      // typeof on a binding not yet initialised throws a ReferenceError
    }

    return service.withLock(scope, callback, options);
  },
});
