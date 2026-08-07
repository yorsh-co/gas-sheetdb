'use strict';
/**
 * Persistent fields managed automatically by the module.
 */
const _GAS_SHEETDB_SYSTEM_FIELDS = Object.freeze({
  ID: '_id',
  CREATED_AT: '_createdAt',
  UPDATED_AT: '_updatedAt',
  IS_DELETED: '_isDeleted',
});
/**
 * Runtime-only fields added while reading rows but not persisted in the sheet.
 */
const _GAS_SHEETDB_NON_PERSISTED_FIELDS = new Set(['_runtime']);
/**
 * Used when no `logger` is injected. Deliberately minimal: it writes to the
 * execution log, keeping the `[GasSheetDb]` prefix and appending meta as
 * JSON so a bare console still shows it. It has no levels, no bindings and
 * no sheet sink — inject a GasLogger for those; that logic lives there and
 * only there.
 */
const _GAS_SHEETDB_DEFAULT_LOGGER = (() => {
  return Object.freeze({
    info: (msg, meta) => console.log(`[GasSheetDb] ${msg}`, meta),
    warn: (msg, meta) => console.warn(`[GasSheetDb] ${msg}`, meta),
  });
})();
/** Scopes accepted by `lockScope`. */
const _GAS_SHEETDB_LOCK_SCOPES = Object.freeze(['script', 'document', 'user']);
/**
 * 'document' locks are unavailable outside a containing-document execution
 * context (e.g. a web app's doGet/doPost) — 'script' is the one that always
 * works.
 */
const _GAS_SHEETDB_DEFAULT_LOCK_SCOPE = 'script';
const _GAS_SHEETDB_DEFAULT_LOCK_TIMEOUT_MS = 30000;
/**
 * Used when no `lockService` is injected. Deliberately minimal: it acquires
 * a LockService lock and releases it, nothing more. It does NOT reproduce
 * gas-lock's reentrancy tracking — a nested call for a scope an outer caller
 * already holds will block until timeout, exactly as raw LockService always
 * has. Inject GasLock to get reentrancy safety; that logic lives there and
 * only there.
 */
const _GAS_SHEETDB_DEFAULT_LOCK_SERVICE = Object.freeze({
  withLock(scope, callback, options = {}) {
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
});
