/**
 * Ambient declaration for gas-lock, the minimum gas-sheetdb needs to detect
 * and call it.
 *
 * gas-lock is expected as a runtime global (Apps Script has no module
 * system) and ships no canonical peer types file yet, so this declares only
 * the `withLock` surface gas-sheetdb uses. Replace it with gas-lock's peer
 * file once one exists.
 *
 * Being a `.d.ts`, this file is type-only: it matches `include:
 * ["src/**\/*.ts"]` but emits nothing, so no counterpart reaches dist/ or
 * the published package.
 *
 * @see https://github.com/yorsh-co/gas-lock
 */
declare const GasLock: GasSheetDbLockService | undefined;
