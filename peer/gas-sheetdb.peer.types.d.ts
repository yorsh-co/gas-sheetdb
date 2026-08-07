/**
 * Ambient declarations for gas-sheetdb, consumed as a peer subtree package.
 *
 * gas-sheetdb is expected as a runtime global (Apps Script has no module
 * system); this file declares its public surface so a consuming package can
 * typecheck standalone via `tsc --noEmit` without vendoring the source.
 *
 * Canonical copy: gas-sheetdb/peer/gas-sheetdb.peer.types.d.ts. Do not edit
 * downstream copies — update the canonical file and re-copy, so every peer
 * stays in step with the published surface.
 *
 * Being a `.d.ts`, this file is type-only: it matches a consumer's
 * `include: ["src/**\/*.ts"]` but emits nothing, so no counterpart reaches
 * dist/ or the published package.
 *
 * Usage: copy into the consuming package as `src/internal/`.
 *
 * @see https://github.com/yorsh-co/gas-sheetdb
 * @version 1.3.0
 */

/** Row/column position config for a table's header and data rows. */
interface GasSheetDbRowsReference {
  columnKeys: number;
  firstData: number;
}

/** A stored entry: persisted system metadata. */
interface GasSheetDbSystemFields {
  _id?: string;
  _createdAt?: Date;
  _updatedAt?: Date;
  _isDeleted?: boolean;
}

/** A stored entry: arbitrary user properties plus persisted system metadata. */
type GasSheetDbEntry = GasSheetDbSystemFields & Record<string, unknown>;

type GasSheetDbLockScope = 'script' | 'document' | 'user';

/**
 * Minimal lock-service shape gas-sheetdb needs to guard its writes.
 * GasLock (github.com/yorsh-co/gas-lock) satisfies this directly — pass
 * it in to get reentrancy-safe nesting with your own LockService calls.
 * Omit it and gas-sheetdb falls back to acquiring Apps Script's
 * LockService directly, with no reentrancy protection.
 */
interface GasSheetDbLockService {
  withLock<T>(
    scope: GasSheetDbLockScope,
    callback: () => T,
    options?: { timeoutMs?: number },
  ): T;
}

/**
 * Minimal logging surface gas-sheetdb needs. GasLogger
 * (github.com/yorsh-co/gas-logger) satisfies it directly — pass an instance,
 * or a `child()` of one, and gas-sheetdb's log lines inherit whatever sinks,
 * levels and bindings that logger is configured with. Plain `console`
 * satisfies it too. Omit it and gas-sheetdb writes to the execution log.
 *
 * Only `info` and `warn` are declared: every failure gas-sheetdb cannot
 * recover from is thrown rather than logged, so there is nothing to report
 * at `error`.
 *
 * Flushing a buffered logger is the caller's responsibility — gas-sheetdb
 * owns no execution boundary to flush on.
 */
interface GasSheetDbLogger {
  info(msg: string, meta?: Record<string, unknown>): void;
  warn(msg: string, meta?: Record<string, unknown>): void;
}

/** Constructor options for `GasSheetDb`. */
interface GasSheetDbOptions {
  spreadsheet?: GoogleAppsScript.Spreadsheet.Spreadsheet;
  spreadsheetUrl?: string;
  spreadsheetId?: string;
  useActiveSpreadsheet?: boolean;
  rowNumbers?: Partial<GasSheetDbRowsReference>;
  /** @default 'script' */
  lockScope?: GasSheetDbLockScope;
  /** @default an internal fallback calling LockService directly, with no reentrancy protection */
  lockService?: GasSheetDbLockService;
  /** @default an internal fallback writing to the execution log via console */
  logger?: GasSheetDbLogger;
}

/** Options for `GasSheetDb.table()`. */
interface GasSheetDbTableOptions {
  sheetName?: string | null;
  rowNumbers?: Partial<GasSheetDbRowsReference>;
  /** Overrides the parent `GasSheetDb` instance's lockScope for this table only. */
  lockScope?: GasSheetDbLockScope;
  /** Overrides the parent `GasSheetDb` instance's logger for this table only. */
  logger?: GasSheetDbLogger;
}

/** Matches entries in the predicate-based query, update and delete methods. */
type GasSheetDbPredicate = (entry: GasSheetDbEntry) => boolean;

/**
 * The change `updateWhere()`/`updateOneWhere()` applies to each matched
 * entry: either a patch object applied to every match, or a function
 * returning a patch for a given entry. A function that mutates its entry in
 * place and returns nothing is honoured too.
 */
type GasSheetDbUpdate =
  GasSheetDbEntry | ((entry: GasSheetDbEntry) => GasSheetDbEntry | void);

/** Options for `_GasSheetDbTable#find()`. */
interface GasSheetDbFindOptions {
  withTrashed?: boolean;
  onlyTrashed?: boolean;
}

/**
 * Manages sheet column keys and schema updates.
 *
 * Reachable as `table.schema`; instances are created by `_GasSheetDbTable`,
 * never by consumers.
 */
declare class _GasSheetDbTableSchema {
  private constructor();

  sheet: GoogleAppsScript.Spreadsheet.Sheet;
  rowNumbers: GasSheetDbRowsReference;
  logger: GasSheetDbLogger;
  /** Column keys as currently loaded from the sheet's header row. */
  columnKeys: string[];

  /** Load column key values from the sheet. */
  loadColumnKeys(): string[];

  /** Reload the column key values from the sheet. */
  reload(): void;

  /**
   * Ensure all column keys exist in the sheet.
   * Missing columns are appended automatically.
   */
  ensureColumns(columnKeys: string[]): void;
}

/**
 * Table-style wrapper around a Google Spreadsheet sheet.
 *
 * Provides object-based querying, inserts, updates, automatic column
 * management, and metadata handling.
 *
 * Instances come from `GasSheetDb#table()` — the constructor takes already
 * resolved collaborators and is not part of the public surface.
 */
declare class _GasSheetDbTable {
  private constructor();

  spreadsheet: GoogleAppsScript.Spreadsheet.Spreadsheet;
  sheetName: string | null;
  /** Base-1 spreadsheet row numbers. */
  rowNumbers: GasSheetDbRowsReference;
  /** Base-0 data array row indexes, derived from `rowNumbers`. */
  rowIndexes: GasSheetDbRowsReference;
  sheet: GoogleAppsScript.Spreadsheet.Sheet;
  schema: _GasSheetDbTableSchema;
  lockScope: GasSheetDbLockScope;
  lockService: GasSheetDbLockService;
  logger: GasSheetDbLogger;

  /**
   * Read all rows as objects.
   * Adds runtime metadata to each entry.
   */
  find(options?: GasSheetDbFindOptions): GasSheetDbEntry[];

  /**
   * Filter entries using a predicate.
   *
   * Returns an empty array when nothing matches.
   */
  findWhere(predicateFn: GasSheetDbPredicate): GasSheetDbEntry[];

  /** Find the first matching entry. */
  findOneWhere(predicateFn: GasSheetDbPredicate): GasSheetDbEntry | null;

  /**
   * Find trashed entries.
   * Optionally, filter trashed entries.
   */
  findTrashed(predicateFn?: GasSheetDbPredicate | null): GasSheetDbEntry[];

  /** Insert a single entry. */
  insert(entry: GasSheetDbEntry): GasSheetDbEntry;

  /**
   * Insert multiple entries.
   * Missing columns are created automatically.
   */
  insertMany(entries: GasSheetDbEntry[]): GasSheetDbEntry[];

  /**
   * Update a single entry.
   * Requires `_id`.
   */
  update(entry: GasSheetDbEntry): GasSheetDbEntry;

  /**
   * Update multiple existing entries.
   * Requires `_id` for each entry.
   */
  updateMany(entries: GasSheetDbEntry[]): GasSheetDbEntry[];

  /**
   * Update every entry matching a predicate, holding a single lock across
   * both the read and the write.
   *
   * Returns the updated entries, or an empty array when nothing matched.
   */
  updateWhere(
    predicateFn: GasSheetDbPredicate,
    update: GasSheetDbUpdate,
    options?: GasSheetDbFindOptions,
  ): GasSheetDbEntry[];

  /**
   * Update the first entry matching a predicate, holding a single lock
   * across both the read and the write.
   *
   * Returns the updated entry, or `null` when nothing matched.
   */
  updateOneWhere(
    predicateFn: GasSheetDbPredicate,
    update: GasSheetDbUpdate,
    options?: GasSheetDbFindOptions,
  ): GasSheetDbEntry | null;

  /**
   * Soft delete a single entry.
   * Requires `_id`.
   */
  softDelete(entry: GasSheetDbEntry): GasSheetDbEntry;

  /**
   * Soft delete multiple existing entries.
   * Requires `_id` for each entry.
   */
  softDeleteMany(entries: GasSheetDbEntry[]): GasSheetDbEntry[];

  /**
   * Soft delete every entry matching a predicate, holding a single lock
   * across both the read and the write.
   *
   * Already-trashed entries are excluded from matching by default.
   *
   * Returns the soft-deleted entries, or an empty array when nothing matched.
   */
  softDeleteWhere(
    predicateFn: GasSheetDbPredicate,
    options?: GasSheetDbFindOptions,
  ): GasSheetDbEntry[];

  /**
   * Restore a single soft-deleted entry.
   * Requires `_id`.
   */
  restore(entry: GasSheetDbEntry): GasSheetDbEntry;

  /**
   * Restore multiple soft-deleted entries.
   * Requires `_id` for each entry.
   */
  restoreMany(entries: GasSheetDbEntry[]): GasSheetDbEntry[];

  /**
   * Permanently delete a single entry.
   * Requires `_id`.
   */
  deleteOne(entry: GasSheetDbEntry): void;

  /**
   * Permanently delete multiple existing entries.
   * Requires `_id` for each entry.
   */
  deleteMany(entries: GasSheetDbEntry[]): void;

  /**
   * Permanently delete every entry matching a predicate, holding a single
   * lock across both the read and the write.
   *
   * Trashed entries are excluded from matching by default; pass
   * `{ onlyTrashed: true }` to purge the trash.
   *
   * Returns the entries as they were before deletion, or an empty array
   * when nothing matched.
   */
  deleteWhere(
    predicateFn: GasSheetDbPredicate,
    options?: GasSheetDbFindOptions,
  ): GasSheetDbEntry[];
}

/**
 * Entry point for database-like table access in sheets.
 */
declare class GasSheetDb {
  constructor(options?: GasSheetDbOptions);

  spreadsheet: GoogleAppsScript.Spreadsheet.Spreadsheet;
  rowNumbers: GasSheetDbRowsReference;
  lockScope: GasSheetDbLockScope;
  lockService: GasSheetDbLockService;
  logger: GasSheetDbLogger;

  /**
   * Create a table wrapper for a sheet. The sheet is created if it does not
   * exist. `lockScope` and `logger` fall back to this instance's values.
   */
  table(options?: GasSheetDbTableOptions): _GasSheetDbTable;
}
