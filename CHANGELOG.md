# Changelog

---

## [Unreleased]

### Added

- `logger` option on the `GasSheetDb` constructor, accepting any object exposing `info(msg, meta?)` and `warn(msg, meta?)`. Pass [`gas-logger`](https://github.com/yorsh-co/gas-logger) — or a `child()` of one — to persist `gas-sheetdb`'s schema and metadata events to a sheet, tag them with your own bindings, and filter them by level. Plain `console` satisfies the option too. Omit it and `gas-sheetdb` writes to the execution log, as before
- `logger` option on `GasSheetDb.table()`, overriding the parent instance's logger for that table alone. A table's own logger takes precedence even when the `GasSheetDb` instance was given none, so a single table can be routed to its own sheet or carry its own bindings without affecting the rest
- Both `logger` options are validated where they are passed, so an object missing `info` or `warn` — passing the `GasLogger` class rather than an instance, typically — fails when the instance or table is created instead of at the first log call, which may not come until a schema change or a decode failure, mid-operation and inside the write lock
- Log lines for the events that alter a sheet without being asked for directly: `Created table sheet` and `Added columns` (`info`), `Backfilled system metadata` with a count of rows repaired (`info`), and `Failed to decode JSON cells` (`warn`). Unrecoverable failures are still thrown rather than logged, so nothing is emitted at `error` and unexpected errors reach the caller's handler unchanged

### Changed

- Sheet creation now logs through the configured logger (`console` by default) instead of `Logger.log`, so the line reaches Cloud Logging alongside every other `gas-sheetdb` event
- A cell whose stored JSON no longer parses is now reported once per operation with a count of affected cells, instead of one `console.error` per cell. A single corrupt column previously emitted a log line per row — with a sheet-backed logger that would mean a spreadsheet write per row, inside the write lock

---

## [1.2.1] - 2026-08-06

### Fixed

- Only remove extra roles in `_GasSheetDbTable._removeExtraRows` if the table body contains data to avoid the 'it is not possible to delete all non-frozen rows' exception thrown by SpreadsheetApp

---

## [1.2.0] - 2026-08-03

### Added

- `updateWhere(predicateFn, update, options?)` and `updateOneWhere(predicateFn, update, options?)`, matching entries with a predicate and writing them back under a **single lock**, so read-modify-write operations (counters, balances, status transitions) can no longer interleave with another execution. `update` is either a patch object applied to every match, or a function returning a patch for a given entry; an updater that mutates its entry in place and returns nothing is honoured too. `_id` is always taken from the matched entry, so an updater cannot drop it or redirect the write to another row
- `softDeleteWhere(predicateFn, options?)` and `deleteWhere(predicateFn, options?)`, applying the same single-lock predicate matching to soft and permanent deletes. `deleteWhere` returns the entries as they were before deletion
- `lockScope` option (`'script' | 'document' | 'user'`) on the `GasSheetDb` constructor, overridable per table via `table({ lockScope })`
- `lockService` option on the `GasSheetDb` constructor, accepting any object exposing `withLock(scope, callback, options?)`. Pass [`gas-lock`](https://github.com/yorsh-co/gas-lock) to share one reentrancy-safe lock registry across every service in an execution

### Changed

- **BREAKING:** Write operations are now guarded by a **script**-scoped lock by default, where they previously used `LockService.getDocumentLock()` unconditionally. Callers depending on document-scoped semantics must now opt in explicitly with `lockScope: 'document'`.
- **BREAKING:** Rename `delete` method as `deleteOne` in `_GasSheetDbTable` to avoid conflict with JS `delete` operator.

### Fixed

- **Regression:** deleting every remaining entry no longer throws `Empty data`. `deleteMany` splices the table body down to zero rows, and `_setTableBodyData` rejected the empty result instead of clearing the orphaned rows off the sheet. This restores the behaviour originally shipped in [0.1.1](#011---2026-07-15), which the TypeScript migration reverted. Previously reachable only by passing every entry to `deleteMany`; `deleteWhere` makes it trivially reachable with a broad predicate
- Concurrent writes are now actually serialized in web app executions. `getDocumentLock()` returns `null` outside the context of a containing document — including a web app's `doGet`/`doPost` — so the previous hard-coded document lock provided no mutual exclusion there, regardless of the script being container-bound. Concurrent writers could interleave between `getMaxRows()` and `getLastRow()`, surfacing as `Those rows are out of bounds.` from `_ensureBlankRows`.
- Add missing `private` keyword before the `_GasSheetDbTable` methods `_insertSheet` and `_withLock`.

### Documentation

- Documented the predicate-based write methods and the single-lock guarantee they provide over a `find...()` followed by a separate write
- Documented the three lock scopes and their execution-context caveats
- Documented injecting `gas-lock` as a `lockService`, and the deadlock it prevents when a caller holds a lock across a `gas-sheetdb` write

---

## [1.1.1] - 2026-07-24

### Changed

- Add README.md to .prettierignore to avoid overwriting `delete(entry)` method formatting

---

## [1.1.0] - 2026-07-24

### Changed

- Migrated the library implementation from JavaScript to TypeScript with no functional changes
- Replaced JSDoc-based type definitions with native TypeScript declarations

### Build

- Added release tooling to publish a consumer-facing `dist` branch
- Added automatic generation of a minimal release `package.json` for published builds

---

## [1.0.1] - 2026-07-22

### Fixed

- Add missing sheet name index incrimination to `_insertSheet` in `_GasSheetDbTable` to fix a potential infinite loop.

---

## [1.0.0] - 2026-07-20

### Added

- Soft delete support via `softDelete`, `softDeleteMany`, `restore`, `restoreMany`, and `findTrashed`
- `find()` now supports `{ withTrashed: true }` and `{ onlyTrashed: true }` for querying soft-deleted entries
- Automatic `_isDeleted` metadata field, including backfilling existing rows that predate soft-delete support

### Changed

- **BREAKING:** `SheetDb` singleton replaced by an instantiable `GasSheetDb` class. Replace `SheetDb.table(name)` with `new GasSheetDb(options).table({ sheetName })`.
- Spreadsheet source is now resolved explicitly per instance via `spreadsheet`, `spreadsheetUrl`, `spreadsheetId`, or `useActiveSpreadsheet` — passing more than one throws instead of falling back to a silent priority order.
- Row configuration (`columnKeys`/`firstData`) is now set via constructor/table options instead of a global config file, with instance-level defaults overridable per table.
- Renamed internal "headers" terminology to "columnKeys" throughout (schema and table classes) for clarity.
- `find()` now excludes soft-deleted entries by default
- `insert`, `insertMany`, `update`, and `updateMany` now return the persisted entries after write operations, including generated metadata and stored values
- `delete` and `deleteMany` now explicitly represent permanent deletion, complementing the new soft-delete workflow

### Removed

- `sheetdb.config.js` and its module-level globals (`SHEETDB_USE_ACTIVE_SPREADSHEET`, `SHEETDB_SPREADSHEET_URL`, `SHEETDB_SHEET_NAMES`, `SHEETDB_ROW_NUMBERS`, `SHEETDB_ROW_INDEXES`).
- `sheetdb.service.js` (superseded by `gas-sheetdb.class.js`).

### Documentation

- Documented the soft-delete lifecycle, including querying, restoring, and permanently deleting entries
- Documented the new `find()` query options for including or filtering soft-deleted entries
- Clarified that write operations return the persisted entries

---

## [0.1.2] - 2026-07-16

### Fixed

- add missing return operator to `find()`

---

## [0.1.1] - 2026-07-15

### Fixed

- `_setTableBodyData` no longer throws when writing an empty data set,
  fixing a crash in `deleteMany` when it removes all remaining rows
  from a table
- `find()` now holds a single lock across metadata backfill and the
  subsequent read, closing a race window where concurrent writers
  could interleave between the two
- `_mapTableBodyRowIndexesById` now throws a descriptive error instead
  of silently mapping every row to an `undefined` key when the `_id`
  column doesn't exist yet
- Corrected the missing-`SHEETDB_SPREADSHEET_URL` error message to
  reference the correct file, `sheetdb.config.js`

### Documentation

- Documented that `insert`/`insertMany`/`update`/`updateMany` mutate
  passed-in entry objects in place (system fields written directly
  onto the caller's objects), in both JSDoc and the README

---

## [0.1.0] - 2026-06-16

Initial public release of `gas-sheetdb`.

### Added

- `SheetDb` library for spreadsheet-backed table storage
- `_SheetDbTable` table wrapper for working with sheet rows as objects
- `find`, `findWhere`, and `findOneWhere` query methods
- `insert`, `insertMany`, `update`, `updateMany`, `delete` and `deleteMany` persistence methods
- Automatic sheet creation for missing tables
- Automatic column creation for new object properties
- Automatic `_id` UUID generation
- Automatic `_createdAt` and `_updatedAt` metadata fields
- Automatic handling of manually inserted entries
- Object and array encoding/decoding for sheet storage
- Spreadsheet schema synchronization via `_SheetDbTableSchema`
- Spreadsheet access configuration for:
  - bound spreadsheet mode
  - standalone spreadsheet mode
- Spreadsheet caching and lazy initialization
- Apps Script `LockService` support for safer concurrent writes
- Git subtree-oriented project structure
- `clasp`-friendly local development workflow support

### Architecture

- Designed around spreadsheet-backed table storage instead of generic spreadsheet helpers
- Uses Google Sheets as lightweight database-style persistence
- Separates persisted metadata from runtime-only metadata
- Uses codec-based value serialization for complex value storage
- Uses schema synchronization to keep sheet headers aligned with object properties
- Uses document locks during inserts and updates to reduce concurrent write conflicts
