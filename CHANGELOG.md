# Changelog

---

## [2.0.0] - 2026-10-27

### Added

- `findWhere()` and `findOneWhere()` take the same `GasSheetDbFindOptions` as `find()` and the `*Where` writers (`withTrashed`, `onlyTrashed`); trashed entries stay excluded by default
- `lockTimeoutMs` option on `GasSheetDb` and `table()` (default `30000`; a table overrides the instance), passed to every `lockService.withLock` call. It must be a positive, finite number
- Typed data errors: a missing `_id` throws a 400 `VALIDATION_ERROR`, an unknown `_id` a 404 `NOT_FOUND` and a duplicate `_id` on insert a 409 `CONFLICT`. When [`gas-error`](https://github.com/yorsh-co/gas-error) is loaded they are its `GasValidationError`, `GasNotFoundError` and `GasConflictError`, so `GasError.handle` reports them correctly; otherwise they are a `GasSheetDbError` with the same `statusCode`, `code`, `details` and `isOperational` fields. Both are still `instanceof Error`. Configuration errors remain plain `Error`s
- Vitest test suite with in-memory Apps Script service mocks that model Sheets' value coercion, grid bounds and frozen-row rules

### Changed

- **BREAKING:** `sheetName` is required. `table()` throws a configuration `Error` when it is missing or not a non-empty string, and unnamed tables (`gsd-table-N`) are gone. `GasSheetDbTableOptions.sheetName` and the table's `sheetName` are now `string`
- **BREAKING:** `new GasSheetDb({ useActiveSpreadsheet: false })` with no other spreadsheet source throws instead of silently using the active spreadsheet
- **BREAKING:** strings are stored as text, exactly as given, so number-, boolean- and date-looking strings (`'0012'`, `'TRUE'`, `'2024-01-02'`, `'1-2'`) now read back as strings rather than as the number, boolean or `Date` Sheets parsed them into. Consumers that relied on that coercion must pass real numbers, booleans and `Date`s, which are still stored natively
- **BREAKING:** `find()`, `findWhere()`, `findOneWhere()` and `findTrashed()` no longer take the lock when the sheet needs no repair, and no longer delete trailing blank rows. A read takes the lock only to add missing system columns or backfill rows, then repairs and reads again under it. Code that relied on a read serializing with writers, or on a read trimming the sheet, must be updated. Writes no longer trim trailing blank rows either: rows are only added by inserts and removed by deletes
- **BREAKING:** when no `lockService` is injected, the `GasLock` global is used if it is loaded, otherwise the minimal `LockService` fallback as before. It is looked up at each acquisition, so file load order does not matter. `lockService` on the instance and on tables is never `undefined`: it is the default service, which resolves `GasLock` itself. Nothing needs to be configured to use [`gas-lock`](https://github.com/yorsh-co/gas-lock)
- **BREAKING:** every error message now carries the `[GasSheetDb]` prefix (`Empty data`, `Cannot map rows by "_id"`, `Entry not found`, `Cannot update entry without "_id"`), and deleting an entry without an `_id` now says `Cannot delete` rather than `Cannot update`. Code that matches on message text must be updated; match on `code` instead
- **BREAKING:** `rowNumbers` is validated when a `GasSheetDb` or a table is created, after the table's override is merged onto the instance's: `columnKeys` and `firstData` must be positive integers and `firstData` must be greater than `columnKeys`. Values that were silently accepted, such as `0`, a fraction or `firstData <= columnKeys`, now throw. Inserts append at `max(lastRow + 1, firstData)`, so a spacer row between the header and `firstData` is never written into
- **BREAKING:** `insertMany([])` returns `[]` instead of throwing `Empty data`; `updateMany`, `softDeleteMany` and `restoreMany` return `[]` and `deleteMany` returns nothing, none of them taking the lock. Code that relied on the throw must check the batch itself
- **BREAKING:** `insert` and `insertMany` throw a 409 `CONFLICT` error, `Duplicate "_id": <id>`, when an entry's `_id` is already stored or appears twice in the batch, before anything is written or any entry is changed. Inserting an entry obtained from `find()` now throws: strip `_id` and the other system fields to insert a copy, or use `update` to change the stored entry
- A table sheet that does not exist is now created under the lock, so `table()` can block, or throw when the lock cannot be acquired
- Lint, TypeScript and Prettier configuration and the test loader come from [`gas-dev-config`](https://github.com/yorsh-co/gas-dev-config)
- Dependabot groups development dependencies into one weekly pull request
- `npm run typecheck` also covers `test/` through `tsconfig.test.json`, and type-checks `peer/gas-sheetdb.peer.types.d.ts` standalone; `npm run check` runs the test suite and `gas-dev-config check`

### Fixed

- The default logger no longer prints a literal `undefined` when a log line has no meta
- Strings written to a sheet were parsed by Sheets as if typed: `=IMPORTXML(...)` or `=HYPERLINK(...)` in user input became a live formula (formula injection), `0012` lost its leading zeros and `TRUE` or `1-2` turned into a boolean or a date. Every string, including JSON-encoded cells, backfilled ids and header keys, is now written as text. Updates still compare logical values, so unchanged string cells are not rewritten
- A string that starts with `__JSON__` was read back as a parsed object. It is now encoded as JSON and decodes to the identical string; cells already stored read as before
- A cell longer than 50,000 characters, a `bigint`, a `symbol`, a `function` and a value that does not serialize to JSON throw a 400 `VALIDATION_ERROR` naming the column, with `details` `{ column, reason }`, before anything is written or any passed object is changed. They used to fail inside Sheets, or write a truncated or corrupt cell
- Fully blank rows inside the data range, such as a row whose contents a person cleared, are no longer entries: reads skip them and the metadata backfill leaves them empty, where they were returned as empty entries and given an `_id`. A row is blank when every cell reads `''`, so a row whose only content is a formula that returns a value is still an entry
- `table()` and the constructor no longer create a sheet when `rowNumbers.columnKeys` is greater than `1`: row banding ran past the end of the grid and threw `Those rows are out of bounds.`
- `updateWhere`, `updateOneWhere`, `deleteWhere` and the other writers that repair the sheet before reading no longer append a second set of system columns when the header was edited after the table was created: the cached header is reloaded first
- Header cells are read as strings, so a numeric header such as `2024` matches the entry key `'2024'`; every write used to append another `2024` column. Blank header cells are ignored: they get no key on entries and their cells are never written. When a header name repeats, the first column wins, later ones are ignored like blanks, and `Duplicate column keys` is logged as a warning once per operation
- `updateMany`, `softDeleteMany` and `restoreMany` validate every entry (`_id` present and found) before changing any passed object, adding columns or writing. A failure on a later entry used to leave the earlier entries mutated, and `softDelete` and `restore` set `_isDeleted` on the passed objects before the lock was taken. Successful writes still update the passed objects' `_updatedAt` and `_isDeleted`
- The `_isDeleted` backfill keeps a non-boolean value's meaning instead of overwriting it with `false`: a blank becomes `false`, booleans are kept and anything else is converted with `Boolean(value)`, so `1` stays deleted and `0` is not. Deleted rows no longer reappear after a hand edit. The rows are counted in `Backfilled system metadata`. Not marked breaking: booleans and blanks behave as before, and the converted cells were previously rewritten to `false`, silently restoring entries
- Duplicate `_id` values are repaired. When several rows share an `_id` the topmost keeps it and the others get a new one, logged once as `Reassigned duplicate ids`. A read that finds duplicates takes the lock to repair them, like any other backfill
- `deleteMany`, `deleteOne` and the other direct deletes resolve the `_id` column from a header read taken under the lock instead of the cached one. After a person inserted or moved columns, the stale index could miss or match the wrong rows. `update`, `insert` and the predicate writers already reloaded the header under the lock
- Formulas, including array-formula spills, in table rows are preserved. `update`, `softDelete`, `restore` and the system metadata backfill rewrote the whole table body from its computed values, turning every formula into a static value
- Writes touch only the cells whose value changes, grouped into one `setValues` call per run of adjacent cells and rows, instead of rewriting the whole body. A cell an entry sets to the value it already has is not written. A manual edit made to another cell while an operation runs is no longer overwritten, and the cost no longer grows with the table size
- Deletes remove the rows in place, bottom-up and one `deleteRows` call per run of adjacent rows, instead of rewriting the body. When every non-frozen row would be deleted, one row is cleared instead, as Sheets refuses to delete them all
- Log lines are no longer emitted while the lock is held. `Created table sheet`, `Added columns`, `Backfilled system metadata` and `Failed to decode JSON cells` are queued and emitted in order once the lock is released, also when the operation throws. A sheet-backed logger such as `gas-logger` takes the script lock itself, and its release ended the `gas-sheetdb` critical section early. A logger that throws no longer fails or masks the operation; the failure is reported through `console`
- Creating a table sheet no longer throws `A sheet with the name ... already exists` when another execution creates it between the lookup and the insert: the lookup is repeated under the lock
- An invalid `lockService`, such as Apps Script's `LockService` global passed instead of `GasLock`, is rejected when the instance is created instead of failing on the first write

### Removed

- The unused `_runtime` non-persisted field and its internal constant; no code ever set it

### Docs

- Corrected the description of `LockService` semantics. Locks are owned per execution: a nested acquire succeeds immediately and does not deadlock, but releasing any handle ends the outer critical section early. The README's deadlock example, and the note that the document-lock default was dropped in `2.0.0` (it shipped in `1.2.0`), are fixed
- The Logging section no longer advises buffering to keep log writes outside the hot path, as logs are emitted after the lock is released
- Added a "Migrating to 2.0.0" section, the `filePushOrder` example now matches the shipped files and the subtree path, and the Logging section recommends `gas-logger` `1.2.0` or later. Fixed broken examples, the return values of the write operations, and the feature and planned-feature lists. Documented the canonical peer type declarations in `peer/gas-sheetdb.peer.types.d.ts`

---

## [1.3.2] - 2026-10-27

### Fixed

- Deleting every remaining entry now empties the sheet. `deleteWhere()`, `deleteMany()` and `deleteOne()` reported the entries deleted but left their rows on the sheet whenever the table body would become empty, so they reappeared on the next read. The previously occupied rows are now cleared and the first body row is kept as an empty row, since Sheets refuses to delete all non-frozen rows. This was a regression of the fix shipped in 1.2.0
- `deleteMany()` no longer removes an unrelated entry when the same entry appears more than once in the batch, or when several entries resolve to the same row. `deleteMany([a, a])` on `[a, b, c]` previously left `[c]`; it now leaves `[b, c]`
- Calling `table()` without `sheetName` while the `gsd-table-0` sheet already exists no longer hangs. The default-name counter was never incremented, so the lookup looped until the execution timed out; the next free `gsd-table-N` is now used. This was a regression of the fix shipped in 1.0.1

### Changed

- Releases are published with [`gas-dev-release`](https://github.com/yorsh-co/gas-dev-release) instead of `scripts/release.sh`, which adds preflight checks and pushes the branch and tags atomically. The empty `gas-sheetdb.types.js` is no longer shipped
- Tags use a `v` prefix, now created by the release tooling: versions are tagged `vX.Y.Z` and releases `dist/vX.Y.Z`
- Bump the transitive dev dependency `brace-expansion` to resolve the high-severity advisory reported by `npm audit`

---

## [1.3.1] - 2026-08-07

### Fixed

- `findWhere()` is no longer declared as returning `GasSheetDbEntry[] | null`. It is implemented with `Array#filter`, which always returns an array — an empty one is still truthy — so the `null` was unreachable and every TypeScript consumer was forced through a null check that could not fire. It now returns `GasSheetDbEntry[]`, matching `findTrashed()`, `updateWhere()`, `softDeleteWhere()` and `deleteWhere()`, which all return an empty array when nothing matches. No runtime change: the `null` branch was never taken, so existing null checks remain valid and simply become redundant

---

## [1.3.0] - 2026-08-07

### Added

- `logger` option on the `GasSheetDb` constructor, accepting any object exposing `info(msg, meta?)` and `warn(msg, meta?)`. Pass [`gas-logger`](https://github.com/yorsh-co/gas-logger) — or a `child()` of one — to persist `gas-sheetdb`'s schema and metadata events to a sheet, tag them with your own bindings, and filter them by level. Plain `console` satisfies the option too. Omit it and `gas-sheetdb` writes to the execution log, as before
- `logger` option on `GasSheetDb.table()`, overriding the parent instance's logger for that table alone. A table's own logger takes precedence even when the `GasSheetDb` instance was given none, so a single table can be routed to its own sheet or carry its own bindings without affecting the rest
- Both `logger` options are validated where they are passed, so an object missing `info` or `warn` — passing the `GasLogger` class rather than an instance, typically — fails when the instance or table is created instead of at the first log call, which may not come until a schema change or a decode failure, mid-operation and inside the write lock
- Log lines for the events that alter a sheet without being asked for directly: `Created table sheet` and `Added columns` (`info`), `Backfilled system metadata` with a count of rows repaired (`info`), and `Failed to decode JSON cells` (`warn`). Unrecoverable failures are still thrown rather than logged, so nothing is emitted at `error` and unexpected errors reach the caller's handler unchanged

### Fixed

- `_ensureRequiredMetadata` no longer rewrites the entire table body on every read. It wrote the rows back unconditionally, so each `find...()` call spent a full-width `setValues` re-persisting values it had just read and found correct. The write now runs only when a row was actually repaired. The orphaned-row trim it also performed is preserved, so the shape of the sheet is unchanged

### Changed

- Sheet creation now logs through the configured logger (`console` by default) instead of `Logger.log`, so the line reaches Cloud Logging alongside every other `gas-sheetdb` event
- A cell whose stored JSON no longer parses is now reported once per operation with a count of affected cells, instead of one `console.error` per cell. A single corrupt column previously emitted a log line per row — with a sheet-backed logger that would mean a spreadsheet write per row, inside the write lock

---

## [1.2.1] - 2026-08-06

### Fixed

- Only remove extra rows in `_GasSheetDbTable._removeExtraRows` if the table body contains data to avoid the 'it is not possible to delete all non-frozen rows' exception thrown by SpreadsheetApp

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

- **Regression:** deleting every remaining entry no longer throws `Empty data`. `deleteMany` splices the table body down to zero rows, and `_setTableBodyData` rejected the empty result instead of clearing the orphaned rows off the sheet. This restores the behaviour originally shipped in [0.1.1](#011---2026-07-15), which the TypeScript migration reverted. (Regressed again in 1.2.1 and restored in 1.3.2) Previously reachable only by passing every entry to `deleteMany`; `deleteWhere` makes it trivially reachable with a broad predicate
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

- Add missing sheet name index incrimination to `_insertSheet` in `_GasSheetDbTable` to fix a potential infinite loop. (Regressed in 1.1.0 by the TypeScript migration and restored in 1.3.2)

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
