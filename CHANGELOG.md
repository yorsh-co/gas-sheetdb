# Changelog

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
