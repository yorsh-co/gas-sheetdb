# Changelog

All notable changes to this project will be documented in this file.

The format loosely follows [Keep a Changelog](https://keepachangelog.com/en/1.1.0/).

---

## [0.1.0] - 2026-05-27

Initial public release of `gas-sheetdb`.

### Added

- `SheetDb` library for spreadsheet-backed table storage
- `_SheetDbTable` table wrapper for working with sheet rows as objects
- `find`, `findWhere`, and `findOneWhere` query methods
- `insert`, `insertMany`, `update`, and `updateMany` persistence methods
- Automatic sheet creation for missing tables
- Automatic column creation for new object properties
- Automatic `_id` UUID generation
- Automatic `_createdAt` and `_updatedAt` metadata fields
- Runtime-only `_runtime.rowNumber` metadata support
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
