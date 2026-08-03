# gas-sheetdb

[![Built with Google Apps Script](https://img.shields.io/badge/Built%20with-Google%20Apps%20Script-4285F4?logo=google&logoColor=white)](https://developers.google.com/apps-script)

## Database-like storage layer for Google Apps Script projects.

> The goal of this project is to use Google Sheets as database storage in Apps Script projects without having to manage sheets manually, store layouts, hard-code ranges, or handle concurrent write logic.

`gas-sheetdb` uses Apps Script's `SpreadsheetApp` to turn sheets in a Google Spreadsheet into tables that store object entries as rows.

Entries can be queried and updated through a `GasSheetDb` instance using ORM-like methods such as `find`, `insert`, `update`, `softDelete`, `restore` and `delete`.

> **Disclaimer:**
> This project and [Yorsh](https://github.com/yorsh-co) are independent and are not affiliated with, endorsed by, or associated with Google LLC.

### Features

- Store object entries as rows in sheets
- Query entries using methods like find, findWhere, and findOneWhere
- Insert, update, soft delete, restore and permanently delete entries using plain JavaScript objects
- Update and delete entries matching a predicate, with the read and the write held under a single lock
- Write operations return the persisted entries, including the generated metadata
- Automatically creates missing columns when new properties appear during inserts or updates
- Automatically adds `_id`, `_createdAt`, `_updatedAt`, and `_isDeleted` metadata fields to new entries
- Automatically mutates the metadata properties of entries passed to write operations
- Returns the full entry object after partial updates
- Instance-based API — configure spreadsheet source and row numbers per instance, with per-table overrides
- Objects and arrays are JSON serialized
- Accepts new entries inserted manually to the spreadsheet
- Supports both bound and standalone spreadsheets
- Uses Apps Script `LockService` for safer concurrent writes, with a configurable lock scope per instance or per table (full transactions are not yet supported)
- Optionally accepts an injected lock service ([`gas-lock`](https://github.com/yorsh-co/gas-lock)) so nested locks across multiple services in one execution reuse a single lock instead of deadlocking- Written in TypeScript with generated JavaScript distribution
- No external dependencies beyond built-in Apps Script services

### Example Usage

```js
// Create a new user
const sheetDb = new GasSheetDb();
const usersTable = sheetDb.table({ sheetName: '👤 Users' });

usersTable.insert({
  name: 'John',
  role: 'admin',
  permissions: ['users:read', 'users:write'],
});

// Filter users
const admins = usersTable.findWhere((user) => user.role === 'admin');

// Update users
admins.forEach((admin) => {
  admin.lastSeenAt = new Date();
});

usersTable.updateMany(admins);

// Soft delete
const revoked = usersTable.findWhere((user) => user.access === 'revoked');

usersTable.softDeleteMany(revoked);

// Restore
const usersToRestore = usersTable.findTrashed((user) => user.project ===='projectA1');

usersTable.restoreMany(usersToRestore);

// Permanently delete
const usersToDelete =usersTable.findTrashed((user) => user.project ===='projectB1');

usersTable.deleteMany(usersToDelete);
```

## Requirements

### Scopes

`gas-sheetdb` requires one spreadsheet access scope to be added to the parent project's `appsscript.json`.

Use the scope that matches the [spreadsheet access mode](#spreadsheet-access-modes) being used:

| Mode                      | Scope                                                        |
| ------------------------- | ------------------------------------------------------------ |
| Active spreadsheet mode   | `"https://www.googleapis.com/auth/spreadsheets.currentonly"` |
| Explicit spreadsheet mode | `"https://www.googleapis.com/auth/spreadsheets"`             |

### Example `appsscript.json`

```js
{
  "timeZone": "America/Sao_Paulo",
  "dependencies": {},
  "exceptionLogging": "STACKDRIVER",
  "runtimeVersion": "V8",

  "oauthScopes": [
    "https://www.googleapis.com/auth/spreadsheets.currentonly"
  ]
}
```

## Spreadsheet Access Modes

`gas-sheetdb` resolves its spreadsheet from the option passed to the `GasSheetDb` constructor.

### Active Spreadsheet Mode

Uses the active spreadsheet attached to the Apps Script project.

```js
const sheetDb = new GasSheetDb({ useActiveSpreadsheet: true });
```

Use this mode for spreadsheet [container-bound scripts](https://developers.google.com/apps-script/guides/bound).

> **Note:** `GasSheetDb` defaults to the active spreadsheet mode if no argument is passed to the class constructor.

### Explicit Spreadsheet Mode

Uses the spreadsheet indicated by the `spreadsheet`, `spreadsheetUrl` or `spreadsheetId` options. Examples below.

Use this mode for:

- standalone Apps Script projects requiring persistent storage
- storing data used by multiple projects
- storing data for a project bound to a different spreadsheet, form or document

#### Using the `spreadsheet` option

```js
const mySpreadsheet = SpreadsheetApp.create('My Spreadsheet');

const sheetDb = new GasSheetDb({
  spreadsheet: mySpreadsheet,
});
```

#### Using the `spreadsheetUrl` option

```js
const mySpreadsheetUrl = 'https://docs.google.com/spreadsheets/d/...';

const sheetDb = new GasSheetDb({
  spreadsheetUrl: mySpreadsheetUrl,
});
```

#### Using the `spreadsheetId` option

```js
const mySpreadsheetId = '1a2b3c...';

const sheetDb = new GasSheetDb({
  spreadsheetId: mySpreadsheetId,
});
```

## Quick Start

It is recommended to use `gas-sheetdb` together with [Google's `clasp` CLI](https://github.com/google/clasp) for local Apps Script development and git-based workflows. See [Setup instructions with `clasp`](#setup-instructions-with-clasp) for more information.

#### 1. Add the library to your Apps Script project

This repository is intended to be added directly into Apps Script projects using git subtree.

```bash
git subtree add \
  --prefix=src/lib/gas-sheetdb \
  https://github.com/yorsh-co/gas-sheetdb.git \
  dist \
  --squash
```

This creates:

```txt
src/lib/gas-sheetdb/
```

#### 2. Configure Apps Script scopes

Add the required spreadsheet scope to the parent project's `appsscript.json`.

See the [Scopes](#scopes) section above.

#### 3. If needed, move `gas-sheetdb` files to the start of the execution order.

This is required for declaring a `GasSheetDb` instance at runtime, as a global variable or inside an IIFE.

See the [Configure the file push order](#6-configure-the-file-push-order) section for details.

#### 4. Declare a GasSheetDb instance with your configuration

```js
const sheetDb = new GasSheetDb({
  useActiveSpreadsheet: true, // or spreadsheet / spreadsheetUrl / spreadsheetId
  rowNumbers: { columnKeys: 2, firstData: 3 }, // optionally, set your own row configuration. Defaults to `{ columnKeys: 1, firstData: 2 }`
  lockScope: 'script', // optionally, set the lock scope for writes. Defaults to `'script'`
  lockService: GasLock, // optionally, configure the instance to use `gas-lock` in the place of the default LockService
});
```

#### 5. Create or link a table

```js
const myTable = sheetDb.table({
  sheetName: 'My Table',
  rowNumbers: { columnKeys: 3, firstData: 4 }, // optionally, set a table-specific row configuration that takes precedence over the `sheetDb` row configuration
  lockScope: 'user', // optionally, override the instance lock scope for this table
});
```

## Setup instructions with `clasp`

`gas-sheetdb` works best with [Google's `clasp` CLI](https://github.com/google/clasp) for local Apps Script development and git-based workflows.

#### 1. Install clasp

```bash
npm install -g @google/clasp
```

#### 2. Enable the [Apps Script API](https://script.google.com/home/usersettings)

#### 3. Login to Google Apps Script

```bash
clasp login
```

#### 4. Clone or create your Apps Script project

Clone an existing project:

```bash
clasp clone <script-id>
```

or create a new project:

```bash
clasp create --type sheets
```

#### 5. Import `gas-sheetdb`

```bash
git subtree add \
  --prefix=src/lib/gas-sheetdb \
  https://github.com/yorsh-co/gas-sheetdb.git \
  dist \
  --squash
```

This creates:

```txt
src/lib/gas-sheetdb/
```

#### 6. Configure the file push order

Apps Script executes files by the order in the Apps Script editor, from top to bottom. By default, `clasp push` orders the files alphabetically, by file name. If a `GasSheetDb` instance is declared at runtime (as a global variable or in an IIFE) in file referencing `GasSheetDb` that is ordered before `gas-sheetdb`'s own files, `clasp push` will succeed but running the project will throw:

```txt
ReferenceError: GasSheetDb is not defined
```

To avoid this, add a [`filePushOrder`](https://github.com/google/clasp#filepushorder-optional) entry to your project's `.clasp.json` that pushes `gas-sheetdb`'s module files ahead of any file that references them:

```json
{
  "filePushOrder": [
    "dist/lib/gas-sheetdb/gas-sheetdb.constants.js",
    "dist/lib/gas-sheetdb/gas-sheetdb.codec.js",
    "dist/lib/gas-sheetdb/gas-sheetdb.schema.js",
    "dist/lib/gas-sheetdb/gas-sheetdb.table.js",
    "dist/lib/gas-sheetdb/gas-sheetdb.class.js",
    "dist/lib/gas-sheetdb/gas-sheetdb.types.js"
  ]
}
```

Alternatively, you can manually move these files to the top of the file list in the Apps Script editor.

> **Note:**
> Any file in your own project that constructs a `GasSheetDb` instance (e.g. `config = new GasSheetDb({ ... })`) must be pushed _after_ the entries above.

#### 7. Push local files to Apps Script

```bash
clasp push
```

#### 8. Configure Apps Script scopes

Add the required spreadsheet scope to the parent project's `appsscript.json`.

See the [Scopes](#scopes) section above.

#### 9. Declare a GasSheetDb instance with your configuration

```js
const sheetDb = new GasSheetDb({
  useActiveSpreadsheet: true, // or spreadsheet / spreadsheetUrl / spreadsheetId
  rowNumbers: { columnKeys: 2, firstData: 3 }, // optionally, set your own row configuration. Defaults to `{ columnKeys: 1, firstData: 2 }`
  lockScope: 'script', // optionally, set the lock scope for writes. Defaults to `'script'`
});
```

#### 10. Create or link a table

```js
const myTable = sheetDb.table({
  sheetName: 'My Table',
  rowNumbers: { columnKeys: 3, firstData: 4 }, // optionally, set a table-specific row configuration that takes precedence over the `sheetDb` row configuration
  lockScope: 'user', // optionally, override the instance lock scope for this table
});
```

## Basic Usage

### Create a `GasSheetDb` instance

```js
const sheetDb = new GasSheetDb({
  spreadsheetUrl: 'https://docs.google.com/spreadsheets/d/...', // or spreadsheet / spreadsheetId / useActiveSpreadsheet
  rowNumbers: { columnKeys: 2, firstData: 3 }, // optionally, set your own row configuration. Defaults to `{ columnKeys: 1, firstData: 2 }`
});
```

### Create a Table

```js
const usersTable = sheetDb.table({
  sheetName: '👤 Users',
  rowNumbers: { columnKeys: 3, firstData: 4 }, // for example, accommodate for a custom header in rows 1 and 2
});
```

### Insert an Entry

```js
usersTable.insert({
  name: 'John',
  email: 'john@email.com',
  role: 'editor',
});
```

### Insert Multiple Entries

```js
usersTable.insertMany([
  {
    name: 'John',
    email: 'john@email.com',
    role: 'editor',
  },
  {
    name: 'Jane',
    email: 'jane@email.com',
    role: 'admin',
  },
]);
```

> **Note:**
> `insert` and `insertMany` mutate the entry objects passed in, adding `_id`, `_createdAt`, `_updatedAt` and `_isDeleted` in place.

### Read All Entries

```js
const users = usersTable.find();
```

### Filter Entries

```js
const admins = usersTable.findWhere((user) => user.role === 'admin');
```

### Find a Single Entry

```js
const user = usersTable.findOneWhere((user) => user._id === 'abc123');
```

> **Note:**
> The `find`, `findWhere` and `findOneWhere` methods exclude soft-deleted entries by default.

### Find Deleted Entries

```js
// return all soft-deleted entries
const deletedUsers = usersTable.findTrashed();

// optionally filter soft-deleted entries
const deletedUsers = usersTable.findTrashed((user) => user._id === 'abc123');

// also supported
const deletedUsers = usersTable.find({ onlyTrashed: true });
```

### Find Active and Deleted Entries

```js
const allUsers = usersTable.find({ withTrashed: true });
```

### Update an Entry

> **Note:**
> The `update` and `updateMany` methods only accept objects containing the original `_id` property returned by one of the `find...()` methods.
>
> Both methods mutate the metadata of the entry objects passed in the argument and also return the full entry object(s) after mutation.

```js
const user = usersTable.findOneWhere(
  (entry) => entry.email === 'john@email.com',
);

user.active = false;

usersTable.update(user);
```

### Update Multiple Entries

> **Note:**
> The `update` and `updateMany` methods only accept objects containing the original `_id` property returned by one of the `find...()` methods.
>
> Both methods mutate the metadata of the entry objects passed in the argument and also return the full entry object(s) after mutation.

```js
const users = usersTable.findWhere((entry) => entry.role === 'editor');

users.forEach((user) => {
  user.active = true;
});

usersTable.updateMany(users);
```

### Partial Update

> **Note:**
> As long as the object contains original `_id` property returned by one of the `find...()` methods, `update` and `updateMany` can accept partial entries.
>
> Both methods return the full entry object(s) after mutation.

```js
const oldUser = usersTable.findOneWhere(
  (entry) => entry.email === 'john@email.com',
);

const userRoleConfig = {
  _id: user._id,
  role: 'editor',
};

const updatedUser = usersTable.update(userRoleConfig); // returns the full entry
```

### Soft Delete an Entry

> **Note:**
> The `softDelete` and `softDeleteMany` methods only accept objects containing the original `_id` property returned by one of the `find...()` methods.
>
> Both methods mutate the metadata of the entry objects passed in the argument and also return the full entry object(s) after mutation.

```js
const user = usersTable.findOneWhere(
  (entry) => entry.email === 'john@email.com',
);

usersTable.softDelete(user);
```

### Soft Delete Multiple Entries

> **Note:**
> The `softDelete` and `softDeleteMany` methods only accept objects containing the original `_id` property returned by one of the `find...()` methods.
>
> Both methods mutate the metadata of the entry objects passed in the argument and also return the full entry object(s) after mutation.

```js
const users = usersTable.findWhere((entry) => entry.access === 'revoked');

usersTable.softDeleteMany(users);
```

### Restore an Entry

> **Note:**
> The `restore` and `restoreMany` methods only accept objects containing the original `_id` property returned by one of the `find...()` methods.
>
> Both methods mutate the metadata of the entry objects passed in the argument and also return the full entry object(s) after mutation.

```js
const user = usersTable.findTrashed(
  (entry) => entry.email === 'john@email.com',
)[0]; // `findTrashed` returns an array of entries

usersTable.restore(user);
```

### Restore Multiple Entries

> **Note:**
> The `restore` and `restoreMany` methods only accept objects containing the original `_id` property returned by one of the `find...()` methods.
>
> Both methods mutate the metadata of the entry objects passed in the argument and also return the full entry object(s) after mutation.

```js
const users = usersTable.findTrashed((entry) => entry.access === 'renewed');

usersTable.restoreMany(users);
```

### Permanently Delete an Entry

> **Note:**
> The `delete` and `deleteMany` methods only accept objects containing the original `_id` property returned by one of the `find...()` methods.

```js
const user = usersTable.findOneWhere(
  (entry) => entry.email === 'john@email.com',
);

usersTable.delete(user);
```

### Permanently Delete Multiple Entries

> **Note:**
> The `delete` and `deleteMany` methods only accept objects containing the original `_id` property returned by one of the `find...()` methods.

```js
const users = usersTable.findWhere((entry) => entry.access === 'revoked');

usersTable.deleteMany(users);
```

## Predicate-Based Writes

Reading with one method and writing with another takes two separate locks, and another execution can write to the table in between. `updateWhere`, `updateOneWhere`, `softDeleteWhere` and `deleteWhere` match entries with a predicate and write them back while **holding a single lock across both the read and the write**.

Use them whenever the new value depends on the value already stored — counters, balances, status transitions, claiming a queued row.

```js
// this is NOT safe: another execution can write between the two calls
const user = usersTable.findOneWhere((entry) => entry.email === email);
usersTable.update({ _id: user._id, visits: user.visits + 1 });

// this is safe: one lock covers the read and the write
usersTable.updateOneWhere(
  (entry) => entry.email === email,
  (entry) => ({ visits: entry.visits + 1 }),
);
```

### Update Entries Matching a Predicate

`updateWhere` updates every match and returns the updated entries, or an empty array when nothing matched. The second argument is either a function returning a patch for a given entry, or a patch object applied to every match.

```js
// computed per entry
const promoted = usersTable.updateWhere(
  (entry) => entry.role === 'editor',
  (entry) => ({ role: 'admin', promotedFrom: entry.role }),
);

// the same patch for every match
usersTable.updateWhere((entry) => entry.access === 'revoked', {
  active: false,
});
```

An updater that mutates its entry in place and returns nothing works as well:

```js
usersTable.updateWhere(
  (entry) => entry.role === 'editor',
  (entry) => {
    entry.active = true;
  },
);
```

> **Note:**
> `_id` is always taken from the matched entry. An updater cannot drop it or redirect the write to a different row.

### Update a Single Entry Matching a Predicate

`updateOneWhere` updates the **first** match only and returns the updated entry, or `null` when nothing matched.

```js
const user = usersTable.updateOneWhere(
  (entry) => entry.email === 'john@email.com',
  { active: false },
);
```

### Soft Delete Entries Matching a Predicate

`softDeleteWhere` returns the soft-deleted entries, or an empty array when nothing matched.

```js
const trashed = usersTable.softDeleteWhere(
  (entry) => entry.access === 'revoked',
);
```

### Permanently Delete Entries Matching a Predicate

`deleteWhere` returns the entries as they were before deletion, or an empty array when nothing matched.

```js
const removed = usersTable.deleteWhere((entry) => entry.access === 'revoked');
```

All four methods exclude soft-deleted entries from matching by default, and accept the same options as `find`. Pass `onlyTrashed` to purge the trash:

```js
usersTable.deleteWhere((entry) => entry._updatedAt < cutoff, {
  onlyTrashed: true,
});
```

## Project Details

### Automatic Column Management

New columns are inserted automatically whenever an entry contains new properties.

Example:

```js
usersTable.insert({
  name: 'James',
  email: 'james@email.com',
  avatarUrl: 'https://example.com/avatar.jpg',
});
```

If `avatarUrl` does not already exist in the sheet headers, the column is appended automatically.

### Metadata

New entries automatically receive:

```js
{
  _id: 'uuid',
  _createdAt: Date,
  _updatedAt: Date,
  _isDeleted: boolean
}
```

> **Note:** All metadata properties are managed automatically and should not normally be modified directly.

### Entry Lifecycle

Entries normally exist in one of two states:

- active
- soft deleted

Soft-deleted entries remain stored in the sheet and can be restored later.
They are excluded from `find()` by default.

Permanent deletion removes rows from the spreadsheet entirely.

### Manual inserts

Entries can be added manually to the sheet table. Metadata will be applied automatically to manually-added entries when they are first read by `table.find()`.

This is useful for importing data directly to the sheet from another source such as a file (.csv, .xlsx etc.) or via copy and paste.

### Encoded Values

`gas-sheetdb` automatically encodes and decodes objects and arrays when storing and reading entries.

Example:

```js
{
  tags: ['crm', 'sales'],
  settings: {
    notifications: true,
  },
}
```

Stored in the sheet as:

| tags                      | settings                         |
| ------------------------- | -------------------------------- |
| `__JSON__["crm","sales"]` | `__JSON__{"notifications":true}` |

Values are automatically decoded when reading rows.

### Locking

Every write operation runs inside an Apps Script `LockService` lock. The scope is configurable per instance, and overridable per table:

```js
const sheetDb = new GasSheetDb({
  useActiveSpreadsheet: true,
  lockScope: 'script', // default
});

const auditTable = sheetDb.table({
  sheetName: 'Audit',
  lockScope: 'user', // per-user writes, no cross-user blocking
});
```

| Scope                | Serializes across                              | Notes                                                                                                                                                      |
| -------------------- | ---------------------------------------------- | ---------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `script` _(default)_ | All users and all executions of the script     | Always available.                                                                                                                                          |
| `document`           | All executions against the containing document | Only available to scripts running in the context of a containing document. Returns no lock — and therefore no mutual exclusion — from a web app execution. |
| `user`               | The current user's own concurrent executions   | Does not protect against concurrent writes by different users.                                                                                             |

> **Note:**
> Prior to `2.0.0`, `gas-sheetdb` always used a document lock. Because that scope silently provides no lock at all in a web app execution, the default is now `script`. Set `lockScope: 'document'` explicitly to restore the old behavior.

#### Injecting a lock service

`LockService` locks aren't reentrant: if your own code holds a lock and then calls a `gas-sheetdb` write that acquires the same scope, the inner acquire blocks until timeout — nothing in a single-threaded execution can release the outer lock while it waits.

```js
// Deadlocks — the insert() blocks on a 'script' lock this callback holds.
GasLock.withLock('script', () => {
  reportsTable.insert(entry);
});
```

Pass [`gas-lock`](https://github.com/yorsh-co/gas-lock) as `lockService` to fix this. It tracks which scopes the current execution already holds, so nested acquires reuse the outer lock rather than waiting on it:

```js
const sheetDb = new GasSheetDb({
  useActiveSpreadsheet: true,
  lockService: GasLock,
});
```

Omit `lockService` and `gas-sheetdb` acquires `LockService` locks directly, with no reentrancy protection — the standalone default, and unchanged behavior for existing callers.

### Entry Point

#### GasSheetDb

Main entry point for the library.

Constructor options:

```js
new GasSheetDb({
  spreadsheet, // or
  spreadsheetUrl, // or
  spreadsheetId, // or
  useActiveSpreadsheet,
  rowNumbers, // optional instance-wide default
  lockScope, // optional, 'script' | 'document' | 'user'. Defaults to 'script'
  lockService, // optional, e.g. GasLock. Defaults to LockService directly
});
```

The resolved spreadsheet is available as sheetDb.spreadsheet.

##### Methods

```js
GasSheetDb(...).table({ sheetName, rowNumbers, lockScope });
```

#### Table instance returned by GasSheetDb.table(...)

Query Methods:

```js
find({
  withTrashed: boolean, // optional
  onlyTrashed: boolean, // optional
}); // hides deleted entries by default
findWhere(predicateFn); // hides deleted entries by default
findOneWhere(predicateFn); // hides deleted entries by default
findTrashed(predicateFn); // `predicateFn` is optional for `findTrashed()`
```

Insert Methods:

```js
insert(entry);
insertMany(entries);
```

Update Methods:

```js
update(entry);
updateMany(entries);
```

Deletion Methods:

```js
softDelete(entry);
softDeleteMany(entry);

restore(entry);
restoreMany(entries);

`delete(entry);`;
deleteMany(entries);
```

> **Note:** All write operations mutate the original object [metadata](#metadata) and also return the persisted entry(ies). This keeps the in-memory instance up to date with the persisted state, while also allowing for partial entries to be passed to the write operation methods as long as the original entry `_id` property is included.

### Example Workflow

```js
// Create a new user
const sheetDb = new GasSheetDb();
const usersTable = sheetDb.table({ sheetName: '👤 Users' });

usersTable.insert({
  name: 'John',
  role: 'admin',
  permissions: ['users:read', 'users:write'],
});

// Filter users
const admins = usersTable.findWhere((user) => user.role === 'admin');

// Update users
admins.forEach((admin) => {
  admin.lastSeenAt = new Date();
});

usersTable.updateMany(admins);

// Soft delete
const revoked = usersTable.findWhere((user) => user.access === 'revoked');

usersTable.softDeleteMany(revoked);

// Restore
const usersToRestore = usersTable.findTrashed((user) => user.project ===='projectA1');

usersTable.restoreMany(usersToRestore);

// Permanently delete
const usersToDelete =usersTable.findTrashed((user) => user.project ===='projectB1');

usersTable.deleteMany(usersToDelete);
```

## Development

The source code is written in TypeScript.

Release builds are compiled to JavaScript before publishing so the distributed library remains compatible with Google Apps Script.

## Planned features

- Catch entry failures
- Support transactions
- Support pagination in `find` to avoid loading the entire table to memory when not required

## License

MIT

See the `LICENSE` file for details.

## Support

Issues and feature requests are welcome via GitHub Issues.

Maintained by [yorsh-co](https://github.com/yorsh-co).
