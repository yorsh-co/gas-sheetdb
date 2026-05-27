# gas-sheetdb

[![Built with Google Apps Script](https://img.shields.io/badge/Built%20with-Google%20Apps%20Script-4285F4?logo=google&logoColor=white)](https://developers.google.com/apps-script)

## Database-like storage layer for Google Apps Script projects.

> The goal of this project is to use Google Sheets as database storage in Apps Script projects without having to manage sheets manually, store layouts, hard-code ranges, or handle concurrent write logic.

`gas-sheetdb` uses Apps Script's `SpreadsheetApp` to turn sheets in a Google Spreadsheet into tables that store object entries as rows.

Entries can be queried and updated through the `SheetDb` API using methods like `find`, `insert`, and `update`.

> **Disclaimer:**
> This project and [Yorsh](https://github.com/yorsh-co) are independent and are not affiliated with, endorsed by, or associated with Google LLC.

### Features

- Store object entries as rows in sheets
- Query entries using methods like find, findWhere, and findOneWhere
- Insert and update entries using plain JavaScript objects
- Automatically creates missing columns when new properties appear during inserts or updates
- Automatically adds `_id`, `_createdAt`, and `_updatedAt` fields to new entries
- Objects and arrays are JSON serialized
- Supports both bound and standalone spreadsheets
- Uses Apps Script `LockService` for safer concurrent writes (full transactions are not yet supported)
- No external dependencies beyond built-in Apps Script services

### Example Usage

```js
// Create a new user
const users = SheetDb.table(SHEETDB_SHEET_NAMES.USERS);

users.insert({
  name: 'John',
  role: 'admin',
  permissions: ['users:read', 'users:write'],
});

// Filter users
const admins = users.findWhere((user) => user.role === 'admin');

// Update users
admins.forEach((admin) => {
  admin.lastSeenAt = new Date();
});

users.updateMany(admins);
```

## Requirements

### Scopes

`gas-sheetdb` requires one spreadsheet access scope to be added to the parent project's `appsscript.json`.

Use the scope that matches the [spreadsheet access mode](#spreadsheet-access-modes) being used:

| Mode                          | Scope                                                        |
| ----------------------------- | ------------------------------------------------------------ |
| Bound spreadsheet mode        | `"https://www.googleapis.com/auth/spreadsheets.currentonly"` |
| Explicit spreadsheet URL mode | `"https://www.googleapis.com/auth/spreadsheets"`             |

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

`gas-sheetdb` supports two spreadsheet access modes.

### Bound Spreadsheet Mode

Uses the spreadsheet attached to the Apps Script project.

```js
const SHEETDB_USE_ACTIVE_SPREADSHEET = true;
```

Use this mode for container-bound scripts.

### Explicit Spreadsheet URL Mode

Uses a spreadsheet by URL.

```js
const SHEETDB_USE_ACTIVE_SPREADSHEET = false;

const SHEETDB_SPREADSHEET_URL = 'https://docs.google.com/spreadsheets/d/...';
```

Use this mode for:

- standalone Apps Script projects requiring persistent storage
- storing data used by multiple projects
- storing data for a project bound to a different spreadsheet, form or document

## Quick Start

It is recommended to use `gas-sheetdb` together with [Google's `clasp` CLI](https://github.com/google/clasp) for local Apps Script development and git-based workflows. See [Setup instructions with `clasp`](#setup-instructions-with-clasp) for more information.

#### 1. Add the library to your Apps Script project

This repository is intended to be added directly into Apps Script projects using git subtree.

```bash
git subtree add \
  --prefix=src/lib/gas-sheetdb \
  https://github.com/yorsh-co/gas-sheetdb.git \
  main \
  --squash
```

This creates:

```txt
src/lib/gas-sheetdb/
```

#### 2. Configure Apps Script scopes

Add the required spreadsheet scope to the parent project's `appsscript.json`.

See the [Required Apps Script Scopes](#required-apps-script-scopes) section above.

#### 3. Review the library configurations

Example `sheetdb.config.js`:

```js
/**
 * Use the spreadsheet bound to the Apps Script project.
 * When false, `SHEETDB_SPREADSHEET_URL` is used instead.
 */
const SHEETDB_USE_ACTIVE_SPREADSHEET = true;

/**
 * Spreadsheet URL used when
 * `SHEETDB_USE_ACTIVE_SPREADSHEET` is false.
 */
const SHEETDB_SPREADSHEET_URL = '';

/**
 * Known sheet names.
 */
const SHEETDB_SHEET_NAMES = Object.freeze({
  SYSTEM: Object.freeze({
    CONFIG: '.config',
    ERRORS: '.errors',
  }),

  USERS: '👤 Users',
});

/**
 * Sheet row numbers.
 */
const SHEETDB_ROW_NUMBERS = {
  headers: 1,
  firstData: 2,
};
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
  main \
  --squash
```

This creates:

```txt
src/lib/gas-sheetdb/
```

#### 6. Push local files to Apps Script

```bash
clasp push
```

### 7. Configure Apps Script scopes

Add the required spreadsheet scope to the parent project's `appsscript.json`.

See the [Required Apps Script Scopes](#required-apps-script-scopes) section above.

### 8. Review the library configurations

Example `sheetdb.config.js`:

```js
/**
 * Use the spreadsheet bound to the Apps Script project.
 * When false, `SHEETDB_SPREADSHEET_URL` is used instead.
 */
const SHEETDB_USE_ACTIVE_SPREADSHEET = true;

/**
 * Spreadsheet URL used when
 * `SHEETDB_USE_ACTIVE_SPREADSHEET` is false.
 */
const SHEETDB_SPREADSHEET_URL = '';

/**
 * Known sheet names.
 */
const SHEETDB_SHEET_NAMES = Object.freeze({
  SYSTEM: Object.freeze({
    CONFIG: '.config',
    ERRORS: '.errors',
  }),

  USERS: '👤 Users',
});

/**
 * Sheet row numbers.
 */
const SHEETDB_ROW_NUMBERS = {
  headers: 1,
  firstData: 2,
};
```

## Basic Usage

### Create a Table

```js
const usersTable = SheetDb.table(SHEETDB_SHEET_NAMES.USERS);
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

### Read All Entries

```js
const users = usersTable.find();
```

### Filter Entries

```js
const admins = usersTable.findWhere((entry) => entry.role === 'admin');
```

### Find a Single Entry

```js
const user = usersTable.findOneWhere((entry) => entry._id === 'abc123');
```

### Update an Entry

> **Note:**
> The `update` and `updateMany` methods currently only accepts entries that were returned from the `find`, `findWhere` or `findOneWhere` and updated in runtime. Updating entries by a specific property such as `_id` is not currently supported.

```js
const user = usersTable.findOneWhere((entry) => entry._id === 'abc123');

user.active = false;

usersTable.update(user);
```

### Update Multiple Entries

> **Note:**
> The `update` and `updateMany` methods currently only accepts entries that were returned from the `find`, `findWhere` or `findOneWhere` and updated in runtime. Updating entries by a specific property such as `_id` is not currently supported.

```js
const users = usersTable.findWhere((entry) => entry.role === 'editor');

users.forEach((user) => {
  user.active = true;
});

usersTable.updateMany(users);
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
}
```

The `_updatedAt` field is refreshed whenever an entry is updated.

Entries returned from tables also include non-persistent runtime metadata:

```js
{
  _runtime: {
    rowNumber: 5,
  },
}
```

`_runtime.rowNumber` is used internally when updating rows and is not stored in the sheet.

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

### Entry Point

#### SheetDb

Main entry point for the library.

Methods:

```js
SheetDb.getSpreadsheet();
SheetDb.table(sheetName);
```

#### Table instance returned by SheetDb.table(...)

Query Methods:

```js
find();
findWhere(predicateFn);
findOneWhere(predicateFn);
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

### Example Workflow

```js
// Create a new user
const users = SheetDb.table(SHEETDB_SHEET_NAMES.USERS);

users.insert({
  name: 'John',
  role: 'admin',
  permissions: ['users:read', 'users:write'],
});

// Filter users
const admins = users.findWhere((user) => user.role === 'admin');

// Update users
admins.forEach((admin) => {
  admin.lastSeenAt = new Date();
});

users.updateMany(admins);
```

## Planned features

- Support entry deletion
- Support updating values by `_id` instead of relying on `_runtime.rowNumber`

## License

MIT

See the `LICENSE` file for details.

## Support

Issues and feature requests are welcome via GitHub Issues.

Maintained by [yorsh-co](https://github.com/yorsh-co).
