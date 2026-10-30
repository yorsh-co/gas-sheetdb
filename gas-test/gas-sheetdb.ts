type TestSpreadsheet = GoogleAppsScript.Spreadsheet.Spreadsheet;
type TestSheet = GoogleAppsScript.Spreadsheet.Sheet;
type TestTable = ReturnType<GasSheetDb['table']>;

/** Strings Sheets would parse if they were written without escaping. */
const SMOKE_STRINGS = [
  '=1+1',
  '=IMPORTXML("https://example.com","//a")',
  '+1',
  '-1',
  '@x',
  '0012',
  '1e3',
  'TRUE',
  'false',
  '2024-01-02',
  '1-2',
  "'quoted",
  "''double",
  "'",
  '__JSON__x',
  '__JSON__{"isAdmin":true}',
  ' padded ',
  'plain text',
];

/** Creates a database over the test spreadsheet. */
const newDb = (spreadsheet: TestSpreadsheet): GasSheetDb =>
  new GasSheetDb({ spreadsheet });

/** The `code` of a thrown value, or undefined when it has none. */
const codeOf = (err: unknown): unknown =>
  typeof err === 'object' && err !== null && 'code' in err
    ? err.code
    : undefined;

/** The message of a thrown value. */
const messageOf = (err: unknown): string =>
  err instanceof Error ? err.message : String(err);

/** 1-based column of a header key. */
function columnOf(sheet: TestSheet, key: string): number {
  const header = sheet
    .getRange(1, 1, 1, sheet.getLastColumn())
    .getValues()[0]
    ?.map(String);
  assert(header, 'the sheet has no header row');
  const column = header.indexOf(key) + 1;
  assert(column > 0, `no "${key}" column`);
  return column;
}

/** 1-based row holding an `_id`. */
function rowOf(sheet: TestSheet, id: unknown): number {
  const column = columnOf(sheet, '_id');
  const ids = sheet
    .getRange(1, column, sheet.getLastRow(), 1)
    .getValues()
    .map((row) => row[0] as unknown);
  const row = ids.indexOf(id) + 1;
  assert(row > 0, `no row with _id ${String(id)}`);
  return row;
}

/** A1 column letter(s) for a 1-based column. */
function letterOf(sheet: TestSheet, column: number): string {
  return sheet.getRange(1, column).getA1Notation().replace(/\d+$/, '');
}

/** Runs `fn` holding a raw script lock; reports whether the lock survived. */
function withRawOuterLock(fn: () => void): string {
  const outer = LockService.getScriptLock();
  outer.waitLock(10000);

  try {
    fn();
    return outer.hasLock()
      ? 'outer lock still held after the write'
      : 'outer lock was released by the nested write (documented hazard: use GasLock.withLock for outer sections)';
  } finally {
    if (outer.hasLock()) outer.releaseLock();
  }
}

/** Asserts the table is empty, still accepts an insert, and reports the grid. */
function assertEmptyThenInsert(table: TestTable): string {
  const sheet = table.sheet;
  assert(
    table.find({ withTrashed: true }).length === 0,
    'entries remain after deleting every entry',
  );
  assert(
    sheet.getMaxRows() > sheet.getFrozenRows(),
    'no non-frozen row is left',
  );

  table.insert({ n: 4 });
  assert(
    table.find().length === 1,
    'an insert after deleting every entry failed',
  );

  return `maxRows ${sheet.getMaxRows()}, frozen ${sheet.getFrozenRows()}`;
}

/** Asserts the ARRAYFORMULA anchor is intact and the spill reads `expected`. */
function assertSpill(
  sheet: TestSheet,
  column: number,
  expected: number[],
): void {
  const anchor = sheet.getRange(2, column).getFormula();
  assert(anchor, 'the ARRAYFORMULA anchor was overwritten');

  const values = sheet
    .getRange(2, column, expected.length, 1)
    .getValues()
    .map((row) => row[0] as unknown);
  assert(
    JSON.stringify(values) === JSON.stringify(expected),
    `the spill now reads ${JSON.stringify(values)}`,
  );
}

/** Raw LockService ownership rules that the 2.0.0 locking design relies on. */
test('LockService: a nested acquire in the same execution succeeds immediately', () => {
  const outer = LockService.getScriptLock();
  outer.waitLock(10000);
  try {
    const inner = LockService.getScriptLock();
    assert(
      inner.tryLock(0),
      'nested tryLock(0) returned false: nested acquires block',
    );
    inner.releaseLock();
  } finally {
    if (outer.hasLock()) outer.releaseLock();
  }
});

test('LockService: releasing a nested handle ends the outer critical section', () => {
  const outer = LockService.getScriptLock();
  outer.waitLock(10000);
  try {
    const inner = LockService.getScriptLock();
    inner.tryLock(0);
    inner.releaseLock();
    assert(
      !outer.hasLock(),
      'outer.hasLock() is still true after the nested release',
    );
  } finally {
    if (outer.hasLock()) outer.releaseLock();
  }
});

/**
 * Strings must round-trip as text through the apostrophe escape. The checks
 * share one `insertMany`, run on first use; if it failed, each check reports
 * that instead of a misleading mismatch.
 */
type StringsFixture = {
  table: TestTable;
  inserted: GasSheetDbEntry[];
  byId: Map<unknown, GasSheetDbEntry>;
};

let strings: StringsFixture | Error | undefined;

function stringsFixture(
  spreadsheet: TestSpreadsheet,
  prefix: string,
): StringsFixture | Error {
  if (strings === undefined) {
    try {
      const table = newDb(spreadsheet).table({
        sheetName: `${prefix}smoke-strings`,
      });
      const inserted = table.insertMany(
        SMOKE_STRINGS.map((value) => ({ value })),
      );
      const byId = new Map(table.find().map((entry) => [entry._id, entry]));
      strings = { table, inserted, byId };
    } catch (err) {
      strings = err instanceof Error ? err : new Error(String(err));
    }
  }
  return strings;
}

/** The shared fixture, failing the calling test when its insert failed. */
function requireStrings(
  spreadsheet: TestSpreadsheet,
  prefix: string,
): StringsFixture {
  const fixture = stringsFixture(spreadsheet, prefix);
  assert(
    !(fixture instanceof Error),
    `insertMany failed: ${fixture instanceof Error ? fixture.message : ''}`,
  );
  return fixture;
}

test('insertMany stores escaped strings without throwing', ({
  spreadsheet,
  prefix,
}) => {
  const { inserted } = requireStrings(spreadsheet, prefix);
  assert(
    inserted.length === SMOKE_STRINGS.length,
    `${inserted.length} rows, expected ${SMOKE_STRINGS.length}`,
  );
});

SMOKE_STRINGS.forEach((expected, index) => {
  test(`round-trips ${JSON.stringify(expected)} as text`, ({
    spreadsheet,
    prefix,
  }) => {
    const { table, inserted, byId } = requireStrings(spreadsheet, prefix);
    const sheet = table.sheet;
    const id = inserted[index]?._id;

    const read = byId.get(id);
    assert(
      read !== undefined && read.value === expected,
      `read back ${JSON.stringify(read?.value)} (${typeof read?.value})`,
    );

    const formula = sheet
      .getRange(rowOf(sheet, id), columnOf(sheet, 'value'))
      .getFormula();
    assert(!formula, `the cell holds the formula ${formula}`);
  });
});

test('update writes a formula-looking string as text', ({
  spreadsheet,
  prefix,
}) => {
  const { table } = requireStrings(spreadsheet, prefix);
  const sheet = table.sheet;
  const entry = table.find()[0];
  assert(entry !== undefined, 'the table has no entries');
  entry.value = '=2+2';
  table.update(entry);

  const read = table.findOneWhere((e) => e._id === entry._id);
  assert(read?.value === '=2+2', `read back ${JSON.stringify(read?.value)}`);
  assert(
    !sheet
      .getRange(rowOf(sheet, entry._id), columnOf(sheet, 'value'))
      .getFormula(),
    'the cell holds a formula',
  );
});

/**
 * The 50,000-character cell limit: the apostrophe escape counts towards it,
 * so text holds at most 49,999 characters.
 */
test('stores a 49,999-character string (the limit for text)', ({
  spreadsheet,
  prefix,
}) => {
  const table = newDb(spreadsheet).table({
    sheetName: `${prefix}smoke-limit`,
  });
  const entry = table.insert({ text: 'x'.repeat(49999) });
  const read = table.findOneWhere((e) => e._id === entry._id);
  const text = read?.text;
  assert(
    typeof text === 'string' && text.length === 49999,
    `read back ${typeof text === 'string' ? text.length : typeof text} characters`,
  );
});

test('rejects a 50,000-character string before writing', ({
  spreadsheet,
  prefix,
}) => {
  const table = newDb(spreadsheet).table({
    sheetName: `${prefix}smoke-limit`,
  });
  const before = table.sheet.getLastRow();

  try {
    table.insert({ text: 'x'.repeat(50000) });
  } catch (err) {
    assert(
      codeOf(err) === 'VALIDATION_ERROR',
      `threw ${err instanceof Error ? err.name : typeof err}: ${messageOf(err)}`,
    );
    assert(table.sheet.getLastRow() === before, 'a row was written');
    return;
  }

  throw new Error('did not throw');
});

/** Writes nested inside an outer lock held by GasLock or raw LockService. */
test('GasLock.withLock around a write raises no lost-lock error', ({
  spreadsheet,
  prefix,
}) => {
  const table = newDb(spreadsheet).table({
    sheetName: `${prefix}smoke-locking`,
  });
  GasLock.withLock('script', () => {
    table.insert({ step: 'gaslock-outer' });
    assert(GasLock.isHeld('script'), 'GasLock no longer holds the scope');
  });
});

info(
  'raw LockService outer lock + fallback lock service',
  ({ spreadsheet, prefix }) => {
    const fallback = new GasSheetDb({
      spreadsheet,
      lockService: _GAS_SHEETDB_FALLBACK_LOCK_SERVICE,
    }).table({ sheetName: `${prefix}smoke-locking` });

    return withRawOuterLock(() =>
      fallback.insert({ step: 'raw-outer-fallback' }),
    );
  },
);

info(
  'raw LockService outer lock + default (GasLock) service',
  ({ spreadsheet, prefix }) => {
    const table = newDb(spreadsheet).table({
      sheetName: `${prefix}smoke-locking`,
    });
    return withRawOuterLock(() => table.insert({ step: 'raw-outer-gaslock' }));
  },
);

info(
  'lockScope "document" in this execution context',
  ({ spreadsheet, prefix }) => {
    try {
      newDb(spreadsheet)
        .table({ sheetName: `${prefix}smoke-locking`, lockScope: 'document' })
        .insert({ step: 'document-scope' });
      return 'document lock available (container-bound context)';
    } catch (err) {
      return `unavailable: ${messageOf(err)}`;
    }
  },
);

/** Deleting every entry must never trip "cannot delete all non-frozen rows". */
test('deleteWhere removes every entry (header row frozen)', ({
  spreadsheet,
  prefix,
}) => {
  const table = newDb(spreadsheet).table({
    sheetName: `${prefix}smoke-delete-a`,
  });
  table.insertMany([{ n: 1 }, { n: 2 }, { n: 3 }]);
  table.deleteWhere(() => true);
  assertEmptyThenInsert(table);
});

test('deleteMany removes every entry with 2 frozen rows', ({
  spreadsheet,
  prefix,
}) => {
  const table = newDb(spreadsheet).table({
    sheetName: `${prefix}smoke-delete-b`,
  });
  table.insertMany([{ n: 1 }, { n: 2 }, { n: 3 }]);
  table.sheet.setFrozenRows(2);
  table.deleteMany(table.find());
  assertEmptyThenInsert(table);
});

test('deleting every entry when the body fills the sheet', ({
  spreadsheet,
  prefix,
}) => {
  const table = newDb(spreadsheet).table({
    sheetName: `${prefix}smoke-delete-c`,
  });
  table.insertMany([{ n: 1 }, { n: 2 }, { n: 3 }]);

  const sheet = table.sheet;
  const spare = sheet.getMaxRows() - sheet.getLastRow();
  if (spare > 0) sheet.deleteRows(sheet.getLastRow() + 1, spare);

  table.deleteWhere(() => true);
  assertEmptyThenInsert(table);
});

/** Formula columns added by a person must survive gas-sheetdb writes. */
test('a per-row formula column survives find → update', ({
  spreadsheet,
  prefix,
}) => {
  const table = newDb(spreadsheet).table({
    sheetName: `${prefix}smoke-formulas`,
  });
  table.insertMany([
    { name: 'a', qty: 1 },
    { name: 'b', qty: 2 },
  ]);

  const sheet = table.sheet;
  const column = sheet.getLastColumn() + 1;
  sheet.getRange(1, column).setValue('calc');
  sheet.getRange(2, column, 2, 1).setFormulas([['=ROW()*10'], ['=ROW()*10']]);
  SpreadsheetApp.flush();

  const entry = table.find().find((e) => e.name === 'a');
  assert(entry !== undefined, 'entry "a" not found');
  assert(entry.calc === 20, `calc read as ${JSON.stringify(entry.calc)}`);

  entry.qty = 5;
  table.update(entry);

  const formulas = sheet.getRange(2, column, 2, 1).getFormulas();
  assert(
    formulas[0]?.[0] === '=ROW()*10' && formulas[1]?.[0] === '=ROW()*10',
    `formulas are now ${JSON.stringify(formulas)}`,
  );
  assert(
    table.findOneWhere((e) => e.name === 'a')?.qty === 5,
    'qty was not updated',
  );
});

/**
 * An ARRAYFORMULA anchored in the first data row fills the rows below it, so
 * those cells must not be overwritten by updates or inserts.
 */
test('an ARRAYFORMULA spill survives find → update', ({
  spreadsheet,
  prefix,
}) => {
  const table = newDb(spreadsheet).table({
    sheetName: `${prefix}smoke-spill`,
  });
  table.insertMany([{ name: 'a' }, { name: 'b' }, { name: 'c' }]);

  const sheet = table.sheet;
  const column = sheet.getLastColumn() + 1;
  const id = letterOf(sheet, columnOf(sheet, '_id'));
  sheet.getRange(1, column).setValue('rowNo');
  sheet.getRange(2, column).setFormula(`=ARRAYFORMULA(ROW(${id}2:${id}4))`);
  SpreadsheetApp.flush();

  // 'b' sits in row 3, inside the spill, not on its anchor
  const entry = table.find().find((e) => e.name === 'b');
  assert(entry !== undefined, 'entry "b" not found');
  entry.name = 'b2';
  table.update(entry);
  SpreadsheetApp.flush();

  assertSpill(sheet, column, [2, 3, 4]);
});

test('an open-ended ARRAYFORMULA spill survives an insert', ({
  spreadsheet,
  prefix,
}) => {
  const table = newDb(spreadsheet).table({
    sheetName: `${prefix}smoke-spill-open`,
  });
  table.insertMany([{ name: 'a' }, { name: 'b' }]);

  const sheet = table.sheet;
  const column = sheet.getLastColumn() + 1;
  const id = letterOf(sheet, columnOf(sheet, '_id'));
  sheet.getRange(1, column).setValue('rowNo');
  sheet
    .getRange(2, column)
    .setFormula(`=ARRAYFORMULA(IF(LEN(${id}2:${id}), ROW(${id}2:${id}), ""))`);
  SpreadsheetApp.flush();

  const lastRowBefore = sheet.getLastRow();
  const entry = table.insert({ name: 'c' });
  SpreadsheetApp.flush();

  const row = rowOf(sheet, entry._id);
  assert(
    row === 4,
    `the new entry landed in row ${row} (getLastRow before the insert: ${lastRowBefore})`,
  );
  assertSpill(sheet, column, [2, 3, 4]);
  assert(
    table.find().length === 3,
    `find() returned ${table.find().length} entries`,
  );
});
