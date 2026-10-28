import { fileURLToPath } from 'node:url';
import { createGasLoader } from 'gas-dev-config/test';
import { createMocks, type MockOptions, type Mocks } from './mocks';

/** Source files in load order; later files depend on earlier ones. */
const SOURCE_FILES = [
  'gas-sheetdb.types.ts',
  'gas-sheetdb.constants.ts',
  'gas-sheetdb.codec.ts',
  'gas-sheetdb.schema.ts',
  'gas-sheetdb.table.ts',
  'gas-sheetdb.class.ts',
];

const load = createGasLoader<{ GasSheetDb: typeof GasSheetDb }>({
  files: SOURCE_FILES.map((file) => `src/${file}`),
  exports: ['GasSheetDb'],
  root: fileURLToPath(new URL('..', import.meta.url)),
});

/**
 * Evaluates the library sources in a fresh context wired to fake Apps Script
 * services, so each call is isolated from every other. `Error` and `Date` are
 * shared with the test realm, so thrown errors and stored dates satisfy
 * `instanceof` checks in assertions.
 */
export function loadGasSheetDb(options: MockOptions = {}): {
  GasSheetDb: typeof GasSheetDb;
  mocks: Mocks;
} {
  const mocks = createMocks(options);
  const { SpreadsheetApp, LockService, Utilities, console } = mocks;
  const { GasSheetDb } = load({
    SpreadsheetApp,
    LockService,
    Utilities,
    console,
    Error,
    Date,
  });

  return { GasSheetDb, mocks };
}
