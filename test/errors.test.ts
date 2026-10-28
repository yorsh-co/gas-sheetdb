import { describe, expect, it } from 'vitest';
import { loadGasSheetDb } from './harness';
import { createLogger } from './helpers';

/**
 * Test-side stand-ins for the gas-error classes, matching its peer shape.
 * Declared as Apps Script would see them: plain globals.
 */
class FakeGasError extends Error {
  statusCode: number;
  code: string;
  details: unknown;
  isOperational = true;

  constructor(
    message: string,
    options: { statusCode?: number; code?: string; details?: unknown } = {},
  ) {
    super(message);
    this.name = this.constructor.name;
    this.statusCode = options.statusCode ?? 500;
    this.code = options.code ?? 'INTERNAL_ERROR';
    this.details = options.details;
  }
}

class GasValidationError extends FakeGasError {
  constructor(message = 'Invalid request', details?: unknown) {
    super(message, { statusCode: 400, code: 'VALIDATION_ERROR', details });
  }
}

class GasNotFoundError extends FakeGasError {
  constructor(message = 'Resource not found') {
    super(message, { statusCode: 404, code: 'NOT_FOUND' });
  }
}

const GAS_ERROR_GLOBALS = {
  GasError: FakeGasError,
  GasValidationError,
  GasNotFoundError,
};

/** A seeded table; `globals` are the optional peer libraries in scope. */
function setupTable(globals: Record<string, unknown> = {}) {
  const { GasSheetDb, mocks } = loadGasSheetDb({}, globals);
  const db = new GasSheetDb({
    useActiveSpreadsheet: true,
    logger: createLogger(),
  });
  const table = db.table({ sheetName: 'people' });
  table.insert({ name: 'Ann' });
  return { GasSheetDb, mocks, db, table };
}

/** Runs `fn` and returns what it throws; fails the test when it does not. */
function thrown(fn: () => unknown): Error {
  try {
    fn();
  } catch (err) {
    return err as Error;
  }
  throw new Error('Expected the callback to throw');
}

describe('without gas-error', () => {
  const missingId = [
    ['update', (t: ReturnType<typeof setupTable>['table']) => t.update({})],
    [
      'softDelete',
      (t: ReturnType<typeof setupTable>['table']) => t.softDelete({}),
    ],
    ['restore', (t: ReturnType<typeof setupTable>['table']) => t.restore({})],
  ] as const;

  it.each(missingId)(
    '%s without an _id is a 400 validation error',
    (_, run) => {
      const err = thrown(() => run(setupTable().table));

      expect(err).toBeInstanceOf(Error);
      expect(err.name).toBe('GasSheetDbError');
      expect(err).toMatchObject({
        message: '[GasSheetDb] Cannot update entry without "_id"',
        statusCode: 400,
        code: 'VALIDATION_ERROR',
        isOperational: true,
      });
      expect((err as unknown as { details: unknown }).details).toBeUndefined();
    },
  );

  it('delete without an _id is a 400 validation error', () => {
    const err = thrown(() => setupTable().table.deleteOne({}));

    expect(err.name).toBe('GasSheetDbError');
    expect(err).toMatchObject({
      message: '[GasSheetDb] Cannot delete entry without "_id"',
      statusCode: 400,
      code: 'VALIDATION_ERROR',
    });
  });

  it('an unknown _id is a 404 not-found error', () => {
    const err = thrown(() => setupTable().table.update({ _id: 'nope' }));

    expect(err).toBeInstanceOf(Error);
    expect(err.name).toBe('GasSheetDbError');
    expect(err).toMatchObject({
      message: '[GasSheetDb] Entry not found: nope',
      statusCode: 404,
      code: 'NOT_FOUND',
      isOperational: true,
    });
  });

  it('keeps a stack trace', () => {
    const err = thrown(() => setupTable().table.update({ _id: 'nope' }));

    expect(err.stack).toContain('GasSheetDbError');
  });
});

describe('with gas-error', () => {
  it('throws GasValidationError for a missing _id', () => {
    const err = thrown(() =>
      setupTable(GAS_ERROR_GLOBALS).table.update({ name: 'x' }),
    );

    expect(err).toBeInstanceOf(GasValidationError);
    expect(err).toBeInstanceOf(FakeGasError);
    expect(err).toMatchObject({
      message: '[GasSheetDb] Cannot update entry without "_id"',
      statusCode: 400,
      code: 'VALIDATION_ERROR',
    });
  });

  it('throws GasNotFoundError for an unknown _id', () => {
    const err = thrown(() =>
      setupTable(GAS_ERROR_GLOBALS).table.deleteOne({ _id: 'nope' }),
    );

    expect(err).toBeInstanceOf(GasNotFoundError);
    expect(err).toMatchObject({
      message: '[GasSheetDb] Entry not found: nope',
      statusCode: 404,
      code: 'NOT_FOUND',
    });
  });

  it('falls back per error class when only some are defined', () => {
    const { table } = setupTable({ GasNotFoundError });

    expect(thrown(() => table.update({}))).toMatchObject({
      name: 'GasSheetDbError',
      statusCode: 400,
    });
    expect(thrown(() => table.update({ _id: 'nope' }))).toBeInstanceOf(
      GasNotFoundError,
    );
  });
});

describe('configuration errors', () => {
  it.each([{}, GAS_ERROR_GLOBALS])('stay plain Error (%#)', (globals) => {
    const { GasSheetDb } = loadGasSheetDb({}, globals);

    const err = thrown(
      () =>
        new GasSheetDb({
          useActiveSpreadsheet: true,
          lockScope: 'global' as GasSheetDbLockScope,
        }),
    );

    expect(err.constructor).toBe(Error);
    expect(err.message).toMatch(/^\[GasSheetDb\] Invalid "lockScope"/);
    expect(err).not.toHaveProperty('statusCode');
  });

  it('prefixes the missing _id column error', () => {
    const { db } = setupTable();
    const fresh = db.table({ sheetName: 'fresh' });

    // Unreachable through the public API, which adds the column first.
    const err = thrown(() =>
      (
        fresh as unknown as { _mapTableBodyRowIndexesById(): unknown }
      )._mapTableBodyRowIndexesById(),
    );

    expect(err.constructor).toBe(Error);
    expect(err.message).toMatch(/^\[GasSheetDb\] Cannot map rows by "_id"/);
  });
});
