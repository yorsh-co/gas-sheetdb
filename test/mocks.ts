import { vi, type Mock } from 'vitest';

export const ROWS_OUT_OF_BOUNDS_MESSAGE = 'Those rows are out of bounds.';
export const COLUMNS_OUT_OF_BOUNDS_MESSAGE = 'Those columns are out of bounds.';
export const DELETE_ALL_ROWS_MESSAGE =
  'Sorry, it is not possible to delete all non-frozen rows.';

/** Dimensions of a sheet created by `Spreadsheet#insertSheet`. */
const NEW_SHEET_ROWS = 1000;
const NEW_SHEET_COLUMNS = 26;

/** Longest input Sheets accepts in a single cell, apostrophe included. */
const MAX_INPUT_LENGTH = 50000;

/** Message Sheets throws when an input exceeds `MAX_INPUT_LENGTH`. */
export const CELL_LIMIT_MESSAGE = `Your input contains more than the maximum of ${MAX_INPUT_LENGTH} characters in a single cell.`;

/** Coordinates of a range operation. */
export interface RangeCall {
  row: number;
  col: number;
  numRows: number;
  numCols: number;
}

/** A stored cell: its value, and its formula when it holds one. */
interface Cell {
  value: unknown;
  formula: string;
}

/** An array formula's output: its anchor cell and the rows it fills. */
interface Spill {
  row: number;
  col: number;
  numRows: number;
  formula: string;
  broken: boolean;
}

/**
 * Computes the value of a formula cell. Receives the formula as stored
 * (leading `=` included) and the sheet that holds it.
 */
export type FormulaEvaluator = (formula: string, sheet: FakeSheet) => unknown;

/** Default evaluator: every formula evaluates to an empty string. */
const DEFAULT_EVALUATOR: FormulaEvaluator = () => '';

const NUMBER_PATTERN = /^[+-]?(\d+\.?\d*|\.\d+)(e[+-]?\d+)?$/i;
const ISO_DATE_PATTERN = /^(\d{4})-(\d{2})-(\d{2})$/;

/**
 * The single place that models how Sheets parses a value written with
 * `setValues` (USER_ENTERED semantics). Non-string values are stored as-is;
 * a string is interpreted, in order, as:
 *
 * 1. a text literal when it starts with `'` (the apostrophe is dropped),
 * 2. a formula when it starts with `=`,
 * 3. a number when it looks numeric (`0012`, `-3.5`, `1e3`),
 * 4. a boolean for `TRUE`/`FALSE`, case-insensitive,
 * 5. a `Date` (UTC midnight) for a valid `YYYY-MM-DD`,
 * 6. otherwise plain text.
 */
export function coerceUserEntered(input: unknown): Cell {
  if (typeof input !== 'string') return { value: input, formula: '' };

  if (input.startsWith("'")) return { value: input.slice(1), formula: '' };
  if (input.startsWith('=')) return { value: '', formula: input };
  if (NUMBER_PATTERN.test(input)) return { value: Number(input), formula: '' };

  const upper = input.toUpperCase();
  if (upper === 'TRUE') return { value: true, formula: '' };
  if (upper === 'FALSE') return { value: false, formula: '' };

  const iso = ISO_DATE_PATTERN.exec(input);
  if (iso) {
    const [, year, month, day] = iso.map(Number) as [
      number,
      number,
      number,
      number,
    ];
    const date = new Date(Date.UTC(year, month - 1, day));
    if (date.getUTCMonth() === month - 1 && date.getUTCDate() === day) {
      return { value: date, formula: '' };
    }
  }

  return { value: input, formula: '' };
}

/** Option switches for `FakeSheet`. */
export interface FakeSheetOptions {
  evaluate?: FormulaEvaluator;
}

/**
 * In-memory stand-in for an Apps Script `Sheet`, backed by a sparse grid of
 * `maxRows` x `maxColumns` cells.
 *
 * Models the Sheets behaviour the library depends on:
 * - `getLastRow`/`getLastColumn` report the last row/column holding any
 *   content (a formula counts), and `getDataRange` spans at least 1x1;
 * - `getRange` throws when the range exceeds the grid, and writes never grow
 *   it: rows and columns must be added with `insertRowsAfter` and
 *   `insertColumnsAfter`;
 * - `deleteRows` refuses to remove every non-frozen row;
 * - `setValues` interprets strings as USER_ENTERED does, see
 *   `coerceUserEntered`, and throws Sheets' own error for a string input
 *   longer than 50,000 characters, a leading apostrophe included.
 *   `getValues` returns formula cells' values from the `evaluate` hook and
 *   `getFormulas` returns their formulas;
 * - `spill` models an array formula's output, see there.
 */
export class FakeSheet {
  maxRows = NEW_SHEET_ROWS;
  maxColumns = NEW_SHEET_COLUMNS;
  evaluate: FormulaEvaluator;
  frozenRows = 0;
  private grid: (Cell | undefined)[][] = [];
  private spills: Spill[] = [];
  /** Anchors (`[row, col]`) of the spills a write has broken. */
  readonly brokenSpills: [number, number][] = [];
  readonly calls = {
    setValues: [] as RangeCall[],
    clearContent: [] as RangeCall[],
    setFontWeight: [] as RangeCall[],
    applyRowBanding: [] as RangeCall[],
  };

  constructor(
    private readonly name: string,
    options: FakeSheetOptions = {},
  ) {
    this.evaluate = options.evaluate ?? DEFAULT_EVALUATOR;
  }

  getSheetName = (): string => this.name;

  getMaxRows = (): number => this.maxRows;

  getMaxColumns = (): number => this.maxColumns;

  /** Last row with any content, or 0 for an empty sheet. */
  getLastRow = (): number => {
    for (let r = this.grid.length - 1; r >= 0; r--) {
      if (
        this.grid[r]?.some((cell, c) => this.hasContentAt(r + 1, c + 1, cell))
      )
        return r + 1;
    }
    return 0;
  };

  /** Last column with any content, or 0 for an empty sheet. */
  getLastColumn = (): number => {
    let last = 0;
    for (const [r, row] of this.grid.entries()) {
      row?.forEach((cell, c) => {
        if (this.hasContentAt(r + 1, c + 1, cell)) last = Math.max(last, c + 1);
      });
    }
    return last;
  };

  getDataRange = (): FakeRange =>
    this.getRange(
      1,
      1,
      Math.max(this.getLastRow(), 1),
      Math.max(this.getLastColumn(), 1),
    );

  getRange = (
    row: number,
    col: number,
    numRows = 1,
    numCols = 1,
  ): FakeRange => {
    if (row < 1 || col < 1 || numRows < 1 || numCols < 1) {
      throw new Error(
        'The coordinates or dimensions of the range are invalid.',
      );
    }
    if (row + numRows - 1 > this.maxRows) {
      throw new Error(ROWS_OUT_OF_BOUNDS_MESSAGE);
    }
    if (col + numCols - 1 > this.maxColumns) {
      throw new Error(COLUMNS_OUT_OF_BOUNDS_MESSAGE);
    }
    return new FakeRange(this, { row, col, numRows, numCols });
  };

  insertRowsAfter = vi.fn((afterPosition: number, howMany: number): this => {
    if (afterPosition < 0 || afterPosition > this.maxRows) {
      throw new Error(ROWS_OUT_OF_BOUNDS_MESSAGE);
    }
    this.grid.length = Math.max(this.grid.length, afterPosition);
    this.grid.splice(afterPosition, 0, ...Array(howMany).fill(undefined));
    this.maxRows += howMany;
    return this;
  });

  insertColumnsAfter = vi.fn((afterPosition: number, howMany: number): this => {
    if (afterPosition < 0 || afterPosition > this.maxColumns) {
      throw new Error(COLUMNS_OUT_OF_BOUNDS_MESSAGE);
    }
    for (const row of this.grid) {
      if (!row) continue;
      row.length = Math.max(row.length, afterPosition);
      row.splice(afterPosition, 0, ...Array(howMany).fill(undefined));
    }
    this.maxColumns += howMany;
    return this;
  });

  deleteRows = vi.fn((rowPosition: number, howMany: number): void => {
    if (rowPosition < 1 || rowPosition + howMany - 1 > this.maxRows) {
      throw new Error(ROWS_OUT_OF_BOUNDS_MESSAGE);
    }
    if (
      rowPosition <= this.frozenRows + 1 &&
      rowPosition + howMany > this.maxRows
    ) {
      throw new Error(DELETE_ALL_ROWS_MESSAGE);
    }
    this.grid.splice(rowPosition - 1, howMany);
    this.maxRows -= howMany;
  });

  setFrozenRows = vi.fn((rows: number): void => {
    this.frozenRows = rows;
  });

  getFrozenRows = (): number => this.frozenRows;

  getBandings = (): unknown[] => [];

  /** Whether a cell counts as content: its own, or an intact spill's output. */
  private hasContentAt(row: number, col: number, cell: Cell | undefined) {
    return (
      hasContent(cell) ||
      this.spills.some(
        (s) =>
          !s.broken && s.col === col && row >= s.row && row < s.row + s.numRows,
      )
    );
  }

  /**
   * Places an array formula at `row`, `col` whose output is `values`, one per
   * row downwards. An open-ended formula such as
   * `ARRAYFORMULA(IF(LEN(B2:B), ROW(B2:B), ""))` outputs `''` down to the last
   * row of the sheet, and Sheets counts those cells as content.
   *
   * The anchor holds `formula` and shows `values[0]`; the cells below hold
   * their output without a formula of their own. Every cell counts as content
   * for `getLastRow`, `getLastColumn` and `getDataRange`, even when its value
   * is `''`. Writing any value into a spilled cell breaks the spill, as Sheets
   * does: the other cells clear and the anchor shows `#REF!`, recorded in
   * `brokenSpills`. A spill does not follow row insertions or deletions.
   */
  spill(row: number, col: number, values: unknown[], formula: string): void {
    this.spills.push({
      row,
      col,
      numRows: values.length,
      formula,
      broken: false,
    });
    values.forEach((value, i) =>
      this.writeCell(row + i, col, {
        value,
        formula: i === 0 ? formula : '',
      }),
    );
  }

  /** @internal Breaks the spill that holds a cell, ahead of a write to it. */
  breakSpillAt(row: number, col: number): void {
    const spill = this.spills.find(
      (s) =>
        !s.broken && s.col === col && row >= s.row && row < s.row + s.numRows,
    );
    if (!spill) return;

    spill.broken = true;
    for (let r = spill.row + 1; r < spill.row + spill.numRows; r++) {
      this.writeCell(r, col, undefined);
    }
    this.writeCell(spill.row, col, { value: '#REF!', formula: spill.formula });
    this.brokenSpills.push([spill.row, col]);
  }

  /**
   * @internal Whether a cell shows the value it stores instead of evaluating
   * its formula: a spill's anchor, and the cells of a spill still intact.
   */
  showsStoredValue(row: number, col: number): boolean {
    return this.spills.some(
      (s) =>
        s.col === col &&
        (row === s.row ||
          (!s.broken && row > s.row && row < s.row + s.numRows)),
    );
  }

  /** @internal Reads one cell; `row` and `col` are 1-based and in bounds. */
  readCell(row: number, col: number): Cell {
    return this.grid[row - 1]?.[col - 1] ?? { value: '', formula: '' };
  }

  /** @internal Writes one cell; `row` and `col` are 1-based and in bounds. */
  writeCell(row: number, col: number, cell: Cell | undefined): void {
    const target = (this.grid[row - 1] ??= []);
    target[col - 1] = cell;
  }
}

/** Whether a cell is non-empty; `0` and `false` count as content. */
function hasContent(cell: Cell | undefined): boolean {
  return !!cell && (cell.formula !== '' || cell.value !== '');
}

/** Range handle returned by `FakeSheet#getRange`. */
export class FakeRange {
  constructor(
    private readonly sheet: FakeSheet,
    private readonly coords: RangeCall,
  ) {}

  getValues(): unknown[][] {
    return this.map((cell, row, col) =>
      cell.formula && !this.sheet.showsStoredValue(row, col)
        ? this.sheet.evaluate(cell.formula, this.sheet)
        : cell.value,
    );
  }

  getFormulas(): string[][] {
    return this.map((cell) => cell.formula);
  }

  /**
   * Writes `values`, whose dimensions must match the range exactly, with
   * the string interpretation described on `coerceUserEntered`.
   */
  setValues(values: unknown[][]): this {
    const { row, col, numRows, numCols } = this.coords;

    if (values.length !== numRows) {
      throw new Error(
        `The number of rows in the data does not match the number of rows in the range. The data has ${values.length} but the range has ${numRows}.`,
      );
    }
    for (const rowValues of values) {
      if (rowValues.length !== numCols) {
        throw new Error(
          `The number of columns in the data does not match the number of columns in the range. The data has ${rowValues.length} but the range has ${numCols}.`,
        );
      }
    }

    for (const rowValues of values) {
      if (
        rowValues.some(
          (v) => typeof v === 'string' && v.length > MAX_INPUT_LENGTH,
        )
      ) {
        throw new Error(CELL_LIMIT_MESSAGE);
      }
    }

    values.forEach((rowValues, i) =>
      rowValues.forEach((value, j) => {
        this.sheet.breakSpillAt(row + i, col + j);
        this.sheet.writeCell(row + i, col + j, coerceUserEntered(value));
      }),
    );
    this.sheet.calls.setValues.push(this.coords);
    return this;
  }

  clearContent(): this {
    const { row, col, numRows, numCols } = this.coords;
    for (let i = 0; i < numRows; i++) {
      for (let j = 0; j < numCols; j++) {
        this.sheet.writeCell(row + i, col + j, undefined);
      }
    }
    this.sheet.calls.clearContent.push(this.coords);
    return this;
  }

  setFontWeight(weight: string): this {
    void weight;
    this.sheet.calls.setFontWeight.push(this.coords);
    return this;
  }

  applyRowBanding(): { getRange: () => FakeRange; setRange: () => void } {
    this.sheet.calls.applyRowBanding.push(this.coords);
    return { getRange: () => this, setRange: () => undefined };
  }

  private map<T>(read: (cell: Cell, row: number, col: number) => T): T[][] {
    const { row, col, numRows, numCols } = this.coords;
    return Array.from({ length: numRows }, (_, i) =>
      Array.from({ length: numCols }, (_, j) =>
        read(this.sheet.readCell(row + i, col + j), row + i, col + j),
      ),
    );
  }
}

/** In-memory stand-in for an Apps Script `Spreadsheet`. */
export class FakeSpreadsheet {
  private readonly sheets: FakeSheet[] = [];

  constructor(private readonly sheetOptions: FakeSheetOptions = {}) {}

  getSheets = (): FakeSheet[] => [...this.sheets];

  getSheetByName = vi.fn(
    (name: string): FakeSheet | null =>
      this.sheets.find((s) => s.getSheetName() === name) ?? null,
  );

  insertSheet = vi.fn((name: string): FakeSheet => {
    if (this.sheets.some((s) => s.getSheetName() === name)) {
      throw new Error(
        `A sheet with the name "${name}" already exists. Please enter another name.`,
      );
    }
    const sheet = new FakeSheet(name, this.sheetOptions);
    this.sheets.push(sheet);
    return sheet;
  });
}

/** Lock handle returned by the fake, with spies on every method. */
export interface MockLock {
  waitLock: Mock<(timeoutMs: number) => void>;
  tryLock: Mock<(timeoutMs: number) => boolean>;
  releaseLock: Mock<() => void>;
  hasLock: Mock<() => boolean>;
}

type Scope = 'script' | 'document' | 'user';

/** Test switches for the Apps Script service fakes. */
export interface MockOptions {
  /** Scopes for which `LockService` returns `null`, e.g. `['document']`. */
  unavailable?: Scope[];
  /** Evaluator for formula cells of every sheet; see `FormulaEvaluator`. */
  evaluate?: FormulaEvaluator;
}

/**
 * Fresh set of Apps Script service fakes injected as globals.
 *
 * `LockService` models Apps Script's ownership rules: locks are owned per
 * execution and per scope, so every handle for a scope shares one "held"
 * flag. Acquiring a scope this execution already holds succeeds at once, and
 * `releaseLock` on any handle clears the flag.
 *
 * `SpreadsheetApp.flush` records whether any lock was held at the moment it
 * ran, in `flushLockStates`. `Utilities.getUuid` returns a deterministic,
 * incrementing UUID-shaped string. `openById` and `openByUrl` return one
 * stable spreadsheet per argument.
 */
export function createMocks(options: MockOptions = {}) {
  const unavailable = new Set(options.unavailable);
  const heldScopes = new Set<Scope>();
  const locks: MockLock[] = [];
  const flushLockStates: boolean[] = [];
  const sheetOptions: FakeSheetOptions =
    options.evaluate === undefined ? {} : { evaluate: options.evaluate };

  const activeSpreadsheet = new FakeSpreadsheet(sheetOptions);
  const opened = new Map<string, FakeSpreadsheet>();
  const open = (key: string): FakeSpreadsheet => {
    let spreadsheet = opened.get(key);
    if (!spreadsheet) {
      spreadsheet = new FakeSpreadsheet(sheetOptions);
      opened.set(key, spreadsheet);
    }
    return spreadsheet;
  };

  const createLock = (scope: Scope): MockLock => {
    let acquired = false;
    const acquire = (): boolean => {
      heldScopes.add(scope);
      acquired = true;
      return true;
    };

    const lock: MockLock = {
      waitLock: vi.fn(() => void acquire()),
      tryLock: vi.fn(() => acquire()),
      releaseLock: vi.fn(() => {
        heldScopes.delete(scope);
      }),
      hasLock: vi.fn(() => acquired && heldScopes.has(scope)),
    };
    locks.push(lock);
    return lock;
  };

  const factory = (scope: Scope) =>
    vi.fn((): MockLock | null =>
      unavailable.has(scope) ? null : createLock(scope),
    );

  let uuidCounter = 0;

  return {
    activeSpreadsheet,
    /** Every lock handle handed out, in request order. */
    locks,
    flushLockStates,
    /** Whether this execution currently holds a lock of any scope. */
    isLocked: (): boolean => heldScopes.size > 0,
    SpreadsheetApp: {
      getActiveSpreadsheet: vi.fn(
        (): FakeSpreadsheet | null => activeSpreadsheet,
      ),
      openById: vi.fn((id: string): FakeSpreadsheet => open(`id:${id}`)),
      openByUrl: vi.fn((url: string): FakeSpreadsheet => open(`url:${url}`)),
      flush: vi.fn((): void => {
        flushLockStates.push(heldScopes.size > 0);
      }),
    },
    LockService: {
      getScriptLock: factory('script'),
      getDocumentLock: factory('document'),
      getUserLock: factory('user'),
    },
    Utilities: {
      getUuid: vi.fn(
        (): string =>
          `00000000-0000-4000-8000-${String(++uuidCounter).padStart(12, '0')}`,
      ),
    },
    console: {
      info: vi.fn(),
      warn: vi.fn(),
      error: vi.fn(),
      log: vi.fn(),
    },
  };
}

export type Mocks = ReturnType<typeof createMocks>;
