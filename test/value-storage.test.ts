import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import { at, header, rows, setupTable } from './helpers';

beforeEach(() => {
  vi.useFakeTimers();
  vi.setSystemTime(new Date('2026-01-01T00:00:00Z'));
});
afterEach(() => vi.useRealTimers());

const TEXTS = [
  '=1+1',
  '=IMPORTXML("http://x","//a")',
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
  "''quoted",
  '__JSON__{"isAdmin":true}',
  '__JSON__x',
  '',
  'plain text',
];

/** The column index (1-based) of `key` in the header row. */
const columnOf = (sheet: Parameters<typeof header>[0], key: string) =>
  header(sheet).indexOf(key) + 1;

describe('text storage (C4, M4)', () => {
  it.each(TEXTS)('round-trips %j through insert and find', (text) => {
    const { table, sheet } = setupTable();

    const inserted = table.insert({ note: text });

    expect(inserted.note).toBe(text);
    expect(table.find()[0]!.note).toBe(text);
    expect(rows(sheet)[0]!.note).toBe(
      text.startsWith('__JSON__') ? `__JSON__${JSON.stringify(text)}` : text,
    );
    expect(
      sheet.getRange(2, columnOf(sheet, 'note')).getFormulas()[0]![0],
    ).toBe('');
  });

  it.each(TEXTS)('round-trips %j through update and find', (text) => {
    const { table } = setupTable();
    const _id = table.insert({ note: 'before', other: 'x' })._id!;

    const updated = table.update({ _id, note: text });

    expect(updated.note).toBe(text);
    expect(table.find()[0]!.note).toBe(text);
  });

  it('round-trips strings inside a backfilled id and header keys', () => {
    const { table, sheet } = setupTable();

    table.insert({ 2024: 'a', TRUE: 'b' });

    expect(header(sheet)).toEqual(expect.arrayContaining(['2024', 'TRUE']));
    expect(table.find()[0]).toMatchObject({ 2024: 'a', TRUE: 'b' });
    expect(table.find()).toHaveLength(1);
  });

  it('stores numbers, booleans and dates natively', () => {
    const { table, sheet } = setupTable();
    const date = new Date('2025-05-05T10:00:00Z');

    table.insert({ n: 12, b: true, d: date });

    const [row] = rows(sheet);
    expect(row!.n).toBe(12);
    expect(row!.b).toBe(true);
    expect(row!.d).toBeInstanceOf(Date);
    expect(table.find()[0]).toMatchObject({ n: 12, b: true, d: date });
  });

  it('decodes a string that starts with the JSON prefix back to the string', () => {
    const { table } = setupTable();

    table.insert({ note: '__JSON__{"isAdmin":true}' });

    expect(table.find()[0]!.note).toBe('__JSON__{"isAdmin":true}');
  });

  it('still decodes objects written with the prefix', () => {
    const { table, sheet } = setupTable();
    table.insert({ settings: { a: 1 } });

    expect(rows(sheet)[0]!.settings).toBe('__JSON__{"a":1}');
    expect(table.find()[0]!.settings).toEqual({ a: 1 });
  });

  it('writes only the changed cells of a row with many string cells', () => {
    const { table, sheet } = setupTable();
    const _id = table.insert({ a: 'a', b: 'b', c: 'c', d: 'd', e: 'e' })._id!;
    vi.advanceTimersByTime(1000);
    sheet.calls.setValues.length = 0;

    table.update({ _id, c: 'changed' });

    const written = sheet.calls.setValues.flatMap((call) =>
      Array.from({ length: call.numCols }, (_, i) =>
        header(sheet).at(call.col - 1 + i),
      ),
    );
    expect(written.sort()).toEqual(['_updatedAt', 'c']);
  });
});

describe('unsupported values (M7)', () => {
  const column = 'note';

  it('stores a text of 49,999 characters, the longest Sheets accepts', () => {
    const { table, sheet } = setupTable();

    table.insert({ [column]: 'x'.repeat(49_999) });

    expect(table.find()[0]![column]).toHaveLength(49_999);
    expect(rows(sheet)[0]![column]).toHaveLength(49_999);
  });

  it('rejects a 50,000 character text on insert, writing and mutating nothing', () => {
    const { table, sheet } = setupTable();
    table.find();
    const before = sheet.calls.setValues.length;
    const good = { name: 'ok' };
    const bad = { [column]: 'x'.repeat(50_000) };

    expect(() => table.insertMany([good, bad])).toThrow(
      expect.objectContaining({
        message: expect.stringMatching(/^\[GasSheetDb\] .*"note"/),
        code: 'VALIDATION_ERROR',
        details: { column, reason: expect.stringContaining('49999') },
      }),
    );
    expect(sheet.calls.setValues.length).toBe(before);
    expect(good).toEqual({ name: 'ok' });
    expect(bad).toEqual({ [column]: 'x'.repeat(50_000) });
  });

  it('counts the text escape towards the limit of an encoded object', () => {
    const { table } = setupTable();
    // `__JSON__` plus the JSON of `{ v }` is 16 characters longer than `v`
    const encoded = (length: number) => ({ v: 'x'.repeat(length - 16) });

    expect(() => table.insert({ [column]: encoded(50_000) })).toThrow(
      expect.objectContaining({ code: 'VALIDATION_ERROR' }),
    );

    table.insert({ [column]: encoded(49_999) });
    expect(table.find()[0]![column]).toEqual(encoded(49_999));
  });

  it('rejects a 50,000 character text on update, writing and mutating nothing', () => {
    const { table, sheet } = setupTable();
    const a = table.insert({ [column]: 'a' });
    const b = table.insert({ [column]: 'b' });
    const before = sheet.calls.setValues.length;
    const first = { ...a, [column]: 'changed' };
    const second = { ...b, [column]: 'x'.repeat(50_000) };

    expect(() => table.updateMany([first, second])).toThrow(
      expect.objectContaining({ code: 'VALIDATION_ERROR' }),
    );
    expect(sheet.calls.setValues.length).toBe(before);
    expect(first[column]).toBe('changed');
    expect(first._updatedAt).toEqual(a._updatedAt);
  });

  it.each([
    ['bigint', 10n],
    ['symbol', Symbol('s')],
    ['function', () => 1],
  ])('rejects a %s', (type, value) => {
    const { table, sheet } = setupTable();
    const stored = table.insert({ name: 'ok' });
    const before = sheet.calls.setValues.length;

    expect(() => table.insert({ [column]: value as never })).toThrow(
      expect.objectContaining({
        message: expect.stringContaining(`"${column}"`),
        details: { column, reason: expect.stringContaining(type) },
      }),
    );
    expect(() =>
      table.update({ _id: stored._id!, [column]: value as never }),
    ).toThrow(/\[GasSheetDb\]/);
    expect(sheet.calls.setValues.length).toBe(before);
    expect(at(table.find(), 0)).not.toHaveProperty(column);
  });

  it('rejects a bigint nested in an object', () => {
    const { table } = setupTable();

    expect(() => table.insert({ [column]: { n: 1n } })).toThrow(
      expect.objectContaining({
        details: { column, reason: expect.any(String) },
      }),
    );
  });
});
