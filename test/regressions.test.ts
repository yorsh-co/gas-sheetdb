import { describe, expect, it } from 'vitest';
import { at, rows, setupTable } from './helpers';

describe('1.3.2 regressions', () => {
  describe('deleting every remaining entry', () => {
    const fill = () => {
      const context = setupTable();
      const entries = context.table.insertMany([
        { name: 'a' },
        { name: 'b' },
        { name: 'c' },
      ]);
      return { ...context, entries };
    };

    it.each([
      [
        'deleteWhere',
        (c: ReturnType<typeof fill>) => c.table.deleteWhere(() => true),
      ],
      [
        'deleteMany',
        (c: ReturnType<typeof fill>) => c.table.deleteMany(c.entries),
      ],
    ])('%s empties the body and keeps a non-frozen row', (_name, run) => {
      const context = fill();

      expect(() => run(context)).not.toThrow();

      expect(context.table.find({ withTrashed: true })).toEqual([]);
      expect(rows(context.sheet)).toEqual([]);
      expect(context.sheet.getLastRow()).toBe(1);
      expect(context.sheet.getMaxRows()).toBeGreaterThan(
        context.sheet.getFrozenRows(),
      );
    });

    it('deleteOne empties a single-entry table', () => {
      const { table, sheet } = setupTable();
      const only = table.insert({ name: 'a' });

      expect(() => table.deleteOne(only)).not.toThrow();

      expect(table.find()).toEqual([]);
      expect(sheet.getLastRow()).toBe(1);
    });

    it('accepts new entries after the body was emptied', () => {
      const { table, entries } = fill();
      table.deleteMany(entries);

      table.insert({ name: 'again' });

      expect(table.find().map((e) => e.name)).toEqual(['again']);
    });
  });

  describe('deleteMany with the same entry twice', () => {
    it('leaves [b, c] for deleteMany([a, a]) on [a, b, c]', () => {
      const { table } = setupTable();
      const [a] = table.insertMany([
        { name: 'a' },
        { name: 'b' },
        { name: 'c' },
      ]);

      table.deleteMany([a!, a!]);

      expect(table.find().map((e) => e.name)).toEqual(['b', 'c']);
    });

    it('deletes once when entries resolve to the same row', () => {
      const { table, entries } = (() => {
        const context = setupTable();
        const inserted = context.table.insertMany([
          { name: 'a' },
          { name: 'b' },
          { name: 'c' },
        ]);
        return { ...context, entries: inserted };
      })();

      table.deleteMany([at(entries, 1), { _id: at(entries, 1)._id! }]);

      expect(table.find().map((e) => e.name)).toEqual(['a', 'c']);
    });
  });
});
