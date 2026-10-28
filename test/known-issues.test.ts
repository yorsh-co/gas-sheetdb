import { describe, it } from 'vitest';

/**
 * Defects scheduled for 2.0.0. They are tracked as todos rather than as
 * assertions of the current behaviour, which must not be locked in.
 */
describe('known issues (2.0.0)', () => {
  it.todo(
    'C4: stores strings verbatim instead of letting Sheets coerce them (numeric-looking text, formula injection)',
  );
  it.todo('M4: a user string starting with "__JSON__" is not decoded as JSON');
  it.todo('M1: blank rows are not returned as entries');
  it.todo('H4: rewriting the table body preserves formulas in untouched cells');
  it.todo('H3: read methods do not take the lock');
  it.todo('C2: duplicate _id values are repaired');
  it.todo('M3: numeric header keys are read back as strings');
  it.todo('M5: insertMany([]) returns an empty array instead of throwing');
  it.todo(
    'M6: a non-boolean _isDeleted is not silently overwritten with false',
  );
});
