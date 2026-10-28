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
});
