/**
 * Encodes complex values for storage in sheets.
 *
 * Objects are serialized with a reserved prefix so they
 * can be restored when reading rows. A string that itself starts with the
 * prefix is serialized the same way, so it reads back unchanged.
 */
const _GasSheetDbValueCodec = (() => {
  const PREFIXES = Object.freeze({
    JSON: '__JSON__',
  });

  /** Maximum number of characters Sheets stores in a cell. */
  const MAX_CELL_LENGTH = 50000;

  /**
   * Convert a JS value into the logical cell value: what a cell holds once
   * stored, and what reading it back returns. Strings are not escaped here;
   * see `toSheetInput`.
   */
  const encode = (value: unknown): GasSheetDbCellValue => {
    if (value === undefined || value === null) {
      return '';
    }

    if (value instanceof Date) {
      return value;
    }

    if (typeof value === 'object') {
      return `${PREFIXES.JSON}${JSON.stringify(value)}`;
    }

    if (typeof value === 'string' && value.startsWith(PREFIXES.JSON)) {
      return `${PREFIXES.JSON}${JSON.stringify(value)}`;
    }

    return value as GasSheetDbCellValue;
  };

  /**
   * Convert a logical cell value into the value to pass to `setValues`.
   *
   * Sheets parses strings as if typed (formulas, numbers, booleans, dates).
   * A leading apostrophe makes it store the rest as text, and reading the
   * cell returns it without the apostrophe. Other values are written as-is.
   * An empty string is left alone: it already stores a blank cell.
   */
  const toSheetInput = (value: GasSheetDbCellValue): GasSheetDbCellValue =>
    typeof value === 'string' && value !== '' ? `'${value}` : value;

  /**
   * Describe why a value cannot be stored, or return `undefined` when it can:
   * a `bigint`, `symbol` or `function`, a value that does not serialize to
   * JSON, or a cell longer than the Sheets limit once encoded.
   */
  const unsupportedReason = (value: unknown): string | undefined => {
    if (
      typeof value === 'bigint' ||
      typeof value === 'symbol' ||
      typeof value === 'function'
    ) {
      return `unsupported ${typeof value} value`;
    }

    let encoded: GasSheetDbCellValue;

    try {
      encoded = encode(value);
    } catch (err) {
      return `value cannot be serialized to JSON: ${(err as Error).message}`;
    }

    if (typeof encoded === 'string' && encoded.length > MAX_CELL_LENGTH) {
      return `value exceeds the ${MAX_CELL_LENGTH} character cell limit (${encoded.length})`;
    }

    return undefined;
  };

  /**
   * Restore encoded sheet values back into JS values.
   *
   * A cell whose payload no longer parses falls back to its raw string.
   * `onError` reports that to the caller rather than logging here: decoding
   * runs once per cell, so a single corrupt column would otherwise emit one
   * log line per row. Callers aggregate and log once per operation.
   */
  const decode = (
    value: GasSheetDbCellValue,
    onError?: (err: unknown) => void,
  ): unknown => {
    if (typeof value !== 'string') {
      return value;
    }

    if (value.startsWith(PREFIXES.JSON)) {
      try {
        return JSON.parse(value.slice(PREFIXES.JSON.length));
      } catch (err) {
        onError?.(err);
      }
    }

    return value;
  };

  return Object.freeze({
    encode,
    toSheetInput,
    unsupportedReason,
    decode,
  });
})();
