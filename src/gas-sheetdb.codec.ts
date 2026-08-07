/**
 * Encodes complex values for storage in sheets.
 *
 * Objects are serialized with a reserved prefix so they
 * can be restored when reading rows.
 */
const _GasSheetDbValueCodec = (() => {
  const PREFIXES = Object.freeze({
    JSON: '__JSON__',
  });

  /**
   * Convert a JS value into a sheet-safe value.
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

    return value as GasSheetDbCellValue;
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
    decode,
  });
})();
