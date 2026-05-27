/**
 * Encodes complex values for storage in sheets.
 *
 * Objects are serialized with a reserved prefix so they
 * can be restored when reading rows.
 */
const _SheetDbValueCodec = (() => {
  const PREFIXES = Object.freeze({
    JSON: '__JSON__',
  });

  /**
   * Convert a JS value into a sheet-safe value.
   *
   * @param {*} value
   * @returns {*}
   */
  const encode = (value) => {
    if (value === undefined || value === null) {
      return '';
    }

    if (value instanceof Date) {
      return value;
    }

    if (typeof value === 'object') {
      return `${PREFIXES.JSON}${JSON.stringify(value)}`;
    }

    return value;
  };

  /**
   * Restore encoded sheet values back into JS values.
   *
   * @param {*} value
   * @returns {*}
   */
  const decode = (value) => {
    if (typeof value !== 'string') {
      return value;
    }

    if (value.startsWith(PREFIXES.JSON)) {
      try {
        return JSON.parse(value.slice(PREFIXES.JSON.length));
      } catch (err) {
        console.error(err);
      }
    }

    return value;
  };

  return Object.freeze({
    encode,
    decode,
  });
})();
