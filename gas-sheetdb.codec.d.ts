/**
 * Encodes complex values for storage in sheets.
 *
 * Objects are serialized with a reserved prefix so they
 * can be restored when reading rows. A string that itself starts with the
 * prefix is serialized the same way, so it reads back unchanged.
 */
declare const _GasSheetDbValueCodec: Readonly<{
  encode: (value: unknown) => GasSheetDbCellValue;
  toSheetInput: (value: GasSheetDbCellValue) => GasSheetDbCellValue;
  unsupportedReason: (value: unknown) => string | undefined;
  decode: (
    value: GasSheetDbCellValue,
    onError?: (err: unknown) => void,
  ) => unknown;
}>;
