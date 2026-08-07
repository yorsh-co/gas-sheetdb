/**
 * Encodes complex values for storage in sheets.
 *
 * Objects are serialized with a reserved prefix so they
 * can be restored when reading rows.
 */
declare const _GasSheetDbValueCodec: Readonly<{
  encode: (value: unknown) => GasSheetDbCellValue;
  decode: (
    value: GasSheetDbCellValue,
    onError?: (err: unknown) => void,
  ) => unknown;
}>;
