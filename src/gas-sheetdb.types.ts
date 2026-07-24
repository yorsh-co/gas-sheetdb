/** Row/column position config for a table's header and data rows. */
interface GasSheetDbRowsReference {
  columnKeys: number;
  firstData: number;
}

/** A stored entry: persisted system metadata.  */
interface GasSheetDbSystemFields {
  _id?: string;
  _createdAt?: Date;
  _updatedAt?: Date;
  _isDeleted?: boolean;
}

/** A stored entry: arbitrary user properties plus persisted system metadata. */
type GasSheetDbEntry = GasSheetDbSystemFields & Record<string, unknown>;

/** A single raw sheet cell value, as returned by Range#getValues(). */
type GasSheetDbCellValue = string | number | boolean | Date;

/** Constructor options for `GasSheetDb`. */
interface GasSheetDbOptions {
  spreadsheet?: GoogleAppsScript.Spreadsheet.Spreadsheet;
  spreadsheetUrl?: string;
  spreadsheetId?: string;
  useActiveSpreadsheet?: boolean;
  rowNumbers?: Partial<GasSheetDbRowsReference>;
}

/** Options for `GasSheetDb.table()`. */
interface GasSheetDbTableOptions {
  sheetName?: string | null;
  rowNumbers?: Partial<GasSheetDbRowsReference>;
}

/** Constructor options for the internal `_GasSheetDbTable` class. */
interface GasSheetDbTableConstructorOptions {
  spreadsheet: GoogleAppsScript.Spreadsheet.Spreadsheet;
  sheetName?: string | null;
  rowNumbers: GasSheetDbRowsReference;
}

/** Options for `_GasSheetDbTable#find()`. */
interface GasSheetDbFindOptions {
  withTrashed?: boolean;
  onlyTrashed?: boolean;
}
