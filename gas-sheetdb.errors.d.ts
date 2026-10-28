/**
 * Expected data failure thrown when gas-error is not loaded. Carries the same
 * fields as gas-error's base error, so a consumer can map either in one line.
 * Configuration and programmer errors are plain `Error`s instead.
 */
declare class GasSheetDbError extends Error {
  statusCode: number;
  code: GasSheetDbErrorCode;
  details: unknown;
  isOperational: true;
  constructor(
    message: string,
    statusCode: number,
    code: GasSheetDbErrorCode,
    details?: unknown,
  );
}
/**
 * Builds the error to throw for an expected data failure: the matching
 * gas-error class when it is loaded, else a `GasSheetDbError`. Availability is
 * checked at throw time, so the Apps Script file push order does not matter.
 */
declare const _GasSheetDbErrors: Readonly<{
  /** 400 `VALIDATION_ERROR`: the caller passed unusable input. */
  validation(message: string, details?: unknown): Error;
  /** 404 `NOT_FOUND`: the addressed entry does not exist. */
  notFound(message: string): Error;
  /** 409 `CONFLICT`: the write would duplicate an existing `_id`. */
  conflict(message: string): Error;
}>;
