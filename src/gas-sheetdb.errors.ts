/**
 * Expected data failure thrown when gas-error is not loaded. Carries the same
 * fields as gas-error's base error, so a consumer can map either in one line.
 * Configuration and programmer errors are plain `Error`s instead.
 */
class GasSheetDbError extends Error {
  statusCode: number;
  code: GasSheetDbErrorCode;
  details: unknown;
  isOperational: true;

  constructor(
    message: string,
    statusCode: number,
    code: GasSheetDbErrorCode,
    details?: unknown,
  ) {
    super(message);

    this.name = this.constructor.name;
    this.statusCode = statusCode;
    this.code = code;
    this.details = details;
    this.isOperational = true;

    Error.captureStackTrace(this, this.constructor);
  }
}

/**
 * Builds the error to throw for an expected data failure: the matching
 * gas-error class when it is loaded, else a `GasSheetDbError`. Availability is
 * checked at throw time, so the Apps Script file push order does not matter.
 */
const _GasSheetDbErrors = Object.freeze({
  /** 400 `VALIDATION_ERROR`: the caller passed unusable input. */
  validation(message: string, details?: unknown): Error {
    return typeof GasValidationError === 'function'
      ? new GasValidationError(message, details)
      : new GasSheetDbError(message, 400, 'VALIDATION_ERROR', details);
  },

  /** 404 `NOT_FOUND`: the addressed entry does not exist. */
  notFound(message: string): Error {
    return typeof GasNotFoundError === 'function'
      ? new GasNotFoundError(message)
      : new GasSheetDbError(message, 404, 'NOT_FOUND');
  },
});
