/**
 * V8 engine extension for capturing a stack trace without the constructor
 * frame. Not part of the ECMAScript spec, so absent from lib.es2021 and
 * @types/google-apps-script — declared here since Apps Script runs on V8.
 * Type-only: it emits nothing, so it does not reach dist/.
 */
declare interface ErrorConstructor {
  captureStackTrace(targetObject: object, constructorOpt?: object): void;
}
