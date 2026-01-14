/**
 * Best-effort detection for environments where the additive `getRangeCompact` API is missing.
 *
 * This can happen when:
 * - the caller is using an older worker bundle (missing RPC method)
 * - the worker is new enough but the loaded WASM build doesn't export `getRangeCompact`
 *
 * Callers should treat this as a capability probe: if it returns `true`, fall back to `getRange`
 * and cache the result to avoid repeatedly paying exception overhead.
 */
export function isMissingGetRangeCompactError(err: unknown): boolean {
  const message = err instanceof Error ? err.message : String(err);
  if (message.includes("unknown method: getRangeCompact")) {
    return true;
  }
  const lower = message.toLowerCase();
  return lower.includes("getrangecompact") && lower.includes("not available");
}

/**
 * Best-effort detection for environments where the additive `supportedLocaleIds` API is missing.
 *
 * This can happen when:
 * - the caller is using an older worker bundle (missing RPC method)
 * - the worker is new enough but the loaded WASM build doesn't export `supportedLocaleIds`
 */
export function isMissingSupportedLocaleIdsError(err: unknown): boolean {
  const message = err instanceof Error ? err.message : String(err);
  if (message.includes("unknown method: supportedLocaleIds")) {
    return true;
  }
  const lower = message.toLowerCase();
  return lower.includes("supportedlocaleids") && lower.includes("does not export");
}

/**
 * Best-effort detection for environments where the additive `getLocaleInfo` API is missing.
 *
 * This can happen when:
 * - the caller is using an older worker bundle (missing RPC method)
 * - the worker is new enough but the loaded WASM build doesn't export `getLocaleInfo`
 */
export function isMissingGetLocaleInfoError(err: unknown): boolean {
  const message = err instanceof Error ? err.message : String(err);
  if (message.includes("unknown method: getLocaleInfo")) {
    return true;
  }
  const lower = message.toLowerCase();
  return lower.includes("getlocaleinfo") && lower.includes("does not export");
}

/**
 * Best-effort detection for environments where the additive `getCellPhonetic` API is missing.
 *
 * This can happen when:
 * - the caller is using an older worker bundle (missing RPC method)
 * - the worker is new enough but the loaded WASM build doesn't export `getCellPhonetic`
 */
export function isMissingGetCellPhoneticError(err: unknown): boolean {
  const message = err instanceof Error ? err.message : String(err);
  if (message.includes("unknown method: getCellPhonetic")) {
    return true;
  }
  const lower = message.toLowerCase();
  return lower.includes("getcellphonetic") && lower.includes("not available");
}

/**
 * Best-effort detection for environments where the additive `setCellPhonetic` API is missing.
 *
 * This can happen when:
 * - the caller is using an older worker bundle (missing RPC method)
 * - the worker is new enough but the loaded WASM build doesn't export `setCellPhonetic`
 */
export function isMissingSetCellPhoneticError(err: unknown): boolean {
  const message = err instanceof Error ? err.message : String(err);
  if (message.includes("unknown method: setCellPhonetic")) {
    return true;
  }
  const lower = message.toLowerCase();
  return lower.includes("setcellphonetic") && lower.includes("not available");
}
