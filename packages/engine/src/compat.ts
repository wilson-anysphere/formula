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

