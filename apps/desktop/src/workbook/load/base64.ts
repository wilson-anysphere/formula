/**
 * Helpers for dealing with base64-encoded binary payloads in workbook import paths.
 *
 * These functions are intentionally small and dependency-free so they can be used in both
 * browser and node/vitest environments.
 */

/**
 * Estimate decoded bytes without decoding.
 *
 * Assumes the input has already been normalized to a raw base64 string without a `data:` prefix.
 */
export function estimateBase64Bytes(base64: string): number {
  const len = base64.length;
  if (len === 0) return 0;
  const padding = base64.endsWith("==") ? 2 : base64.endsWith("=") ? 1 : 0;
  return Math.floor((len * 3) / 4) - padding;
}

export function normalizeBase64String(raw: string): string | null {
  if (typeof raw !== "string") return null;
  const trimmed = raw.trim();
  if (!trimmed) return null;

  // Strip optional data URL prefix.
  let base64 = trimmed;
  if (base64.startsWith("data:")) {
    const comma = base64.indexOf(",");
    if (comma === -1) return null;
    base64 = base64.slice(comma + 1);
  }

  base64 = base64.trim();
  if (!base64) return null;

  // Be tolerant of whitespace/newlines.
  if (/\s/.test(base64)) {
    base64 = base64.replace(/\s+/g, "");
  }

  // Support base64url; normalize to standard base64.
  if (base64.includes("-") || base64.includes("_")) {
    base64 = base64.replace(/-/g, "+").replace(/_/g, "/");
  }

  const mod = base64.length % 4;
  if (mod === 1) return null;
  if (mod === 2) base64 += "==";
  else if (mod === 3) base64 += "=";

  return base64 || null;
}

/**
 * Normalize and enforce a decoded byte-size cap without decoding.
 *
 * Returns the normalized base64 string (standard base64, no whitespace/data prefix) when the
 * payload is within `maxBytes`; otherwise returns null.
 */
export function coerceBase64StringWithinLimit(base64Raw: string, maxBytes: number): string | null {
  if (typeof base64Raw !== "string") return null;
  const max = Number.isFinite(maxBytes) && maxBytes > 0 ? Math.floor(maxBytes) : 0;
  if (max === 0) return null;

  // Fast length guard before we perform any normalization that could allocate large intermediate
  // strings (trim/replace).
  //
  // Base64 expands bytes by ~4/3 plus padding. Allow generous overhead for optional `data:` prefixes
  // and whitespace/newlines.
  const roughMaxChars = Math.ceil(((max + 2) * 4) / 3) + 128;
  if (base64Raw.length > roughMaxChars * 2) return null;

  const base64 = normalizeBase64String(base64Raw);
  if (!base64) return null;

  if (estimateBase64Bytes(base64) > max) return null;
  return base64;
}

