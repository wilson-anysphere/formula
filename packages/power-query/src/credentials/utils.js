/**
 * Browser-safe random ID generator.
 *
 * Credential IDs do not need to be cryptographically unpredictable; they need to
 * be collision-resistant and stable so cache keys can safely vary by credential
 * without embedding secret material.
 *
 * @param {number} [bytes]
 */
export function randomId(bytes = 16) {
  // Prefer the Web Crypto API when available (browser + Node 18+).
  if (globalThis.crypto && typeof globalThis.crypto.getRandomValues === "function") {
    const buf = new Uint8Array(bytes);
    globalThis.crypto.getRandomValues(buf);
    return Array.from(buf, (b) => b.toString(16).padStart(2, "0")).join("");
  }

  if (globalThis.crypto && typeof globalThis.crypto.randomUUID === "function") {
    return globalThis.crypto.randomUUID();
  }

  // Last-resort fallback (still fine for cache partitioning in tests).
  return `${Date.now().toString(16)}-${Math.random().toString(16).slice(2)}`;
}

