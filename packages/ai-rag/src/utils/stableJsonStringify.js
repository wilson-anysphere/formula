/**
 * Deterministic JSON stringification with stable object key ordering.
 *
 * This is intended for hashing/cache keys, not for human readability.
 *
 * - Object keys are sorted lexicographically at every level.
 * - Arrays preserve order.
 * - `toJSON` is respected (since JSON.stringify invokes it before the replacer).
 *
 * @param {any} value
 * @returns {string}
 */
export function stableJsonStringify(value) {
  return JSON.stringify(value, (_key, val) => {
    if (typeof val === "bigint") return val.toString();
    if (!val || typeof val !== "object") return val;
    if (Array.isArray(val)) return val;
    const keys = Object.keys(val).sort();
    /** @type {Record<string, any>} */
    const out = {};
    for (const k of keys) out[k] = val[k];
    return out;
  });
}

