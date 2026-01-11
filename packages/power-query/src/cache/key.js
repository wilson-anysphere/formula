/**
 * Deterministic serialization + hashing utilities used for cache keys.
 *
 * We intentionally avoid Node-only crypto dependencies so this can run in the
 * browser, workers, and Node.
 */

/**
 * @param {unknown} value
 * @param {Map<object, string>} seen
 * @param {string} path
 * @returns {unknown}
 */
function canonicalize(value, seen, path) {
  if (value === null) return null;
  const type = typeof value;
  if (type === "string" || type === "boolean") return value;
  if (type === "number") {
    if (Number.isFinite(value)) return value;
    return { $type: "number", value: String(value) };
  }
  if (type === "bigint") return { $type: "bigint", value: value.toString() };
  if (type === "undefined") return { $type: "undefined" };
  if (type === "symbol") return { $type: "symbol", value: String(value) };
  if (type === "function") return { $type: "function", value: value.name || "<anonymous>" };

  if (value instanceof Date) {
    return { $type: "date", value: value.toISOString() };
  }

  if (Array.isArray(value)) {
    return value.map((item, idx) => canonicalize(item, seen, `${path}[${idx}]`));
  }

  if (value instanceof Map) {
    const entries = Array.from(value.entries()).map(([k, v]) => [
      canonicalize(k, seen, `${path}.<mapKey>`),
      canonicalize(v, seen, `${path}.<mapValue>`),
    ]);
    entries.sort((a, b) => stableStringify(a[0]).localeCompare(stableStringify(b[0])));
    return { $type: "map", entries };
  }

  if (value instanceof Set) {
    const entries = Array.from(value.values()).map((v, idx) => canonicalize(v, seen, `${path}.<set>${idx}`));
    entries.sort((a, b) => stableStringify(a).localeCompare(stableStringify(b)));
    return { $type: "set", entries };
  }

  if (type === "object") {
    const obj = /** @type {object} */ (value);
    const existing = seen.get(obj);
    if (existing) {
      return { $type: "circular", ref: existing };
    }
    seen.set(obj, path);

    const out = {};
    const keys = Object.keys(obj).sort();
    for (const key of keys) {
      // @ts-ignore - runtime indexing
      out[key] = canonicalize(obj[key], seen, `${path}.${key}`);
    }
    return out;
  }

  return { $type: "unknown", value: String(value) };
}

/**
 * Deterministically stringify a JS value: stable object key ordering + handling
 * of Dates/Maps/Sets/undefined.
 *
 * @param {unknown} value
 * @returns {string}
 */
export function stableStringify(value) {
  const canonical = canonicalize(value, new Map(), "$");
  return JSON.stringify(canonical);
}

/**
 * FNV-1a 64-bit hash (as hex) for deterministic cache keys.
 *
 * @param {string} input
 * @returns {string}
 */
export function fnv1a64(input) {
  let hash = 0xcbf29ce484222325n;
  const prime = 0x100000001b3n;

  for (let i = 0; i < input.length; i++) {
    hash ^= BigInt(input.charCodeAt(i));
    hash = (hash * prime) & 0xffffffffffffffffn;
  }

  return hash.toString(16).padStart(16, "0");
}

/**
 * Hash any value using `stableStringify` + `fnv1a64`.
 *
 * @param {unknown} value
 * @returns {string}
 */
export function hashValue(value) {
  return fnv1a64(stableStringify(value));
}

