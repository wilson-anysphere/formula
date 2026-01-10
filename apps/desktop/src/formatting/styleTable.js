/**
 * @typedef {Record<string, any>} CellStyle
 */

function isPlainObject(value) {
  return Boolean(value) && typeof value === "object" && !Array.isArray(value);
}

function stableStringify(value) {
  if (value === undefined) return "undefined";
  if (value == null || typeof value !== "object") return JSON.stringify(value);
  if (Array.isArray(value)) return `[${value.map(stableStringify).join(",")}]`;
  const keys = Object.keys(value).sort();
  const entries = keys.map((k) => `${JSON.stringify(k)}:${stableStringify(value[k])}`);
  return `{${entries.join(",")}}`;
}

function deepMerge(base, patch) {
  if (!isPlainObject(base) || !isPlainObject(patch)) return patch;
  const out = { ...base };
  for (const [key, value] of Object.entries(patch)) {
    if (value === undefined) continue;
    if (isPlainObject(value) && isPlainObject(out[key])) {
      out[key] = deepMerge(out[key], value);
    } else {
      out[key] = value;
    }
  }
  return out;
}

/**
 * Apply a patch to a style object (deep merge).
 *
 * @param {CellStyle} base
 * @param {CellStyle | null} patch
 * @returns {CellStyle}
 */
export function applyStylePatch(base, patch) {
  if (patch == null) return {};
  if (!isPlainObject(base)) return deepMerge({}, patch);
  return deepMerge(base, patch);
}

/**
 * Deduplicated table of style objects. Styles are immutable once interned.
 */
export class StyleTable {
  constructor() {
    /** @type {CellStyle[]} */
    this.styles = [{}];
    /** @type {Map<string, number>} */
    this.index = new Map([[stableStringify({}), 0]]);
  }

  /**
   * @param {CellStyle | null | undefined} style
   * @returns {number}
   */
  intern(style) {
    const normalized = style == null ? {} : style;
    const key = stableStringify(normalized);
    const existing = this.index.get(key);
    if (existing != null) return existing;
    const id = this.styles.length;
    const structuredCloneFn =
      typeof globalThis.structuredClone === "function" ? globalThis.structuredClone : null;
    this.styles.push(structuredCloneFn ? structuredCloneFn(normalized) : JSON.parse(JSON.stringify(normalized)));
    this.index.set(key, id);
    return id;
  }

  /**
   * @param {number} styleId
   * @returns {CellStyle}
   */
  get(styleId) {
    return this.styles[styleId] ?? {};
  }

  get size() {
    return this.styles.length;
  }
}
