/**
 * Deep equality for JSON-ish data structures.
 *
 * This is intentionally small and browser-safe (no `node:*` imports) while still
 * being defensive against unexpected inputs like cycles.
 *
 * Supported inputs:
 * - primitives (including `undefined`, `null`, `NaN`)
 * - arrays (including sparse arrays; holes are not treated as `undefined`)
 * - plain objects (including null-prototype objects)
 *
 * A handful of common non-JSON built-ins are handled in a value-ish way to avoid
 * surprising diffs when snapshots contain them via `structuredClone`:
 * - boxed primitives (`new Number(…)`, `new String(…)`, `new Boolean(…)`)
 * - `Date`
 * - `ArrayBuffer` and views (TypedArrays, DataView)
 *
 * Other object types are compared by reference only.
 *
 * @param {any} a
 * @param {any} b
 * @returns {boolean}
 */
export function deepEqual(a, b) {
  /** @type {WeakMap<object, object>} */
  const aToB = new WeakMap();
  /** @type {WeakMap<object, object>} */
  const bToA = new WeakMap();
  return innerDeepEqual(a, b, aToB, bToA);
}

/**
 * @param {any} a
 * @param {any} b
 * @param {WeakMap<object, object>} aToB
 * @param {WeakMap<object, object>} bToA
 * @returns {boolean}
 */
function innerDeepEqual(a, b, aToB, bToA) {
  if (Object.is(a, b)) return true;

  if (a === null || b === null) return false;
  if (a === undefined || b === undefined) return false;

  const typeA = typeof a;
  const typeB = typeof b;
  if (typeA !== typeB) return false;

  // Primitives (and functions) that weren't `Object.is` equal are not equal.
  if (typeA !== "object") return false;

  // Objects: handle cycles / repeated references.
  if (aToB.has(a)) return aToB.get(a) === b;
  if (bToA.has(b)) return bToA.get(b) === a;

  const tagA = Object.prototype.toString.call(a);
  const tagB = Object.prototype.toString.call(b);
  if (tagA !== tagB) return false;

  // Built-in types (value-ish comparisons).
  switch (tagA) {
    case "[object Date]":
      return Object.is(a.getTime(), b.getTime());
    case "[object RegExp]":
      return a.source === b.source && a.flags === b.flags;
    case "[object Number]":
    case "[object String]":
    case "[object Boolean]":
    case "[object BigInt]":
      // Boxed primitives.
      return Object.is(a.valueOf(), b.valueOf());
    case "[object ArrayBuffer]":
      return arrayBufferEqual(a, b);
    default:
      break;
  }

  // TypedArrays / DataView.
  if (ArrayBuffer.isView(a)) {
    if (!ArrayBuffer.isView(b)) return false;
    // Ensure we don't treat different view types (e.g. Uint8Array vs Int8Array) as equal.
    if (a.constructor !== b.constructor) return false;
    return arrayBufferViewEqual(a, b);
  }

  // Only recurse into arrays + plain objects. Everything else is by-reference.
  const protoA = Object.getPrototypeOf(a);
  const protoB = Object.getPrototypeOf(b);
  if (protoA !== protoB) return false;

  // Record mapping before recursing so cycles short-circuit correctly.
  aToB.set(a, b);
  bToA.set(b, a);

  if (Array.isArray(a)) return arrayEqual(a, b, aToB, bToA);

  if (protoA !== Object.prototype && protoA !== null) {
    // Unknown object type (class instances, Maps, Sets, Errors, Yjs types, …).
    // We intentionally avoid walking these (they can be cyclic / huge) and fall
    // back to reference equality (already handled by Object.is above).
    return false;
  }

  return plainObjectEqual(a, b, aToB, bToA);
}

/**
 * @param {ArrayBuffer} a
 * @param {ArrayBuffer} b
 */
function arrayBufferEqual(a, b) {
  if (a.byteLength !== b.byteLength) return false;
  const va = new Uint8Array(a);
  const vb = new Uint8Array(b);
  for (let i = 0; i < va.length; i++) {
    if (va[i] !== vb[i]) return false;
  }
  return true;
}

/**
 * @param {ArrayBufferView} a
 * @param {ArrayBufferView} b
 */
function arrayBufferViewEqual(a, b) {
  if (a.byteLength !== b.byteLength) return false;
  const va = new Uint8Array(a.buffer, a.byteOffset, a.byteLength);
  const vb = new Uint8Array(b.buffer, b.byteOffset, b.byteLength);
  for (let i = 0; i < va.length; i++) {
    if (va[i] !== vb[i]) return false;
  }
  return true;
}

/**
 * @param {any[]} a
 * @param {any[]} b
 * @param {WeakMap<object, object>} aToB
 * @param {WeakMap<object, object>} bToA
 */
function arrayEqual(a, b, aToB, bToA) {
  if (!Array.isArray(b)) return false;
  if (a.length !== b.length) return false;

  // Compare index entries, treating holes distinctly from `undefined`.
  for (let i = 0; i < a.length; i++) {
    const aHas = Object.prototype.hasOwnProperty.call(a, i);
    const bHas = Object.prototype.hasOwnProperty.call(b, i);
    if (aHas !== bHas) return false;
    if (aHas && !innerDeepEqual(a[i], b[i], aToB, bToA)) return false;
  }

  // Compare any additional enumerable properties (rare, but keeps behavior closer
  // to Node's deep equality for arrays).
  const aKeys = Object.keys(a).filter((k) => !isArrayIndexKey(k));
  const bKeys = Object.keys(b).filter((k) => !isArrayIndexKey(k));
  if (aKeys.length !== bKeys.length) return false;
  aKeys.sort();
  bKeys.sort();
  for (let i = 0; i < aKeys.length; i++) {
    if (aKeys[i] !== bKeys[i]) return false;
  }
  for (const key of aKeys) {
    if (!innerDeepEqual(a[key], b[key], aToB, bToA)) return false;
  }
  return true;
}

/**
 * @param {string} key
 * @returns {boolean}
 */
function isArrayIndexKey(key) {
  // Avoid treating "01" as an index.
  if (key === "") return false;
  const n = Number(key);
  if (!Number.isInteger(n) || n < 0) return false;
  return String(n) === key;
}

/**
 * @param {Record<string, any>} a
 * @param {Record<string, any>} b
 * @param {WeakMap<object, object>} aToB
 * @param {WeakMap<object, object>} bToA
 */
function plainObjectEqual(a, b, aToB, bToA) {
  /** @type {string[]} */
  const aKeys = Object.keys(a);
  /** @type {string[]} */
  const bKeys = Object.keys(b);
  if (aKeys.length !== bKeys.length) return false;
  aKeys.sort();
  bKeys.sort();
  for (let i = 0; i < aKeys.length; i++) {
    if (aKeys[i] !== bKeys[i]) return false;
  }
  for (const key of aKeys) {
    if (!innerDeepEqual(a[key], b[key], aToB, bToA)) return false;
  }
  return true;
}

