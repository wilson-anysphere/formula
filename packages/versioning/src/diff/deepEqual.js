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
  /** @type {any[]} */
  const aStack = [];
  /** @type {any[]} */
  const bStack = [];
  // Track only the *current recursion stack* (not all visited nodes). This matches
  // Node's deep equality semantics:
  // - repeated (non-cyclic) references do not need to correspond across the two values
  // - references to an *ancestor* (cycles) must correspond to the same ancestor
  /** @type {WeakMap<object, number>} */
  const aIndex = new WeakMap();
  /** @type {WeakMap<object, number>} */
  const bIndex = new WeakMap();
  return innerDeepEqual(a, b, aStack, bStack, aIndex, bIndex);
}

/**
 * @param {any} a
 * @param {any} b
 * @param {any[]} aStack
 * @param {any[]} bStack
 * @param {WeakMap<object, number>} aIndex
 * @param {WeakMap<object, number>} bIndex
 * @returns {boolean}
 */
function innerDeepEqual(a, b, aStack, bStack, aIndex, bIndex) {
  if (Object.is(a, b)) return true;

  if (a === null || b === null) return false;
  if (a === undefined || b === undefined) return false;

  const typeA = typeof a;
  const typeB = typeof b;
  if (typeA !== typeB) return false;

  // Primitives (and functions) that weren't `Object.is` equal are not equal.
  if (typeA !== "object") return false;

  // Objects: handle cycles (references to an ancestor in the current recursion stack).
  // Important: do NOT enforce a global a<->b bijection. Non-cyclic repeated references
  // are allowed to compare equal even if the two graphs don't share the same aliasing.
  const aPos = aIndex.get(a);
  if (aPos !== undefined) return bStack[aPos] === b;
  const bPos = bIndex.get(b);
  if (bPos !== undefined) return aStack[bPos] === a;

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

  // Track ancestry for cycle detection.
  const depth = aStack.length;
  aStack.push(a);
  bStack.push(b);
  aIndex.set(a, depth);
  bIndex.set(b, depth);

  try {
    if (Array.isArray(a)) return arrayEqual(a, b, aStack, bStack, aIndex, bIndex);

    if (protoA !== Object.prototype && protoA !== null) {
      // Unknown object type (class instances, Maps, Sets, Errors, Yjs types, …).
      // We intentionally avoid walking these (they can be cyclic / huge) and fall
      // back to reference equality (already handled by Object.is above).
      return false;
    }

    return plainObjectEqual(a, b, aStack, bStack, aIndex, bIndex);
  } finally {
    aStack.pop();
    bStack.pop();
    aIndex.delete(a);
    bIndex.delete(b);
  }
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
 * @param {any[]} aStack
 * @param {any[]} bStack
 * @param {WeakMap<object, number>} aIndex
 * @param {WeakMap<object, number>} bIndex
 */
function arrayEqual(a, b, aStack, bStack, aIndex, bIndex) {
  if (!Array.isArray(b)) return false;
  if (a.length !== b.length) return false;

  // Compare index entries, treating holes distinctly from `undefined`.
  for (let i = 0; i < a.length; i++) {
    const aHas = Object.prototype.hasOwnProperty.call(a, i);
    const bHas = Object.prototype.hasOwnProperty.call(b, i);
    if (aHas !== bHas) return false;
    if (aHas && !innerDeepEqual(a[i], b[i], aStack, bStack, aIndex, bIndex)) return false;
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
    if (!innerDeepEqual(a[key], b[key], aStack, bStack, aIndex, bIndex)) return false;
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
 * @param {any[]} aStack
 * @param {any[]} bStack
 * @param {WeakMap<object, number>} aIndex
 * @param {WeakMap<object, number>} bIndex
 */
function plainObjectEqual(a, b, aStack, bStack, aIndex, bIndex) {
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
    if (!innerDeepEqual(a[key], b[key], aStack, bStack, aIndex, bIndex)) return false;
  }
  return true;
}
