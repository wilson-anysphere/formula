const CIRCULAR_PLACEHOLDER = "[Circular]";
const UNSERIALIZABLE_PLACEHOLDER = "[Unserializable]";

/**
 * A tiny, dependency-free stable JSON stringify.
 *
 * This is equivalent to `JSON.stringify` for JSON-compatible inputs, but guarantees
 * stable object key ordering across runs and is resilient to BigInt/cycles.
 */
export function stableStringify(value: unknown): string {
  // JSON.stringify can return `undefined` for unsupported top-level inputs
  // (e.g. `undefined`). Since this helper returns `string`, normalize those
  // cases to `"null"` for deterministic output.
  return JSON.stringify(stableJsonValue(value, new WeakSet())) ?? "null";
}

/**
 * Convert an arbitrary value to a deterministic, human-readable string.
 *
 * - Strings are returned as-is.
 * - Other values are serialized via `stableJsonValue` and then JSON-stringified.
 */
export function stableValueToDisplayString(value: unknown): string {
  if (typeof value === "string") return value;

  const stable = stableJsonValue(value, new WeakSet());
  if (typeof stable === "string") return stable;

  return JSON.stringify(stable) ?? "null";
}

export function stableJsonValue(value: unknown, ancestors: WeakSet<object>): unknown {
  if (value === null) return null;

  const t = typeof value;
  if (t === "string" || t === "number" || t === "boolean") return value;
  // `t` is derived from `typeof value`, but TypeScript can't use that to narrow `value` here.
  // Use an explicit cast so bigint values remain serializable under strict TS settings.
  if (t === "bigint") return (value as bigint).toString();
  if (t === "undefined" || t === "function" || t === "symbol") return undefined;

  if (Array.isArray(value)) {
    if (ancestors.has(value)) return CIRCULAR_PLACEHOLDER;
    ancestors.add(value);
    const out: unknown[] = [];
    for (let i = 0; i < value.length; i++) {
      let item: unknown;
      try {
        item = value[i];
      } catch {
        item = UNSERIALIZABLE_PLACEHOLDER;
      }
      out.push(stableJsonValue(item, ancestors));
    }
    ancestors.delete(value);
    return out;
  }

  if (t !== "object") return undefined;

  const obj = value as Record<string, unknown>;

  if (ancestors.has(obj)) return CIRCULAR_PLACEHOLDER;

  // Preserve JSON.stringify behavior for objects with toJSON (e.g. Date), while
  // staying resilient to pathological implementations.
  //
  // Notes:
  // - Some `toJSON()` implementations return `this` (a no-op for JSON.stringify).
  //   Treat that as "no toJSON" so we still serialize the object's enumerable
  //   properties instead of collapsing to `[Circular]`.
  // - Other `toJSON()` implementations can create recursion (e.g. `toJSON() { return { self: this } }`).
  //   Track the object in `ancestors` while serializing the `toJSON()` result so any
  //   references back to the original object are replaced with a stable placeholder.
  if (typeof (obj as { toJSON?: unknown }).toJSON === "function") {
    let replacement: unknown;
    try {
      replacement = (obj as { toJSON: () => unknown }).toJSON();
    } catch {
      return UNSERIALIZABLE_PLACEHOLDER;
    }

    if (replacement !== obj) {
      ancestors.add(obj);
      try {
        return stableJsonValue(replacement, ancestors);
      } catch {
        return UNSERIALIZABLE_PLACEHOLDER;
      } finally {
        ancestors.delete(obj);
      }
    }
    // replacement === obj => fall through to normal object serialization.
  }

  ancestors.add(obj);

  // Use a null-prototype object to avoid special-casing keys like `__proto__`
  // (which can otherwise mutate the prototype chain and drop data).
  const sorted: Record<string, unknown> = Object.create(null);
  let keys: string[];
  try {
    keys = Object.keys(obj).sort();
  } catch {
    ancestors.delete(obj);
    return UNSERIALIZABLE_PLACEHOLDER;
  }

  for (const key of keys) {
    let child: unknown;
    try {
      child = obj[key];
    } catch {
      child = UNSERIALIZABLE_PLACEHOLDER;
    }
    sorted[key] = stableJsonValue(child, ancestors);
  }

  ancestors.delete(obj);
  return sorted;
}
