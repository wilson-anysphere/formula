import * as Y from "yjs";

function isYMap(value) {
  if (value instanceof Y.Map) return true;
  if (!value || typeof value !== "object") return false;
  const maybe = /** @type {any} */ (value);
  if (maybe.constructor?.name !== "YMap") return false;
  return typeof maybe.forEach === "function" && typeof maybe.get === "function" && typeof maybe.set === "function";
}

function isYArray(value) {
  if (value instanceof Y.Array) return true;
  if (!value || typeof value !== "object") return false;
  const maybe = /** @type {any} */ (value);
  if (maybe.constructor?.name !== "YArray") return false;
  return (
    typeof maybe.toArray === "function" &&
    typeof maybe.get === "function" &&
    typeof maybe.push === "function" &&
    typeof maybe.delete === "function"
  );
}

function isYText(value) {
  if (value instanceof Y.Text) return true;
  if (!value || typeof value !== "object") return false;
  const maybe = /** @type {any} */ (value);
  if (maybe.constructor?.name !== "YText") return false;
  return (
    typeof maybe.toDelta === "function" &&
    typeof maybe.applyDelta === "function" &&
    typeof maybe.insert === "function" &&
    typeof maybe.delete === "function"
  );
}

/**
 * Deep-clone a Yjs value from one document into a freshly-constructed Yjs type
 * (not yet integrated into any document).
 *
 * This is used for restoring snapshots: Yjs types cannot be moved between docs,
 * so we recreate equivalent content in the target doc.
 *
 * @param {any} value
 * @param {{ Map?: new () => any, Array?: new () => any, Text?: new () => any } | null} [constructors]
 * @returns {any}
 */
export function cloneYjsValue(value, constructors = null) {
  if (isYMap(value)) {
    // NOTE: Yjs types cannot be moved between documents (or between different Yjs
    // module instances). Construct new instances using the source value's
    // constructor so the clone matches the module instance that produced the
    // original. This is critical when a doc has been hydrated by a different
    // loader (e.g. CJS vs ESM) and `instanceof` checks are unreliable.
    const MapCtor = constructors?.Map ?? /** @type {any} */ (value).constructor;
    const out = new MapCtor();
    value.forEach((v, k) => {
      out.set(k, cloneYjsValue(v, constructors));
    });
    return out;
  }

  if (isYArray(value)) {
    const ArrayCtor = constructors?.Array ?? /** @type {any} */ (value).constructor;
    const out = new ArrayCtor();
    for (const item of value.toArray()) {
      out.push([cloneYjsValue(item, constructors)]);
    }
    return out;
  }

  if (isYText(value)) {
    const TextCtor = constructors?.Text ?? /** @type {any} */ (value).constructor;
    const out = new TextCtor();
    // Preserve formatting by cloning the Y.Text delta instead of just its plain string.
    // Note: it's safe to call `applyDelta` on an un-integrated Y.Text, but you
    // must integrate it into a Y.Doc before reading it (toString/toDelta).
    out.applyDelta(structuredClone(value.toDelta()));
    return out;
  }

  // For plain JSON-ish values we can use structuredClone to avoid sharing
  // object identity between docs (Yjs stores JSON by-value).
  if (value && typeof value === "object") {
    return structuredClone(value);
  }

  return value;
}
