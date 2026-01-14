import { getYArray, getYMap, getYText } from "../../../collab/yjs-utils/src/index.ts";

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
  const map = getYMap(value);
  if (map) {
    // NOTE: Yjs types cannot be moved between documents (or between different Yjs
    // module instances). Construct new instances using the source value's
    // constructor so the clone matches the module instance that produced the
    // original. This is critical when a doc has been hydrated by a different
    // loader (e.g. CJS vs ESM) and `instanceof` checks are unreliable.
    const MapCtor = constructors?.Map ?? /** @type {any} */ (map).constructor;
    const out = new MapCtor();
    map.forEach((v, k) => {
      out.set(k, cloneYjsValue(v, constructors));
    });
    return out;
  }

  const array = getYArray(value);
  if (array) {
    const ArrayCtor = constructors?.Array ?? /** @type {any} */ (array).constructor;
    const out = new ArrayCtor();
    for (const item of array.toArray()) {
      out.push([cloneYjsValue(item, constructors)]);
    }
    return out;
  }

  const text = getYText(value);
  if (text) {
    const TextCtor = constructors?.Text ?? /** @type {any} */ (text).constructor;
    const out = new TextCtor();
    // Preserve formatting by cloning the Y.Text delta instead of just its plain string.
    // Note: it's safe to call `applyDelta` on an un-integrated Y.Text, but you
    // must integrate it into a Y.Doc before reading it (toString/toDelta).
    out.applyDelta(structuredClone(text.toDelta()));
    return out;
  }

  // For plain JSON-ish values we can use structuredClone to avoid sharing
  // object identity between docs (Yjs stores JSON by-value).
  if (value && typeof value === "object") {
    return structuredClone(value);
  }

  return value;
}
