import * as Y from "yjs";

/**
 * Deep-clone a Yjs value from one document into a freshly-constructed Yjs type
 * (not yet integrated into any document).
 *
 * This is used for restoring snapshots: Yjs types cannot be moved between docs,
 * so we recreate equivalent content in the target doc.
 *
 * @param {any} value
 * @returns {any}
 */
export function cloneYjsValue(value) {
  if (value instanceof Y.Map) {
    const out = new Y.Map();
    value.forEach((v, k) => {
      out.set(k, cloneYjsValue(v));
    });
    return out;
  }

  if (value instanceof Y.Array) {
    const out = new Y.Array();
    for (const item of value.toArray()) {
      out.push([cloneYjsValue(item)]);
    }
    return out;
  }

  if (value instanceof Y.Text) {
    const out = new Y.Text();
    out.insert(0, value.toString());
    return out;
  }

  // For plain JSON-ish values we can use structuredClone to avoid sharing
  // object identity between docs (Yjs stores JSON by-value).
  if (value && typeof value === "object") {
    return structuredClone(value);
  }

  return value;
}

