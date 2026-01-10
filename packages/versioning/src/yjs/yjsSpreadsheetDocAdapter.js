import * as Y from "yjs";
import { cloneYjsValue } from "./cloneYjsValue.js";

/**
 * @typedef {{ name: string, kind: "map" | "array" }} RootTypeSpec
 */

/**
 * Create a VersionManager-compatible adapter around a Y.Doc.
 *
 * Note: restoring a snapshot is implemented by replacing known root types
 * (cells/sheets/metadata/namedRanges). This keeps the doc instance stable so
 * other systems (providers/awareness) can keep references to it.
 *
 * @param {Y.Doc} doc
 * @param {{ roots?: RootTypeSpec[] }} [opts]
 */
export function createYjsSpreadsheetDocAdapter(doc, opts = {}) {
  /** @type {RootTypeSpec[]} */
  const roots = opts.roots ?? [
    { name: "sheets", kind: "array" },
    { name: "cells", kind: "map" },
    { name: "metadata", kind: "map" },
    { name: "namedRanges", kind: "map" },
  ];

  return {
    encodeState() {
      return Y.encodeStateAsUpdate(doc);
    },
    /**
     * @param {Uint8Array} snapshot
     */
    applyState(snapshot) {
      const restored = new Y.Doc();
      Y.applyUpdate(restored, snapshot);

      doc.transact(() => {
        for (const root of roots) {
          if (root.kind === "map") {
            const target = doc.getMap(root.name);
            const source = restored.getMap(root.name);

            for (const key of Array.from(target.keys())) {
              target.delete(key);
            }

            source.forEach((value, key) => {
              target.set(key, cloneYjsValue(value));
            });
            continue;
          }

          if (root.kind === "array") {
            const target = doc.getArray(root.name);
            const source = restored.getArray(root.name);

            if (target.length > 0) {
              target.delete(0, target.length);
            }

            for (const value of source.toArray()) {
              target.push([cloneYjsValue(value)]);
            }
            continue;
          }

          /** @type {never} */
          const _exhaustive = root;
          throw new Error(`Unsupported root kind: ${_exhaustive}`);
        }
      }, "versioning-restore");
    },
    /**
     * @param {"update"} event
     * @param {() => void} listener
     */
    on(event, listener) {
      if (event !== "update") {
        throw new Error(`Unsupported event: ${event}`);
      }
      doc.on("update", () => listener());
    },
  };
}

