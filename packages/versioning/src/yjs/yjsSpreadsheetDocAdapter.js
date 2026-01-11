import * as Y from "yjs";
import { cloneYjsValue } from "./cloneYjsValue.js";

/**
 * @typedef {{ name: string, kind: "map" | "array" | "text" }} RootTypeSpec
 */

/**
 * Create a VersionManager-compatible adapter around a Y.Doc.
 *
 * Note: restoring a snapshot is implemented by mutating the current `doc` in
 * place (clearing and rehydrating root types). This keeps the doc instance
 * stable so other systems (providers/awareness) can keep references to it.
 *
 * @param {Y.Doc} doc
 * @param {{ roots?: RootTypeSpec[] }} [opts]
 */
export function createYjsSpreadsheetDocAdapter(doc, opts = {}) {
  /** @type {RootTypeSpec[] | null} */
  const configuredRoots = opts.roots ?? null;

  /**
   * Best-effort kind detection for non-default roots when restoring snapshots.
   *
   * Note: after `Y.applyUpdate` into a fresh doc, root types can exist as a
   * generic `AbstractType` until a constructor is chosen via `getMap/getArray`.
   * For roots with content we can infer Map vs Array by inspecting the internal
   * state.
   *
   * @param {Y.Doc} snapshotDoc
   * @param {string} name
   * @returns {"map" | "array" | null}
   */
  function detectSnapshotRootKind(snapshotDoc, name) {
    const placeholder = snapshotDoc.share.get(name);
    if (!placeholder) return null;
    if (placeholder instanceof Y.Map) return "map";
    if (placeholder instanceof Y.Array) return "array";

    const hasStart = placeholder?._start != null;
    const mapSize = placeholder?._map instanceof Map ? placeholder._map.size : 0;

    if (mapSize > 0) return "map";
    if (hasStart) return "array";
    return null;
  }

  /** @returns {RootTypeSpec[]} */
  function resolveRoots() {
    if (configuredRoots) return configuredRoots;

    // Default spreadsheet roots. We seed these so the adapter works even if a
    // doc hasn't touched all root types yet.
    /** @type {Map<string, RootTypeSpec>} */
    const roots = new Map([
      ["sheets", { name: "sheets", kind: "array" }],
      ["cells", { name: "cells", kind: "map" }],
      ["metadata", { name: "metadata", kind: "map" }],
      ["namedRanges", { name: "namedRanges", kind: "map" }],
    ]);

    // Add any other root types already defined in this doc. Note that Yjs root
    // types are schema-defined: you must know whether a key is an Array or Map.
    // We can safely restore additional roots that are already instantiated in
    // the current doc (e.g. comments).
    for (const [name, value] of doc.share.entries()) {
      if (roots.has(name)) continue;
      if (value instanceof Y.Map) roots.set(name, { name, kind: "map" });
      else if (value instanceof Y.Array) roots.set(name, { name, kind: "array" });
      else if (value instanceof Y.Text) roots.set(name, { name, kind: "text" });
    }

    return Array.from(roots.values());
  }

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

      const roots = resolveRoots();

      // Best-effort: if the snapshot contains collaboration roots (like
      // comments) that haven't been instantiated in the current doc yet, add
      // them so restoration doesn't silently drop data.
      if (!configuredRoots) {
        const names = new Set(roots.map((root) => root.name));
        if (!names.has("comments") && restored.share.has("comments")) {
          const kind = detectSnapshotRootKind(restored, "comments");
          if (kind) roots.push({ name: "comments", kind });
        }
      }

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

          if (root.kind === "text") {
            const target = doc.getText(root.name);
            const source = restored.getText(root.name);
            if (target.length > 0) target.delete(0, target.length);
            const text = source.toString();
            if (text) target.insert(0, text);
            continue;
          }
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
