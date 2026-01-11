import * as Y from "yjs";
import { cloneYjsValue } from "./cloneYjsValue.js";

/**
 * @typedef {{ name: string, kind: "map" | "array" | "text" }} RootTypeSpec
 */

/**
 * @param {any} value
 * @returns {string | null}
 */
function coerceString(value) {
  if (value instanceof Y.Text) return value.toString();
  if (typeof value === "string") return value;
  if (value == null) return null;
  return String(value);
}

/**
 * Recover list items (sequence entries with `parentSub === null`) stored on a map
 * root.
 *
 * This can happen if a document originally used a legacy Array schema, but the
 * root was later instantiated as a Map (e.g. by calling `doc.getMap(name)` first
 * while the root was still a placeholder). In that case the list content is
 * invisible via `map.keys()` but still exists in the CRDT.
 *
 * @param {any} mapType
 * @returns {Y.Map<any>[]}
 */
function legacyListItemsFromMapRoot(mapType) {
  /** @type {Y.Map<any>[]} */
  const out = [];
  let item = mapType?._start ?? null;
  while (item) {
    if (!item.deleted && item.parentSub === null) {
      const content = item.content?.getContent?.() ?? [];
      for (const value of content) {
        if (value instanceof Y.Map) out.push(value);
      }
    }
    item = item.right;
  }
  return out;
}

/**
 * Create a VersionManager-compatible adapter around a Y.Doc.
 *
 * Note: restoring a snapshot is implemented by mutating the current `doc` in
 * place (clearing and rehydrating root types). This keeps the doc instance
 * stable so other systems (providers/awareness) can keep references to it.
 *
 * @param {Y.Doc} doc
 * @param {{ roots?: RootTypeSpec[], excludeRoots?: string[] }} [opts]
 */
export function createYjsSpreadsheetDocAdapter(doc, opts = {}) {
  /** @type {RootTypeSpec[] | null} */
  const configuredRoots = opts.roots ?? null;
  const excludedRoots = Array.isArray(opts.excludeRoots) ? new Set(opts.excludeRoots) : null;

  /**
   * @param {string} name
   */
  function isExcludedRoot(name) {
    return Boolean(excludedRoots?.has(name));
  }

  /**
   * @param {unknown} value
   * @returns {RootTypeSpec["kind"] | null}
   */
  function rootKindFromValue(value) {
    if (value instanceof Y.Map) return "map";
    if (value instanceof Y.Array) return "array";
    if (value instanceof Y.Text) return "text";

    // When applying a snapshot update into a doc that hasn't instantiated a
    // root type (via getMap/getArray/getText), Yjs represents that root as a
    // generic `AbstractType` placeholder. Infer the intended kind from the
    // placeholder's internal structure.
    if (value instanceof Y.AbstractType) {
      if (value._map instanceof Map && value._map.size > 0) {
        return "map";
      }

      if (value._start) {
        let item = value._start;
        for (let i = 0; item && i < 1000; i += 1) {
          if (!item.deleted) {
            const content = item.content;
            if (content && typeof content === "object") {
              if ("str" in content) return "text";
              if ("key" in content && "value" in content) return "text";
              if ("embed" in content) return "text";
            }
            return "array";
          }
          item = item.right;
        }
      }
    }
    return null;
  }

  return {
    encodeState() {
      if (!excludedRoots || excludedRoots.size === 0) {
        return Y.encodeStateAsUpdate(doc);
      }

      // Fast path: if none of the excluded roots exist, there is nothing to filter.
      let hasExcluded = false;
      for (const name of excludedRoots) {
        if (doc.share.has(name)) {
          hasExcluded = true;
          break;
        }
      }
      if (!hasExcluded) {
        return Y.encodeStateAsUpdate(doc);
      }

      const snapshotDoc = new Y.Doc();

      /** @type {Map<string, { kind: RootTypeSpec["kind"], source: string }>} */
      const roots = new Map();

      /**
       * @param {string} name
       * @param {RootTypeSpec["kind"]} kind
       * @param {string} source
       */
      function addRoot(name, kind, source) {
        if (isExcludedRoot(name)) return;
        const existing = roots.get(name);
        if (!existing) {
          roots.set(name, { kind, source });
          return;
        }
        if (existing.kind !== kind) {
          throw new Error(
            `Yjs root schema mismatch for "${name}": ${existing.source} is "${existing.kind}" but ${source} is "${kind}"`,
          );
        }
      }

      if (configuredRoots) {
        for (const root of configuredRoots) {
          addRoot(root.name, root.kind, "configured roots");
        }
      } else {
        addRoot("sheets", "array", "default roots");
        addRoot("cells", "map", "default roots");
        addRoot("metadata", "map", "default roots");
        addRoot("namedRanges", "map", "default roots");
      }

      for (const [name, value] of doc.share.entries()) {
        if (isExcludedRoot(name)) continue;
        const kind = rootKindFromValue(value);
        if (!kind) {
          throw new Error(
            `Unsupported Yjs root type for "${name}" in current doc: ${value?.constructor?.name ?? typeof value}`,
          );
        }
        addRoot(name, kind, "current doc");
      }

      for (const [name, { kind }] of roots.entries()) {
        if (kind === "map") {
          const source = doc.getMap(name);
          const target = snapshotDoc.getMap(name);
          source.forEach((value, key) => {
            target.set(key, cloneYjsValue(value));
          });
          continue;
        }

        if (kind === "array") {
          const source = doc.getArray(name);
          const target = snapshotDoc.getArray(name);
          for (const value of source.toArray()) {
            target.push([cloneYjsValue(value)]);
          }
          continue;
        }

        if (kind === "text") {
          const source = doc.getText(name);
          const target = snapshotDoc.getText(name);
          target.applyDelta(structuredClone(source.toDelta()));
          continue;
        }
      }

      return Y.encodeStateAsUpdate(snapshotDoc);
    },
    /**
     * @param {Uint8Array} snapshot
     */
    applyState(snapshot) {
      const restored = new Y.Doc();
      Y.applyUpdate(restored, snapshot);

      /** @type {Map<string, { kind: RootTypeSpec["kind"], source: string }>} */
      const roots = new Map();

      /**
       * @param {string} name
       * @param {RootTypeSpec["kind"]} kind
       * @param {string} source
       */
      function addRoot(name, kind, source) {
        if (isExcludedRoot(name)) return;
        const existing = roots.get(name);
        if (!existing) {
          roots.set(name, { kind, source });
          return;
        }
        if (existing.kind !== kind) {
          throw new Error(
            `Yjs root schema mismatch for "${name}": ${existing.source} is "${existing.kind}" but ${source} is "${kind}"`,
          );
        }
      }

      if (configuredRoots) {
        for (const root of configuredRoots) {
          addRoot(root.name, root.kind, "configured roots");
        }
      } else {
        // Default spreadsheet roots. We seed these so the adapter works even if a
        // doc hasn't touched all root types yet.
        addRoot("sheets", "array", "default roots");
        addRoot("cells", "map", "default roots");
        addRoot("metadata", "map", "default roots");
        addRoot("namedRanges", "map", "default roots");
      }

      // Include any other root types already instantiated in either the current
      // doc or the snapshot doc so restoring doesn't silently drop data.
      for (const [name, value] of doc.share.entries()) {
        if (isExcludedRoot(name)) continue;
        const kind = rootKindFromValue(value);
        if (!kind) {
          throw new Error(
            `Unsupported Yjs root type for "${name}" in current doc: ${value?.constructor?.name ?? typeof value}`,
          );
        }
        addRoot(name, kind, "current doc");
      }

      for (const [name, value] of restored.share.entries()) {
        if (isExcludedRoot(name)) continue;
        const kind = rootKindFromValue(value);
        if (!kind) {
          throw new Error(
            `Unsupported Yjs root type for "${name}" in snapshot: ${value?.constructor?.name ?? typeof value}`,
          );
        }
        addRoot(name, kind, "snapshot");
      }

      doc.transact(() => {
        for (const [name, { kind }] of roots.entries()) {
          if (kind === "map") {
            const target = doc.getMap(name);
            const source = restored.getMap(name);
       

            for (const key of Array.from(target.keys())) {
              target.delete(key);
            }

            source.forEach((value, key) => {
              target.set(key, cloneYjsValue(value));
            });

            // Special-case: comments historically existed as a list (Array) but
            // could be accidentally instantiated as a Map. If that happens, the
            // legacy list items still exist on the Map root (as list entries with
            // `parentSub === null`) but are invisible via `map.keys()`. Preserve
            // them by migrating into proper map entries keyed by comment id.
            if (name === "comments") {
              for (const item of legacyListItemsFromMapRoot(source)) {
                const id = coerceString(item.get("id"));
                if (!id) continue;
                if (target.has(id)) continue;
                target.set(id, cloneYjsValue(item));
              }
            }
            continue;
          }

          if (kind === "array") {
            const target = doc.getArray(name);
            const source = restored.getArray(name);

            if (target.length > 0) {
              target.delete(0, target.length);
            }

            for (const value of source.toArray()) {
              target.push([cloneYjsValue(value)]);
            }
            continue;
          }

          if (kind === "text") {
            const target = doc.getText(name);
            const source = restored.getText(name);
            if (target.length > 0) target.delete(0, target.length);
            target.applyDelta(structuredClone(source.toDelta()));
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
      if (!excludedRoots || excludedRoots.size === 0) {
        doc.on("update", () => listener());
        return;
      }

      doc.on("update", (_update, _origin, _doc, transaction) => {
        // We only want to surface changes that touch non-excluded roots.
        // When using YjsVersionStore the version-history itself lives inside the
        // same Y.Doc. Without this filter, saving/pruning versions would mark the
        // workbook as dirty and trigger redundant snapshots.
        const changedParentTypes = /** @type {any} */ (transaction)?.changedParentTypes;
        const changedTypes = /** @type {any} */ (transaction)?.changed;

        if (!(changedParentTypes instanceof Map) && !(changedTypes instanceof Map)) {
          // Defensive fallback: if we can't introspect the transaction, treat it
          // as a meaningful update rather than risking missed changes.
          listener();
          return;
        }

        const hasTypeChange = (type) =>
          (changedParentTypes instanceof Map && changedParentTypes.has(type)) ||
          (changedTypes instanceof Map && changedTypes.has(type));

        for (const [name, value] of doc.share.entries()) {
          if (isExcludedRoot(name)) continue;
          if (hasTypeChange(value)) {
            listener();
            return;
          }
        }
      });
    },
  };
}
