import * as Y from "yjs";
import {
  cloneYjsValue,
  getArrayRoot,
  getDocTypeConstructors,
  getMapRoot,
  getTextRoot,
  getYArray,
  getYMap,
  getYText,
  isYAbstractType,
  yjsValueToJson,
} from "../../../collab/yjs-utils/src/index.ts";

/**
 * @typedef {{ name: string, kind: "map" | "array" | "text" }} RootTypeSpec
 */

/**
 * @typedef {import("../../../collab/yjs-utils/src/index.ts").DocTypeConstructors} YjsTypeConstructors
 */

/**
 * Returns true if the placeholder contains no visible data (no map entries and
 * no non-deleted list items). In this case we can safely ignore the root for
 * snapshot/restore purposes because it cannot affect user-visible state.
 *
 * @param {any} value
 */
function isEmptyPlaceholderRoot(value) {
  if (!isYAbstractType(value)) return false;

  const map = value?._map;
  if (map instanceof Map) {
    for (const item of map.values()) {
      if (item && !item.deleted) return false;
    }
  }

  // For Array/Text roots, Yjs tracks the count of non-deleted items in `_length`.
  // Prefer it over scanning list items so we don't misclassify a non-empty root
  // that happens to have a long chain of deleted items.
  if (typeof value?._length === "number") {
    return value._length === 0;
  }

  let item = value?._start ?? null;
  for (let i = 0; item && i < 1000; i += 1) {
    if (!item.deleted) return false;
    item = item.right;
  }

  // If we hit our scan limit while there are still list items, we can't safely
  // conclude the root is empty. Treat it as non-empty to avoid dropping data.
  if (item) return false;

  return true;
}

/**
 * @param {any} value
 * @returns {string | null}
 */
function coerceString(value) {
  const text = getYText(value);
  if (text) return yjsValueToJson(text);
  if (typeof value === "string") return value;
  if (value == null) return null;
  return String(value);
}

function isRecord(value) {
  return value !== null && typeof value === "object" && !Array.isArray(value);
}

function isPlainObject(value) {
  if (!value || typeof value !== "object") return false;
  if (Array.isArray(value)) return false;
  const proto = Object.getPrototypeOf(value);
  return proto === Object.prototype || proto === null;
}

// Drawing ids can be authored via remote/shared state (sheet view state). Keep validation strict
// so version snapshots/restores can't be DoS'd by cloning pathological ids (e.g. multi-megabyte
// `drawings[*].id` strings or Y.Text values).
const MAX_DRAWING_ID_STRING_CHARS = 4096;

/**
 * @param {unknown} value
 * @returns {string | number | null}
 */
function normalizeDrawingIdValue(value) {
  const text = getYText(value);
  if (text) {
    // Avoid `text.toString()` for oversized ids: it would allocate a large JS string.
    if (typeof text.length === "number" && text.length > MAX_DRAWING_ID_STRING_CHARS) return null;
    value = yjsValueToJson(text);
  }

  if (typeof value === "string") {
    if (value.length > MAX_DRAWING_ID_STRING_CHARS) return null;
    const trimmed = value.trim();
    if (!trimmed) return null;
    return trimmed;
  }

  if (typeof value === "number") {
    if (!Number.isSafeInteger(value)) return null;
    return value;
  }

  return null;
}

/**
 * Convert a `drawings` list into JSON without materializing oversized `drawing.id` strings.
 *
 * @param {unknown} raw
 * @returns {any[] | null}
 */
function drawingsValueToJsonSafe(raw) {
  if (raw === null) return null;
  if (raw === undefined) return null;

  const yArr = getYArray(raw);
  const isArr = Array.isArray(raw);
  if (!yArr && !isArr) return null;

  /** @type {any[]} */
  const out = [];
  const len = yArr ? yArr.length : raw.length;

  for (let idx = 0; idx < len; idx += 1) {
    const entry = yArr ? yArr.get(idx) : raw[idx];

    const map = getYMap(entry);
    if (map) {
      const normalizedId = normalizeDrawingIdValue(map.get("id"));
      if (normalizedId == null) continue;

      /** @type {any} */
      const obj = { id: normalizedId };
      const keys = Array.from(map.keys()).sort();
      for (const key of keys) {
        if (key === "id") continue;
        obj[String(key)] = yjsValueToJson(map.get(key));
      }
      out.push(obj);
      continue;
    }

    if (isPlainObject(entry)) {
      const normalizedId = normalizeDrawingIdValue(entry.id);
      if (normalizedId == null) continue;

      /** @type {any} */
      const obj = { id: normalizedId };
      const keys = Object.keys(entry).sort();
      for (const key of keys) {
        if (key === "id") continue;
        obj[key] = yjsValueToJson(entry[key]);
      }
      out.push(obj);
    }
  }

  return out;
}

/**
 * Convert a sheet `view` object into JSON, treating `view.drawings` specially so we don't
 * materialize oversized `drawing.id` strings.
 *
 * @param {unknown} rawView
 * @returns {any}
 */
function sheetViewValueToJsonSafe(rawView) {
  if (rawView == null) return yjsValueToJson(rawView);

  const map = getYMap(rawView);
  if (map) {
    /** @type {Record<string, any>} */
    const out = {};
    const keys = Array.from(map.keys()).sort();
    for (const key of keys) {
      if (key === "drawings") {
        const rawDrawings = map.get(key);
        if (rawDrawings === null) out.drawings = null;
        else out.drawings = drawingsValueToJsonSafe(rawDrawings) ?? [];
        continue;
      }
      out[String(key)] = yjsValueToJson(map.get(key));
    }
    return out;
  }

  if (isPlainObject(rawView)) {
    /** @type {Record<string, any>} */
    const out = {};
    const keys = Object.keys(rawView).sort();
    for (const key of keys) {
      if (key === "drawings") {
        const rawDrawings = rawView.drawings;
        if (rawDrawings === null) out.drawings = null;
        else out.drawings = drawingsValueToJsonSafe(rawDrawings) ?? [];
        continue;
      }
      out[key] = yjsValueToJson(rawView[key]);
    }
    return out;
  }

  // Unknown/invalid view type. Avoid materializing it (e.g. huge Y.Text); treat as absent.
  return null;
}

/**
 * Clone a sheet entry map while sanitizing potentially large `drawings[*].id` values.
 *
 * This is used by version snapshots/restores when we have to clone a doc to filter
 * excluded roots.
 *
 * @param {any} entry
 * @param {YjsTypeConstructors} constructors
 * @returns {any}
 */
function cloneSheetEntryWithSanitizedView(entry, constructors) {
  const map = getYMap(entry);
  const MapCtor = constructors?.Map ?? (map ? map.constructor : Y.Map);
  const out = new MapCtor();

  if (map) {
    const keys = Array.from(map.keys()).sort();
    for (const key of keys) {
      if (key === "view") {
        out.set("view", sheetViewValueToJsonSafe(map.get(key)));
        continue;
      }
      if (key === "drawings") {
        const raw = map.get(key);
        out.set("drawings", raw === null ? null : drawingsValueToJsonSafe(raw) ?? []);
        continue;
      }
      out.set(String(key), cloneYjsValue(map.get(key), constructors));
    }
    return out;
  }

  if (isPlainObject(entry)) {
    const keys = Object.keys(entry).sort();
    for (const key of keys) {
      if (key === "view") {
        out.set("view", sheetViewValueToJsonSafe(entry[key]));
        continue;
      }
      if (key === "drawings") {
        const raw = entry[key];
        out.set("drawings", raw === null ? null : drawingsValueToJsonSafe(raw) ?? []);
        continue;
      }
      out.set(key, cloneYjsValue(entry[key], constructors));
    }
  }

  return out;
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
        const map = getYMap(value);
        if (map) out.push(map);
      }
    }
    item = item.right;
  }
  return out;
}

/**
 * Delete any legacy list items (sequence entries with `parentSub === null`) from
 * an instantiated map root.
 *
 * @param {Y.Transaction} transaction
 * @param {any} mapType
 */
function deleteLegacyListItemsFromMapRoot(transaction, mapType) {
  let item = mapType?._start ?? null;
  while (item) {
    if (!item.deleted && item.parentSub === null) {
      item.delete(transaction);
    }
    item = item.right;
  }
}

/**
 * Delete any map entries (keyed items) from an instantiated array root.
 *
 * This can happen if a map schema was instantiated as an Array: map entries are
 * stored in `type._map` and are invisible to `array.toArray()`.
 *
 * @param {Y.Transaction} transaction
 * @param {any} arrayType
 */
function deleteMapEntriesFromArrayRoot(transaction, arrayType) {
  const map = arrayType?._map;
  if (!(map instanceof Map)) return;
  for (const item of map.values()) {
    if (!item?.deleted) item.delete(transaction);
  }
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
    if (getYMap(value)) return "map";
    if (getYArray(value)) return "array";
    if (getYText(value)) return "text";

    // When applying a snapshot update into a doc that hasn't instantiated a
    // root type (via getMap/getArray/getText), Yjs represents that root as a
    // generic `AbstractType` placeholder. Infer the intended kind from the
    // placeholder's internal structure.
    if (isYAbstractType(value)) {
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
      /** @type {YjsTypeConstructors} */
      const snapshotConstructors = { Map: Y.Map, Array: Y.Array, Text: Y.Text };

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
          if (roots.has(name)) continue;
          if (isEmptyPlaceholderRoot(value)) continue;
          throw new Error(
            `Unsupported Yjs root type for "${name}" in current doc: ${value?.constructor?.name ?? typeof value}`,
          );
        }
        addRoot(name, kind, "current doc");
      }

      for (const [name, { kind }] of roots.entries()) {
        if (kind === "map") {
          const source = getMapRoot(doc, name);
          const target = snapshotDoc.getMap(name);
          source.forEach((value, key) => {
            target.set(key, cloneYjsValue(value, snapshotConstructors));
          });
          continue;
        }

        if (kind === "array") {
          const source = getArrayRoot(doc, name);
          const target = snapshotDoc.getArray(name);
          for (let i = 0; i < source.length; i++) {
            const value = source.get(i);
            target.push([
              name === "sheets"
                ? cloneSheetEntryWithSanitizedView(value, snapshotConstructors)
                : cloneYjsValue(value, snapshotConstructors),
            ]);
          }
          continue;
        }

        if (kind === "text") {
          const source = getTextRoot(doc, name);
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
      const docConstructors = getDocTypeConstructors(doc);

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
          if (roots.has(name)) continue;
          if (isEmptyPlaceholderRoot(value)) continue;
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
          if (roots.has(name)) continue;
          // Some root types can be present in the snapshot as an empty
          // `AbstractType` placeholder (no visible map/list items). In that case
          // the placeholder doesn't provide enough structure to infer whether it
          // was a Map/Array/Text. If the root is empty, skipping it is safe:
          // there is no user-visible content to restore.
          if (isEmptyPlaceholderRoot(value)) continue;
          throw new Error(
            `Unsupported Yjs root type for "${name}" in snapshot: ${value?.constructor?.name ?? typeof value}`,
          );
        }
        addRoot(name, kind, "snapshot");
      }

      doc.transact((transaction) => {
        // When the target doc comes from a different Yjs module instance (e.g. CJS vs ESM),
        // we must construct cloned values using the target instance's constructors. Otherwise
        // `target.set(key, value)` throws ("Unexpected content type").
        for (const [name, { kind }] of roots.entries()) {
          if (kind === "map") {
            const target = getMapRoot(doc, name);
            const source = getMapRoot(restored, name);


            for (const key of Array.from(target.keys())) {
              target.delete(key);
            }
            // Clear any legacy list items on the target root so restore doesn't
            // accidentally retain clobbered Array-schema content.
            if (name === "comments") {
              deleteLegacyListItemsFromMapRoot(transaction, target);
            }

            source.forEach((value, key) => {
              target.set(key, cloneYjsValue(value, docConstructors));
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
                target.set(id, cloneYjsValue(item, docConstructors));
              }
            }
            continue;
          }

          if (kind === "array") {
            const target = getArrayRoot(doc, name);
            const source = getArrayRoot(restored, name);

            if (name === "comments") {
              // Clear any clobbered map entries stored on the array root.
              deleteMapEntriesFromArrayRoot(transaction, target);
            }
            if (target.length > 0) {
              target.delete(0, target.length);
            }

            for (let i = 0; i < source.length; i++) {
              const value = source.get(i);
              target.push([
                name === "sheets"
                  ? cloneSheetEntryWithSanitizedView(value, docConstructors)
                  : cloneYjsValue(value, docConstructors),
              ]);
            }
            continue;
          }

          if (kind === "text") {
            const target = getTextRoot(doc, name);
            const source = getTextRoot(restored, name);
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
     * @returns {() => void} unsubscribe
     */
    on(event, listener) {
      if (event !== "update") {
        throw new Error(`Unsupported event: ${event}`);
      }
      if (!excludedRoots || excludedRoots.size === 0) {
        const wrappedListener = () => listener();
        doc.on("update", wrappedListener);
        return () => doc.off("update", wrappedListener);
      }

      const wrappedListener = (_update, _origin, _doc, transaction) => {
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
      };
      doc.on("update", wrappedListener);
      return () => doc.off("update", wrappedListener);
    },
  };
}
