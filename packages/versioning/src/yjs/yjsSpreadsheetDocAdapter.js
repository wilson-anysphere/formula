import * as Y from "yjs";
import { cloneYjsValue } from "./cloneYjsValue.js";

/**
 * @typedef {{ name: string, kind: "map" | "array" | "text" }} RootTypeSpec
 */

/**
 * @typedef {{ Map: new () => any, Array: new () => any, Text: new () => any }} YjsTypeConstructors
 */

/**
 * Return constructors for Y.Map/Y.Array/Y.Text that match the module instance
 * used to create `doc`.
 *
 * In pnpm workspaces it is possible to load both the ESM + CJS builds of Yjs in
 * the same process (for example via y-websocket). Yjs types cannot be moved
 * across module instances; the safest approach is to clone nested types using
 * constructors from the target doc's module instance.
 *
 * @param {any} doc
 * @returns {YjsTypeConstructors}
 */
function getDocConstructors(doc) {
  const DocCtor = /** @type {any} */ (doc)?.constructor;
  if (typeof DocCtor !== "function") {
    return { Map: Y.Map, Array: Y.Array, Text: Y.Text };
  }

  try {
    const probe = new DocCtor();
    return {
      Map: probe.getMap("__ctor_probe_map").constructor,
      Array: probe.getArray("__ctor_probe_array").constructor,
      Text: probe.getText("__ctor_probe_text").constructor,
    };
  } catch {
    return { Map: Y.Map, Array: Y.Array, Text: Y.Text };
  }
}

function isYMap(value) {
  if (value instanceof Y.Map) return true;
  if (!value || typeof value !== "object") return false;
  const maybe = /** @type {any} */ (value);
  // Bundlers can rename constructors and pnpm workspaces can load multiple `yjs`
  // module instances (ESM + CJS). Avoid relying on `constructor.name`; prefer a
  // structural check instead.
  if (typeof maybe.forEach !== "function") return false;
  if (typeof maybe.get !== "function") return false;
  if (typeof maybe.set !== "function") return false;
  if (typeof maybe.delete !== "function") return false;
  // Plain JS Maps also have get/set/delete/forEach, so additionally require Yjs'
  // deep observer APIs.
  if (typeof maybe.observeDeep !== "function") return false;
  if (typeof maybe.unobserveDeep !== "function") return false;
  return true;
}

function isYArray(value) {
  if (value instanceof Y.Array) return true;
  if (!value || typeof value !== "object") return false;
  const maybe = /** @type {any} */ (value);
  return (
    typeof maybe.toArray === "function" &&
    typeof maybe.get === "function" &&
    typeof maybe.push === "function" &&
    typeof maybe.delete === "function" &&
    typeof maybe.observeDeep === "function" &&
    typeof maybe.unobserveDeep === "function"
  );
}

function isYText(value) {
  if (value instanceof Y.Text) return true;
  if (!value || typeof value !== "object") return false;
  const maybe = /** @type {any} */ (value);
  return (
    typeof maybe.toDelta === "function" &&
    typeof maybe.applyDelta === "function" &&
    typeof maybe.insert === "function" &&
    typeof maybe.delete === "function" &&
    typeof maybe.observeDeep === "function" &&
    typeof maybe.unobserveDeep === "function"
  );
}

function isYAbstractType(value) {
  if (value instanceof Y.AbstractType) return true;
  if (!value || typeof value !== "object") return false;
  const maybe = /** @type {any} */ (value);
  // When different Yjs module instances are loaded (ESM vs CJS), `instanceof`
  // checks can fail even though the object is a valid Yjs type. Use a small
  // duck-type check so restores/snapshots work regardless of module loader.
  return Boolean(maybe._map instanceof Map || maybe._start || maybe._item || maybe._length != null);
}

function replaceForeignRootType(params) {
  const { doc, name, existing, create } = params;
  const t = create();

  // Mirror Yjs' own Doc.get conversion logic for AbstractType placeholders, but
  // also support roots instantiated by a different Yjs module instance (e.g.
  // CJS `require("yjs")`).
  //
  // We intentionally only do this replacement when `doc` is from this module's
  // Yjs instance (i.e. `doc instanceof Y.Doc`). If the entire doc was created by
  // a foreign Yjs build, inserting local types into it can cause the same
  // cross-instance integration errors we're trying to avoid.
  (t)._map = existing?._map;
  (t)._start = existing?._start;
  (t)._length = existing?._length;

  const map = existing?._map;
  if (map instanceof Map) {
    map.forEach((item) => {
      for (let n = item; n !== null; n = n.left) {
        n.parent = t;
      }
    });
  }

  for (let n = existing?._start ?? null; n !== null; n = n.right) {
    n.parent = t;
  }

  doc.share.set(name, t);
  t._integrate(doc, null);
  return t;
}

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
  if (isYText(value)) return value.toString();
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
        if (isYMap(value)) out.push(value);
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
 * Safely access a root type without relying on `doc.getMap/getArray/getText`
 * `instanceof` checks that can fail when the document contains types created by
 * a different Yjs module instance (ESM vs CJS).
 *
 * @param {Y.Doc} doc
 * @param {string} name
 */
function getMapRoot(doc, name) {
  const existing = doc.share.get(name);
  if (existing == null) return doc.getMap(name);

  if (isYMap(existing)) {
    // If the map root was created by a different Yjs module instance (ESM vs CJS),
    // `instanceof` checks fail and inserting local nested types can throw
    // ("Unexpected content type"). Normalize the root to this module instance.
    if (!(existing instanceof Y.Map) && doc instanceof Y.Doc) {
      return replaceForeignRootType({ doc, name, existing, create: () => new Y.Map() });
    }
    return existing;
  }

  if (isYArray(existing)) {
    throw new Error(`Unsupported Yjs root type for "${name}" in current doc: Y.Array`);
  }
  if (isYText(existing)) {
    throw new Error(`Unsupported Yjs root type for "${name}" in current doc: Y.Text`);
  }

  // Placeholder root types should be coerced via Yjs' own constructors.
  //
  // Note: other parts of the system patch foreign prototype chains so foreign
  // types can pass `instanceof Y.AbstractType` checks. Use constructor identity
  // to detect placeholders created by *this* Yjs module instance.
  if (existing instanceof Y.AbstractType && existing.constructor === Y.AbstractType) return doc.getMap(name);
  if (isYAbstractType(existing) && doc instanceof Y.Doc) {
    return replaceForeignRootType({ doc, name, existing, create: () => new Y.Map() });
  }
  if (isYAbstractType(existing)) return doc.getMap(name);

  throw new Error(`Unsupported Yjs root type for "${name}" in current doc`);
}

/**
 * @param {Y.Doc} doc
 * @param {string} name
 */
function getArrayRoot(doc, name) {
  const existing = doc.share.get(name);
  if (existing == null) return doc.getArray(name);

  if (isYArray(existing)) {
    if (!(existing instanceof Y.Array) && doc instanceof Y.Doc) {
      return replaceForeignRootType({ doc, name, existing, create: () => new Y.Array() });
    }
    return existing;
  }

  if (isYMap(existing)) {
    throw new Error(`Unsupported Yjs root type for "${name}" in current doc: Y.Map`);
  }
  if (isYText(existing)) {
    throw new Error(`Unsupported Yjs root type for "${name}" in current doc: Y.Text`);
  }

  if (existing instanceof Y.AbstractType && existing.constructor === Y.AbstractType) return doc.getArray(name);
  if (isYAbstractType(existing) && doc instanceof Y.Doc) {
    return replaceForeignRootType({ doc, name, existing, create: () => new Y.Array() });
  }
  if (isYAbstractType(existing)) return doc.getArray(name);

  throw new Error(`Unsupported Yjs root type for "${name}" in current doc`);
}

/**
 * @param {Y.Doc} doc
 * @param {string} name
 */
function getTextRoot(doc, name) {
  const existing = doc.share.get(name);
  if (existing == null) return doc.getText(name);

  if (isYText(existing)) {
    if (!(existing instanceof Y.Text) && doc instanceof Y.Doc) {
      return replaceForeignRootType({ doc, name, existing, create: () => new Y.Text() });
    }
    return existing;
  }

  if (isYMap(existing)) {
    throw new Error(`Unsupported Yjs root type for "${name}" in current doc: Y.Map`);
  }
  if (isYArray(existing)) {
    throw new Error(`Unsupported Yjs root type for "${name}" in current doc: Y.Array`);
  }

  if (existing instanceof Y.AbstractType && existing.constructor === Y.AbstractType) return doc.getText(name);
  if (isYAbstractType(existing) && doc instanceof Y.Doc) {
    return replaceForeignRootType({ doc, name, existing, create: () => new Y.Text() });
  }
  if (isYAbstractType(existing)) return doc.getText(name);

  throw new Error(`Unsupported Yjs root type for "${name}" in current doc`);
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
    if (isYMap(value)) return "map";
    if (isYArray(value)) return "array";
    if (isYText(value)) return "text";

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
          for (const value of source.toArray()) {
            target.push([cloneYjsValue(value, snapshotConstructors)]);
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
      const docConstructors = getDocConstructors(doc);

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

            for (const value of source.toArray()) {
              target.push([cloneYjsValue(value, docConstructors)]);
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
