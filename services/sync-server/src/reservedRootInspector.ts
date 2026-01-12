import { Y } from "./yjs.js";

export type ReservedRootKeyPathSegment = string | number;

export type ReservedRootInspectionHit = {
  root: string;
  keyPath: ReservedRootKeyPathSegment[];
};

function collectKeysFromEventChanges(event: any): string[] {
  const keys = event?.changes?.keys;
  if (!keys) return [];

  const out: string[] = [];
  if (typeof keys.entries === "function") {
    for (const [key] of keys.entries()) {
      if (typeof key === "string" && key.length > 0) out.push(key);
    }
    return out;
  }

  if (typeof keys.keys === "function") {
    for (const key of keys.keys()) {
      if (typeof key === "string" && key.length > 0) out.push(key);
    }
  }
  return out;
}

/**
 * Best-effort inspection helper that attributes a Yjs update to specific reserved roots
 * and nested map-key paths.
 *
 * Why this exists: updates that mutate nested Y.Types (e.g. pushing to a `Y.Array` stored
 * under a map key like `versionsMeta.order`) encode leaf Items with `parentSub=null`.
 * To attribute the mutation to the reserved root + key path, we must rely on Yjs'
 * internal type graph (available when the update is applied against a doc seeded with
 * the current server state).
 *
 * This helper uses a shadow doc (seeded from `baseDoc`) and `observeDeep` to capture
 * the path information that would otherwise be lost when only looking at leaf Items.
 */
export function inspectReservedRootUpdate(params: {
  baseDoc: Y.Doc;
  update: Uint8Array;
  reservedRoots: string[];
}): ReservedRootInspectionHit[] {
  const { baseDoc, update, reservedRoots } = params;
  const roots = Array.from(new Set(reservedRoots)).filter((r) => typeof r === "string" && r.length > 0);
  if (roots.length === 0) return [];

  const shadow = new Y.Doc();

  // Seed shadow doc with the current server state so parent type insertion items exist.
  try {
    Y.applyUpdate(shadow, Y.encodeStateAsUpdate(baseDoc));
  } catch {
    // Best-effort; if seeding fails we still try to apply the update, but deep path
    // resolution may be incomplete (pending structs).
  }

  const hitsByRoot = new Map<string, Map<string, ReservedRootKeyPathSegment[]>>();
  const observers: Array<{ rootType: any; observer: (events: any[]) => void }> = [];

  for (const rootName of roots) {
    const hitsForRoot = new Map<string, ReservedRootKeyPathSegment[]>();
    hitsByRoot.set(rootName, hitsForRoot);

    // We primarily care about roots that are Maps (e.g. versions/versionsMeta), but
    // avoid failing if the root isn't instantiated yet.
    let rootType: any;
    try {
      rootType = shadow.share.get(rootName);
      if (!rootType) rootType = shadow.getMap(rootName);
    } catch {
      // If the root is an Array/Text (schema mismatch), fall back to an `AbstractType`
      // placeholder lookup.
      rootType = shadow.share.get(rootName);
    }

    if (!rootType || typeof rootType.observeDeep !== "function") continue;

    const observer = (events: any[]) => {
      for (const event of events ?? []) {
        const path = Array.isArray(event?.path) ? (event.path as any[]) : [];
        const changedKeys = collectKeysFromEventChanges(event);

        if (changedKeys.length > 0) {
          for (const key of changedKeys) {
            const keyPath = [...path, key] as ReservedRootKeyPathSegment[];
            const id = JSON.stringify(keyPath);
            if (!hitsForRoot.has(id)) hitsForRoot.set(id, keyPath);
          }
          continue;
        }

        if (path.length > 0) {
          const keyPath = path as ReservedRootKeyPathSegment[];
          const id = JSON.stringify(keyPath);
          if (!hitsForRoot.has(id)) hitsForRoot.set(id, keyPath);
        }
      }
    };

    try {
      rootType.observeDeep(observer);
      observers.push({ rootType, observer });
    } catch {
      // ignore
    }
  }

  try {
    Y.applyUpdate(shadow, update);
  } catch {
    // ignore (best-effort decoding)
  } finally {
    for (const { rootType, observer } of observers) {
      try {
        rootType.unobserveDeep(observer);
      } catch {
        // ignore
      }
    }
    try {
      shadow.destroy();
    } catch {
      // ignore
    }
  }

  const out: ReservedRootInspectionHit[] = [];
  for (const [root, hits] of hitsByRoot.entries()) {
    for (const keyPath of hits.values()) {
      out.push({ root, keyPath });
    }
  }
  return out;
}

