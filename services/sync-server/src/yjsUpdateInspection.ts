import { Y } from "./yjs.js";

export type ReservedRootTouchKind = "insert" | "delete" | "gc" | "unknown";

export type ReservedRootTouch = {
  root: string;
  keyPath: string[];
  kind: ReservedRootTouchKind;
};

export type InspectUpdateResult = {
  touchesReserved: boolean;
  touches: ReservedRootTouch[];
  unknownReason?: string;
};

export type CollectTouchedRootMapKeysResult = {
  touched: Map<string, Set<string>>;
  unknownReason?: string;
};

export type InspectUpdateParams = {
  ydoc: unknown;
  update: Uint8Array;
  reservedRootNames: ReadonlySet<string>;
  reservedRootPrefixes: readonly string[];
  /**
   * Maximum number of touches to report before early-exiting.
   * Defaults to `1` (the inspector is typically used as a boolean guard).
   */
  maxTouches?: number;
};

type IDLike = { client: number; clock: number };

type StoreLike = {
  clients?: Map<number, unknown[]>;
  pendingStructs?: unknown;
  pendingDs?: unknown;
};

type DocLike = {
  store?: StoreLike;
  share?: Map<string, unknown>;
};

type ItemStructLike = {
  id: IDLike;
  length: number;
  parent: unknown;
  parentSub: unknown;
  origin: unknown;
  rightOrigin: unknown;
};

type AbstractTypeLike = {
  _item: unknown;
  doc?: unknown;
};

type StructIndex = Map<number, unknown[]>;

const MAX_RESOLUTION_DEPTH = 64;

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null;
}

function isIdLike(value: unknown): value is IDLike {
  return (
    isRecord(value) &&
    typeof (value as Record<string, unknown>).client === "number" &&
    typeof (value as Record<string, unknown>).clock === "number"
  );
}

function isAbstractTypeLike(value: unknown): value is AbstractTypeLike {
  return isRecord(value) && Object.prototype.hasOwnProperty.call(value, "_item");
}

function isItemStructLike(value: unknown): value is ItemStructLike {
  if (!isRecord(value)) return false;
  const v = value as Record<string, unknown>;
  return (
    isIdLike(v.id) &&
    typeof v.length === "number" &&
    Object.prototype.hasOwnProperty.call(v, "parent") &&
    Object.prototype.hasOwnProperty.call(v, "parentSub") &&
    Object.prototype.hasOwnProperty.call(v, "origin") &&
    Object.prototype.hasOwnProperty.call(v, "rightOrigin")
  );
}

function isGcStruct(value: unknown): boolean {
  // Avoid relying on constructor names (bundlers can rename `GC`).
  return value instanceof (Y as any).GC;
}

function safeStructLen(struct: unknown): number | null {
  if (!isRecord(struct)) return null;
  const len = (struct as any).length;
  return typeof len === "number" ? len : null;
}

function safeStructId(struct: unknown): IDLike | null {
  if (!isRecord(struct)) return null;
  const id = (struct as any).id;
  return isIdLike(id) ? id : null;
}

function structRangeContains(struct: unknown, clock: number): boolean {
  const id = safeStructId(struct);
  const len = safeStructLen(struct);
  if (!id || len === null) return false;
  return id.clock <= clock && clock < id.clock + len;
}

function lowerBoundByClock(structs: readonly unknown[], clock: number): number {
  let left = 0;
  let right = structs.length;
  while (left < right) {
    const mid = (left + right) >> 1;
    const midId = safeStructId(structs[mid]);
    const midClock = midId?.clock;
    if (typeof midClock !== "number") {
      // Malformed struct list. Degrade to a safe linear scan starting at 0.
      return 0;
    }
    if (midClock < clock) left = mid + 1;
    else right = mid;
  }
  return left;
}

function findStructInSortedArray(structs: readonly unknown[], id: IDLike): unknown | null {
  if (structs.length === 0) return null;
  let left = 0;
  let right = structs.length - 1;
  while (left <= right) {
    const mid = (left + right) >> 1;
    const midStruct = structs[mid];
    const midId = safeStructId(midStruct);
    const midLen = safeStructLen(midStruct);
    if (!midId || midLen === null) return null;
    if (id.clock < midId.clock) {
      right = mid - 1;
    } else if (id.clock >= midId.clock + midLen) {
      left = mid + 1;
    } else {
      return midStruct;
    }
  }
  return null;
}

function buildStructIndex(structs: readonly unknown[]): StructIndex {
  const byClient: StructIndex = new Map();
  for (const s of structs) {
    const id = safeStructId(s);
    if (!id) continue;
    const arr = byClient.get(id.client);
    if (arr) arr.push(s);
    else byClient.set(id.client, [s]);
  }
  for (const [client, arr] of byClient) {
    // Sort by clock to enable binary search. Copy first so we don't mutate caller's array.
    const sorted = arr.slice().sort((a, b) => {
      const aId = safeStructId(a);
      const bId = safeStructId(b);
      return (aId?.clock ?? 0) - (bId?.clock ?? 0);
    });
    byClient.set(client, sorted);
  }
  return byClient;
}

function resolveStructById(params: {
  id: IDLike;
  decodedIndex: StructIndex;
  store: StoreLike | undefined;
}): unknown | null {
  const { id, decodedIndex, store } = params;
  const decoded = decodedIndex.get(id.client);
  if (decoded) {
    const s = findStructInSortedArray(decoded, id);
    if (s) return s;
  }

  const storeClients = store?.clients;
  if (storeClients) {
    const structs = storeClients.get(id.client);
    if (structs) {
      // StructStore clients arrays are already sorted by clock, and ranges are contiguous.
      return findStructInSortedArray(structs, id);
    }
  }
  return null;
}

type ParentResolution =
  | { ok: true; parent: unknown; parentSub: string | null }
  | { ok: false; kind: "gc" | "unknown"; reason: string };

function resolveEffectiveParentInfo(
  item: unknown,
  ctx: { decodedIndex: StructIndex; store: StoreLike | undefined },
  opts?: { depth?: number; visited?: Set<unknown> }
): ParentResolution {
  const depth = opts?.depth ?? 0;
  if (depth > MAX_RESOLUTION_DEPTH) {
    return { ok: false, kind: "unknown", reason: "max_parent_resolution_depth" };
  }
  if (!isItemStructLike(item)) {
    return { ok: false, kind: "unknown", reason: "not_item_struct" };
  }

  const parent = item.parent;
  const parentSub = typeof item.parentSub === "string" ? item.parentSub : null;
  if (parent !== null && parent !== undefined) {
    return { ok: true, parent, parentSub };
  }

  const visited = opts?.visited ?? new Set<unknown>();
  if (visited.has(item)) {
    return { ok: false, kind: "unknown", reason: "cycle_in_parent_resolution" };
  }
  visited.add(item);

  // Mirror the "copy parent info from origin/rightOrigin" behavior from Item.getMissing.
  const originId = isIdLike(item.origin) ? item.origin : null;
  const rightOriginId = isIdLike(item.rightOrigin) ? item.rightOrigin : null;

  const leftStruct =
    originId !== null ? resolveStructById({ id: originId, decodedIndex: ctx.decodedIndex, store: ctx.store }) : null;
  const rightStruct =
    rightOriginId !== null
      ? resolveStructById({ id: rightOriginId, decodedIndex: ctx.decodedIndex, store: ctx.store })
      : null;

  if ((leftStruct && isGcStruct(leftStruct)) || (rightStruct && isGcStruct(rightStruct))) {
    return { ok: false, kind: "gc", reason: "origin_or_right_origin_is_gc" };
  }

  if (leftStruct && isItemStructLike(leftStruct)) {
    return resolveEffectiveParentInfo(leftStruct, ctx, { depth: depth + 1, visited });
  }
  if (rightStruct && isItemStructLike(rightStruct)) {
    return resolveEffectiveParentInfo(rightStruct, ctx, { depth: depth + 1, visited });
  }

  return { ok: false, kind: "unknown", reason: "unable_to_resolve_parent_info" };
}

type PathResolution =
  | { ok: true; root: string; keyPath: string[] }
  | { ok: false; kind: "gc" | "unknown"; reason: string };

type RootResolution =
  | { ok: true; root: string }
  | { ok: false; kind: "gc" | "unknown"; reason: string };

function computePathForType(
  type: AbstractTypeLike,
  opts: { depth?: number; visitedTypes?: Set<unknown> } = {}
): PathResolution {
  const depth = opts.depth ?? 0;
  if (depth > MAX_RESOLUTION_DEPTH) {
    return { ok: false, kind: "unknown", reason: "max_type_walk_depth" };
  }
  const keys: string[] = [];
  let current: unknown = type;
  const visitedTypes = opts.visitedTypes ?? new Set<unknown>();
  for (let i = 0; i < MAX_RESOLUTION_DEPTH; i += 1) {
    if (!isAbstractTypeLike(current)) {
      return { ok: false, kind: "unknown", reason: "type_walk_non_abstract_type" };
    }
    if (visitedTypes.has(current)) {
      return { ok: false, kind: "unknown", reason: "cycle_in_type_walk" };
    }
    visitedTypes.add(current);
    const insertionItem = (current as any)._item;
    if (!insertionItem) break;

    const sub = (insertionItem as any).parentSub;
    if (typeof sub === "string") {
      // We walk from inner -> outer, so collect in reverse.
      keys.unshift(sub);
    }
    current = (insertionItem as any).parent;
    if (current == null) {
      return { ok: false, kind: "unknown", reason: "type_walk_missing_parent" };
    }
  }

  if (!isAbstractTypeLike(current)) {
    return { ok: false, kind: "unknown", reason: "type_walk_invalid_root_type" };
  }

  let root: string | null = null;
  try {
    // `Y.findRootTypeKey` throws on failure.
    root = (Y as any).findRootTypeKey(current);
  } catch {
    root = null;
  }

  if (typeof root !== "string") {
    return { ok: false, kind: "unknown", reason: "unable_to_find_root_type_key" };
  }

  return { ok: true, root, keyPath: keys };
}

function computeRootForType(type: AbstractTypeLike): RootResolution {
  let current: unknown = type;

  for (let i = 0; i < MAX_RESOLUTION_DEPTH; i += 1) {
    if (!isAbstractTypeLike(current)) {
      return { ok: false, kind: "unknown", reason: "type_walk_non_abstract_type" };
    }
    const insertionItem = (current as any)._item;
    if (!insertionItem) {
      // `current` is a root type.
      let root: string | null = null;
      try {
        root = (Y as any).findRootTypeKey(current);
      } catch {
        root = null;
      }
      if (typeof root !== "string") {
        return { ok: false, kind: "unknown", reason: "unable_to_find_root_type_key" };
      }
      return { ok: true, root };
    }

    current = (insertionItem as any).parent;
    if (current == null) {
      return { ok: false, kind: "unknown", reason: "type_walk_missing_parent" };
    }
  }

  // We didn't reach a root type within our depth budget, so we can't confidently resolve the root.
  return { ok: false, kind: "unknown", reason: "max_type_walk_depth" };
}

function computeRootForItem(
  item: unknown,
  ctx: { decodedIndex: StructIndex; store: StoreLike | undefined },
  depth: number = 0
): RootResolution {
  if (depth > MAX_RESOLUTION_DEPTH) {
    return { ok: false, kind: "unknown", reason: "max_item_walk_depth" };
  }
  if (!isItemStructLike(item)) {
    return { ok: false, kind: "unknown", reason: "not_item_struct" };
  }

  let parent: unknown = item.parent;
  if (parent === null || parent === undefined) {
    const parentRes = resolveEffectiveParentInfo(item, ctx);
    if (!parentRes.ok) return parentRes;
    parent = parentRes.parent;
  }

  if (typeof parent === "string") {
    if (parent.length === 0) {
      return { ok: false, kind: "unknown", reason: "missing_container_root" };
    }
    return { ok: true, root: parent };
  }

  if (isIdLike(parent)) {
    const insertionStruct = resolveStructById({ id: parent, decodedIndex: ctx.decodedIndex, store: ctx.store });
    if (!insertionStruct) {
      return { ok: false, kind: "unknown", reason: "parent_id_not_found" };
    }
    if (isGcStruct(insertionStruct)) {
      return { ok: false, kind: "gc", reason: "parent_id_points_to_gc" };
    }
    return computeRootForItem(insertionStruct, ctx, depth + 1);
  }

  if (isAbstractTypeLike(parent)) {
    return computeRootForType(parent);
  }

  return { ok: false, kind: "unknown", reason: "unsupported_parent_type" };
}

function computeRootAndKeyPathForItem(
  item: unknown,
  ctx: { decodedIndex: StructIndex; store: StoreLike | undefined }
): PathResolution {
  return computeRootAndKeyPathForItemInner(item, ctx, { depth: 0, visitedItems: new Set<unknown>() });
}

function computeRootAndKeyPathForItemInner(
  item: unknown,
  ctx: { decodedIndex: StructIndex; store: StoreLike | undefined },
  opts: { depth: number; visitedItems: Set<unknown> }
): PathResolution {
  if (opts.depth > MAX_RESOLUTION_DEPTH) {
    return { ok: false, kind: "unknown", reason: "max_item_walk_depth" };
  }
  if (!isItemStructLike(item)) {
    return { ok: false, kind: "unknown", reason: "not_item_struct" };
  }
  if (opts.visitedItems.has(item)) {
    return { ok: false, kind: "unknown", reason: "cycle_in_item_walk" };
  }
  opts.visitedItems.add(item);

  const parentRes = resolveEffectiveParentInfo(item, ctx);
  if (!parentRes.ok) return parentRes;

  const leafKey = typeof parentRes.parentSub === "string" ? parentRes.parentSub : null;
  const parent = parentRes.parent;

  // Determine the container path (root + keys to the parent type).
  let containerRoot: string | null = null;
  let containerKeys: string[] = [];

  if (typeof parent === "string") {
    containerRoot = parent;
  } else if (isIdLike(parent)) {
    // Parent references the insertion item of the parent type.
    const insertionStruct = resolveStructById({ id: parent, decodedIndex: ctx.decodedIndex, store: ctx.store });
    if (!insertionStruct) {
      return { ok: false, kind: "unknown", reason: "parent_id_not_found" };
    }
    if (isGcStruct(insertionStruct)) {
      return { ok: false, kind: "gc", reason: "parent_id_points_to_gc" };
    }
    const insertionPath = computeRootAndKeyPathForItemInner(insertionStruct, ctx, {
      depth: opts.depth + 1,
      visitedItems: opts.visitedItems,
    });
    if (!insertionPath.ok) return insertionPath;
    containerRoot = insertionPath.root;
    containerKeys = insertionPath.keyPath;
  } else if (isAbstractTypeLike(parent)) {
    const typePath = computePathForType(parent);
    if (!typePath.ok) return typePath;
    containerRoot = typePath.root;
    containerKeys = typePath.keyPath;
  } else {
    return { ok: false, kind: "unknown", reason: "unsupported_parent_type" };
  }

  if (typeof containerRoot !== "string" || containerRoot.length === 0) {
    return { ok: false, kind: "unknown", reason: "missing_container_root" };
  }

  const keyPath = leafKey !== null ? [...containerKeys, leafKey] : containerKeys.slice();
  return { ok: true, root: containerRoot, keyPath };
}

function isReservedRoot(root: string, reservedRootNames: ReadonlySet<string>, reservedRootPrefixes: readonly string[]) {
  if (reservedRootNames.has(root)) return true;
  for (const prefix of reservedRootPrefixes) {
    if (root.startsWith(prefix)) return true;
  }
  return false;
}

function failClosed(reason: string, kind: ReservedRootTouchKind = "unknown"): InspectUpdateResult {
  return {
    touchesReserved: true,
    touches: [{ root: "<unknown>", keyPath: [], kind }],
    unknownReason: reason,
  };
}

function safeDecodeUpdate(update: Uint8Array): { structs: unknown[]; ds: unknown } | null {
  const decodeV1 = (): { structs: unknown[]; ds: unknown } | null => {
    try {
      const decoded = (Y as any).decodeUpdate(update) as unknown;
      if (isRecord(decoded) && Array.isArray((decoded as any).structs)) {
        return { structs: (decoded as any).structs as unknown[], ds: (decoded as any).ds };
      }
    } catch {
      // ignore
    }
    return null;
  };

  const decodeV2 = (): { structs: unknown[]; ds: unknown } | null => {
    try {
      const decoded = (Y as any).decodeUpdateV2(update) as unknown;
      if (isRecord(decoded) && Array.isArray((decoded as any).structs)) {
        return { structs: (decoded as any).structs as unknown[], ds: (decoded as any).ds };
      }
    } catch {
      // ignore
    }
    return null;
  };

  const dsClientCount = (ds: unknown): number | null => {
    if (!isRecord(ds)) return null;
    const clients = (ds as any).clients;
    if (!(clients instanceof Map)) return null;
    return clients.size;
  };

  const v1 = decodeV1();
  if (v1) {
    // Yjs v1 decoding can successfully return a *no-op* (0 structs, empty delete set)
    // when given a v2 update (e.g. `encodeStateAsUpdateV2`). Detect this case and
    // fall back to v2 decoding.
    const v1DsSize = dsClientCount(v1.ds);
    if (v1.structs.length === 0 && v1DsSize === 0) {
      const v2 = decodeV2();
      if (v2) {
        const v2DsSize = dsClientCount(v2.ds);
        const v2HasContent = v2.structs.length > 0 || (v2DsSize !== null && v2DsSize > 0);
        if (v2HasContent) return v2;
      }
    }
    return v1;
  }

  return decodeV2();
}

/**
 * Optimized reserved-root inspection intended for per-message websocket guards.
 *
 * This avoids building `keyPath` arrays for every struct in the common case where the
 * update does *not* touch a reserved root. If a reserved root touch is detected, we
 * compute the full `keyPath` only for that first touch (sufficient for logging).
 */
export function inspectUpdateForReservedRootGuard(params: InspectUpdateParams): InspectUpdateResult {
  const maxTouches = params.maxTouches ?? 1;
  const touches: ReservedRootTouch[] = [];

  const ydoc = params.ydoc as DocLike;
  const store = ydoc?.store;
  if (!store || store.pendingStructs || store.pendingDs) {
    return failClosed("ydoc_store_pending");
  }

  const decoded = safeDecodeUpdate(params.update);
  if (!decoded) {
    return failClosed("decode_failed");
  }

  const decodedIndex = buildStructIndex(decoded.structs);

  const maybeReserved = (root: string) =>
    isReservedRoot(root, params.reservedRootNames, params.reservedRootPrefixes);

  const recordFirstTouchForStruct = (struct: unknown, kind: ReservedRootTouchKind) => {
    const pathRes = computeRootAndKeyPathForItem(struct, { decodedIndex, store });
    if (!pathRes.ok) {
      return failClosed(pathRes.reason, pathRes.kind);
    }
    touches.push({ root: pathRes.root, keyPath: pathRes.keyPath, kind });
    return touches.length >= maxTouches ? { touchesReserved: true, touches } : null;
  };

  // Inspect newly inserted structs.
  for (const struct of decoded.structs) {
    if (!isItemStructLike(struct)) continue;
    const rootRes = computeRootForItem(struct, { decodedIndex, store });
    if (!rootRes.ok) {
      // If we can't confidently inspect, fail closed.
      return failClosed(rootRes.reason, rootRes.kind);
    }
    if (maybeReserved(rootRes.root)) {
      const res = recordFirstTouchForStruct(struct, "insert");
      if (res) return res;
    }
  }

  // Inspect delete set ranges using the server doc store.
  const ds = decoded.ds as any;
  const dsClients: Map<number, unknown[]> | null =
    ds && ds.clients instanceof Map ? (ds.clients as Map<number, unknown[]>) : null;

  if (dsClients) {
    for (const [client, deletes] of dsClients.entries()) {
      if (!Array.isArray(deletes) || deletes.length === 0) continue;
      const storeStructs = store.clients?.get(client) ?? null;
      const decodedStructs = decodedIndex.get(client) ?? null;
      if (!storeStructs && !decodedStructs) {
        return failClosed("delete_set_client_missing_in_store_and_update");
      }

      const storeState = (() => {
        if (!storeStructs || storeStructs.length === 0) return 0;
        const last = storeStructs[storeStructs.length - 1];
        const lastId = safeStructId(last);
        const lastLen = safeStructLen(last);
        if (!lastId || lastLen === null) return null;
        return lastId.clock + lastLen;
      })();

      const decodedState = (() => {
        if (!decodedStructs || decodedStructs.length === 0) return 0;
        const last = decodedStructs[decodedStructs.length - 1];
        const lastId = safeStructId(last);
        const lastLen = safeStructLen(last);
        if (!lastId || lastLen === null) return null;
        return lastId.clock + lastLen;
      })();

      if (storeState === null || decodedState === null) {
        return failClosed("malformed_store_or_update_struct");
      }

      const clientState = Math.max(storeState, decodedState);

      for (const del of deletes) {
        const clock = typeof (del as any)?.clock === "number" ? (del as any).clock : null;
        const len = typeof (del as any)?.len === "number" ? (del as any).len : null;
        if (clock === null || len === null) {
          return failClosed("malformed_delete_set");
        }
        if (len <= 0) continue;

        const endClock = clock + len;
        if (clock < 0 || endClock > clientState) {
          // We can't resolve the full delete range against the current store state
          // without splitting or pending updates. Fail closed.
          return failClosed("delete_set_range_out_of_bounds");
        }

        const processStructArray = (structs: readonly unknown[]) => {
          // Find first struct that overlaps [clock, endClock)
          let index = lowerBoundByClock(structs, clock);
          if (index > 0 && structRangeContains(structs[index - 1], clock)) {
            index -= 1;
          }
          for (let i = index; i < structs.length; i += 1) {
            const s = structs[i];
            const sId = safeStructId(s);
            const sLen = safeStructLen(s);
            if (!sId || sLen === null) {
              throw new Error("malformed_struct");
            }
            if (sId.clock >= endClock) break;
            // s overlaps [clock, endClock)
            if (!isItemStructLike(s)) continue;
            const rootRes = computeRootForItem(s, { decodedIndex, store });
            if (!rootRes.ok) {
              throw new Error(`${rootRes.kind}:${rootRes.reason}`);
            }
            if (maybeReserved(rootRes.root)) {
              const pathRes = computeRootAndKeyPathForItem(s, { decodedIndex, store });
              if (!pathRes.ok) {
                throw new Error(`${pathRes.kind}:${pathRes.reason}`);
              }
              touches.push({ root: pathRes.root, keyPath: pathRes.keyPath, kind: "delete" });
              if (touches.length >= maxTouches) return;
            }
          }
        };

        // Ensure we can resolve the starting point in either store structs or update structs.
        const startId: IDLike = { client, clock };
        const startInStore = storeStructs ? findStructInSortedArray(storeStructs, startId) : null;
        const startInUpdate = decodedStructs ? findStructInSortedArray(decodedStructs, startId) : null;
        if (!startInStore && !startInUpdate) {
          return failClosed("delete_set_start_not_found");
        }

        try {
          if (storeStructs) processStructArray(storeStructs);
          if (touches.length >= maxTouches) {
            return { touchesReserved: true, touches: touches.slice(0, maxTouches) };
          }
          if (decodedStructs) processStructArray(decodedStructs);
        } catch (err) {
          const msg = err instanceof Error ? err.message : String(err);
          if (msg.startsWith("gc:")) return failClosed(msg.slice(3), "gc");
          if (msg.startsWith("unknown:")) return failClosed(msg.slice(8), "unknown");
          return failClosed("malformed_store_struct");
        }

        if (touches.length >= maxTouches) {
          return { touchesReserved: true, touches: touches.slice(0, maxTouches) };
        }
      }
    }
  }

  return { touchesReserved: touches.length > 0, touches };
}

export type InspectUpdateAllowedRootsResult =
  | { allowed: true }
  | {
      allowed: false;
      touch: ReservedRootTouch;
      /**
       * Present when we failed to confidently inspect the update.
       *
       * Callers should treat this as fail-closed.
       */
      unknownReason?: string;
    };

/**
 * Best-effort inspection to enforce that an update only touches an allowlist of root names.
 *
 * Intended for coarse-grained role enforcement (e.g. allowing a `commenter` role to mutate
 * the `comments` root while rejecting all workbook mutations).
 *
 * Like the reserved-root guard, this is designed to be safe for websocket message guards:
 * - It avoids constructing full key paths for every struct in the common case where updates
 *   are allowed.
 * - When a disallowed root touch is detected, it computes the full key path only for that
 *   first violation (sufficient for logging).
 */
export function inspectUpdateForAllowedRoots(params: {
  ydoc: unknown;
  update: Uint8Array;
  allowedRoots: ReadonlySet<string>;
}): InspectUpdateAllowedRootsResult {
  const ydoc = params.ydoc as DocLike;
  const store = ydoc?.store;
  if (!store || store.pendingStructs || store.pendingDs) {
    return {
      allowed: false,
      touch: { root: "<unknown>", keyPath: [], kind: "unknown" },
      unknownReason: "ydoc_store_pending",
    };
  }

  const decoded = safeDecodeUpdate(params.update);
  if (!decoded) {
    return {
      allowed: false,
      touch: { root: "<unknown>", keyPath: [], kind: "unknown" },
      unknownReason: "decode_failed",
    };
  }

  const decodedIndex = buildStructIndex(decoded.structs);

  const isAllowed = (root: string) => params.allowedRoots.has(root);

  const recordViolation = (struct: unknown, kind: ReservedRootTouchKind): InspectUpdateAllowedRootsResult => {
    const pathRes = computeRootAndKeyPathForItem(struct, { decodedIndex, store });
    if (!pathRes.ok) {
      return {
        allowed: false,
        touch: { root: "<unknown>", keyPath: [], kind: pathRes.kind },
        unknownReason: pathRes.reason,
      };
    }
    return { allowed: false, touch: { root: pathRes.root, keyPath: pathRes.keyPath, kind } };
  };

  // Inspect newly inserted structs.
  for (const struct of decoded.structs) {
    if (!isItemStructLike(struct)) continue;
    const rootRes = computeRootForItem(struct, { decodedIndex, store });
    if (!rootRes.ok) {
      return {
        allowed: false,
        touch: { root: "<unknown>", keyPath: [], kind: rootRes.kind },
        unknownReason: rootRes.reason,
      };
    }
    if (!isAllowed(rootRes.root)) {
      return recordViolation(struct, "insert");
    }
  }

  // Inspect delete set ranges using the server doc store.
  const ds = decoded.ds as any;
  const dsClients: Map<number, unknown[]> | null =
    ds && ds.clients instanceof Map ? (ds.clients as Map<number, unknown[]>) : null;

  if (dsClients) {
    for (const [client, deletes] of dsClients.entries()) {
      if (!Array.isArray(deletes) || deletes.length === 0) continue;
      const storeStructs = store.clients?.get(client) ?? null;
      const decodedStructs = decodedIndex.get(client) ?? null;
      if (!storeStructs && !decodedStructs) {
        return {
          allowed: false,
          touch: { root: "<unknown>", keyPath: [], kind: "unknown" },
          unknownReason: "delete_set_client_missing_in_store_and_update",
        };
      }

      const storeState = (() => {
        if (!storeStructs || storeStructs.length === 0) return 0;
        const last = storeStructs[storeStructs.length - 1];
        const lastId = safeStructId(last);
        const lastLen = safeStructLen(last);
        if (!lastId || lastLen === null) return null;
        return lastId.clock + lastLen;
      })();

      const decodedState = (() => {
        if (!decodedStructs || decodedStructs.length === 0) return 0;
        const last = decodedStructs[decodedStructs.length - 1];
        const lastId = safeStructId(last);
        const lastLen = safeStructLen(last);
        if (!lastId || lastLen === null) return null;
        return lastId.clock + lastLen;
      })();

      if (storeState === null || decodedState === null) {
        return {
          allowed: false,
          touch: { root: "<unknown>", keyPath: [], kind: "unknown" },
          unknownReason: "malformed_store_or_update_struct",
        };
      }

      const clientState = Math.max(storeState, decodedState);

      for (const del of deletes) {
        const clock = typeof (del as any)?.clock === "number" ? (del as any).clock : null;
        const len = typeof (del as any)?.len === "number" ? (del as any).len : null;
        if (clock === null || len === null) {
          return {
            allowed: false,
            touch: { root: "<unknown>", keyPath: [], kind: "unknown" },
            unknownReason: "malformed_delete_set",
          };
        }
        if (len <= 0) continue;

        const endClock = clock + len;
        if (clock < 0 || endClock > clientState) {
          return {
            allowed: false,
            touch: { root: "<unknown>", keyPath: [], kind: "unknown" },
            unknownReason: "delete_set_range_out_of_bounds",
          };
        }

        let violation: InspectUpdateAllowedRootsResult | null = null;

        const processStructArray = (
          structs: readonly unknown[]
        ): InspectUpdateAllowedRootsResult | null => {
          let index = lowerBoundByClock(structs, clock);
          if (index > 0 && structRangeContains(structs[index - 1], clock)) {
            index -= 1;
          }
          for (let i = index; i < structs.length; i += 1) {
            const s = structs[i];
            const sId = safeStructId(s);
            const sLen = safeStructLen(s);
            if (!sId || sLen === null) {
              throw new Error("malformed_struct");
            }
            if (sId.clock >= endClock) break;
            if (!isItemStructLike(s)) continue;

            const rootRes = computeRootForItem(s, { decodedIndex, store });
            if (!rootRes.ok) {
              throw new Error(`${rootRes.kind}:${rootRes.reason}`);
            }
            if (!isAllowed(rootRes.root)) {
              return recordViolation(s, "delete");
            }
          }
          return null;
        };

        const startId: IDLike = { client, clock };
        const startInStore = storeStructs ? findStructInSortedArray(storeStructs, startId) : null;
        const startInUpdate = decodedStructs ? findStructInSortedArray(decodedStructs, startId) : null;
        if (!startInStore && !startInUpdate) {
          return {
            allowed: false,
            touch: { root: "<unknown>", keyPath: [], kind: "unknown" },
            unknownReason: "delete_set_start_not_found",
          };
        }

        try {
          if (storeStructs) violation = processStructArray(storeStructs);
          if (!violation && decodedStructs) violation = processStructArray(decodedStructs);
        } catch (err) {
          const msg = err instanceof Error ? err.message : String(err);
          if (msg.startsWith("gc:")) {
            return {
              allowed: false,
              touch: { root: "<unknown>", keyPath: [], kind: "gc" },
              unknownReason: msg.slice(3),
            };
          }
          if (msg.startsWith("unknown:")) {
            return {
              allowed: false,
              touch: { root: "<unknown>", keyPath: [], kind: "unknown" },
              unknownReason: msg.slice(8),
            };
          }
          return {
            allowed: false,
            touch: { root: "<unknown>", keyPath: [], kind: "unknown" },
            unknownReason: "malformed_store_struct",
          };
        }

        if (violation && violation.allowed === false) {
          return violation;
        }
      }
    }
  }

  return { allowed: true };
}

export function inspectUpdate(params: InspectUpdateParams): InspectUpdateResult {
  const maxTouches = params.maxTouches ?? 1;
  const touches: ReservedRootTouch[] = [];

  const ydoc = params.ydoc as DocLike;
  const store = ydoc?.store;
  if (!store || store.pendingStructs || store.pendingDs) {
    return failClosed("ydoc_store_pending");
  }

  const decoded = safeDecodeUpdate(params.update);
  if (!decoded) {
    return failClosed("decode_failed");
  }

  const decodedIndex = buildStructIndex(decoded.structs);

  const recordTouch = (touch: ReservedRootTouch) => {
    touches.push(touch);
  };

  const maybeReserved = (root: string) =>
    isReservedRoot(root, params.reservedRootNames, params.reservedRootPrefixes);

  // Inspect newly inserted structs.
  for (const struct of decoded.structs) {
    if (!isItemStructLike(struct)) continue;
    const pathRes = computeRootAndKeyPathForItem(struct, { decodedIndex, store });
    if (!pathRes.ok) {
      // If we can't confidently inspect, fail closed.
      return failClosed(pathRes.reason, pathRes.kind);
    }
    if (maybeReserved(pathRes.root)) {
      recordTouch({ root: pathRes.root, keyPath: pathRes.keyPath, kind: "insert" });
      if (touches.length >= maxTouches) {
        return { touchesReserved: true, touches };
      }
    }
  }

  // Inspect delete set ranges using the server doc store.
  const ds = decoded.ds as any;
  const dsClients: Map<number, unknown[]> | null =
    ds && ds.clients instanceof Map ? (ds.clients as Map<number, unknown[]>) : null;

  if (dsClients) {
    for (const [client, deletes] of dsClients.entries()) {
      if (!Array.isArray(deletes) || deletes.length === 0) continue;
      const storeStructs = store.clients?.get(client) ?? null;
      const decodedStructs = decodedIndex.get(client) ?? null;
      if (!storeStructs && !decodedStructs) {
        return failClosed("delete_set_client_missing_in_store_and_update");
      }

      const storeState = (() => {
        if (!storeStructs || storeStructs.length === 0) return 0;
        const last = storeStructs[storeStructs.length - 1];
        const lastId = safeStructId(last);
        const lastLen = safeStructLen(last);
        if (!lastId || lastLen === null) return null;
        return lastId.clock + lastLen;
      })();

      const decodedState = (() => {
        if (!decodedStructs || decodedStructs.length === 0) return 0;
        const last = decodedStructs[decodedStructs.length - 1];
        const lastId = safeStructId(last);
        const lastLen = safeStructLen(last);
        if (!lastId || lastLen === null) return null;
        return lastId.clock + lastLen;
      })();

      if (storeState === null || decodedState === null) {
        return failClosed("malformed_store_or_update_struct");
      }

      const clientState = Math.max(storeState, decodedState);

      for (const del of deletes) {
        const clock = typeof (del as any)?.clock === "number" ? (del as any).clock : null;
        const len = typeof (del as any)?.len === "number" ? (del as any).len : null;
        if (clock === null || len === null) {
          return failClosed("malformed_delete_set");
        }
        if (len <= 0) continue;

        const endClock = clock + len;
        if (clock < 0 || endClock > clientState) {
          // We can't resolve the full delete range against the current store state
          // without splitting or pending updates. Fail closed.
          return failClosed("delete_set_range_out_of_bounds");
        }

        const processStructArray = (structs: readonly unknown[]) => {
          // Find first struct that overlaps [clock, endClock)
          let index = lowerBoundByClock(structs, clock);
          if (index > 0 && structRangeContains(structs[index - 1], clock)) {
            index -= 1;
          }
          for (let i = index; i < structs.length; i += 1) {
            const s = structs[i];
            const sId = safeStructId(s);
            const sLen = safeStructLen(s);
            if (!sId || sLen === null) {
              throw new Error("malformed_struct");
            }
            if (sId.clock >= endClock) break;
            // s overlaps [clock, endClock)
            if (!isItemStructLike(s)) continue;
            const pathRes = computeRootAndKeyPathForItem(s, { decodedIndex, store });
            if (!pathRes.ok) {
              throw new Error(`${pathRes.kind}:${pathRes.reason}`);
            }
            if (maybeReserved(pathRes.root)) {
              recordTouch({ root: pathRes.root, keyPath: pathRes.keyPath, kind: "delete" });
            }
          }
        };

        // Ensure we can resolve the starting point in either store structs or update structs.
        const startId: IDLike = { client, clock };
        const startInStore = storeStructs ? findStructInSortedArray(storeStructs, startId) : null;
        const startInUpdate = decodedStructs ? findStructInSortedArray(decodedStructs, startId) : null;
        if (!startInStore && !startInUpdate) {
          return failClosed("delete_set_start_not_found");
        }

        try {
          if (storeStructs) processStructArray(storeStructs);
          if (decodedStructs) processStructArray(decodedStructs);
        } catch (err) {
          const msg = err instanceof Error ? err.message : String(err);
          if (msg.startsWith("gc:")) return failClosed(msg.slice(3), "gc");
          if (msg.startsWith("unknown:")) return failClosed(msg.slice(8), "unknown");
          return failClosed("malformed_store_struct");
        }

        if (touches.length >= maxTouches) {
          return { touchesReserved: true, touches: touches.slice(0, maxTouches) };
        }
      }
    }
  }

  return { touchesReserved: touches.length > 0, touches };
}

/**
 * Collects *direct* root Y.Map key touches for the specified root names.
 *
 * This is primarily useful for server-side quota enforcement where we only need to
 * know which keys were set/updated on a top-level root map (e.g. `versions.set(id, ...)`),
 * and we want to support incremental updates (client clock > 0) without applying the
 * update to a temporary document.
 */
export function collectTouchedRootMapKeys(params: {
  ydoc: unknown;
  update: Uint8Array;
  rootNames: readonly string[];
}): CollectTouchedRootMapKeysResult {
  const touched = new Map<string, Set<string>>();
  for (const rootName of params.rootNames) {
    touched.set(rootName, new Set());
  }

  const ydoc = params.ydoc as DocLike;
  const store = ydoc?.store;
  if (!store || store.pendingStructs || store.pendingDs) {
    return { touched, unknownReason: "ydoc_store_pending" };
  }

  const decoded = safeDecodeUpdate(params.update);
  if (!decoded) {
    return { touched, unknownReason: "decode_failed" };
  }

  const decodedIndex = buildStructIndex(decoded.structs);

  for (const struct of decoded.structs) {
    if (!isItemStructLike(struct)) continue;
    const parentRes = resolveEffectiveParentInfo(struct, { decodedIndex, store });
    if (!parentRes.ok) {
      return { touched, unknownReason: parentRes.reason };
    }
    if (typeof parentRes.parent !== "string") continue;
    const rootSet = touched.get(parentRes.parent);
    if (!rootSet) continue;
    if (typeof parentRes.parentSub === "string" && parentRes.parentSub.length > 0) {
      rootSet.add(parentRes.parentSub);
    }
  }

  return { touched };
}
