import { stableJsonStringify } from "../../../../../packages/ai-context/src/tokenBudget.js";

function safeList<T>(fn: () => T[] | null | undefined): T[] {
  try {
    const out = fn();
    return Array.isArray(out) ? out : [];
  } catch {
    return [];
  }
}

function safeStableJsonStringify(value: unknown): string {
  try {
    return stableJsonStringify(value);
  } catch {
    try {
      return JSON.stringify(value) ?? "";
    } catch {
      return "";
    }
  }
}

function fnv1a32(value: string): number {
  // 32-bit FNV-1a hash. (Stable across runs.)
  let hash = 0x811c9dc5;
  for (let i = 0; i < value.length; i++) {
    hash ^= value.charCodeAt(i);
    hash = Math.imul(hash, 0x01000193);
  }
  return hash >>> 0;
}

function fnv1a32Update(hash: number, value: string): number {
  // Incremental 32-bit FNV-1a hash update.
  let out = hash >>> 0;
  for (let i = 0; i < value.length; i++) {
    out ^= value.charCodeAt(i);
    out = Math.imul(out, 0x01000193);
  }
  return out >>> 0;
}

function hashString(value: string): string {
  return fnv1a32(value).toString(16);
}

function normalizeLabels(labelsRaw: unknown): string[] {
  const labels = Array.isArray(labelsRaw) ? labelsRaw : [];
  const normalized = labels
    .map((l) => {
      try {
        return String(l);
      } catch {
        return "";
      }
    })
    .map((l) => l.trim())
    .filter(Boolean);
  return Array.from(new Set(normalized)).sort((a, b) => a.localeCompare(b));
}

function normalizeClassificationForCacheKey(value: unknown): unknown {
  if (!value || typeof value !== "object") return value ?? null;
  const obj = value as any;
  return {
    level: typeof obj.level === "string" ? obj.level : null,
    labels: normalizeLabels(obj.labels),
  };
}

function normalizeNonNegativeInt(value: unknown): number | null {
  if (typeof value !== "number" || !Number.isFinite(value)) return null;
  if (!Number.isInteger(value) || value < 0) return null;
  return value;
}

function normalizeSelectorForCacheKey(value: unknown): unknown {
  if (!value || typeof value !== "object") return value ?? null;
  const selector = value as any;
  const scope = typeof selector.scope === "string" ? selector.scope : "";
  const documentId = typeof selector.documentId === "string" ? selector.documentId : null;

  // When we can't recognize the selector shape, fall back to a stable JSON form so
  // the cache key remains sensitive to changes.
  const fallback = () => safeStableJsonStringify(value);

  switch (scope) {
    case "document":
      return { scope, documentId };
    case "sheet":
      return {
        scope,
        documentId,
        sheetId: typeof selector.sheetId === "string" ? selector.sheetId : null,
      };
    case "column": {
      const out: any = {
        scope,
        documentId,
        sheetId: typeof selector.sheetId === "string" ? selector.sheetId : null,
      };
      const columnIndex = normalizeNonNegativeInt(selector.columnIndex);
      if (columnIndex !== null) out.columnIndex = columnIndex;
      if (typeof selector.columnId === "string") out.columnId = selector.columnId;
      if (typeof selector.tableId === "string") out.tableId = selector.tableId;
      return out;
    }
    case "cell": {
      const row = normalizeNonNegativeInt(selector.row);
      const col = normalizeNonNegativeInt(selector.col);
      const out: any = {
        scope,
        documentId,
        sheetId: typeof selector.sheetId === "string" ? selector.sheetId : null,
        row,
        col,
      };
      if (typeof selector.tableId === "string") out.tableId = selector.tableId;
      if (typeof selector.columnId === "string") out.columnId = selector.columnId;
      return out;
    }
    case "range": {
      if (!selector.range || typeof selector.range !== "object") {
        return {
          scope,
          documentId,
          sheetId: typeof selector.sheetId === "string" ? selector.sheetId : null,
          range: null,
        };
      }
      const start = (selector.range as any).start;
      const end = (selector.range as any).end;
      if (!start || typeof start !== "object" || !end || typeof end !== "object") return fallback();
      const startRow = normalizeNonNegativeInt((start as any).row);
      const startCol = normalizeNonNegativeInt((start as any).col);
      const endRow = normalizeNonNegativeInt((end as any).row);
      const endCol = normalizeNonNegativeInt((end as any).col);
      if ([startRow, startCol, endRow, endCol].some((n) => n === null)) return fallback();
      const r0 = Math.min(startRow!, endRow!);
      const r1 = Math.max(startRow!, endRow!);
      const c0 = Math.min(startCol!, endCol!);
      const c1 = Math.max(startCol!, endCol!);
      return {
        scope,
        documentId,
        sheetId: typeof selector.sheetId === "string" ? selector.sheetId : null,
        range: { start: { row: r0, col: c0 }, end: { row: r1, col: c1 } },
      };
    }
    default:
      return fallback();
  }
}

function normalizedClassificationRecordKeysForCacheKey(records: unknown): string[] {
  const list = Array.isArray(records) ? records : [];
  const keys = list.map((record) =>
    safeStableJsonStringify({
      selector: normalizeSelectorForCacheKey((record as any)?.selector ?? null),
      classification: normalizeClassificationForCacheKey((record as any)?.classification ?? null),
    }),
  );
  // Ensure deterministic ordering even if the backing classification store does not
  // guarantee record order.
  keys.sort((a, b) => a.localeCompare(b));
  return keys;
}

function hashJsonArray(keys: string[]): { length: number; hash: string } {
  // Hash the JSON array string `[k1,k2,...]` without materializing a potentially huge
  // intermediate string.
  const length = 2 + keys.reduce((acc, key) => acc + key.length, 0) + Math.max(0, keys.length - 1);
  let hash = 0x811c9dc5;
  hash = fnv1a32Update(hash, "[");
  for (let i = 0; i < keys.length; i++) {
    if (i > 0) hash = fnv1a32Update(hash, ",");
    hash = fnv1a32Update(hash, keys[i]!);
  }
  hash = fnv1a32Update(hash, "]");
  return { length, hash: (hash >>> 0).toString(16) };
}

/**
 * Stable, safe-to-store DLP cache key.
 *
 * This key is designed to be:
 * - deterministic across record ordering
 * - sensitive to policy / classification / includeRestrictedContent changes
 * - cheap to compare (short hash-based string)
 * - safe to log/store (no raw policy JSON embedded)
 */
export function computeDlpCacheKey(dlp: any): string {
  if (!dlp) return "no_dlp";

  const includeRestrictedContent = Boolean(dlp.includeRestrictedContent ?? dlp.include_restricted_content ?? false);

  // Allow callers to pre-compute and attach the key (or reuse a previously computed
  // key) to avoid re-hashing large policies / record sets in hot paths.
  try {
    if (typeof dlp === "object" && dlp !== null) {
      const hasExplicitRecords = Array.isArray((dlp as any).classificationRecords) || Array.isArray((dlp as any).classification_records);
      const hasStore = Boolean((dlp as any).classificationStore || (dlp as any).classification_store);
      const existing = (dlp as any).cacheKey ?? (dlp as any).cache_key;
      if ((hasExplicitRecords || !hasStore) && typeof existing === "string" && existing.trim()) {
        const trimmed = existing.trim();
        // Ensure a cached key cannot be reused if includeRestrictedContent was toggled on the
        // same object (cache keys are prefixed by incl/excl).
        const expectedPrefix = `dlp:${includeRestrictedContent ? "incl" : "excl"}:`;
        if (trimmed.startsWith(expectedPrefix)) return trimmed;
      }
    }
  } catch {
    // ignore
  }

  const policyJson = safeStableJsonStringify(dlp.policy ?? null);
  const policyKey = `${policyJson.length}:${hashString(policyJson)}`;

  // Prefer the explicit record list when available (callers often fetch it once for
  // both ToolExecutor and cache key computation). Fall back to a provided store if
  // the records array is omitted.
  const explicitRecords: Array<any> | null = Array.isArray(dlp.classificationRecords)
    ? dlp.classificationRecords
    : Array.isArray(dlp.classification_records)
      ? dlp.classification_records
      : null;

  const records: Array<any> =
    explicitRecords ??
    safeList(() => {
      const store = dlp.classificationStore ?? dlp.classification_store;
      if (!store || typeof store.list !== "function") return [];
      const docId =
        typeof dlp.documentId === "string"
          ? dlp.documentId
          : typeof dlp.document_id === "string"
            ? dlp.document_id
            : "";
      if (!docId) return [];
      const out = store.list(docId);
      return Array.isArray(out) ? out : [];
    });

  // Cache keys must be sensitive to selector/classification changes; relying only on
  // timestamps like `updatedAt` is unsafe in distributed systems (clock skew) and
  // for callers that omit timestamps entirely.
  const recordKeys = normalizedClassificationRecordKeysForCacheKey(records);
  const hashed = hashJsonArray(recordKeys);
  const recordsKey = `${hashed.length}:${hashed.hash}`;

  const key = `dlp:${includeRestrictedContent ? "incl" : "excl"}:${policyKey}:${recordsKey}`;

  // Memoize when it is safe to do so:
  // - When an explicit snapshot list of classification records is provided, OR
  // - When no classification store is provided (records are effectively immutable/empty).
  //
  // If a caller passes only a dynamic classification store (no explicit records), caching
  // could become unsafe because the store contents can change over time.
  try {
    if (typeof dlp === "object" && dlp !== null) {
      const obj = dlp as any;
      const hasExplicitRecords = Array.isArray(obj.classificationRecords) || Array.isArray(obj.classification_records);
      const hasStore = Boolean(obj.classificationStore || obj.classification_store);
      if (hasExplicitRecords || !hasStore) {
        try {
          Object.defineProperty(obj, "cacheKey", { value: key, enumerable: false, configurable: true });
        } catch {
          try {
            obj.cacheKey = key;
          } catch {
            // ignore
          }
        }
      }
    }
  } catch {
    // ignore
  }

  return key;
}
