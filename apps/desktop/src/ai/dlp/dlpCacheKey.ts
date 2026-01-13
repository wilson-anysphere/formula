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

function hashString(value: string): string {
  return fnv1a32(value).toString(16);
}

function normalizeClassificationRecordsForCacheKey(
  records: unknown,
): Array<{ selector: unknown; classification: unknown }> {
  const list = Array.isArray(records) ? records : [];
  const normalized = list.map((record) => ({
    selector: (record as any)?.selector ?? null,
    classification: (record as any)?.classification ?? null,
  }));

  // Ensure deterministic ordering even if the backing classification store does not
  // guarantee record order.
  const keyed = normalized.map((r) => ({ key: safeStableJsonStringify(r), value: r }));
  keyed.sort((a, b) => a.key.localeCompare(b.key));
  return keyed.map((r) => r.value);
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

  // Allow callers to pre-compute and attach the key (or reuse a previously computed
  // key) to avoid re-hashing large policies / record sets in hot paths.
  try {
    if (typeof dlp === "object" && dlp !== null) {
      const existing = (dlp as any).cacheKey ?? (dlp as any).cache_key;
      if (typeof existing === "string" && existing.trim()) return existing.trim();
    }
  } catch {
    // ignore
  }

  const includeRestrictedContent = Boolean(dlp.includeRestrictedContent ?? dlp.include_restricted_content ?? false);

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
  const normalized = normalizeClassificationRecordsForCacheKey(records);
  const recordsJson = safeStableJsonStringify(normalized);
  const recordsKey = `${recordsJson.length}:${hashString(recordsJson)}`;

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
