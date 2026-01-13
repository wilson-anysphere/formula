import { LocalClassificationStore, createMemoryStorage } from "../../../../../packages/security/dlp/src/classificationStore.js";
import { InMemoryAuditLogger } from "../../../../../packages/security/dlp/src/audit.js";
import { createDefaultOrgPolicy, mergePolicies } from "../../../../../packages/security/dlp/src/policy.js";
import { LocalPolicyStore } from "../../../../../packages/security/dlp/src/policyStore.js";
import type { SheetNameResolver } from "../../sheet/sheetNameResolver.js";
import { computeDlpCacheKey } from "./dlpCacheKey.js";

type StorageLike = { getItem(key: string): string | null; setItem(key: string, value: string): void; removeItem(key: string): void };

const AI_DLP_AUDIT_LOGGER_RETENTION_CAP = 1_000;

class CappedInMemoryAuditLogger extends InMemoryAuditLogger {
  private readonly retentionCap: number;

  constructor(retentionCap: number) {
    super();
    this.retentionCap = retentionCap;
  }

  override log(event: any): string {
    const id = super.log(event);
    const excess = this.events.length - this.retentionCap;
    if (excess > 0) this.events.splice(0, excess);
    return id;
  }
}

function createAiDlpAuditLogger(): InMemoryAuditLogger {
  return new CappedInMemoryAuditLogger(AI_DLP_AUDIT_LOGGER_RETENTION_CAP);
}

let sharedAuditLogger: InMemoryAuditLogger | null = null;

export function getAiDlpAuditLogger(): InMemoryAuditLogger {
  if (!sharedAuditLogger) sharedAuditLogger = createAiDlpAuditLogger();
  return sharedAuditLogger;
}

/**
 * Test helper: resets the shared audit logger instance.
 *
 * Desktop production code can replace this with a real audit pipeline later; the
 * singleton keeps orchestration surfaces consistent and makes unit testing deterministic.
 */
export function resetAiDlpAuditLoggerForTests(): void {
  sharedAuditLogger = createAiDlpAuditLogger();
}

type AiCloudDlpCacheEntry = {
  rawOrgPolicy: string | null;
  rawDocPolicy: string | null;
  rawClassifications: string | null;
  storedOrgPolicy: any | null;
  storedDocumentPolicy: any | null;
  policy: any;
  classificationRecords: Array<{ selector: any; classification: any }>;
  classificationStore: LocalClassificationStore;
  hasDlpConfig: boolean;
};

const aiCloudDlpCache = new Map<string, AiCloudDlpCacheEntry>();

function safeStorage(storage: StorageLike): StorageLike {
  return {
    getItem(key) {
      try {
        return storage.getItem(key);
      } catch {
        return null;
      }
    },
    setItem(key, value) {
      try {
        storage.setItem(key, value);
      } catch {
        // ignore
      }
    },
    removeItem(key) {
      try {
        storage.removeItem(key);
      } catch {
        // ignore
      }
    }
  };
}

function getLocalStorageOrNull(): StorageLike | null {
  if (typeof window !== "undefined") {
    try {
      const storage = window.localStorage as any;
      if (storage && typeof storage.getItem === "function" && typeof storage.setItem === "function") {
        return safeStorage(storage);
      }
    } catch {
      // ignore
    }
  }

  try {
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const storage = (globalThis as any)?.localStorage as any;
    if (storage && typeof storage.getItem === "function" && typeof storage.setItem === "function") {
      return safeStorage(storage);
    }
  } catch {
    // ignore
  }

  return null;
}

function loadActiveOrgId(storage: StorageLike): string {
  // Desktop currently doesn't have a first-class org/session concept. Allow tests or
  // future integrations to configure the active org id via localStorage.
  const candidates = ["dlp:activeOrgId", "formula:activeOrgId", "formula:orgId"];
  for (const key of candidates) {
    const value = storage.getItem(key);
    if (typeof value === "string" && value.trim()) return value.trim();
  }
  return "default";
}

function mergeEffectivePolicy(params: { orgPolicy: any; documentPolicy: any }): any {
  try {
    return mergePolicies({ orgPolicy: params.orgPolicy, documentPolicy: params.documentPolicy }).policy;
  } catch {
    // Corrupt localStorage / invalid policies should never take down AI surfaces.
    // Prefer falling back to the org policy (if valid) before using the hard-coded default.
    try {
      return mergePolicies({ orgPolicy: params.orgPolicy, documentPolicy: undefined }).policy;
    } catch {
      // Safe baseline.
      return createDefaultOrgPolicy();
    }
  }
}

function loadEffectivePolicy(params: { documentId: string; orgId?: string; storage: StorageLike }): any {
  const policyStore = new LocalPolicyStore({ storage: params.storage });

  const orgId = params.orgId ?? loadActiveOrgId(params.storage);

  const storedOrgPolicy = policyStore.getOrgPolicy(orgId);
  const storedDocumentPolicy = policyStore.getDocumentPolicy(params.documentId);

  let orgPolicy: any = createDefaultOrgPolicy();
  if (storedOrgPolicy) orgPolicy = storedOrgPolicy;

  return mergeEffectivePolicy({ orgPolicy, documentPolicy: storedDocumentPolicy });
}

function memoizedDlpData(params: { storage: StorageLike; documentId: string; orgId: string }): AiCloudDlpCacheEntry {
  const cacheKey = `${params.orgId}\u0000${params.documentId}`;

  const rawOrgPolicy = params.storage.getItem(`dlp:orgPolicy:${params.orgId}`);
  const rawDocPolicy = params.storage.getItem(`dlp:docPolicy:${params.documentId}`);
  const rawClassifications = params.storage.getItem(`dlp:classifications:${params.documentId}`);

  const cached = aiCloudDlpCache.get(cacheKey);
  if (
    cached &&
    cached.rawOrgPolicy === rawOrgPolicy &&
    cached.rawDocPolicy === rawDocPolicy &&
    cached.rawClassifications === rawClassifications
  ) {
    return cached;
  }

  const classificationStore = cached?.classificationStore ?? new LocalClassificationStore({ storage: params.storage });

  let storedOrgPolicy = cached?.storedOrgPolicy ?? null;
  let storedDocumentPolicy = cached?.storedDocumentPolicy ?? null;
  let policy: any = cached?.policy;

  if (!cached || cached.rawOrgPolicy !== rawOrgPolicy || cached.rawDocPolicy !== rawDocPolicy) {
    const policyStore = new LocalPolicyStore({ storage: params.storage });
    storedOrgPolicy = policyStore.getOrgPolicy(params.orgId);
    storedDocumentPolicy = policyStore.getDocumentPolicy(params.documentId);

    let orgPolicy: any = createDefaultOrgPolicy();
    if (storedOrgPolicy) orgPolicy = storedOrgPolicy;
    policy = mergeEffectivePolicy({ orgPolicy, documentPolicy: storedDocumentPolicy });
  }

  let classificationRecords = cached?.classificationRecords ?? [];
  if (!cached || cached.rawClassifications !== rawClassifications) {
    classificationRecords = classificationStore.list(params.documentId);
  }

  const hasDlpConfig = Boolean(storedOrgPolicy || storedDocumentPolicy || classificationRecords.length > 0);

  const next: AiCloudDlpCacheEntry = {
    rawOrgPolicy,
    rawDocPolicy,
    rawClassifications,
    storedOrgPolicy,
    storedDocumentPolicy,
    policy,
    classificationRecords,
    classificationStore,
    hasDlpConfig,
  };

  aiCloudDlpCache.set(cacheKey, next);
  return next;
}

export type AiCloudDlpOptions = {
  /**
   * Optional, safe-to-log cache key for the effective DLP state (policy +
   * classification records + includeRestrictedContent).
   *
   * Surfaces can use this to cheaply detect DLP changes and avoid re-indexing /
   * reusing caches under a stricter policy.
   */
  cacheKey?: string;
  cache_key?: string;
  // ContextManager (camelCase)
  documentId: string;
  sheetId?: string;
  sheetNameResolver?: SheetNameResolver | null;
  policy: any;
  classificationRecords: Array<{ selector: any; classification: any }>;
  classificationStore: LocalClassificationStore;
  includeRestrictedContent: boolean;
  auditLogger: { log(event: any): void };
  // ToolExecutorOptions (snake_case)
  document_id: string;
  sheet_id?: string;
  sheet_name_resolver?: SheetNameResolver | null;
  classification_records: Array<{ selector: any; classification: any }>;
  classification_store: LocalClassificationStore;
  include_restricted_content: boolean;
  audit_logger: { log(event: any): void };
};

/**
 * Returns a DLP options object that can be passed to both:
 * - `ContextManager.buildWorkbookContextFromSpreadsheetApi({ dlp: ... })`
 * - `SpreadsheetLLMToolExecutor({ dlp: ... })`
 *
  * IMPORTANT: This helper is the bridge that wires enterprise DLP policy + per-cell
  * classifications into desktop AI surfaces (chat/agent/inline-edit).
  */
export function getAiCloudDlpOptions(params: {
  documentId: string;
  sheetId?: string;
  orgId?: string;
  sheetNameResolver?: SheetNameResolver | null;
}): AiCloudDlpOptions {
  const localStorage = getLocalStorageOrNull();
  const storage = localStorage ?? createMemoryStorage();
  const orgId = params.orgId ?? loadActiveOrgId(storage);

  let policy: any;
  let classificationStore: LocalClassificationStore;
  let classificationRecords: Array<{ selector: any; classification: any }>;

  if (localStorage !== null) {
    ({ policy, classificationRecords, classificationStore } = memoizedDlpData({ storage, documentId: params.documentId, orgId }));
  } else {
    policy = loadEffectivePolicy({ documentId: params.documentId, orgId, storage });
    classificationStore = new LocalClassificationStore({ storage });
    classificationRecords = classificationStore.list(params.documentId);
  }

  const auditLogger = getAiDlpAuditLogger();

  const out: AiCloudDlpOptions = {
    documentId: params.documentId,
    sheetId: params.sheetId,
    ...(params.sheetNameResolver !== undefined
      ? { sheetNameResolver: params.sheetNameResolver ?? null, sheet_name_resolver: params.sheetNameResolver ?? null }
      : {}),
    policy,
    classificationRecords,
    classificationStore,
    includeRestrictedContent: false,
    auditLogger,
    document_id: params.documentId,
    sheet_id: params.sheetId,
    classification_records: classificationRecords,
    classification_store: classificationStore,
    include_restricted_content: false,
    audit_logger: auditLogger
  };

  const cacheKey = computeDlpCacheKey(out);
  // `computeDlpCacheKey` may memoize `cacheKey` as a non-writable property. Avoid
  // assigning to it directly; just attach the snake_case alias for tool surfaces.
  try {
    Object.defineProperty(out, "cache_key", { value: cacheKey, enumerable: false, configurable: true });
  } catch {
    out.cache_key = cacheKey;
  }
  return out;
}

/**
 * Returns DLP options only when DLP has been configured in localStorage.
 *
 * Desktop AI surfaces treat DLP as an enterprise feature; when no org/document policy
 * or classification metadata is configured, we omit `dlp` entirely so downstream
 * systems (notably RAG indexing) can use their cheaper "no DLP" paths.
 */
export function maybeGetAiCloudDlpOptions(params: {
  documentId: string;
  sheetId?: string;
  orgId?: string;
  sheetNameResolver?: SheetNameResolver | null;
}): AiCloudDlpOptions | null {
  const localStorage = getLocalStorageOrNull();
  if (!localStorage) return null;

  const storage = localStorage;
  const orgId = params.orgId ?? loadActiveOrgId(storage);
  const memoized = memoizedDlpData({ storage, documentId: params.documentId, orgId });
  if (!memoized.hasDlpConfig) return null;

  const auditLogger = getAiDlpAuditLogger();

  const out: AiCloudDlpOptions = {
    documentId: params.documentId,
    sheetId: params.sheetId,
    ...(params.sheetNameResolver !== undefined
      ? { sheetNameResolver: params.sheetNameResolver ?? null, sheet_name_resolver: params.sheetNameResolver ?? null }
      : {}),
    policy: memoized.policy,
    classificationRecords: memoized.classificationRecords,
    classificationStore: memoized.classificationStore,
    includeRestrictedContent: false,
    auditLogger,
    document_id: params.documentId,
    sheet_id: params.sheetId,
    classification_records: memoized.classificationRecords,
    classification_store: memoized.classificationStore,
    include_restricted_content: false,
    audit_logger: auditLogger
  };

  const cacheKey = computeDlpCacheKey(out);
  try {
    Object.defineProperty(out, "cache_key", { value: cacheKey, enumerable: false, configurable: true });
  } catch {
    out.cache_key = cacheKey;
  }
  return out;
}
