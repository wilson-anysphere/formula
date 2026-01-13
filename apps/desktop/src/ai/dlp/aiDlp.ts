import { LocalClassificationStore, createMemoryStorage } from "../../../../../packages/security/dlp/src/classificationStore.js";
import { InMemoryAuditLogger } from "../../../../../packages/security/dlp/src/audit.js";
import { createDefaultOrgPolicy, mergePolicies } from "../../../../../packages/security/dlp/src/policy.js";
import { LocalPolicyStore } from "../../../../../packages/security/dlp/src/policyStore.js";

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

function loadEffectivePolicy(params: { documentId: string; orgId?: string; storage: StorageLike }): any {
  const policyStore = new LocalPolicyStore({ storage: params.storage });

  const orgId = params.orgId ?? loadActiveOrgId(params.storage);

  let orgPolicy: any = createDefaultOrgPolicy();
  const storedOrg = policyStore.getOrgPolicy(orgId);
  if (storedOrg) orgPolicy = storedOrg;

  const documentPolicy = policyStore.getDocumentPolicy(params.documentId);

  try {
    return mergePolicies({ orgPolicy, documentPolicy }).policy;
  } catch {
    // Corrupt localStorage / invalid policies should never take down AI surfaces.
    // Prefer falling back to the org policy (if valid) before using the hard-coded default.
    try {
      return mergePolicies({ orgPolicy, documentPolicy: undefined }).policy;
    } catch {
      // Safe baseline.
      return createDefaultOrgPolicy();
    }
  }
}

export type AiCloudDlpOptions = {
  // ContextManager (camelCase)
  documentId: string;
  sheetId?: string;
  policy: any;
  classificationRecords: Array<{ selector: any; classification: any }>;
  classificationStore: LocalClassificationStore;
  includeRestrictedContent: boolean;
  auditLogger: { log(event: any): void };
  // ToolExecutorOptions (snake_case)
  document_id: string;
  sheet_id?: string;
  classification_records: Array<{ selector: any; classification: any }>;
  classification_store: LocalClassificationStore;
  include_restricted_content: boolean;
  audit_logger: { log(event: any): void };
};

function hasDlpConfig(params: { storage: StorageLike; documentId: string; orgId?: string }): boolean {
  const policyStore = new LocalPolicyStore({ storage: params.storage });
  const orgId = params.orgId ?? loadActiveOrgId(params.storage);
  const storedOrg = policyStore.getOrgPolicy(orgId);
  const storedDoc = policyStore.getDocumentPolicy(params.documentId);
  const classificationStore = new LocalClassificationStore({ storage: params.storage });
  const classificationRecords = classificationStore.list(params.documentId);
  return Boolean(storedOrg || storedDoc || classificationRecords.length > 0);
}

/**
 * Returns a DLP options object that can be passed to both:
 * - `ContextManager.buildWorkbookContextFromSpreadsheetApi({ dlp: ... })`
 * - `SpreadsheetLLMToolExecutor({ dlp: ... })`
 *
 * IMPORTANT: This helper is the bridge that wires enterprise DLP policy + per-cell
 * classifications into desktop AI surfaces (chat/agent/inline-edit).
 */
export function getAiCloudDlpOptions(params: { documentId: string; sheetId?: string; orgId?: string }): AiCloudDlpOptions {
  const storage = getLocalStorageOrNull() ?? createMemoryStorage();

  const policy = loadEffectivePolicy({ documentId: params.documentId, orgId: params.orgId, storage });
  const classificationStore = new LocalClassificationStore({ storage });
  const classificationRecords = classificationStore.list(params.documentId);

  const auditLogger = getAiDlpAuditLogger();

  return {
    documentId: params.documentId,
    sheetId: params.sheetId,
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
}): AiCloudDlpOptions | null {
  const storage = getLocalStorageOrNull() ?? createMemoryStorage();
  if (!hasDlpConfig({ storage, documentId: params.documentId, orgId: params.orgId })) return null;
  return getAiCloudDlpOptions(params);
}
