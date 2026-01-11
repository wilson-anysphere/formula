import { InMemoryAuditLogger } from "../../../../packages/security/dlp/src/audit.js";
import { LocalClassificationStore, createMemoryStorage } from "../../../../packages/security/dlp/src/classificationStore.js";
import { createDefaultOrgPolicy, mergePolicies } from "../../../../packages/security/dlp/src/policy.js";
import { LocalPolicyStore } from "../../../../packages/security/dlp/src/policyStore.js";

type StorageLike = {
  getItem(key: string): string | null;
  setItem(key: string, value: string): void;
  removeItem(key: string): void;
};

export type DesktopDlpContext = {
  orgId: string;
  documentId: string;
  policy: any;
  classificationStore: LocalClassificationStore;
  auditLogger: InMemoryAuditLogger;
};

export function createDesktopDlpContext(params: {
  documentId: string;
  orgId?: string;
  storage?: StorageLike | null;
  auditLogger?: InMemoryAuditLogger;
  classificationStore?: LocalClassificationStore;
}): DesktopDlpContext {
  const storage: StorageLike = safeStorage(params.storage ?? getLocalStorageOrNull() ?? createMemoryStorage());
  const documentId = params.documentId;
  const orgId =
    typeof params.orgId === "string" && params.orgId.trim().length > 0 ? params.orgId.trim() : loadActiveOrgId(storage);

  const policyStore = new LocalPolicyStore({ storage });
  let orgPolicy: any = createDefaultOrgPolicy();
  const storedOrg = policyStore.getOrgPolicy(orgId);
  if (storedOrg) orgPolicy = storedOrg;
  const documentPolicy = policyStore.getDocumentPolicy(documentId) ?? undefined;

  // Corrupt localStorage / invalid policies should never take down desktop surfaces.
  // Prefer falling back to a safe baseline policy rather than throwing.
  let policy: any;
  try {
    policy = mergePolicies({ orgPolicy, documentPolicy }).policy;
  } catch {
    try {
      policy = mergePolicies({ orgPolicy, documentPolicy: undefined }).policy;
    } catch {
      policy = createDefaultOrgPolicy();
    }
  }

  const classificationStore = params.classificationStore ?? new LocalClassificationStore({ storage });
  const auditLogger = params.auditLogger ?? new InMemoryAuditLogger();

  return { orgId, documentId, policy, classificationStore, auditLogger };
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
    },
  };
}

function getLocalStorageOrNull(): StorageLike | null {
  if (typeof window !== "undefined") {
    try {
      const storage = window.localStorage as unknown as StorageLike | undefined;
      if (!storage) return null;
      if (typeof storage.getItem !== "function" || typeof storage.setItem !== "function") return null;
      return safeStorage(storage);
    } catch {
      // ignore
    }
  }

  try {
    if (typeof globalThis === "undefined") return null;
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const storage = (globalThis as any).localStorage as StorageLike | undefined;
    if (!storage) return null;
    if (typeof storage.getItem !== "function" || typeof storage.setItem !== "function") return null;
    return safeStorage(storage);
  } catch {
    return null;
  }
}
