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
  const orgId = params.orgId ?? "local-org";
  const documentId = params.documentId;

  const storage: StorageLike = params.storage ?? getLocalStorageOrNull() ?? createMemoryStorage();

  const policyStore = new LocalPolicyStore({ storage });
  const orgPolicy = policyStore.getOrgPolicy(orgId) ?? createDefaultOrgPolicy();
  const documentPolicy = policyStore.getDocumentPolicy(documentId) ?? undefined;
  const { policy } = mergePolicies({ orgPolicy, documentPolicy });

  const classificationStore = params.classificationStore ?? new LocalClassificationStore({ storage });
  const auditLogger = params.auditLogger ?? new InMemoryAuditLogger();

  return { orgId, documentId, policy, classificationStore, auditLogger };
}

function getLocalStorageOrNull(): StorageLike | null {
  if (typeof window !== "undefined") {
    try {
      const storage = window.localStorage as unknown as StorageLike | undefined;
      if (!storage) return null;
      if (typeof storage.getItem !== "function" || typeof storage.setItem !== "function") return null;
      return storage;
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
    return storage;
  } catch {
    return null;
  }
}
