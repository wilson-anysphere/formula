import { LocalClassificationStore, createMemoryStorage } from "../../../../packages/security/dlp/src/classificationStore.js";
import { createDefaultOrgPolicy, mergePolicies } from "../../../../packages/security/dlp/src/policy.js";
import { LocalPolicyStore } from "../../../../packages/security/dlp/src/policyStore.js";

import type { DesktopQueryEngineOptions } from "./engine.ts";

type StorageLike = { getItem(key: string): string | null; setItem(key: string, value: string): void; removeItem(key: string): void };

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
      const storage = window.localStorage as any;
      if (storage && typeof storage.getItem === "function" && typeof storage.setItem === "function") {
        return safeStorage(storage);
      }
    } catch {
      // ignore
    }
  }

  try {
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
    // Corrupt localStorage / invalid policies should never take down the Power Query UX.
    try {
      return mergePolicies({ orgPolicy, documentPolicy: undefined }).policy;
    } catch {
      return createDefaultOrgPolicy();
    }
  }
}

function hasDlpConfig(params: { storage: StorageLike; documentId: string; orgId?: string }): boolean {
  const policyStore = new LocalPolicyStore({ storage: params.storage });
  const orgId = params.orgId ?? loadActiveOrgId(params.storage);
  const storedOrg = policyStore.getOrgPolicy(orgId);
  const storedDoc = policyStore.getDocumentPolicy(params.documentId);
  const classificationStore = new LocalClassificationStore({ storage: params.storage });
  const records = classificationStore.list(params.documentId);
  return Boolean(storedOrg || storedDoc || records.length > 0);
}

/**
 * Best-effort helper to wire desktop Power Query executions into the local DLP policy engine.
 *
 * Desktop does not currently have first-class org/session state, so we read policy +
 * classification metadata from localStorage when available.
 */
export function maybeGetPowerQueryDlpContext(params: {
  documentId: string;
  sheetId?: string;
  range?: unknown;
  orgId?: string;
}): DesktopQueryEngineOptions["dlp"] | null {
  const storage = getLocalStorageOrNull() ?? createMemoryStorage();
  if (!hasDlpConfig({ storage, documentId: params.documentId, orgId: params.orgId })) return null;
  const policy = () => loadEffectivePolicy({ documentId: params.documentId, orgId: params.orgId, storage });
  const classificationStore = new LocalClassificationStore({ storage });

  return {
    documentId: params.documentId,
    sheetId: params.sheetId,
    range: params.range,
    classificationStore,
    policy,
  };
}
