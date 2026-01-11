export type StorageLike = {
  getItem(key: string): string | null;
  setItem(key: string, value: string): void;
  removeItem(key: string): void;
};

export class LocalPolicyStore {
  constructor(params: { storage: StorageLike });
  getOrgPolicy(orgId: string): any | null;
  setOrgPolicy(orgId: string, policy: any): void;
  getDocumentPolicy(documentId: string): any | null;
  setDocumentPolicy(documentId: string, policy: any): void;
  clearDocumentPolicy(documentId: string): void;
}

export class CloudOrgPolicyStore {
  constructor(params: { fetchImpl?: typeof fetch; baseUrl: string; authToken?: string });
  get(orgId: string): Promise<any | null>;
  set(orgId: string, policy: any): Promise<any>;
}

export class CloudDocumentPolicyStore {
  constructor(params: { fetchImpl?: typeof fetch; baseUrl: string; authToken?: string });
  get(documentId: string): Promise<any | null>;
  set(documentId: string, policy: any): Promise<any>;
}

