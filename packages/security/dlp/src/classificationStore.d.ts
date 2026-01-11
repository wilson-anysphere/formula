export type StorageLike = {
  getItem(key: string): string | null;
  setItem(key: string, value: string): void;
  removeItem(key: string): void;
};

export function createMemoryStorage(): StorageLike;

export class LocalClassificationStore {
  constructor(params: { storage: StorageLike });
  list(documentId: string): Array<{ selector: any; classification: any; updatedAt: string }>;
  upsert(documentId: string, selector: any, classification: any): void;
  remove(documentId: string, selector: any): void;
}

export class CloudClassificationStore {
  constructor(params: { fetchImpl?: typeof fetch; baseUrl: string; authToken?: string });
  list(documentId: string): Promise<any[]>;
  upsert(documentId: string, selector: any, classification: any): Promise<void>;
  remove(documentId: string, selector: any): Promise<void>;
}

export class HybridClassificationStore {
  constructor(params: { local: LocalClassificationStore; cloud: CloudClassificationStore });
  list(documentId: string): Array<{ selector: any; classification: any; updatedAt: string }>;
  syncFromCloud(documentId: string): Promise<any[]>;
  upsert(documentId: string, selector: any, classification: any): Promise<void>;
  remove(documentId: string, selector: any): Promise<void>;
}

