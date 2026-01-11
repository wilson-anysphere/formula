import { createHash } from "node:crypto";

export function sha256Hex(value: string): string {
  return createHash("sha256").update(value).digest("hex");
}

export function persistedDocName(docName: string, hashingEnabled: boolean): string {
  return hashingEnabled ? sha256Hex(docName) : docName;
}

export type LeveldbPersistenceLike = {
  flushDocument: (docName: string) => Promise<void>;
  getYDoc: (docName: string) => Promise<any>;
  storeUpdate: (docName: string, update: Uint8Array) => Promise<any>;
  getStateVector?: (docName: string) => Promise<Uint8Array>;
  getDiff?: (docName: string, stateVector: Uint8Array) => Promise<Uint8Array>;
  clearDocument?: (docName: string) => Promise<void>;
  setMeta?: (docName: string, metaKey: string, value: unknown) => Promise<void>;
  delMeta?: (docName: string, metaKey: string) => Promise<void>;
  getMeta?: (docName: string, metaKey: string) => Promise<unknown>;
  getMetas?: (docName: string) => Promise<Map<string, unknown>>;
  getAllDocNames?: () => Promise<string[]>;
  getAllDocStateVectors?: () => Promise<
    Array<{ name: string; sv: Uint8Array; clock: number }>
  >;
  destroy: () => Promise<void>;
  clearAll?: () => Promise<void>;
};

export class LeveldbDocNameHashingLayer {
  constructor(
    private readonly inner: LeveldbPersistenceLike,
    private readonly hashingEnabled: boolean
  ) {}

  persistedName(docName: string): string {
    return persistedDocName(docName, this.hashingEnabled);
  }

  flushDocument(docName: string) {
    return this.inner.flushDocument(this.persistedName(docName));
  }

  getYDoc(docName: string) {
    return this.inner.getYDoc(this.persistedName(docName));
  }

  storeUpdate(docName: string, update: Uint8Array) {
    return this.inner.storeUpdate(this.persistedName(docName), update);
  }

  getStateVector(docName: string) {
    if (!this.inner.getStateVector) {
      throw new Error("LevelDB provider does not implement getStateVector()");
    }
    return this.inner.getStateVector(this.persistedName(docName));
  }

  getDiff(docName: string, stateVector: Uint8Array) {
    if (!this.inner.getDiff) {
      throw new Error("LevelDB provider does not implement getDiff()");
    }
    return this.inner.getDiff(this.persistedName(docName), stateVector);
  }

  clearDocument(docName: string) {
    if (!this.inner.clearDocument) {
      throw new Error("LevelDB provider does not implement clearDocument()");
    }
    return this.inner.clearDocument(this.persistedName(docName));
  }

  setMeta(docName: string, metaKey: string, value: unknown) {
    if (!this.inner.setMeta) {
      throw new Error("LevelDB provider does not implement setMeta()");
    }
    return this.inner.setMeta(this.persistedName(docName), metaKey, value);
  }

  delMeta(docName: string, metaKey: string) {
    if (!this.inner.delMeta) {
      throw new Error("LevelDB provider does not implement delMeta()");
    }
    return this.inner.delMeta(this.persistedName(docName), metaKey);
  }

  getMeta(docName: string, metaKey: string) {
    if (!this.inner.getMeta) {
      throw new Error("LevelDB provider does not implement getMeta()");
    }
    return this.inner.getMeta(this.persistedName(docName), metaKey);
  }

  getMetas(docName: string) {
    if (!this.inner.getMetas) {
      throw new Error("LevelDB provider does not implement getMetas()");
    }
    return this.inner.getMetas(this.persistedName(docName));
  }

  getAllDocNames() {
    if (!this.inner.getAllDocNames) {
      throw new Error("LevelDB provider does not implement getAllDocNames()");
    }
    return this.inner.getAllDocNames();
  }

  getAllDocStateVectors() {
    if (!this.inner.getAllDocStateVectors) {
      throw new Error(
        "LevelDB provider does not implement getAllDocStateVectors()"
      );
    }
    return this.inner.getAllDocStateVectors();
  }

  destroy() {
    return this.inner.destroy();
  }

  clearAll() {
    if (!this.inner.clearAll) {
      throw new Error("LevelDB provider does not implement clearAll()");
    }
    return this.inner.clearAll();
  }
}
