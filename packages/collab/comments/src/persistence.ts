import * as Y from "yjs";

export interface KeyValueStorage {
  getItem(key: string): string | null;
  setItem(key: string, value: string): void;
  removeItem?(key: string): void;
}

export function loadDocFromStorage(doc: Y.Doc, storage: KeyValueStorage, key: string): boolean {
  const encoded = storage.getItem(key);
  if (!encoded) return false;

  const update = base64ToBytes(encoded);
  Y.applyUpdate(doc, update);
  return true;
}

export function saveDocToStorage(doc: Y.Doc, storage: KeyValueStorage, key: string): void {
  const update = Y.encodeStateAsUpdate(doc);
  storage.setItem(key, bytesToBase64(update));
}

export function bindDocToStorage(doc: Y.Doc, storage: KeyValueStorage, key: string): () => void {
  loadDocFromStorage(doc, storage, key);

  const handler = (): void => {
    saveDocToStorage(doc, storage, key);
  };

  doc.on("update", handler);
  return () => {
    doc.off("update", handler);
  };
}

function bytesToBase64(bytes: Uint8Array): string {
  const bufferCtor = (globalThis as any).Buffer as
    | (typeof Buffer & { from(data: Uint8Array): Buffer })
    | undefined;
  if (bufferCtor) {
    return bufferCtor.from(bytes).toString("base64");
  }

  let binary = "";
  for (const byte of bytes) {
    binary += String.fromCharCode(byte);
  }
  // eslint-disable-next-line no-undef
  return btoa(binary);
}

function base64ToBytes(encoded: string): Uint8Array {
  const bufferCtor = (globalThis as any).Buffer as
    | (typeof Buffer & { from(data: string, encoding: string): Buffer })
    | undefined;
  if (bufferCtor) {
    return Uint8Array.from(bufferCtor.from(encoded, "base64"));
  }

  // eslint-disable-next-line no-undef
  const binary = atob(encoded);
  const out = new Uint8Array(binary.length);
  for (let i = 0; i < binary.length; i++) {
    out[i] = binary.charCodeAt(i);
  }
  return out;
}

