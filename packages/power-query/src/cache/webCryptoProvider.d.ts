export type WebCryptoCacheProvider = {
  keyVersion: number;
  encryptBytes: (
    plaintext: Uint8Array,
    aad?: Uint8Array,
  ) => Promise<{ keyVersion: number; iv: Uint8Array; tag: Uint8Array; ciphertext: Uint8Array }>;
  decryptBytes: (
    payload: { keyVersion: number; iv: Uint8Array; tag: Uint8Array; ciphertext: Uint8Array },
    aad?: Uint8Array,
  ) => Promise<Uint8Array>;
};

export function createWebCryptoCacheProvider(options: { keyVersion: number; keyBytes: Uint8Array }): Promise<WebCryptoCacheProvider>;
