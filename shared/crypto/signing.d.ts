export const SIGNATURE_ALGORITHM: string;

export function sha256(buffer: Buffer | Uint8Array | ArrayBuffer | string): string;

export function signBytes(
  bytes: Buffer | Uint8Array | ArrayBuffer,
  privateKeyPem: string,
  options?: { algorithm?: string },
): string;

export function verifyBytesSignature(
  bytes: Buffer | Uint8Array | ArrayBuffer,
  signatureBase64: string,
  publicKeyPem: string,
  options?: { algorithm?: string },
): boolean;

export function generateEd25519KeyPair(): { publicKeyPem: string; privateKeyPem: string };

