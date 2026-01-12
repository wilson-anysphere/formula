export function detectExtensionPackageFormatVersion(packageBytes: Uint8Array): number;

export function createExtensionPackage(
  extensionDir: string,
  options?: { formatVersion?: number; privateKeyPem?: string | null },
): Promise<Uint8Array>;

export function readExtensionPackage(packageBytes: Uint8Array): any;

export function extractExtensionPackage(packageBytes: Uint8Array, destDir: string): Promise<void>;

export function verifyAndExtractExtensionPackage(
  packageBytes: Uint8Array,
  destDir: string,
  options: { publicKeyPem: string; signatureBase64?: string | null; formatVersion?: number; expectedId?: string | null; expectedVersion?: string | null },
): Promise<any>;

// v1 exports (backcompat)
export const createExtensionPackageV1: any;
export const readExtensionPackageV1: any;
export const extractExtensionPackageV1: any;
export const loadExtensionManifest: any;

// v2 exports
export const createExtensionPackageV2: any;
export const readExtensionPackageV2: any;
export const extractExtensionPackageV2: any;
export const verifyExtensionPackageV2: any;
export const createSignaturePayloadBytes: any;
export const canonicalJsonBytes: any;

// extracted directory integrity
export const verifyExtractedExtensionDir: any;

