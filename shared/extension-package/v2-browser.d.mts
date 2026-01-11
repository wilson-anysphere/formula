export interface VerifiedExtensionPackageFile {
  path: string;
  sha256: string;
  size: number;
}

export interface VerifiedExtensionPackageV2 {
  manifest: Record<string, any>;
  signatureBase64: string;
  files: VerifiedExtensionPackageFile[];
  unpackedSize: number;
  fileCount: number;
  readme: string;
}

export interface ReadExtensionPackageV2Result {
  formatVersion: number;
  manifest: Record<string, any>;
  checksums: Record<string, any>;
  signature: Record<string, any>;
  files: Map<string, Uint8Array>;
}

export function verifyExtensionPackageV2Browser(
  packageBytes: Uint8Array,
  publicKeyPem: string
): Promise<VerifiedExtensionPackageV2>;

export function readExtensionPackageV2(packageBytes: Uint8Array): ReadExtensionPackageV2Result;

export const PACKAGE_FORMAT_VERSION: number;

