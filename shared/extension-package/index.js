const v1 = require("./v1");
const v2 = require("./v2");

function detectExtensionPackageFormatVersion(packageBytes) {
  if (!packageBytes || packageBytes.length < 2) {
    throw new Error("Invalid extension package (empty)");
  }

  // v1 packages are gzipped JSON.
  if (packageBytes[0] === 0x1f && packageBytes[1] === 0x8b) return 1;

  // v2 packages are deterministic tar archives with required entries.
  return 2;
}

async function createExtensionPackage(extensionDir, options = {}) {
  const formatVersion = options.formatVersion ?? 2;
  if (formatVersion === 1) {
    return v1.createExtensionPackageV1(extensionDir);
  }
  if (formatVersion !== 2) {
    throw new Error(`Unsupported extension package formatVersion: ${formatVersion}`);
  }
  return v2.createExtensionPackageV2(extensionDir, { privateKeyPem: options.privateKeyPem });
}

function readExtensionPackage(packageBytes) {
  const version = detectExtensionPackageFormatVersion(packageBytes);
  if (version === 1) return v1.readExtensionPackageV1(packageBytes);
  return v2.readExtensionPackageV2(packageBytes);
}

async function extractExtensionPackage(packageBytes, destDir) {
  const version = detectExtensionPackageFormatVersion(packageBytes);
  if (version === 1) return v1.extractExtensionPackageV1(packageBytes, destDir);
  return v2.extractExtensionPackageV2(packageBytes, destDir);
}

module.exports = {
  detectExtensionPackageFormatVersion,

  createExtensionPackage,
  readExtensionPackage,
  extractExtensionPackage,

  // v1 exports (backward compatibility)
  createExtensionPackageV1: v1.createExtensionPackageV1,
  readExtensionPackageV1: v1.readExtensionPackageV1,
  extractExtensionPackageV1: v1.extractExtensionPackageV1,
  loadExtensionManifest: v1.loadExtensionManifest,

  // v2 exports
  createExtensionPackageV2: v2.createExtensionPackageV2,
  readExtensionPackageV2: v2.readExtensionPackageV2,
  extractExtensionPackageV2: v2.extractExtensionPackageV2,
  verifyExtensionPackageV2: v2.verifyExtensionPackageV2,
  createSignaturePayloadBytes: v2.createSignaturePayloadBytes,
  canonicalJsonBytes: v2.canonicalJsonBytes,
};

