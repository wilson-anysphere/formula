const v1 = require("./v1");
const v2 = require("./v2");
const fs = require("node:fs/promises");
const path = require("node:path");
const crypto = require("node:crypto");

const { verifyBytesSignature } = require("../crypto/signing");

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

async function verifyAndExtractExtensionPackage(packageBytes, destDir, options = {}) {
  if (!options || typeof options !== "object") {
    throw new Error("verifyAndExtractExtensionPackage: options must be an object");
  }
  if (!options.publicKeyPem || typeof options.publicKeyPem !== "string") {
    throw new Error("verifyAndExtractExtensionPackage: publicKeyPem is required");
  }
  if (!destDir || typeof destDir !== "string") {
    throw new Error("verifyAndExtractExtensionPackage: destDir must be a string");
  }

  const formatVersion = options.formatVersion ?? detectExtensionPackageFormatVersion(packageBytes);

  const expectedId = options.expectedId ?? null;
  const expectedVersion = options.expectedVersion ?? null;

  /** @type {any} */
  let manifest = null;
  let signatureBase64 = null;

  /** @type {(stagingDir: string) => Promise<void>} */
  let writeToStaging = null;

  if (formatVersion === 1) {
    if (!options.signatureBase64 || typeof options.signatureBase64 !== "string") {
      throw new Error("verifyAndExtractExtensionPackage: signatureBase64 is required for v1 packages");
    }

    const signatureOk = verifyBytesSignature(packageBytes, options.signatureBase64, options.publicKeyPem);
    if (!signatureOk) {
      throw new Error("Extension signature verification failed (mandatory)");
    }

    const bundle = v1.readExtensionPackageV1(packageBytes);
    manifest = bundle.manifest;
    signatureBase64 = options.signatureBase64;
    writeToStaging = (stagingDir) => v1.extractExtensionPackageV1FromBundle(bundle, stagingDir);
  } else if (formatVersion === 2) {
    const parsed = v2.readExtensionPackageV2(packageBytes);

    let verified;
    try {
      verified = v2.verifyExtensionPackageV2Parsed(parsed, options.publicKeyPem);
    } catch (error) {
      throw new Error(`Extension signature verification failed (mandatory): ${error?.message ?? String(error)}`);
    }

    manifest = verified.manifest;
    signatureBase64 = verified.signatureBase64;

    // Optional transport cross-check: if the server included an X-Package-Signature header,
    // ensure it matches the signed payload inside the package.
    if (options.signatureBase64 && options.signatureBase64 !== signatureBase64) {
      throw new Error("Marketplace signature header does not match package signature");
    }

    writeToStaging = (stagingDir) => v2.extractExtensionPackageV2FromParsed(parsed, stagingDir);
  } else {
    throw new Error(`Unsupported extension package formatVersion: ${formatVersion}`);
  }

  if (!manifest || typeof manifest !== "object") {
    throw new Error("Invalid extension package: missing manifest");
  }
  const actualId = `${manifest.publisher}.${manifest.name}`;
  if (expectedId && actualId !== expectedId) {
    throw new Error(`Package id mismatch: expected ${expectedId} but got ${actualId}`);
  }
  if (expectedVersion && manifest.version !== expectedVersion) {
    throw new Error(`Package version mismatch: expected ${expectedVersion} but got ${manifest.version}`);
  }

  await fs.mkdir(path.dirname(destDir), { recursive: true });

  const stagingDir = await fs.mkdtemp(path.join(path.dirname(destDir), `.${path.basename(destDir)}.staging-`));
  let committed = false;

  try {
    await writeToStaging(stagingDir);
    await atomicReplaceDir(stagingDir, destDir);
    committed = true;
  } finally {
    if (!committed) {
      await fs.rm(stagingDir, { recursive: true, force: true }).catch(() => {});
    }
  }

  return { manifest, formatVersion, signatureBase64 };
}

async function atomicReplaceDir(stagingDir, destDir) {
  const parent = path.dirname(destDir);
  const base = path.basename(destDir);
  const backupDir = path.join(parent, `.${base}.backup-${crypto.randomUUID()}`);

  let hadExisting = false;
  try {
    const st = await fs.lstat(destDir);
    if (!st.isDirectory()) {
      throw new Error(`Extension install destination exists and is not a directory: ${destDir}`);
    }
    hadExisting = true;
  } catch (error) {
    if (error && (error.code === "ENOENT" || error.code === "ENOTDIR")) {
      hadExisting = false;
    } else {
      throw error;
    }
  }

  if (hadExisting) {
    await fs.rename(destDir, backupDir);
  }

  try {
    await fs.rename(stagingDir, destDir);
  } catch (error) {
    if (hadExisting) {
      try {
        await fs.rename(backupDir, destDir);
      } catch (restoreError) {
        throw new Error(
          `Failed to finalize extension install, and failed to restore previous install: ${restoreError?.message ?? String(restoreError)}`,
          { cause: error },
        );
      }
    }
    throw error;
  }

  if (hadExisting) {
    await fs.rm(backupDir, { recursive: true, force: true }).catch(() => {});
  }
}

module.exports = {
  detectExtensionPackageFormatVersion,

  createExtensionPackage,
  readExtensionPackage,
  extractExtensionPackage,
  verifyAndExtractExtensionPackage,

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
