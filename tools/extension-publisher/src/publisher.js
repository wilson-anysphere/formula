const fs = require("node:fs/promises");
const path = require("node:path");

const { createExtensionPackage, loadExtensionManifest, readExtensionPackage } = require("../../../shared/extension-package");
const { signBytes } = require("../../../shared/crypto/signing");
const { isValidSemver } = require("../../../shared/semver");

const NAME_RE = /^[a-z0-9][a-z0-9-]*$/;

function validateManifest(manifest) {
  if (!manifest || typeof manifest !== "object") throw new Error("Manifest must be an object");

  const required = ["name", "publisher", "version", "main"];
  for (const field of required) {
    if (!manifest[field] || typeof manifest[field] !== "string") {
      throw new Error(`Manifest missing required string field: ${field}`);
    }
  }

  if (!NAME_RE.test(manifest.name)) {
    throw new Error(`Invalid extension name "${manifest.name}" (expected ${NAME_RE})`);
  }

  if (!NAME_RE.test(manifest.publisher)) {
    throw new Error(`Invalid publisher "${manifest.publisher}" (expected ${NAME_RE})`);
  }

  if (!isValidSemver(manifest.version)) {
    throw new Error(`Invalid version "${manifest.version}" (expected semver)`);
  }

  if (!manifest.engines || typeof manifest.engines !== "object" || typeof manifest.engines.formula !== "string") {
    throw new Error("Manifest missing required engines.formula string field");
  }

  return true;
}

async function packageExtension(extensionDir, { privateKeyPem } = {}) {
  const manifest = await loadExtensionManifest(extensionDir);
  validateManifest(manifest);

  // Ensure the declared entrypoint exists before packaging. This prevents publishing
  // broken extensions when build output (e.g. dist/extension.js) is missing.
  const root = path.resolve(extensionDir);
  const mainRel = String(manifest.main);
  const mainPath = path.resolve(root, mainRel);
  if (!mainPath.startsWith(root + path.sep)) {
    throw new Error(`Manifest main must resolve inside extensionDir (got ${manifest.main})`);
  }
  try {
    const stat = await fs.stat(mainPath);
    if (!stat.isFile()) {
      throw new Error(`Manifest main is not a file: ${manifest.main}`);
    }
  } catch (error) {
    if (error && error.code === "ENOENT") {
      throw new Error(
        `Manifest main entrypoint is missing: ${manifest.main}. Did you forget to build the extension?`
      );
    }
    throw error;
  }

  const packageBytes = await createExtensionPackage(extensionDir, {
    formatVersion: 2,
    privateKeyPem,
  });

  const parsed = readExtensionPackage(packageBytes);
  const signatureBase64 = parsed?.signature?.signatureBase64 || null;

  return { manifest, packageBytes, signatureBase64 };
}

async function signExtensionPackage(packageBytes, privateKeyPemOrPath) {
  let privateKeyPem = privateKeyPemOrPath;
  if (privateKeyPemOrPath.includes(path.sep) || privateKeyPemOrPath.includes(".pem")) {
    privateKeyPem = await fs.readFile(privateKeyPemOrPath, "utf8");
  }
  return signBytes(packageBytes, privateKeyPem);
}

async function publishExtension({ extensionDir, marketplaceUrl, token, privateKeyPemOrPath }) {
  if (!extensionDir) throw new Error("extensionDir is required");
  if (!marketplaceUrl) throw new Error("marketplaceUrl is required");
  if (!token) throw new Error("token is required");
  if (!privateKeyPemOrPath) throw new Error("privateKeyPemOrPath is required");

  let privateKeyPem = privateKeyPemOrPath;
  if (privateKeyPemOrPath.includes(path.sep) || privateKeyPemOrPath.includes(".pem")) {
    privateKeyPem = await fs.readFile(privateKeyPemOrPath, "utf8");
  }

  const { manifest, packageBytes, signatureBase64 } = await packageExtension(extensionDir, { privateKeyPem });

  const response = await fetch(`${marketplaceUrl.replace(/\/$/, "")}/api/publish`, {
    method: "POST",
    headers: {
      Authorization: `Bearer ${token}`,
      "Content-Type": "application/json",
    },
    body: JSON.stringify({
      packageBase64: packageBytes.toString("base64"),
      signatureBase64,
    }),
  });

  if (!response.ok) {
    const text = await response.text();
    throw new Error(`Publish failed (${response.status}): ${text}`);
  }

  const result = await response.json();
  return { ...result, manifest };
}

module.exports = {
  packageExtension,
  publishExtension,
  signExtensionPackage,
  validateManifest,
};
