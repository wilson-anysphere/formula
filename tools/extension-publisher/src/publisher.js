const fs = require("node:fs/promises");
const path = require("node:path");

const { createExtensionPackage, loadExtensionManifest, readExtensionPackage } = require("../../../shared/extension-package");
const { signBytes, sha256 } = require("../../../shared/crypto/signing");
const { validateExtensionManifest } = require("../../../shared/extension-manifest");

const NAME_RE = /^[a-z0-9][a-z0-9-]*$/;

function looksLikePemPrivateKey(value) {
  const text = String(value ?? "").trim();
  return text.startsWith("-----BEGIN") && text.includes("PRIVATE KEY");
}

async function resolvePrivateKeyPem(privateKeyPemOrPath) {
  const raw = String(privateKeyPemOrPath ?? "");
  const trimmed = raw.trim();
  if (!trimmed) throw new Error("privateKeyPemOrPath must be a non-empty string");
  if (looksLikePemPrivateKey(trimmed)) return trimmed;
  return await fs.readFile(raw, "utf8");
}

function validateManifest(manifest) {
  if (!manifest || typeof manifest !== "object") throw new Error("Manifest must be an object");

  const validated = validateExtensionManifest(manifest, { enforceEngine: false });

  if (!NAME_RE.test(validated.name)) {
    throw new Error(`Invalid extension name "${validated.name}" (expected ${NAME_RE})`);
  }

  if (!NAME_RE.test(validated.publisher)) {
    throw new Error(`Invalid publisher "${validated.publisher}" (expected ${NAME_RE})`);
  }

  return validated;
}

async function packageExtension(extensionDir, { privateKeyPem, formatVersion = 2 } = {}) {
  let manifest = await loadExtensionManifest(extensionDir);
  manifest = validateManifest(manifest);

  // Ensure the declared entrypoint exists before packaging. This prevents publishing
  // broken extensions when build output (e.g. dist/extension.js) is missing.
  const root = path.resolve(extensionDir);
  async function assertEntrypoint(fieldName, relPath) {
    const entryRel = String(relPath);
    const entryPath = path.resolve(root, entryRel);
    if (!entryPath.startsWith(root + path.sep)) {
      throw new Error(`Manifest ${fieldName} must resolve inside extensionDir (got ${relPath})`);
    }
    try {
      const stat = await fs.stat(entryPath);
      if (!stat.isFile()) {
        throw new Error(`Manifest ${fieldName} is not a file: ${relPath}`);
      }
    } catch (error) {
      if (error && error.code === "ENOENT") {
        throw new Error(
          `Manifest ${fieldName} entrypoint is missing: ${relPath}. Did you forget to build the extension?`
        );
      }
      throw error;
    }
  }

  await assertEntrypoint("main", manifest.main);
  if (typeof manifest.module === "string" && manifest.module.trim().length > 0) {
    await assertEntrypoint("module", manifest.module);
  }
  if (typeof manifest.browser === "string" && manifest.browser.trim().length > 0) {
    await assertEntrypoint("browser", manifest.browser);
  }

  const packageBytes = await createExtensionPackage(extensionDir, {
    formatVersion,
    privateKeyPem: formatVersion === 2 ? privateKeyPem : undefined,
  });

  let signatureBase64 = null;
  if (formatVersion === 1) {
    if (!privateKeyPem) throw new Error("privateKeyPem is required to sign v1 extension packages");
    signatureBase64 = signBytes(packageBytes, privateKeyPem);
  } else {
    const parsed = readExtensionPackage(packageBytes);
    signatureBase64 = parsed?.signature?.signatureBase64 || null;
  }

  return { manifest, packageBytes, signatureBase64, formatVersion };
}

async function signExtensionPackage(packageBytes, privateKeyPemOrPath) {
  const privateKeyPem = await resolvePrivateKeyPem(privateKeyPemOrPath);
  return signBytes(packageBytes, privateKeyPem);
}

async function publishExtension({ extensionDir, marketplaceUrl, token, privateKeyPemOrPath, formatVersion = 2 }) {
  if (!extensionDir) throw new Error("extensionDir is required");
  if (!marketplaceUrl) throw new Error("marketplaceUrl is required");
  if (!token) throw new Error("token is required");
  if (!privateKeyPemOrPath) throw new Error("privateKeyPemOrPath is required");

  // Publisher tooling expects `marketplaceUrl` to be the marketplace origin and appends `/api/...`.
  // To reduce common confusion (Desktop/Tauri clients typically use an API base URL that includes `/api`),
  // accept a trailing `/api` and normalize it away here.
  const normalizedMarketplaceUrl = (() => {
    const raw = String(marketplaceUrl ?? "").trim();
    // A defensive normalization: treat query/hash as invalid for the base URL and strip them so we
    // don't accidentally generate URLs like `https://host/api?x=y/api/publish-bin`.
    const withoutQueryHash = raw.split("#", 1)[0].split("?", 1)[0];
    const withoutTrailingSlash = withoutQueryHash.replace(/\/+$/, "");
    if (!withoutTrailingSlash) {
      throw new Error("marketplaceUrl must be a non-empty URL (e.g. https://marketplace.formula.app)");
    }
    if (withoutTrailingSlash.endsWith("/api")) {
      const stripped = withoutTrailingSlash.slice(0, -4);
      if (!stripped) {
        throw new Error("marketplaceUrl must be an origin URL, not a bare '/api' path");
      }
      return stripped;
    }
    return withoutTrailingSlash;
  })();

  const privateKeyPem = await resolvePrivateKeyPem(privateKeyPemOrPath);

  const { manifest, packageBytes, signatureBase64 } = await packageExtension(extensionDir, { privateKeyPem, formatVersion });

  const base = normalizedMarketplaceUrl;
  const packageSha256 = sha256(packageBytes);
  let response = await fetch(`${base}/api/publish-bin`, {
    method: "POST",
    headers: {
      Authorization: `Bearer ${token}`,
      "Content-Type": "application/vnd.formula.extension-package",
      "X-Package-Sha256": packageSha256,
      ...(formatVersion === 1 ? { "X-Package-Signature": signatureBase64 } : {}),
    },
    body: packageBytes,
  });

  if (response.status === 404) {
    // Backward compatibility: older marketplace servers only support JSON+base64 publishing.
    response = await fetch(`${base}/api/publish`, {
      method: "POST",
      headers: {
        Authorization: `Bearer ${token}`,
        "Content-Type": "application/json",
      },
      body: JSON.stringify({
        packageBase64: packageBytes.toString("base64"),
        ...(formatVersion === 1 ? { signatureBase64 } : {}),
      }),
    });
  }

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
