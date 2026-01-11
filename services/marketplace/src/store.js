const fs = require("node:fs/promises");
const path = require("node:path");
const crypto = require("node:crypto");
const zlib = require("node:zlib");

const { verifyBytesSignature, sha256 } = require("../../../shared/crypto/signing");
const {
  canonicalJsonBytes,
  detectExtensionPackageFormatVersion,
  readExtensionPackageV2,
  verifyExtensionPackageV2,
} = require("../../../shared/extension-package");
const { validateExtensionManifest } = require("../../../shared/extension-manifest");
const { compareSemver, maxSemver } = require("../../../shared/semver");

const { SqliteFileDb } = require("./db/sqlite");

function normalizeStringArray(value) {
  if (!Array.isArray(value)) return [];
  return value
    .filter((v) => typeof v === "string")
    .map((v) => v.trim())
    .filter((v) => v.length > 0);
}

function extensionIdFromManifest(manifest) {
  if (!manifest?.publisher || !manifest?.name) {
    throw new Error("Manifest is missing required fields (publisher, name)");
  }
  return `${manifest.publisher}.${manifest.name}`;
}

function tokenize(query) {
  const tokens = String(query || "")
    .toLowerCase()
    .split(/\s+/)
    .map((t) => t.trim())
    .filter(Boolean);
  return Array.from(new Set(tokens));
}

function calculateSearchScore(ext, tokens) {
  if (tokens.length === 0) return 0;

  let score = 0;

  const fields = [
    { text: ext.id, weight: 12 },
    { text: ext.displayName, weight: 10 },
    { text: ext.description, weight: 6 },
    { text: ext.publisher, weight: 4 },
    { text: (ext.tags || []).join(" "), weight: 3 },
    { text: (ext.categories || []).join(" "), weight: 2 },
  ];

  for (const token of tokens) {
    let bestWeight = 0;
    let matchCount = 0;
    for (const field of fields) {
      if (String(field.text || "").toLowerCase().includes(token)) {
        matchCount += 1;
        if (field.weight > bestWeight) bestWeight = field.weight;
      }
    }
    // All tokens must match somewhere (across any fields).
    if (bestWeight === 0) return 0;

    score += bestWeight;
    // Small bonus for tokens that match across multiple fields.
    score += Math.min(2, Math.max(0, matchCount - 1));
  }

  const phrase = tokens.join(" ");
  if (
    phrase &&
    (String(ext.id || "").toLowerCase().includes(phrase) || String(ext.displayName || "").toLowerCase().includes(phrase))
  ) {
    score += 5;
  }

  if (ext.featured) score += 8;
  if (ext.verified) score += 3;
  score += Math.min(5, Math.log10((ext.downloadCount || 0) + 1));

  return score;
}

const PACKAGE_FORMAT = "formula-extension-package";
const PACKAGE_FORMAT_VERSION = 1;
const NAME_RE = /^[a-z0-9][a-z0-9-]*$/;

const PACKAGE_SCAN_STATUS = {
  PENDING: "pending",
  PASSED: "passed",
  FAILED: "failed",
};

const DEFAULT_JS_HEURISTIC_RULES = [
  {
    id: "js.child_process",
    message: "References Node.js child_process APIs",
    test: (source) => /\bchild_process\b/.test(source),
  },
  {
    id: "js.eval",
    message: "Uses eval()",
    test: (source) => /\beval\s*\(/.test(source),
  },
  {
    id: "js.new_function",
    message: "Uses new Function()",
    test: (source) => /\bnew\s+Function\s*\(/.test(source),
  },
  {
    id: "js.obfuscation.hex_escape",
    message: "Contains repeated hex escape sequences (possible obfuscation)",
    test: (source) => /(\\x[0-9a-fA-F]{2}){16,}/.test(source),
  },
];

function nowIso() {
  return new Date().toISOString();
}

function parseCsvSet(value) {
  if (!value) return new Set();
  return new Set(
    String(value)
      .split(",")
      .map((v) => v.trim())
      .filter(Boolean)
  );
}

function sha256Utf8(value) {
  return sha256(Buffer.from(String(value ?? ""), "utf8"));
}

function parseJsonArray(json, fallback = []) {
  try {
    const parsed = JSON.parse(String(json || "[]"));
    return Array.isArray(parsed) ? parsed : fallback;
  } catch {
    return fallback;
  }
}

function safeJsonStringify(value) {
  return JSON.stringify(value ?? null);
}

function randomId() {
  return crypto.randomUUID();
}

function normalizePublicKeyPem(publicKeyPem) {
  const trimmed = String(publicKeyPem || "").trim();
  return trimmed ? trimmed + "\n" : "";
}

function publisherKeyIdFromPublicKeyPem(publicKeyPem) {
  const key = crypto.createPublicKey(publicKeyPem);
  if (key.asymmetricKeyType !== "ed25519") {
    throw new Error(`Unsupported public key type: ${key.asymmetricKeyType} (expected ed25519)`);
  }
  const der = key.export({ type: "spki", format: "der" });
  return crypto.createHash("sha256").update(der).digest("hex");
}

function safeJoin(baseDir, relPath) {
  const normalized = String(relPath || "").replace(/\\\\/g, "/");
  if (normalized.startsWith("/") || normalized.includes("..")) {
    throw new Error(`Invalid relative path: ${relPath}`);
  }
  const full = path.join(baseDir, normalized);
  const relative = path.relative(baseDir, full);
  if (relative.startsWith("..") || path.isAbsolute(relative)) {
    throw new Error(`Path traversal in path: ${relPath}`);
  }
  return full;
}

function resolvePackagePath(dataDir, storedPath) {
  if (!storedPath) return null;
  const raw = String(storedPath);
  if (path.isAbsolute(raw)) {
    const full = path.resolve(raw);
    const base = path.resolve(dataDir);
    if (full === base || full.startsWith(base + path.sep)) return full;
    throw new Error(`Invalid package path (outside dataDir): ${storedPath}`);
  }
  return safeJoin(dataDir, raw);
}

function packageRelPath(extensionId, version) {
  return path.posix.join("packages", String(extensionId), `${String(version)}.fextpkg`);
}

async function ensurePackageFile({ dataDir, relPath, bytes, expectedSha256 }) {
  const fullPath = safeJoin(dataDir, relPath);
  await fs.mkdir(path.dirname(fullPath), { recursive: true });

  try {
    const existing = await fs.readFile(fullPath);
    const existingSha = sha256(existing);
    if (existingSha !== expectedSha256) {
      throw new Error("Package file already exists with different contents");
    }
    return fullPath;
  } catch (error) {
    if (error && (error.code === "ENOENT" || error.code === "ENOTDIR")) {
      // continue
    } else if (error) {
      throw error;
    }
  }

  const tmpPath = `${fullPath}.${randomId()}.tmp`;
  await fs.writeFile(tmpPath, bytes);
  try {
    await fs.link(tmpPath, fullPath);
    return fullPath;
  } catch (error) {
    if (error && error.code === "EEXIST") {
      const existing = await fs.readFile(fullPath);
      const existingSha = sha256(existing);
      if (existingSha !== expectedSha256) {
        throw new Error("Package file already exists with different contents");
      }
      return fullPath;
    }
    throw error;
  } finally {
    await fs.unlink(tmpPath).catch(() => {});
  }
}

async function gunzipWithLimit(packageBytes, maxBytes) {
  return new Promise((resolve, reject) => {
    const gunzip = zlib.createGunzip();
    const chunks = [];
    let size = 0;
    gunzip.on("data", (chunk) => {
      size += chunk.length;
      if (size > maxBytes) {
        gunzip.destroy(new Error("Package exceeds maximum uncompressed size"));
        return;
      }
      chunks.push(chunk);
    });
    gunzip.on("error", (err) => reject(err));
    gunzip.on("end", () => resolve(Buffer.concat(chunks)));
    gunzip.end(packageBytes);
  });
}

async function readExtensionPackageSafe(packageBytes, { maxUncompressedBytes }) {
  const jsonBytes = await gunzipWithLimit(packageBytes, maxUncompressedBytes);
  const parsed = JSON.parse(jsonBytes.toString("utf8"));

  if (parsed?.format !== PACKAGE_FORMAT || parsed?.formatVersion !== PACKAGE_FORMAT_VERSION) {
    throw new Error("Unsupported extension package format");
  }

  if (!Array.isArray(parsed.files) || typeof parsed.manifest !== "object" || !parsed.manifest) {
    throw new Error("Invalid extension package contents");
  }

  return parsed;
}

function validateManifest(manifest) {
  if (!manifest || typeof manifest !== "object") throw new Error("Manifest must be an object");

  let validated;
  try {
    validated = validateExtensionManifest(manifest, { enforceEngine: false });
  } catch (error) {
    if (error && typeof error === "object" && error.name === "ManifestError") {
      throw new Error(`Invalid manifest: ${error.message}`);
    }
    throw error;
  }

  if (!NAME_RE.test(validated.name)) {
    throw new Error(`Invalid extension name "${validated.name}" (expected ${NAME_RE})`);
  }

  if (!NAME_RE.test(validated.publisher)) {
    throw new Error(`Invalid publisher "${validated.publisher}" (expected ${NAME_RE})`);
  }

  const categories = normalizeStringArray(validated.categories);
  const tags = normalizeStringArray(validated.tags);

  if (categories.length > 5) throw new Error("Manifest categories must have at most 5 entries");
  if (tags.length > 10) throw new Error("Manifest tags must have at most 10 entries");

  for (const c of categories) {
    if (c.length > 32) throw new Error("Manifest category entries must be <= 32 characters");
  }
  for (const t of tags) {
    if (t.length > 32) throw new Error("Manifest tag entries must be <= 32 characters");
  }

  return validated;
}

function isAllowedFilePath(filePath) {
  const normalized = String(filePath || "").replace(/\\/g, "/");
  if (normalized.length === 0) return false;
  if (normalized.startsWith("/")) return false;
  if (normalized.includes("\0")) return false;
  const parts = normalized.split("/");
  if (parts.some((p) => p === "" || p === "." || p === "..")) return false;
  if (parts.some((p) => p.includes(":"))) return false;
  const windowsReservedRe = /^(con|prn|aux|nul|com[1-9]|lpt[1-9])(\..*)?$/i;
  for (const part of parts) {
    if (/[<>:"|?*]/.test(part)) return false;
    if (part.endsWith(" ") || part.endsWith(".")) return false;
    if (windowsReservedRe.test(part)) return false;
  }

  const lower = normalized.toLowerCase();
  const ext = path.extname(lower);
  const base = path.basename(lower);

  // Allow a conservative set of extensions used by Formula extensions.
  const allowedExts = new Set([
    ".js",
    ".cjs",
    ".mjs",
    ".json",
    ".md",
    ".txt",
    ".css",
    ".html",
    ".svg",
    ".png",
    ".jpg",
    ".jpeg",
    ".gif",
    ".wasm",
    ".map",
  ]);

  if (allowedExts.has(ext)) return true;

  // Allow common metadata files without extensions.
  if (ext === "" && (base === "license" || base === "notice" || base === "readme")) return true;
  return false;
}

function normalizeManifestFilePath(filePath) {
  const normalized = String(filePath || "").trim().replace(/\\/g, "/");
  if (!normalized) return "";
  if (normalized.startsWith("/")) throw new Error(`Manifest file path must be relative (got ${filePath})`);
  if (normalized.includes("\0")) throw new Error("Manifest file path contains NUL byte");
  let rel = normalized;
  while (rel.startsWith("./")) rel = rel.slice(2);
  const parts = rel.split("/");
  if (parts.some((p) => p === "" || p === "." || p === "..")) {
    throw new Error(`Invalid manifest file path: ${filePath}`);
  }
  return parts.join("/");
}

function validateEntrypoint(manifest, fileRecords) {
  const files = new Set((fileRecords || []).map((f) => String(f?.path || "")));

  const main = normalizeManifestFilePath(manifest?.main);
  if (!main) throw new Error("Manifest main entrypoint is required");
  if (!isAllowedFilePath(main)) {
    throw new Error(`Manifest main entrypoint must be an allowed file type (got ${manifest.main})`);
  }
  if (!files.has(main)) {
    throw new Error(`Manifest main entrypoint is missing from extension package: ${manifest.main}`);
  }

  const moduleRaw = manifest?.module;
  if (moduleRaw !== undefined && moduleRaw !== null) {
    const moduleEntry = normalizeManifestFilePath(moduleRaw);
    if (!moduleEntry) throw new Error("Manifest module entrypoint must be a non-empty string");
    if (!isAllowedFilePath(moduleEntry)) {
      throw new Error(`Manifest module entrypoint must be an allowed file type (got ${moduleRaw})`);
    }
    if (!files.has(moduleEntry)) {
      throw new Error(`Manifest module entrypoint is missing from extension package: ${moduleRaw}`);
    }
  }

  const browserRaw = manifest?.browser;
  if (browserRaw !== undefined && browserRaw !== null) {
    const browserEntry = normalizeManifestFilePath(browserRaw);
    if (!browserEntry) throw new Error("Manifest browser entrypoint must be a non-empty string");
    if (!isAllowedFilePath(browserEntry)) {
      throw new Error(`Manifest browser entrypoint must be an allowed file type (got ${browserRaw})`);
    }
    if (!files.has(browserEntry)) {
      throw new Error(`Manifest browser entrypoint is missing from extension package: ${browserRaw}`);
    }
  }
}

class MarketplaceStore {
  constructor({ dataDir, scanAllowlist = null, requireScanPassedForDownload = null } = {}) {
    this.dataDir = dataDir;
    this.legacyStorePath = path.join(this.dataDir, "store.json");
    this.legacyPackageDir = path.join(this.dataDir, "packages");
    this.dbPath = path.join(this.dataDir, "marketplace.sqlite");
    this.db = new SqliteFileDb({
      filePath: this.dbPath,
      migrationsDir: path.join(__dirname, "db", "migrations"),
    });

    const envAllowlist = parseCsvSet(process.env.MARKETPLACE_SCAN_ALLOWLIST);
    this.scanAllowlist = scanAllowlist ? new Set(scanAllowlist) : envAllowlist;
    this.requireScanPassedForDownload =
      requireScanPassedForDownload !== null && requireScanPassedForDownload !== undefined
        ? Boolean(requireScanPassedForDownload)
        : process.env.MARKETPLACE_REQUIRE_SCAN_PASSED === "1";
  }

  async init() {
    await this.db.getDb();
    await this._maybeMigrateLegacyStore();
  }

  async _maybeMigrateLegacyStore() {
    let legacyRaw = null;
    try {
      legacyRaw = await fs.readFile(this.legacyStorePath, "utf8");
    } catch {
      return;
    }

    const { hasPublishers, hasExtensions } = await this.db.withRead((db) => {
      const publishers = db.exec("SELECT publisher FROM publishers LIMIT 1")[0]?.values?.length > 0;
      const extensions = db.exec("SELECT id FROM extensions LIMIT 1")[0]?.values?.length > 0;
      return { hasPublishers: publishers, hasExtensions: extensions };
    });
    if (hasPublishers || hasExtensions) return;

    const legacy = JSON.parse(legacyRaw);
    if (!legacy || typeof legacy !== "object") return;

    const now = new Date().toISOString();
    const publisherKeyByPublisher = new Map();

    await this.db.withTransaction(async (tx) => {
      for (const publisherRecord of Object.values(legacy.publishers || {})) {
        if (!publisherRecord?.publisher) continue;
        const normalizedKeyPem = normalizePublicKeyPem(publisherRecord.publicKeyPem);
        let keyId = null;
        try {
          keyId = normalizedKeyPem ? publisherKeyIdFromPublicKeyPem(normalizedKeyPem) : null;
        } catch {
          keyId = null;
        }
        tx.run(
          `INSERT INTO publishers (publisher, token_sha256, public_key_pem, verified, created_at)
           VALUES (?, ?, ?, ?, ?)
           ON CONFLICT(publisher) DO UPDATE SET
             token_sha256 = excluded.token_sha256,
             public_key_pem = excluded.public_key_pem,
             verified = excluded.verified`,
          [
            publisherRecord.publisher,
            publisherRecord.tokenSha256,
            normalizedKeyPem,
            publisherRecord.verified ? 1 : 0,
            publisherRecord.createdAt || now,
          ]
        );

        if (keyId) {
          // Guard against reusing the same key id across multiple publishers.
          const existingKeyStmt = tx.prepare(`SELECT publisher FROM publisher_keys WHERE id = ? LIMIT 1`);
          existingKeyStmt.bind([keyId]);
          if (existingKeyStmt.step()) {
            const row = existingKeyStmt.getAsObject();
            const existingPublisher = String(row.publisher || "");
            existingKeyStmt.free();
            if (existingPublisher && existingPublisher !== publisherRecord.publisher) {
              throw new Error(`Signing key id already registered to a different publisher: ${keyId}`);
            }
          } else {
            existingKeyStmt.free();
          }

          publisherKeyByPublisher.set(publisherRecord.publisher, { id: keyId, publicKeyPem: normalizedKeyPem });
          tx.run(
            `INSERT INTO publisher_keys (id, publisher, public_key_pem, created_at, revoked, revoked_at, is_primary)
             VALUES (?, ?, ?, ?, 0, NULL, 1)
             ON CONFLICT(id) DO UPDATE SET
               public_key_pem = excluded.public_key_pem`,
            [keyId, publisherRecord.publisher, normalizedKeyPem, publisherRecord.createdAt || now]
          );
        }
      }

      for (const extRecord of Object.values(legacy.extensions || {})) {
        if (!extRecord?.id || !extRecord?.publisher) continue;
        tx.run(
          `INSERT INTO extensions
            (id, name, display_name, publisher, description, categories_json, tags_json, screenshots_json,
             verified, featured, deprecated, blocked, malicious, download_count, created_at, updated_at)
           VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 0, 0, 0, ?, ?, ?)
           ON CONFLICT(id) DO UPDATE SET
             name = excluded.name,
             display_name = excluded.display_name,
             description = excluded.description,
             categories_json = excluded.categories_json,
             tags_json = excluded.tags_json,
             screenshots_json = excluded.screenshots_json,
             verified = excluded.verified,
             featured = excluded.featured,
             download_count = excluded.download_count,
             updated_at = excluded.updated_at`,
          [
            extRecord.id,
            extRecord.name,
            extRecord.displayName,
            extRecord.publisher,
            extRecord.description || "",
            JSON.stringify(extRecord.categories || []),
            JSON.stringify(extRecord.tags || []),
            JSON.stringify(extRecord.screenshots || []),
            extRecord.verified ? 1 : 0,
            extRecord.featured ? 1 : 0,
            Number(extRecord.downloadCount || 0),
            extRecord.createdAt || now,
            extRecord.updatedAt || extRecord.createdAt || now,
          ]
        );

        const versions = extRecord.versions || {};
        for (const v of Object.values(versions)) {
          if (!v?.version) continue;
          let absPackagePath = null;
          let storedPackagePath = null;
          try {
            absPackagePath = v.packagePath
              ? resolvePackagePath(this.dataDir, v.packagePath)
              : safeJoin(this.dataDir, packageRelPath(extRecord.id, v.version));
            storedPackagePath = path.relative(this.dataDir, absPackagePath).replace(/\\/g, "/");
            if (storedPackagePath.startsWith("..") || path.isAbsolute(storedPackagePath)) {
              storedPackagePath = null;
            }
          } catch {
            absPackagePath = null;
            storedPackagePath = null;
          }

          if (!absPackagePath || !storedPackagePath) continue;

          let pkgBytes = null;
          try {
            pkgBytes = await fs.readFile(absPackagePath);
          } catch {
            pkgBytes = null;
          }
          if (!pkgBytes) continue;

          let manifestJson = "{}";
          try {
            const bundle = await readExtensionPackageSafe(pkgBytes, { maxUncompressedBytes: 50 * 1024 * 1024 });
            manifestJson = safeJsonStringify(bundle.manifest);
          } catch {
            // Keep best-effort; package bytes are stored on disk for downloads.
          }

          const publisherKey = publisherKeyByPublisher.get(String(extRecord.publisher));

          tx.run(
            `INSERT INTO extension_versions
              (extension_id, version, sha256, signature_base64, manifest_json, readme, package_bytes, package_path, uploaded_at, yanked, yanked_at,
               signing_key_id, signing_public_key_pem)
             VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, 0, NULL, ?, ?)
             ON CONFLICT(extension_id, version) DO NOTHING`,
            [
              extRecord.id,
              v.version,
              v.sha256,
              v.signatureBase64,
              manifestJson,
              v.readme || "",
              Buffer.alloc(0),
              storedPackagePath,
              v.uploadedAt || now,
              publisherKey ? String(publisherKey.id) : null,
              publisherKey ? String(publisherKey.publicKeyPem) : null,
            ]
          );
        }
      }

      const audit = tx.prepare(
        `INSERT INTO audit_log (id, actor, action, extension_id, version, ip, details_json, created_at)
         VALUES (?, ?, ?, ?, ?, ?, ?, ?)`
      );
      audit.run([randomId(), "system", "legacy.migration", null, null, null, "{}", now]);
      audit.free();
    });

    // Keep a backup for debugging, but prevent the legacy JSON store from being reused accidentally.
    try {
      await fs.rename(this.legacyStorePath, `${this.legacyStorePath}.bak`);
    } catch {
      // ignore
    }
  }

  async registerPublisher({ publisher, tokenSha256, publicKeyPem, verified = false }) {
    if (!publisher || typeof publisher !== "string") throw new Error("publisher is required");
    if (!tokenSha256 || typeof tokenSha256 !== "string") throw new Error("tokenSha256 is required");
    if (!publicKeyPem || typeof publicKeyPem !== "string") throw new Error("publicKeyPem is required");

    const now = new Date().toISOString();
    const normalizedPublicKeyPem = normalizePublicKeyPem(publicKeyPem);
    const keyId = publisherKeyIdFromPublicKeyPem(normalizedPublicKeyPem);

    await this.db.withTransaction((db) => {
      // Preserve the previous primary key (if any) before we overwrite publishers.public_key_pem.
      // This is required so that rotating keys doesn't "forget" the old key for already-published
      // extension versions when upgrading from a schema that only stored one key.
      const existingPublisherStmt = db.prepare(
        `SELECT public_key_pem, created_at
         FROM publishers
         WHERE publisher = ?
         LIMIT 1`
      );
      existingPublisherStmt.bind([publisher]);
      if (existingPublisherStmt.step()) {
        const existing = existingPublisherStmt.getAsObject();
        existingPublisherStmt.free();
        const existingPem = existing.public_key_pem ? normalizePublicKeyPem(String(existing.public_key_pem)) : "";
        if (existingPem) {
          let existingKeyId = null;
          try {
            existingKeyId = publisherKeyIdFromPublicKeyPem(existingPem);
          } catch {
            existingKeyId = null;
          }

          if (existingKeyId) {
            // Guard against reusing the same key id across multiple publishers.
            const existingKeyStmt = db.prepare(`SELECT publisher FROM publisher_keys WHERE id = ? LIMIT 1`);
            existingKeyStmt.bind([existingKeyId]);
            if (existingKeyStmt.step()) {
              const row = existingKeyStmt.getAsObject();
              const existingPublisher = String(row.publisher || "");
              existingKeyStmt.free();
              if (existingPublisher && existingPublisher !== publisher) {
                throw new Error(`Signing key id already registered to a different publisher: ${existingKeyId}`);
              }
            } else {
              existingKeyStmt.free();
            }

            db.run(
              `INSERT INTO publisher_keys (id, publisher, public_key_pem, created_at, revoked, revoked_at, is_primary)
               VALUES (?, ?, ?, ?, 0, NULL, 0)
               ON CONFLICT(id) DO UPDATE SET
                 public_key_pem = excluded.public_key_pem`,
              [existingKeyId, publisher, existingPem, existing.created_at ? String(existing.created_at) : now]
            );

            // Best-effort backfill: versions published before key history existed all used the
            // then-current publisher key. Capture it so downloads can include X-Publisher-Key-Id.
            db.run(
              `UPDATE extension_versions
               SET signing_key_id = ?, signing_public_key_pem = ?
               WHERE signing_key_id IS NULL
                 AND extension_id IN (SELECT id FROM extensions WHERE publisher = ?)`,
              [existingKeyId, existingPem, publisher]
            );
          }
        }
      } else {
        existingPublisherStmt.free();
      }

      db.run(
        `
          INSERT INTO publishers (publisher, token_sha256, public_key_pem, verified, created_at)
          VALUES (?, ?, ?, ?, ?)
          ON CONFLICT(publisher) DO UPDATE SET
            token_sha256 = excluded.token_sha256,
            public_key_pem = excluded.public_key_pem,
            verified = excluded.verified
        `,
        [publisher, tokenSha256, normalizedPublicKeyPem, verified ? 1 : 0, now]
      );

      // Guard against reusing the same key id across multiple publishers.
      const existingKeyStmt = db.prepare(`SELECT publisher FROM publisher_keys WHERE id = ? LIMIT 1`);
      existingKeyStmt.bind([keyId]);
      if (existingKeyStmt.step()) {
        const row = existingKeyStmt.getAsObject();
        const existingPublisher = String(row.publisher || "");
        existingKeyStmt.free();
        if (existingPublisher && existingPublisher !== publisher) {
          throw new Error(`Signing key id already registered to a different publisher: ${keyId}`);
        }
      } else {
        existingKeyStmt.free();
      }

      db.run(
        `INSERT INTO publisher_keys (id, publisher, public_key_pem, created_at, revoked, revoked_at, is_primary)
         VALUES (?, ?, ?, ?, 0, NULL, 0)
         ON CONFLICT(id) DO UPDATE SET
           public_key_pem = excluded.public_key_pem`,
        [keyId, publisher, normalizedPublicKeyPem, now]
      );

      // Update primary key for this publisher (do not delete historical keys).
      const revokedStmt = db.prepare(`SELECT revoked FROM publisher_keys WHERE id = ? LIMIT 1`);
      revokedStmt.bind([keyId]);
      const isRevoked = revokedStmt.step() ? Boolean(revokedStmt.getAsObject().revoked) : false;
      revokedStmt.free();
      if (isRevoked) {
        throw new Error("Cannot set a revoked signing key as primary");
      }
      db.run(`UPDATE publisher_keys SET is_primary = 0 WHERE publisher = ?`, [publisher]);
      db.run(`UPDATE publisher_keys SET is_primary = 1 WHERE id = ?`, [keyId]);

      const audit = db.prepare(
        `INSERT INTO audit_log (id, actor, action, extension_id, version, ip, details_json, created_at)
         VALUES (?, ?, ?, ?, ?, ?, ?, ?)`
      );
      audit.run([
        randomId(),
        "admin",
        "publisher.register",
        null,
        null,
        null,
        safeJsonStringify({ publisher, keyId }),
        now,
      ]);
      audit.free();
    });
  }

  async getPublisherByTokenSha256(tokenSha256) {
    return this.db.withRead((db) => {
      const stmt = db.prepare(
        `SELECT publisher, token_sha256, public_key_pem, verified, revoked, revoked_at, created_at
         FROM publishers WHERE token_sha256 = ? LIMIT 1`
      );
      stmt.bind([tokenSha256]);
      if (!stmt.step()) {
        stmt.free();
        return null;
      }
      const row = stmt.getAsObject();
      stmt.free();
      return {
        publisher: row.publisher,
        tokenSha256: row.token_sha256,
        publicKeyPem: row.public_key_pem,
        verified: Boolean(row.verified),
        revoked: Boolean(row.revoked),
        revokedAt: row.revoked_at ? String(row.revoked_at) : null,
        createdAt: row.created_at,
      };
    });
  }

  async getPublisher(publisher) {
    return this.db.withRead((db) => {
      const stmt = db.prepare(
        `SELECT publisher, token_sha256, public_key_pem, verified, revoked, revoked_at, created_at
         FROM publishers WHERE publisher = ? LIMIT 1`
      );
      stmt.bind([publisher]);
      if (!stmt.step()) {
        stmt.free();
        return null;
      }
      const row = stmt.getAsObject();
      stmt.free();
      return {
        publisher: row.publisher,
        tokenSha256: row.token_sha256,
        publicKeyPem: row.public_key_pem,
        verified: Boolean(row.verified),
        revoked: Boolean(row.revoked),
        revokedAt: row.revoked_at ? String(row.revoked_at) : null,
        createdAt: row.created_at,
      };
    });
  }

  async getPublisherKeys(publisher, { includeRevoked = true } = {}) {
    return this.db.withRead((db) => {
      const condition = includeRevoked ? "" : "AND revoked = 0";
      const stmt = db.prepare(
        `SELECT id, public_key_pem, created_at, revoked, revoked_at, is_primary
         FROM publisher_keys
         WHERE publisher = ? ${condition}
         ORDER BY is_primary DESC, created_at ASC`
      );
      stmt.bind([publisher]);
      const out = [];
      while (stmt.step()) {
        const row = stmt.getAsObject();
        out.push({
          id: String(row.id),
          publicKeyPem: String(row.public_key_pem),
          createdAt: String(row.created_at),
          revoked: Boolean(row.revoked),
          revokedAt: row.revoked_at ? String(row.revoked_at) : null,
          isPrimary: Boolean(row.is_primary),
        });
      }
      stmt.free();
      return out;
    });
  }

  async revokePublisherKey(publisher, keyId, { actor = "admin", ip = null } = {}) {
    const now = new Date().toISOString();

    await this.db.withTransaction((db) => {
      const stmt = db.prepare(
        `SELECT id, revoked, is_primary
         FROM publisher_keys
         WHERE publisher = ? AND id = ? LIMIT 1`
      );
      stmt.bind([publisher, keyId]);
      if (!stmt.step()) {
        stmt.free();
        throw new Error("Publisher key not found");
      }
      const row = stmt.getAsObject();
      stmt.free();

      if (!row.revoked) {
        db.run(`UPDATE publisher_keys SET revoked = 1, revoked_at = ?, is_primary = 0 WHERE id = ?`, [now, keyId]);
      } else {
        // Ensure we still clear primary in case it was left set by older schemas.
        db.run(`UPDATE publisher_keys SET is_primary = 0 WHERE id = ?`, [keyId]);
      }

      // If the primary key was revoked, promote the newest non-revoked key to primary (if any).
      if (row.is_primary) {
        const nextStmt = db.prepare(
          `SELECT id, public_key_pem
           FROM publisher_keys
           WHERE publisher = ? AND revoked = 0
           ORDER BY created_at DESC
           LIMIT 1`
        );
        nextStmt.bind([publisher]);
        if (nextStmt.step()) {
          const next = nextStmt.getAsObject();
          nextStmt.free();
          db.run(`UPDATE publisher_keys SET is_primary = 0 WHERE publisher = ?`, [publisher]);
          db.run(`UPDATE publisher_keys SET is_primary = 1 WHERE id = ?`, [String(next.id)]);
          db.run(`UPDATE publishers SET public_key_pem = ? WHERE publisher = ?`, [String(next.public_key_pem), publisher]);
        } else {
          nextStmt.free();
        }
      }

      const audit = db.prepare(
        `INSERT INTO audit_log (id, actor, action, extension_id, version, ip, details_json, created_at)
         VALUES (?, ?, ?, ?, ?, ?, ?, ?)`
      );
      audit.run([
        randomId(),
        actor,
        "publisher.key.revoke",
        null,
        null,
        ip,
        safeJsonStringify({ publisher, keyId }),
        now,
      ]);
      audit.free();
    });

    return { publisher, keyId, revoked: true };
  }

  async rotatePublisherToken(publisher, { token = null, actor = "admin", ip = null } = {}) {
    if (!publisher || typeof publisher !== "string") throw new Error("publisher is required");

    const newToken =
      token && typeof token === "string"
        ? token
        : crypto
            .randomBytes(32)
            .toString("base64")
            .replace(/\+/g, "-")
            .replace(/\//g, "_")
            .replace(/=+$/g, "");
    const tokenSha256 = sha256Utf8(newToken);
    const now = nowIso();

    await this.db.withTransaction((db) => {
      const stmt = db.prepare(`SELECT publisher FROM publishers WHERE publisher = ? LIMIT 1`);
      stmt.bind([publisher]);
      const exists = stmt.step();
      stmt.free();
      if (!exists) throw new Error("Publisher not found");

      db.run(`UPDATE publishers SET token_sha256 = ? WHERE publisher = ?`, [tokenSha256, publisher]);

      const audit = db.prepare(
        `INSERT INTO audit_log (id, actor, action, extension_id, version, ip, details_json, created_at)
         VALUES (?, ?, ?, ?, ?, ?, ?, ?)`
      );
      audit.run([randomId(), actor, "publisher.rotate_token", null, null, ip, safeJsonStringify({ publisher }), now]);
      audit.free();
    });

    return { publisher, token: newToken, tokenSha256 };
  }

  async rotatePublisherPublicKey(
    publisher,
    { publicKeyPem, overlapMs = null, actor = "admin", ip = null } = {}
  ) {
    if (!publisher || typeof publisher !== "string") throw new Error("publisher is required");
    if (!publicKeyPem || typeof publicKeyPem !== "string") throw new Error("publicKeyPem is required");

    const now = nowIso();
    const normalizedPublicKeyPem = normalizePublicKeyPem(publicKeyPem);
    const keyId = publisherKeyIdFromPublicKeyPem(normalizedPublicKeyPem);

    await this.db.withTransaction((db) => {
      const existingPublisherStmt = db.prepare(
        `SELECT public_key_pem, created_at
         FROM publishers
         WHERE publisher = ?
         LIMIT 1`
      );
      existingPublisherStmt.bind([publisher]);
      if (!existingPublisherStmt.step()) {
        existingPublisherStmt.free();
        throw new Error("Publisher not found");
      }
      const existing = existingPublisherStmt.getAsObject();
      existingPublisherStmt.free();

      const existingPem = existing.public_key_pem ? normalizePublicKeyPem(String(existing.public_key_pem)) : "";
      const existingCreatedAt = existing.created_at ? String(existing.created_at) : now;
      if (existingPem) {
        let existingKeyId = null;
        try {
          existingKeyId = publisherKeyIdFromPublicKeyPem(existingPem);
        } catch {
          existingKeyId = null;
        }

        if (existingKeyId) {
          const existingKeyStmt = db.prepare(`SELECT publisher FROM publisher_keys WHERE id = ? LIMIT 1`);
          existingKeyStmt.bind([existingKeyId]);
          if (existingKeyStmt.step()) {
            const row = existingKeyStmt.getAsObject();
            const existingPublisher = String(row.publisher || "");
            existingKeyStmt.free();
            if (existingPublisher && existingPublisher !== publisher) {
              throw new Error(`Signing key id already registered to a different publisher: ${existingKeyId}`);
            }
          } else {
            existingKeyStmt.free();
          }

          db.run(
            `INSERT INTO publisher_keys (id, publisher, public_key_pem, created_at, revoked, revoked_at, is_primary)
             VALUES (?, ?, ?, ?, 0, NULL, 0)
             ON CONFLICT(id) DO UPDATE SET
               public_key_pem = excluded.public_key_pem`,
            [existingKeyId, publisher, existingPem, existingCreatedAt]
          );

          db.run(
            `UPDATE extension_versions
             SET signing_key_id = ?, signing_public_key_pem = ?
             WHERE signing_key_id IS NULL
               AND extension_id IN (SELECT id FROM extensions WHERE publisher = ?)`,
            [existingKeyId, existingPem, publisher]
          );
        }
      }

      db.run(`UPDATE publishers SET public_key_pem = ? WHERE publisher = ?`, [normalizedPublicKeyPem, publisher]);

      const existingKeyStmt = db.prepare(`SELECT publisher FROM publisher_keys WHERE id = ? LIMIT 1`);
      existingKeyStmt.bind([keyId]);
      if (existingKeyStmt.step()) {
        const row = existingKeyStmt.getAsObject();
        const existingPublisher = String(row.publisher || "");
        existingKeyStmt.free();
        if (existingPublisher && existingPublisher !== publisher) {
          throw new Error(`Signing key id already registered to a different publisher: ${keyId}`);
        }
      } else {
        existingKeyStmt.free();
      }

      db.run(
        `INSERT INTO publisher_keys (id, publisher, public_key_pem, created_at, revoked, revoked_at, is_primary)
         VALUES (?, ?, ?, ?, 0, NULL, 0)
         ON CONFLICT(id) DO UPDATE SET
           public_key_pem = excluded.public_key_pem`,
        [keyId, publisher, normalizedPublicKeyPem, now]
      );

      const revokedStmt = db.prepare(`SELECT revoked FROM publisher_keys WHERE id = ? LIMIT 1`);
      revokedStmt.bind([keyId]);
      const isRevoked = revokedStmt.step() ? Boolean(revokedStmt.getAsObject().revoked) : false;
      revokedStmt.free();
      if (isRevoked) {
        throw new Error("Cannot set a revoked signing key as primary");
      }

      db.run(`UPDATE publisher_keys SET is_primary = 0 WHERE publisher = ?`, [publisher]);
      db.run(`UPDATE publisher_keys SET is_primary = 1 WHERE id = ?`, [keyId]);

      const audit = db.prepare(
        `INSERT INTO audit_log (id, actor, action, extension_id, version, ip, details_json, created_at)
         VALUES (?, ?, ?, ?, ?, ?, ?, ?)`
      );
      audit.run([
        randomId(),
        actor,
        "publisher.rotate_key",
        null,
        null,
        ip,
        safeJsonStringify({ publisher, keyId, overlapMs }),
        now,
      ]);
      audit.free();
    });

    return { publisher, keyId, publicKeyPem: normalizedPublicKeyPem, overlapMs };
  }

  async revokePublisher(publisher, { revoked = true, actor = "admin", ip = null } = {}) {
    if (!publisher || typeof publisher !== "string") throw new Error("publisher is required");

    const now = nowIso();
    await this.db.withTransaction((db) => {
      const stmt = db.prepare(`SELECT publisher FROM publishers WHERE publisher = ? LIMIT 1`);
      stmt.bind([publisher]);
      const exists = stmt.step();
      stmt.free();
      if (!exists) throw new Error("Publisher not found");

      db.run(`UPDATE publishers SET revoked = ?, revoked_at = ? WHERE publisher = ?`, [
        revoked ? 1 : 0,
        revoked ? now : null,
        publisher,
      ]);

      const audit = db.prepare(
        `INSERT INTO audit_log (id, actor, action, extension_id, version, ip, details_json, created_at)
         VALUES (?, ?, ?, ?, ?, ?, ?, ?)`
      );
      audit.run([randomId(), actor, "publisher.revoke", null, null, ip, safeJsonStringify({ publisher, revoked }), now]);
      audit.free();
    });

    return this.getPublisher(publisher);
  }

  async setExtensionFlags(id, { verified, featured, deprecated, blocked, malicious }, { actor = "admin", ip = null } = {}) {
    const now = new Date().toISOString();
    await this.db.withTransaction((db) => {
      const existingStmt = db.prepare("SELECT id FROM extensions WHERE id = ? LIMIT 1");
      existingStmt.bind([id]);
      const exists = existingStmt.step();
      existingStmt.free();
      if (!exists) throw new Error("Extension not found");

      const patch = [];
      const params = [];

      function add(field, value) {
        patch.push(`${field} = ?`);
        params.push(value);
      }

      if (verified !== undefined) add("verified", verified ? 1 : 0);
      if (featured !== undefined) add("featured", featured ? 1 : 0);
      if (deprecated !== undefined) add("deprecated", deprecated ? 1 : 0);
      if (blocked !== undefined) add("blocked", blocked ? 1 : 0);
      if (malicious !== undefined) add("malicious", malicious ? 1 : 0);

      if (patch.length > 0) {
        add("updated_at", now);
        params.push(id);
        db.run(`UPDATE extensions SET ${patch.join(", ")} WHERE id = ?`, params);
      }

      const audit = db.prepare(
        `INSERT INTO audit_log (id, actor, action, extension_id, version, ip, details_json, created_at)
         VALUES (?, ?, ?, ?, ?, ?, ?, ?)`
      );
      audit.run([
        randomId(),
        actor,
        "extension.flags",
        id,
        null,
        ip,
        safeJsonStringify({ verified, featured, deprecated, blocked, malicious }),
        now,
      ]);
      audit.free();
    });

    return this.getExtension(id, { includeHidden: true });
  }

  async publishExtension({ publisher, packageBytes, signatureBase64 }) {
    const publisherRecord = await this.getPublisher(publisher);
    if (!publisherRecord) throw new Error(`Unknown publisher: ${publisher}`);
    if (publisherRecord.revoked) {
      throw new Error("Publisher revoked");
    }

    const now = new Date().toISOString();

    // Allow publisher signing key rotation: verify against any non-revoked key for the publisher,
    // and persist which key succeeded so old versions remain verifiable after rotation.
    //
    // Important: if the publisher has *no* keys recorded in publisher_keys (e.g. upgraded from an
    // older schema), we fall back to publishers.public_key_pem as the single known active key.
    // However, if keys exist but are all revoked, we must *not* resurrect the legacy key.
    let publisherKeys = await this.getPublisherKeys(publisher, { includeRevoked: false });
    if (publisherKeys.length === 0) {
      const allKeys = await this.getPublisherKeys(publisher, { includeRevoked: true });
      if (allKeys.length > 0) {
        throw new Error("Package signature verification failed (all publisher signing keys are revoked)");
      }

      if (publisherRecord.publicKeyPem) {
        const fallbackPem = normalizePublicKeyPem(publisherRecord.publicKeyPem);
        if (fallbackPem) {
          const keyId = publisherKeyIdFromPublicKeyPem(fallbackPem);
          await this.db.withTransaction((db) => {
            const existingKeyStmt = db.prepare(`SELECT publisher FROM publisher_keys WHERE id = ? LIMIT 1`);
            existingKeyStmt.bind([keyId]);
            if (existingKeyStmt.step()) {
              const row = existingKeyStmt.getAsObject();
              const existingPublisher = String(row.publisher || "");
              existingKeyStmt.free();
              if (existingPublisher && existingPublisher !== publisher) {
                throw new Error(`Signing key id already registered to a different publisher: ${keyId}`);
              }
            } else {
              existingKeyStmt.free();
            }

            db.run(`UPDATE publisher_keys SET is_primary = 0 WHERE publisher = ?`, [publisher]);
            db.run(
              `INSERT INTO publisher_keys (id, publisher, public_key_pem, created_at, revoked, revoked_at, is_primary)
               VALUES (?, ?, ?, ?, 0, NULL, 1)
               ON CONFLICT(id) DO UPDATE SET
                 public_key_pem = excluded.public_key_pem,
                 revoked = 0,
                 revoked_at = NULL,
                 is_primary = 1`,
              [keyId, publisher, fallbackPem, now]
            );

            db.run(
              `UPDATE extension_versions
               SET signing_key_id = ?, signing_public_key_pem = ?
               WHERE signing_key_id IS NULL
                 AND extension_id IN (SELECT id FROM extensions WHERE publisher = ?)`,
              [keyId, fallbackPem, publisher]
            );
          });
          publisherKeys = await this.getPublisherKeys(publisher, { includeRevoked: false });
        }
      }
    }
    if (publisherKeys.length === 0) {
      throw new Error("Package signature verification failed (publisher has no signing keys)");
    }

    const MAX_COMPRESSED_BYTES = 10 * 1024 * 1024;
    const MAX_UNCOMPRESSED_BYTES = 50 * 1024 * 1024;
    const MAX_UNPACKED_BYTES = 50 * 1024 * 1024;
    const MAX_FILES = 500;
    const MAX_SINGLE_FILE_BYTES = 10 * 1024 * 1024;

    const formatVersion = detectExtensionPackageFormatVersion(packageBytes);

    /** @type {any} */
    let manifest = null;
    /** @type {{path: string, sha256: string, size: number}[]} */
    let fileRecords = [];
    let fileCount = 0;
    let unpackedSize = 0;
    let readme = "";
    let storedSignatureBase64 = null;
    /** @type {string | null} */
    let signingKeyId = null;
    /** @type {string | null} */
    let signingPublicKeyPem = null;

    if (formatVersion === 1) {
      if (!signatureBase64) throw new Error("signatureBase64 is required for v1 extension packages");
      if (packageBytes.length > MAX_COMPRESSED_BYTES) {
        throw new Error("Package exceeds maximum compressed size");
      }

      let matchingKey = null;
      for (const key of publisherKeys) {
        if (verifyBytesSignature(packageBytes, signatureBase64, key.publicKeyPem)) {
          matchingKey = key;
          break;
        }
      }
      if (!matchingKey) throw new Error("Package signature verification failed");
      storedSignatureBase64 = String(signatureBase64);
      signingKeyId = String(matchingKey.id);
      signingPublicKeyPem = String(matchingKey.publicKeyPem);

      const bundle = await readExtensionPackageSafe(packageBytes, { maxUncompressedBytes: MAX_UNCOMPRESSED_BYTES });
      manifest = bundle.manifest;

      if (bundle.files.length > MAX_FILES) throw new Error("Extension package contains too many files");

      const seen = new Set();
      const seenFolded = new Set();
      let packageJsonBytes = null;
      for (const file of bundle.files) {
        if (!file?.path || typeof file.path !== "string" || typeof file.dataBase64 !== "string") {
          throw new Error("Invalid file entry in extension package");
        }
        if (!isAllowedFilePath(file.path)) {
          throw new Error(`Disallowed file type in extension package: ${file.path}`);
        }

        const normalizedPath = file.path.replace(/\\/g, "/");
        if (seen.has(normalizedPath)) throw new Error(`Duplicate file in extension package: ${normalizedPath}`);
        seen.add(normalizedPath);
        const folded = normalizedPath.toLowerCase();
        if (seenFolded.has(folded)) {
          throw new Error(`Duplicate file path in extension package (case-insensitive): ${normalizedPath}`);
        }
        seenFolded.add(folded);

        const bytes = Buffer.from(file.dataBase64, "base64");
        if (bytes.length > MAX_SINGLE_FILE_BYTES) {
          throw new Error(`Extension package contains oversized file: ${normalizedPath}`);
        }
        if (normalizedPath === "package.json") {
          packageJsonBytes = bytes;
        }
        unpackedSize += bytes.length;
        fileRecords.push({ path: normalizedPath, sha256: sha256(bytes), size: bytes.length });

        if (normalizedPath.toLowerCase() === "readme.md") {
          readme = bytes.toString("utf8");
        }
      }

      if (!packageJsonBytes) {
        throw new Error("Invalid extension package: missing package.json file");
      }
      let packageJson = null;
      try {
        packageJson = JSON.parse(packageJsonBytes.toString("utf8"));
      } catch {
        throw new Error("Invalid extension package: package.json is not valid JSON");
      }
      if (!packageJson || typeof packageJson !== "object" || Array.isArray(packageJson)) {
        throw new Error("Invalid extension package: package.json must be a JSON object");
      }
      if (!canonicalJsonBytes(packageJson).equals(canonicalJsonBytes(manifest))) {
        throw new Error("Invalid extension package: package.json does not match embedded manifest");
      }

      fileRecords.sort((a, b) => (a.path < b.path ? -1 : a.path > b.path ? 1 : 0));
      fileCount = fileRecords.length;
      if (unpackedSize > MAX_UNPACKED_BYTES) {
        throw new Error("Package exceeds maximum uncompressed payload size");
      }
    } else if (formatVersion === 2) {
      /** @type {ReturnType<typeof verifyExtensionPackageV2> | null} */
      let verified = null;
      let matchingKey = null;
      let sawSignatureFailure = false;
      for (const key of publisherKeys) {
        try {
          verified = verifyExtensionPackageV2(packageBytes, key.publicKeyPem);
          matchingKey = key;
          break;
        } catch (error) {
          const message = String(error?.message || error);
          if (message.toLowerCase().includes("signature verification failed")) {
            sawSignatureFailure = true;
            continue;
          }
          throw error;
        }
      }
      if (!verified || !matchingKey) {
        if (sawSignatureFailure) {
          throw new Error("Package signature verification failed");
        }
        throw new Error("Package signature verification failed");
      }

      manifest = verified.manifest;
      storedSignatureBase64 = verified.signatureBase64;
      fileRecords = verified.files;
      fileCount = verified.fileCount;
      unpackedSize = verified.unpackedSize;
      readme = verified.readme || "";
      signingKeyId = String(matchingKey.id);
      signingPublicKeyPem = String(matchingKey.publicKeyPem);

      if (fileCount > MAX_FILES) throw new Error("Extension package contains too many files");
      if (unpackedSize > MAX_UNPACKED_BYTES) {
        throw new Error("Package exceeds maximum uncompressed payload size");
      }

      for (const file of fileRecords) {
        if (!isAllowedFilePath(file.path)) {
          throw new Error(`Disallowed file type in extension package: ${file.path}`);
        }
        if (typeof file.size === "number" && file.size > MAX_SINGLE_FILE_BYTES) {
          throw new Error(`Extension package contains oversized file: ${file.path}`);
        }
      }
    } else {
      throw new Error(`Unsupported extension package formatVersion: ${formatVersion}`);
    }

    manifest = validateManifest(manifest);
    validateEntrypoint(manifest, fileRecords);

    if (manifest.publisher !== publisher) {
      throw new Error("Manifest publisher does not match authenticated publisher");
    }

    const id = extensionIdFromManifest(manifest);
    const version = manifest.version;

    const pkgSha = sha256(packageBytes);
    const filesJson = safeJsonStringify(fileRecords);
    const storedPackagePath = packageRelPath(id, version);

    await ensurePackageFile({
      dataDir: this.dataDir,
      relPath: storedPackagePath,
      bytes: packageBytes,
      expectedSha256: pkgSha,
    });

    const categories = normalizeStringArray(manifest.categories);
    const tags = normalizeStringArray(manifest.tags);
    const screenshots = normalizeStringArray(manifest.screenshots);

    await this.db.withTransaction((db) => {
      const existingStmt = db.prepare(
        `SELECT id, verified, featured, deprecated, blocked, malicious, download_count, created_at
         FROM extensions WHERE id = ? LIMIT 1`
      );
      existingStmt.bind([id]);
      const hasExisting = existingStmt.step();
      const existing = hasExisting ? existingStmt.getAsObject() : null;
      existingStmt.free();

      if (!hasExisting) {
        db.run(
          `INSERT INTO extensions
            (id, name, display_name, publisher, description, categories_json, tags_json, screenshots_json,
             verified, featured, deprecated, blocked, malicious, download_count, created_at, updated_at)
           VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 0, 0, 0, 0, ?, ?)`,
          [
            id,
            manifest.name,
            manifest.displayName || manifest.name,
            manifest.publisher,
            manifest.description || "",
            JSON.stringify(categories),
            JSON.stringify(tags),
            JSON.stringify(screenshots),
            publisherRecord.verified ? 1 : 0,
            0,
            now,
            now,
          ]
        );
      } else {
        db.run(
          `UPDATE extensions SET
            name = ?,
            display_name = ?,
            description = ?,
            categories_json = ?,
            tags_json = ?,
            screenshots_json = ?,
            updated_at = ?
          WHERE id = ?`,
          [
            manifest.name,
            manifest.displayName || manifest.name,
            manifest.description || "",
            JSON.stringify(categories),
            JSON.stringify(tags),
            JSON.stringify(screenshots),
            now,
            id,
          ]
        );
      }

      // Unique constraint on (extension_id, version) ensures concurrency safety.
      db.run(
        `INSERT INTO extension_versions
          (extension_id, version, sha256, signature_base64, manifest_json, readme, package_bytes, package_path,
           uploaded_at, yanked, yanked_at, format_version, file_count, unpacked_size, files_json,
           signing_key_id, signing_public_key_pem)
         VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, 0, NULL, ?, ?, ?, ?, ?, ?)`,
        [
          id,
          version,
          pkgSha,
          String(storedSignatureBase64),
          safeJsonStringify(manifest),
          readme,
          Buffer.alloc(0),
          storedPackagePath,
          now,
          formatVersion,
          fileCount,
          unpackedSize,
          filesJson,
          signingKeyId,
          signingPublicKeyPem,
        ]
      );

      db.run(
        `INSERT OR IGNORE INTO package_scans
          (extension_id, version, status, findings_json, scanned_at)
         VALUES (?, ?, ?, '[]', NULL)`,
        [id, version, PACKAGE_SCAN_STATUS.PENDING]
      );
    }).catch((err) => {
      if (String(err?.message || "").includes("UNIQUE constraint failed: extension_versions.extension_id, extension_versions.version")) {
        throw new Error("That extension version is already published");
      }
      throw err;
    });

    try {
      await this.scanExtensionVersion(id, version, { packageBytes });
    } catch {
    }

    return { id, version };
  }

  async getPackageScan(extensionId, version) {
    return this.db.withRead((db) => {
      const stmt = db.prepare(
        `SELECT status, findings_json, scanned_at
         FROM package_scans
         WHERE extension_id = ? AND version = ?
         LIMIT 1`
      );
      stmt.bind([extensionId, version]);
      if (!stmt.step()) {
        stmt.free();
        return null;
      }
      const row = stmt.getAsObject();
      stmt.free();

      let findings = null;
      try {
        findings = JSON.parse(String(row.findings_json || "[]"));
      } catch {
        findings = [];
      }

      return {
        extensionId,
        version,
        status: String(row.status),
        findings,
        scannedAt: row.scanned_at ? String(row.scanned_at) : null,
      };
    });
  }

  async rescanExtensionVersion(extensionId, version, { actor = "admin", ip = null } = {}) {
    const now = nowIso();
    await this.db.withTransaction((db) => {
      const existingStmt = db.prepare(
        `SELECT extension_id
         FROM extension_versions
         WHERE extension_id = ? AND version = ?
         LIMIT 1`
      );
      existingStmt.bind([extensionId, version]);
      const exists = existingStmt.step();
      existingStmt.free();
      if (!exists) throw new Error("Extension version not found");

      db.run(
        `INSERT OR IGNORE INTO package_scans
          (extension_id, version, status, findings_json, scanned_at)
         VALUES (?, ?, ?, '[]', NULL)`,
        [extensionId, version, PACKAGE_SCAN_STATUS.PENDING]
      );
      db.run(
        `UPDATE package_scans
         SET status = ?, findings_json = '[]', scanned_at = NULL
         WHERE extension_id = ? AND version = ?`,
        [PACKAGE_SCAN_STATUS.PENDING, extensionId, version]
      );

      const audit = db.prepare(
        `INSERT INTO audit_log (id, actor, action, extension_id, version, ip, details_json, created_at)
         VALUES (?, ?, ?, ?, ?, ?, ?, ?)`
      );
      audit.run([randomId(), actor, "package_scan.rescan", extensionId, version, ip, "{}", now]);
      audit.free();
    });

    await this.scanExtensionVersion(extensionId, version);
    return this.getPackageScan(extensionId, version);
  }

  async scanExtensionVersion(extensionId, version, { packageBytes = null } = {}) {
    const meta = await this.db.withRead((db) => {
      const stmt = db.prepare(
        `SELECT sha256, package_bytes, package_path
         FROM extension_versions
         WHERE extension_id = ? AND version = ?
         LIMIT 1`
      );
      stmt.bind([extensionId, version]);
      if (!stmt.step()) {
        stmt.free();
        return null;
      }
      const row = stmt.getAsObject();
      stmt.free();
      return {
        expectedSha256: String(row.sha256),
        packageBytes: row.package_bytes,
        packagePath: row.package_path ? String(row.package_path) : null,
      };
    });

    if (!meta) throw new Error("Extension version not found");

    /** @type {Buffer | null} */
    let bytes = packageBytes ? Buffer.from(packageBytes) : null;
    if (!bytes) {
      if (meta.packagePath) {
        const fullPath = resolvePackagePath(this.dataDir, meta.packagePath);
        try {
          bytes = await fs.readFile(fullPath);
        } catch {
          bytes = null;
        }
      } else {
        bytes = Buffer.from(meta.packageBytes);
      }
    }

    const allowlist = this.scanAllowlist || new Set();
    const findings = [];
    const actualSha = bytes ? sha256(bytes) : null;
    if (!bytes || bytes.length === 0) {
      findings.push({ id: "package.missing_bytes", message: "Package bytes missing on disk" });
    } else if (actualSha !== meta.expectedSha256) {
      findings.push({
        id: "package.sha256_mismatch",
        message: "Stored package sha256 does not match on-disk bytes",
        expected: meta.expectedSha256,
        actual: actualSha,
      });
    }

    let scanData = {
      formatVersion: null,
      fileRecords: [],
      fileCount: 0,
      unpackedSize: 0,
      findings: [],
    };

    if (bytes && bytes.length > 0) {
      try {
        scanData = await this._computePackageScan(bytes, { allowlist });
      } catch (err) {
        scanData = {
          formatVersion: null,
          fileRecords: [],
          fileCount: 0,
          unpackedSize: 0,
          findings: [{ id: "scan.error", message: String(err?.message || err) }],
        };
      }
    }

    const mergedFindings = [...findings, ...(scanData.findings || [])].filter((f) => !allowlist.has(f.id));
    const status = mergedFindings.length > 0 ? PACKAGE_SCAN_STATUS.FAILED : PACKAGE_SCAN_STATUS.PASSED;
    const scannedAt = nowIso();

    const filesJson = safeJsonStringify(scanData.fileRecords || []);
    const findingsJson = safeJsonStringify({
      formatVersion: scanData.formatVersion,
      fileCount: scanData.fileCount,
      unpackedSize: scanData.unpackedSize,
      findings: mergedFindings,
    });

    await this.db.withTransaction((db) => {
      if (Array.isArray(scanData.fileRecords) && scanData.fileRecords.length > 0) {
        db.run(
          `UPDATE extension_versions
           SET format_version = ?, file_count = ?, unpacked_size = ?, files_json = ?
           WHERE extension_id = ? AND version = ?`,
          [scanData.formatVersion || 1, scanData.fileCount || 0, scanData.unpackedSize || 0, filesJson, extensionId, version]
        );
      }

      db.run(
        `INSERT OR IGNORE INTO package_scans
          (extension_id, version, status, findings_json, scanned_at)
         VALUES (?, ?, ?, ?, ?)`,
        [extensionId, version, status, findingsJson, scannedAt]
      );
      db.run(
        `UPDATE package_scans
         SET status = ?, findings_json = ?, scanned_at = ?
         WHERE extension_id = ? AND version = ?`,
        [status, findingsJson, scannedAt, extensionId, version]
      );
    });

    return { extensionId, version, status, scannedAt, findings: mergedFindings };
  }

  async _computePackageScan(packageBytes, { allowlist }) {
    const findings = [];
    const formatVersion = detectExtensionPackageFormatVersion(packageBytes);

    /** @type {any} */
    let manifest = null;
    /** @type {{path: string, sha256: string, size: number}[]} */
    const fileRecords = [];
    /** @type {Map<string, Buffer>} */
    const files = new Map();
    let unpackedSize = 0;

    const MAX_UNCOMPRESSED_BYTES = 50 * 1024 * 1024;
    const MAX_UNPACKED_BYTES = 50 * 1024 * 1024;
    const MAX_FILES = 500;

    if (formatVersion === 1) {
      const bundle = await readExtensionPackageSafe(packageBytes, { maxUncompressedBytes: MAX_UNCOMPRESSED_BYTES });
      manifest = bundle.manifest;
      for (const file of bundle.files || []) {
        if (!file?.path || typeof file.path !== "string" || typeof file.dataBase64 !== "string") {
          continue;
        }
        const normalizedPath = file.path.replace(/\\/g, "/");
        const bytes = Buffer.from(file.dataBase64, "base64");
        unpackedSize += bytes.length;
        files.set(normalizedPath, bytes);
        fileRecords.push({ path: normalizedPath, sha256: sha256(bytes), size: bytes.length });
      }
    } else if (formatVersion === 2) {
      const parsed = readExtensionPackageV2(packageBytes);
      manifest = parsed.manifest;

      const checksums = parsed.checksums;
      const checksumEntries = new Map();
      if (checksums?.algorithm === "sha256" && typeof checksums.files === "object" && checksums.files) {
        for (const [relPath, entry] of Object.entries(checksums.files)) {
          checksumEntries.set(relPath, entry);
        }
      } else {
        findings.push({ id: "package.v2.invalid_checksums", message: "Invalid checksums.json" });
      }

      const expectedPaths = new Set(checksumEntries.keys());
      for (const [relPath, data] of parsed.files.entries()) {
        unpackedSize += data.length;
        files.set(relPath, data);
        const actualSha = sha256(data);
        fileRecords.push({ path: relPath, sha256: actualSha, size: data.length });

        const expected = checksumEntries.get(relPath);
        if (!expected) {
          findings.push({ id: "package.v2.missing_checksum", message: `checksums.json missing entry for ${relPath}` });
        } else {
          const expectedSha = typeof expected.sha256 === "string" ? expected.sha256.toLowerCase() : null;
          const expectedSize = expected.size;
          if (expectedSha && expectedSha !== actualSha) {
            findings.push({ id: "package.v2.checksum_mismatch", message: `Checksum mismatch for ${relPath}` });
          }
          if (typeof expectedSize === "number" && expectedSize !== data.length) {
            findings.push({ id: "package.v2.size_mismatch", message: `Size mismatch for ${relPath}` });
          }
        }
        expectedPaths.delete(relPath);
      }
      if (expectedPaths.size > 0) {
        findings.push({
          id: "package.v2.extra_checksums",
          message: `checksums.json contains extra entries: ${[...expectedPaths].join(", ")}`,
        });
      }
    } else {
      findings.push({ id: "package.unknown_format", message: `Unsupported extension package formatVersion: ${formatVersion}` });
    }

    fileRecords.sort((a, b) => (a.path < b.path ? -1 : a.path > b.path ? 1 : 0));

    if (fileRecords.length > MAX_FILES) {
      findings.push({ id: "package.too_many_files", message: `Extension package contains too many files (${fileRecords.length})` });
    }

    if (unpackedSize > MAX_UNPACKED_BYTES) {
      findings.push({ id: "package.unpacked_size", message: `Package exceeds maximum uncompressed payload size (${unpackedSize})` });
    }

    for (const record of fileRecords) {
      if (!isAllowedFilePath(record.path)) {
        findings.push({ id: "package.disallowed_path", message: `Disallowed file type in extension package: ${record.path}` });
      }
    }

    const packageJsonBytes = files.get("package.json");
    if (!packageJsonBytes) {
      findings.push({ id: "manifest.missing_package_json", message: "Missing package.json in extension package" });
    } else {
      try {
        const pkg = JSON.parse(packageJsonBytes.toString("utf8"));
        if (!canonicalJsonBytes(pkg).equals(canonicalJsonBytes(manifest))) {
          findings.push({ id: "manifest.package_json_mismatch", message: "package.json does not match embedded manifest" });
        }
      } catch (err) {
        findings.push({ id: "manifest.package_json_invalid", message: String(err?.message || err) });
      }
    }

    let validatedManifest = null;
    try {
      validatedManifest = validateManifest(manifest);
    } catch (err) {
      findings.push({ id: "manifest.invalid", message: String(err?.message || err) });
      validatedManifest = null;
    }

    if (validatedManifest) {
      try {
        validateEntrypoint(validatedManifest, fileRecords);
      } catch (err) {
        findings.push({ id: "manifest.entrypoint", message: String(err?.message || err) });
      }
    }

    for (const [relPath, data] of files.entries()) {
      const lower = relPath.toLowerCase();
      if (!(lower.endsWith(".js") || lower.endsWith(".cjs") || lower.endsWith(".mjs"))) continue;

      const source = data.toString("utf8");
      for (const rule of DEFAULT_JS_HEURISTIC_RULES) {
        if (allowlist.has(rule.id)) continue;
        if (rule.test(source)) {
          findings.push({ id: rule.id, message: rule.message, file: relPath });
        }
      }
    }

    return {
      formatVersion,
      fileRecords,
      fileCount: fileRecords.length,
      unpackedSize,
      findings,
    };
  }

  async search({
    q,
    category,
    tag,
    verified,
    featured,
    sort = "relevance",
    limit = 20,
    offset = 0,
    cursor = null,
  }) {
    const tokens = tokenize(q);
    const normalizedCategory = category ? String(category).toLowerCase() : null;
    const normalizedTag = tag ? String(tag).toLowerCase() : null;

    const { extRows, versionMap } = await this.db.withRead((db) => {
      // Pull extensions + versions into memory for flexible semver + cursor pagination.
      const rows =
        db.exec(
          `SELECT id, name, display_name, publisher, description, categories_json, tags_json, screenshots_json,
                  verified, featured, deprecated, blocked, malicious, download_count, updated_at
           FROM extensions`
        )[0]?.values || [];

      /** @type {Record<string, string[]>} */
      const versions = {};
      const versionRows = db.exec(`SELECT extension_id, version, yanked FROM extension_versions`)[0]?.values;
      if (versionRows) {
        for (const [extensionId, version, yanked] of versionRows) {
          if (yanked) continue;
          const key = String(extensionId);
          versions[key] = versions[key] || [];
          versions[key].push(String(version));
        }
      }

      return { extRows: rows, versionMap: versions };
    });

    const entries = extRows
      .map((row) => {
        const [
          id,
          name,
          displayName,
          publisher,
          description,
          categoriesJson,
          tagsJson,
          screenshotsJson,
          vVerified,
          vFeatured,
          vDeprecated,
          vBlocked,
          vMalicious,
          downloadCount,
          updatedAt,
        ] = row;

        const categories = normalizeStringArray(parseJsonArray(categoriesJson));
        const tags = normalizeStringArray(parseJsonArray(tagsJson));
        const screenshots = normalizeStringArray(parseJsonArray(screenshotsJson));

        const latestVersion = maxSemver(versionMap[String(id)] || []) || null;

        return {
          id: String(id),
          name: String(name),
          displayName: String(displayName),
          publisher: String(publisher),
          description: String(description || ""),
          categories,
          tags,
          screenshots,
          verified: Boolean(vVerified),
          featured: Boolean(vFeatured),
          deprecated: Boolean(vDeprecated),
          blocked: Boolean(vBlocked),
          malicious: Boolean(vMalicious),
          downloadCount: Number(downloadCount || 0),
          updatedAt: String(updatedAt || ""),
          latestVersion,
        };
      })
      .filter((ext) => {
        if (ext.blocked || ext.malicious || ext.deprecated) return false;
        if (!ext.latestVersion) return false;
        if (verified !== undefined && Boolean(verified) !== ext.verified) return false;
        if (featured !== undefined && Boolean(featured) !== ext.featured) return false;

        if (normalizedCategory && !ext.categories.some((c) => String(c).toLowerCase() === normalizedCategory)) {
          return false;
        }
        if (normalizedTag && !ext.tags.some((t) => String(t).toLowerCase() === normalizedTag)) {
          return false;
        }
        if (tokens.length === 0) return true;
        return calculateSearchScore(ext, tokens) > 0;
      })
      .map((ext) => ({
        ...ext,
        score: calculateSearchScore(ext, tokens),
      }));

    const mode = sort || (tokens.length > 0 ? "relevance" : "updated");
    entries.sort((a, b) => {
      if (mode === "downloads") {
        if ((a.downloadCount || 0) !== (b.downloadCount || 0)) return (b.downloadCount || 0) - (a.downloadCount || 0);
        return a.id < b.id ? -1 : 1;
      }
      if (mode === "updated") {
        if (a.updatedAt !== b.updatedAt) return a.updatedAt < b.updatedAt ? 1 : -1;
        return a.id < b.id ? -1 : 1;
      }

      // relevance (default): preserve old sorting behavior.
      if (a.score !== b.score) return b.score - a.score;
      if (a.featured !== b.featured) return a.featured ? -1 : 1;
      if (a.verified !== b.verified) return a.verified ? -1 : 1;
      if ((a.downloadCount || 0) !== (b.downloadCount || 0)) return (b.downloadCount || 0) - (a.downloadCount || 0);
      return a.id < b.id ? -1 : 1;
    });

    function decodeCursor(raw) {
      try {
        return JSON.parse(Buffer.from(String(raw), "base64url").toString("utf8"));
      } catch {
        return null;
      }
    }

    function encodeCursorPayload(payload) {
      return Buffer.from(JSON.stringify(payload)).toString("base64url");
    }

    function cursorPayloadForItem(item) {
      if (mode === "downloads") {
        return { sort: mode, downloadCount: item.downloadCount, id: item.id };
      }
      if (mode === "updated") {
        return { sort: mode, updatedAt: item.updatedAt, id: item.id };
      }
      return {
        sort: "relevance",
        score: item.score,
        featured: item.featured ? 1 : 0,
        verified: item.verified ? 1 : 0,
        downloadCount: item.downloadCount,
        id: item.id,
      };
    }

    function isAfterCursor(item, cur) {
      if (!cur || typeof cur !== "object") return true;
      if (cur.sort !== mode && !(mode === "relevance" && cur.sort === "relevance")) return true;

      const curId = String(cur.id || "");
      if (!curId) return true;

      if (mode === "downloads") {
        const curDl = Number(cur.downloadCount || 0);
        if (item.downloadCount !== curDl) return item.downloadCount < curDl;
        return item.id > curId;
      }

      if (mode === "updated") {
        const curUpdatedAt = String(cur.updatedAt || "");
        if (item.updatedAt !== curUpdatedAt) return item.updatedAt < curUpdatedAt;
        return item.id > curId;
      }

      const curScore = Number(cur.score || 0);
      const curFeatured = Number(cur.featured || 0);
      const curVerified = Number(cur.verified || 0);
      const curDl = Number(cur.downloadCount || 0);

      if (item.score !== curScore) return item.score < curScore;
      if ((item.featured ? 1 : 0) !== curFeatured) return (item.featured ? 1 : 0) < curFeatured;
      if ((item.verified ? 1 : 0) !== curVerified) return (item.verified ? 1 : 0) < curVerified;
      if (item.downloadCount !== curDl) return item.downloadCount < curDl;
      return item.id > curId;
    }

    const boundedLimit = Math.max(1, Math.min(100, Number(limit) || 20));

    let filtered = entries;
    if (cursor) {
      const decoded = decodeCursor(cursor);
      filtered = entries.filter((item) => isAfterCursor(item, decoded));
    } else {
      const start = Math.max(0, Number(offset) || 0);
      filtered = entries.slice(start);
    }

    const page = filtered.slice(0, boundedLimit);
    const next = filtered.length > boundedLimit ? page[page.length - 1] : null;

    const results = page.map((ext) => ({
      id: ext.id,
      name: ext.name,
      displayName: ext.displayName,
      publisher: ext.publisher,
      description: ext.description,
      latestVersion: ext.latestVersion,
      verified: ext.verified,
      featured: ext.featured,
      categories: ext.categories,
      tags: ext.tags,
      screenshots: ext.screenshots,
      downloadCount: ext.downloadCount,
      updatedAt: ext.updatedAt,
    }));

    return {
      total: entries.length,
      results,
      nextCursor: next ? encodeCursorPayload(cursorPayloadForItem(next)) : null,
    };
  }

  async getExtension(id, { includeHidden = false } = {}) {
    return this.db.withRead((db) => {
      const stmt = db.prepare(
        `SELECT id, name, display_name, publisher, description, categories_json, tags_json, screenshots_json,
                verified, featured, deprecated, blocked, malicious, download_count, created_at, updated_at
         FROM extensions WHERE id = ? LIMIT 1`
      );
      stmt.bind([id]);
      if (!stmt.step()) {
        stmt.free();
        return null;
      }
      const row = stmt.getAsObject();
      stmt.free();

      const hidden = Boolean(row.blocked) || Boolean(row.malicious);
      if (hidden && !includeHidden) return null;

      const versionsStmt = db.prepare(
        `SELECT version, sha256, uploaded_at, yanked, readme
         FROM extension_versions WHERE extension_id = ?`
      );
      versionsStmt.bind([id]);

      const versions = [];
      const unyanked = [];
      /** @type {Record<string, { readme: string }>} */
      const versionMeta = {};
      while (versionsStmt.step()) {
        const v = versionsStmt.getAsObject();
        const ver = String(v.version);
        const yanked = Boolean(v.yanked);
        versions.push({
          version: ver,
          sha256: String(v.sha256),
          uploadedAt: String(v.uploaded_at),
          yanked,
        });
        versionMeta[ver] = { readme: String(v.readme || "") };
        if (!yanked) unyanked.push(ver);
      }
      versionsStmt.free();

      versions.sort((a, b) => compareSemver(b.version, a.version));
      const latestVersion = maxSemver(unyanked) || null;
      const readme = latestVersion ? versionMeta[latestVersion]?.readme || "" : "";

      const publisherStmt = db.prepare(`SELECT public_key_pem FROM publishers WHERE publisher = ? LIMIT 1`);
      publisherStmt.bind([row.publisher]);
      const hasPublisher = publisherStmt.step();
      const publisherRow = hasPublisher ? publisherStmt.getAsObject() : null;
      publisherStmt.free();

      const publisherKeysStmt = db.prepare(
        `SELECT id, public_key_pem, revoked, is_primary
         FROM publisher_keys
         WHERE publisher = ?
         ORDER BY is_primary DESC, created_at ASC`
      );
      publisherKeysStmt.bind([row.publisher]);
      /** @type {{ id: string, publicKeyPem: string, revoked: boolean }[]} */
      const publisherKeys = [];
      let primaryKeyPem = null;
      let firstActiveKeyPem = null;
      while (publisherKeysStmt.step()) {
        const keyRow = publisherKeysStmt.getAsObject();
        const revoked = Boolean(keyRow.revoked);
        const isPrimary = Boolean(keyRow.is_primary);
        const pem = String(keyRow.public_key_pem || "");
        publisherKeys.push({
          id: String(keyRow.id),
          publicKeyPem: pem,
          revoked,
        });
        if (!revoked && !firstActiveKeyPem) {
          firstActiveKeyPem = pem;
        }
        if (!primaryKeyPem && isPrimary && !revoked) {
          primaryKeyPem = pem;
        }
      }
      publisherKeysStmt.free();

      const legacyPublisherKeyPem = publisherRow ? String(publisherRow.public_key_pem) : null;
      const publisherPublicKeyPem =
        primaryKeyPem || firstActiveKeyPem || (publisherKeys.length === 0 && legacyPublisherKeyPem ? legacyPublisherKeyPem : null);
      if (publisherKeys.length === 0 && publisherPublicKeyPem) {
        try {
          const normalized = normalizePublicKeyPem(publisherPublicKeyPem);
          publisherKeys.push({
            id: publisherKeyIdFromPublicKeyPem(normalized),
            publicKeyPem: normalized,
            revoked: false,
          });
        } catch {
          // Ignore; publisherPublicKeyPem is still returned for backward compatibility.
        }
      }

      return {
        id: String(row.id),
        name: String(row.name),
        displayName: String(row.display_name),
        publisher: String(row.publisher),
        description: String(row.description || ""),
        categories: normalizeStringArray(parseJsonArray(row.categories_json)),
        tags: normalizeStringArray(parseJsonArray(row.tags_json)),
        screenshots: normalizeStringArray(parseJsonArray(row.screenshots_json)),
        verified: Boolean(row.verified),
        featured: Boolean(row.featured),
        deprecated: Boolean(row.deprecated),
        blocked: Boolean(row.blocked),
        malicious: Boolean(row.malicious),
        downloadCount: Number(row.download_count || 0),
        latestVersion,
        versions,
        readme,
        publisherPublicKeyPem,
        publisherKeys,
        updatedAt: String(row.updated_at || ""),
        createdAt: String(row.created_at || ""),
      };
    });
  }

  async getPackage(id, version, { includeBytes = true, incrementDownloadCount = includeBytes, includePath = false } = {}) {
    const meta = await this.db.withRead((db) => {
      const stmt = db.prepare(
        `SELECT e.publisher, e.blocked, e.malicious,
                v.signature_base64, v.sha256, v.package_path, v.yanked, v.format_version,
                v.signing_key_id,
                v.files_json,
                s.status AS scan_status
         FROM extensions e
         JOIN extension_versions v ON v.extension_id = e.id
         LEFT JOIN package_scans s ON s.extension_id = v.extension_id AND s.version = v.version
         WHERE e.id = ? AND v.version = ? LIMIT 1`
      );
      stmt.bind([id, version]);
      if (!stmt.step()) {
        stmt.free();
        return null;
      }
      const row = stmt.getAsObject();
      stmt.free();

      if (row.blocked || row.malicious || row.yanked) {
        return null;
      }

      const scanStatusRaw = row.scan_status ? String(row.scan_status) : null;
      if (scanStatusRaw === PACKAGE_SCAN_STATUS.FAILED) {
        return null;
      }
      if (this.requireScanPassedForDownload && scanStatusRaw !== PACKAGE_SCAN_STATUS.PASSED) {
        return null;
      }

      const filesJson = row.files_json ? String(row.files_json) : "[]";
      const filesSha256 = sha256Utf8(filesJson);

      const packagePath = row.package_path ? String(row.package_path) : null;
      let packageBytes = null;
      if (includeBytes && !packagePath) {
        const bytesStmt = db.prepare(
          `SELECT package_bytes FROM extension_versions WHERE extension_id = ? AND version = ? LIMIT 1`
        );
        bytesStmt.bind([id, version]);
        if (!bytesStmt.step()) {
          bytesStmt.free();
          return null;
        }
        const bytesRow = bytesStmt.getAsObject();
        bytesStmt.free();
        packageBytes = bytesRow.package_bytes;
      }

      return {
        publisher: String(row.publisher),
        signatureBase64: String(row.signature_base64),
        sha256: String(row.sha256),
        formatVersion: Number(row.format_version || 1),
        signingKeyId: row.signing_key_id ? String(row.signing_key_id) : null,
        scanStatus: scanStatusRaw || "unknown",
        filesSha256,
        packageBytes,
        packagePath,
      };
    });

    if (!meta) return null;

    const fullPath =
      meta.packagePath && (includeBytes || includePath) ? resolvePackagePath(this.dataDir, meta.packagePath) : null;

    /** @type {Buffer | null} */
    let bytes = null;
    if (includeBytes) {
      if (fullPath) {
        try {
          bytes = await fs.readFile(fullPath);
        } catch {
          bytes = null;
        }
      } else {
        bytes = Buffer.from(meta.packageBytes);
      }

      if (!bytes || bytes.length === 0) return null;
    }

    if (incrementDownloadCount) {
      await this.db.withTransaction((db) => {
        db.run(`UPDATE extensions SET download_count = download_count + 1 WHERE id = ?`, [id]);
      });
    }

    return {
      bytes,
      signatureBase64: meta.signatureBase64,
      sha256: meta.sha256,
      formatVersion: meta.formatVersion,
      publisher: meta.publisher,
      signingKeyId: meta.signingKeyId,
      scanStatus: meta.scanStatus,
      filesSha256: meta.filesSha256,
      packagePath: includePath ? fullPath : null,
    };
  }

  async incrementDownloadCount(id) {
    await this.db.withTransaction((db) => {
      db.run(`UPDATE extensions SET download_count = download_count + 1 WHERE id = ?`, [id]);
    });
  }

  async setVersionFlags(
    id,
    version,
    { yanked },
    { actor = "admin", ip = null } = {}
  ) {
    const now = new Date().toISOString();
    await this.db.withTransaction((db) => {
      const stmt = db.prepare(
        `SELECT extension_id, version FROM extension_versions WHERE extension_id = ? AND version = ? LIMIT 1`
      );
      stmt.bind([id, version]);
      const exists = stmt.step();
      stmt.free();
      if (!exists) throw new Error("Extension version not found");

      if (yanked !== undefined) {
        db.run(
          `UPDATE extension_versions SET yanked = ?, yanked_at = ? WHERE extension_id = ? AND version = ?`,
          [yanked ? 1 : 0, yanked ? now : null, id, version]
        );
        db.run(`UPDATE extensions SET updated_at = ? WHERE id = ?`, [now, id]);
      }

      const audit = db.prepare(
        `INSERT INTO audit_log (id, actor, action, extension_id, version, ip, details_json, created_at)
         VALUES (?, ?, ?, ?, ?, ?, ?, ?)`
      );
      audit.run([
        randomId(),
        actor,
        "extension.version.flags",
        id,
        version,
        ip,
        safeJsonStringify({ yanked }),
        now,
      ]);
      audit.free();
    });

    return { id, version, yanked: Boolean(yanked) };
  }

  async listAuditLog({ limit = 50, offset = 0 } = {}) {
    const boundedLimit = Math.max(1, Math.min(200, Number(limit) || 50));
    const boundedOffset = Math.max(0, Number(offset) || 0);

    return this.db.withRead((db) => {
      const stmt = db.prepare(
        `SELECT id, actor, action, extension_id, version, ip, details_json, created_at
         FROM audit_log
         ORDER BY created_at DESC
         LIMIT ? OFFSET ?`
      );
      stmt.bind([boundedLimit, boundedOffset]);
      const out = [];
      while (stmt.step()) {
        const row = stmt.getAsObject();
        out.push({
          id: String(row.id),
          actor: String(row.actor),
          action: String(row.action),
          extensionId: row.extension_id ? String(row.extension_id) : null,
          version: row.version ? String(row.version) : null,
          ip: row.ip ? String(row.ip) : null,
          details: (() => {
            try {
              return JSON.parse(String(row.details_json || "{}"));
            } catch {
              return {};
            }
          })(),
          createdAt: String(row.created_at),
        });
      }
      stmt.free();
      return out;
    });
  }

  close() {
    this.db.close();
  }
}

module.exports = {
  MarketplaceStore,
  extensionIdFromManifest,
};
