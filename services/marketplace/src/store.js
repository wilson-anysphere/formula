const fs = require("node:fs/promises");
const path = require("node:path");
const crypto = require("node:crypto");
const zlib = require("node:zlib");

const { verifyBytesSignature, sha256 } = require("../../../shared/crypto/signing");
const {
  canonicalJsonBytes,
  detectExtensionPackageFormatVersion,
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

  const validated = validateExtensionManifest(manifest, { enforceEngine: false });

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
  constructor({ dataDir }) {
    this.dataDir = dataDir;
    this.legacyStorePath = path.join(this.dataDir, "store.json");
    this.legacyPackageDir = path.join(this.dataDir, "packages");
    this.dbPath = path.join(this.dataDir, "marketplace.sqlite");
    this.db = new SqliteFileDb({
      filePath: this.dbPath,
      migrationsDir: path.join(__dirname, "db", "migrations"),
    });
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

    await this.db.withTransaction(async (tx) => {
      for (const publisherRecord of Object.values(legacy.publishers || {})) {
        if (!publisherRecord?.publisher) continue;
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
            publisherRecord.publicKeyPem,
            publisherRecord.verified ? 1 : 0,
            publisherRecord.createdAt || now,
          ]
        );
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

          tx.run(
            `INSERT INTO extension_versions
              (extension_id, version, sha256, signature_base64, manifest_json, readme, package_bytes, package_path, uploaded_at, yanked, yanked_at)
             VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, 0, NULL)
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
    await this.db.withTransaction((db) => {
      db.run(
        `
          INSERT INTO publishers (publisher, token_sha256, public_key_pem, verified, created_at)
          VALUES (?, ?, ?, ?, ?)
          ON CONFLICT(publisher) DO UPDATE SET
            token_sha256 = excluded.token_sha256,
            public_key_pem = excluded.public_key_pem,
            verified = excluded.verified
        `,
        [publisher, tokenSha256, publicKeyPem, verified ? 1 : 0, now]
      );

      const audit = db.prepare(
        `INSERT INTO audit_log (id, actor, action, extension_id, version, ip, details_json, created_at)
         VALUES (?, ?, ?, ?, ?, ?, ?, ?)`
      );
      audit.run([randomId(), "admin", "publisher.register", null, null, null, safeJsonStringify({ publisher }), now]);
      audit.free();
    });
  }

  async getPublisherByTokenSha256(tokenSha256) {
    return this.db.withRead((db) => {
      const stmt = db.prepare(
        `SELECT publisher, token_sha256, public_key_pem, verified, created_at
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
        createdAt: row.created_at,
      };
    });
  }

  async getPublisher(publisher) {
    return this.db.withRead((db) => {
      const stmt = db.prepare(
        `SELECT publisher, token_sha256, public_key_pem, verified, created_at
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
        createdAt: row.created_at,
      };
    });
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

    if (formatVersion === 1) {
      if (!signatureBase64) throw new Error("signatureBase64 is required for v1 extension packages");
      if (packageBytes.length > MAX_COMPRESSED_BYTES) {
        throw new Error("Package exceeds maximum compressed size");
      }

      const signatureOk = verifyBytesSignature(packageBytes, signatureBase64, publisherRecord.publicKeyPem);
      if (!signatureOk) throw new Error("Package signature verification failed");
      storedSignatureBase64 = String(signatureBase64);

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

      fileCount = fileRecords.length;
      if (unpackedSize > MAX_UNPACKED_BYTES) {
        throw new Error("Package exceeds maximum uncompressed payload size");
      }
    } else if (formatVersion === 2) {
      const verified = verifyExtensionPackageV2(packageBytes, publisherRecord.publicKeyPem);
      manifest = verified.manifest;
      storedSignatureBase64 = verified.signatureBase64;
      fileRecords = verified.files;
      fileCount = verified.fileCount;
      unpackedSize = verified.unpackedSize;
      readme = verified.readme || "";

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

    const now = new Date().toISOString();
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
           uploaded_at, yanked, yanked_at, format_version, file_count, unpacked_size, files_json)
         VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, 0, NULL, ?, ?, ?, ?)`,
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
        ]
      );
    }).catch((err) => {
      if (String(err?.message || "").includes("UNIQUE constraint failed: extension_versions.extension_id, extension_versions.version")) {
        throw new Error("That extension version is already published");
      }
      throw err;
    });

    return { id, version };
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
        publisherPublicKeyPem: publisherRow ? String(publisherRow.public_key_pem) : null,
        updatedAt: String(row.updated_at || ""),
        createdAt: String(row.created_at || ""),
      };
    });
  }

  async getPackage(id, version) {
    const meta = await this.db.withRead((db) => {
      const stmt = db.prepare(
        `SELECT e.publisher, e.blocked, e.malicious,
                v.signature_base64, v.sha256, v.package_bytes, v.package_path, v.yanked, v.format_version
         FROM extensions e
         JOIN extension_versions v ON v.extension_id = e.id
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

      return {
        publisher: String(row.publisher),
        signatureBase64: String(row.signature_base64),
        sha256: String(row.sha256),
        formatVersion: Number(row.format_version || 1),
        packageBytes: row.package_bytes,
        packagePath: row.package_path ? String(row.package_path) : null,
      };
    });

    if (!meta) return null;

    /** @type {Buffer | null} */
    let bytes = null;
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

    if (!bytes || bytes.length === 0) return null;

    await this.db.withTransaction((db) => {
      db.run(`UPDATE extensions SET download_count = download_count + 1 WHERE id = ?`, [id]);
    });

    return {
      bytes,
      signatureBase64: meta.signatureBase64,
      sha256: meta.sha256,
      formatVersion: meta.formatVersion,
      publisher: meta.publisher,
    };
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
