const fs = require("node:fs/promises");
const path = require("node:path");

const { verifyBytesSignature, sha256 } = require("../../../shared/crypto/signing");
const { readExtensionPackage } = require("../../../shared/extension-package");
const { compareSemver, isValidSemver, maxSemver } = require("../../../shared/semver");

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

async function atomicWriteJson(filePath, data) {
  const dir = path.dirname(filePath);
  await fs.mkdir(dir, { recursive: true });

  const tmpPath = `${filePath}.tmp`;
  await fs.writeFile(tmpPath, JSON.stringify(data, null, 2));
  await fs.rename(tmpPath, filePath);
}

async function readJsonIfExists(filePath, fallback) {
  try {
    const raw = await fs.readFile(filePath, "utf8");
    return JSON.parse(raw);
  } catch (error) {
    if (error && (error.code === "ENOENT" || error.code === "ENOTDIR")) return fallback;
    throw error;
  }
}

function tokenize(query) {
  return String(query || "")
    .toLowerCase()
    .split(/\s+/)
    .map((t) => t.trim())
    .filter(Boolean);
}

function textIncludesAllTokens(text, tokens) {
  const haystack = String(text || "").toLowerCase();
  return tokens.every((t) => haystack.includes(t));
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

  for (const field of fields) {
    if (textIncludesAllTokens(field.text, tokens)) score += field.weight;
  }

  if (ext.featured) score += 8;
  if (ext.verified) score += 3;
  score += Math.min(5, Math.log10((ext.downloadCount || 0) + 1));

  return score;
}

class MarketplaceStore {
  constructor({ dataDir }) {
    this.dataDir = dataDir;
    this.storePath = path.join(this.dataDir, "store.json");
    this.packageDir = path.join(this.dataDir, "packages");
    this.state = { publishers: {}, extensions: {} };
  }

  async init() {
    await fs.mkdir(this.dataDir, { recursive: true });
    await fs.mkdir(this.packageDir, { recursive: true });
    this.state = await readJsonIfExists(this.storePath, { publishers: {}, extensions: {} });
  }

  async persist() {
    await atomicWriteJson(this.storePath, this.state);
  }

  async registerPublisher({ publisher, tokenSha256, publicKeyPem, verified = false }) {
    if (!publisher || typeof publisher !== "string") throw new Error("publisher is required");
    if (!tokenSha256 || typeof tokenSha256 !== "string") throw new Error("tokenSha256 is required");
    if (!publicKeyPem || typeof publicKeyPem !== "string") throw new Error("publicKeyPem is required");

    this.state.publishers[publisher] = {
      publisher,
      tokenSha256,
      publicKeyPem,
      verified: Boolean(verified),
      createdAt: new Date().toISOString(),
    };
    await this.persist();
  }

  getPublisherByTokenSha256(tokenSha256) {
    return Object.values(this.state.publishers).find((p) => p.tokenSha256 === tokenSha256) || null;
  }

  getPublisher(publisher) {
    return this.state.publishers[publisher] || null;
  }

  async setExtensionFlags(id, { verified, featured }) {
    const ext = this.state.extensions[id];
    if (!ext) throw new Error("Extension not found");
    if (verified !== undefined) ext.verified = Boolean(verified);
    if (featured !== undefined) ext.featured = Boolean(featured);
    await this.persist();
    return ext;
  }

  async publishExtension({ publisher, packageBytes, signatureBase64 }) {
    const publisherRecord = this.getPublisher(publisher);
    if (!publisherRecord) throw new Error(`Unknown publisher: ${publisher}`);

    const signatureOk = verifyBytesSignature(packageBytes, signatureBase64, publisherRecord.publicKeyPem);
    if (!signatureOk) throw new Error("Package signature verification failed");

    const bundle = readExtensionPackage(packageBytes);
    const manifest = bundle.manifest;

    if (manifest.publisher !== publisher) {
      throw new Error("Manifest publisher does not match authenticated publisher");
    }

    if (!isValidSemver(manifest.version)) throw new Error("Manifest version is not valid semver");

    const id = extensionIdFromManifest(manifest);
    const version = manifest.version;

    const existing = this.state.extensions[id];
    if (existing?.versions?.[version]) {
      throw new Error("That extension version is already published");
    }

    const pkgSha = sha256(packageBytes);
    const pkgPath = path.join(this.packageDir, id, `${version}.fextpkg`);
    await fs.mkdir(path.dirname(pkgPath), { recursive: true });
    await fs.writeFile(pkgPath, packageBytes);

    const readmeEntry =
      bundle.files.find((f) => typeof f?.path === "string" && f.path.toLowerCase() === "readme.md") || null;
    const readme = readmeEntry ? Buffer.from(readmeEntry.dataBase64, "base64").toString("utf8") : "";

    const categories = normalizeStringArray(manifest.categories);
    const tags = normalizeStringArray(manifest.tags);
    const screenshots = normalizeStringArray(manifest.screenshots);

    const extensionRecord = existing || {
      id,
      name: manifest.name,
      displayName: manifest.displayName || manifest.name,
      publisher: manifest.publisher,
      description: manifest.description || "",
      categories,
      tags,
      screenshots,
      verified: Boolean(publisherRecord.verified),
      featured: false,
      downloadCount: 0,
      createdAt: new Date().toISOString(),
      latestVersion: version,
      versions: {},
    };

    extensionRecord.name = manifest.name;
    extensionRecord.displayName = manifest.displayName || manifest.name;
    extensionRecord.description = manifest.description || "";
    extensionRecord.categories = categories;
    extensionRecord.tags = tags;
    extensionRecord.screenshots = screenshots;

    extensionRecord.versions[version] = {
      version,
      sha256: pkgSha,
      signatureBase64,
      packagePath: pkgPath,
      readme,
      uploadedAt: new Date().toISOString(),
    };

    const versions = Object.keys(extensionRecord.versions);
    extensionRecord.latestVersion = maxSemver(versions) || version;
    extensionRecord.updatedAt = new Date().toISOString();

    this.state.extensions[id] = extensionRecord;
    await this.persist();

    return { id, version };
  }

  search({ q, category, limit = 20, offset = 0 }) {
    const tokens = tokenize(q);
    const normalizedCategory = category ? String(category).toLowerCase() : null;

    const entries = Object.values(this.state.extensions)
      .filter((ext) => {
        if (normalizedCategory && !ext.categories.some((c) => String(c).toLowerCase() === normalizedCategory)) {
          return false;
        }
        if (tokens.length === 0) return true;
        return calculateSearchScore(ext, tokens) > 0;
      })
      .map((ext) => ({
        ...ext,
        score: calculateSearchScore(ext, tokens),
      }))
      .sort((a, b) => {
        if (a.score !== b.score) return b.score - a.score;
        if (a.featured !== b.featured) return a.featured ? -1 : 1;
        if (a.verified !== b.verified) return a.verified ? -1 : 1;
        if ((a.downloadCount || 0) !== (b.downloadCount || 0)) return (b.downloadCount || 0) - (a.downloadCount || 0);
        return a.id < b.id ? -1 : 1;
      });

    const slice = entries.slice(offset, offset + limit).map((ext) => ({
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
      results: slice,
    };
  }

  getExtension(id) {
    const ext = this.state.extensions[id];
    if (!ext) return null;

    const versions = Object.keys(ext.versions).sort((a, b) => compareSemver(b, a));

    const latest = ext.versions[ext.latestVersion] || null;
    const publisherRecord = this.getPublisher(ext.publisher);

    return {
      id: ext.id,
      name: ext.name,
      displayName: ext.displayName,
      publisher: ext.publisher,
      description: ext.description,
      categories: ext.categories,
      tags: ext.tags,
      screenshots: ext.screenshots,
      verified: ext.verified,
      featured: ext.featured,
      downloadCount: ext.downloadCount,
      latestVersion: ext.latestVersion,
      versions: versions.map((v) => ({
        version: v,
        sha256: ext.versions[v].sha256,
        uploadedAt: ext.versions[v].uploadedAt,
      })),
      readme: latest?.readme || "",
      publisherPublicKeyPem: publisherRecord?.publicKeyPem || null,
      updatedAt: ext.updatedAt,
      createdAt: ext.createdAt,
    };
  }

  async getPackage(id, version) {
    const ext = this.state.extensions[id];
    if (!ext) return null;
    const entry = ext.versions[version];
    if (!entry) return null;

    const bytes = await fs.readFile(entry.packagePath);

    ext.downloadCount = (ext.downloadCount || 0) + 1;
    await this.persist();

    return {
      bytes,
      signatureBase64: entry.signatureBase64,
      sha256: entry.sha256,
      publisher: ext.publisher,
    };
  }
}

module.exports = {
  MarketplaceStore,
  extensionIdFromManifest,
};

