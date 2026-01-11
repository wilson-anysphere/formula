const fs = require("node:fs/promises");
const path = require("node:path");
const zlib = require("node:zlib");

const PACKAGE_FORMAT = "formula-extension-package";
const PACKAGE_FORMAT_VERSION = 1;
const MAX_UNCOMPRESSED_BYTES = 50 * 1024 * 1024;

function isPlainObject(value) {
  return Boolean(value) && typeof value === "object" && !Array.isArray(value);
}

function canonicalJsonString(value) {
  if (value === null) return "null";
  if (typeof value === "string") return JSON.stringify(value);
  if (typeof value === "number") return JSON.stringify(value);
  if (typeof value === "boolean") return value ? "true" : "false";
  if (Array.isArray(value)) {
    return `[${value.map((v) => canonicalJsonString(v)).join(",")}]`;
  }
  if (isPlainObject(value)) {
    const keys = Object.keys(value).sort();
    const items = keys
      .filter((k) => value[k] !== undefined)
      .map((k) => `${JSON.stringify(k)}:${canonicalJsonString(value[k])}`);
    return `{${items.join(",")}}`;
  }
  return JSON.stringify(value);
}

function manifestsMatch(a, b) {
  return canonicalJsonString(a) === canonicalJsonString(b);
}

function normalizePath(relPath) {
  const normalized = String(relPath).replace(/\\/g, "/");
  if (normalized.startsWith("/") || normalized.includes("\0")) {
    throw new Error(`Invalid path in extension package: ${relPath}`);
  }
  const parts = normalized.split("/");
  if (parts.some((p) => p === "" || p === "." || p === "..")) {
    throw new Error(`Invalid path in extension package: ${relPath}`);
  }
  if (parts.some((p) => p.includes(":"))) {
    throw new Error(`Invalid path in extension package: ${relPath}`);
  }
  const windowsReservedRe = /^(con|prn|aux|nul|com[1-9]|lpt[1-9])(\..*)?$/i;
  for (const part of parts) {
    if (/[<>:"|?*]/.test(part)) {
      throw new Error(`Invalid path in extension package: ${relPath}`);
    }
    if (part.endsWith(" ") || part.endsWith(".")) {
      throw new Error(`Invalid path in extension package: ${relPath}`);
    }
    if (windowsReservedRe.test(part)) {
      throw new Error(`Invalid path in extension package: ${relPath}`);
    }
  }
  return parts.join("/");
}

async function walkFiles(rootDir) {
  const results = [];

  async function visit(dir) {
    const entries = await fs.readdir(dir, { withFileTypes: true });
    for (const entry of entries) {
      const abs = path.join(dir, entry.name);
      const rel = path.relative(rootDir, abs).replace(/\\/g, "/");

      if (rel === "" || rel.startsWith("..") || path.isAbsolute(rel)) continue;

      if (entry.isDirectory()) {
        if (entry.name === "node_modules" || entry.name === ".git") continue;
        await visit(abs);
        continue;
      }

      if (!entry.isFile()) continue;
      results.push({ abs, rel });
    }
  }

  await visit(rootDir);
  results.sort((a, b) => (a.rel < b.rel ? -1 : a.rel > b.rel ? 1 : 0));
  return results;
}

async function loadExtensionManifest(extensionDir) {
  const manifestPath = path.join(extensionDir, "package.json");
  const raw = await fs.readFile(manifestPath, "utf8");
  return JSON.parse(raw);
}

async function createExtensionPackageV1(extensionDir) {
  const manifest = await loadExtensionManifest(extensionDir);

  const files = [];
  const entries = await walkFiles(extensionDir);
  for (const entry of entries) {
    const data = await fs.readFile(entry.abs);
    files.push({
      path: entry.rel,
      dataBase64: data.toString("base64"),
    });
  }

  const bundle = {
    format: PACKAGE_FORMAT,
    formatVersion: PACKAGE_FORMAT_VERSION,
    createdAt: new Date().toISOString(),
    manifest,
    files,
  };

  const jsonBytes = Buffer.from(JSON.stringify(bundle), "utf8");
  return zlib.gzipSync(jsonBytes, { level: 9 });
}

function readExtensionPackageV1(packageBytes) {
  let jsonBytes;
  try {
    jsonBytes = zlib.gunzipSync(packageBytes, { maxOutputLength: MAX_UNCOMPRESSED_BYTES });
  } catch (error) {
    if (error && error.code === "ERR_BUFFER_TOO_LARGE") {
      throw new Error("Extension package exceeds maximum uncompressed size");
    }
    throw error;
  }
  const parsed = JSON.parse(jsonBytes.toString("utf8"));

  if (parsed?.format !== PACKAGE_FORMAT || parsed?.formatVersion !== PACKAGE_FORMAT_VERSION) {
    throw new Error("Unsupported extension package format");
  }

  if (!Array.isArray(parsed.files) || typeof parsed.manifest !== "object" || !parsed.manifest) {
    throw new Error("Invalid extension package contents");
  }

  const seenFoldedPaths = new Set();
  const packageJsonEntry =
    parsed.files.find((f) => typeof f?.path === "string" && normalizePath(f.path) === "package.json") || null;
  if (!packageJsonEntry?.dataBase64 || typeof packageJsonEntry.dataBase64 !== "string") {
    throw new Error("Invalid extension package: missing package.json file");
  }

  for (const file of parsed.files) {
    if (!file?.path || typeof file.path !== "string") {
      throw new Error("Invalid file entry in extension package");
    }
    const normalizedPath = normalizePath(file.path);
    const folded = normalizedPath.toLowerCase();
    if (seenFoldedPaths.has(folded)) {
      throw new Error(`Invalid extension package: duplicate file path (case-insensitive): ${normalizedPath}`);
    }
    seenFoldedPaths.add(folded);
  }

  let packageJson = null;
  try {
    packageJson = JSON.parse(Buffer.from(packageJsonEntry.dataBase64, "base64").toString("utf8"));
  } catch {
    throw new Error("Invalid extension package: package.json is not valid JSON");
  }
  if (!isPlainObject(packageJson)) {
    throw new Error("Invalid extension package: package.json must be a JSON object");
  }
  if (!manifestsMatch(packageJson, parsed.manifest)) {
    throw new Error("Invalid extension package: package.json does not match embedded manifest");
  }

  return parsed;
}

function safeJoin(baseDir, relPath) {
  const normalized = normalizePath(relPath);
  const full = path.join(baseDir, normalized);
  const relative = path.relative(baseDir, full);
  if (relative.startsWith("..") || path.isAbsolute(relative)) {
    throw new Error(`Path traversal in extension package: ${relPath}`);
  }
  return full;
}

async function extractExtensionPackageV1(packageBytes, destDir) {
  const bundle = readExtensionPackageV1(packageBytes);

  await fs.mkdir(destDir, { recursive: true });

  for (const file of bundle.files) {
    if (!file?.path || typeof file.path !== "string" || typeof file.dataBase64 !== "string") {
      throw new Error("Invalid file entry in extension package");
    }
    const outPath = safeJoin(destDir, file.path);
    await fs.mkdir(path.dirname(outPath), { recursive: true });
    await fs.writeFile(outPath, Buffer.from(file.dataBase64, "base64"));
  }

  return bundle.manifest;
}

module.exports = {
  PACKAGE_FORMAT,
  PACKAGE_FORMAT_VERSION,

  createExtensionPackageV1,
  extractExtensionPackageV1,
  loadExtensionManifest,
  readExtensionPackageV1,
};
