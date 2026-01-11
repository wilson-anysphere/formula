const fs = require("node:fs/promises");
const path = require("node:path");
const zlib = require("node:zlib");

const PACKAGE_FORMAT = "formula-extension-package";
const PACKAGE_FORMAT_VERSION = 1;

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
  const jsonBytes = zlib.gunzipSync(packageBytes);
  const parsed = JSON.parse(jsonBytes.toString("utf8"));

  if (parsed?.format !== PACKAGE_FORMAT || parsed?.formatVersion !== PACKAGE_FORMAT_VERSION) {
    throw new Error("Unsupported extension package format");
  }

  if (!Array.isArray(parsed.files) || typeof parsed.manifest !== "object" || !parsed.manifest) {
    throw new Error("Invalid extension package contents");
  }

  return parsed;
}

function safeJoin(baseDir, relPath) {
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
    if (part.endsWith(" ") || part.endsWith(".")) {
      throw new Error(`Invalid path in extension package: ${relPath}`);
    }
    if (windowsReservedRe.test(part)) {
      throw new Error(`Invalid path in extension package: ${relPath}`);
    }
  }
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
