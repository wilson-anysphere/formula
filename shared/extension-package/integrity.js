const crypto = require("node:crypto");
const fs = require("node:fs/promises");
const path = require("node:path");

function sha256Hex(bytes) {
  return crypto.createHash("sha256").update(bytes).digest("hex");
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
  return parts.join("/");
}

async function walkDirFiles(rootDir) {
  /** @type {{ relPath: string, absPath: string }[]} */
  const files = [];

  async function visit(dir) {
    const entries = await fs.readdir(dir, { withFileTypes: true });
    for (const entry of entries) {
      const absPath = path.join(dir, entry.name);
      const rel = path.relative(rootDir, absPath).replace(/\\/g, "/");
      if (rel === "" || rel.startsWith("..") || path.isAbsolute(rel)) continue;

      if (entry.isSymbolicLink()) {
        throw new Error(`Unexpected symlink in extension directory: ${rel}`);
      }

      if (entry.isDirectory()) {
        await visit(absPath);
        continue;
      }

      if (entry.isFile()) {
        files.push({ relPath: normalizePath(rel), absPath });
        continue;
      }

      throw new Error(`Unsupported filesystem entry in extension directory: ${rel}`);
    }
  }

  await visit(rootDir);
  files.sort((a, b) => (a.relPath < b.relPath ? -1 : a.relPath > b.relPath ? 1 : 0));
  return files;
}

/**
 * Verifies an extracted extension directory matches an expected file manifest.
 *
 * @param {string} extensionDir
 * @param {{ path: string, sha256: string, size: number }[]} expectedFiles
 * @param {{ ignoreExtraPaths?: string[] }} [options]
 * @returns {Promise<{ ok: boolean, reason?: string }>}
 */
async function verifyExtractedExtensionDir(extensionDir, expectedFiles, options = {}) {
  if (!Array.isArray(expectedFiles) || expectedFiles.length === 0) {
    return { ok: false, reason: "Missing integrity metadata (no expected file list)" };
  }

  const ignoreExtraExactPaths = new Set();
  const ignoreExtraBasenames = new Set();
  for (const entry of options.ignoreExtraPaths ?? []) {
    const raw = String(entry);
    if (!raw) continue;
    if (!raw.includes("/")) {
      ignoreExtraBasenames.add(raw);
      continue;
    }
    try {
      ignoreExtraExactPaths.add(normalizePath(raw));
    } catch {
      // ignore invalid ignore entries
    }
  }

  /** @type {Map<string, { sha256: string, size: number }>} */
  const expected = new Map();
  for (const file of expectedFiles) {
    if (!file || typeof file !== "object") {
      return { ok: false, reason: "Invalid integrity metadata (expected file record objects)" };
    }
    if (typeof file.path !== "string" || file.path.trim().length === 0) {
      return { ok: false, reason: "Invalid integrity metadata (file path)" };
    }
    const relPath = normalizePath(file.path);
    const sha = typeof file.sha256 === "string" ? file.sha256.toLowerCase() : "";
    if (!/^[0-9a-f]{64}$/.test(sha)) {
      return { ok: false, reason: `Invalid integrity metadata (sha256 for ${relPath})` };
    }
    const size = file.size;
    if (
      typeof size !== "number" ||
      !Number.isFinite(size) ||
      size < 0 ||
      !Number.isInteger(size) ||
      size > Number.MAX_SAFE_INTEGER
    ) {
      return { ok: false, reason: `Invalid integrity metadata (size for ${relPath})` };
    }
    if (expected.has(relPath)) {
      return { ok: false, reason: `Invalid integrity metadata (duplicate path: ${relPath})` };
    }
    expected.set(relPath, { sha256: sha, size });
  }

  let dirStat;
  try {
    dirStat = await fs.stat(extensionDir);
  } catch (error) {
    if (error && error.code === "ENOENT") {
      return { ok: false, reason: "Extension directory is missing" };
    }
    return { ok: false, reason: `Failed to read extension directory: ${error?.message ?? String(error)}` };
  }

  if (!dirStat.isDirectory()) {
    return { ok: false, reason: "Extension path is not a directory" };
  }

  let actualFiles;
  try {
    actualFiles = await walkDirFiles(extensionDir);
  } catch (error) {
    return { ok: false, reason: error?.message ?? String(error) };
  }

  const actualPaths = new Set(actualFiles.map((f) => f.relPath));

  for (const relPath of expected.keys()) {
    if (!actualPaths.has(relPath)) {
      return { ok: false, reason: `Missing expected file: ${relPath}` };
    }
  }

  for (const relPath of actualPaths) {
    if (ignoreExtraExactPaths.has(relPath)) continue;
    if (ignoreExtraBasenames.has(path.posix.basename(relPath))) continue;
    if (!expected.has(relPath)) {
      return { ok: false, reason: `Unexpected file in extension directory: ${relPath}` };
    }
  }

  for (const file of actualFiles) {
    const expectedEntry = expected.get(file.relPath);
    if (!expectedEntry) continue;

    let stat;
    try {
      stat = await fs.stat(file.absPath);
    } catch (error) {
      if (error && error.code === "ENOENT") {
        return { ok: false, reason: `Missing expected file: ${file.relPath}` };
      }
      return { ok: false, reason: `Failed to stat file ${file.relPath}: ${error?.message ?? String(error)}` };
    }

    if (!stat.isFile()) {
      return { ok: false, reason: `Expected file is not a regular file: ${file.relPath}` };
    }

    if (stat.size !== expectedEntry.size) {
      return { ok: false, reason: `Size mismatch for ${file.relPath}` };
    }

    let bytes;
    try {
      bytes = await fs.readFile(file.absPath);
    } catch (error) {
      return { ok: false, reason: `Failed to read file ${file.relPath}: ${error?.message ?? String(error)}` };
    }

    const actualSha = sha256Hex(bytes);
    if (actualSha !== expectedEntry.sha256) {
      return { ok: false, reason: `Checksum mismatch for ${file.relPath}` };
    }
  }

  return { ok: true };
}

module.exports = {
  sha256Hex,
  verifyExtractedExtensionDir,
};
