const fs = require("node:fs/promises");
const path = require("node:path");

async function installExtensionFromDirectory(sourceDir, installRoot) {
  if (!sourceDir) throw new Error("sourceDir is required");
  if (!installRoot) throw new Error("installRoot is required");

  const manifestPath = path.join(sourceDir, "package.json");
  const raw = await fs.readFile(manifestPath, "utf8");
  const manifest = JSON.parse(raw);
  const extensionId = `${manifest.publisher}.${manifest.name}`;

  const destDir = path.join(installRoot, extensionId);
  await fs.mkdir(installRoot, { recursive: true });
  await fs.rm(destDir, { recursive: true, force: true });
  await fs.cp(sourceDir, destDir, { recursive: true });

  return destDir;
}

async function uninstallExtension(installRoot, extensionId) {
  if (!installRoot) throw new Error("installRoot is required");
  if (!extensionId) throw new Error("extensionId is required");
  await fs.rm(path.join(installRoot, extensionId), { recursive: true, force: true });
}

async function listInstalledExtensions(installRoot) {
  try {
    const entries = await fs.readdir(installRoot, { withFileTypes: true });
    return entries.filter((e) => e.isDirectory()).map((e) => e.name);
  } catch {
    return [];
  }
}

module.exports = {
  installExtensionFromDirectory,
  uninstallExtension,
  listInstalledExtensions
};

