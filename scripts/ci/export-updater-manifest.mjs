#!/usr/bin/env node
/**
 * Export the per-platform Tauri updater manifest (`latest.json`) generated during a
 * release build into a stable path so the workflow can upload it as an artifact.
 *
 * This is used to merge/update `latest.json` after parallel matrix builds complete,
 * avoiding "last writer wins" races when multiple jobs upload updater metadata.
 *
 * Usage:
 *   node scripts/ci/export-updater-manifest.mjs <output-path>
 */
import fs from "node:fs";
import path from "node:path";

/**
 * @param {string} message
 */
function fatal(message) {
  console.error(message);
  process.exit(1);
}

/**
 * @param {string} p
 */
function isDir(p) {
  try {
    return fs.statSync(p).isDirectory();
  } catch {
    return false;
  }
}

/**
 * @param {string} dir
 * @param {(p: string) => boolean} predicate
 * @param {string[]} out
 */
function walkSync(dir, predicate, out) {
  let entries;
  try {
    entries = fs.readdirSync(dir, { withFileTypes: true });
  } catch {
    return;
  }
  for (const ent of entries) {
    const full = path.join(dir, ent.name);
    if (ent.isDirectory()) {
      walkSync(full, predicate, out);
    } else if (ent.isFile()) {
      if (predicate(full)) out.push(full);
    }
  }
}

/**
 * Finds:
 * - <target>/release/bundle
 * - <target>/<triple>/release/bundle
 *
 * @param {string} targetDir
 * @returns {string[]}
 */
function findBundleDirs(targetDir) {
  /** @type {string[]} */
  const dirs = [];

  const native = path.join(targetDir, "release", "bundle");
  if (isDir(native)) dirs.push(native);

  let entries;
  try {
    entries = fs.readdirSync(targetDir, { withFileTypes: true });
  } catch {
    return dirs;
  }

  for (const ent of entries) {
    if (!ent.isDirectory()) continue;
    const maybe = path.join(targetDir, ent.name, "release", "bundle");
    if (isDir(maybe)) dirs.push(maybe);
  }

  return dirs.slice().sort();
}

/**
 * @param {string[]} files
 */
function pickBestManifest(files) {
  return files.slice().sort()[0] ?? null;
}

function main() {
  const outPath = process.argv[2];
  if (!outPath) {
    fatal("Missing output path. Usage: node scripts/ci/export-updater-manifest.mjs <output-path>");
  }

  const repoRoot = process.cwd();
  const candidates = [
    path.join(repoRoot, "apps", "desktop", "src-tauri", "target"),
    path.join(repoRoot, "target"),
  ].filter(isDir);

  if (candidates.length === 0) {
    fatal(
      `No candidate target directories found (expected apps/desktop/src-tauri/target or target). Repo root: ${repoRoot}`,
    );
  }

  /** @type {string[]} */
  const found = [];
  for (const dir of candidates) {
    const bundleDirs = findBundleDirs(dir);
    for (const bundleDir of bundleDirs) {
      walkSync(bundleDir, (p) => path.basename(p) === "latest.json", found);
    }
  }

  const best = pickBestManifest(found);
  if (!best) {
    fatal(
      `No latest.json found under: ${candidates.map((c) => c.split(path.sep).join("/")).join(", ")}`,
    );
  }

  fs.mkdirSync(path.dirname(outPath), { recursive: true });
  fs.copyFileSync(best, outPath);

  console.log(`export-updater-manifest: copied ${best} -> ${outPath}`);
}

main();
