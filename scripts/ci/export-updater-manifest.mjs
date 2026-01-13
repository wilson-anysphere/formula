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
  const rootManifest = path.join(repoRoot, "latest.json");

  /**
   * `tauri-apps/tauri-action` writes `latest.json` to the workflow working directory (the repo
   * root in GitHub Actions) before uploading it to the GitHub Release.
   *
   * Prefer that file when present, but keep scanning Cargo bundle directories as a fallback in
   * case upstream tooling/layouts change.
   */
  if (fs.existsSync(rootManifest)) {
    fs.mkdirSync(path.dirname(outPath), { recursive: true });
    fs.copyFileSync(rootManifest, outPath);
    console.log(`export-updater-manifest: copied ${rootManifest} -> ${outPath}`);
    return;
  }
  /** @type {string[]} */
  const candidates = [];

  const cargoTargetDir = process.env.CARGO_TARGET_DIR?.trim() ?? "";
  if (cargoTargetDir) {
    candidates.push(path.isAbsolute(cargoTargetDir) ? cargoTargetDir : path.join(repoRoot, cargoTargetDir));
  }

  candidates.push(
    path.join(repoRoot, "apps", "desktop", "src-tauri", "target"),
    path.join(repoRoot, "apps", "desktop", "target"),
    path.join(repoRoot, "target"),
  );

  const candidateDirs = candidates.filter(isDir);

  if (candidateDirs.length === 0) {
    fatal(
      `No candidate Cargo target directories found. Looked in: CARGO_TARGET_DIR=${cargoTargetDir || "(unset)"} apps/desktop/src-tauri/target apps/desktop/target target. Repo root: ${repoRoot}`,
    );
  }

  /** @type {string[]} */
  const found = [];
  for (const dir of candidateDirs) {
    const bundleDirs = findBundleDirs(dir);
    for (const bundleDir of bundleDirs) {
      walkSync(bundleDir, (p) => path.basename(p) === "latest.json", found);
    }
  }

  const best = pickBestManifest(found);
  if (!best) {
    fatal(
      `No latest.json found under: ${candidateDirs.map((c) => c.split(path.sep).join("/")).join(", ")}`,
    );
  }

  fs.mkdirSync(path.dirname(outPath), { recursive: true });
  fs.copyFileSync(best, outPath);

  console.log(`export-updater-manifest: copied ${best} -> ${outPath}`);
}

main();
