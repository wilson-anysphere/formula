#!/usr/bin/env node
/**
 * CI guardrail for desktop releases.
 *
 * Why this exists:
 * - The release workflow relies on Tauri producing specific installer/update artifacts.
 * - It's easy for a Tauri upgrade/config change to silently stop emitting one of them.
 * - Missing artifacts only show up after tagging a release unless CI enforces expectations.
 *
 * This script scans for bundle outputs under the Cargo target directory, e.g.:
 *   target/release/bundle
 *   target/<target-triple>/release/bundle
 *
 * It then asserts that:
 * - Per-OS required artifacts exist (installers + updater archives) and have matching `.sig` files
 * - Updater metadata exists: `latest.json` + `latest.json.sig`
 *
 * Fork/dry-run behavior:
 * - On forks (or any run) where updater signing secrets are not configured, CI may choose to
 *   validate that *installer artifacts* exist without enforcing updater signature files (`*.sig`).
 * - This is controlled via `FORMULA_REQUIRE_TAURI_UPDATER_SIGNATURES` and
 *   `FORMULA_HAS_TAURI_UPDATER_KEY` env vars (see `.github/workflows/release.yml`).
 *
 * Usage:
 *   node scripts/ci/check-desktop-release-artifacts.mjs
 *   node scripts/ci/check-desktop-release-artifacts.mjs --os linux
 */

import fs from "node:fs";
import path from "node:path";
import process from "node:process";
import { fileURLToPath } from "node:url";

const repoRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "../..");

/**
 * @param {string} message
 */
function die(message) {
  console.error(message);
  process.exitCode = 1;
}

/**
 * @param {string} raw
 * @returns {"linux"|"windows"|"macos"|null}
 */
function normalizeOs(raw) {
  const val = raw.trim().toLowerCase();
  if (!val) return null;
  if (val === "macos" || val === "mac" || val === "darwin" || val === "osx") return "macos";
  if (val === "windows" || val === "win" || val === "win32") return "windows";
  if (val === "linux") return "linux";
  if (val === "macosx") return "macos";
  // GitHub Actions RUNNER_OS values:
  if (val === "macos") return "macos";
  if (val === "windows") return "windows";
  return null;
}

/**
 * @param {string} name
 * @returns {boolean | undefined}
 */
function envBool(name) {
  const raw = process.env[name];
  if (raw === undefined) return undefined;
  const val = raw.trim().toLowerCase();
  if (!val) return undefined;
  if (val === "1" || val === "true" || val === "yes") return true;
  if (val === "0" || val === "false" || val === "no") return false;
  return undefined;
}

/**
 * @param {string[]} argv
 * @returns {{ os: string | null, bundleDirs: string[] }}
 */
function parseArgs(argv) {
  /** @type {string | null} */
  let os = null;
  /** @type {string[]} */
  const bundleDirs = [];

  for (let i = 0; i < argv.length; i++) {
    const arg = argv[i];
    if (arg === "--help" || arg === "-h") {
      console.log(
        [
          "check-desktop-release-artifacts",
          "",
          "Usage:",
          "  node scripts/ci/check-desktop-release-artifacts.mjs [--os <linux|windows|macos>] [--bundle-dir <dir> ...]",
          "",
          "Options:",
          "  --os         Override the OS used for validation (otherwise uses RUNNER_OS).",
          "  --bundle-dir Explicit bundle directory to scan (can be specified multiple times).",
        ].join("\n"),
      );
      process.exit(0);
    }
    if (arg === "--os") {
      os = argv[i + 1] ?? null;
      i++;
      continue;
    }
    if (arg.startsWith("--os=")) {
      os = arg.slice("--os=".length);
      continue;
    }
    if (arg === "--bundle-dir") {
      const dir = argv[i + 1];
      if (!dir) {
        die("Missing value for --bundle-dir");
        return { os, bundleDirs };
      }
      bundleDirs.push(dir);
      i++;
      continue;
    }
    if (arg.startsWith("--bundle-dir=")) {
      bundleDirs.push(arg.slice("--bundle-dir=".length));
      continue;
    }

    die(`Unknown argument: ${arg}\nRun with --help for usage.`);
    return { os, bundleDirs };
  }
  return { os, bundleDirs };
}

/**
 * @param {string} p
 * @returns {boolean}
 */
function isDir(p) {
  try {
    return fs.statSync(p).isDirectory();
  } catch {
    return false;
  }
}

/**
 * @param {string} p
 * @returns {boolean}
 */
function isFile(p) {
  try {
    return fs.statSync(p).isFile();
  } catch {
    return false;
  }
}

/**
 * @param {string} p
 * @returns {string}
 */
function relPath(p) {
  const rel = path.relative(repoRoot, p);
  // Normalise for readability in logs on Windows.
  return rel.split(path.sep).join("/");
}

/**
 * @param {string} message
 * @param {string[]} details
 */
function dieBlock(message, details) {
  die(`\n${message}\n${details.map((d) => `  - ${d}`).join("\n")}\n`);
}

/**
 * @returns {string[]}
 */
function candidateTargetDirs() {
  /** @type {string[]} */
  const candidates = [];

  // Respect `CARGO_TARGET_DIR` if set, since some CI/caching setups override it.
  // Cargo interprets relative paths relative to the working directory used for the build.
  const cargoTargetDirEnv = process.env.CARGO_TARGET_DIR;
  const cargoTargetDir = typeof cargoTargetDirEnv === "string" ? cargoTargetDirEnv.trim() : "";
  if (cargoTargetDir !== "") {
    const resolved = path.isAbsolute(cargoTargetDir) ? cargoTargetDir : path.join(repoRoot, cargoTargetDir);
    if (isDir(resolved)) candidates.push(resolved);
  }

  // Common locations:
  // - workspace builds: <repo>/target
  // - standalone Tauri app builds: apps/desktop/src-tauri/target
  for (const p of [
    path.join(repoRoot, "apps", "desktop", "src-tauri", "target"),
    path.join(repoRoot, "apps", "desktop", "target"),
    path.join(repoRoot, "target"),
  ]) {
    if (isDir(p)) candidates.push(p);
  }

  if (candidates.length > 0) return dedupeRealpaths(candidates);

  // Fallback: search for src-tauri directories with tauri.conf.json (skip huge dirs).
  // This is a last resort: most CI builds should have one of the standard Cargo target
  // directories present (apps/desktop/src-tauri/target, apps/desktop/target, target, or
  // CARGO_TARGET_DIR).
  //
  // Keep the walk bounded so we don't traverse an entire monorepo tree (or extracted
  // build artifacts) when something goes wrong in CI.
  const maxDepth = 8;
  const skipDirNames = new Set([
    ".git",
    "node_modules",
    ".pnpm-store",
    ".turbo",
    ".cache",
    ".vite",
    "dist",
    "build",
    "coverage",
    "target",
    "security-report",
    "test-results",
    "playwright-report",
    ".cargo",
  ]);

  /** @type {string[]} */
  const stack = [{ dir: repoRoot, depth: 0 }];
  while (stack.length > 0) {
    const popped = stack.pop();
    const dir = popped?.dir;
    const depth = popped?.depth ?? maxDepth;
    if (!dir) break;
    let entries;
    try {
      entries = fs.readdirSync(dir, { withFileTypes: true });
    } catch {
      continue;
    }
    for (const ent of entries) {
      if (ent.isDirectory()) {
        if (skipDirNames.has(ent.name)) continue;
        const nextDepth = depth + 1;
        if (nextDepth <= maxDepth) {
          stack.push({ dir: path.join(dir, ent.name), depth: nextDepth });
        }
        continue;
      }
      if (ent.isFile() && ent.name === "tauri.conf.json") {
        const srcTauriDir = dir;
        const targetDir = path.join(srcTauriDir, "target");
        if (isDir(targetDir)) candidates.push(targetDir);
      }
    }
  }

  return dedupeRealpaths(candidates);
}

/**
 * @param {string[]} paths
 * @returns {string[]}
 */
function dedupeRealpaths(paths) {
  /** @type {string[]} */
  const uniq = [];
  /** @type {Set<string>} */
  const seen = new Set();
  for (const p of paths) {
    let key = p;
    try {
      key = fs.realpathSync(p);
    } catch {
      // ok
    }
    if (seen.has(key)) continue;
    seen.add(key);
    uniq.push(p);
  }
  return uniq;
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

  return dedupeRealpaths(dirs);
}

/**
 * @param {string[]} bundleDirs
 * @returns {{ allFiles: string[], byKind: Map<string, string[]>, fileSet: Set<string> }}
 */
function scanBundleDirs(bundleDirs) {
  /** @type {string[]} */
  const allFiles = [];
  /** @type {Map<string, string[]>} */
  const byKind = new Map();
  /** @type {Set<string>} */
  const fileSet = new Set();

  const skipDirNames = new Set(["node_modules", ".git"]);
  const skipDirSuffixes = [".app", ".dSYM", ".framework", ".xcframework"];

  /**
   * @param {string} filePath
   */
  function addFile(filePath) {
    const resolved = path.resolve(filePath);
    allFiles.push(resolved);
    fileSet.add(resolved);

    const kind = classifyArtifact(resolved);
    if (!kind) return;

    const arr = byKind.get(kind) ?? [];
    arr.push(resolved);
    byKind.set(kind, arr);
  }

  for (const bundleDir of bundleDirs) {
    /** @type {string[]} */
    const stack = [bundleDir];
    while (stack.length > 0) {
      const dir = stack.pop();
      if (!dir) break;
      let entries;
      try {
        entries = fs.readdirSync(dir, { withFileTypes: true });
      } catch {
        continue;
      }

      for (const ent of entries) {
        const full = path.join(dir, ent.name);
        if (ent.isSymbolicLink()) {
          // Avoid surprises/loops; bundle output should be plain files/directories.
          continue;
        }
        if (ent.isDirectory()) {
          if (skipDirNames.has(ent.name)) continue;
          if (skipDirSuffixes.some((s) => ent.name.endsWith(s))) continue;
          stack.push(full);
          continue;
        }
        if (ent.isFile()) {
          addFile(full);
        }
      }
    }
  }

  // Stabilize output (GitHub logs are easier to diff).
  allFiles.sort();
  for (const [kind, paths] of byKind.entries()) {
    paths.sort();
    byKind.set(kind, paths);
  }

  return { allFiles, byKind, fileSet };
}

/**
 * @param {string} filePath
 * @returns {string | null}
 */
function classifyArtifact(filePath) {
  const name = path.basename(filePath);

  if (name === "latest.json") return "latest.json";
  if (name === "latest.json.sig") return "latest.json.sig";

  if (name.endsWith(".dmg")) return "dmg";
  if (name.endsWith(".msi")) return "msi";
  if (name.endsWith(".exe")) return "exe";
  if (name.endsWith(".AppImage")) return "AppImage";
  if (name.endsWith(".deb")) return "deb";
  if (name.endsWith(".rpm")) return "rpm";
  if (name.endsWith(".tar.gz")) return "tar.gz";
  if (name.endsWith(".tgz")) return "tgz";
  if (name.endsWith(".sig")) return "sig";

  return null;
}

/**
 * @param {string[]} headers
 * @param {Array<Array<string | number>>} rows
 * @returns {string}
 */
function markdownTable(headers, rows) {
  const esc = (v) => String(v).replaceAll("|", "\\|");
  const header = `| ${headers.map(esc).join(" | ")} |`;
  const sep = `| ${headers.map(() => "---").join(" | ")} |`;
  const lines = rows.map((r) => `| ${r.map(esc).join(" | ")} |`);
  return [header, sep, ...lines].join("\n");
}

/**
 * @param {Map<string, string[]>} byKind
 * @returns {string}
 */
function renderDiscoveredArtifacts(byKind) {
  const kinds = [
    "latest.json",
    "latest.json.sig",
    "dmg",
    "msi",
    "exe",
    "AppImage",
    "deb",
    "rpm",
    "tar.gz",
    "tgz",
    "sig",
  ];

  /** @type {Array<Array<string>>} */
  const rows = [];
  for (const kind of kinds) {
    const paths = byKind.get(kind) ?? [];
    for (const p of paths) {
      rows.push([kind, relPath(p)]);
    }
  }

  if (rows.length === 0) {
    return "_No relevant artifacts discovered under bundle directories._";
  }

  return markdownTable(["Kind", "Path"], rows);
}

/**
 * @param {string} os
 * @returns {{ label: string, matchBase: (filePath: string) => boolean }[]}
 */
function requirementsForOs(os) {
  if (os === "macos") {
    return [
      { label: "macOS installer (.dmg)", matchBase: (p) => p.endsWith(".dmg") },
    ];
  }
  if (os === "windows") {
    const isWindowsBundleMsi = (p) => {
      // Only count Tauri-produced installers, not any random `.msi` that might appear under
      // `bundle/**` (or future helper tooling output).
      const normalized = p.split(path.sep).join("/").toLowerCase();
      return normalized.endsWith(".msi") && normalized.includes("/release/bundle/msi/");
    };

    const isWindowsBundleExe = (p) => {
      // Only count Tauri-produced NSIS installers under:
      // - bundle/nsis/*.exe
      // - bundle/nsis-web/*.exe
      //
      // Exclude embedded helper installers (notably WebView2 bootstrapper/runtime installers),
      // which may be present as standalone files in the bundle dir but are not the Formula
      // installer we ship on GitHub Releases.
      const normalized = p.split(path.sep).join("/").toLowerCase();
      if (!normalized.endsWith(".exe")) return false;
      if (
        !normalized.includes("/release/bundle/nsis/") &&
        !normalized.includes("/release/bundle/nsis-web/")
      ) {
        return false;
      }
      const base = path.basename(p).toLowerCase();
      if (base.startsWith("microsoftedgewebview2")) return false;
      return true;
    };

    return [
      { label: "Windows installer (.msi)", matchBase: isWindowsBundleMsi },
      { label: "Windows installer (.exe)", matchBase: isWindowsBundleExe },
    ];
  }
  if (os === "linux") {
    return [
      { label: "Linux bundle (.AppImage)", matchBase: (p) => p.endsWith(".AppImage") },
      { label: "Linux package (.deb)", matchBase: (p) => p.endsWith(".deb") },
      { label: "Linux package (.rpm)", matchBase: (p) => p.endsWith(".rpm") },
    ];
  }
  return [];
}

/**
 * @param {string} os
 * @param {{ allFiles: string[], byKind: Map<string, string[]>, fileSet: Set<string> }} scan
 * @param {{ requireUpdaterSignatures: boolean }} opts
 * @returns {{ ok: boolean, failures: string[], requirementRows: Array<Array<string | number>> }}
 */
function validate(os, scan, opts) {
  const requireUpdaterSignatures = opts.requireUpdaterSignatures;

  /** @type {string[]} */
  const failures = [];
  /** @type {Array<Array<string | number>>} */
  const requirementRows = [];

  /**
   * @param {string} label
   * @param {string[]} baseFiles
   */
  function validateBase(label, baseFiles) {
    if (baseFiles.length === 0) {
      failures.push(`Missing ${label}.`);
      requirementRows.push([label, "MISSING", 0, ""]);
      return;
    }

    if (!requireUpdaterSignatures) {
      requirementRows.push([label, "OK (unsigned)", baseFiles.length, relPath(baseFiles[0])]);
      return;
    }

    /** @type {string[]} */
    const missingSig = [];
    for (const base of baseFiles) {
      const sig = `${base}.sig`;
      if (!scan.fileSet.has(path.resolve(sig))) {
        missingSig.push(relPath(base));
      }
    }

    if (missingSig.length > 0) {
      failures.push(
        `${label} signature(s) missing for:\n${missingSig.map((p) => `  - ${p}`).join("\n")}`,
      );
      requirementRows.push([label, "MISSING .sig", baseFiles.length, relPath(baseFiles[0])]);
      return;
    }

    requirementRows.push([label, "OK", baseFiles.length, relPath(baseFiles[0])]);
  }

  // Updater metadata is only meaningful when updater signatures are enabled.
  if (requireUpdaterSignatures) {
    const latestJsonFiles = scan.allFiles.filter((p) => path.basename(p) === "latest.json");
    validateBase("Updater metadata (latest.json)", latestJsonFiles);
  } else {
    requirementRows.push(["Updater metadata (latest.json)", "SKIPPED", 0, ""]);
  }

  // OS-specific requirements.
  const osReqs = requirementsForOs(os);
  if (os === "macos" && requireUpdaterSignatures) {
    osReqs.push({
      label: "macOS updater archive (.app.tar.gz preferred; allow .tar.gz/.tgz)",
      matchBase: (p) => {
        const base = path.basename(p).toLowerCase();
        if (base.endsWith(".appimage.tar.gz") || base.endsWith(".appimage.tgz")) return false;
        return base.endsWith(".tar.gz") || base.endsWith(".tgz");
      },
    });
  }

  for (const req of osReqs) {
    const baseFiles = scan.allFiles.filter(req.matchBase);
    validateBase(req.label, baseFiles);
  }

  return { ok: failures.length === 0, failures, requirementRows };
}

function main() {
  const { os: osArg, bundleDirs: bundleDirsArg } = parseArgs(process.argv.slice(2));
  if (process.exitCode) return;

  const osFromEnv = process.env.RUNNER_OS?.trim() || "";
  const osFromPlatform = process.platform;

  const os =
    (osArg ? normalizeOs(osArg) : null) ??
    (osFromEnv ? normalizeOs(osFromEnv) : null) ??
    normalizeOs(osFromPlatform) ??
    null;

  if (!os) {
    dieBlock("Unable to determine OS for artifact validation.", [
      `RUNNER_OS=${JSON.stringify(process.env.RUNNER_OS ?? "")}`,
      `process.platform=${JSON.stringify(process.platform)}`,
      `Pass --os <linux|windows|macos> to override.`,
    ]);
    return;
  }

  /** @type {string[]} */
  let bundleDirs;
  if (bundleDirsArg.length > 0) {
    bundleDirs = bundleDirsArg.map((p) => path.resolve(process.cwd(), p));
  } else {
    const targetDirs = candidateTargetDirs();
    /** @type {string[]} */
    const found = [];
    for (const t of targetDirs) {
      found.push(...findBundleDirs(t));
    }
    bundleDirs = dedupeRealpaths(found);
  }

  if (bundleDirs.length === 0) {
    const targetDirs = candidateTargetDirs();
    dieBlock("No Tauri bundle directories found.", [
      `Expected at least one bundle dir like: target/release/bundle or target/<triple>/release/bundle`,
      `Repo root: ${repoRoot}`,
      `Candidate target dirs: ${targetDirs.length ? targetDirs.map(relPath).join(", ") : "(none found)"}`,
      `Tip: run this script after a Tauri release build, or pass --bundle-dir <dir>.`,
    ]);
    return;
  }

  const scan = scanBundleDirs(bundleDirs);

  // In the upstream repo we require updater signatures (so auto-update always works). For forks or
  // dry-run builds, CI may disable signature validation when the required secrets are not present.
  const requireUpdaterSignatures = envBool("FORMULA_REQUIRE_TAURI_UPDATER_SIGNATURES") ?? true;
  const hasUpdaterKey = envBool("FORMULA_HAS_TAURI_UPDATER_KEY") ?? false;

  if (requireUpdaterSignatures && !hasUpdaterKey) {
    dieBlock("Updater signature validation is required but TAURI_PRIVATE_KEY is not configured.", [
      `Set the TAURI_PRIVATE_KEY secret in GitHub Actions to enable updater signatures.`,
      `Or (forks/dry-runs) set FORMULA_REQUIRE_TAURI_UPDATER_SIGNATURES=false to skip signature validation.`,
    ]);
    return;
  }

  const result = validate(os, scan, { requireUpdaterSignatures });

  if (!result.ok) {
    console.error(`\nDesktop release artifact check failed (os=${os}).\n`);
    console.error("Bundle directories scanned:");
    for (const d of bundleDirs) console.error(`- ${relPath(d)}`);

    console.error("\nRequired artifacts:");
    console.error(markdownTable(["Requirement", "Status", "Count", "Example"], result.requirementRows));

    console.error("\nFailure details:");
    for (const msg of result.failures) {
      console.error(`- ${msg}`);
    }

    console.error("\nDiscovered artifacts:");
    console.error(renderDiscoveredArtifacts(scan.byKind));

    process.exitCode = 1;
    return;
  }

  console.log(`Desktop release artifact check passed (os=${os}).`);
  console.log("Bundle directories scanned:");
  for (const d of bundleDirs) console.log(`- ${relPath(d)}`);
  console.log("");
  console.log("Required artifacts:");
  console.log(markdownTable(["Requirement", "Status", "Count", "Example"], result.requirementRows));
}

main();
