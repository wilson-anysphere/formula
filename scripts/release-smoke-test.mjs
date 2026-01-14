#!/usr/bin/env node
/**
 * Release smoke test runner.
 *
 * Goal: give maintainers a single command to sanity-check a tag/release locally.
 *
 * Runs:
 *  - scripts/check-desktop-version.mjs <tag>
 *  - scripts/check-updater-config.mjs
 *  - scripts/verify-desktop-release-assets.mjs --tag <tag> --repo <repo>
 *
 * Optional:
 *  - With --local-bundles, runs any platform-specific local bundle validator scripts
 *    (if present) against locally-built Tauri bundles.
 */

import { spawn, spawnSync } from "node:child_process";
import { existsSync } from "node:fs";
import { readdir } from "node:fs/promises";
import path from "node:path";
import process from "node:process";
import { fileURLToPath } from "node:url";

const repoRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");

/**
 * @param {string} message
 */
function die(message) {
  console.error(message);
  process.exit(1);
}

function printUsage() {
  const usage = `
Release smoke test

Usage:
  node scripts/release-smoke-test.mjs --tag vX.Y.Z [--repo owner/name] [--token <token>] [--local-bundles] [expectations...]

Options:
  --tag            Required. Release tag (example: v0.2.3).
  --repo           GitHub repo in owner/name form. Defaults to:
                     - $GITHUB_REPOSITORY (if set)
                     - or inferred from git remote "origin" (if possible)
  --token          GitHub token. Defaults to $GITHUB_TOKEN / $GH_TOKEN (if set).
  --local-bundles  Also run any platform-specific local bundle validators (if present).

Verifier options (forwarded to scripts/verify-desktop-release-assets.mjs; optional):
  --dry-run                Validate manifest/assets only (skip bundle hashing)
  --verify-assets          Download updater assets referenced in latest.json and verify their signatures (slow)
  --out <path>             Output path for SHA256SUMS.txt (default: ./SHA256SUMS.txt)
  --all-assets             Hash all release assets (still excludes .sig by default)
  --include-sigs           Include .sig assets in SHA256SUMS (use with --all-assets to match CI)
  --allow-windows-msi      Allow raw .msi in latest.json Windows entries (defaults to disallowed)
  --allow-windows-exe      Allow raw .exe in latest.json Windows entries (defaults to disallowed)

Expectations (also forwarded; optional):
  --expectations <file>     Load expected targets from a JSON config file
  --expect-windows-x64
  --expect-windows-arm64
  --expect-macos-universal
  --expect-macos-x64
  --expect-macos-arm64
  --expect-linux-x64
  --expect-linux-arm64

  -h, --help       Print this help.
`;
  console.log(usage.trimEnd());
}

/**
 * Minimal argv parser (no deps).
 * @param {string[]} argv
 */
function parseArgs(argv) {
  /** @type {Record<string, unknown>} */
  const out = {};
  for (let i = 0; i < argv.length; i++) {
    const arg = argv[i];
    if (arg === "--help" || arg === "-h") {
      out.help = true;
      continue;
    }
    if (arg === "--local-bundles") {
      out.localBundles = true;
      continue;
    }
    if (arg === "--dry-run") {
      out.dryRun = true;
      continue;
    }
    if (arg === "--verify-assets") {
      out.verifyAssets = true;
      continue;
    }
    if (arg === "--all-assets") {
      out.allAssets = true;
      continue;
    }
    if (arg === "--include-sigs") {
      out.includeSigs = true;
      continue;
    }
    if (arg === "--allow-windows-msi") {
      out.allowWindowsMsi = true;
      continue;
    }
    if (arg === "--allow-windows-exe") {
      out.allowWindowsExe = true;
      continue;
    }
    if (arg === "--out" || arg.startsWith("--out=")) {
      const value = arg === "--out" ? argv[i + 1] : arg.slice("--out=".length);
      if (!value || value.startsWith("-")) {
        die(`Missing value for --out.\n\nRun with --help for usage.`);
      }
      out.out = value;
      if (arg === "--out") i++;
      continue;
    }
    if (arg === "--expectations" || arg.startsWith("--expectations=")) {
      const value =
        arg === "--expectations" ? argv[i + 1] : arg.slice("--expectations=".length);
      if (!value || value.startsWith("-")) {
        die(`Missing value for --expectations.\n\nRun with --help for usage.`);
      }
      out.expectations = value;
      if (arg === "--expectations") i++;
      continue;
    }
    if (arg.startsWith("--expect-")) {
      const list = Array.isArray(out.expectFlags) ? out.expectFlags : [];
      list.push(arg);
      out.expectFlags = list;
      continue;
    }
    if (arg === "--tag" || arg === "--repo" || arg === "--token") {
      const value = argv[i + 1];
      if (!value || value.startsWith("-")) {
        die(`Missing value for ${arg}.\n\nRun with --help for usage.`);
      }
      if (arg === "--tag") out.tag = value;
      if (arg === "--repo") out.repo = value;
      if (arg === "--token") out.token = value;
      i++;
      continue;
    }
    die(`Unknown argument: ${arg}\n\nRun with --help for usage.`);
  }
  return out;
}

/**
 * @param {string} maybeRepo
 */
function normalizeRepo(maybeRepo) {
  const trimmed = maybeRepo.trim();
  if (!trimmed) return "";
  // allow "owner/name" only (no protocol)
  if (!/^[A-Za-z0-9_.-]+\/[A-Za-z0-9_.-]+$/.test(trimmed)) return "";
  return trimmed;
}

/**
 * @param {string} remoteUrl
 */
function parseRepoFromRemoteUrl(remoteUrl) {
  const trimmed = remoteUrl.trim();
  if (!trimmed) return "";

  // Examples:
  // - https://github.com/owner/name.git
  // - git@github.com:owner/name.git
  // - ssh://git@github.com/owner/name.git
  const httpsMatch = trimmed.match(/^https?:\/\/github\.com\/([^/]+)\/([^/]+?)(?:\.git)?$/i);
  if (httpsMatch) return normalizeRepo(`${httpsMatch[1]}/${httpsMatch[2]}`);

  const sshMatch = trimmed.match(/^git@github\.com:([^/]+)\/([^/]+?)(?:\.git)?$/i);
  if (sshMatch) return normalizeRepo(`${sshMatch[1]}/${sshMatch[2]}`);

  const sshProtoMatch = trimmed.match(/^ssh:\/\/git@github\.com\/([^/]+)\/([^/]+?)(?:\.git)?$/i);
  if (sshProtoMatch) return normalizeRepo(`${sshProtoMatch[1]}/${sshProtoMatch[2]}`);

  return "";
}

/**
 * @returns {Promise<string>}
 */
async function detectDefaultRepo() {
  const envRepo = normalizeRepo(process.env.GITHUB_REPOSITORY ?? "");
  if (envRepo) return envRepo;

  try {
    const remoteUrl = await runAndCapture("git", ["remote", "get-url", "origin"], {
      cwd: repoRoot,
    });
    const parsed = parseRepoFromRemoteUrl(remoteUrl);
    if (parsed) return parsed;
  } catch {
    // ignore
  }

  return "";
}

/**
 * @typedef {Object} RunCaptureOpts
 * @property {string} [cwd]
 * @property {NodeJS.ProcessEnv} [env]
 */

/**
 * Run a command and capture stdout (trimmed).
 * @param {string} command
 * @param {string[]} args
 * @param {RunCaptureOpts} [opts]
 * @returns {Promise<string>}
 */
function runAndCapture(command, args, opts = {}) {
  return new Promise((resolve, reject) => {
    const child = spawn(command, args, {
      cwd: opts.cwd,
      env: opts.env,
      stdio: ["ignore", "pipe", "pipe"],
    });
    /** @type {Buffer[]} */
    const out = [];
    /** @type {Buffer[]} */
    const err = [];
    child.stdout.on("data", (d) => out.push(Buffer.from(d)));
    child.stderr.on("data", (d) => err.push(Buffer.from(d)));
    child.on("error", reject);
    child.on("close", (code) => {
      if (code === 0) {
        resolve(Buffer.concat(out).toString("utf8").trim());
      } else {
        const msg =
          Buffer.concat(err).toString("utf8").trim() ||
          Buffer.concat(out).toString("utf8").trim() ||
          `Command failed: ${command} ${args.join(" ")} (exit ${code ?? "unknown"})`;
        reject(new Error(msg));
      }
    });
  });
}

/**
 * @typedef {Object} Step
 * @property {string} id
 * @property {string} title
 * @property {string} command
 * @property {string[]} args
 * @property {NodeJS.ProcessEnv} [env]
 * @property {boolean} [skipIfMissing]
 * @property {string} [fileToCheck]
 * @property {string} [skipReason]
 */

/**
 * @typedef {Object} StepResult
 * @property {Step} step
 * @property {"pass" | "fail" | "skip"} status
 * @property {number | null} exitCode
 * @property {string} [reason]
 */

/**
 * @param {Step} step
 * @returns {Promise<StepResult>}
 */
function runStep(step) {
  return new Promise((resolve) => {
    if (step.skipReason) {
      resolve({ step, status: "skip", exitCode: null, reason: step.skipReason });
      return;
    }

    const checkPath = step.fileToCheck ?? step.args[0];
    if (step.skipIfMissing && typeof checkPath === "string" && checkPath.length > 0) {
      if (!existsSync(checkPath)) {
        resolve({
          step,
          status: "skip",
          exitCode: null,
          reason: `Missing file: ${path.relative(repoRoot, checkPath)}`,
        });
        return;
      }
    }

    console.log(`\n=== ${step.title} ===`);
    const child = spawn(step.command, step.args, {
      cwd: repoRoot,
      env: step.env ?? process.env,
      stdio: "inherit",
    });
    child.on("error", (err) => {
      resolve({
        step,
        status: "fail",
        exitCode: 1,
        reason: err instanceof Error ? err.message : String(err),
      });
    });
    child.on("close", (code) => {
      if (code === 0) {
        resolve({ step, status: "pass", exitCode: 0 });
      } else {
        resolve({
          step,
          status: "fail",
          exitCode: typeof code === "number" ? code : 1,
        });
      }
    });
  });
}

/**
 * @param {string} platform
 */
function platformKey(platform) {
  if (platform === "darwin") return "macos";
  if (platform === "win32") return "windows";
  if (platform === "linux") return "linux";
  return platform;
}

function pickPowerShellCommand() {
  // Prefer PowerShell 7 (`pwsh`) when available; fall back to Windows PowerShell
  // (`powershell`) on older environments.
  for (const cmd of ["pwsh", "powershell"]) {
    const res = spawnSync(cmd, ["-NoProfile", "-Command", "$PSVersionTable.PSVersion.Major"], {
      stdio: "ignore",
    });
    if (!res.error) return cmd;
  }
  return "pwsh";
}

/**
 * Best-effort check: do we appear to have any local Tauri `release/bundle` output directories?
 * This avoids surprising `--local-bundles` failures when the user hasn't built anything yet.
 */
async function detectLocalBundleOutputs() {
  /** @type {string[]} */
  const roots = [];

  if (process.env.CARGO_TARGET_DIR) {
    const raw = process.env.CARGO_TARGET_DIR.trim();
    if (raw) {
      roots.push(path.isAbsolute(raw) ? raw : path.join(repoRoot, raw));
    }
  }

  roots.push(
    path.join(repoRoot, "apps", "desktop", "src-tauri", "target"),
    path.join(repoRoot, "apps", "desktop", "target"),
    path.join(repoRoot, "target")
  );

  for (const root of roots) {
    if (!existsSync(root)) continue;

    if (existsSync(path.join(root, "release", "bundle"))) return true;

    // Tauri sometimes nests by target triple:
    //   <target>/<triple>/release/bundle/...
    try {
      const children = await readdir(root, { withFileTypes: true });
      for (const child of children) {
        if (!child.isDirectory()) continue;
        if (existsSync(path.join(root, child.name, "release", "bundle"))) return true;
      }
    } catch {
      // ignore
    }
  }

  return false;
}

/**
 * @param {string} validatorPath
 * @param {string} key
 * @returns {Step}
 */
function makeValidatorStep(validatorPath, key) {
  const base = path.basename(validatorPath);
  const ext = path.extname(base).toLowerCase();

  if (ext === ".mjs") {
    return {
      id: `local-bundles:${base}`,
      title: `Validate local bundles (${key}): ${base}`,
      command: process.execPath,
      args: [validatorPath],
      skipIfMissing: true,
      fileToCheck: validatorPath,
    };
  }

  if (ext === ".sh") {
    return {
      id: `local-bundles:${base}`,
      title: `Validate local bundles (${key}): ${base}`,
      command: "bash",
      args: [validatorPath],
      skipIfMissing: true,
      fileToCheck: validatorPath,
    };
  }

  if (ext === ".ps1") {
    const pwsh = pickPowerShellCommand();
    return {
      id: `local-bundles:${base}`,
      title: `Validate local bundles (${key}): ${base}`,
      command: pwsh,
      args: ["-NoProfile", "-ExecutionPolicy", "Bypass", "-File", validatorPath],
      skipIfMissing: true,
      fileToCheck: validatorPath,
    };
  }

  return {
    id: `local-bundles:${base}`,
    title: `Validate local bundles (${key}): ${base}`,
    command: process.execPath,
    args: [],
    skipReason: `Unsupported validator script type: ${base}`,
  };
}

/**
 * Discover local bundle validator scripts for the current platform.
 *
 * This is intentionally permissive: the exact validator script names may evolve,
 * but we want `--local-bundles` to "just work" as new validators land.
 *
 * @param {string} scriptsDir
 * @param {string} key
 */
async function discoverLocalBundleValidators(scriptsDir, key) {
  /** @type {string[]} */
  const discovered = [];
  const entries = await readdir(scriptsDir, { withFileTypes: true });

  const allowedExt = new Set([".mjs", ".sh", ".ps1"]);
  const keywords = [
    "bundle",
    "bundles",
    "installer",
    "appimage",
    "dmg",
    "msi",
    "nsis",
    "rpm",
    "deb",
  ];

  for (const entry of entries) {
    if (!entry.isFile()) continue;
    const name = entry.name;
    const ext = path.extname(name).toLowerCase();
    if (!allowedExt.has(ext)) continue;

    const lower = name.toLowerCase();
    if (!lower.startsWith("validate-")) continue;
    if (!keywords.some((k) => lower.includes(k))) continue;

    // Platform filter: run scripts explicitly scoped to this OS, and also any
    // generic desktop bundle validators (e.g. validate-desktop-bundles.mjs).
    const platformSpecific =
      lower.includes(key) ||
      (key === "macos" && lower.includes("darwin")) ||
      (key === "windows" && lower.includes("win"));

    const generic = lower.includes("desktop") && (lower.includes("bundle") || lower.includes("bundles"));

    if (platformSpecific || generic) {
      discovered.push(path.join(scriptsDir, name));
    }
  }

  // Stable, explicit fallback names (if the discovery heuristic misses).
  const explicitFallbacks =
    key === "macos"
      ? ["validate-macos-bundle.sh"]
      : key === "windows"
        ? ["validate-windows-bundles.ps1"]
        : key === "linux"
          ? ["validate-linux-appimage.sh", "validate-linux-rpm.sh"]
          : [];

  const fallbacks = [
    ...explicitFallbacks,
    `validate-desktop-bundle-${key}.mjs`,
    `validate-desktop-bundle-${key}.sh`,
    `validate-desktop-bundle-${key}.ps1`,
    `validate-desktop-bundles-${key}.mjs`,
    `validate-desktop-bundles-${key}.sh`,
    `validate-desktop-bundles-${key}.ps1`,
    `validate-${key}-bundle.mjs`,
    `validate-${key}-bundle.sh`,
    `validate-${key}-bundle.ps1`,
    `validate-${key}-bundles.mjs`,
    `validate-${key}-bundles.sh`,
    `validate-${key}-bundles.ps1`,
  ]
    .map((n) => path.join(scriptsDir, n))
    .filter((p) => existsSync(p));

  const out = Array.from(new Set([...discovered, ...fallbacks]));
  out.sort();
  return out;
}

/**
 * @param {StepResult[]} results
 */
function printSummary(results) {
  console.log("\n=== Summary ===");

  const maxTitle = Math.min(
    50,
    results.reduce((m, r) => Math.max(m, r.step.title.length), 0)
  );

  for (const r of results) {
    const statusLabel = r.status.toUpperCase().padEnd(4);
    const title = r.step.title.length > maxTitle ? `${r.step.title.slice(0, maxTitle - 1)}…` : r.step.title;
    const exit = r.exitCode === null ? "" : ` (exit ${r.exitCode})`;
    const reason = r.reason ? `\n    ${r.reason}` : "";
    console.log(`[${statusLabel}] ${title}${exit}${reason}`);
  }

  const failed = results.filter((r) => r.status === "fail");
  if (failed.length > 0) {
    console.log(`\nRelease smoke test FAILED (${failed.length} failing step${failed.length === 1 ? "" : "s"}).`);
  } else {
    console.log("\nRelease smoke test PASSED.");
  }
}

async function main() {
  const args = parseArgs(process.argv.slice(2));
  if (args.help) {
    printUsage();
    return;
  }

  const tag = typeof args.tag === "string" ? args.tag : "";
  if (!tag) {
    printUsage();
    die("\nError: --tag is required.");
  }

  const explicitRepo = typeof args.repo === "string" ? normalizeRepo(args.repo) : "";
  const repo = explicitRepo || (await detectDefaultRepo());
  if (!repo) {
    die(
      `Missing --repo (owner/name).\n\n` +
        `Set --repo explicitly, or ensure $GITHUB_REPOSITORY is set, or ensure git remote "origin" points at GitHub.`
    );
  }

  const token =
    typeof args.token === "string" && args.token.trim().length > 0
      ? args.token.trim()
      : (process.env.GITHUB_TOKEN ?? process.env.GH_TOKEN ?? "").trim();

  if (!token) {
    console.warn(
      "Warning: no GitHub token provided. scripts/verify-desktop-release-assets.mjs requires GITHUB_TOKEN/GH_TOKEN; the GitHub release asset verification step will fail.\nSet GITHUB_TOKEN=... (recommended) or pass --token ... to run the full smoke test."
    );
  }

  const steps = /** @type {Step[]} */ ([
    {
      id: "desktop-version",
      title: "Check desktop version matches tag",
      command: process.execPath,
      args: [path.join(repoRoot, "scripts", "check-desktop-version.mjs"), tag],
    },
    {
      id: "updater-config",
      title: "Check updater config",
      command: process.execPath,
      args: [path.join(repoRoot, "scripts", "check-updater-config.mjs")],
    },
    {
      id: "release-assets",
      title: "Verify GitHub release assets + manifests",
      command: process.execPath,
      args: [
        path.join(repoRoot, "scripts", "verify-desktop-release-assets.mjs"),
        "--tag",
        tag,
        "--repo",
        repo,
        ...(args.dryRun === true ? ["--dry-run"] : []),
        ...(args.verifyAssets === true ? ["--verify-assets"] : []),
        ...(typeof args.out === "string" && args.out.trim().length > 0 ? ["--out", args.out.trim()] : []),
        ...(args.allAssets === true ? ["--all-assets"] : []),
        ...(args.includeSigs === true ? ["--include-sigs"] : []),
        ...(args.allowWindowsMsi === true ? ["--allow-windows-msi"] : []),
        ...(args.allowWindowsExe === true ? ["--allow-windows-exe"] : []),
        ...(typeof args.expectations === "string" && args.expectations.trim().length > 0
          ? ["--expectations", args.expectations.trim()]
          : []),
        ...(Array.isArray(args.expectFlags)
          ? args.expectFlags.map((v) => String(v)).filter((v) => v.startsWith("--"))
          : []),
      ],
      env: token ? { ...process.env, GITHUB_TOKEN: token, GH_TOKEN: token } : process.env,
      skipIfMissing: false,
    },
  ]);

  const localBundlesRequested = args.localBundles === true;
  if (localBundlesRequested) {
    const key = platformKey(process.platform);
    const scriptsDir = path.join(repoRoot, "scripts");

    const hasBundles = await detectLocalBundleOutputs();
    if (!hasBundles) {
      steps.push({
        id: "local-bundles",
        title: `Local bundle validation (${key})`,
        command: process.execPath,
        args: [],
        skipReason:
          "No local Tauri bundles found (expected output under apps/desktop/src-tauri/target/**/release/bundle). Build with: cd apps/desktop && bash ../../scripts/cargo_agent.sh tauri build",
      });
    } else {
      const validators = await discoverLocalBundleValidators(scriptsDir, key);
      if (validators.length === 0) {
        steps.push({
          id: "local-bundles",
          title: `Local bundle validation (${key})`,
          command: process.execPath,
          args: [],
          skipReason: `No local bundle validator scripts found for platform "${key}" in scripts/.`,
        });
      } else {
        for (const validator of validators) {
          steps.push(makeValidatorStep(validator, key));
        }
      }
    }
  }

  /** @type {StepResult[]} */
  const results = [];
  for (const step of steps) {
    results.push(await runStep(step));
  }

  printSummary(results);

  const failed = results.some((r) => r.status === "fail");
  process.exitCode = failed ? 1 : 0;
}

await main();
