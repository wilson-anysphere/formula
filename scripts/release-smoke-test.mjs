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

import { spawn } from "node:child_process";
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
  node scripts/release-smoke-test.mjs --tag vX.Y.Z [--repo owner/name] [--token <token>] [--local-bundles]

Options:
  --tag            Required. Release tag (example: v0.2.3).
  --repo           GitHub repo in owner/name form. Defaults to:
                     - $GITHUB_REPOSITORY (if set)
                     - or inferred from git remote "origin" (if possible)
  --token          GitHub token. Defaults to $GITHUB_TOKEN (if set).
  --local-bundles  Also run any platform-specific local bundle validators (if present).
  -h, --help       Print this help.
`;
  console.log(usage.trimEnd());
}

/**
 * Minimal argv parser (no deps).
 * @param {string[]} argv
 */
function parseArgs(argv) {
  /** @type {Record<string, string | boolean>} */
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
    const scriptPath = step.args[0];
    if (step.skipIfMissing && typeof scriptPath === "string" && scriptPath.endsWith(".mjs")) {
      if (!existsSync(scriptPath)) {
        resolve({
          step,
          status: "skip",
          exitCode: null,
          reason: `Missing file: ${path.relative(repoRoot, scriptPath)}`,
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
  for (const entry of entries) {
    if (!entry.isFile()) continue;
    const name = entry.name;
    if (!name.endsWith(".mjs")) continue;

    // Target only bundle validators (avoid running unrelated scripts).
    const isValidator =
      name.startsWith("validate-") &&
      (name.includes("bundle") || name.includes("bundles") || name.includes("installer"));
    if (!isValidator) continue;

    // Platform filter: run scripts explicitly scoped to this OS, and also any
    // generic desktop bundle validators (e.g. validate-desktop-bundles.mjs).
    const lower = name.toLowerCase();
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
  const fallbacks = [
    `validate-desktop-bundle-${key}.mjs`,
    `validate-desktop-bundles-${key}.mjs`,
    `validate-${key}-bundle.mjs`,
    `validate-${key}-bundles.mjs`,
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
      : (process.env.GITHUB_TOKEN ?? "").trim();

  if (!token) {
    console.warn(
      "Warning: no GitHub token provided (use --token or set GITHUB_TOKEN). Public releases may work without a token, but you may hit rate limits."
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
      ],
      env: token ? { ...process.env, GITHUB_TOKEN: token } : process.env,
      skipIfMissing: false,
    },
  ]);

  const localBundlesRequested = args.localBundles === true;
  if (localBundlesRequested) {
    const key = platformKey(process.platform);
    const scriptsDir = path.join(repoRoot, "scripts");
    const bundleDir = path.join(
      repoRoot,
      "apps",
      "desktop",
      "src-tauri",
      "target",
      "release",
      "bundle"
    );

    if (!existsSync(bundleDir)) {
      steps.push({
        id: "local-bundles",
        title: `Local bundle validation (${key})`,
        command: process.execPath,
        args: [path.join(repoRoot, "scripts", "nonexistent-local-bundles-placeholder.mjs")],
        skipIfMissing: true,
      });
      // We'll override the placeholder to a skip result via summary.
    } else {
      const validators = await discoverLocalBundleValidators(scriptsDir, key);
      if (validators.length === 0) {
        steps.push({
          id: "local-bundles",
          title: `Local bundle validation (${key})`,
          command: process.execPath,
          args: [path.join(repoRoot, "scripts", "nonexistent-local-bundles-placeholder.mjs")],
          skipIfMissing: true,
        });
      } else {
        for (const validator of validators) {
          steps.push({
            id: `local-bundles:${path.basename(validator)}`,
            title: `Validate local bundles (${key}): ${path.basename(validator)}`,
            command: process.execPath,
            args: [validator],
            skipIfMissing: true,
          });
        }
      }
    }
  }

  /** @type {StepResult[]} */
  const results = [];
  for (const step of steps) {
    // Special-case: our placeholder "local bundles" step is always a skip, but
    // we want a helpful reason string.
    if (
      step.id === "local-bundles" &&
      step.args[0]?.includes("nonexistent-local-bundles-placeholder.mjs")
    ) {
      const key = platformKey(process.platform);
      const bundleDir = path.join(
        repoRoot,
        "apps",
        "desktop",
        "src-tauri",
        "target",
        "release",
        "bundle"
      );
      const exists = existsSync(bundleDir);
      results.push({
        step,
        status: "skip",
        exitCode: null,
        reason: !exists
          ? `No local bundles found at ${path.relative(repoRoot, bundleDir)} (build with: cd apps/desktop && bash ../../scripts/cargo_agent.sh tauri build)`
          : `No local bundle validator scripts found for platform "${key}" in scripts/ (expected from bundle validator tasks).`,
      });
      continue;
    }

    results.push(await runStep(step));
  }

  printSummary(results);

  const failed = results.some((r) => r.status === "fail");
  process.exitCode = failed ? 1 : 0;
}

await main();
