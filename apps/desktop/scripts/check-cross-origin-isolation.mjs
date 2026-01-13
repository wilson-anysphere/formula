import { spawn } from "node:child_process";
import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const desktopDir = path.resolve(__dirname, "..");
const repoRoot = path.resolve(desktopDir, "../..");

const DESKTOP_TAURI_PACKAGE = "formula-desktop-tauri";
const DESKTOP_BINARY_NAME = "formula-desktop";

function isTruthyEnv(val) {
  if (!val) return false;
  return ["1", "true", "yes", "y", "on"].includes(val.trim().toLowerCase());
}

function cargoTargetDir() {
  // Respect `CARGO_TARGET_DIR` if set, since some developer/CI environments override it
  // for caching. Cargo interprets relative paths relative to the working directory used
  // for `cargo build`, which this script sets to `repoRoot`.
  const targetDir = process.env.CARGO_TARGET_DIR;
  if (!targetDir || targetDir.trim() === "") return path.join(repoRoot, "target");
  return path.isAbsolute(targetDir) ? targetDir : path.join(repoRoot, targetDir);
}

function run(
  cmd,
  args,
  { cwd = repoRoot, env = process.env, shell = process.platform === "win32" } = {},
) {
  return new Promise((resolve, reject) => {
    const child = spawn(cmd, args, {
      cwd,
      env,
      stdio: "inherit",
      // Allow running pnpm/cargo on Windows without needing `.cmd` suffixes.
      shell,
    });
    child.on("error", reject);
    child.on("exit", (code, signal) => {
      if (signal) {
        reject(new Error(`${cmd} exited with signal ${signal}`));
        return;
      }
      resolve(code ?? 1);
    });
  });
}

function usage() {
  console.log(`Usage: pnpm -C apps/desktop check:coi [-- --no-build] [-- --bin <path>]

Options:
  --no-build            Run the COI smoke check against already-built artifacts.
                        Also enabled via FORMULA_COI_NO_BUILD=1.
  --bin <path>          Path to an existing built desktop binary to run.
                        Useful with --no-build when auto-detection picks the wrong binary.
  --help                Show this help.

Examples:
  # Default (build + run)
  pnpm -C apps/desktop check:coi

  # CI/release (run-only, uses artifacts built by cargo tauri build / tauri-action)
  pnpm -C apps/desktop check:coi -- --no-build

  # Explicit binary override
  pnpm -C apps/desktop check:coi -- --no-build --bin target/release/formula-desktop
`);
}

function desktopBinaryPath() {
  const exe = process.platform === "win32" ? `${DESKTOP_BINARY_NAME}.exe` : DESKTOP_BINARY_NAME;
  return path.join(cargoTargetDir(), "release", exe);
}

function resolveBinPath(raw) {
  if (path.isAbsolute(raw)) return raw;

  // Prefer resolving relative to the repo root (matches the auto-detection paths), but also
  // accept paths relative to the current working directory for convenience.
  const fromCwd = path.resolve(process.cwd(), raw);
  if (fs.existsSync(fromCwd)) return fromCwd;
  return path.resolve(repoRoot, raw);
}

function statIsFile(p) {
  try {
    return fs.statSync(p).isFile();
  } catch {
    return false;
  }
}

function maybeAddCandidate(candidates, p) {
  if (!p) return;
  const key = path.resolve(p);
  if (candidates.some((c) => path.resolve(c) === key)) return;
  candidates.push(p);
}

function findBinaryInTargetDir(targetDir, exeName) {
  /** @type {string[]} */
  const candidates = [];
  if (!targetDir || !fs.existsSync(targetDir)) return candidates;
  try {
    if (!fs.statSync(targetDir).isDirectory()) return candidates;
  } catch {
    return candidates;
  }

  // Direct: target/release/<exe>
  maybeAddCandidate(candidates, path.join(targetDir, "release", exeName));

  // Bounded search: target/*/release/<exe>
  let entries = [];
  try {
    entries = fs.readdirSync(targetDir, { withFileTypes: true });
  } catch {
    return candidates;
  }

  for (const entry of entries) {
    if (!entry.isDirectory()) continue;
    // Skip very common non-target-triple directories to avoid pointless fs calls.
    if (entry.name === "release" || entry.name === "debug") continue;
    maybeAddCandidate(candidates, path.join(targetDir, entry.name, "release", exeName));
  }

  return candidates;
}

function pickMostRecentBinary(paths) {
  let best = null;
  let bestMtime = -1;
  for (const p of paths) {
    if (!statIsFile(p)) continue;
    try {
      const mtime = fs.statSync(p).mtimeMs;
      if (mtime > bestMtime) {
        bestMtime = mtime;
        best = p;
      }
    } catch {
      // Ignore.
    }
  }
  return best;
}

function detectBuiltDesktopBinary() {
  const exe = process.platform === "win32" ? `${DESKTOP_BINARY_NAME}.exe` : DESKTOP_BINARY_NAME;

  /** @type {string[]} */
  const candidates = [];

  // Common locations:
  // - workspace build: target/release
  // - standalone Tauri app: apps/desktop/src-tauri/target/release
  // - builds invoked from `apps/desktop`: apps/desktop/target/release
  // - cross-compiled: target/<triple>/release
  for (const targetDir of [
    cargoTargetDir(),
    path.join(repoRoot, "target"),
    path.join(repoRoot, "apps", "desktop", "src-tauri", "target"),
    path.join(repoRoot, "apps", "desktop", "target"),
  ]) {
    for (const c of findBinaryInTargetDir(targetDir, exe)) {
      maybeAddCandidate(candidates, c);
    }
  }

  // Prefer the most recently built binary if multiple exist (common in CI where both
  // `target/release` and `target/<triple>/release` may exist).
  return pickMostRecentBinary(candidates);
}

async function cargoBuildDesktopBinary() {
  const cargoArgs = [
    "build",
    "-p",
    DESKTOP_TAURI_PACKAGE,
    "--features",
    "desktop",
    "--bin",
    DESKTOP_BINARY_NAME,
    "--release",
  ];

  // Prefer the repo's agent-safe wrapper on macOS/Linux.
  if (process.platform !== "win32") {
    return await run("bash", ["scripts/cargo_agent.sh", ...cargoArgs]);
  }

  // On Windows, we *try* bash first (Git Bash, MSYS2, etc). If it's not available,
  // fall back to `cargo` so the script still works in plain Windows terminals.
  try {
    return await run("bash", ["scripts/cargo_agent.sh", ...cargoArgs], { shell: false });
  } catch (err) {
    if (err && typeof err === "object" && "code" in err && err.code === "ENOENT") {
      console.warn("[coi-check] `bash` not found; falling back to `cargo`.");
      return await run("cargo", cargoArgs);
    }
    throw err;
  }
}

async function main() {
  const argv = process.argv.slice(2);
  let noBuild = isTruthyEnv(process.env.FORMULA_COI_NO_BUILD);
  /** @type {string | null} */
  let explicitBin = null;

  for (let i = 0; i < argv.length; i++) {
    const arg = argv[i];
    if (arg === "--help" || arg === "-h") {
      usage();
      process.exit(0);
    }
    if (arg === "--no-build") {
      noBuild = true;
      continue;
    }
    if (arg === "--bin") {
      const next = argv[i + 1];
      if (!next) {
        console.error("[coi-check] ERROR: --bin requires a path argument.");
        usage();
        process.exit(2);
      }
      explicitBin = next;
      i++;
      continue;
    }
    console.error(`[coi-check] ERROR: unknown argument: ${arg}`);
    usage();
    process.exit(2);
  }

  const frontendDistDir = path.join(desktopDir, "dist");
  const frontendDistIndex = path.join(frontendDistDir, "index.html");

  if (!noBuild) {
    console.log("[coi-check] Building desktop frontend (Vite)...");
    const buildFrontendCode = await run("pnpm", ["build"], { cwd: desktopDir });
    if (buildFrontendCode !== 0) process.exit(buildFrontendCode);
  } else {
    console.log("[coi-check] --no-build enabled; skipping frontend + Rust builds.");
    if (!fs.existsSync(frontendDistIndex)) {
      console.error(`[coi-check] ERROR: expected frontend dist is missing: ${frontendDistIndex}`);
      console.error("[coi-check] Hint: build the desktop frontend with:");
      console.error("  pnpm -C apps/desktop build");
      console.error("[coi-check] Or run the COI check without --no-build to build automatically:");
      console.error("  pnpm -C apps/desktop check:coi");
      process.exit(1);
    }
  }

  if (!noBuild) {
    console.log("[coi-check] Building Tauri desktop binary (release)...");
    const buildDesktopCode = await cargoBuildDesktopBinary();
    if (buildDesktopCode !== 0) process.exit(buildDesktopCode);
  }

  const binary = explicitBin ? resolveBinPath(explicitBin) : noBuild ? detectBuiltDesktopBinary() : desktopBinaryPath();

  if (!binary || !fs.existsSync(binary)) {
    const exe = process.platform === "win32" ? `${DESKTOP_BINARY_NAME}.exe` : DESKTOP_BINARY_NAME;
    const searched = [
      path.join(repoRoot, "target", "release", exe),
      path.join(repoRoot, "apps", "desktop", "src-tauri", "target", "release", exe),
      path.join(repoRoot, "apps", "desktop", "target", "release", exe),
      path.join(repoRoot, "target", "<triple>", "release", exe),
    ];
    console.error("[coi-check] ERROR: could not find a built desktop binary to run.");
    if (explicitBin) {
      console.error(`[coi-check] --bin was provided but does not exist: ${binary}`);
    } else if (noBuild) {
      console.error("[coi-check] Searched common locations such as:");
      for (const p of searched) console.error(`  - ${p}`);
      console.error("[coi-check] Hint: build the app first with one of:");
      console.error("  (cd apps/desktop && bash ../../scripts/cargo_agent.sh tauri build)");
      console.error("  # or");
      console.error(
        "  bash scripts/cargo_agent.sh build -p formula-desktop-tauri --features desktop --bin formula-desktop --release",
      );
      console.error("[coi-check] Or run the COI check without --no-build to build automatically:");
      console.error("  pnpm -C apps/desktop check:coi");
      console.error("[coi-check] You can also pass an explicit binary path via:");
      console.error("  pnpm -C apps/desktop check:coi -- --no-build --bin <path>");
    } else {
      console.error(`[coi-check] Expected desktop binary was not found at: ${desktopBinaryPath()}`);
      console.error(
        "[coi-check] Cargo build succeeded but the output binary is missing. Check the package name and --bin value.",
      );
    }
    process.exit(1);
  }
  console.log(`[coi-check] Running packaged app check: ${binary}`);

  const args = ["--cross-origin-isolation-check"];

  let runCmd = binary;
  let runArgs = args;
  if (process.platform === "linux") {
    // Ensure we have a virtual display in CI/headless environments.
    // xvfb-run-safe.sh is also safe to use on developer machines: if a working
    // DISPLAY is already available it will simply `exec` the command directly.
    runCmd = path.join(repoRoot, "scripts", "xvfb-run-safe.sh");
    runArgs = [binary, ...args];
  }

  const runCode = await run(runCmd, runArgs, {
    // Running the produced `.exe` directly is more reliable than going through a shell
    // (avoids quoting issues when the repo path contains spaces).
    shell: process.platform === "win32" ? false : undefined,
  });
  if (runCode !== 0) {
    console.error(
      `[coi-check] FAILED: packaged Tauri build is missing cross-origin isolation and/or Worker support (exit code ${runCode}).`,
    );
  } else {
    console.log("[coi-check] OK: packaged Tauri build is cross-origin isolated (SharedArrayBuffer + Worker ready).");
  }
  process.exit(runCode);
}

main().catch((err) => {
  console.error("[coi-check] ERROR:", err);
  process.exit(1);
});
