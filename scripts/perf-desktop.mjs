import { spawnSync } from "node:child_process";
import {
  existsSync,
  mkdirSync,
  readdirSync,
  readFileSync,
  rmSync,
  statSync,
} from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

const repoRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");

// Gracefully exit when piping output into a consumer that closes early (e.g. `head`), to avoid a
// noisy EPIPE stack trace.
const onEpipe = (err) => {
  if (err && err.code === "EPIPE") process.exit(0);
};
process.stdout.on("error", onEpipe);
process.stderr.on("error", onEpipe);

function usage() {
  // eslint-disable-next-line no-console
  console.log(`Usage: node scripts/perf-desktop.mjs <startup|memory|size> [-- <args...>]

Commands:
  startup  Run the desktop startup benchmark.
           - default: full app startup (builds apps/desktop/dist + desktop binary)
           - CI default: shell startup only (skips frontend build; requires --startup-bench support in the binary)
  memory   Build the desktop app (dist + binary) and run the desktop idle memory benchmark.
  size     Report desktop size: frontend dist + compressed asset download, binary (+ bloat breakdown), and (if present) installer artifacts.

Environment (shared):
  FORMULA_PERF_HOME             Override the isolated HOME dir (default: target/perf-home)
  FORMULA_PERF_PRESERVE_HOME=1  Skip clearing FORMULA_PERF_HOME before running benchmarks
  FORMULA_PERF_ALLOW_UNSAFE_CLEAN=1
                               Allow clearing FORMULA_PERF_HOME even when it is outside target/
                               (DANGEROUS: could delete user directories if misconfigured)

Environment (startup):
  FORMULA_DESKTOP_STARTUP_RUNS
  FORMULA_DESKTOP_STARTUP_TIMEOUT_MS
  FORMULA_DESKTOP_STARTUP_BENCH_KIND=shell|full
  FORMULA_DESKTOP_STARTUP_MODE=cold|warm
    - cold (default): each measured run uses a fresh profile dir (true cold start)
    - warm: one warmup run primes caches, then measured runs reuse the same profile
  FORMULA_DESKTOP_RSS_IDLE_DELAY_MS
    - delay after startup metrics before sampling idle RSS (startup bench only)
  FORMULA_DESKTOP_RSS_TARGET_MB
    - p95 budget for startup bench idle RSS metric (default: 100MB)
  FORMULA_DESKTOP_WEBVIEW_LOADED_TARGET_MS
    - p95 budget for webview_loaded_ms (native WebView page-load complete; defaults to 800ms)
  FORMULA_DESKTOP_SHELL_WEBVIEW_LOADED_TARGET_MS
    - optional shell-only override (falls back to FORMULA_DESKTOP_WEBVIEW_LOADED_TARGET_MS)

Environment (memory):
  FORMULA_DESKTOP_MEMORY_RUNS
  FORMULA_DESKTOP_MEMORY_TIMEOUT_MS
  FORMULA_DESKTOP_MEMORY_SETTLE_MS
  FORMULA_DESKTOP_IDLE_RSS_TARGET_MB / FORMULA_DESKTOP_MEMORY_TARGET_MB
    - p95 budget for idle RSS (default: 100MB)

Notes:
  - These commands are safe to run locally: they use a repo-local HOME so they don't touch
    ~/.config, ~/Library, etc.
  - Pass extra args after "--" to forward them to the underlying runner script.
      pnpm perf:desktop-startup -- --mode warm
  - To see runner options without building:
      pnpm perf:desktop-startup -- --help
      pnpm perf:desktop-memory -- --help
  - For shell startup benchmarking (no apps/desktop/dist required), pass:
      pnpm perf:desktop-startup -- --startup-bench
    or set:
      FORMULA_DESKTOP_STARTUP_BENCH_KIND=shell
`);
}

function isTruthyEnv(value) {
  if (!value) return false;
  const v = String(value).trim().toLowerCase();
  return v !== "" && v !== "0" && v !== "false" && v !== "no";
}

function isSubpath(parentDir, maybeChild) {
  const rel = path.relative(parentDir, maybeChild);
  if (rel === "") return true;
  if (rel.startsWith("..")) return false;
  // On Windows, `path.relative()` can return an absolute path when drives differ.
  if (path.isAbsolute(rel)) return false;
  return true;
}

function formatPerfPath(p) {
  const abs = path.resolve(path.isAbsolute(p) ? p : path.resolve(repoRoot, p));
  const rel = path.relative(repoRoot, abs);
  if (!rel || rel.startsWith("..") || path.isAbsolute(rel)) return abs;
  return rel;
}

function resolvePerfHome() {
  const fromEnv = process.env.FORMULA_PERF_HOME;
  const trimmed = typeof fromEnv === "string" ? fromEnv.trim() : "";
  if (trimmed !== "") {
    // Always normalize (collapse `..` segments) so safety checks below can't be bypassed by
    // path tricks like `target/perf-home/..` (which would otherwise resolve to `target` at
    // deletion time).
    const candidate = path.isAbsolute(trimmed) ? trimmed : path.resolve(repoRoot, trimmed);
    return path.resolve(candidate);
  }
  return path.resolve(repoRoot, "target", "perf-home");
}

function ensureCleanPerfHome(perfHome) {
  if (isTruthyEnv(process.env.FORMULA_PERF_PRESERVE_HOME)) {
    mkdirSync(perfHome, { recursive: true });
    return;
  }

  // Extra guardrails: never allow `rm -rf /` or `rm -rf <repoRoot>` even if users
  // force-enable unsafe deletion.
  const rootDir = path.parse(perfHome).root;
  if (perfHome === rootDir || perfHome === repoRoot) {
    throw new Error(
      `[perf-desktop] Refusing to reset unsafe perf home dir: ${perfHome}\n` +
        `Pick a path under target/ (recommended) or a dedicated temp dir (e.g. /tmp/formula-perf-home).`,
    );
  }

  const safeRoot = path.resolve(repoRoot, "target");
  const allowUnsafe = isTruthyEnv(process.env.FORMULA_PERF_ALLOW_UNSAFE_CLEAN);

  // Never delete the entire `target/` directory, even with the unsafe override. This directory
  // contains build artifacts and other tooling state; deleting it is almost never intended.
  if (perfHome === safeRoot) {
    if (allowUnsafe) {
      throw new Error(
        `[perf-desktop] Refusing to reset FORMULA_PERF_HOME=${perfHome} because it points at target/ itself.\n` +
          "Pick a subdirectory like target/perf-home (recommended).",
      );
    }
    // eslint-disable-next-line no-console
    console.warn(
      `[perf-desktop] WARN refusing to delete FORMULA_PERF_HOME=${perfHome} because it points at target/ itself.\n` +
        "  - Pick a subdirectory like target/perf-home (recommended)\n",
    );
    mkdirSync(perfHome, { recursive: true });
    return;
  }

  const safeToDelete = perfHome !== safeRoot && isSubpath(safeRoot, perfHome);
  if (!safeToDelete && !allowUnsafe) {
    // eslint-disable-next-line no-console
    console.warn(
      `[perf-desktop] WARN refusing to delete FORMULA_PERF_HOME=${perfHome} because it is outside ${safeRoot}.\n` +
        `  - To suppress this warning and reuse the directory, set FORMULA_PERF_PRESERVE_HOME=1\n` +
        `  - To use a clean dir, pick a path under target/ (recommended)\n` +
        `  - To force deletion anyway, set FORMULA_PERF_ALLOW_UNSAFE_CLEAN=1 (DANGEROUS)\n`,
    );
    mkdirSync(perfHome, { recursive: true });
    return;
  }

  if (!safeToDelete && allowUnsafe) {
    // eslint-disable-next-line no-console
    console.warn(
      `[perf-desktop] WARN FORMULA_PERF_ALLOW_UNSAFE_CLEAN=1: deleting perf home outside target/: ${perfHome}`,
    );
  }

  mkdirSync(path.dirname(perfHome), { recursive: true });
  rmSync(perfHome, { recursive: true, force: true });
  mkdirSync(perfHome, { recursive: true });
}

function perfEnv(extra = {}) {
  const perfHome = resolvePerfHome();
  return {
    perfHome,
    env: {
      ...process.env,
      FORMULA_PERF_HOME: perfHome,
      ...extra,
    },
  };
}

function run(command, args, { cwd = repoRoot, env = process.env } = {}) {
  const proc = spawnSync(command, args, {
    cwd,
    env,
    stdio: "inherit",
    encoding: "utf8",
  });
  if (proc.error) throw proc.error;
  if (proc.status !== 0) process.exit(proc.status ?? 1);
}

function runOptional(command, args, { cwd = repoRoot, env = process.env, label } = {}) {
  const proc = spawnSync(command, args, {
    cwd,
    env,
    stdio: "inherit",
    encoding: "utf8",
  });
  if (proc.error) {
    // eslint-disable-next-line no-console
    console.warn(`[perf-desktop] WARN ${label ?? command}: ${proc.error.message}`);
    return proc.status ?? 1;
  }
  if (proc.status && proc.status !== 0) {
    // eslint-disable-next-line no-console
    console.warn(`[perf-desktop] WARN ${label ?? command} exited with status ${proc.status}`);
  }
  return proc.status ?? 0;
}

function parseDesktopStartupBenchKind({ env, forwardedArgs }) {
  const rawEnv = String(env.FORMULA_DESKTOP_STARTUP_BENCH_KIND ?? "").trim().toLowerCase();
  let envKind = null;
  if (rawEnv !== "") {
    if (rawEnv === "shell") envKind = "shell";
    else if (rawEnv === "full") envKind = "full";
    else throw new Error(`Invalid FORMULA_DESKTOP_STARTUP_BENCH_KIND=${JSON.stringify(rawEnv)}`);
  }

  const defaultKind = envKind ?? (isTruthyEnv(env.CI) ? "shell" : "full");
  let kind = defaultKind;

  // Mirror the startup runner's behavior: later flags win.
  for (const arg of forwardedArgs) {
    if (arg === "--startup-bench" || arg === "--shell") kind = "shell";
    else if (arg === "--full") kind = "full";
  }

  return kind;
}

function detectDesktopPackageName() {
  // The desktop shell package has historically been named both `desktop` and
  // `formula-desktop-tauri`. Prefer reading the current name from the crate's
  // Cargo.toml so local perf tooling stays stable across renames.
  const cargoToml = path.join(repoRoot, "apps", "desktop", "src-tauri", "Cargo.toml");
  try {
    if (!existsSync(cargoToml)) return "formula-desktop-tauri";
    const lines = readFileSync(cargoToml, "utf8").split(/\r?\n/);
    let inPackage = false;
    for (const raw of lines) {
      const line = raw.trim();
      if (!line || line.startsWith("#")) continue;
      if (line.startsWith("[") && line.endsWith("]")) {
        inPackage = line === "[package]";
        continue;
      }
      if (!inPackage) continue;
      if (!line.startsWith("name")) continue;
      const parts = line.split("=", 2);
      if (parts.length !== 2) continue;
      const rhs = parts[1].trim();
      const m = rhs.match(/^"([^"]+)"/);
      if (m?.[1]) return m[1];
    }
  } catch {
    // ignore and fall back
  }
  return "formula-desktop-tauri";
}

function buildDesktop({ env, buildFrontend = true } = {}) {
  if (buildFrontend) {
    // eslint-disable-next-line no-console
    console.log("[perf-desktop] Building frontend (apps/desktop/dist)...");
    run("pnpm", ["-C", "apps/desktop", "build"], { env });
  } else {
    // eslint-disable-next-line no-console
    console.log(
      "[perf-desktop] Skipping frontend build (shell startup benchmark does not require apps/desktop/dist)...",
    );
  }

  // eslint-disable-next-line no-console
  console.log("[perf-desktop] Building desktop binary (target/release/formula-desktop)...");
  const pkg = detectDesktopPackageName();
  // `scripts/cargo_agent.sh` defaults `CARGO_PROFILE_RELEASE_CODEGEN_UNITS` based on its job count
  // for stability on multi-agent hosts. For perf + size measurements we want a binary that matches
  // the repo's Cargo.toml release profile (`codegen-units = 1`), so default it here (callers can
  // still override via the environment).
  const buildEnv = {
    ...env,
    CARGO_PROFILE_RELEASE_CODEGEN_UNITS: env.CARGO_PROFILE_RELEASE_CODEGEN_UNITS || "1",
  };
  run(
    "bash",
    [
      "scripts/cargo_agent.sh",
      "build",
      "-p",
      pkg,
      "--bin",
      "formula-desktop",
      "--release",
      "--features",
      "desktop",
    ],
    { env: buildEnv },
  );
}

function humanBytes(bytes) {
  const units = ["B", "KB", "MB", "GB", "TB"];
  let size = bytes;
  let unit = units[0];
  for (let i = 0; i < units.length - 1 && size >= 1000; i++) {
    size /= 1000;
    unit = units[i + 1];
  }
  if (unit === "B") return `${bytes} ${unit}`;
  return `${size.toFixed(1)} ${unit}`;
}

function bytesToMb(bytes) {
  // Decimal MB, matching existing size budgets in CI.
  return bytes / 1_000_000;
}

function formatMb(value) {
  return `${value.toFixed(3)}MB`;
}

function parseOptionalTargetMb(name) {
  const raw = process.env[name];
  if (raw == null) return null;
  if (String(raw).trim() === "") return null;
  const val = Number(raw);
  if (!Number.isFinite(val) || val <= 0) {
    throw new Error(`Invalid ${name}=${JSON.stringify(raw)} (expected a number > 0)`);
  }
  return val;
}

function dirSizeBytes(dir) {
  let total = 0;
  const entries = readdirSync(dir, { withFileTypes: true });
  for (const ent of entries) {
    const p = path.join(dir, ent.name);
    if (ent.isDirectory()) total += dirSizeBytes(p);
    else if (ent.isFile()) total += statSync(p).st_size;
  }
  return total;
}

function listLargestFiles(dir, limit = 10) {
  /** @type {{path: string, size: number}[]} */
  const files = [];
  const stack = [dir];
  while (stack.length > 0) {
    const cur = stack.pop();
    if (!cur) continue;
    for (const ent of readdirSync(cur, { withFileTypes: true })) {
      const p = path.join(cur, ent.name);
      if (ent.isDirectory()) stack.push(p);
      else if (ent.isFile()) files.push({ path: p, size: statSync(p).st_size });
    }
  }
  files.sort((a, b) => b.size - a.size);
  return files.slice(0, limit);
}

function defaultDesktopBinPath() {
  const exe = process.platform === "win32" ? "formula-desktop.exe" : "formula-desktop";
  const candidates = [
    path.resolve(repoRoot, "target", "release", exe),
    path.resolve(repoRoot, "target", "debug", exe),
    path.resolve(repoRoot, "apps", "desktop", "src-tauri", "target", "release", exe),
    path.resolve(repoRoot, "apps", "desktop", "src-tauri", "target", "debug", exe),
  ];
  for (const p of candidates) {
    if (existsSync(p)) return p;
  }
  return null;
}

function findBundleDirs() {
  /** @type {string[]} */
  const out = [];
  const roots = [
    path.join(repoRoot, "target"),
    path.join(repoRoot, "apps", "desktop", "src-tauri", "target"),
  ];
  for (const root of roots) {
    if (!existsSync(root)) continue;
    const direct = path.join(root, "release", "bundle");
    if (existsSync(direct)) out.push(direct);
    for (const ent of readdirSync(root, { withFileTypes: true })) {
      if (!ent.isDirectory()) continue;
      const candidate = path.join(root, ent.name, "release", "bundle");
      if (existsSync(candidate)) out.push(candidate);
    }
  }
  return [...new Set(out)];
}

function tryRunPython(args, { env } = {}) {
  const candidates = [];
  if (process.env.PYTHON && String(process.env.PYTHON).trim() !== "") {
    candidates.push(String(process.env.PYTHON).trim());
  }
  candidates.push("python3", "python");

  for (const python of candidates) {
    const proc = spawnSync(python, args, {
      cwd: repoRoot,
      env,
      stdio: "inherit",
      encoding: "utf8",
    });
    if (proc.error) {
      // If the interpreter isn't present, try the next candidate.
      if (proc.error.code === "ENOENT") continue;
      throw proc.error;
    }
    return { ran: true, status: proc.status ?? 1 };
  }

  return { ran: false, status: null };
}

function tryCreateDistTarGzBytes(distDir) {
  const outDir = path.join(repoRoot, "target", "perf-artifacts");
  mkdirSync(outDir, { recursive: true });
  const outFile = path.join(outDir, "desktop-dist.tar.gz");
  rmSync(outFile, { force: true });

  const proc = spawnSync("tar", ["-czf", outFile, "-C", distDir, "."], {
    cwd: repoRoot,
    encoding: "utf8",
    stdio: ["ignore", "ignore", "pipe"],
  });
  if (proc.error || proc.status !== 0) {
    rmSync(outFile, { force: true });
    return null;
  }

  try {
    const bytes = statSync(outFile).st_size;
    rmSync(outFile, { force: true });
    return bytes;
  } catch {
    rmSync(outFile, { force: true });
    return null;
  }
}

function readJsonIfExists(p) {
  try {
    if (!existsSync(p)) return null;
    return JSON.parse(readFileSync(p, "utf8"));
  } catch {
    return null;
  }
}

function reportSize({ env }) {
  let failed = false;

  // NOTE: there are two "budget" knobs in the repo:
  // - *_SIZE_LIMIT_MB: used by `scripts/desktop_size_report.py` (PR gating)
  // - *_SIZE_TARGET_MB: used by `pnpm benchmark` size metrics (optional)
  //
  // `pnpm perf:desktop-size` supports both, preferring the *_TARGET_MB variants.
  const binaryTargetMb =
    parseOptionalTargetMb("FORMULA_DESKTOP_BINARY_SIZE_TARGET_MB") ??
    parseOptionalTargetMb("FORMULA_DESKTOP_BINARY_SIZE_LIMIT_MB");
  const distTargetMb =
    parseOptionalTargetMb("FORMULA_DESKTOP_DIST_SIZE_TARGET_MB") ??
    parseOptionalTargetMb("FORMULA_DESKTOP_DIST_SIZE_LIMIT_MB");
  const distGzipTargetMb = parseOptionalTargetMb("FORMULA_DESKTOP_DIST_GZIP_SIZE_TARGET_MB");

  // If python is available, prefer the repo's canonical size report (also supports *_SIZE_LIMIT_MB gating).
  // We still compute/check *_SIZE_TARGET_MB below because the python script doesn't know about those.
  // eslint-disable-next-line no-console
  console.log("\n[desktop-size] Summary (binary + dist):\n");

  const artifactDir = path.join(repoRoot, "target", "perf-artifacts");
  mkdirSync(artifactDir, { recursive: true });
  const jsonOut = path.join(artifactDir, "desktop-size.json");
  rmSync(jsonOut, { force: true });

  const sizeReport = tryRunPython(["scripts/desktop_size_report.py", "--json-out", jsonOut], { env });
  if (!sizeReport.ran) {
    // eslint-disable-next-line no-console
    console.log("[desktop-size] python not found; skipping scripts/desktop_size_report.py");
  } else if (sizeReport.status !== 0) {
    failed = true;
  }

  const sizeJson = readJsonIfExists(jsonOut);

  // Rust binary size breakdown (crates + symbols).
  //
  // This is best-effort: it relies on `cargo-bloat` for the most useful output,
  // but will fall back to `llvm-size`/`size` when available. Keep it non-fatal so
  // `pnpm perf:desktop-size` still works in minimal local environments.
  // eslint-disable-next-line no-console
  console.log("\n[desktop-size] Rust binary breakdown (cargo-bloat / llvm-size):\n");
  const binJsonOut = path.join(artifactDir, "desktop-binary-size.json");
  rmSync(binJsonOut, { force: true });
  const binReport = tryRunPython(
    ["scripts/desktop_binary_size_report.py", "--no-build", "--json-out", binJsonOut],
    { env },
  );
  if (!binReport.ran) {
    // eslint-disable-next-line no-console
    console.log("[desktop-size] python not found; skipping scripts/desktop_binary_size_report.py");
  } else if (binReport.status !== 0) {
    // eslint-disable-next-line no-console
    console.log(`[desktop-size] WARN desktop_binary_size_report exited with status ${binReport.status}`);
    failed = true;
  } else {
    // eslint-disable-next-line no-console
    console.log(`[desktop-size] wrote binary bloat JSON: ${formatPerfPath(binJsonOut)}`);
  }

  const distDir = path.join(repoRoot, "apps", "desktop", "dist");
  const distLabel = formatPerfPath(distDir);
  if (existsSync(distDir)) {
    const total = dirSizeBytes(distDir);
    const totalMb = bytesToMb(total);
    const distStatus = distTargetMb == null || totalMb <= distTargetMb ? "PASS" : "FAIL";
    if (distTargetMb != null && distStatus === "FAIL") failed = true;
    // eslint-disable-next-line no-console
    console.log(
      `\n[desktop-size] dist/ total: ${humanBytes(total)} (${formatMb(totalMb)})  (${distLabel})` +
        (distTargetMb == null ? "" : `  ${distStatus} target=${formatMb(distTargetMb)}`),
    );
    const largest = listLargestFiles(distDir, 10);
    if (largest.length > 0) {
      // eslint-disable-next-line no-console
      console.log("[desktop-size] largest dist assets:");
      for (const f of largest) {
        const rel = formatPerfPath(f.path);
        // eslint-disable-next-line no-console
        console.log(`  - ${humanBytes(f.size).padStart(10)}  ${rel}`);
      }
    }

    let gzMb = null;
    if (typeof sizeJson?.dist_tar_gz?.size_mb === "number" && Number.isFinite(sizeJson.dist_tar_gz.size_mb)) {
      gzMb = sizeJson.dist_tar_gz.size_mb;
      // eslint-disable-next-line no-console
      console.log(`[desktop-size] dist.tar.gz: ${formatMb(gzMb)} (from desktop_size_report.py)`);
    } else {
      const gzBytes = tryCreateDistTarGzBytes(distDir);
      if (gzBytes == null) {
        // eslint-disable-next-line no-console
        console.log("[desktop-size] dist.tar.gz: unavailable (tar failed/missing)");
      } else {
        gzMb = bytesToMb(gzBytes);
        // eslint-disable-next-line no-console
        console.log(`[desktop-size] dist.tar.gz: ${humanBytes(gzBytes)} (${formatMb(gzMb)})`);
      }
    }

    if (distGzipTargetMb != null && gzMb != null) {
      const gzStatus = gzMb <= distGzipTargetMb ? "PASS" : "FAIL";
      if (gzStatus === "FAIL") failed = true;
      // eslint-disable-next-line no-console
      console.log(
        `[desktop-size] dist.tar.gz budget: ${gzStatus} value=${formatMb(gzMb)} target=${formatMb(distGzipTargetMb)}`,
      );
    }
  } else {
    // eslint-disable-next-line no-console
    console.log(`\n[desktop-size] dist/ not found at ${distLabel} (run: pnpm -C apps/desktop build)`);
  }

  const binPath = process.env.FORMULA_DESKTOP_BIN
    ? path.resolve(repoRoot, process.env.FORMULA_DESKTOP_BIN)
    : defaultDesktopBinPath();
  if (binPath && existsSync(binPath)) {
    const size = statSync(binPath).st_size;
    const sizeMb = bytesToMb(size);
    const binStatus = binaryTargetMb == null || sizeMb <= binaryTargetMb ? "PASS" : "FAIL";
    if (binaryTargetMb != null && binStatus === "FAIL") failed = true;
      // eslint-disable-next-line no-console
      console.log(
        `\n[desktop-size] binary: ${humanBytes(size)} (${formatMb(sizeMb)})  (${formatPerfPath(binPath)})` +
          (binaryTargetMb == null ? "" : `  ${binStatus} target=${formatMb(binaryTargetMb)}`),
      );
  } else {
    // eslint-disable-next-line no-console
    console.log(
      `\n[desktop-size] desktop binary not found (expected target/(release|debug)/formula-desktop or apps/desktop/src-tauri/target/(release|debug)/formula-desktop).\n` +
        `Build it via: bash scripts/cargo_agent.sh build -p formula-desktop-tauri --bin formula-desktop --release --features desktop\n` +
        `Or set FORMULA_DESKTOP_BIN=/path/to/formula-desktop`,
    );
  }

  // Approximate the *network download cost* of the frontend by summing per-asset Brotli/gzip sizes.
  // This is intentionally separate from installer artifact size budgets.
  // eslint-disable-next-line no-console
  console.log("\n[desktop-size] frontend asset download size (compressed JS/CSS/WASM):\n");
  const assetStatus = runOptional(process.execPath, ["scripts/frontend_asset_size_report.mjs", "--dist", "apps/desktop/dist"], {
    env,
    label: "frontend_asset_size_report",
  });
  if (assetStatus !== 0) failed = true;

  const bundleDirs = findBundleDirs();
  if (bundleDirs.length === 0) {
    // eslint-disable-next-line no-console
    console.log(
      `\n[desktop-size] No installer artifacts found (expected <target>/release/bundle or <target>/<triple>/release/bundle).\n` +
        `To generate installers/bundles, run:\n` +
        `  (cd apps/desktop && bash ../../scripts/cargo_agent.sh tauri build)\n` +
        `Then re-run: pnpm perf:desktop-size\n`,
    );
    if (failed) process.exitCode = 1;
    return;
  }

  // eslint-disable-next-line no-console
  console.log(`\n[desktop-size] Installer artifacts (override limit via FORMULA_BUNDLE_SIZE_LIMIT_MB):\n`);

  const bundleReport = tryRunPython(["scripts/desktop_bundle_size_report.py"], { env });
  if (!bundleReport.ran) {
    // eslint-disable-next-line no-console
    console.log("[desktop-size] python not found; skipping bundle size report");
  } else if (bundleReport.status !== 0) {
    failed = true;
  }

  if (failed) process.exitCode = 1;
}

function main() {
  const cmd = process.argv[2];
  const passthroughIdx = process.argv.indexOf("--");
  const forwardedArgs =
    passthroughIdx >= 0 ? process.argv.slice(passthroughIdx + 1) : process.argv.slice(3);
  const forwardedWantsHelp = forwardedArgs.includes("-h") || forwardedArgs.includes("--help");

  if (!cmd || cmd === "-h" || cmd === "--help") {
    usage();
    process.exit(cmd ? 0 : 2);
  }

  // Convenience: allow `pnpm perf:desktop-startup --help` to show the underlying runner help
  // without building the app first.
  if (cmd === "startup" && forwardedWantsHelp) {
    run(process.execPath, [
      "scripts/run-node-ts.mjs",
      "apps/desktop/tests/performance/desktop-startup-runner.ts",
      ...forwardedArgs,
    ]);
    return;
  }
  if (cmd === "memory" && forwardedWantsHelp) {
    run(process.execPath, [
      "scripts/run-node-ts.mjs",
      "apps/desktop/tests/performance/desktop-memory-runner.ts",
      ...forwardedArgs,
    ]);
    return;
  }
  if (cmd === "size" && forwardedWantsHelp) {
    usage();
    return;
  }

  const { perfHome, env } = perfEnv();
  ensureCleanPerfHome(perfHome);

  // eslint-disable-next-line no-console
  console.log(`[perf-desktop] Using isolated desktop HOME root=${formatPerfPath(perfHome)}`);
  // eslint-disable-next-line no-console
  console.log(
    "[perf-desktop] Tip: set FORMULA_PERF_PRESERVE_HOME=1 to avoid clearing the perf HOME between runs.\n" +
      "             set FORMULA_PERF_HOME=/path/to/dir to override.\n",
  );

  if (cmd === "startup") {
    const benchKind = parseDesktopStartupBenchKind({ env, forwardedArgs });
    buildDesktop({ env, buildFrontend: benchKind === "full" });

    // eslint-disable-next-line no-console
    console.log(`\n[perf-desktop] Running desktop startup benchmark (kind=${benchKind})...\n`);
    run(
      process.execPath,
      ["scripts/run-node-ts.mjs", "apps/desktop/tests/performance/desktop-startup-runner.ts", ...forwardedArgs],
      {
        env: { ...env, FORMULA_RUN_DESKTOP_STARTUP_BENCH: "1" },
      },
    );
    return;
  }

  if (cmd === "memory") {
    buildDesktop({ env });

    // eslint-disable-next-line no-console
    console.log("\n[perf-desktop] Running desktop idle memory benchmark...\n");
    run(
      process.execPath,
      ["scripts/run-node-ts.mjs", "apps/desktop/tests/performance/desktop-memory-runner.ts", ...forwardedArgs],
      {
        env: { ...env, FORMULA_RUN_DESKTOP_MEMORY_BENCH: "1" },
      },
    );
    return;
  }

  if (cmd === "size") {
    buildDesktop({ env });
    reportSize({ env });
    return;
  }

  // eslint-disable-next-line no-console
  console.error(`Unknown command: ${cmd}`);
  usage();
  process.exit(2);
}

main();
