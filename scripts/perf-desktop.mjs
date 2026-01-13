import { spawnSync } from "node:child_process";
import { existsSync, mkdirSync, readdirSync, rmSync, statSync } from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

const repoRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");

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

Environment (startup):
  FORMULA_DESKTOP_STARTUP_MODE=cold|warm
    - cold (default): each measured run uses a fresh profile dir (true cold start)
    - warm: one warmup run primes caches, then measured runs reuse the same profile
  FORMULA_DESKTOP_WEBVIEW_LOADED_TARGET_MS
    - p95 budget for `webview_loaded_ms` (native WebView page-load complete; defaults to 800ms)
  FORMULA_DESKTOP_SHELL_WEBVIEW_LOADED_TARGET_MS
    - optional shell-only override (falls back to FORMULA_DESKTOP_WEBVIEW_LOADED_TARGET_MS)

Notes:
  - These commands are safe to run locally: they use a repo-local HOME so they don't touch
    ~/.config, ~/Library, etc.
  - Pass extra args after "--" to forward them to the underlying runner script.
      pnpm perf:desktop-startup -- --mode warm
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

function resolvePerfHome() {
  const fromEnv = process.env.FORMULA_PERF_HOME;
  if (fromEnv && fromEnv.trim() !== "") {
    return path.isAbsolute(fromEnv) ? fromEnv : path.resolve(repoRoot, fromEnv);
  }
  return path.resolve(repoRoot, "target", "perf-home");
}

function ensureCleanPerfHome(perfHome) {
  mkdirSync(path.dirname(perfHome), { recursive: true });
  if (isTruthyEnv(process.env.FORMULA_PERF_PRESERVE_HOME)) {
    mkdirSync(perfHome, { recursive: true });
    return;
  }
  rmSync(perfHome, { recursive: true, force: true });
  mkdirSync(perfHome, { recursive: true });
}

function perfEnv(extra = {}) {
  const perfHome = resolvePerfHome();
  // Keep XDG state alongside HOME so libraries don't write to the real user profile.
  const xdgBase = perfHome;
  return {
    perfHome,
    env: {
      ...process.env,
      FORMULA_PERF_HOME: perfHome,
      HOME: perfHome,
      USERPROFILE: perfHome,
      XDG_CACHE_HOME: path.join(xdgBase, ".cache"),
      XDG_CONFIG_HOME: path.join(xdgBase, ".config"),
      XDG_STATE_HOME: path.join(xdgBase, ".local", "state"),
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
  run("bash", ["scripts/cargo_agent.sh", "build", "-p", "formula-desktop-tauri", "--bin", "formula-desktop", "--release", "--features", "desktop"], {
    env,
  });
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
  const roots = [path.join(repoRoot, "target"), path.join(repoRoot, "apps", "desktop", "src-tauri", "target")];
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

function runPython(script, args, { env } = {}) {
  // Prefer python3, fall back to python.
  const python = process.env.PYTHON || "python3";
  const proc = spawnSync(python, [script, ...args], {
    cwd: repoRoot,
    env,
    stdio: "inherit",
    encoding: "utf8",
  });
  if (proc.status === 0) return;

  // Retry with `python` if python3 isn't available.
  if (python === "python3" && proc.error) {
    const retry = spawnSync("python", [script, ...args], {
      cwd: repoRoot,
      env,
      stdio: "inherit",
      encoding: "utf8",
    });
    if (retry.error) throw retry.error;
    if (retry.status !== 0) process.exit(retry.status ?? 1);
    return;
  }

  if (proc.error) throw proc.error;
  process.exit(proc.status ?? 1);
}

function reportSize({ env }) {
  // eslint-disable-next-line no-console
  console.log("\n[desktop-size] Summary (binary + dist):\n");
  runPython("scripts/desktop_size_report.py", [], { env });

  // Rust binary size breakdown (crates + symbols).
  //
  // This is best-effort: it relies on `cargo-bloat` for the most useful output,
  // but will fall back to `llvm-size`/`size` when available. Keep it non-fatal so
  // `pnpm perf:desktop-size` still works in minimal local environments.
  // eslint-disable-next-line no-console
  console.log("\n[desktop-size] Rust binary breakdown (cargo-bloat / llvm-size):\n");
  runOptional(process.env.PYTHON || "python3", ["scripts/desktop_binary_size_report.py", "--no-build"], {
    env,
    label: "desktop_binary_size_report",
  });

  const distDir = path.join(repoRoot, "apps", "desktop", "dist");
  if (existsSync(distDir)) {
    const largest = listLargestFiles(distDir, 10);
    if (largest.length > 0) {
      // eslint-disable-next-line no-console
      console.log("[desktop-size] largest dist assets:");
      for (const f of largest) {
        const rel = path.relative(repoRoot, f.path);
        // eslint-disable-next-line no-console
        console.log(`  - ${humanBytes(f.size).padStart(10)}  ${rel}`);
      }
    }
  }

  // Approximate the *network download cost* of the frontend by summing per-asset Brotli/gzip sizes.
  // This is intentionally separate from installer artifact size budgets.
  // eslint-disable-next-line no-console
  console.log("\n[desktop-size] frontend asset download size (compressed JS/CSS/WASM):\n");
  runOptional(process.execPath, ["scripts/frontend_asset_size_report.mjs", "--dist", "apps/desktop/dist"], {
    env,
    label: "frontend_asset_size_report",
  });

  const bundleDirs = findBundleDirs();
  if (bundleDirs.length === 0) {
    // eslint-disable-next-line no-console
    console.log(
      `\n[desktop-size] No Tauri bundle artifacts found (target/**/release/bundle).\n` +
        `To generate installers/bundles, run:\n` +
        `  (cd apps/desktop && bash ../../scripts/cargo_agent.sh tauri build)\n` +
        `Then re-run: pnpm perf:desktop-size\n`,
    );
    return;
  }

  // eslint-disable-next-line no-console
  console.log(`\n[desktop-size] Installer artifacts (override limit via FORMULA_BUNDLE_SIZE_LIMIT_MB):\n`);

  runPython("scripts/desktop_bundle_size_report.py", [], { env });
}

function main() {
  const cmd = process.argv[2];
  const passthroughIdx = process.argv.indexOf("--");
  const forwardedArgs = passthroughIdx >= 0 ? process.argv.slice(passthroughIdx + 1) : process.argv.slice(3);

  if (!cmd || cmd === "-h" || cmd === "--help") {
    usage();
    process.exit(cmd ? 0 : 2);
  }

  const { perfHome, env } = perfEnv();
  ensureCleanPerfHome(perfHome);

  // eslint-disable-next-line no-console
  console.log(`[perf-desktop] Using isolated HOME=${path.relative(repoRoot, perfHome)}`);
  // eslint-disable-next-line no-console
  console.log(
    "[perf-desktop] Tip: set FORMULA_PERF_PRESERVE_HOME=1 to reuse caches between runs.\n" +
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
