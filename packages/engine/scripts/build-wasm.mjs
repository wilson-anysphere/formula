import { spawnSync } from "node:child_process";
import { existsSync } from "node:fs";
import { copyFile, mkdir, readFile, readdir, rename, rm, stat } from "node:fs/promises";
import os from "node:os";
import path from "node:path";
import { fileURLToPath } from "node:url";

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

// `packages/engine/scripts/*` → repo root
const repoRoot = path.resolve(__dirname, "..", "..", "..");

const defaultGlobalCargoHome = path.resolve(os.homedir(), ".cargo");
const envCargoHome = process.env.CARGO_HOME;
const normalizedEnvCargoHome = envCargoHome ? path.resolve(envCargoHome) : null;
const cargoHome =
  !envCargoHome ||
  (!process.env.CI &&
    !process.env.FORMULA_ALLOW_GLOBAL_CARGO_HOME &&
    normalizedEnvCargoHome === defaultGlobalCargoHome)
    ? path.join(repoRoot, "target", "cargo-home")
    : envCargoHome;
await mkdir(cargoHome, { recursive: true });
const childEnv = { ...process.env, CARGO_HOME: cargoHome };
// `RUSTUP_TOOLCHAIN` overrides the repo's `rust-toolchain.toml` pin. Some environments set it
// globally (often to `stable`), which would bypass the pinned toolchain and reintroduce drift when
// running cargo (via wasm-pack) from this script.
if (childEnv.RUSTUP_TOOLCHAIN && existsSync(path.join(repoRoot, "rust-toolchain.toml"))) {
  delete childEnv.RUSTUP_TOOLCHAIN;
}
const cargoBinDir = path.join(cargoHome, "bin");
await mkdir(cargoBinDir, { recursive: true });
if (!childEnv.PATH?.split(path.delimiter).includes(cargoBinDir)) {
  childEnv.PATH = childEnv.PATH ? `${cargoBinDir}${path.delimiter}${childEnv.PATH}` : cargoBinDir;
}

const crateDir = path.join(repoRoot, "crates", "formula-wasm");

const outDir = path.join(repoRoot, "packages", "engine", "pkg");
// Build into a separate directory and swap it into place on success. This prevents an
// interrupted build (Ctrl+C, OOM, CI cancellation, etc) from deleting the last-known-good
// WASM bundle and forcing every subsequent workflow (dev server / e2e) to do a full rebuild.
const buildOutDir = `${outDir}-build`;
const backupOutDir = `${outDir}-backup`;
// Note: `wasm-pack build --out-dir` is documented as a *relative* path and is
// resolved from the crate directory, not `cwd`. Use an absolute path to ensure
// output always lands in this repo's deterministic location.
const wrapper = path.join(outDir, "formula_wasm.js");
const wasm = path.join(outDir, "formula_wasm_bg.wasm");
const buildWrapper = path.join(buildOutDir, "formula_wasm.js");
const buildWasm = path.join(buildOutDir, "formula_wasm_bg.wasm");

// If the previous run was interrupted during the final swap step, we may have a complete
// `pkg-backup/` directory but no `pkg/` directory. Restore it so subsequent runs can still
// short-circuit when nothing has changed.
if (!existsSync(outDir) && existsSync(backupOutDir)) {
  try {
    await rename(backupOutDir, outDir);
  } catch {
    // ignore
  }
}

const targets = [
  path.join(repoRoot, "apps", "web", "public", "engine"),
  path.join(repoRoot, "apps", "desktop", "public", "engine")
];

const wasmPackBin = process.platform === "win32" ? "wasm-pack.exe" : "wasm-pack";

const releaseWorkflowPath = path.join(repoRoot, ".github", "workflows", "release.yml");

async function readPinnedWasmPackVersion() {
  try {
    const text = await readFile(releaseWorkflowPath, "utf8");
    const match = text.match(/^[\t ]*WASM_PACK_VERSION:[\t ]*["']?([^"'\n]+)["']?/m);
    return match ? match[1].trim() : null;
  } catch {
    return null;
  }
}

const scriptArgs = process.argv.slice(2);
const forceBuild =
  scriptArgs.includes("--force") ||
  scriptArgs.includes("-f") ||
  process.env.FORMULA_WASM_FORCE_BUILD === "1";
const showHelp = scriptArgs.includes("--help") || scriptArgs.includes("-h");
if (showHelp) {
  console.log(
    [
      "Build the Rust/WASM engine (used by apps/*).",
      "",
      "Usage:",
      "  pnpm -C packages/engine build:wasm [-- --force]",
      "",
      "Options:",
      "  --force, -f    Rebuild even if artifacts appear up to date.",
      "",
      "Environment:",
      "  FORMULA_WASM_FORCE_BUILD=1   Same as --force.",
    ].join("\n"),
  );
  process.exit(0);
}

function cargoAgentPath() {
  const abs = path.join(repoRoot, "scripts", "cargo_agent.sh");
  const rel = path.relative(process.cwd(), abs);
  return rel || abs;
}

function fatal(message) {
  console.error(message);
  process.exit(1);
}

async function latestMtime(entryPath) {
  const info = await stat(entryPath);
  if (!info.isDirectory()) {
    return info.mtimeMs;
  }

  // Ignore generated/build output directories when determining whether the Rust
  // sources have changed. These can be updated by unrelated workflows (e.g.
  // `wasm-pack --target nodejs` populating `pkg-node/`) and would otherwise force
  // unnecessary rebuilds of the web WASM bundle.
  const base = path.basename(entryPath);
  if (base === "target" || base === "pkg" || base === "pkg-node" || base === "node_modules") {
    return 0;
  }

  // Directory mtimes update when entries are added/removed, and also when ignored
  // build artifacts change. Use the maximum of child entry mtimes instead so
  // generated outputs like `pkg-node/` don't force rebuilds of the web bundle.
  let latest = 0;
  const entries = await readdir(entryPath, { withFileTypes: true });
  for (const entry of entries) {
    const childPath = path.join(entryPath, entry.name);
    const mtime = await latestMtime(childPath);
    latest = Math.max(latest, mtime);
  }

  return latest;
}

function isDependencyKeyValueTable(tableName) {
  if (tableName === "dependencies" || tableName === "build-dependencies") return true;
  if (!tableName.startsWith("target.")) return false;
  return tableName.endsWith(".dependencies") || tableName.endsWith(".build-dependencies");
}

function isDependencyDetailTable(tableName) {
  if (tableName.startsWith("dependencies.") || tableName.startsWith("build-dependencies.")) return true;
  if (!tableName.startsWith("target.")) return false;
  return tableName.includes(".dependencies.") || tableName.includes(".build-dependencies.");
}

function extractPathDependencies(cargoToml) {
  const deps = new Set();
  let currentTable = null;

  for (const rawLine of cargoToml.split(/\r?\n/)) {
    const withoutComment = rawLine.split("#")[0];
    const line = withoutComment.trim();
    if (!line) continue;

    const headerMatch = line.match(/^\[([^\]]+)\]\s*$/);
    if (headerMatch) {
      currentTable = headerMatch[1].trim();
      continue;
    }

    if (currentTable && isDependencyKeyValueTable(currentTable)) {
      const inlineTableMatch = line.match(/^[A-Za-z0-9_-]+\s*=\s*\{([^}]*)\}\s*$/);
      if (!inlineTableMatch) continue;

      const body = inlineTableMatch[1];
      const pathMatch = body.match(/\bpath\s*=\s*(?:"([^"]+)"|'([^']+)')/);
      const depPath = pathMatch?.[1] ?? pathMatch?.[2];
      if (depPath) deps.add(depPath);
      continue;
    }

    if (currentTable && isDependencyDetailTable(currentTable)) {
      const pathMatch = line.match(/^path\s*=\s*(?:"([^"]+)"|'([^']+)')\s*$/);
      const depPath = pathMatch?.[1] ?? pathMatch?.[2];
      if (depPath) deps.add(depPath);
    }
  }

  return Array.from(deps);
}

async function collectDependencyCrates(entryDirs) {
  const visited = new Set();
  const queue = [...entryDirs];

  while (queue.length > 0) {
    const cratePath = queue.pop();
    if (!cratePath) continue;

    const crateDirPath = path.resolve(cratePath);
    if (visited.has(crateDirPath)) continue;

    const manifestPath = path.join(crateDirPath, "Cargo.toml");
    if (!existsSync(manifestPath)) continue;

    visited.add(crateDirPath);

    const cargoToml = await readFile(manifestPath, "utf8");
    const depPaths = extractPathDependencies(cargoToml);
    for (const depPath of depPaths) {
      const resolved = path.resolve(crateDirPath, depPath);
      if (existsSync(path.join(resolved, "Cargo.toml"))) {
        queue.push(resolved);
      }
    }
  }

  return visited;
}

async function copyRuntimeAssets(sourceDir, destDir) {
  await mkdir(destDir, { recursive: true });
  const entries = await readdir(sourceDir, { withFileTypes: true });

  for (const entry of entries) {
    if (entry.name === ".gitignore") continue;
    if (entry.name === "package.json") continue;
    if (entry.name.endsWith(".d.ts")) continue;

    const from = path.join(sourceDir, entry.name);
    const to = path.join(destDir, entry.name);

    if (entry.isDirectory()) {
      await copyRuntimeAssets(from, to);
      continue;
    }

    if (!entry.isFile()) continue;
    await copyFile(from, to);
  }
}

async function copyToPublic() {
  for (const targetDir of targets) {
    await copyRuntimeAssets(outDir, targetDir);
  }
}

// Ensure the crate path exists (helps when running from unexpected working dirs).
if (!existsSync(path.join(crateDir, "Cargo.toml"))) {
  fatal(
    `[formula] Expected WASM crate at ${path.relative(repoRoot, crateDir)} (relative to repo root), but it was not found.`
  );
}

const outputExists = existsSync(wrapper) && existsSync(wasm);
if (outputExists) {
  const outputStamp = Math.min((await stat(wrapper)).mtimeMs, (await stat(wasm)).mtimeMs);
  const dependencyCrates = await collectDependencyCrates([crateDir]);

  let sourceStamp = await latestMtime(path.join(repoRoot, "Cargo.lock"));
  for (const dependencyCrate of dependencyCrates) {
    sourceStamp = Math.max(sourceStamp, await latestMtime(dependencyCrate));
  }

  if (outputStamp >= sourceStamp && !forceBuild) {
    console.log("[formula] WASM artifacts up to date; copying runtime assets into apps/*/public/engine.");
    await copyToPublic();
    process.exit(0);
  }
  if (outputStamp >= sourceStamp && forceBuild) {
    console.log("[formula] --force specified; rebuilding WASM artifacts.");
  }
}

// Validate `wasm-pack` is installed (only required when we need to rebuild).
{
  const check = spawnSync(wasmPackBin, ["--version"], { encoding: "utf8", env: childEnv });
  if (check.error) {
    const pinnedWasmPack = await readPinnedWasmPackVersion();
    const pinnedHint = pinnedWasmPack
      ? `  - bash ${cargoAgentPath()} install wasm-pack --version "${pinnedWasmPack}" --locked --force`
      : `  - bash ${cargoAgentPath()} install wasm-pack`;
    fatal(
      [
        "[formula] wasm-pack is required to build the Rust/WASM engine but was not found on PATH.",
        "",
        "Install it with one of:",
        pinnedHint,
        "  - https://rustwasm.github.io/wasm-pack/installer/",
        "",
        `Original error: ${check.error?.message ?? "unknown"}`
      ].join("\n")
    );
  }

  if (check.status !== 0) {
    const stderr = (check.stderr ?? "").trim();
    const stdout = (check.stdout ?? "").trim();
    fatal(
      [
        "[formula] wasm-pack is installed but failed to run.",
        "",
        `Exit status: ${check.status ?? "unknown"}${check.signal ? ` (signal ${check.signal})` : ""}`,
        stdout ? `stdout:\n${stdout}` : null,
        stderr ? `stderr:\n${stderr}` : null
      ]
        .filter(Boolean)
        .join("\n")
    );
  }
}

console.log("[formula] Building WASM artifacts via wasm-pack…");

// Some environments configure Cargo to use `sccache` via `build.rustc-wrapper` or
// other wrapper settings. When the wrapper is unavailable/misconfigured, wasm-pack
// builds can fail even for `cargo metadata`/`rustc -vV`. Default to disabling any
// configured wrapper unless the user explicitly overrides it in the environment.
const rustcWrapper = process.env.RUSTC_WRAPPER ?? process.env.CARGO_BUILD_RUSTC_WRAPPER ?? "";
const rustcWorkspaceWrapper =
  process.env.RUSTC_WORKSPACE_WRAPPER ??
  process.env.CARGO_BUILD_RUSTC_WORKSPACE_WRAPPER ??
  "";
const wasmPackEnv = {
  ...childEnv,
  RUSTC_WRAPPER: rustcWrapper,
  RUSTC_WORKSPACE_WRAPPER: rustcWorkspaceWrapper,
  // Cargo config can also be controlled via `CARGO_BUILD_RUSTC_WRAPPER` env vars; set these so
  // we reliably override any global config (e.g. `build.rustc-wrapper=sccache`) when callers
  // didn't explicitly opt in.
  CARGO_BUILD_RUSTC_WRAPPER: rustcWrapper,
  CARGO_BUILD_RUSTC_WORKSPACE_WRAPPER: rustcWorkspaceWrapper,
};

function filterRustflagsForWasm(raw) {
  if (typeof raw !== "string") return raw;
  const trimmed = raw.trim();
  if (!trimmed) return raw;
  const tokens = trimmed.split(/\s+/).filter(Boolean);
  const out = [];
  for (let i = 0; i < tokens.length; i += 1) {
    const token = tokens[i];
    const next = i + 1 < tokens.length ? tokens[i + 1] : null;

    // Some build wrappers (e.g. `scripts/cargo_agent.sh`) append host-only linker flags like
    // `-C link-arg=-Wl,--threads=1` to reduce lld thread usage. Those flags are not understood by
    // `rust-lld -flavor wasm`, causing wasm builds to fail.
    if (token === "-C" && next && next.startsWith("link-arg=-Wl,--threads=")) {
      i += 1;
      continue;
    }
    if (token.startsWith("link-arg=-Wl,--threads=")) {
      // Defensive: handle cases where the `-C` token was stripped elsewhere.
      if (out[out.length - 1] === "-C") out.pop();
      continue;
    }
    if (token.startsWith("-Clink-arg=-Wl,--threads=")) {
      continue;
    }

    out.push(token);
  }
  return out.join(" ");
}

// Ensure target-specific wasm builds do not inherit host-only linker flags from wrapper scripts.
const inheritedRustflags = process.env.RUSTFLAGS;
const filteredRustflags = filterRustflagsForWasm(inheritedRustflags);
if (filteredRustflags !== inheritedRustflags) {
  if (typeof filteredRustflags === "string" && filteredRustflags.trim()) {
    wasmPackEnv.RUSTFLAGS = filteredRustflags;
  } else {
    delete wasmPackEnv.RUSTFLAGS;
  }
}

function isPositiveIntegerString(value) {
  return typeof value === "string" && /^[1-9]\d*$/.test(value.trim());
}

function defaultWasmConcurrency() {
  const rawFormulaJobs = process.env.FORMULA_CARGO_JOBS;
  const rawCargoJobs = process.env.CARGO_BUILD_JOBS;
  const jobsFromEnv =
    (isPositiveIntegerString(rawFormulaJobs) ? rawFormulaJobs.trim() : null) ??
    (isPositiveIntegerString(rawCargoJobs) ? rawCargoJobs.trim() : null);
  // Default to conservative parallelism. On high-core-count machines the Rust linker
  // (lld) can spawn many threads per invocation; combining that with Cargo-level
  // parallelism can exceed sandbox process/thread limits and cause flaky "Resource
  // temporarily unavailable" failures.
  const cpuCount = os.cpus().length;
  const defaultJobs = cpuCount >= 64 ? "2" : "4";
  const jobs = jobsFromEnv ?? defaultJobs;
  return {
    jobs,
    makeflags: process.env.MAKEFLAGS ?? `-j${jobs}`,
    // Prefer setting codegen units via Cargo profile env vars rather than forcing `RUSTFLAGS`.
    // This avoids overriding other profile configuration and matches the approach used by
    // `scripts/cargo_agent.sh` on multi-agent hosts.
    releaseCodegenUnits: process.env.CARGO_PROFILE_RELEASE_CODEGEN_UNITS?.trim() || jobs,
    rayonThreads:
      process.env.RAYON_NUM_THREADS ??
      process.env.FORMULA_RAYON_NUM_THREADS ??
      // When jobs are explicitly configured, prefer that.
      jobsFromEnv ??
      jobs,
    // wasm-pack's release mode runs wasm-opt (Binaryen) which can spawn one worker per CPU core.
    // Mirror our Cargo concurrency defaults so constrained environments don't fail with
    // `Resource temporarily unavailable` due to thread/process limits.
    binaryenCores:
      process.env.BINARYEN_CORES ??
      process.env.FORMULA_BINARYEN_CORES ??
      // When jobs are explicitly configured, prefer that.
      jobsFromEnv ??
      jobs,
  };
}

function runWasmPack({ jobs, makeflags, releaseCodegenUnits, rayonThreads, binaryenCores }) {
  const limitAs = process.env.FORMULA_CARGO_LIMIT_AS ?? "14G";
  const runLimited = path.join(repoRoot, "scripts", "run_limited.sh");
  const canUseRunLimited = process.platform !== "win32" && existsSync(runLimited);
  const verbose =
    process.env.FORMULA_WASM_PACK_VERBOSE === "1" || process.env.FORMULA_WASM_PACK_VERBOSE === "true";
  // `wasm-pack build` inherits cargo's per-crate compile output, which can be extremely verbose.
  // In CI/agent environments where stdout isn't a TTY this can create enormous logs and even hit
  // output capture limits. Pass `--quiet` through to cargo in those cases unless callers
  // explicitly opt into verbose output.
  const cargoExtraArgs = ["--locked"];
  if (!verbose && !process.stdout.isTTY) cargoExtraArgs.push("--quiet");

  const wasmPackArgs = [
    "build",
    crateDir,
    "--target",
    "web",
    "--release",
    // Binaryen's wasm-opt validator can lag behind newly-emitted wasm features
    // (e.g. bulk memory ops, non-trapping float-to-int). Skip wasm-opt so engine
    // builds remain stable across environments and e2e workflows.
    "--no-opt",
    "--out-dir",
    buildOutDir,
    "--out-name",
    "formula_wasm",
    // Avoid generating a nested package.json in the output directory; consumers
    // import the wrapper by URL and do not need `wasm-pack`'s npm packaging.
    "--no-pack",
    ...cargoExtraArgs,
  ];

  const env = {
    ...wasmPackEnv,
    // Keep builds safe in high-core-count environments (e.g. agent sandboxes) even
    // if the caller didn't initialize via `scripts/agent-init.sh`.
    CARGO_BUILD_JOBS: jobs,
    MAKEFLAGS: makeflags,
    CARGO_PROFILE_RELEASE_CODEGEN_UNITS: releaseCodegenUnits,
    // Rayon defaults to spawning one thread per core. On multi-agent hosts this can be very
    // large and can even fail to initialize ("Resource temporarily unavailable"). Default it
    // to our safe cargo job count unless explicitly overridden by the caller.
    RAYON_NUM_THREADS: rayonThreads,
    BINARYEN_CORES: binaryenCores,
  };

  return spawnSync(
    ...(canUseRunLimited
      ? ["bash", [runLimited, "--as", limitAs, "--", wasmPackBin, ...wasmPackArgs]]
      : [wasmPackBin, wasmPackArgs]),
    {
      cwd: repoRoot,
      stdio: "inherit",
      env,
    }
  );
}

// `wasm-pack` refuses to overwrite some files if the output already exists. Build into a
// dedicated directory and swap it into place once the build completes successfully.
await rm(buildOutDir, { recursive: true, force: true });
await rm(backupOutDir, { recursive: true, force: true });

const concurrency = defaultWasmConcurrency();
let result = runWasmPack({
  jobs: concurrency.jobs,
  makeflags: concurrency.makeflags,
  releaseCodegenUnits: concurrency.releaseCodegenUnits,
  rayonThreads: concurrency.rayonThreads,
  binaryenCores: concurrency.binaryenCores,
});

// When running on heavily loaded/locked-down hosts, even modest parallelism can fail with
// `Resource temporarily unavailable` (thread or process spawn failures). If the caller didn't
// explicitly opt into custom concurrency settings, retry once with the most conservative config.
if ((result.status ?? 0) !== 0) {
  const rustflagsSetsCodegenUnits =
    typeof process.env.RUSTFLAGS === "string" && process.env.RUSTFLAGS.includes("codegen-units");
  const userProvidedConcurrency =
    process.env.FORMULA_CARGO_JOBS ||
    process.env.CARGO_BUILD_JOBS ||
    process.env.MAKEFLAGS ||
    // Treat RUSTFLAGS as a concurrency override only when it sets `-C codegen-units=...`.
    // Many environments (including CI) set RUSTFLAGS for unrelated reasons (e.g. `-D warnings`);
    // we still want to retry with a conservative `-j1` config when the failure is due to thread/
    // process limits.
    (rustflagsSetsCodegenUnits ? process.env.RUSTFLAGS : undefined) ||
    process.env.CARGO_PROFILE_RELEASE_CODEGEN_UNITS ||
    process.env.RAYON_NUM_THREADS ||
    process.env.FORMULA_RAYON_NUM_THREADS ||
    process.env.BINARYEN_CORES ||
    process.env.FORMULA_BINARYEN_CORES;
  if (!userProvidedConcurrency && concurrency.jobs !== "1") {
    console.warn("[formula] wasm-pack build failed; retrying with CARGO_BUILD_JOBS=1 for stability.");
    await rm(buildOutDir, { recursive: true, force: true });
    result = runWasmPack({
      jobs: "1",
      makeflags: "-j1",
      releaseCodegenUnits: "1",
      rayonThreads: "1",
      binaryenCores: "1",
    });
  }
}

if (result.error) {
  fatal(`[formula] Failed to run wasm-pack: ${result.error.message}`);
}

if (result.status !== 0) {
  process.exit(result.status ?? 1);
}

if (!existsSync(buildWrapper) || !existsSync(buildWasm)) {
  const missing = [];
  if (!existsSync(buildWrapper)) missing.push(`Missing: ${path.relative(repoRoot, buildWrapper)}`);
  if (!existsSync(buildWasm)) missing.push(`Missing: ${path.relative(repoRoot, buildWasm)}`);
  fatal(["[formula] wasm-pack completed but build artifacts are missing.", ...missing].join("\n"));
}

// Swap the freshly built artifacts into place so downstream workflows (dev servers, tests)
// always see a complete `pkg/` directory.
try {
  if (existsSync(outDir)) {
    await rm(backupOutDir, { recursive: true, force: true });
    await rename(outDir, backupOutDir);
  }
  await rename(buildOutDir, outDir);
  await rm(backupOutDir, { recursive: true, force: true });
} catch (err) {
  console.error("[formula] Failed to swap WASM build output into place:", err);
  try {
    // Best-effort restore of the previous artifacts.
    if (!existsSync(outDir) && existsSync(backupOutDir)) {
      await rename(backupOutDir, outDir);
    }
  } catch {
    // ignore
  }
  process.exit(1);
}

if (!existsSync(wrapper) || !existsSync(wasm)) {
  const missing = [];
  if (!existsSync(wrapper)) missing.push(`Missing: ${path.relative(repoRoot, wrapper)}`);
  if (!existsSync(wasm)) missing.push(`Missing: ${path.relative(repoRoot, wasm)}`);
  fatal(["[formula] WASM output swap succeeded but expected artifacts are missing.", ...missing].join("\n"));
}

await copyToPublic();
