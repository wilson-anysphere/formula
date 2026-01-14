import { spawnSync } from "node:child_process";
import { existsSync, mkdirSync, readFileSync, readdirSync, rmSync, statSync } from "node:fs";
import os from "node:os";
import path from "node:path";
import { fileURLToPath, pathToFileURL } from "node:url";

const repoRoot = path.resolve(fileURLToPath(new URL("..", import.meta.url)));
const crateDir = path.join(repoRoot, "crates", "formula-wasm");

// `RUSTUP_TOOLCHAIN` overrides the repo's `rust-toolchain.toml` pin. Some environments set it
// globally (often to `stable`), which would bypass the pinned toolchain and reintroduce drift.
if (process.env.RUSTUP_TOOLCHAIN && existsSync(path.join(repoRoot, "rust-toolchain.toml"))) {
  delete process.env.RUSTUP_TOOLCHAIN;
}
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
mkdirSync(cargoHome, { recursive: true });
const cargoBinDir = path.join(cargoHome, "bin");
mkdirSync(cargoBinDir, { recursive: true });
if (process.env.PATH) {
  const pathEntries = process.env.PATH.split(path.delimiter);
  if (!pathEntries.includes(cargoBinDir)) {
    process.env.PATH = `${cargoBinDir}${path.delimiter}${process.env.PATH}`;
  }
} else {
  process.env.PATH = cargoBinDir;
}

// Default to conservative parallelism. On high-core-count machines the Rust linker
// (lld) can spawn many threads per invocation; combining that with Cargo-level
// parallelism can exceed sandbox process/thread limits and cause flaky "Resource
// temporarily unavailable" failures.
const cpuCount = os.cpus().length;
const defaultJobs = cpuCount >= 64 ? "2" : "4";

const outDir = path.join(crateDir, "pkg-node");
const outPackageJsonPath = path.join(outDir, "package.json");
const outEntryJsPath = path.join(outDir, "formula_wasm.js");
const outWasmPath = path.join(outDir, "formula_wasm_bg.wasm");
const releaseWorkflowPath = path.join(repoRoot, ".github", "workflows", "release.yml");

function readPinnedWasmPackVersion() {
  try {
    const text = readFileSync(releaseWorkflowPath, "utf8");
    const match = text.match(/^[\t ]*WASM_PACK_VERSION:[\t ]*["']?([^"'\n]+)["']?/m);
    return match ? match[1].trim() : null;
  } catch {
    return null;
  }
}

/**
 * Ensure we have a Node-compatible (wasm-bindgen `--target nodejs`) build of
 * `crates/formula-wasm` available for vitest/Node consumers.
 *
 * The output directory is intentionally stable + gitignored:
 *   `crates/formula-wasm/pkg-node/`
 *
 * Importing from TypeScript (ESM) in Node/Vitest:
 *
 * ```ts
 * import { pathToFileURL } from "node:url";
 * import path from "node:path";
 *
 * const entry = pathToFileURL(
 *   path.resolve("crates/formula-wasm/pkg-node/formula_wasm.js")
 * ).href;
 *
 * // `--target nodejs` generates CommonJS; ESM gets it under `default`.
 * // (If Vite tries to pre-bundle the file URL, add the `@vite-ignore` magic
 * // comment inside the import call.)
 * const mod = await import(entry);
 * const wasm = (mod as any).default ?? mod;
 *
 * const workbook = new wasm.WasmWorkbook();
 * ```
 *
 * @param {{ force?: boolean }} [options]
 * @returns {{ outDir: string; entryJsPath: string; rebuilt: boolean }}
 */
export function ensureFormulaWasmNodeBuild(options = {}) {
  const force = options.force === true;

  const outputsExist = existsSync(outPackageJsonPath) && existsSync(outEntryJsPath) && existsSync(outWasmPath);
  const outputsStale = outputsExist ? isOutputStale() : true;

  if (!force && outputsExist && !outputsStale) {
    return { outDir, entryJsPath: outEntryJsPath, rebuilt: false };
  }

  assertPrereqs();
  buildWithWasmPack();
  assertOutputsExist();
  return { outDir, entryJsPath: outEntryJsPath, rebuilt: true };
}

/**
 * @returns {string} file:// URL to the JS entry point (`formula_wasm.js`).
 */
export function formulaWasmNodeEntryUrl() {
  return pathToFileURL(outEntryJsPath).href;
}

function assertPrereqs() {
  const cargoAgent = path.relative(process.cwd(), path.join(repoRoot, "scripts", "cargo_agent.sh"));
  const cargoAgentCmd = cargoAgent || path.join(repoRoot, "scripts", "cargo_agent.sh");
  const pinnedWasmPack = readPinnedWasmPackVersion();
  const pinnedHint = pinnedWasmPack
    ? `  bash ${cargoAgentCmd} install wasm-pack --version "${pinnedWasmPack}" --locked --force`
    : `  bash ${cargoAgentCmd} install wasm-pack`;
  assertCommand("wasm-pack", ["--version"], `Missing \`wasm-pack\`.

Install it from https://rustwasm.github.io/wasm-pack/installer/ (recommended),
or via the repo cargo wrapper (agent-safe):
${pinnedHint}
`);

  // wasm-pack ultimately needs this rust target.
  const rustup = spawnSync("rustup", ["target", "list", "--installed"], {
    encoding: "utf8",
    // Run from the repo root so rustup can reliably pick up `rust-toolchain.toml` even if callers
    // execute this script from another working directory.
    cwd: repoRoot,
  });
  if (rustup.error && rustup.error.code === "ENOENT") {
    throw new Error(`Missing \`rustup\` (required to validate/install the wasm32 target).

Install Rust via https://rustup.rs/ then run:
  rustup target add wasm32-unknown-unknown
`);
  }
  if (rustup.status !== 0) {
    throw new Error(`Failed to run \`rustup target list --installed\` while preparing formula-wasm.

stderr:
${rustup.stderr || "<empty>"}
`);
  }
  if (!String(rustup.stdout).split(/\s+/).includes("wasm32-unknown-unknown")) {
    throw new Error(`Rust target \`wasm32-unknown-unknown\` is not installed.

Install it with:
  rustup target add wasm32-unknown-unknown
`);
  }
}

/**
 * @param {string} cmd
 * @param {string[]} args
 * @param {string} message
 */
function assertCommand(cmd, args, message) {
  const res = spawnSync(cmd, args, { stdio: "ignore" });
  if (res.error && res.error.code === "ENOENT") {
    throw new Error(message);
  }
  if (res.status !== 0) {
    throw new Error(`${message.trim()}

(\`${cmd} ${args.join(" ")}\` exited with code ${res.status ?? "unknown"})`);
  }
}

function buildWithWasmPack() {
  // wasm-pack will read an existing `package.json` from the output directory and treat it as a
  // template. Its parser is stricter than npm's (e.g. `repository` must be a string, not the
  // `{ type, url }` object form), which can make incremental builds fail after a successful build.
  // Since we only call wasm-pack when we intend to rebuild, wipe the output dir to match CI's
  // clean builds and avoid brittle template parsing.
  rmSync(outDir, { recursive: true, force: true });

  const jobs = process.env.FORMULA_CARGO_JOBS ?? process.env.CARGO_BUILD_JOBS ?? defaultJobs;
  const makeflags = process.env.MAKEFLAGS ?? `-j${jobs}`;
  const rayonThreads = process.env.RAYON_NUM_THREADS ?? process.env.FORMULA_RAYON_NUM_THREADS ?? jobs;
  const limitAs = process.env.FORMULA_CARGO_LIMIT_AS ?? "14G";
  const runLimited = path.join(repoRoot, "scripts", "run_limited.sh");
  const canUseRunLimited = process.platform !== "win32" && existsSync(runLimited);
  const verbose =
    process.env.FORMULA_WASM_PACK_VERBOSE === "1" || process.env.FORMULA_WASM_PACK_VERBOSE === "true";
  // `wasm-pack build` inherits cargo's per-crate compile output, which is very verbose. In CI/agent
  // environments where stdout isn't a TTY this can create enormous logs and even hit output capture
  // limits. Pass `--quiet` through to cargo in those cases unless callers explicitly opt into
  // verbose output.
  const cargoExtraArgs = ["--locked"];
  if (!verbose && !process.stdout.isTTY) cargoExtraArgs.push("--quiet");
  const rustcWrapperRaw = process.env.RUSTC_WRAPPER ?? process.env.CARGO_BUILD_RUSTC_WRAPPER ?? "";
  const rustcWorkspaceWrapperRaw =
    process.env.RUSTC_WORKSPACE_WRAPPER ?? process.env.CARGO_BUILD_RUSTC_WORKSPACE_WRAPPER ?? "";
  const rustcWrapper = normalizeRustcWrapper(rustcWrapperRaw);
  const rustcWorkspaceWrapper = normalizeRustcWrapper(rustcWorkspaceWrapperRaw);

  // Equivalent to: `wasm-pack build crates/formula-wasm --target nodejs --out-dir pkg-node`
  // but avoids any ambiguity around relative output paths by running from the crate dir.
  //
  // Note: some environments configure Cargo to use `sccache` via `build.rustc-wrapper`.
  // When `sccache` is unavailable/misconfigured, wasm-pack builds can fail even for
  // `cargo metadata`/`rustc -vV`. Explicitly setting `RUSTC_WRAPPER=""` disables
  // any configured wrapper unless the user overrides it in the environment.
  //
  // NOTE: Cargo treats empty wrapper env vars as "unset" and can fall back to a global
  // `build.rustc-wrapper` config. To reliably bypass global wrappers we default to using
  // the `env` wrapper, which simply execs the underlying rustc (`env rustc ...`).
  const wasmPackArgs = ["build", "--target", "nodejs", "--out-dir", "pkg-node", "--dev", ...cargoExtraArgs];
  const baseEnv = {
    ...process.env,
    CARGO_HOME: cargoHome,
    // Keep builds safe in high-core-count environments (e.g. agent sandboxes) even
    // if the caller didn't initialize via `scripts/agent-init.sh`.
    CARGO_BUILD_JOBS: jobs,
    MAKEFLAGS: makeflags,
    CARGO_PROFILE_DEV_CODEGEN_UNITS: process.env.CARGO_PROFILE_DEV_CODEGEN_UNITS ?? jobs,
    // Rayon defaults to spawning one worker per core; cap it for multi-agent hosts unless
    // callers explicitly override it.
    RAYON_NUM_THREADS: rayonThreads,
  };

  const run = (env) =>
    canUseRunLimited
      ? spawnSync("bash", [runLimited, "--as", limitAs, "--", "wasm-pack", ...wasmPackArgs], {
          cwd: crateDir,
          env,
          stdio: "inherit",
        })
      : spawnSync("wasm-pack", wasmPackArgs, {
          cwd: crateDir,
          env,
          stdio: "inherit",
        });

  const res = run(withRustcWrappers(baseEnv, rustcWrapper, rustcWorkspaceWrapper));
  if (res.error) throw res.error;
  if (res.status === 0) return;

  // If a host-configured sccache daemon dies under load, builds can fail spuriously with
  // "failed to execute compile" / connection reset errors. Retry once with wrappers disabled
  // so vitest can proceed (even if caching is unavailable).
  const wrapperLooksLikeSccache =
    rustcWrapperRaw.trim().toLowerCase().includes("sccache") || rustcWorkspaceWrapperRaw.trim().toLowerCase().includes("sccache");
  if (wrapperLooksLikeSccache && process.platform !== "win32") {
    const fallback = "env";
    const retry = run(withRustcWrappers(baseEnv, fallback, fallback));
    if (retry.error) throw retry.error;
    if (retry.status === 0) return;
  }

  throw new Error(`wasm-pack build failed (exit code ${res.status ?? "unknown"})`);
}

function assertOutputsExist() {
  const missing = [];
  if (!existsSync(outPackageJsonPath)) missing.push(outPackageJsonPath);
  if (!existsSync(outEntryJsPath)) missing.push(outEntryJsPath);
  if (!existsSync(outWasmPath)) missing.push(outWasmPath);
  if (missing.length > 0) {
    throw new Error(`formula-wasm Node build completed but expected output files are missing:
${missing.map((p) => `  - ${p}`).join("\n")}
`);
  }
}

function isOutputStale() {
  const outputMtimeMs = Math.min(
    statSync(outPackageJsonPath).mtimeMs,
    statSync(outEntryJsPath).mtimeMs,
    statSync(outWasmPath).mtimeMs
  );

  const inputMtimeMs = newestInputMtimeMs();
  return inputMtimeMs > outputMtimeMs;
}

function newestInputMtimeMs() {
  const watchedCrates = collectPathDependencyClosure(crateDir);
  const inputs = [path.join(repoRoot, "Cargo.lock")];

  for (const dir of watchedCrates) {
    inputs.push(path.join(dir, "Cargo.toml"));
    inputs.push(path.join(dir, "src"));
  }

  return newestMtimeMs(inputs);
}

/**
 * Normalize a rustc wrapper value for use by Cargo.
 *
 * Cargo treats an empty wrapper env var as "unset" and can fall back to a global
 * `build.rustc-wrapper` config. When no wrapper is provided, use `env` to reliably
 * bypass any global wrapper configuration.
 */
function normalizeRustcWrapper(wrapper) {
  const trimmed = typeof wrapper === "string" ? wrapper.trim() : "";
  if (!trimmed) {
    // `env` exists on Unix-like systems. On Windows we return an empty string and rely on the
    // existing environment/configuration.
    return process.platform === "win32" ? "" : "env";
  }
  return trimmed;
}

function withRustcWrappers(baseEnv, rustcWrapper, rustcWorkspaceWrapper) {
  return {
    ...baseEnv,
    RUSTC_WRAPPER: rustcWrapper,
    RUSTC_WORKSPACE_WRAPPER: rustcWorkspaceWrapper,
    // Cargo can also read wrapper settings via `CARGO_BUILD_RUSTC_WRAPPER`. Set it explicitly so
    // a global Cargo config (`build.rustc-wrapper`) cannot accidentally re-enable a flaky wrapper.
    CARGO_BUILD_RUSTC_WRAPPER: rustcWrapper,
    CARGO_BUILD_RUSTC_WORKSPACE_WRAPPER: rustcWorkspaceWrapper,
  };
}

/**
 * @param {string[]} roots
 * @returns {number}
 */
function newestMtimeMs(roots) {
  let newest = 0;
  for (const root of roots) {
    newest = Math.max(newest, newestMtimeMsRecursive(root));
  }
  return newest;
}

/**
 * @param {string} entry
 * @returns {number}
 */
function newestMtimeMsRecursive(entry) {
  if (!existsSync(entry)) return 0;
  const stat = statSync(entry);
  if (stat.isFile()) return stat.mtimeMs;
  if (!stat.isDirectory()) return stat.mtimeMs;

  let newest = stat.mtimeMs;
  const entries = readdirSync(entry, { withFileTypes: true });
  for (const child of entries) {
    // Skip the build outputs if someone runs this from the crate directory.
    if (child.isDirectory() && (child.name === "target" || child.name === "pkg" || child.name === "pkg-node")) {
      continue;
    }
    newest = Math.max(newest, newestMtimeMsRecursive(path.join(entry, child.name)));
  }
  return newest;
}

/**
 * Collect `crateDir` and any transitive path dependencies declared in Cargo.toml.
 * This lets us skip wasm-pack builds when the Rust sources haven't changed.
 *
 * @param {string} rootCrateDir
 * @returns {Set<string>}
 */
function collectPathDependencyClosure(rootCrateDir) {
  const visited = new Set();
  const queue = [rootCrateDir];

  while (queue.length > 0) {
    const next = queue.pop();
    if (!next || visited.has(next)) continue;
    visited.add(next);

    const cargoTomlPath = path.join(next, "Cargo.toml");
    if (!existsSync(cargoTomlPath)) continue;

    const toml = readFileSync(cargoTomlPath, "utf8");
    for (const depPath of parseTomlPathDependencies(toml)) {
      const resolved = path.resolve(next, depPath);
      if (!resolved.startsWith(repoRoot)) continue;
      if (!existsSync(path.join(resolved, "Cargo.toml"))) continue;
      queue.push(resolved);
    }
  }

  return visited;
}

/**
 * Extremely small TOML "parser" to find path dependencies. We only need to
 * understand `path = "../some-crate"` occurrences.
 *
 * @param {string} toml
 * @returns {string[]}
 */
function parseTomlPathDependencies(toml) {
  /** @type {string[]} */
  const paths = [];
  // We only need best-effort detection of `path = "../some-crate"` dependencies. Keep it
  // line-based so we can cheaply ignore TOML comments without pulling in a full parser.
  //
  // Note: Use both quote styles; Cargo accepts either in `Cargo.toml`.
  const regex = /path\s*=\s*(?:"([^"]+)"|'([^']+)')/g;
  for (const rawLine of toml.split(/\r?\n/)) {
    // Strip trailing comments. (This is intentionally simplistic; path values should not
    // contain `#` in practice.)
    const line = rawLine.split("#")[0];
    if (!line) continue;
    regex.lastIndex = 0;
    let match;
    while ((match = regex.exec(line))) {
      paths.push(match[1] ?? match[2]);
    }
  }
  return paths;
}

if (process.argv[1] && path.resolve(process.argv[1]) === fileURLToPath(import.meta.url)) {
  try {
    const { rebuilt } = ensureFormulaWasmNodeBuild();
    console.log(rebuilt ? `Built formula-wasm (Node) -> ${outDir}` : `formula-wasm (Node) is up to date -> ${outDir}`);
  } catch (err) {
    const message = err instanceof Error ? err.message : String(err);
    console.error(message);
    process.exitCode = 1;
  }
}
