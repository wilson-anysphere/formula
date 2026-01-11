import { spawnSync } from "node:child_process";
import { existsSync } from "node:fs";
import { copyFile, mkdir, readFile, readdir, rm, stat } from "node:fs/promises";
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
const cargoBinDir = path.join(cargoHome, "bin");
await mkdir(cargoBinDir, { recursive: true });
if (!childEnv.PATH?.split(path.delimiter).includes(cargoBinDir)) {
  childEnv.PATH = childEnv.PATH ? `${cargoBinDir}${path.delimiter}${childEnv.PATH}` : cargoBinDir;
}

const crateDir = path.join(repoRoot, "crates", "formula-wasm");

const outDir = path.join(repoRoot, "packages", "engine", "pkg");
// Note: `wasm-pack build --out-dir` is documented as a *relative* path and is
// resolved from the crate directory, not `cwd`. Use an absolute path to ensure
// output always lands in this repo's deterministic location.
const wrapper = path.join(outDir, "formula_wasm.js");
const wasm = path.join(outDir, "formula_wasm_bg.wasm");

const targets = [
  path.join(repoRoot, "apps", "web", "public", "engine"),
  path.join(repoRoot, "apps", "desktop", "public", "engine")
];

const wasmPackBin = process.platform === "win32" ? "wasm-pack.exe" : "wasm-pack";

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

  if (outputStamp >= sourceStamp) {
    console.log("[formula] WASM artifacts up to date; copying runtime assets into apps/*/public/engine.");
    await copyToPublic();
    process.exit(0);
  }
}

// Validate `wasm-pack` is installed (only required when we need to rebuild).
{
  const check = spawnSync(wasmPackBin, ["--version"], { encoding: "utf8", env: childEnv });
  if (check.error) {
    fatal(
      [
        "[formula] wasm-pack is required to build the Rust/WASM engine but was not found on PATH.",
        "",
        "Install it with one of:",
        "  - cargo install wasm-pack",
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

// `wasm-pack` refuses to overwrite some files if the output already exists.
await rm(outDir, { recursive: true, force: true });

// Some environments configure Cargo to use `sccache` via `build.rustc-wrapper` or
// other wrapper settings. When the wrapper is unavailable/misconfigured, wasm-pack
// builds can fail even for `cargo metadata`/`rustc -vV`. Default to disabling any
// configured wrapper unless the user explicitly overrides it in the environment.
const wasmPackEnv = {
  ...childEnv,
  RUSTC_WRAPPER: process.env.RUSTC_WRAPPER ?? "",
  RUSTC_WORKSPACE_WRAPPER: process.env.RUSTC_WORKSPACE_WRAPPER ?? "",
};

const result = spawnSync(
  wasmPackBin,
  [
    "build",
    crateDir,
    "--target",
    "web",
    "--release",
    "--out-dir",
    outDir,
    "--out-name",
    "formula_wasm",
    // Avoid generating a nested package.json in the output directory; consumers
    // import the wrapper by URL and do not need `wasm-pack`'s npm packaging.
    "--no-pack"
  ],
  { cwd: repoRoot, stdio: "inherit", env: wasmPackEnv }
);

if (result.error) {
  fatal(`[formula] Failed to run wasm-pack: ${result.error.message}`);
}

if (result.status !== 0) {
  process.exit(result.status ?? 1);
}

if (!existsSync(wrapper) || !existsSync(wasm)) {
  const missing = [];
  if (!existsSync(wrapper)) missing.push(`Missing: ${path.relative(repoRoot, wrapper)}`);
  if (!existsSync(wasm)) missing.push(`Missing: ${path.relative(repoRoot, wasm)}`);
  fatal(["[formula] wasm-pack completed but expected artifacts are missing.", ...missing].join("\n"));
}

await copyToPublic();
