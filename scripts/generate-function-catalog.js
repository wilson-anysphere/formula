import { spawn } from "node:child_process";
import { existsSync } from "node:fs";
import { mkdir, writeFile } from "node:fs/promises";
import os from "node:os";
import path from "node:path";
import { fileURLToPath } from "node:url";

// Generates `shared/functionCatalog.json` (and a JS-friendly wrapper module)
// by enumerating the Rust formula engine's
// inventory-backed registry of built-in functions.
//
// This is intentionally opt-in. CI/tests consume the committed JSON artifact so
// JavaScript/TypeScript workflows do not require compiling Rust.
const repoRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const outputJsonPath = path.join(repoRoot, "shared", "functionCatalog.json");
const outputModulePath = path.join(repoRoot, "shared", "functionCatalog.mjs");
// When the runtime module is `.mjs`, TypeScript expects an ESM-flavored declaration file (`.d.mts`).
const outputTypesPath = path.join(repoRoot, "shared", "functionCatalog.d.mts");
const outputNamesModulePath = path.join(repoRoot, "shared", "functionNames.mjs");
const outputNamesTypesPath = path.join(repoRoot, "shared", "functionNames.d.mts");

const baseEnv = { ...process.env };
// `RUSTUP_TOOLCHAIN` overrides the repo's `rust-toolchain.toml` pin. Some environments set it
// globally (often to `stable`), which would bypass the pinned toolchain and reintroduce drift when
// this script falls back to invoking `cargo` directly (notably on Windows).
if (baseEnv.RUSTUP_TOOLCHAIN && existsSync(path.join(repoRoot, "rust-toolchain.toml"))) {
  delete baseEnv.RUSTUP_TOOLCHAIN;
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
await mkdir(cargoHome, { recursive: true });

/**
 * @param {string} command
 * @param {string[]} args
 * @param {{ env?: NodeJS.ProcessEnv }} [options]
 * @returns {Promise<string>}
 */
function run(command, args, options = {}) {
  return new Promise((resolve, reject) => {
    const child = spawn(command, args, {
      cwd: repoRoot,
      stdio: ["ignore", "pipe", "inherit"],
      env: options.env ?? process.env,
    });

    /** @type {Buffer[]} */
    const chunks = [];
    child.stdout.on("data", (chunk) => chunks.push(chunk));
    child.on("error", reject);
    child.on("close", (code) => {
      if (code !== 0) {
        reject(new Error(`${command} ${args.join(" ")} exited with code ${code}`));
        return;
      }
      resolve(Buffer.concat(chunks).toString("utf8"));
    });
  });
}

function isCatalogShape(value) {
  return value && typeof value === "object" && Array.isArray(value.functions);
}

function isCatalogValueType(value) {
  return value === "any" || value === "number" || value === "text" || value === "bool";
}

function isCatalogVolatility(value) {
  return value === "non_volatile" || value === "volatile";
}

// Default to conservative parallelism. On very high core-count hosts, the Rust linker
// (lld) can spawn many worker threads per link step; combining that with Cargo-level
// parallelism can exceed sandbox process/thread limits and cause flaky
// "Resource temporarily unavailable" failures.
const cpuCount = os.cpus().length;
const defaultJobs = cpuCount >= 64 ? "2" : "4";

const jobs = process.env.FORMULA_CARGO_JOBS ?? process.env.CARGO_BUILD_JOBS ?? defaultJobs;
const rayonThreads = process.env.RAYON_NUM_THREADS ?? process.env.FORMULA_RAYON_NUM_THREADS ?? jobs;
// Some environments configure Cargo to use `sccache` via `build.rustc-wrapper` (or equivalent
// env-var configuration). Default to disabling any configured wrapper unless the user explicitly
// opts in via environment variables.
const rustcWrapper = process.env.RUSTC_WRAPPER ?? process.env.CARGO_BUILD_RUSTC_WRAPPER ?? "";
const rustcWorkspaceWrapper =
  process.env.RUSTC_WORKSPACE_WRAPPER ?? process.env.CARGO_BUILD_RUSTC_WORKSPACE_WRAPPER ?? "";

const cargoAgentPath = path.join(repoRoot, "scripts", "cargo_agent.sh");
const useCargoAgent = process.platform !== "win32" && existsSync(cargoAgentPath);

const cargoCommand = useCargoAgent ? "bash" : "cargo";
const cargoArgs = useCargoAgent
  ? [
      cargoAgentPath,
      "run",
      "--quiet",
      "--locked",
      "-p",
      "formula-engine",
      "--bin",
      "function_catalog",
    ]
  : ["run", "--quiet", "--locked", "-p", "formula-engine", "--bin", "function_catalog"];

const raw = await run(cargoCommand, cargoArgs, {
  env: {
    ...baseEnv,
    CARGO_HOME: cargoHome,
    // Keep builds safe in high-core-count environments (e.g. agent sandboxes) even
    // if the caller didn't initialize via `scripts/agent-init.sh`.
    CARGO_BUILD_JOBS: jobs,
    MAKEFLAGS: process.env.MAKEFLAGS ?? `-j${jobs}`,
    CARGO_PROFILE_DEV_CODEGEN_UNITS: process.env.CARGO_PROFILE_DEV_CODEGEN_UNITS ?? jobs,
    // Rayon defaults to spawning one worker per core; cap it for multi-agent hosts unless
    // callers explicitly override it.
    RAYON_NUM_THREADS: rayonThreads,
    RUSTC_WRAPPER: rustcWrapper,
    RUSTC_WORKSPACE_WRAPPER: rustcWorkspaceWrapper,
    // Cargo config can also be set via `CARGO_BUILD_RUSTC_WRAPPER`; include it so we reliably
    // override global config (and avoid flaky sccache wrappers) when the caller didn't opt in.
    CARGO_BUILD_RUSTC_WRAPPER: rustcWrapper,
    CARGO_BUILD_RUSTC_WORKSPACE_WRAPPER: rustcWorkspaceWrapper,
  },
});

/** @type {any} */
let parsed;
try {
  parsed = JSON.parse(raw);
} catch (err) {
  throw new Error(`Rust function_catalog output was not valid JSON: ${err}`);
}

if (!isCatalogShape(parsed)) {
  throw new Error("Rust function_catalog output did not match expected shape: { functions: [...] }");
}

for (const entry of parsed.functions) {
  if (!entry || typeof entry.name !== "string" || entry.name.length === 0) {
    throw new Error("Rust function_catalog output contained invalid function entry");
  }
  if (!Number.isInteger(entry.min_args) || entry.min_args < 0) {
    throw new Error(`Rust function_catalog output contained invalid min_args for ${entry.name}`);
  }
  if (!Number.isInteger(entry.max_args) || entry.max_args < entry.min_args) {
    throw new Error(`Rust function_catalog output contained invalid max_args for ${entry.name}`);
  }
  if (!isCatalogVolatility(entry.volatility)) {
    throw new Error(`Rust function_catalog output contained invalid volatility for ${entry.name}`);
  }
  if (!isCatalogValueType(entry.return_type)) {
    throw new Error(`Rust function_catalog output contained invalid return_type for ${entry.name}`);
  }
  if (!Array.isArray(entry.arg_types) || !entry.arg_types.every(isCatalogValueType)) {
    throw new Error(`Rust function_catalog output contained invalid arg_types for ${entry.name}`);
  }
}

await mkdir(path.dirname(outputJsonPath), { recursive: true });
await writeFile(outputJsonPath, JSON.stringify(parsed, null, 2) + "\n", "utf8");

// Node/Vite/TS interoperability note: JSON module import syntax has changed over
// time (`assert { type: "json" }` vs `with { type: "json" }`). To keep consumer
// code simple (and syntax-compatible across runtimes/tools), we also emit an
// `.mjs` wrapper that exports the same object.
const moduleContents = [
  "// This file is generated by scripts/generate-function-catalog.js. Do not edit.\n",
  `export default ${JSON.stringify(parsed, null, 2)};\n`,
].join("");
await writeFile(outputModulePath, moduleContents, "utf8");

const dtsContents = `// This file is generated by scripts/generate-function-catalog.js. Do not edit.
declare const catalog: {
  functions: Array<{
    name: string;
    min_args: number;
    max_args: number;
    arg_types: Array<"any" | "number" | "text" | "bool">;
    volatility: "non_volatile" | "volatile";
    return_type: "any" | "number" | "text" | "bool";
  }>;
};
export default catalog;
`;
await writeFile(outputTypesPath, dtsContents, "utf8");

// Keep the name list sorted for deterministic diffs. Use the default JS string sort
// (Unicode codepoint order) rather than localeCompare to avoid locale-dependent output.
const functionNames = parsed.functions.map((fn) => fn.name).sort();
const namesModuleContents = [
  "// This file is generated by scripts/generate-function-catalog.js. Do not edit.\n",
  `export default ${JSON.stringify(functionNames, null, 2)};\n`,
].join("");
await writeFile(outputNamesModulePath, namesModuleContents, "utf8");

const namesDtsContents = `// This file is generated by scripts/generate-function-catalog.js. Do not edit.
declare const names: string[];
export default names;
`;
await writeFile(outputNamesTypesPath, namesDtsContents, "utf8");
console.log(
  `Wrote ${path.relative(repoRoot, outputJsonPath)} + ${path.relative(repoRoot, outputModulePath)} + ${path.relative(repoRoot, outputTypesPath)} + ${path.relative(repoRoot, outputNamesModulePath)} + ${path.relative(repoRoot, outputNamesTypesPath)} (${parsed.functions.length} functions)`
);
