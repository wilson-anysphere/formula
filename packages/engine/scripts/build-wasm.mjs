import { spawnSync } from "node:child_process";
import { existsSync } from "node:fs";
import { copyFile, mkdir, readdir, rm, stat } from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

// `packages/engine/scripts/*` → repo root
const repoRoot = path.resolve(__dirname, "..", "..", "..");

const crateDir = path.join(repoRoot, "crates", "formula-wasm");
const coreDir = path.join(repoRoot, "crates", "formula-core");

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

  let latest = info.mtimeMs;
  const entries = await readdir(entryPath, { withFileTypes: true });
  for (const entry of entries) {
    const childPath = path.join(entryPath, entry.name);
    const mtime = await latestMtime(childPath);
    latest = Math.max(latest, mtime);
  }

  return latest;
}

async function copyToPublic() {
  for (const targetDir of targets) {
    await mkdir(targetDir, { recursive: true });
    await copyFile(wrapper, path.join(targetDir, "formula_wasm.js"));
    await copyFile(wasm, path.join(targetDir, "formula_wasm_bg.wasm"));
  }
}

// Validate `wasm-pack` is installed early with a good error message.
{
  const check = spawnSync(wasmPackBin, ["--version"], { encoding: "utf8" });
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

// Ensure the crate path exists (helps when running from unexpected working dirs).
if (!existsSync(path.join(crateDir, "Cargo.toml"))) {
  fatal(
    `[formula] Expected WASM crate at ${path.relative(repoRoot, crateDir)} (relative to repo root), but it was not found.`
  );
}

const outputExists = existsSync(wrapper) && existsSync(wasm);
if (outputExists) {
  const outputStamp = Math.min((await stat(wrapper)).mtimeMs, (await stat(wasm)).mtimeMs);
  const sourceStamp = Math.max(
    await latestMtime(crateDir),
    await latestMtime(coreDir),
    await latestMtime(path.join(repoRoot, "Cargo.lock"))
  );

  if (outputStamp >= sourceStamp) {
    await copyToPublic();
    process.exit(0);
  }
}

// `wasm-pack` refuses to overwrite some files if the output already exists.
await rm(outDir, { recursive: true, force: true });

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
  { cwd: repoRoot, stdio: "inherit" }
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
