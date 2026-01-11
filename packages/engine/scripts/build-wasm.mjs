import { spawnSync } from "node:child_process";
import { existsSync, rmSync } from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

// `packages/engine/scripts/*` → repo root
const repoRoot = path.resolve(__dirname, "..", "..", "..");

const cratePath = "crates/formula-wasm";
const outDir = "packages/engine/pkg";
const outName = "formula_wasm";

const crateAbs = path.join(repoRoot, cratePath);
const outDirAbs = path.join(repoRoot, outDir);
const wrapperAbs = path.join(outDirAbs, `${outName}.js`);
const wasmAbs = path.join(outDirAbs, `${outName}_bg.wasm`);

const wasmPackBin = process.platform === "win32" ? "wasm-pack.exe" : "wasm-pack";

function fatal(message) {
  console.error(message);
  process.exit(1);
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
        `Original error: ${check.error.message}`
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
if (!existsSync(path.join(crateAbs, "Cargo.toml"))) {
  fatal(`[formula] Expected WASM crate at ${cratePath} (relative to repo root), but it was not found.`);
}

// `wasm-pack` refuses to overwrite some files if the output already exists.
rmSync(outDirAbs, { recursive: true, force: true });

const result = spawnSync(
  wasmPackBin,
  ["build", crateAbs, "--target", "web", "--out-dir", outDirAbs, "--out-name", outName],
  { cwd: repoRoot, stdio: "inherit" }
);

if (result.error) {
  fatal(`[formula] Failed to run wasm-pack: ${result.error.message}`);
}

if (result.status !== 0) {
  process.exit(result.status ?? 1);
}

if (!existsSync(wrapperAbs) || !existsSync(wasmAbs)) {
  const missing = [];
  if (!existsSync(wrapperAbs)) missing.push(`Missing: ${path.relative(repoRoot, wrapperAbs)}`);
  if (!existsSync(wasmAbs)) missing.push(`Missing: ${path.relative(repoRoot, wasmAbs)}`);
  fatal(
    ["[formula] wasm-pack completed but expected artifacts are missing.", ...missing].join("\n")
  );
}

console.log(`[formula] Built WASM package: ${path.relative(repoRoot, wrapperAbs)}`);
