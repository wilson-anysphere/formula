import { accessSync, constants } from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

// `packages/engine/scripts/*` → repo root
const repoRoot = path.resolve(__dirname, "..", "..", "..");

const outDir = path.join(repoRoot, "packages", "engine", "pkg");
const wrapper = path.join(outDir, "formula_wasm.js");
const wasm = path.join(outDir, "formula_wasm_bg.wasm");

const publicTargets = [
  path.join(repoRoot, "apps", "web", "public", "engine"),
  path.join(repoRoot, "apps", "desktop", "public", "engine")
];

function assertReadable(filePath) {
  try {
    accessSync(filePath, constants.R_OK);
  } catch {
    console.error(`[formula] Missing WASM artifact: ${path.relative(repoRoot, filePath)}`);
    console.error("Run `pnpm build:wasm` first.");
    process.exit(1);
  }
}

assertReadable(wrapper);
assertReadable(wasm);

for (const dir of publicTargets) {
  assertReadable(path.join(dir, "formula_wasm.js"));
  assertReadable(path.join(dir, "formula_wasm_bg.wasm"));
}

console.log("[formula] WASM artifacts present.");
