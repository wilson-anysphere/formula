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

console.log("[formula] WASM artifacts present.");
