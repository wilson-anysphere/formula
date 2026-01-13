import { ensureFormulaWasmNodeBuild } from "./build-formula-wasm-node.mjs";
import { existsSync } from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

export default async function globalSetup() {
  if (process.env.FORMULA_SKIP_WASM_BUILD === "1" || process.env.FORMULA_SKIP_WASM_BUILD === "true") {
    return;
  }

  if (shouldSkipWasmBuildForCurrentRun()) {
    return;
  }

  ensureFormulaWasmNodeBuild();
}

function shouldSkipWasmBuildForCurrentRun() {
  // Vitest's globalSetup runs even when executing a single test file (e.g.
  // `pnpm vitest run packages/ai-audit/test/export.test.ts`). Building the
  // formula-wasm Node bundle requires a Rust+wasm toolchain, which is not
  // necessary for the vast majority of unit tests.
  //
  // To keep `vitest run <path>` usable in environments without wasm-pack,
  // only trigger the wasm build when:
  //   - no explicit test paths were provided (full test suite), OR
  //   - any provided path suggests we are running engine/wasm-backed tests.
  const args = process.argv.slice(2);
  const positional = args.filter((arg) => typeof arg === "string" && !arg.startsWith("-"));
  const filtered = positional.filter((arg) => arg !== "run" && arg !== "watch" && arg !== "dev");

  if (filtered.length === 0) {
    // Full suite run; keep the existing behavior (ensure wasm bundle is available).
    return false;
  }

  const repoRoot = fileURLToPath(new URL("..", import.meta.url));
  const engineDir = ensureTrailingSep(path.resolve(repoRoot, "packages", "engine"));
  const formulaWasmDir = ensureTrailingSep(path.resolve(repoRoot, "crates", "formula-wasm"));

  // Only skip the build when the user provided *concrete* filesystem paths (not
  // globs / regex patterns), and none of those paths touch wasm-backed suites.
  const isAmbiguousSelector = (arg) =>
    arg.includes("*") ||
    arg.includes("?") ||
    arg.includes("{") ||
    arg.includes("}") ||
    arg.includes("[") ||
    arg.includes("]") ||
    arg.includes("(") ||
    arg.includes(")") ||
    arg.includes("|");

  const resolved = [];
  for (const arg of filtered) {
    if (isAmbiguousSelector(arg)) {
      // Globs/patterns may include wasm-backed tests; keep the conservative default.
      return false;
    }

    const abs = path.isAbsolute(arg) ? arg : path.resolve(process.cwd(), arg);
    if (!existsSync(abs)) {
      // If this doesn't exist, treat it as a non-path selector; keep the conservative default.
      return false;
    }
    resolved.push(abs);
  }

  const needsWasmBuild = resolved.some((abs) => {
    const normalized = ensureTrailingSep(path.resolve(abs));
    return normalized.startsWith(engineDir) || normalized.startsWith(formulaWasmDir);
  });

  return !needsWasmBuild;
}

function ensureTrailingSep(p) {
  return p.endsWith(path.sep) ? p : `${p}${path.sep}`;
}
