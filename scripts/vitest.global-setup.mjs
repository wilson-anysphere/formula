import { ensureFormulaWasmNodeBuild } from "./build-formula-wasm-node.mjs";

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

  // Only treat args as explicit test paths when they're concrete filesystem-like paths
  // (not glob patterns / regex filters).
  const looksLikeConcretePath = (arg) =>
    (arg.includes("/") || arg.includes("\\")) &&
    !arg.includes("*") &&
    !arg.includes("?") &&
    !arg.includes("{") &&
    !arg.includes("}") &&
    !arg.includes("[") &&
    !arg.includes("]") &&
    !arg.includes("(") &&
    !arg.includes(")") &&
    !arg.includes("|");

  if (!filtered.every(looksLikeConcretePath)) {
    // Ambiguous selection (globs / patterns); default to building so wasm-backed
    // tests don't fail with missing artifacts.
    return false;
  }

  const needsWasmBuild = filtered.some((arg) => {
    const normalized = arg.replaceAll("\\", "/");
    return normalized.includes("/packages/engine/") || normalized.includes("/crates/formula-wasm/");
  });

  return !needsWasmBuild;
}
