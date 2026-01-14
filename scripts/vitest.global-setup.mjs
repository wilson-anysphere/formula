import { ensureFormulaWasmNodeBuild } from "./build-formula-wasm-node.mjs";
import { existsSync, readFileSync, statSync } from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

export default async function globalSetup() {
  // Keep `pnpm test:vitest --silent` readable by default. Individual suites can
  // override by setting LOG_LEVEL explicitly (or by constructing a logger with an
  // explicit level/stream).
  if (!process.env.LOG_LEVEL) {
    process.env.LOG_LEVEL = "silent";
  }

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

    let abs = path.isAbsolute(arg) ? arg : path.resolve(process.cwd(), arg);
    if (!existsSync(abs)) {
      // Users sometimes pass repo-rooted paths while running vitest from a package subdirectory.
      // (e.g. `vitest run packages/ai-audit/test/export.test.ts` from `packages/ai-audit/`).
      // Fall back to resolving the path relative to the repo root before treating it as a
      // non-path selector.
      const repoAbs = path.resolve(repoRoot, arg);
      if (existsSync(repoAbs)) {
        abs = repoAbs;
      } else {
        // If this doesn't exist, treat it as a non-path selector; keep the conservative default.
        return false;
      }
    }
    resolved.push(abs);
  }

  const needsWasmBuild = resolved.some((abs) => {
    const resolvedAbs = path.resolve(abs);
    const normalized = ensureTrailingSep(resolvedAbs);

    // The formula-wasm crate itself always requires a wasm build.
    if (overlapsPath(normalized, formulaWasmDir)) {
      return true;
    }

    // Engine tests are a mix of pure unit tests (no wasm) and wasm-backed integration suites.
    // To keep `vitest run packages/engine/src/foo.test.ts` usable in environments without
    // wasm-pack, only require the wasm build when we detect a wasm-backed test.
    if (overlapsPath(normalized, engineDir)) {
      try {
        const stat = statSync(resolvedAbs);
        if (stat.isDirectory()) {
          // Directories (or ancestor paths like `vitest run packages`) may contain wasm-backed suites;
          // keep the conservative default.
          return true;
        }
      } catch {
        // If we can't stat it, keep the conservative default.
        return true;
      }

      // File path heuristics: `.wasm.test.ts(x)` is used for suites that load the wasm bundle.
      const basename = path.basename(resolvedAbs);
      if (basename.includes(".wasm.test.")) {
        return true;
      }

      // Content heuristics: most wasm-backed engine tests load the Node-compatible wasm bundle
      // using the shared helper `formulaWasmNodeEntryUrl()`.
      try {
        const text = readFileSync(resolvedAbs, "utf8");
        if (
          text.includes("formulaWasmNodeEntryUrl") ||
          text.includes("ensureFormulaWasmNodeBuild") ||
          text.includes("crates/formula-wasm") ||
          text.includes("pkg-node")
        ) {
          return true;
        }
      } catch {
        // If we can't read the file, keep the conservative default.
        return true;
      }

      return false;
    }

    return false;
  });

  return !needsWasmBuild;
}

function ensureTrailingSep(p) {
  return p.endsWith(path.sep) ? p : `${p}${path.sep}`;
}

function overlapsPath(candidatePath, dirPath) {
  // Candidate may be a file or directory path. Consider it overlapping if:
  // - it's inside the wasm-backed suite dir, OR
  // - it's an ancestor that contains the suite dir (e.g. `vitest run packages`).
  return candidatePath.startsWith(dirPath) || dirPath.startsWith(candidatePath);
}
