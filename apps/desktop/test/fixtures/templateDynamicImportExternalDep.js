export async function loadEsbuild() {
  // `esbuild` is a representative third-party dependency used by the node:test runners to
  // detect whether external packages are installed. When `node_modules` is missing, the
  // runner should skip any test files that reach this import.
  const mod = await import("esbuild");
  return mod?.build ? mod : mod?.default;
}

