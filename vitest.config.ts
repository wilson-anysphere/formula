import { existsSync } from "node:fs";
import { dirname, resolve } from "node:path";
import { defineConfig } from "vitest/config";

function resolveJsToTs() {
  return {
    name: "formula:resolve-js-to-ts",
    enforce: "pre" as const,
    /**
     * Workspace packages use ESM-style `.js` import specifiers in TS source
     * files (e.g. `import './types.js'` next to `types.ts`). TypeScript can
     * resolve these, but Vite/Vitest will not unless we explicitly map.
     *
     * When a relative `.js` import target doesn't exist, fall back to `.ts`/`.tsx`.
     */
    resolveId(source: string, importer?: string) {
      if (!importer) return null;
      if (!source.endsWith(".js")) return null;
      if (!(source.startsWith("./") || source.startsWith("../"))) return null;

      const importerPath = importer.split("?", 1)[0]!;
      const resolved = resolve(dirname(importerPath), source);
      if (existsSync(resolved)) return null;

      const ts = resolved.slice(0, -3) + ".ts";
      if (existsSync(ts)) return ts;

      const tsx = resolved.slice(0, -3) + ".tsx";
      if (existsSync(tsx)) return tsx;

      return null;
    },
  };
}

export default defineConfig({
  plugins: [resolveJsToTs()],
  test: {
    // The repo includes several integration-style suites (API, sandboxed runtimes,
    // wasm-backed rendering) that can exceed Vitest's default 10s hook timeout on
    // shared/contended runners.
    testTimeout: 30_000,
    hookTimeout: 30_000,
    include: [
      "packages/**/*.test.ts",
      "packages/**/*.test.tsx",
      "packages/**/*.vitest.ts",
      "packages/**/*.vitest.tsx",
      "apps/**/*.test.ts",
      "apps/**/*.test.tsx",
      "apps/**/*.vitest.ts",
      "apps/**/*.vitest.tsx",
      "services/api/src/__tests__/**/*.test.ts",
      "services/api/src/__tests__/**/*.vitest.ts",
    ],
    environment: "node",
    setupFiles: ["./vitest.setup.ts"],
    globalSetup: "./scripts/vitest.global-setup.mjs",
  },
});
