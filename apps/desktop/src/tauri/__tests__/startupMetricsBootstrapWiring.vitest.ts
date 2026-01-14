import { readFileSync } from "node:fs";

import { describe, expect, it } from "vitest";

describe("startup metrics bootstrap wiring", () => {
  it("loads startupMetricsBootstrap before main.ts in index.html (allows JS to report before the full app graph loads)", () => {
    const indexUrl = new URL("../../../index.html", import.meta.url);
    const source = readFileSync(indexUrl, "utf8");

    const bootstrapIdx = source.indexOf('src="/src/tauri/startupMetricsBootstrap.ts"');
    const mainIdx = source.indexOf('src="/src/main.ts"');
    expect(bootstrapIdx).toBeGreaterThanOrEqual(0);
    expect(mainIdx).toBeGreaterThanOrEqual(0);
    expect(bootstrapIdx).toBeLessThan(mainIdx);
  });

  it("imports startupMetricsBootstrap as the first runtime import in main.ts (fallback guardrail)", () => {
    // `main.ts` has a lot of side effects and isn't safe to import in unit tests.
    // Read the source and assert the bootstrap module is imported first, so the
    // host-side reporting also happens early if `main.ts` is used as an entrypoint.
    const mainUrl = new URL("../../main.ts", import.meta.url);
    const source = readFileSync(mainUrl, "utf8");

    // Only consider runtime imports: `import type ...` is erased and does not affect module
    // evaluation order in the built JS bundle.
    const firstImport = source.match(/^\s*import(?!\s+type\b)\s+.*$/m)?.[0] ?? "";
    expect(firstImport).toMatch(/import\s+["']\.\/tauri\/startupMetricsBootstrap\.js["']\s*;/);
  });
});
