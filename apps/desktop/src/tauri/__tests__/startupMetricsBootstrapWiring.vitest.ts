import { readFileSync } from "node:fs";

import { describe, expect, it } from "vitest";

describe("startup metrics bootstrap wiring", () => {
  it("imports startupMetricsBootstrap as the first import in main.ts (minimizes dropped startup events)", () => {
    // `main.ts` has a lot of side effects and isn't safe to import in unit tests.
    // Read the source and assert the bootstrap module is imported first, so the
    // frontend installs its startup timing listeners as early as possible.
    const mainUrl = new URL("../../main.ts", import.meta.url);
    const source = readFileSync(mainUrl, "utf8");

    // Only consider runtime imports: `import type ...` is erased and does not affect module
    // evaluation order in the built JS bundle.
    const firstImport = source.match(/^\s*import(?!\s+type\b)\s+.*$/m)?.[0] ?? "";
    expect(firstImport).toMatch(/import\s+["']\.\/tauri\/startupMetricsBootstrap\.js["']\s*;/);
  });
});
