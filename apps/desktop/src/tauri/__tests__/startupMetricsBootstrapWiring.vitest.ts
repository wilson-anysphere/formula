import { readFileSync } from "node:fs";

import { describe, expect, it } from "vitest";

describe("startup metrics bootstrap wiring", () => {
  it("imports startupMetricsBootstrap as the first import in main.ts (minimizes metric skew)", () => {
    // `main.ts` has a lot of side effects and isn't safe to import in unit tests.
    // Read the source and assert the bootstrap module is imported first, so the
    // host-side startup timing capture runs before the rest of the module graph.
    const mainUrl = new URL("../../main.ts", import.meta.url);
    const source = readFileSync(mainUrl, "utf8");

    const firstImport = source.match(/^\s*import\s+.*$/m)?.[0] ?? "";
    expect(firstImport).toMatch(/import\s+["']\.\/tauri\/startupMetricsBootstrap\.js["']\s*;/);
  });
});

