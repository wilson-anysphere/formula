import { existsSync, readFileSync } from "node:fs";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";

import { describe, expect, it } from "vitest";

describe("@formula/workbook-backend index exports", () => {
  it("avoids .ts import specifiers (breaks repo-level typecheck)", () => {
    const testDir = dirname(fileURLToPath(import.meta.url));
    const indexPath = join(testDir, "..", "index.ts");
    const src = readFileSync(indexPath, "utf8");

    // Repo-wide `pnpm -w typecheck` uses a TS config that does not enable
    // `allowImportingTsExtensions`, so `.ts` specifiers in source imports/exports
    // break the build.
    expect(src).not.toMatch(/from\s+['"]\.\.?\/[^'"\n]+\.ts['"]/);

    // Keep Node ESM runnable by ensuring any `.js` specifier we use is backed by
    // a real `.js` file on disk.
    expect(src).toMatch(/from\s+['"]\.\/sheetNameValidation\.js['"]/);
    expect(existsSync(join(testDir, "..", "sheetNameValidation.js"))).toBe(true);
  });
});
