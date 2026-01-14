import { existsSync, readFileSync } from "node:fs";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";

import { describe, expect, it } from "vitest";

import { stripComments } from "../../../../apps/desktop/test/sourceTextUtils.js";

describe("@formula/workbook-backend index exports", () => {
  it("avoids .ts import specifiers (breaks repo-level typecheck)", () => {
    const testDir = dirname(fileURLToPath(import.meta.url));
    const indexPath = join(testDir, "..", "index.ts");
    // Strip JS/TS comments so commented-out imports cannot satisfy or fail assertions.
    const src = stripComments(readFileSync(indexPath, "utf8"));

    // Repo-wide `pnpm -w typecheck` uses a TS config that does not enable
    // `allowImportingTsExtensions`, so `.ts` specifiers in source imports/exports
    // break the build.
    const tsSpecifierRe =
      /(?:\bfrom\s+|\bimport\s*\(\s*|\bimport\s+)\s*['"]\.\.?\/[^'"\n]+?\.(?:ts|tsx)(?:[?#][^'"\n]*)?['"]/;
    expect(src).not.toMatch(tsSpecifierRe);

    // Keep Node ESM runnable by ensuring any `.js` specifier we use is backed by
    // a real `.js` file on disk.
    expect(src).toMatch(/from\s+['"]\.\/sheetNameValidation\.js['"]/);
    expect(existsSync(join(testDir, "..", "sheetNameValidation.js"))).toBe(true);
  });
});
