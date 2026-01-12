import { readFileSync } from "node:fs";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";

import { describe, expect, it } from "vitest";

describe("@formula/workbook-backend index exports", () => {
  it("keeps Node ESM import specifiers explicit for TS entrypoints", () => {
    const testDir = dirname(fileURLToPath(import.meta.url));
    const indexPath = join(testDir, "..", "index.ts");
    const src = readFileSync(indexPath, "utf8");

    // `@formula/workbook-backend` exports TypeScript source directly
    // (package.json `exports` points at `./src/index.ts`). When executing these
    // sources under Node ESM (e.g. via `--experimental-strip-types`), Node does
    // *not* apply TypeScript-style extension resolution for relative imports.
    //
    // Keep internal re-exports using explicit `.ts` specifiers so the package is
    // directly runnable in Node ESM without a build step.
    expect(src).toMatch(/from\s+\"\.\/sheetNameValidation\.ts\"/);
  });
});
