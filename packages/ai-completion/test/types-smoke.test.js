import assert from "node:assert/strict";
import { spawnSync } from "node:child_process";
import { createRequire } from "node:module";
import test from "node:test";
import { fileURLToPath } from "node:url";

const require = createRequire(import.meta.url);
let tscPath = null;
try {
  tscPath = require.resolve("typescript/bin/tsc");
} catch {
  tscPath = null;
}

test(
  "TypeScript can compile ai-completion public API (including suggestRanges maxScanCols)",
  { skip: !tscPath },
  async () => {
    assert.ok(tscPath);
    const fixture = fileURLToPath(new URL("./types.smoke.ts", import.meta.url));

    const result = spawnSync(
      process.execPath,
      [
        tscPath,
        "--noEmit",
        "--pretty",
        "false",
        "--target",
        "ES2022",
        "--lib",
        "ES2022,DOM",
        "--module",
        "ESNext",
        "--moduleResolution",
        "Bundler",
        fixture,
      ],
      { encoding: "utf8" },
    );

    if (result.status !== 0) {
      throw new Error([result.stdout, result.stderr].filter(Boolean).join("\n"));
    }
  },
);

