import assert from "node:assert/strict";
import fs from "node:fs";
import test from "node:test";
import { fileURLToPath } from "node:url";

import { stripComments } from "../../../apps/desktop/test/sourceTextUtils.js";

test("power-query: node.d.ts avoids .d.ts specifiers and keeps referenced types in-scope", () => {
  const nodeDtsPath = fileURLToPath(new URL("../src/node.d.ts", import.meta.url));
  const content = stripComments(fs.readFileSync(nodeDtsPath, "utf8"));

  // TypeScript's `--moduleResolution Bundler` can be sensitive to explicit `.d.ts` specifiers.
  // Ensure we only reference the runtime `.js` entrypoint, letting TS follow it to the
  // corresponding declaration output.
  assert.equal(
    /(?:import|export)\s+[^;]*from\s+["'][^"']*\.d\.ts["']/.test(content),
    false,
    "Expected packages/power-query/src/node.d.ts to avoid importing/re-exporting .d.ts specifiers",
  );

  // Ensure the node surface re-exports the main entrypoint types.
  assert.match(content, /export\s+\*\s+from\s+["']\.\/index\.js["'];/);

  // Ensure types referenced by declarations in this file are imported into scope (not just re-exported).
  for (const name of ["CacheStore", "CacheEntry", "CacheCryptoProvider", "CredentialStore"]) {
    assert.match(
      content,
      new RegExp(`import\\s+type\\s+\\{[^}]*\\b${name}\\b[^}]*\\}\\s+from\\s+["']\\.\\/index\\.js["']`),
      `Expected node.d.ts to import type ${name} from ./index.js`,
    );
  }
});
