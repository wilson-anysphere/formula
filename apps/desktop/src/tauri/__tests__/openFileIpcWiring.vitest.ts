import { readFileSync } from "node:fs";

import { describe, expect, it } from "vitest";

import { stripComments } from "../../__tests__/sourceTextUtils";

describe("openFileIpc wiring", () => {
  it("installs the open-file IPC handshake in main.ts (prevents cold-start drops)", () => {
    // `main.ts` has a lot of side effects and isn't safe to import in unit tests. Instead,
    // validate (lightly) that it wires the open-file IPC helper responsible for emitting
    // `open-file-ready` *after* the listener is registered.
    const mainUrl = new URL("../../main.ts", import.meta.url);
    const source = readFileSync(mainUrl, "utf8");
    const code = stripComments(source);

    // Ensure the helper is imported (as an actual import statement, not just mentioned in a comment).
    // Allow optional `.ts`/`.js` extensions so this guardrail doesn't break on harmless specifier refactors.
    expect(code).toMatch(
      /^\s*import\s+\{[^}]*\binstallOpenFileIpc\b[^}]*\}\s+from\s+["']\.\/tauri\/openFileIpc(?:\.(?:ts|js))?["']/m,
    );

    // Ensure the helper is actually used. This guards against a regression where the helper
    // remains in the tree but the startup wiring is removed.
    expect(code).toMatch(/^\s*(?:void\s+)?installOpenFileIpc\(\s*\{[\s\S]*?\blisten\b[\s\S]*?\bemit\b[\s\S]*?\bonOpenPath\b/m);
  });
});
