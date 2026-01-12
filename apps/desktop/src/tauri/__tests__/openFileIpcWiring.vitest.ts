import { readFileSync } from "node:fs";

import { describe, expect, it } from "vitest";

describe("openFileIpc wiring", () => {
  it("installs the open-file IPC handshake in main.ts (prevents cold-start drops)", () => {
    // `main.ts` has a lot of side effects and isn't safe to import in unit tests. Instead,
    // validate (lightly) that it wires the open-file IPC helper responsible for emitting
    // `open-file-ready` *after* the listener is registered.
    const mainUrl = new URL("../../main.ts", import.meta.url);
    const source = readFileSync(mainUrl, "utf8");

    // Ensure the helper is imported.
    expect(source).toMatch(/from\s+["']\.\/tauri\/openFileIpc["']/);

    // Ensure the helper is actually used. This guards against a regression where the helper
    // remains in the tree but the startup wiring is removed.
    expect(source).toMatch(
      /installOpenFileIpc\(\s*\{[\s\S]*?\blisten\b[\s\S]*?\bemit\b[\s\S]*?\bonOpenPath\b[\s\S]*?\}\s*\)/,
    );
  });
});
