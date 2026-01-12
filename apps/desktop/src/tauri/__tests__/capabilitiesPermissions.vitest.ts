import { readFileSync } from "node:fs";
import { fileURLToPath } from "node:url";

import { describe, expect, it } from "vitest";

describe("Tauri capabilities", () => {
  it("grants the updater permissions required by the frontend restart/install flow", () => {
    const capabilityPath = fileURLToPath(new URL("../../../src-tauri/capabilities/main.json", import.meta.url));
    const capability = JSON.parse(readFileSync(capabilityPath, "utf8")) as { permissions?: unknown };

    expect(Array.isArray(capability.permissions)).toBe(true);

    const permissions = capability.permissions as unknown[];

    expect(permissions).toContain("updater:allow-check");
    expect(permissions).toContain("updater:allow-download");
    expect(permissions).toContain("updater:allow-install");
    // Ensure we keep the updater permission surface minimal/explicit.
    expect(permissions).not.toContain("updater:default");
    expect(permissions).not.toContain("updater:allow-download-and-install");
  });
});
