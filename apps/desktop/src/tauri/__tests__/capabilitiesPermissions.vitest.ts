import { readFileSync } from "node:fs";
import { fileURLToPath } from "node:url";

import { describe, expect, it } from "vitest";

describe("Tauri capabilities", () => {
  function readPermissions(): unknown[] {
    const capabilityPath = fileURLToPath(new URL("../../../src-tauri/capabilities/main.json", import.meta.url));
    const capability = JSON.parse(readFileSync(capabilityPath, "utf8")) as { permissions?: unknown };
    expect(Array.isArray(capability.permissions)).toBe(true);
    return capability.permissions as unknown[];
  }

  it("does not rely on core:allow-invoke permissions (commands must validate in Rust)", () => {
    const permissions = readPermissions();

    // Some Tauri toolchains expose a `core:allow-invoke` permission for per-command allowlisting,
    // but the current config/schema used in this repo does not.
    const hasAllowInvoke = permissions.some(
      (permission) =>
        permission === "core:allow-invoke" ||
        (Boolean(permission) && typeof permission === "object" && (permission as any).identifier === "core:allow-invoke"),
    );
    expect(hasAllowInvoke).toBe(false);
  });

  it("grants the dialog + clipboard permissions required by the frontend", () => {
    const permissions = readPermissions();

    expect(permissions).toContain("dialog:allow-open");
    expect(permissions).toContain("dialog:allow-save");
    expect(permissions).toContain("dialog:allow-confirm");
    expect(permissions).toContain("dialog:allow-message");
    // Keep dialog permission surface minimal.
    expect(permissions).not.toContain("dialog:default");

    expect(permissions).toContain("clipboard-manager:allow-read-text");
    expect(permissions).toContain("clipboard-manager:allow-write-text");
    // Keep clipboard permission surface minimal.
    expect(permissions).not.toContain("clipboard-manager:default");
  });

  it("grants the window permissions required by the UI window helpers", () => {
    const permissions = readPermissions();
    expect(permissions).toContain("core:window:allow-hide");
    expect(permissions).toContain("core:window:allow-show");
    expect(permissions).toContain("core:window:allow-set-focus");
    expect(permissions).toContain("core:window:allow-close");
    expect(permissions).toContain("core:window:allow-minimize");
    expect(permissions).toContain("core:window:allow-toggle-maximize");
    expect(permissions).toContain("core:window:allow-maximize");
    expect(permissions).toContain("core:window:allow-unmaximize");
    expect(permissions).toContain("core:window:allow-is-maximized");

    // Keep window permission surface minimal.
    expect(permissions).not.toContain("core:window:default");
    expect(permissions).not.toContain("window:default");
  });

  it("does not grant shell open permissions to the frontend (external navigation goes through Rust)", () => {
    const permissions = readPermissions();
    expect(permissions).not.toContain("shell:allow-open");
    expect(permissions).not.toContain("shell:default");
  });

  it("grants the updater permissions required by the frontend restart/install flow", () => {
    const permissions = readPermissions();

    expect(permissions).toContain("updater:allow-check");
    expect(permissions).toContain("updater:allow-download");
    expect(permissions).toContain("updater:allow-install");
    // Ensure we keep the updater permission surface minimal/explicit.
    expect(permissions).not.toContain("updater:default");
    expect(permissions).not.toContain("updater:allow-download-and-install");
  });

  it("does not grant notification plugin permissions to the webview (notifications go through show_system_notification)", () => {
    const permissions = readPermissions();

    const identifiers = permissions
      .map((permission) => {
        if (typeof permission === "string") return permission;
        if (permission && typeof permission === "object") {
          const identifier = (permission as Record<string, unknown>).identifier;
          if (typeof identifier === "string") return identifier;
        }
        return null;
      })
      .filter((permission): permission is string => typeof permission === "string");

    expect(identifiers.some((permission) => permission.startsWith("notification:"))).toBe(false);
    // Some Tauri versions may namespace this as a core permission.
    expect(identifiers.some((permission) => permission.startsWith("core:notification:"))).toBe(false);
  });
});
