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

  it("uses the application allow-invoke permission (and does not grant unscoped core:allow-invoke)", () => {
    const permissions = readPermissions();

    // Application commands (`#[tauri::command]`) are allowlisted via `src-tauri/permissions/allow-invoke.json`
    // and granted through `capabilities/main.json` as an application permission entry.
    expect(permissions).toContain("allow-invoke");

    // Some Tauri toolchains expose a `core:allow-invoke` permission for per-command allowlisting.
    //
    // If present, it must be scoped via the object form (explicit allowlist). The plain string
    // form would grant the default/unscoped allowlist, which is not acceptable for hardened
    // desktop builds.
    expect(permissions).not.toContain("core:allow-invoke");
    const coreAllowInvoke = permissions.find(
      (permission): permission is { identifier: string; allow?: unknown } =>
        Boolean(permission) && typeof permission === "object" && (permission as any).identifier === "core:allow-invoke",
    );

    if (coreAllowInvoke) {
      const allow = (coreAllowInvoke as any).allow;
      expect(Array.isArray(allow)).toBe(true);
      expect((allow as unknown[]).length).toBeGreaterThan(0);

      const commands = (allow as any[])
        .map((scope) => scope?.command)
        .filter((name): name is string => typeof name === "string");
      expect(new Set(commands).size).toBe(commands.length);

      for (const cmd of commands) {
        expect(cmd.trim()).not.toBe("");
        // Disallow wildcard/pattern scopes; we want explicit command names only.
        expect(cmd).not.toContain("*");
      }
    }
  });

  it("defines an explicit invoke command allowlist (no wildcards)", () => {
    const permissionPath = fileURLToPath(new URL("../../../src-tauri/permissions/allow-invoke.json", import.meta.url));
    const permissionFile = JSON.parse(readFileSync(permissionPath, "utf8")) as any;
    expect(Array.isArray(permissionFile?.permission)).toBe(true);

    const entry = (permissionFile.permission as any[]).find((p) => p?.identifier === "allow-invoke") as any;
    expect(entry).toBeTruthy();

    const allow = entry?.commands?.allow as unknown;
    expect(Array.isArray(allow)).toBe(true);
    expect((allow as unknown[]).length).toBeGreaterThan(0);

    const commands = (allow as unknown[]).filter((command): command is string => typeof command === "string");
    expect(commands.length).toBe((allow as unknown[]).length);
    expect(new Set(commands).size).toBe(commands.length);

    for (const command of commands) {
      expect(command.trim()).not.toBe("");
      expect(command).not.toContain("*");
    }
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
