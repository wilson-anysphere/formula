import { readdirSync, readFileSync } from "node:fs";
import path from "node:path";
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

  it("keeps src-tauri/permissions/allow-invoke.json in sync with frontend invoke() usage", () => {
    const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "../../../../..");

    const allowInvokePath = path.join(root, "apps/desktop/src-tauri/permissions/allow-invoke.json");
    const allowInvoke = JSON.parse(readFileSync(allowInvokePath, "utf8")) as any;
    const allow = allowInvoke?.permission?.[0]?.commands?.allow;

    expect(Array.isArray(allow)).toBe(true);
    expect(allow.length).toBeGreaterThan(0);
    expect(allow.every((cmd: any) => typeof cmd === "string")).toBe(true);

    const allowlistedCommands = new Set<string>(allow);
    // No duplicates.
    expect(allowlistedCommands.size).toBe(allow.length);

    for (const cmd of allowlistedCommands) {
      expect(cmd.trim()).not.toBe("");
      // Keep the command allowlist explicit.
      expect(cmd).not.toContain("*");
    }

    const isRuntimeSource = (filePath: string): boolean => {
      const normalized = filePath.replace(/\\/g, "/");
      if (normalized.includes("/__tests__/")) return false;
      if (normalized.match(/\.(test|spec|vitest)\./)) return false;
      if (normalized.endsWith(".md")) return false;
      if (normalized.includes("/src-tauri/capabilities/")) return false;
      const ext = path.extname(filePath);
      return ext === ".ts" || ext === ".tsx" || ext === ".js" || ext === ".jsx" || ext === ".mjs";
    };

    const walk = (dir: string): string[] => {
      const out: string[] = [];
      for (const entry of readdirSync(dir, { withFileTypes: true })) {
        const full = path.join(dir, entry.name);
        if (entry.isDirectory()) {
          if (entry.name === "node_modules" || entry.name === "__tests__") continue;
          out.push(...walk(full));
        } else if (entry.isFile()) {
          if (isRuntimeSource(full)) out.push(full);
        }
      }
      return out;
    };

    const runtimeFiles = [
      ...walk(path.join(root, "apps/desktop/src")),
      ...walk(path.join(root, "packages/extension-marketplace/src")),
      ...walk(path.join(root, "packages/extension-host/src/browser")),
      ...walk(path.join(root, "shared/extension-package")),
    ].sort();

    const runtimeText = runtimeFiles.map((p) => readFileSync(p, "utf8")).join("\n");

    const invokedCommands = new Set<string>();
    // Capture command invocations from helpers like:
    // - invoke("...")
    // - tauriInvoke("...")
    // - invokeFn("...")
    // - args.invoke("...") (the `.invoke` segment still matches)
    const invokeCall = /\b[\w$]*invoke[\w$]*\s*\(\s*(["'])([^"']+)\1/gim;
    for (const match of runtimeText.matchAll(invokeCall)) {
      invokedCommands.add(match[2]);
    }

    // Some code uses indirection (e.g. VBA event macros pass the command name into a helper).
    {
      const eventMacrosPath = path.join(root, "apps/desktop/src/macros/event_macros.ts");
      const eventMacrosText = readFileSync(eventMacrosPath, "utf8");
      const runEventMacroCall = /runEventMacro\s*\(\s*[^,]+,\s*(["'])([^"']+)\1/gm;
      for (const match of eventMacrosText.matchAll(runEventMacroCall)) {
        invokedCommands.add(match[2]);
      }
    }

    for (const cmd of invokedCommands) {
      expect(allowlistedCommands.has(cmd)).toBe(true);
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
    // Guard against legacy/unprefixed identifiers too (schema drift).
    expect(permissions).not.toContain("clipboard:default");
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
