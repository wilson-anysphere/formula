import { readdirSync, readFileSync } from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

import { describe, expect, it } from "vitest";

import { stripComments } from "../../__tests__/sourceTextUtils";

describe("Tauri capabilities", () => {
  function readPermissions(): unknown[] {
    const capabilityPath = fileURLToPath(new URL("../../../src-tauri/capabilities/main.json", import.meta.url));
    const capability = JSON.parse(readFileSync(capabilityPath, "utf8")) as { permissions?: unknown };
    expect(Array.isArray(capability.permissions)).toBe(true);
    return capability.permissions as unknown[];
  }

  it("grants the application allow-invoke permission (IPC commands)", () => {
    const permissions = readPermissions();

    // Application commands (`#[tauri::command]`) are allowlisted via
    // `apps/desktop/src-tauri/permissions/allow-invoke.json` and granted through
    // `capabilities/main.json` as an application permission entry.
    expect(permissions).toContain("allow-invoke");
  });

  it("does not grant unscoped core:allow-invoke (requires explicit allowlist when present)", () => {
    const permissions = readPermissions();

    // The string form would grant the default/unscoped allowlist, which is not acceptable for hardened desktop builds.
    expect(permissions).not.toContain("core:allow-invoke");

    // Some toolchains model command allowlisting exclusively via `permissions/allow-invoke.json`.
    // If `core:allow-invoke` is present, ensure it is scoped and explicit (no allow-all).
    const allowInvoke = permissions.find(
      (permission) => Boolean(permission) && typeof permission === "object" && (permission as any).identifier === "core:allow-invoke",
    ) as any;
    if (!allowInvoke) return;

    const allow = allowInvoke.allow;
    expect(Array.isArray(allow)).toBe(true);
    expect(allow.length).toBeGreaterThan(0);

    const commands = allow
      .map((entry: any) => entry?.command)
      .filter((cmd: any): cmd is string => typeof cmd === "string");
    expect(commands.length).toBe(allow.length);

    // No duplicates.
    expect(new Set(commands).size).toBe(commands.length);

    for (const cmd of commands) {
      expect(cmd.trim()).not.toBe("");
      // Disallow wildcard/pattern scopes; keep commands explicit.
      expect(cmd).not.toContain("*");
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

    for (const cmd of commands) {
      expect(cmd.trim()).not.toBe("");
      expect(cmd).not.toContain("*");
    }
  });

  it("keeps invoke allowlists in sync with frontend invoke() usage", () => {
    const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "../../../../..");

    const allowInvokePath = path.join(root, "apps/desktop/src-tauri/permissions/allow-invoke.json");
    const allowInvokeFile = JSON.parse(readFileSync(allowInvokePath, "utf8")) as any;
    expect(Array.isArray(allowInvokeFile?.permission)).toBe(true);

    const allowInvokeEntry = (allowInvokeFile.permission as any[]).find((p) => p?.identifier === "allow-invoke") as any;
    expect(allowInvokeEntry).toBeTruthy();

    const allow = allowInvokeEntry?.commands?.allow;

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

    // Some toolchains also include an explicit per-command allowlist in the capability itself
    // (`core:allow-invoke`). When present, keep it scoped and in sync with actual invoke() usage.
    const permissions = readPermissions();
    // The string form would grant the default/unscoped allowlist.
    expect(permissions).not.toContain("core:allow-invoke");

    const coreAllowInvoke = permissions.find(
      (permission) => Boolean(permission) && typeof permission === "object" && (permission as any).identifier === "core:allow-invoke",
    ) as any;
    const coreAllow = coreAllowInvoke && Array.isArray(coreAllowInvoke?.allow) ? coreAllowInvoke.allow : [];
    const coreAllowlistedList = coreAllow
      .map((entry: any) => entry?.command)
      .filter((cmd: any): cmd is string => typeof cmd === "string");
    if (coreAllowInvoke) {
      expect(coreAllowlistedList.length).toBe(coreAllow.length);
      // No duplicates.
      expect(new Set(coreAllowlistedList).size).toBe(coreAllowlistedList.length);
    }
    const coreAllowlistedCommands = coreAllowInvoke ? new Set(coreAllowlistedList) : null;

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

    const invokedCommands = new Set<string>();
    // Capture command invocations from helpers like:
    // - invoke("...")
    // - tauriInvoke("...")
    // - invokeFn("...")
    // - args.invoke("...") (the `.invoke` segment still matches)
    const invokeCall = /\b[\w$]*invoke[\w$]*\s*\(\s*(["'])([^"']+)\1/gim;
    // Some code uses indirection (e.g. VBA event macros pass the command name into a helper).
    const runEventMacroCall = /runEventMacro\s*\(\s*[^,]+,\s*(["'])([^"']+)\1/gm;

    // Avoid concatenating all runtime source into one giant string (perf/memory).
    for (const runtimePath of runtimeFiles) {
      const runtimeText = stripComments(readFileSync(runtimePath, "utf8"));

      for (const match of runtimeText.matchAll(invokeCall)) {
        const cmd = match[2].trim();
        // Tauri `#[tauri::command]` names in this repo are snake_case.
        // Filter out unrelated `*invoke*("...")` calls that happen to appear in runtime code.
        if (/^[a-z0-9_]+$/.test(cmd)) {
          invokedCommands.add(cmd);
        }
      }

      // Only `apps/desktop/src/macros/event_macros.ts` uses `runEventMacro(..., "command_name")`.
      if (runtimePath.replace(/\\/g, "/").endsWith("/macros/event_macros.ts")) {
        for (const match of runtimeText.matchAll(runEventMacroCall)) {
          const cmd = match[2].trim();
          if (/^[a-z0-9_]+$/.test(cmd)) {
            invokedCommands.add(cmd);
          }
        }
      }
    }

    for (const cmd of invokedCommands) {
      expect(cmd.trim(), "invoke() command names must not be empty/whitespace").not.toBe("");
      // Keep the command allowlist explicit.
      expect(cmd).not.toContain("*");
    }

    const missingInAllowInvoke = Array.from(invokedCommands)
      .filter((cmd) => !allowlistedCommands.has(cmd))
      .sort();
    expect(
      missingInAllowInvoke,
      `allow-invoke.json is missing commands used by the frontend: ${missingInAllowInvoke.join(", ")}`,
    ).toEqual([]);

    if (coreAllowlistedCommands) {
      const missingInCore = Array.from(invokedCommands)
        .filter((cmd) => !coreAllowlistedCommands.has(cmd))
        .sort();
      expect(
        missingInCore,
        `capabilities/main.json core:allow-invoke is missing commands used by the frontend: ${missingInCore.join(", ")}`,
      ).toEqual([]);

      const extraInCore = Array.from(coreAllowlistedCommands)
        .filter((cmd) => !invokedCommands.has(cmd))
        .sort();
      expect(
        extraInCore,
        `capabilities/main.json core:allow-invoke should match actual invoke() usage; remove unused commands: ${extraInCore.join(", ")}`,
      ).toEqual([]);
    }
  });

  it("grants the dialog permissions required by the frontend", () => {
    const permissions = readPermissions();

    expect(permissions).toContain("dialog:allow-open");
    expect(permissions).toContain("dialog:allow-save");
    expect(permissions).toContain("dialog:allow-confirm");
    expect(permissions).toContain("dialog:allow-message");
    // Keep dialog permission surface minimal.
    expect(permissions).not.toContain("dialog:default");

    // Clipboard operations go through Formula's explicit IPC commands (`clipboard_read` / `clipboard_write`),
    // which are gated by window + stable-origin checks and enforce resource limits during deserialization.
    // Do not grant the clipboard-manager plugin API surface to the webview.
    expect(permissions).not.toContain("clipboard-manager:allow-read-text");
    expect(permissions).not.toContain("clipboard-manager:allow-write-text");
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
