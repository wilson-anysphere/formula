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

  it("scopes custom Rust command invocation via core:allow-invoke", () => {
    const permissions = readPermissions();
    // Use the object form so we can keep the command allowlist explicit (no allow-all).
    expect(permissions).not.toContain("core:allow-invoke");

    const allowInvoke = permissions.find(
      (permission) => Boolean(permission) && typeof permission === "object" && (permission as any).identifier === "core:allow-invoke",
    );
    expect(allowInvoke).toBeTruthy();

    const allow = (allowInvoke as any).allow;
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

  it("grants the dialog + clipboard permissions required by the frontend", () => {
    const permissions = readPermissions();

    expect(permissions).toContain("dialog:allow-open");
    expect(permissions).toContain("dialog:allow-save");
    expect(permissions).toContain("dialog:allow-confirm");
    expect(permissions).toContain("dialog:allow-message");
    // Keep dialog permission surface minimal.
    expect(permissions).not.toContain("dialog:default");

    // Clipboard permission identifiers differ slightly between toolchains:
    // - `clipboard:*` (older Tauri/plugin naming)
    // - `clipboard-manager:*` (tauri-plugin-clipboard-manager)
    expect(
      permissions.includes("clipboard:allow-read-text") || permissions.includes("clipboard-manager:allow-read-text"),
    ).toBe(true);
    expect(
      permissions.includes("clipboard:allow-write-text") || permissions.includes("clipboard-manager:allow-write-text"),
    ).toBe(true);
    // Keep clipboard permission surface minimal (no broad defaults).
    expect(permissions).not.toContain("clipboard:default");
    expect(permissions).not.toContain("clipboard-manager:default");
  });

  it("grants the window permissions required by the UI window helpers", () => {
    const permissions = readPermissions();
    const hasWindowPerm = (name: string): boolean =>
      permissions.includes(name) || permissions.includes(`core:${name}`);

    expect(hasWindowPerm("window:allow-hide")).toBe(true);
    expect(hasWindowPerm("window:allow-show")).toBe(true);
    expect(hasWindowPerm("window:allow-set-focus")).toBe(true);
    expect(hasWindowPerm("window:allow-close")).toBe(true);
    expect(hasWindowPerm("window:allow-minimize")).toBe(true);
    expect(hasWindowPerm("window:allow-toggle-maximize")).toBe(true);
    expect(hasWindowPerm("window:allow-is-maximized")).toBe(true);

    // Keep window permission surface minimal.
    expect(permissions).not.toContain("window:default");
    expect(permissions).not.toContain("core:window:default");
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

  it("keeps core:allow-invoke in sync with the frontend's invoke() usage", () => {
    const permissions = readPermissions();

    const allowInvoke = permissions.find(
      (permission) => Boolean(permission) && typeof permission === "object" && (permission as any).identifier === "core:allow-invoke",
    ) as any;
    expect(allowInvoke).toBeTruthy();

    const allow = Array.isArray(allowInvoke?.allow) ? allowInvoke.allow : [];
    const allowlistedCommands = new Set(
      allow.map((entry: any) => entry?.command).filter((cmd: any): cmd is string => typeof cmd === "string"),
    );

    const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "../../../../..");

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

    // Capture command invocations from helpers like:
    // - invoke("...")
    // - tauriInvoke("...")
    // - invokeFn("...")
    // - args.invoke("...") (the `.invoke` segment still matches)
    const invokedCommands = new Set<string>();
    const invokeCall = /\b[\w$]*invoke[\w$]*\s*\(\s*(["'])([^"']+)\1/gim;
    for (const match of runtimeText.matchAll(invokeCall)) {
      invokedCommands.add(match[2]);
    }

    // Some code uses indirection (e.g. VBA event macros pass the command name into a helper).
    // Keep those allowlisted too.
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

    // Keep the allowlist as small as possible: it should match actual invoke usage.
    expect(Array.from(allowlistedCommands).sort()).toEqual(Array.from(invokedCommands).sort());
  });
});

