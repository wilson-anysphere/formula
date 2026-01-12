import { describe, expect, it } from "vitest";

import { readFileSync } from "node:fs";

type CapabilityPermission =
  | string
  | {
      identifier: string;
      allow?: unknown;
      deny?: unknown;
    };

describe("tauri capability event permissions", () => {
  it("scopes event.listen / event.emit (no allow-all event permissions)", () => {
    const capabilityUrl = new URL("../../../src-tauri/capabilities/main.json", import.meta.url);
    const capability = JSON.parse(readFileSync(capabilityUrl, "utf8")) as {
      permissions?: CapabilityPermission[];
    };

    const permissions = Array.isArray(capability.permissions) ? capability.permissions : [];

    // Using the string form grants the permission with its default (unscoped) allowlist. For
    // `event:allow-listen` / `event:allow-emit` that effectively becomes "allow all events", which
    // is not acceptable for hardened desktop builds.
    expect(permissions).not.toContain("event:allow-listen");
    expect(permissions).not.toContain("event:allow-emit");

    const listen = permissions.find(
      (p): p is Exclude<CapabilityPermission, string> =>
        typeof p === "object" && p != null && (p as any).identifier === "event:allow-listen",
    );
    const emit = permissions.find(
      (p): p is Exclude<CapabilityPermission, string> =>
        typeof p === "object" && p != null && (p as any).identifier === "event:allow-emit",
    );

    expect(listen).toBeTruthy();
    expect(emit).toBeTruthy();

    for (const entry of [listen, emit]) {
      const allow = (entry as any).allow;
      expect(Array.isArray(allow)).toBe(true);
      expect((allow as unknown[]).length).toBeGreaterThan(0);
      for (const scope of allow as any[]) {
        expect(scope).toBeTruthy();
        expect(typeof scope.event).toBe("string");
        expect(scope.event.trim()).not.toBe("");
        // Disallow wildcard/pattern scopes; we want explicit event names only.
        expect(scope.event).not.toContain("*");
      }
    }
  });

  it("includes the desktop shell event names used by the frontend", () => {
    const capabilityUrl = new URL("../../../src-tauri/capabilities/main.json", import.meta.url);
    const capability = JSON.parse(readFileSync(capabilityUrl, "utf8")) as {
      permissions?: CapabilityPermission[];
    };

    const permissions = Array.isArray(capability.permissions) ? capability.permissions : [];

    const allowListen = permissions.find(
      (p): p is Exclude<CapabilityPermission, string> =>
        typeof p === "object" && p != null && (p as any).identifier === "event:allow-listen",
    ) as any;
    const allowEmit = permissions.find(
      (p): p is Exclude<CapabilityPermission, string> =>
        typeof p === "object" && p != null && (p as any).identifier === "event:allow-emit",
    ) as any;

    const listenEvents = new Set(
      Array.isArray(allowListen?.allow) ? allowListen.allow.map((entry: any) => entry?.event).filter(Boolean) : [],
    );
    const emitEvents = new Set(
      Array.isArray(allowEmit?.allow) ? allowEmit.allow.map((entry: any) => entry?.event).filter(Boolean) : [],
    );

    // Rust -> JS (frontend listens)
    const requiredListen = [
      // Close flow
      "close-prep",
      "close-requested",

      // File open flows
      "open-file",
      "file-dropped",

      // Tray
      "tray-open",
      "tray-new",
      "tray-quit",

      // Shortcuts
      "shortcut-quick-open",
      "shortcut-command-palette",

      // Menu bar
      "menu-open",
      "menu-new",
      "menu-save",
      "menu-save-as",
      "menu-close-window",
      "menu-quit",
      "menu-undo",
      "menu-redo",
      "menu-cut",
      "menu-copy",
      "menu-paste",
      "menu-select-all",
      "menu-about",
      "menu-check-updates",

      // Updater
      "update-check-started",
      "update-check-already-running",
      "update-not-available",
      "update-check-error",
      "update-available",

      // Startup instrumentation
      "startup:window-visible",
      "startup:webview-loaded",
      "startup:tti",
      "startup:metrics",

      // Deep links
      "oauth-redirect",
    ];

    for (const event of requiredListen) {
      expect(listenEvents.has(event)).toBe(true);
    }

    // JS -> Rust (frontend emits)
    const requiredEmit = [
      "close-prep-done",
      "close-handled",
      "open-file-ready",
      "updater-ui-ready",
      // Emitted by the COI check harness (window.eval) to report results back to Rust.
      "coi-check-result",
    ];

    for (const event of requiredEmit) {
      expect(emitEvents.has(event)).toBe(true);
    }

    // Sanity check: events outside the allowlist should be denied by Tauri's permission system.
    // (We can't assert the runtime error message here without running the desktop shell, but we
    // can assert the capability file does not include arbitrary names.)
    expect(listenEvents.has("totally-not-a-real-event")).toBe(false);
    expect(emitEvents.has("totally-not-a-real-event")).toBe(false);
  });
});
