import { describe, expect, it } from "vitest";

import { readdirSync, readFileSync } from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

import { stripComments, stripRustComments } from "../../__tests__/sourceTextUtils";

type CapabilityPermission =
  | string
  | {
      identifier: string;
      allow?: unknown;
      deny?: unknown;
    };

describe("tauri capability event permissions", () => {
  const allowListenIdentifiers = ["core:event:allow-listen"] as const;
  const allowEmitIdentifiers = ["core:event:allow-emit"] as const;

  it("is scoped to the main window label via the capability file", () => {
    const tauriConfUrl = new URL("../../../src-tauri/tauri.conf.json", import.meta.url);
    const tauriConf = JSON.parse(readFileSync(tauriConfUrl, "utf8")) as any;

    const windows = Array.isArray(tauriConf?.app?.windows) ? tauriConf.app.windows : [];
    const mainWindow = windows.find((w: any) => w?.label === "main");
    expect(mainWindow).toBeTruthy();

    const mainWindowLabel = String(mainWindow?.label ?? "");
    expect(mainWindowLabel).toBe("main");

    const capabilityUrl = new URL("../../../src-tauri/capabilities/main.json", import.meta.url);
    const capability = JSON.parse(readFileSync(capabilityUrl, "utf8")) as any;
    expect(capability?.identifier).toBe("main");
    expect(Array.isArray(capability?.windows)).toBe(true);
    expect(capability.windows).toContain(mainWindowLabel);
  });

  it("scopes event.listen / event.emit (no allow-all event permissions)", () => {
    const capabilityUrl = new URL("../../../src-tauri/capabilities/main.json", import.meta.url);
    const capability = JSON.parse(readFileSync(capabilityUrl, "utf8")) as {
      permissions?: CapabilityPermission[];
    };

    const permissions = Array.isArray(capability.permissions) ? capability.permissions : [];

    // Using the string form grants the permission with its default (unscoped) allowlist. For
    // `core:event:allow-listen` / `core:event:allow-emit` that effectively becomes "allow all
    // events", which is not acceptable for hardened desktop builds.
    for (const id of allowListenIdentifiers) expect(permissions).not.toContain(id);
    for (const id of allowEmitIdentifiers) expect(permissions).not.toContain(id);

    const listen = permissions.find(
      (p): p is Exclude<CapabilityPermission, string> =>
        typeof p === "object" && p != null && allowListenIdentifiers.includes((p as any).identifier),
    );
    const emit = permissions.find(
      (p): p is Exclude<CapabilityPermission, string> =>
        typeof p === "object" && p != null && allowEmitIdentifiers.includes((p as any).identifier),
    );

    expect(listen).toBeTruthy();
    expect(emit).toBeTruthy();

    for (const entry of [listen, emit]) {
      const allow = (entry as any).allow;
      expect(Array.isArray(allow)).toBe(true);
      expect((allow as unknown[]).length).toBeGreaterThan(0);

      const rawEvents = (allow as any[]).map((scope) => scope?.event).filter((name) => typeof name === "string");
      expect(new Set(rawEvents).size).toBe(rawEvents.length);

      for (const scope of allow as any[]) {
        expect(scope).toBeTruthy();
        expect(typeof scope.event).toBe("string");
        expect(scope.event.trim()).not.toBe("");
        // Disallow wildcard/pattern scopes; we want explicit event names only.
        expect(scope.event).not.toContain("*");
      }
    }
  });

  it("does not grant broad default core permissions (least privilege)", () => {
    const capabilityUrl = new URL("../../../src-tauri/capabilities/main.json", import.meta.url);
    const capability = JSON.parse(readFileSync(capabilityUrl, "utf8")) as {
      permissions?: CapabilityPermission[];
    };

    const permissions = Array.isArray(capability.permissions) ? capability.permissions : [];

    // `core:default` would implicitly grant broad access to many core plugins (event/window/etc),
    // defeating the point of the explicit allowlists in this hardened build.
    const hasPermission = (id: string): boolean =>
      permissions.some(
        (p) =>
          (typeof p === "string" && p === id) ||
          (typeof p === "object" && p != null && (p as any).identifier === id),
      );

    expect(hasPermission("core:default")).toBe(false);
    expect(hasPermission("core:event:default")).toBe(false);
    expect(hasPermission("core:window:default")).toBe(false);
    // Guard against legacy/unprefixed identifiers too (schema drift).
    expect(hasPermission("event:default")).toBe(false);
    expect(hasPermission("window:default")).toBe(false);
  });

  it("includes the desktop shell event names used by the frontend", () => {
    const capabilityUrl = new URL("../../../src-tauri/capabilities/main.json", import.meta.url);
    const capability = JSON.parse(readFileSync(capabilityUrl, "utf8")) as {
      permissions?: CapabilityPermission[];
    };

    const permissions = Array.isArray(capability.permissions) ? capability.permissions : [];

    const allowListen = permissions.find(
      (p): p is Exclude<CapabilityPermission, string> =>
        typeof p === "object" && p != null && allowListenIdentifiers.includes((p as any).identifier),
    ) as any;
    const allowEmit = permissions.find(
      (p): p is Exclude<CapabilityPermission, string> =>
        typeof p === "object" && p != null && allowEmitIdentifiers.includes((p as any).identifier),
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
      "menu-print",
      "menu-print-preview",
      "menu-export-pdf",
      "menu-close-window",
      "menu-quit",
      "menu-undo",
      "menu-redo",
      "menu-cut",
      "menu-copy",
      "menu-paste",
      "menu-paste-special",
      "menu-select-all",
      "menu-zoom-in",
      "menu-zoom-out",
      "menu-zoom-reset",
      "menu-about",
      "menu-check-updates",
      "menu-open-release-page",

      // Updater
      "update-check-started",
      "update-check-already-running",
      "update-not-available",
      "update-check-error",
      "update-available",
      "update-download-started",
      "update-download-progress",
      "update-downloaded",
      "update-download-error",

      // Pyodide (Python runtime download/install)
      "pyodide-download-progress",

      // Startup instrumentation
      "startup:window-visible",
      "startup:webview-loaded",
      "startup:first-render",
      "startup:tti",
      "startup:metrics",

      // Pyodide asset download
      "pyodide-download-progress",

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
      "oauth-redirect-ready",
      "updater-ui-ready",
      // Emitted by the COI check harness (window.eval) to report results back to Rust.
      "coi-check-result",
    ];

    for (const event of requiredEmit) {
      expect(emitEvents.has(event)).toBe(true);
    }

    // Keep the allowlists as small as possible: if you add a new desktop event, update both
    // `src-tauri/capabilities/main.json` and this test. This prevents accidental over-broad
    // capability grants (e.g. "just allow one more event" turning into allow-all).
    expect(Array.from(listenEvents).sort()).toEqual(Array.from(new Set(requiredListen)).sort());
    expect(Array.from(emitEvents).sort()).toEqual(Array.from(new Set(requiredEmit)).sort());

    // Sanity check: events outside the allowlist should be denied by Tauri's permission system.
    // (We can't assert the runtime error message here without running the desktop shell, but we
    // can assert the capability file does not include arbitrary names.)
    expect(listenEvents.has("totally-not-a-real-event")).toBe(false);
    expect(emitEvents.has("totally-not-a-real-event")).toBe(false);
  });

  it("does not grant event permissions for unused event names", () => {
    const capabilityUrl = new URL("../../../src-tauri/capabilities/main.json", import.meta.url);
    const capability = JSON.parse(readFileSync(capabilityUrl, "utf8")) as {
      permissions?: CapabilityPermission[];
    };

    const permissions = Array.isArray(capability.permissions) ? capability.permissions : [];

    const allowListen = permissions.find(
      (p): p is Exclude<CapabilityPermission, string> =>
        typeof p === "object" && p != null && allowListenIdentifiers.includes((p as any).identifier),
    ) as any;
    const allowEmit = permissions.find(
      (p): p is Exclude<CapabilityPermission, string> =>
        typeof p === "object" && p != null && allowEmitIdentifiers.includes((p as any).identifier),
    ) as any;

    const allowlistedEvents = [
      ...(Array.isArray(allowListen?.allow) ? allowListen.allow.map((entry: any) => entry?.event).filter(Boolean) : []),
      ...(Array.isArray(allowEmit?.allow) ? allowEmit.allow.map((entry: any) => entry?.event).filter(Boolean) : []),
    ];

    const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "../../../../..");

    const isRuntimeSource = (filePath: string): boolean => {
      const normalized = filePath.replace(/\\/g, "/");
      if (normalized.includes("/__tests__/")) return false;
      if (normalized.match(/\\.(test|spec|vitest)\\./)) return false;
      if (normalized.endsWith(".md")) return false;
      if (normalized.includes("/src-tauri/capabilities/")) return false;
      const ext = path.extname(filePath);
      return ext === ".ts" || ext === ".tsx" || ext === ".js" || ext === ".jsx" || ext === ".rs";
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
      ...walk(path.join(root, "apps/desktop/src-tauri/src")),
    ];

    const runtimeText = runtimeFiles
      .map((p) => {
        const raw = readFileSync(p, "utf8");
        const ext = path.extname(p);
        // Strip comments so commented-out `listen("...")` / `emit("...")` calls and event names
        // cannot satisfy or fail these guardrail assertions.
        return ext === ".rs" ? stripRustComments(raw) : stripComments(raw);
      })
      .join("\n");

    // Ensure the frontend's *direct* `listen("...")` / `emit("...")` calls are all allowlisted.
    // (Some subsystems use indirection, e.g. iterating over an array of updater event names; those
    // are covered by the explicit allowlist checks above.)
    const listenedByFrontend = new Set<string>();
    const emittedByFrontend = new Set<string>();

    const listenCall = /(^|[^.#A-Za-z0-9_])listen\s*\(\s*(["'])([^"']+)\2/gm;
    for (const match of runtimeText.matchAll(listenCall)) {
      listenedByFrontend.add(match[3]);
    }

    const emitCall = /(^|[^.#A-Za-z0-9_])emit\s*\(\s*(["'])([^"']+)\2/gm;
    for (const match of runtimeText.matchAll(emitCall)) {
      emittedByFrontend.add(match[3]);
    }

    const allowListenEvents = new Set(
      Array.isArray(allowListen?.allow) ? allowListen.allow.map((entry: any) => entry?.event).filter(Boolean) : [],
    );
    const allowEmitEvents = new Set(
      Array.isArray(allowEmit?.allow) ? allowEmit.allow.map((entry: any) => entry?.event).filter(Boolean) : [],
    );

    // Rust-side event usage:
    // - `.emit(...)` => must be allowlisted for the frontend to listen
    // - `.listen(...)` / `.listen_global(...)` => must be allowlisted for the frontend to emit
    {
      const rustFiles = runtimeFiles
        .filter((filePath) => path.extname(filePath) === ".rs")
        // Ensure deterministic behavior if the filesystem ordering changes.
        .sort();

      // Some Rust modules use overlapping constant names (e.g. both `tray.rs` and `menu.rs` define
      // `ITEM_OPEN`). To avoid resolving those incorrectly, we:
      // - resolve constants within their declaring file when possible
      // - only fall back to a global constant map when the name is unambiguous (defined exactly once)
      const globalConstDefs = new Map<string, Set<string>>();
      for (const rustPath of rustFiles) {
        const rustText = stripRustComments(readFileSync(rustPath, "utf8"));
        for (const match of rustText.matchAll(/\b(?:pub\s+)?const\s+([A-Z0-9_]+)\s*:\s*&str\s*=\s*"([^"]+)"/g)) {
          const name = match[1];
          const value = match[2];
          const set = globalConstDefs.get(name) ?? new Set<string>();
          set.add(value);
          globalConstDefs.set(name, set);
        }
      }
      const globalUniqueConsts = new Map<string, string>();
      for (const [name, values] of globalConstDefs) {
        if (values.size === 1) {
          globalUniqueConsts.set(name, Array.from(values)[0]);
        }
      }

      const rustEmits = new Set<string>();
      const rustListens = new Set<string>();

      for (const rustPath of rustFiles) {
        const rustText = stripRustComments(readFileSync(rustPath, "utf8"));

        const localConsts = new Map<string, string>();
        for (const match of rustText.matchAll(/\b(?:pub\s+)?const\s+([A-Z0-9_]+)\s*:\s*&str\s*=\s*"([^"]+)"/g)) {
          localConsts.set(match[1], match[2]);
        }

        const resolveConst = (name: string): string | null => {
          return localConsts.get(name) ?? globalUniqueConsts.get(name) ?? null;
        };

        for (const match of rustText.matchAll(/\.emit\s*\(\s*"([^"]+)"/g)) {
          rustEmits.add(match[1]);
        }
        for (const match of rustText.matchAll(/\.emit\s*\(\s*([A-Z0-9_]+)\s*,/g)) {
          const value = resolveConst(match[1]);
          if (value) rustEmits.add(value);
        }

        for (const match of rustText.matchAll(/\.listen\s*\(\s*"([^"]+)"/g)) {
          rustListens.add(match[1]);
        }
        for (const match of rustText.matchAll(/\.listen\s*\(\s*([A-Z0-9_]+)\s*,/g)) {
          const value = resolveConst(match[1]);
          if (value) rustListens.add(value);
        }
        for (const match of rustText.matchAll(/\.listen_global\s*\(\s*"([^"]+)"/g)) {
          rustListens.add(match[1]);
        }
        for (const match of rustText.matchAll(/\.listen_global\s*\(\s*([A-Z0-9_]+)\s*,/g)) {
          const value = resolveConst(match[1]);
          if (value) rustListens.add(value);
        }
      }

      for (const event of rustEmits) {
        expect(allowListenEvents.has(event)).toBe(true);
      }
      for (const event of rustListens) {
        expect(allowEmitEvents.has(event)).toBe(true);
      }
    }

    // Some subsystems (e.g. the updater UI) `listen(...)` via an indirection (iterating over a
    // constant list of event names) instead of calling `listen("...")` directly. Keep those
    // allowlisted too.
    {
      const updaterUiPath = path.join(root, "apps/desktop/src/tauri/updaterUi.ts");
      const updaterUiText = stripComments(readFileSync(updaterUiPath, "utf8"));

      const eventsList = updaterUiText.match(/const\s+events\s*:\s*UpdaterEventName\[\]\s*=\s*\[([\s\S]*?)\];/);
      expect(eventsList).toBeTruthy();

      const updaterEvents = Array.from(eventsList![1].matchAll(/["']([^"']+)["']/g)).map((m) => m[1]);
      expect(updaterEvents.length).toBeGreaterThan(0);
      for (const event of updaterEvents) {
        expect(allowListenEvents.has(event)).toBe(true);
      }
    }

    for (const event of listenedByFrontend) {
      expect(allowListenEvents.has(event)).toBe(true);
    }
    for (const event of emittedByFrontend) {
      expect(allowEmitEvents.has(event)).toBe(true);
    }

    for (const event of allowlistedEvents) {
      // Only treat an allowlisted event as "used" if it appears as a string literal in runtime
      // source. This intentionally avoids counting documentation comments as usage (see the
      // canonical event lists near the desktop event wiring).
      const quoted = [`"${String(event)}"`, `'${String(event)}'`];
      expect(quoted.some((q) => runtimeText.includes(q))).toBe(true);
    }
  });
});
