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
});
