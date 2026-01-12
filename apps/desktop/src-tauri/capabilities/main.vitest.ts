import { readFileSync } from "node:fs";

import { describe, expect, it } from "vitest";

describe("Tauri capabilities (desktop main window)", () => {
  it('allows emitting the updater handshake event ("updater-ui-ready")', () => {
    const contents = readFileSync(new URL("./main.json", import.meta.url), "utf8");
    const capabilities = JSON.parse(contents) as {
      permissions?: Array<unknown>;
    };

    const permissions = capabilities.permissions;
    expect(Array.isArray(permissions)).toBe(true);

    const allowEmit = (permissions ?? []).find(
      (entry): entry is { identifier?: string; allow?: Array<{ event?: string }> } =>
        typeof entry === "object" && entry != null && (entry as any).identifier === "event:allow-emit",
    );

    expect(allowEmit).toBeTruthy();

    const allowedEvents = (allowEmit?.allow ?? [])
      .map((item) => item?.event)
      .filter((event): event is string => typeof event === "string");

    expect(allowedEvents).toContain("updater-ui-ready");
  });
});

