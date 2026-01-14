import { describe, expect, it } from "vitest";
import { readFileSync } from "node:fs";

import { stripComments } from "../../__tests__/sourceTextUtils";

describe("desktop updater listener consolidation", () => {
  it("does not register updater UX toast listeners in main.ts (handled by tauri/updaterUi.ts)", () => {
    const mainUrl = new URL("../../main.ts", import.meta.url);
    const source = stripComments(readFileSync(mainUrl, "utf8"));

    const escapeRegExp = (value: string) => value.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
    // `main.ts` wraps listen() calls in `listenBestEffort(...)` so they can't throw during
    // startup. Count either form to ensure we don't regress and add extra updater listeners.
    const listenCallRe = (eventName: string) =>
      new RegExp(String.raw`\b(?:listen|listenBestEffort)\(\s*['"]${escapeRegExp(eventName)}['"]`, "g");
    const listenCallCount = (eventName: string) => (source.match(listenCallRe(eventName)) ?? []).length;

    // These updater UX events should be handled by `tauri/updaterUi.ts` so that manual checks
    // don't produce duplicate toasts and startup checks don't accidentally surface dialogs/toasts.
    expect(listenCallCount("update-not-available")).toBe(0);
    expect(listenCallCount("update-check-error")).toBe(0);
    expect(listenCallCount("update-available")).toBe(0);

    // `main.ts` may listen to some updater events for non-UX bookkeeping, but it should not
    // install toast-producing listeners. Only the dedicated `recordManualUpdateCheckEvent` listener
    // (used to suppress duplicate backend checks) is permitted.
    expect(listenCallCount("update-check-started")).toBe(1);
    expect(source).toMatch(
      /\b(?:listen|listenBestEffort)\(\s*['"]update-check-started['"]\s*,\s*recordManualUpdateCheckEvent\b/,
    );

    expect(listenCallCount("update-check-already-running")).toBe(1);
    expect(source).toMatch(
      /\b(?:listen|listenBestEffort)\(\s*['"]update-check-already-running['"]\s*,\s*recordManualUpdateCheckEvent\b/,
    );

    // The `updater-ui-ready` handshake should be tied to the updater UI listeners being installed,
    // without adding extra updater listeners (like a second `update-available` handler).
    expect(source).toMatch(/\bemit\(\s*['"]updater-ui-ready['"]/);
    // Use a lazy match so this doesn't become brittle as `main.ts` grows.
    expect(source).toMatch(/void\s+updaterUiListeners[\s\S]*?emit\(\s*['"]updater-ui-ready['"]/);
    expect(source).not.toMatch(/Promise\.all\(\s*\[\s*updaterUiListeners\s*,/);
  });
});
