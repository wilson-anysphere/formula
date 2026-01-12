import { describe, expect, it } from "vitest";
import { readFileSync } from "node:fs";

describe("desktop updater listener consolidation", () => {
  it("does not register updater UX toast listeners in main.ts (handled by tauri/updaterUi.ts)", () => {
    const mainUrl = new URL("../../main.ts", import.meta.url);
    const source = readFileSync(mainUrl, "utf8");

    // These updater UX events should be handled by `tauri/updaterUi.ts` so that manual checks
    // don't produce duplicate toasts and startup checks don't accidentally surface dialogs/toasts.
    expect(source).not.toContain('listen("update-not-available"');
    expect(source).not.toContain('listen("update-check-error"');
    expect(source).not.toContain('listen("update-available"');

    // `main.ts` may listen to some updater events for non-UX bookkeeping, but it should not
    // install inline handlers that could easily re-introduce user-facing toast duplication.
    expect(source).not.toMatch(/listen\("update-check-started",\s*(?:async\s*)?\(/);
    expect(source).not.toMatch(/listen\("update-check-started",\s*function\b/);
    expect(source).not.toMatch(/listen\("update-check-already-running",\s*(?:async\s*)?\(/);
    expect(source).not.toMatch(/listen\("update-check-already-running",\s*function\b/);
  });
});

