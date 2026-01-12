import { expect, test } from "@playwright/test";

import { gotoDesktop, waitForDesktopReady } from "./helpers";

async function waitForIdle(page: import("@playwright/test").Page): Promise<void> {
  await page.evaluate(() => (window as any).__formulaApp.whenIdle());
}

test.describe("collab: beforeunload unsaved-changes prompt", () => {
  test("does not show a beforeunload confirm dialog when a collab session is active", async ({ page }) => {
    await gotoDesktop(page);

    // Create a user gesture + local edit so browsers are allowed to show a beforeunload prompt.
    // Click inside A1 (avoid the shared-grid corner header/select-all region).
    await page.click("#grid", { position: { x: 80, y: 40 } });
    await page.keyboard.press("h");
    await page.keyboard.type("ello");
    await page.keyboard.press("Enter");
    await waitForIdle(page);

    await expect.poll(() => page.evaluate(() => (window as any).__formulaApp.getDocument().isDirty)).toBe(true);

    // Simulate collab mode by attaching the API expected by main.ts. In real collab mode
    // this is provided by the collaboration bootstrap layer.
    await page.evaluate(() => {
      const app = (window as any).__formulaApp;
      if (!app) throw new Error("Missing __formulaApp");

      const sheetId = (typeof app.getCurrentSheetId === "function" ? app.getCurrentSheetId() : null) ?? "Sheet1";
      const sheetsState: Array<{ id: string; name: string }> = [{ id: sheetId, name: sheetId }];

      const sheets = {
        toArray: () => sheetsState.map((s) => ({ ...s })),
        observeDeep: () => {},
        unobserveDeep: () => {},
        get length() {
          return sheetsState.length;
        },
        get: (idx: number) => {
          const entry = sheetsState[idx];
          if (!entry) return null;
          return {
            get: (key: string) => (key === "id" ? entry.id : key === "name" ? entry.name : undefined),
            set: (key: string, value: unknown) => {
              if (key === "id") entry.id = String(value ?? "");
              if (key === "name") entry.name = String(value ?? "");
            },
          };
        },
        insert: (index: number, items: any[]) => {
          if (!Array.isArray(items) || items.length === 0) return;
          const normalized: Array<{ id: string; name: string }> = [];
          for (const item of items) {
            const id = String((item as any)?.get?.("id") ?? (item as any)?.id ?? "").trim();
            if (!id) continue;
            const name = String((item as any)?.get?.("name") ?? (item as any)?.name ?? id);
            normalized.push({ id, name });
          }
          sheetsState.splice(index, 0, ...normalized);
        },
        delete: (index: number, count: number) => {
          const n = Number.isFinite(count) ? Math.max(0, Math.trunc(count)) : 0;
          sheetsState.splice(index, n);
        },
      };

      const session = {
        id: "e2e-collab-session",
        docId: "e2e-doc",
        sheets,
        transactLocal: (fn: () => void) => fn(),
      };

      app.getCollabSession = () => session;
    });

    let beforeUnloadDialogs = 0;
    page.on("dialog", async (dialog) => {
      if (dialog.type() === "beforeunload") beforeUnloadDialogs += 1;
      await dialog.accept();
    });

    await page.reload();
    await waitForDesktopReady(page);

    expect(beforeUnloadDialogs).toBe(0);
  });
});
