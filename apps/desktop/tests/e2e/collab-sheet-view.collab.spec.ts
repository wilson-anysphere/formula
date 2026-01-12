import { expect, test } from "@playwright/test";

import { randomUUID } from "node:crypto";
import { mkdtemp, rm } from "node:fs/promises";
import os from "node:os";
import path from "node:path";

import { getAvailablePort, startSyncServer } from "../../../../services/sync-server/test/test-helpers";
import { gotoDesktop } from "./helpers";

test.describe("collaboration: sheet view state", () => {
  test("syncs freeze panes + column resize across clients via session.sheets[i].view", async ({ browser }, testInfo) => {
    // Collab startup can be slow on first-run (WASM, python runtime, Vite optimize),
    // and we spin up two independent clients + a sync server.
    test.setTimeout(240_000);

    const baseURL = testInfo.project.use.baseURL;
    if (!baseURL) throw new Error("Playwright baseURL is required for collaboration e2e");

    const dataDir = await mkdtemp(path.join(os.tmpdir(), "formula-sync-"));
    const server = await startSyncServer({
      port: await getAvailablePort(),
      dataDir,
      auth: { mode: "opaque", token: "test-token" },
    });

    const contextA = await browser.newContext({ baseURL });
    const contextB = await browser.newContext({ baseURL });
    const pageA = await contextA.newPage();
    const pageB = await contextB.newPage();

    try {
      const docId = randomUUID();

      const makeUrl = (user: { id: string; name: string }): string => {
        const params = new URLSearchParams({
          collab: "1",
          wsUrl: server.wsUrl,
          docId,
          token: "test-token",
          userId: user.id,
          userName: user.name,
          // Ensure sync goes through the websocket server (not BroadcastChannel).
          disableBc: "1",
          // Enable axis resize interactions for the column resize assertion.
          grid: "shared",
        });
        return `/?${params.toString()}`;
      };

      await Promise.all([
        gotoDesktop(pageA, makeUrl({ id: "u-a", name: "User A" }), { idleTimeoutMs: 10_000 }),
        gotoDesktop(pageB, makeUrl({ id: "u-b", name: "User B" }), { idleTimeoutMs: 10_000 }),
      ]);

      // Wait for providers to complete initial sync before applying view-state edits.
      await Promise.all([
        pageA.waitForFunction(() => {
          const app = (window as any).__formulaApp;
          const session = app?.getCollabSession?.() ?? null;
          return Boolean(session?.provider?.synced);
        }, undefined, { timeout: 60_000 }),
        pageB.waitForFunction(() => {
          const app = (window as any).__formulaApp;
          const session = app?.getCollabSession?.() ?? null;
          return Boolean(session?.provider?.synced);
        }, undefined, { timeout: 60_000 }),
      ]);

      // ---- Freeze panes sync ----
      const gridA = pageA.locator("#grid");

      await pageA.waitForFunction(() => {
        const app = (window as any).__formulaApp;
        const rect = app?.getCellRectA1?.("B3");
        return rect && typeof rect.x === "number" && rect.width > 0 && rect.height > 0;
      }, undefined, { timeout: 30_000 });

      const b3Rect = await pageA.evaluate(() => (window as any).__formulaApp.getCellRectA1("B3"));
      if (!b3Rect) throw new Error("Missing B3 rect");

      // Select B3.
      await gridA.click({ position: { x: b3Rect.x + b3Rect.width / 2, y: b3Rect.y + b3Rect.height / 2 } });
      await expect(pageA.getByTestId("active-cell")).toHaveText("B3");

      // Open command palette and run "Freeze Panes".
      const primary = process.platform === "darwin" ? "Meta" : "Control";
      await pageA.keyboard.press(`${primary}+Shift+P`);
      await expect(pageA.getByTestId("command-palette")).toBeVisible();
      await pageA.keyboard.type("Freeze Panes");
      await expect(pageA.getByTestId("command-palette-list")).toContainText("View");
      await pageA.keyboard.press("Enter");

      await expect
        .poll(() => pageA.evaluate(() => (window as any).__formulaApp.getFrozen()), { timeout: 30_000 })
        .toEqual({ frozenRows: 2, frozenCols: 1 });

      await expect
        .poll(() => pageB.evaluate(() => (window as any).__formulaApp.getFrozen()), { timeout: 30_000 })
        .toEqual({ frozenRows: 2, frozenCols: 1 });

      // ---- Column resize sync ----
      const beforeB1 = await pageB.evaluate(() => (window as any).__formulaApp.getCellRectA1("B1"));
      if (!beforeB1) throw new Error("Missing B1 rect on page B");

      const gridBoxA = await pageA.locator("#grid").boundingBox();
      if (!gridBoxA) throw new Error("Missing grid bounding box");

      const b1RectA = await pageA.evaluate(() => (window as any).__formulaApp.getCellRectA1("B1"));
      if (!b1RectA) throw new Error("Missing B1 rect on page A");

      // Drag the boundary between columns A and B in the header row to make column A wider.
      const boundaryX = b1RectA.x;
      const boundaryY = b1RectA.y / 2;

      await pageA.mouse.move(gridBoxA.x + boundaryX, gridBoxA.y + boundaryY);
      await pageA.mouse.down();
      await pageA.mouse.move(gridBoxA.x + boundaryX + 80, gridBoxA.y + boundaryY, { steps: 4 });
      await pageA.mouse.up();

      // Verify the resize took effect locally before waiting for remote propagation.
      await pageA.waitForFunction(
        (threshold) => {
          const rect = (window as any).__formulaApp.getCellRectA1("B1");
          return rect && rect.x > threshold;
        },
        b1RectA.x + 30,
        { timeout: 30_000 },
      );

      await expect
        .poll(() => pageB.evaluate(() => (window as any).__formulaApp.getCellRectA1("B1")?.x ?? -1), { timeout: 30_000 })
        .toBeGreaterThan(beforeB1.x + 30);
    } finally {
      await Promise.allSettled([contextA.close(), contextB.close()]);
      await server.stop().catch(() => {});
      await rm(dataDir, { recursive: true, force: true }).catch(() => {});
    }
  });
});

