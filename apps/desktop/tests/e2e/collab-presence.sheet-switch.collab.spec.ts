import { expect, test } from "@playwright/test";
import path from "node:path";
import { fileURLToPath } from "node:url";

import { gotoDesktop } from "./helpers";
import { startLocalSyncServer, type LocalSyncServerHandle } from "./sync-server";

test.describe("collab presence: sheet switching", () => {
  test.describe.configure({ mode: "serial" });

  let syncServer: LocalSyncServerHandle | null = null;
  let wsUrl: string;

  test.beforeAll(async () => {
    const here = path.dirname(fileURLToPath(import.meta.url));
    const desktopRoot = path.resolve(here, "../..");
    const repoRoot = path.resolve(desktopRoot, "../..");
    // Reuse the same stable hash base as other collab e2e tests, but offset to avoid
    // port collisions when multiple sync-server-backed specs run in parallel.
    syncServer = await startLocalSyncServer({ repoRoot, portOffset: 1 });
    wsUrl = syncServer.wsUrl;
  });

  test.afterAll(async () => {
    const child = syncServer;
    syncServer = null;
    if (!child) return;
    await child.stop();
  });

  const GRID_MODES = ["legacy", "shared"] as const;

  for (const gridMode of GRID_MODES) {
    test(`filters remote presences by the local active sheet (${gridMode})`, async ({ browser }, testInfo) => {
      test.setTimeout(120_000);

      const baseURL = testInfo.project.use.baseURL;
      if (!baseURL) throw new Error("Playwright baseURL is required for collab presence e2e");

      // Use independent browser contexts so presence sync must travel through the
      // websocket server (not BroadcastChannel/shared storage).
      const context1 = await browser.newContext({ baseURL });
      const context2 = await browser.newContext({ baseURL });
      const page = await context1.newPage();
      const page2 = await context2.newPage();
      const docId = `e2e-doc-${gridMode}-${Date.now()}`;
      const token = "dev-token";

      const urlForUser = (userId: string): string => {
        const params = new URLSearchParams({
          grid: gridMode,
          collab: "1",
          docId,
          wsUrl,
          token,
          userId,
          userName: userId,
          // Ensure presence sync goes through the websocket server (not BroadcastChannel),
          // otherwise multi-tab tests can trigger awareness spoof filtering.
          disableBc: "1",
        });
        return `/?${params.toString()}`;
      };

      try {
        await Promise.all([gotoDesktop(page, urlForUser("user1")), gotoDesktop(page2, urlForUser("user2"))]);

        await expect(page.getByTestId("collab-status")).toBeVisible();
        await expect(page2.getByTestId("collab-status")).toBeVisible();
        await expect(page.getByTestId("collab-status")).toHaveAttribute("data-collab-sync", "synced", { timeout: 30_000 });
        await expect(page2.getByTestId("collab-status")).toHaveAttribute("data-collab-sync", "synced", { timeout: 30_000 });

        const remotePresenceNames = async (targetPage: typeof page): Promise<string[]> => {
          return targetPage.evaluate(() => {
            const app = window.__formulaApp as any;
            if (!app) return [];
            const session = app.getCollabSession?.() ?? null;
            const presence = session?.presence ?? null;
            if (!presence || typeof presence.getRemotePresences !== "function") return [];
            try {
              const presences = presence.getRemotePresences();
              const names = Array.isArray(presences)
                ? presences.map((p: any) => p?.name).filter((name: any) => typeof name === "string" && name.trim() !== "")
                : [];
              return names.sort();
            } catch {
              return [];
            }
          });
        };

        await Promise.all([
          expect.poll(() => page.evaluate(() => (window as any).__formulaApp.getGridMode?.() ?? null)).toBe(gridMode),
          expect.poll(() => page2.evaluate(() => (window as any).__formulaApp.getGridMode?.() ?? null)).toBe(gridMode),
        ]);

        // With both clients on the initial sheet, user2 should see user1's presence.
        await expect
          .poll(async () => remotePresenceNames(page2), { timeout: 30_000 })
          .toEqual(["user1"]);

        // Add a second sheet in user1.
        await page.getByTestId("sheet-add").click();

        await expect(page.getByRole("tab", { name: "Sheet2" })).toBeVisible({ timeout: 30_000 });
        await expect(page2.getByRole("tab", { name: "Sheet2" })).toBeVisible({ timeout: 30_000 });
        const sheet2Id = await page.getByRole("tab", { name: "Sheet2" }).getAttribute("data-sheet-id");
        if (!sheet2Id) throw new Error("Expected Sheet2 tab to have data-sheet-id attribute");

        // Explicitly activate the newly created sheet in user1. (Some builds keep the
        // original sheet active after creation.)
        await page.getByRole("tab", { name: "Sheet2" }).click();
        await expect(page.getByRole("tab", { name: "Sheet2" })).toHaveAttribute("aria-selected", "true", { timeout: 30_000 });

        // user2 stays on Sheet1, so user1 should no longer appear in user2's filtered remote presences.
        await expect(page2.getByRole("tab", { name: "Sheet1" })).toHaveAttribute("aria-selected", "true", { timeout: 30_000 });
        await expect
          .poll(async () => remotePresenceNames(page2), { timeout: 30_000 })
          .toEqual([]);

        // Ensure the newly created sheet is materialized in the underlying document model. The
        // DocumentController creates sheets lazily, and activation can fail if a sheet exists
        // only in metadata (tab strip) but has not been referenced by any cell edits yet.
        await page.evaluate((id) => {
          const app = (window as any).__formulaApp as any;
          const doc = app?.getDocument?.() ?? null;
          doc?.setCellValue?.(id, { row: 0, col: 0 }, "X");
        }, sheet2Id);
        await expect
          .poll(
            async () =>
              page2.evaluate((id) => {
                const app = (window as any).__formulaApp as any;
                const doc = app?.getDocument?.() ?? null;
                const ids = doc?.getSheetIds?.() ?? [];
                return Array.isArray(ids) ? ids.includes(id) : false;
              }, sheet2Id),
            { timeout: 30_000 },
          )
          .toBe(true);

        // When user2 switches to the same sheet, they should see user1 again.
        await page2.getByRole("tab", { name: "Sheet2" }).click();
        await expect
          .poll(async () => page2.evaluate(() => (window as any).__formulaApp.getCurrentSheetId?.() ?? null), { timeout: 30_000 })
          .toBe(sheet2Id);

        await expect
          .poll(async () => remotePresenceNames(page2), { timeout: 30_000 })
          .toEqual(["user1"]);
      } finally {
        await context1.close();
        await context2.close();
      }
    });
  }
});
