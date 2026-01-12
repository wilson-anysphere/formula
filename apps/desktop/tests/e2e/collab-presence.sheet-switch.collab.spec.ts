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
    test(`filters remote presences by the local active sheet (${gridMode})`, async ({ page }) => {
      test.setTimeout(120_000);

      const page2 = await page.context().newPage();
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
        });
        return `/?${params.toString()}`;
      };

      await Promise.all([gotoDesktop(page, urlForUser("user1")), gotoDesktop(page2, urlForUser("user2"))]);

      await expect(page.getByTestId("collab-status")).toBeVisible();
      await expect(page2.getByTestId("collab-status")).toBeVisible();
      await expect(page.getByTestId("collab-status")).toHaveAttribute("data-collab-sync", "synced", { timeout: 30_000 });
      await expect(page2.getByTestId("collab-status")).toHaveAttribute("data-collab-sync", "synced", { timeout: 30_000 });

      const remotePresenceIds = async (targetPage: typeof page): Promise<string[]> => {
        return targetPage.evaluate(() => {
          const app = window.__formulaApp as any;
          if (!app) return [];
          const mode = app.getGridMode?.() ?? null;

        // Shared-grid mode: presences are pushed into the CanvasGridRenderer selection layer.
        if (mode === "shared") {
          const renderer = (app as any).sharedGrid?.renderer;
          const presences = renderer?.remotePresences ?? [];
          const ids = Array.isArray(presences)
            ? presences.map((presence: any) => presence?.id).filter((id: any) => typeof id === "string")
            : [];
          return ids.sort();
        }

          // Legacy mode: remote presences are stored on the app and rendered via `renderPresence()`.
          const presences = (app as any).remotePresences ?? [];
          const ids = Array.isArray(presences)
            ? presences.map((presence: any) => presence?.id).filter((id: any) => typeof id === "string")
            : [];
          return ids.sort();
        });
      };

      // With both clients on the initial sheet, user2 should see user1's presence.
      await expect
        .poll(async () => remotePresenceIds(page2), { timeout: 30_000 })
        .toEqual(["user1"]);

      // Add a second sheet in user1; the desktop UI automatically activates it.
      await page.getByTestId("sheet-add").click();

      await expect(page.getByRole("tab", { name: "Sheet2" })).toBeVisible({ timeout: 30_000 });
      await expect(page.getByRole("tab", { name: "Sheet2" })).toHaveAttribute("aria-selected", "true", { timeout: 30_000 });
      await expect(page2.getByRole("tab", { name: "Sheet2" })).toBeVisible({ timeout: 30_000 });

      // user2 stays on Sheet1, so user1 should no longer appear in user2's filtered remote presences.
      await expect
        .poll(async () => remotePresenceIds(page2), { timeout: 30_000 })
        .toEqual([]);

      // When user2 switches to the same sheet, they should see user1 again.
      await page2.getByRole("tab", { name: "Sheet2" }).click();
      await expect(page2.getByRole("tab", { name: "Sheet2" })).toHaveAttribute("aria-selected", "true", { timeout: 30_000 });

      await expect
        .poll(async () => remotePresenceIds(page2), { timeout: 30_000 })
        .toEqual(["user1"]);
    });
  }
});
