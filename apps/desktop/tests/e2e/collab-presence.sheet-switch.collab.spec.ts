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

  test("filters remote presences by the local active sheet", async ({ page }) => {
    test.setTimeout(120_000);

    const page2 = await page.context().newPage();
    const docId = `e2e-doc-${Date.now()}`;
    const token = "dev-token";

    const urlForUser = (userId: string): string => {
      const params = new URLSearchParams({
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

    // With both clients on the initial sheet, user2 should see user1's presence.
    await expect
      .poll(async () => {
        return page2.evaluate(() => {
          const app = (window as any).__formulaApp;
          const session = app?.getCollabSession?.() ?? null;
          const presences = session?.presence?.getRemotePresences?.() ?? [];
          return presences.map((p: any) => p?.id ?? null).filter((id: any) => typeof id === "string");
        });
      })
      .toEqual(["user1"]);

    // Add a second sheet in user1; the desktop UI automatically activates it.
    await page.getByTestId("sheet-add").click();

    await expect(page.getByRole("tab", { name: "Sheet2" })).toBeVisible({ timeout: 30_000 });
    await expect(page.getByRole("tab", { name: "Sheet2" })).toHaveAttribute("aria-selected", "true", { timeout: 30_000 });
    await expect(page2.getByRole("tab", { name: "Sheet2" })).toBeVisible({ timeout: 30_000 });

    // user2 stays on Sheet1, so user1 should no longer appear in user2's filtered remote presences.
    await expect
      .poll(async () => {
        return page2.evaluate(() => {
          const app = (window as any).__formulaApp;
          const session = app?.getCollabSession?.() ?? null;
          const presences = session?.presence?.getRemotePresences?.() ?? [];
          return presences.map((p: any) => p?.id ?? null).filter((id: any) => typeof id === "string");
        });
      })
      .toEqual([]);

    // When user2 switches to the same sheet, they should see user1 again.
    await page2.getByRole("tab", { name: "Sheet2" }).click();
    await expect(page2.getByRole("tab", { name: "Sheet2" })).toHaveAttribute("aria-selected", "true", { timeout: 30_000 });

    await expect
      .poll(async () => {
        return page2.evaluate(() => {
          const app = (window as any).__formulaApp;
          const session = app?.getCollabSession?.() ?? null;
          const presences = session?.presence?.getRemotePresences?.() ?? [];
          return presences.map((p: any) => p?.id ?? null).filter((id: any) => typeof id === "string");
        });
      })
      .toEqual(["user1"]);
  });
});
