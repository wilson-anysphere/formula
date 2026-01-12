import { expect, test } from "@playwright/test";
import path from "node:path";
import { fileURLToPath } from "node:url";

import { gotoDesktop } from "./helpers";
import { startLocalSyncServer, type LocalSyncServerHandle } from "./sync-server";

test.describe("collab status indicator (collab mode)", () => {
  test.describe.configure({ mode: "serial" });

  let syncServer: LocalSyncServerHandle | null = null;
  let wsUrl: string;

  test.beforeAll(async () => {
    const here = path.dirname(fileURLToPath(import.meta.url));
    const desktopRoot = path.resolve(here, "../..");
    const repoRoot = path.resolve(desktopRoot, "../..");
    syncServer = await startLocalSyncServer({ repoRoot, portOffset: 0 });
    wsUrl = syncServer.wsUrl;
  });

  test.afterAll(async () => {
    const child = syncServer;
    syncServer = null;
    if (!child) return;
    await child.stop();
  });

  test("shows Synced after connecting to the sync server", async ({ page }) => {
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
    await expect(page.getByTestId("collab-status")).toHaveAttribute("data-collab-mode", "collab");
    await expect(page2.getByTestId("collab-status")).toHaveAttribute("data-collab-mode", "collab");
    await expect(page.getByTestId("collab-status")).toHaveAttribute("data-collab-doc-id", docId);
    await expect(page2.getByTestId("collab-status")).toHaveAttribute("data-collab-doc-id", docId);

    await expect(page.getByTestId("collab-status")).toContainText(docId);
    await expect(page2.getByTestId("collab-status")).toContainText(docId);

    await expect(page.getByTestId("collab-status")).toHaveAttribute("data-collab-sync", "synced", { timeout: 30_000 });
    await expect(page2.getByTestId("collab-status")).toHaveAttribute("data-collab-sync", "synced", { timeout: 30_000 });
    await expect(page.getByTestId("collab-status")).toHaveAttribute("data-collab-conn", "connected");
    await expect(page2.getByTestId("collab-status")).toHaveAttribute("data-collab-conn", "connected");
  });
});
