import { expect, test } from "@playwright/test";
import path from "node:path";
import { fileURLToPath } from "node:url";

import { gotoDesktop, waitForDesktopReady } from "./helpers";
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

    // Toggling network offline should transition to a disconnected/reconnecting state.
    await page.context().setOffline(true);
    await expect(page.getByTestId("collab-status")).toHaveAttribute("data-collab-conn", "offline", { timeout: 10_000 });
    await expect(page2.getByTestId("collab-status")).toHaveAttribute("data-collab-conn", "offline", { timeout: 10_000 });
    await expect(page.getByTestId("collab-status")).toHaveAttribute("data-collab-sync", "offline", { timeout: 10_000 });
    await expect(page2.getByTestId("collab-status")).toHaveAttribute("data-collab-sync", "offline", { timeout: 10_000 });
    await expect
      .poll(async () => page.getByTestId("collab-status").getAttribute("data-collab-conn"))
      .not.toBe("connected");
    await expect
      .poll(async () => page2.getByTestId("collab-status").getAttribute("data-collab-conn"))
      .not.toBe("connected");

    await page.context().setOffline(false);
    await expect(page.getByTestId("collab-status")).toHaveAttribute("data-collab-sync", "synced", { timeout: 60_000 });
    await expect(page2.getByTestId("collab-status")).toHaveAttribute("data-collab-sync", "synced", { timeout: 60_000 });

    // Ensure collab mode never triggers the browser/Tauri beforeunload unsaved prompt, even if
    // DocumentController is dirty.
    await page.locator("#grid").click({ position: { x: 80, y: 40 } });
    await page.keyboard.press("F2");
    const editor = page.locator("textarea.cell-editor");
    await expect(editor).toBeVisible();
    await editor.fill("dirty");
    await page.keyboard.press("Enter");
    await expect.poll(() => page.evaluate(() => (window.__formulaApp as any).getDocument().isDirty)).toBe(true);

    let beforeUnloadDialogs = 0;
    page.on("dialog", async (dialog) => {
      if (dialog.type() === "beforeunload") beforeUnloadDialogs += 1;
      await dialog.accept();
    });

    await page.reload({ waitUntil: "domcontentloaded" });
    await waitForDesktopReady(page, { idleTimeoutMs: 10_000 });

    expect(beforeUnloadDialogs).toBe(0);
  });
});
