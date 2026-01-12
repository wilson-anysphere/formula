import { expect, test } from "@playwright/test";

import { mkdtemp, rm } from "node:fs/promises";
import os from "node:os";
import path from "node:path";
import { randomUUID } from "node:crypto";

import { getAvailablePort, startSyncServer } from "../../../../services/sync-server/test/test-helpers";
import { gotoDesktop } from "./helpers";

test.describe("collaboration", () => {
  test("syncs cells, comments, and presence across two clients", async ({ browser }, testInfo) => {
    // Desktop startup can be slow on first-run (WASM, python runtime, Vite optimize),
    // and this test spins up two independent clients + a sync server.
    test.setTimeout(240_000);

    const baseURL = testInfo.project.use.baseURL;
    if (!baseURL) throw new Error("Playwright baseURL is required for collaboration e2e");

    const dataDir = await mkdtemp(path.join(os.tmpdir(), "formula-sync-"));
    const server = await startSyncServer({
      // Prefer an explicit port so parallel test workers don't accidentally collide.
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

      const makeUrl = (user: { id: string; name: string; color: string }): string => {
        const params = new URLSearchParams({
          collab: "1",
          wsUrl: server.wsUrl,
          docId,
          token: "test-token",
          userId: user.id,
          userName: user.name,
          userColor: user.color,
          // Ensure sync goes through the websocket server (not BroadcastChannel).
          disableBc: "1",
        });
        return `/?${params.toString()}`;
      };

      // Load each client in its own browser context so sync must travel through the
      // websocket server (not BroadcastChannel).
      await Promise.all([
        gotoDesktop(pageA, makeUrl({ id: "u-a", name: "User A", color: "#ff0000" }), { idleTimeoutMs: 10_000 }),
        gotoDesktop(pageB, makeUrl({ id: "u-b", name: "User B", color: "#0000ff" }), { idleTimeoutMs: 10_000 }),
      ]);

      // Wait for the websocket providers to complete initial sync before applying edits.
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

      // --- 1) Cell sync ----------------------------------------------------

      await pageA.locator("#grid").focus();
      await expect(pageA.getByTestId("active-cell")).toHaveText("A1");

      await pageB.locator("#grid").focus();
      await expect(pageB.getByTestId("active-cell")).toHaveText("A1");

      const cellText = `collab-${Date.now()}`;
      await pageA.keyboard.press("F2");
      const editor = pageA.locator("textarea.cell-editor");
      await expect(editor).toBeVisible();
      await editor.fill(cellText);
      await pageA.keyboard.press("Enter");

      await pageB.waitForFunction(
        async (expected) => {
          const app = (window as any).__formulaApp;
          const value = await app.getCellValueA1("A1");
          return value === expected;
        },
        cellText,
        { timeout: 60_000 }
      );

      // --- 2) Comment sync -------------------------------------------------

      // Editing a cell and pressing Enter advances selection to the next row. To
      // keep both clients aligned, re-select A1 before adding the comment.
      await pageA.locator("#grid").focus();
      await pageA.keyboard.press("ArrowUp");
      await expect(pageA.getByTestId("active-cell")).toHaveText("A1");
      await pageB.locator("#grid").focus();
      await expect(pageB.getByTestId("active-cell")).toHaveText("A1");

      const commentText = `comment-${Date.now()}`;

      await pageA.getByTestId("ribbon-root").getByTestId("open-comments-panel").click();
      const panelA = pageA.getByTestId("comments-panel");
      await expect(panelA).toBeVisible();
      await panelA.getByTestId("new-comment-input").fill(commentText);
      await panelA.getByTestId("submit-comment").click();
      await expect(panelA.getByTestId("comment-thread").first()).toContainText(commentText);

      await pageB.getByTestId("ribbon-root").getByTestId("open-comments-panel").click();
      const panelB = pageB.getByTestId("comments-panel");
      await expect(panelB).toBeVisible();

      await pageB.waitForFunction(
        (text) => {
          const threads = Array.from(document.querySelectorAll('[data-testid="comment-thread"]'));
          return threads.some((el) => el.textContent?.includes(text));
        },
        commentText,
        { timeout: 60_000 }
      );

      await expect(panelB.getByTestId("comment-thread").first()).toContainText(commentText);

      // --- 3) Presence -----------------------------------------------------

      // Move selection on page A to ensure we publish a presence cursor update.
      await pageA.locator("#grid").focus();
      await pageA.keyboard.press("ArrowRight"); // B1
      await pageA.keyboard.press("ArrowDown"); // B2
      await expect(pageA.getByTestId("active-cell")).toHaveText("B2");

      await pageB.waitForFunction(() => {
        const app = (window as any).__formulaApp;
        const session = app?.getCollabSession?.() ?? null;
        const presence = session?.presence ?? null;
        if (!presence) return false;
        try {
          return presence.getRemotePresences({ includeOtherSheets: true }).length > 0;
        } catch {
          return false;
        }
      }, undefined, { timeout: 60_000 });
    } finally {
      await Promise.allSettled([contextA.close(), contextB.close()]);
      await server.stop().catch(() => {});
      await rm(dataDir, { recursive: true, force: true }).catch(() => {});
    }
  });
});
