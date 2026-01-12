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
        gotoDesktop(pageA, makeUrl({ id: "u-a", name: "User A", color: "#ff0000" }), { idleTimeoutMs: 10_000, appReadyTimeoutMs: 120_000 }),
        gotoDesktop(pageB, makeUrl({ id: "u-b", name: "User B", color: "#0000ff" }), { idleTimeoutMs: 10_000, appReadyTimeoutMs: 120_000 }),
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

      // --- 2.5) Multi-sheet comment isolation ---------------------------------
      //
      // Historically, desktop comments were keyed by A1 only, which caused collisions
      // across sheets (Sheet1!A1 vs Sheet2!A1). Ensure sheet-qualified comment refs
      // keep threads independent in real collab sessions.
      await pageA.evaluate(() => {
        const app = (window as any).__formulaApp;
        const session = app?.getCollabSession?.() ?? null;
        if (!session) throw new Error("Missing collab session");

        session.transactLocal(() => {
          const existingIds = new Set(
            (session.sheets?.toArray?.() ?? [])
              .map((entry: any) => String(entry?.get?.("id") ?? entry?.id ?? "").trim())
              .filter(Boolean),
          );
          if (existingIds.has("Sheet2")) return;

          const firstSheet: any = session.sheets.get(0);
          const MapCtor = firstSheet?.constructor ?? session.cells?.constructor ?? null;
          if (typeof MapCtor !== "function") throw new Error("Missing Y.Map constructor");
          const sheet = new MapCtor();
          sheet.set("id", "Sheet2");
          sheet.set("name", "Sheet2");
          sheet.set("visibility", "visible");
          session.sheets.insert(1, [sheet]);
        });
      });

      await Promise.all([
        expect(pageA.getByTestId("sheet-tab-Sheet2")).toBeVisible({ timeout: 30_000 }),
        expect(pageB.getByTestId("sheet-tab-Sheet2")).toBeVisible({ timeout: 30_000 }),
      ]);

      await Promise.all([
        pageA.getByTestId("sheet-tab-Sheet2").click(),
        pageB.getByTestId("sheet-tab-Sheet2").click(),
      ]);

      await Promise.all([
        expect.poll(() => pageA.evaluate(() => (window as any).__formulaApp.getCurrentSheetId())).toBe("Sheet2"),
        expect.poll(() => pageB.evaluate(() => (window as any).__formulaApp.getCurrentSheetId())).toBe("Sheet2"),
      ]);

      await pageA.locator("#grid").focus();
      await expect(pageA.getByTestId("active-cell")).toHaveText("A1");
      await pageB.locator("#grid").focus();
      await expect(pageB.getByTestId("active-cell")).toHaveText("A1");

      const commentTextSheet2 = `comment-sheet2-${Date.now()}`;
      await panelA.getByTestId("new-comment-input").fill(commentTextSheet2);
      await panelA.getByTestId("submit-comment").click();
      await expect(panelA.getByTestId("comment-thread").first()).toContainText(commentTextSheet2);

      // Wait for Sheet2 comment to sync to the other client.
      await pageB.waitForFunction(
        (text) => {
          const threads = Array.from(document.querySelectorAll('[data-testid="comment-thread"]'));
          return threads.some((el) => el.textContent?.includes(text));
        },
        commentTextSheet2,
        { timeout: 60_000 },
      );

      // Switch back to Sheet1 and ensure the Sheet2 A1 comment doesn't collide.
      await pageB.getByTestId("sheet-tab-Sheet1").click();
      await expect.poll(() => pageB.evaluate(() => (window as any).__formulaApp.getCurrentSheetId())).toBe("Sheet1");
      await expect(panelB).toContainText(commentText);
      await expect(panelB).not.toContainText(commentTextSheet2);

      await pageB.getByTestId("sheet-tab-Sheet2").click();
      await expect.poll(() => pageB.evaluate(() => (window as any).__formulaApp.getCurrentSheetId())).toBe("Sheet2");
      await expect(panelB).toContainText(commentTextSheet2);
      await expect(panelB).not.toContainText(commentText);

      // Restore both clients to Sheet1 so the presence assertions below remain stable.
      await Promise.all([
        pageA.getByTestId("sheet-tab-Sheet1").click(),
        pageB.getByTestId("sheet-tab-Sheet1").click(),
      ]);

      await Promise.all([
        expect.poll(() => pageA.evaluate(() => (window as any).__formulaApp.getCurrentSheetId())).toBe("Sheet1"),
        expect.poll(() => pageB.evaluate(() => (window as any).__formulaApp.getCurrentSheetId())).toBe("Sheet1"),
      ]);

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
