import { expect, test } from "@playwright/test";

import { mkdtemp, rm } from "node:fs/promises";
import os from "node:os";
import path from "node:path";
import { randomUUID } from "node:crypto";

import { getAvailablePort, startSyncServer } from "../../../../services/sync-server/test/test-helpers";
import { gotoDesktop } from "./helpers";

test.describe("collaboration: local persistence", () => {
  test("restores edits from IndexedDB before provider sync", async ({ browser }, testInfo) => {
    test.setTimeout(240_000);

    const baseURL = testInfo.project.use.baseURL;
    if (!baseURL) throw new Error("Playwright baseURL is required for collaboration e2e");

    const dataDir = await mkdtemp(path.join(os.tmpdir(), "formula-sync-"));
    const server = await startSyncServer({
      port: await getAvailablePort(),
      dataDir,
      auth: { mode: "opaque", token: "test-token" },
    });

    const docId = randomUUID();

    const makeUrl = (): string => {
      const params = new URLSearchParams({
        collab: "1",
        wsUrl: server.wsUrl,
        docId,
        token: "test-token",
        userId: "u-persist",
        userName: "Persist",
        // Ensure sync goes through the websocket server (not BroadcastChannel).
        disableBc: "1",
      });
      return `/?${params.toString()}`;
    };

    const context = await browser.newContext({ baseURL });
    const pageA = await context.newPage();

    try {
      await gotoDesktop(pageA, makeUrl(), { idleTimeoutMs: 10_000, appReadyTimeoutMs: 120_000 });

      // Wait for the provider to fully sync once so the session schema is initialized.
      await pageA.waitForFunction(() => {
        const app = (window as any).__formulaApp;
        const session = app?.getCollabSession?.() ?? null;
        return Boolean(session?.provider?.synced);
      }, undefined, { timeout: 60_000 });

      // Apply an edit and force persistence flush so the update is durably stored.
      await pageA.evaluate(async () => {
        const app = (window as any).__formulaApp;
        const sheetId = app.getCurrentSheetId();
        const doc = app.getDocument();
        const session = app.getCollabSession();

        doc.setCellValue(sheetId, { row: 0, col: 0 }, "persisted");
        doc.markDirty();

        // Wait a microtask so the DocumentController→Yjs binder can propagate before flushing.
        await new Promise<void>((resolve) => queueMicrotask(resolve));

        // Ensure the session sees the value in Yjs before flushing.
        for (let i = 0; i < 100; i += 1) {
          const cell = await session.getCell(`${sheetId}:0:0`);
          if (cell?.value === "persisted") break;
          await new Promise<void>((resolve) => setTimeout(resolve, 10));
        }

        await session.flushLocalPersistence();
      });

      await pageA.close();

      // Take the sync server down so the next session cannot report `provider.synced=true`.
      await server.stop();

      const pageB = await context.newPage();
      await gotoDesktop(pageB, makeUrl(), { waitForIdle: false });

      // Ensure local persistence completes even without a sync provider.
      await pageB.evaluate(async () => {
        const app = (window as any).__formulaApp;
        const session = app.getCollabSession();
        await session.whenLocalPersistenceLoaded();
      });

      // Assert the provider is not synced (the server is down), but the value is present.
      const providerSynced = await pageB.evaluate(() => {
        const app = (window as any).__formulaApp;
        const session = app?.getCollabSession?.() ?? null;
        return Boolean(session?.provider?.synced);
      });
      expect(providerSynced).toBe(false);

      await expect
        .poll(
          () =>
            pageB.evaluate(() => {
              const app = (window as any).__formulaApp;
              const doc = app.getDocument();
              const sheetId = app.getCurrentSheetId();
              return doc.getCell(sheetId, { row: 0, col: 0 })?.value ?? null;
            }),
          { timeout: 30_000 },
        )
        .toBe("persisted");

      await pageB.close();
    } finally {
      await Promise.allSettled([pageA.close(), context.close()]);
      await server.stop().catch(() => {});
      await rm(dataDir, { recursive: true, force: true }).catch(() => {});
    }
  });
});
