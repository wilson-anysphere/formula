import { expect, test } from "@playwright/test";

import { mkdtemp, rm } from "node:fs/promises";
import os from "node:os";
import path from "node:path";
import { randomUUID } from "node:crypto";

import { getAvailablePort, startSyncServer } from "../../../../services/sync-server/test/test-helpers";
import { gotoDesktop } from "./helpers";

test.describe("collaboration: sheet metadata", () => {
  test("syncs sheet list order + names from Yjs session.sheets across clients", async ({ browser }, testInfo) => {
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
        });
        return `/?${params.toString()}`;
      };

      await Promise.all([
        gotoDesktop(pageA, makeUrl({ id: "u-a", name: "User A" }), { idleTimeoutMs: 10_000 }),
        gotoDesktop(pageB, makeUrl({ id: "u-b", name: "User B" }), { idleTimeoutMs: 10_000 }),
      ]);

      // Wait for providers to complete initial sync before applying edits.
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

      // 1) Add a new sheet entry directly in Yjs (simulates version restore / branch checkout).
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

          const MapCtor = session.cells?.constructor ?? null;
          if (typeof MapCtor !== "function") throw new Error("Missing Y.Map constructor");
          const sheet = new MapCtor();
          sheet.set("id", "Sheet2");
          sheet.set("name", "Sheet2");
          sheet.set("visibility", "visible");
          session.sheets.insert(1, [sheet]);
        });
      });

      await expect(pageB.getByTestId("sheet-tab-Sheet2")).toBeVisible({ timeout: 30_000 });

      // 2) Rename Sheet1 by updating Yjs metadata (remote-driven rename).
      await pageA.evaluate(() => {
        const app = (window as any).__formulaApp;
        const session = app?.getCollabSession?.() ?? null;
        if (!session) throw new Error("Missing collab session");

        session.transactLocal(() => {
          for (let i = 0; i < session.sheets.length; i += 1) {
            const entry: any = session.sheets.get(i);
            const id = String(entry?.get?.("id") ?? entry?.id ?? "").trim();
            if (id !== "Sheet1") continue;
            if (typeof entry?.set !== "function") throw new Error("Sheet entry is not a Y.Map");
            entry.set("name", "Budget");
            return;
          }
          throw new Error("Sheet1 not found in session.sheets");
        });
      });

      await expect(pageB.getByTestId("sheet-tab-Sheet1").locator(".sheet-tab__name")).toHaveText("Budget", {
        timeout: 30_000,
      });

      // 3) Reorder Sheet2 before Sheet1 in Yjs (remote-driven reorder).
      await pageA.evaluate(() => {
        const app = (window as any).__formulaApp;
        const session = app?.getCollabSession?.() ?? null;
        if (!session) throw new Error("Missing collab session");

        session.transactLocal(() => {
          let fromIndex = -1;
          for (let i = 0; i < session.sheets.length; i += 1) {
            const entry: any = session.sheets.get(i);
            const id = String(entry?.get?.("id") ?? entry?.id ?? "").trim();
            if (id === "Sheet2") {
              fromIndex = i;
              break;
            }
          }
          if (fromIndex < 0) throw new Error("Sheet2 not found");
          if (fromIndex === 0) return;

          const entry: any = session.sheets.get(fromIndex);
          const MapCtor = session.cells?.constructor ?? null;
          if (typeof MapCtor !== "function") throw new Error("Missing Y.Map constructor");
          const clone = new MapCtor();
          if (entry && typeof entry.forEach === "function") {
            entry.forEach((v: any, k: string) => {
              clone.set(k, v);
            });
          } else if (entry && typeof entry === "object") {
            for (const [k, v] of Object.entries(entry)) {
              clone.set(k, v);
            }
          }

          session.sheets.delete(fromIndex, 1);
          session.sheets.insert(0, [clone]);
        });
      });

      await expect
        .poll(() =>
          pageB.evaluate(() =>
            Array.from(document.querySelectorAll("#sheet-tabs .sheet-tabs [data-sheet-id]")).map((el) =>
              (el as HTMLElement).getAttribute("data-sheet-id"),
            ),
          ),
        )
        .toEqual(["Sheet2", "Sheet1"]);

      // 4) Remove the currently active sheet (Sheet1) and ensure the UI auto-switches.
      await expect
        .poll(() => pageB.evaluate(() => (window as any).__formulaApp.getCurrentSheetId()))
        .toBe("Sheet1");

      await pageA.evaluate(() => {
        const app = (window as any).__formulaApp;
        const session = app?.getCollabSession?.() ?? null;
        if (!session) throw new Error("Missing collab session");

        session.transactLocal(() => {
          for (let i = 0; i < session.sheets.length; i += 1) {
            const entry: any = session.sheets.get(i);
            const id = String(entry?.get?.("id") ?? entry?.id ?? "").trim();
            if (id !== "Sheet1") continue;
            session.sheets.delete(i, 1);
            return;
          }
          throw new Error("Sheet1 not found for deletion");
        });
      });

      await expect
        .poll(() => pageB.evaluate(() => (window as any).__formulaApp.getCurrentSheetId()))
        .toBe("Sheet2");
    } finally {
      await Promise.allSettled([contextA.close(), contextB.close()]);
      await server.stop().catch(() => {});
      await rm(dataDir, { recursive: true, force: true }).catch(() => {});
    }
  });
});
