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

      await expect(pageB.getByTestId("sheet-tab-Sheet2")).toBeVisible({ timeout: 30_000 });

      // 1.5) Perform sheet-tab UI actions on client A and assert they propagate to client B
      // via the shared session.sheets schema (this exercises CollabWorkbookSheetStore write-backs).
      //
      // Hide Sheet2.
      await pageA.getByTestId("sheet-tab-Sheet2").click({ button: "right" });
      const menuA = pageA.getByTestId("sheet-tab-context-menu");
      await expect(menuA).toBeVisible();
      await menuA.getByRole("button", { name: "Hide" }).click();
      await expect(pageA.locator('[data-testid="sheet-tab-Sheet2"]')).toHaveCount(0);
      await expect(pageB.locator('[data-testid="sheet-tab-Sheet2"]')).toHaveCount(0, { timeout: 30_000 });

      // Unhide Sheet2.
      await pageA.getByTestId("sheet-tab-Sheet1").click({ button: "right" });
      await expect(menuA).toBeVisible();
      await menuA.getByRole("button", { name: "Unhide…" }).click();
      await menuA.getByRole("button", { name: "Sheet2" }).click();
      await expect(pageA.getByTestId("sheet-tab-Sheet2")).toBeVisible();
      await expect(pageB.getByTestId("sheet-tab-Sheet2")).toBeVisible({ timeout: 30_000 });

      // Set Sheet2 tab color (pick a non-red color so we can distinguish from Sheet1 later).
      await pageA.getByTestId("sheet-tab-Sheet2").click({ button: "right" });
      await expect(menuA).toBeVisible();
      await menuA.getByRole("button", { name: "Tab Color" }).click();
      await menuA.getByRole("button", { name: "Blue" }).click();
      await expect(pageB.getByTestId("sheet-tab-Sheet2")).toHaveAttribute("data-tab-color", "#0070c0", {
        timeout: 30_000,
      });

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

      // 2.5) Set Sheet1 tab color via Yjs metadata (remote-driven tab color).
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
            entry.set("tabColor", "FFFF0000");
            return;
          }
          throw new Error("Sheet1 not found in session.sheets");
        });
      });

      await expect(pageB.getByTestId("sheet-tab-Sheet1")).toHaveAttribute("data-tab-color", "#ff0000", {
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
          const MapCtor = entry?.constructor ?? session.cells?.constructor ?? null;
          if (typeof MapCtor !== "function") throw new Error("Missing Y.Map constructor");
          const clone = new MapCtor();
          // Only copy stable sheet-list metadata fields. Sheet entries can contain nested
          // Yjs types (e.g. `view`), and reusing those objects during a move can violate
          // Yjs' parentage rules. (The production code uses a deep clone for this.)
          const id = String(entry?.get?.("id") ?? entry?.id ?? "").trim();
          if (!id) throw new Error("Sheet entry missing id");
          clone.set("id", id);
          const name = entry?.get?.("name") ?? entry?.name;
          if (name != null) clone.set("name", name);
          const visibility = entry?.get?.("visibility") ?? entry?.visibility;
          if (visibility != null) clone.set("visibility", visibility);
          const tabColor = entry?.get?.("tabColor") ?? entry?.tabColor;
          if (tabColor != null) clone.set("tabColor", tabColor);

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

      // 3.5) Mark Sheet2 as "veryHidden" and ensure it is not shown in the tab UI.
      await pageA.evaluate(() => {
        const app = (window as any).__formulaApp;
        const session = app?.getCollabSession?.() ?? null;
        if (!session) throw new Error("Missing collab session");

        session.transactLocal(() => {
          for (let i = 0; i < session.sheets.length; i += 1) {
            const entry: any = session.sheets.get(i);
            const id = String(entry?.get?.("id") ?? entry?.id ?? "").trim();
            if (id !== "Sheet2") continue;
            if (typeof entry?.set !== "function") throw new Error("Sheet entry is not a Y.Map");
            entry.set("visibility", "veryHidden");
            return;
          }
          throw new Error("Sheet2 not found in session.sheets");
        });
      });

      await expect(pageB.locator('[data-testid="sheet-tab-Sheet2"]')).toHaveCount(0, { timeout: 30_000 });

      // 3.6) Restore Sheet2 to visible and ensure it reappears.
      await pageA.evaluate(() => {
        const app = (window as any).__formulaApp;
        const session = app?.getCollabSession?.() ?? null;
        if (!session) throw new Error("Missing collab session");

        session.transactLocal(() => {
          for (let i = 0; i < session.sheets.length; i += 1) {
            const entry: any = session.sheets.get(i);
            const id = String(entry?.get?.("id") ?? entry?.id ?? "").trim();
            if (id !== "Sheet2") continue;
            if (typeof entry?.set !== "function") throw new Error("Sheet entry is not a Y.Map");
            entry.set("visibility", "visible");
            return;
          }
          throw new Error("Sheet2 not found in session.sheets");
        });
      });

      await expect(pageB.getByTestId("sheet-tab-Sheet2")).toBeVisible({ timeout: 30_000 });
      await expect(pageB.getByTestId("sheet-tab-Sheet1")).toHaveAttribute("data-tab-color", "#ff0000", {
        timeout: 30_000,
      });
      await expect(pageB.getByTestId("sheet-tab-Sheet2")).toHaveAttribute("data-tab-color", "#0070c0", {
        timeout: 30_000,
      });

      // 4) Hide the currently active sheet (Sheet1) and ensure the UI auto-switches.
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
            entry.set("visibility", "hidden");
            return;
          }
          throw new Error("Sheet1 not found for hide");
        });
      });

      await expect(pageB.locator('[data-testid="sheet-tab-Sheet1"]')).toHaveCount(0, { timeout: 30_000 });
      await expect
        .poll(() => pageB.evaluate(() => (window as any).__formulaApp.getCurrentSheetId()))
        .toBe("Sheet2");

      // 5) Unhide Sheet1 and ensure the tab returns (while staying on Sheet2).
      await pageA.evaluate(() => {
        const app = (window as any).__formulaApp;
        const session = app?.getCollabSession?.() ?? null;
        if (!session) throw new Error("Missing collab session");

        session.transactLocal(() => {
          for (let i = 0; i < session.sheets.length; i += 1) {
            const entry: any = session.sheets.get(i);
            const id = String(entry?.get?.("id") ?? entry?.id ?? "").trim();
            if (id !== "Sheet1") continue;
            entry.set("visibility", "visible");
            return;
          }
          throw new Error("Sheet1 not found for unhide");
        });
      });

      await expect(pageB.getByTestId("sheet-tab-Sheet1")).toBeVisible({ timeout: 30_000 });
      await expect(pageB.getByTestId("sheet-tab-Sheet1")).toHaveAttribute("data-tab-color", "#ff0000", {
        timeout: 30_000,
      });
      await expect
        .poll(() => pageB.evaluate(() => (window as any).__formulaApp.getCurrentSheetId()))
        .toBe("Sheet2");

      // 6) Remove Sheet1 entirely and ensure the remaining client stays on Sheet2.
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

      await expect(pageB.locator('[data-testid="sheet-tab-Sheet1"]')).toHaveCount(0, { timeout: 30_000 });
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
