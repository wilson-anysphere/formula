import { expect, test } from "@playwright/test";

import { mkdtemp, rm } from "node:fs/promises";
import os from "node:os";
import path from "node:path";
import { randomUUID } from "node:crypto";

import { getAvailablePort, startSyncServer } from "../../../../services/sync-server/test/test-helpers";
import { expectSheetPosition, gotoDesktop, openSheetTabContextMenu } from "./helpers";

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
      const pollSheetSwitcher = async (expected: { value: string; options: Array<{ value: string; label: string }> }) => {
        await expect
          .poll(
            () =>
              pageB.evaluate(() => {
                const el = document.querySelector<HTMLSelectElement>('[data-testid="sheet-switcher"]');
                if (!el) return null;
                return {
                  value: el.value,
                  options: Array.from(el.options).map((opt) => ({ value: opt.value, label: opt.textContent ?? "" })),
                };
              }),
            { timeout: 30_000 },
          )
          .toEqual(expected);
      };

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
          // IndexedDB-backed collab persistence can delay sheet schema initialization on first-run.
          // Disable it in this e2e so Sheet1 is deterministically present before we mutate `session.sheets`.
          collabPersistence: "0",
        });
        return `/?${params.toString()}`;
      };

      // Start clients sequentially to avoid concurrent "brand new doc" schema init races that can
      // briefly create duplicate Sheet1 entries (and transient empty-sheet lists) on first sync.
      await gotoDesktop(pageA, makeUrl({ id: "u-a", name: "User A" }), { idleTimeoutMs: 10_000, appReadyTimeoutMs: 120_000 });
      await gotoDesktop(pageB, makeUrl({ id: "u-b", name: "User B" }), { idleTimeoutMs: 10_000, appReadyTimeoutMs: 120_000 });

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

      // Wait for the default sheet metadata ("Sheet1") to be materialized in the Yjs sheet list.
      // In collab mode, `provider.synced` can flip true before schema initialization finishes.
      await Promise.all([
        pageA.waitForFunction(() => {
          const app = (window as any).__formulaApp;
          const session = app?.getCollabSession?.() ?? null;
          if (!session?.provider?.synced) return false;
          if (session?.sheets?.length !== 1) return false;
          const entry = session?.sheets?.get?.(0) ?? null;
          const id = String(entry?.get?.("id") ?? entry?.id ?? "").trim();
          return id === "Sheet1";
        }, undefined, { timeout: 60_000 }),
        pageB.waitForFunction(() => {
          const app = (window as any).__formulaApp;
          const session = app?.getCollabSession?.() ?? null;
          if (!session?.provider?.synced) return false;
          if (session?.sheets?.length !== 1) return false;
          const entry = session?.sheets?.get?.(0) ?? null;
          const id = String(entry?.get?.("id") ?? entry?.id ?? "").trim();
          return id === "Sheet1";
        }, undefined, { timeout: 60_000 }),
      ]);

      // Ensure the UI sheet switcher reflects the initial state before edits.
      await pollSheetSwitcher({
        value: "Sheet1",
        options: [{ value: "Sheet1", label: "Sheet1" }],
      });
      await expectSheetPosition(pageB, { position: 1, total: 1 }, { timeoutMs: 30_000 });

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
          // Insert Sheet2 after Sheet1 (if present), rather than assuming Sheet1 is at index 0.
          // Collab schema init can temporarily create placeholder sheet entries that have not had
          // their `id` fields applied yet; those placeholders should not influence the user-visible
          // ordering expectations in this test.
          let insertIndex = session.sheets.length;
          for (let i = 0; i < session.sheets.length; i += 1) {
            const entry: any = session.sheets.get(i);
            const id = String(entry?.get?.("id") ?? entry?.id ?? "").trim();
            if (id === "Sheet1") {
              insertIndex = i + 1;
              break;
            }
          }
          // Clamp to the current array length (Yjs throws if index > length).
          insertIndex = Math.max(0, Math.min(insertIndex, session.sheets.length));
          session.sheets.insert(insertIndex, [sheet]);
        });
      });

      await expect(pageB.getByTestId("sheet-tab-Sheet2")).toBeVisible({ timeout: 30_000 });
      await pollSheetSwitcher({
        value: "Sheet1",
        options: [
          { value: "Sheet1", label: "Sheet1" },
          { value: "Sheet2", label: "Sheet2" },
        ],
      });
      await expectSheetPosition(pageB, { position: 1, total: 2 }, { timeoutMs: 30_000 });

      const menuA = pageA.getByTestId("sheet-tab-context-menu");
      const openSheetTabContextMenuA = async (sheetId: string) => {
        await openSheetTabContextMenu(pageA, sheetId);
      };

      // 1.5) Perform sheet-tab UI actions on client A and assert they propagate to client B
      // via the shared session.sheets schema (this exercises CollabWorkbookSheetStore write-backs).
      //
      // Hide Sheet2.
      const sheet2TabA = pageA.getByTestId("sheet-tab-Sheet2");
      await expect(sheet2TabA).toBeVisible();
      await openSheetTabContextMenuA("Sheet2");
      await menuA.getByRole("button", { name: "Hide", exact: true }).click();
      await expect(pageA.locator('[data-testid="sheet-tab-Sheet2"]')).toHaveCount(0);
      await expect(pageB.locator('[data-testid="sheet-tab-Sheet2"]')).toHaveCount(0, { timeout: 30_000 });
      await expectSheetPosition(pageB, { position: 1, total: 1 }, { timeoutMs: 30_000 });

      // Unhide Sheet2.
      const sheet1TabA = pageA.getByTestId("sheet-tab-Sheet1");
      await openSheetTabContextMenuA("Sheet1");
      await menuA.getByRole("button", { name: "Unhide…", exact: true }).click();
      await menuA.getByRole("button", { name: "Sheet2" }).click();
      await expect(pageA.getByTestId("sheet-tab-Sheet2")).toBeVisible();
      await expect(pageB.getByTestId("sheet-tab-Sheet2")).toBeVisible({ timeout: 30_000 });
      await expectSheetPosition(pageB, { position: 1, total: 2 }, { timeoutMs: 30_000 });

      // Set Sheet2 tab color (pick a non-red color so we can distinguish from Sheet1 later).
      await openSheetTabContextMenuA("Sheet2");
      // Avoid a flaky click on the submenu button (menu can reflow during collab updates).
      // Navigate via keyboard instead: Rename -> Hide -> Tab Color -> Blue.
      await pageA.keyboard.press("ArrowDown");
      await pageA.keyboard.press("ArrowDown");
      await pageA.keyboard.press("ArrowRight");
      // No Color -> Red -> Orange -> Yellow -> Green -> Teal -> Blue
      for (let i = 0; i < 6; i += 1) {
        await pageA.keyboard.press("ArrowDown");
      }
      await pageA.keyboard.press("Enter");
      await expect(pageB.getByTestId("sheet-tab-Sheet2")).toHaveAttribute("data-tab-color", "#0070c0", {
        timeout: 30_000,
      });

      // Rename Sheet2 via the UI on client A and ensure it propagates to client B.
      await sheet2TabA.dblclick();
      const renameInputA = sheet2TabA.locator("input.sheet-tab__input");
      await expect(renameInputA).toBeVisible();
      await expect(renameInputA).toBeFocused();
      await renameInputA.fill("Expenses");
      await renameInputA.press("Enter");

      await expect(sheet2TabA.locator(".sheet-tab__name")).toHaveText("Expenses");
      await expect(pageB.getByTestId("sheet-tab-Sheet2").locator(".sheet-tab__name")).toHaveText("Expenses", { timeout: 30_000 });

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
      await pollSheetSwitcher({
        value: "Sheet1",
        options: [
          { value: "Sheet1", label: "Budget" },
          { value: "Sheet2", label: "Expenses" },
        ],
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

      // 3) Reorder Sheet2 before Sheet1 via the UI on client A.
      //
      // Avoid Playwright's `dragTo` in the desktop shell (can hang); dispatch a synthetic drop event instead.
      await pageA.evaluate(() => {
        const fromId = "Sheet2";
        const target = document.querySelector('[data-testid="sheet-tab-Sheet1"]') as HTMLElement | null;
        if (!target) throw new Error("Missing Sheet1 tab");
        const rect = target.getBoundingClientRect();

        const dt = new DataTransfer();
        dt.setData("text/sheet-id", fromId);
        dt.setData("text/plain", fromId);

        const drop = new DragEvent("drop", {
          bubbles: true,
          cancelable: true,
          clientX: rect.left + 1,
          clientY: rect.top + rect.height / 2,
        });
        Object.defineProperty(drop, "dataTransfer", { value: dt });
        target.dispatchEvent(drop);
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
      await pollSheetSwitcher({
        value: "Sheet1",
        options: [
          { value: "Sheet2", label: "Expenses" },
          { value: "Sheet1", label: "Budget" },
        ],
      });
      await expectSheetPosition(pageB, { position: 2, total: 2 }, { timeoutMs: 30_000 });

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
      await pollSheetSwitcher({
        value: "Sheet1",
        options: [{ value: "Sheet1", label: "Budget" }],
      });
      await expectSheetPosition(pageB, { position: 1, total: 1 }, { timeoutMs: 30_000 });

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
      await pollSheetSwitcher({
        value: "Sheet1",
        options: [
          { value: "Sheet2", label: "Expenses" },
          { value: "Sheet1", label: "Budget" },
        ],
      });
      await expectSheetPosition(pageB, { position: 2, total: 2 }, { timeoutMs: 30_000 });

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
      await pollSheetSwitcher({
        value: "Sheet2",
        options: [{ value: "Sheet2", label: "Expenses" }],
      });
      await expectSheetPosition(pageB, { position: 1, total: 1 }, { timeoutMs: 30_000 });

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
      await pollSheetSwitcher({
        value: "Sheet2",
        options: [
          { value: "Sheet2", label: "Expenses" },
          { value: "Sheet1", label: "Budget" },
        ],
      });
      await expectSheetPosition(pageB, { position: 1, total: 2 }, { timeoutMs: 30_000 });

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
      await pollSheetSwitcher({
        value: "Sheet2",
        options: [{ value: "Sheet2", label: "Expenses" }],
      });
      await expectSheetPosition(pageB, { position: 1, total: 1 }, { timeoutMs: 30_000 });
    } finally {
      await Promise.allSettled([contextA.close(), contextB.close()]);
      await server.stop().catch(() => {});
      await rm(dataDir, { recursive: true, force: true }).catch(() => {});
    }
  });
});
