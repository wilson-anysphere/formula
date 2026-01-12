import { expect, test, type Page } from "@playwright/test";

import { gotoDesktop, waitForDesktopReady } from "./helpers";

async function grantSampleHelloPermissions(page: Page): Promise<void> {
  await page.evaluate(() => {
    const extensionId = "formula.sample-hello";
    const key = "formula.extensionHost.permissions";
    const existing = (() => {
      try {
        const raw = localStorage.getItem(key);
        return raw ? JSON.parse(raw) : {};
      } catch {
        return {};
      }
    })();

    existing[extensionId] = {
      ...(existing[extensionId] ?? {}),
      "ui.commands": true,
      "ui.panels": true,
      "cells.read": true,
      "cells.write": true,
    };

    localStorage.setItem(key, JSON.stringify(existing));
  });
}

test.describe("Extensions UI integration", () => {
  // The desktop shell has a large ribbon; the default Playwright viewport height can
  // leave too little space for the grid, making hit-testing unreliable. Use a
  // taller viewport so context menu/selection interactions have room.
  test.use({ viewport: { width: 1280, height: 900 } });

  test("shows extension commands in the command palette after lazy-loading extensions", async ({ page }) => {
    await gotoDesktop(page);

    // Open the command palette first (without opening the Extensions panel) and
    // ensure that extension-contributed commands appear once the extension host loads.
    const primary = process.platform === "darwin" ? "Meta" : "Control";
    await page.keyboard.press(`${primary}+Shift+P`);
    await expect(page.getByTestId("command-palette")).toBeVisible();

    await page.getByTestId("command-palette-input").fill("Sum Selection");

    // Command palette groups commands by category, rendering the category as a group header
    // and the command title as the selectable row.
    const list = page.getByTestId("command-palette-list");
    await expect(list).toContainText("Sample Hello", { timeout: 10_000 });
    await expect(list).toContainText("Sum Selection", { timeout: 10_000 });
  });

  test("runs sampleHello.openPanel and renders the panel webview", async ({ page }) => {
    await gotoDesktop(page);
    await grantSampleHelloPermissions(page);

    await page.getByTestId("open-extensions-panel").click();
    const openPanelBtn = page.getByTestId("run-command-sampleHello.openPanel");
    await expect(openPanelBtn).toBeVisible({ timeout: 30_000 });
    // Avoid hit-target flakiness from fixed overlays by dispatching a click directly.
    await openPanelBtn.dispatchEvent("click");

    await expect(page.getByTestId("panel-sampleHello.panel")).toBeAttached();
    const frame = page.frameLocator('iframe[data-testid="extension-webview-sampleHello.panel"]');
    await expect(frame.locator("h1")).toHaveText("Sample Hello Panel");
    await expect(
      frame.locator('meta[http-equiv="Content-Security-Policy"]'),
      "webview should inject a restrictive CSP meta tag",
    ).toHaveCount(1);
  });

  test("runs sampleHello.sumSelection via the Extensions panel and shows a toast", async ({ page }) => {
    await gotoDesktop(page);
    await grantSampleHelloPermissions(page);

    await page.evaluate(() => {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const app: any = (window as any).__formulaApp;
      if (!app) throw new Error("Missing window.__formulaApp (desktop e2e harness)");
      const doc = app.getDocument();
      const sheetId = app.getCurrentSheetId();

      doc.setCellValue(sheetId, { row: 0, col: 0 }, 1);
      doc.setCellValue(sheetId, { row: 0, col: 1 }, 2);
      doc.setCellValue(sheetId, { row: 1, col: 0 }, 3);
      doc.setCellValue(sheetId, { row: 1, col: 1 }, 4);

      app.selectRange({
        sheetId,
        range: { startRow: 0, startCol: 0, endRow: 1, endCol: 1 },
      });
    });

    await page.getByTestId("open-extensions-panel").click();
    const sumSelectionBtn = page.getByTestId("run-command-sampleHello.sumSelection");
    await expect(sumSelectionBtn).toBeVisible({ timeout: 30_000 });
    await sumSelectionBtn.dispatchEvent("click");

    await expect(page.getByTestId("toast-root")).toContainText("Sum: 10");
  });

  test("persists an extension panel in the layout and re-activates it after reload", async ({ page }) => {
    await gotoDesktop(page);
    await grantSampleHelloPermissions(page);

    await page.getByTestId("open-extensions-panel").click();
    const openPanelBtn = page.getByTestId("run-command-sampleHello.openPanel");
    await expect(openPanelBtn).toBeVisible({ timeout: 30_000 });
    await openPanelBtn.dispatchEvent("click");

    await expect(page.getByTestId("panel-sampleHello.panel")).toBeAttached();
    const frame = page.frameLocator('iframe[data-testid="extension-webview-sampleHello.panel"]');
    await expect(frame.locator("h1")).toHaveText("Sample Hello Panel");

    await page.reload();
    await waitForDesktopReady(page);
    await grantSampleHelloPermissions(page);

    // Reloading the page resets the extension host; opening the command palette triggers
    // the lazy extension host boot so persisted extension panels can re-activate.
    const primary = process.platform === "darwin" ? "Meta" : "Control";
    await page.keyboard.press(`${primary}+Shift+P`);
    await expect(page.getByTestId("command-palette")).toBeVisible();
    await page.getByTestId("command-palette-input").fill("Sum Selection");
    await expect(page.getByTestId("command-palette-list")).toContainText("Sample Hello: Sum Selection", {
      timeout: 10_000,
    });
    await page.keyboard.press("Escape");

    await expect(page.getByTestId("panel-sampleHello.panel")).toBeAttached();
    const frameAfter = page.frameLocator('iframe[data-testid="extension-webview-sampleHello.panel"]');
    await expect(frameAfter.locator("h1")).toHaveText("Sample Hello Panel");
  });

  test("executes a contributed keybinding when its when-clause matches", async ({ page }) => {
    await gotoDesktop(page);
    await grantSampleHelloPermissions(page);

    await page.evaluate(() => {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const app: any = (window as any).__formulaApp;
      if (!app) throw new Error("Missing window.__formulaApp (desktop e2e harness)");
      const doc = app.getDocument();
      const sheetId = app.getCurrentSheetId();

      doc.setCellValue(sheetId, { row: 0, col: 0 }, 1);
      doc.setCellValue(sheetId, { row: 0, col: 1 }, 2);
      doc.setCellValue(sheetId, { row: 1, col: 0 }, 3);
      doc.setCellValue(sheetId, { row: 1, col: 1 }, 4);

      app.selectRange({
        sheetId,
        range: { startRow: 0, startCol: 0, endRow: 1, endCol: 1 },
      });
    });

    await page.getByTestId("open-extensions-panel").click();
    await expect(page.getByTestId("run-command-sampleHello.sumSelection")).toBeVisible();

    await page.keyboard.press("Control+Shift+Y");
    await expect(page.getByTestId("toast-root")).toContainText("Sum: 10");
  });

  test("does not execute a keybinding when its when-clause fails", async ({ page }) => {
    await gotoDesktop(page);
    await grantSampleHelloPermissions(page);

    await page.evaluate(() => {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const app: any = (window as any).__formulaApp;
      if (!app) throw new Error("Missing window.__formulaApp (desktop e2e harness)");
      const doc = app.getDocument();
      const sheetId = app.getCurrentSheetId();
      doc.setCellValue(sheetId, { row: 0, col: 0 }, 5);
    });

    await page.getByTestId("open-extensions-panel").click();
    await expect(page.getByTestId("run-command-sampleHello.sumSelection")).toBeVisible();

    const before = await page.evaluate(() => {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const app: any = (window as any).__formulaApp;
      const doc = app.getDocument();
      const sheetId = app.getCurrentSheetId();
      const cell = doc.getCell(sheetId, { row: 2, col: 0 }) as any;
      return cell?.value ?? null;
    });

    // Default selection is a single cell, so `hasSelection` should be false and the keybinding should be ignored.
    await page.keyboard.press("Control+Shift+Y");

    // Give the UI a brief moment in case a command mistakenly fires.
    await page.waitForTimeout(250);

    const after = await page.evaluate(() => {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const app: any = (window as any).__formulaApp;
      const doc = app.getDocument();
      const sheetId = app.getCurrentSheetId();
      const cell = doc.getCell(sheetId, { row: 2, col: 0 }) as any;
      return cell?.value ?? null;
    });

    expect(after).toEqual(before);
  });

  test("loads extensions when opening the grid context menu", async ({ page }) => {
    await gotoDesktop(page);
    await grantSampleHelloPermissions(page);

    // Open the context menu without first opening the Extensions panel.
    await page.locator("#grid").click({ button: "right", position: { x: 100, y: 40 } });
    const menu = page.getByTestId("context-menu");
    await expect(menu).toBeVisible();

    // Extension contributions should appear once the lazy-loaded extension host finishes
    // initializing.
    const item = menu.getByRole("button", { name: "Sample Hello: Open Sample Panel" });
    await expect(item).toBeVisible({ timeout: 30_000 });
  });

  test("executes a contributed context menu item when its when-clause matches", async ({ page }) => {
    await gotoDesktop(page);
    await grantSampleHelloPermissions(page);

    await page.evaluate(() => {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const app: any = (window as any).__formulaApp;
      if (!app) throw new Error("Missing window.__formulaApp (desktop e2e harness)");
      const doc = app.getDocument();
      const sheetId = app.getCurrentSheetId();

      doc.setCellValue(sheetId, { row: 0, col: 0 }, 1);
      doc.setCellValue(sheetId, { row: 0, col: 1 }, 2);
      doc.setCellValue(sheetId, { row: 1, col: 0 }, 3);
      doc.setCellValue(sheetId, { row: 1, col: 1 }, 4);

      app.selectRange({
        sheetId,
        range: { startRow: 0, startCol: 0, endRow: 1, endCol: 1 },
      });
    });

    await page.getByTestId("open-extensions-panel").click();
    await expect(page.getByTestId("run-command-sampleHello.sumSelection")).toBeVisible();

    // Right-click inside the selection so the selection remains intact and `hasSelection` stays true.
    await page.evaluate(() => {
      const grid = document.getElementById("grid");
      if (!grid) throw new Error("Missing #grid container");
      const rect = grid.getBoundingClientRect();
      grid.dispatchEvent(
        new MouseEvent("contextmenu", {
          bubbles: true,
          cancelable: true,
          button: 2,
          clientX: rect.left + 100,
          clientY: rect.top + 40,
        }),
      );
    });
    const menu = page.getByTestId("context-menu");
    await expect(menu).toBeVisible();

    const item = menu.getByRole("button", { name: "Sample Hello: Sum Selection" });
    await expect(item).toBeEnabled();
    await item.click();

    await expect(page.getByTestId("toast-root")).toContainText("Sum: 10");
  });

  test("right-clicking outside a multi-cell selection moves the active cell before showing the menu", async ({ page }) => {
    await gotoDesktop(page);
    await grantSampleHelloPermissions(page);

    await page.evaluate(() => {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const app: any = (window as any).__formulaApp;
      if (!app) throw new Error("Missing window.__formulaApp (desktop e2e harness)");
      const sheetId = app.getCurrentSheetId();
      app.selectRange({
        sheetId,
        range: { startRow: 0, startCol: 0, endRow: 1, endCol: 1 },
      });
    });

    await expect(page.getByTestId("active-cell")).toHaveText("A1");
    await expect(page.getByTestId("selection-range")).toHaveText("A1:B2");

    // Ensure the extensions host is running so the contributed context menu renders.
    await page.getByTestId("open-extensions-panel").click();
    await expect(page.getByTestId("run-command-sampleHello.openPanel")).toBeVisible();

    // Ensure the grid has a usable hit-test surface. In headless e2e environments the
    // surrounding shell (ribbon/status bar) can leave the grid with near-zero layout
    // height, which makes `pickCellAtClientPoint` return null for all coordinates.
    await page.evaluate(() => {
      const grid = document.getElementById("grid");
      if (!grid) return;
      grid.style.height = "600px";
      grid.style.minHeight = "600px";
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const app: any = (window as any).__formulaApp;
      try {
        app?.onResize?.();
      } catch {
        // ignore
      }
    });

    // Wait for the grid renderer to fully initialize its viewport mapping so hit-testing
    // works reliably (otherwise `pickCellAtClientPoint` can report A1 for all points).
    await page.waitForFunction(() => {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const app: any = (window as any).__formulaApp;
      const rect = app?.getCellRectA1?.("D4");
      return Boolean(rect && rect.width > 0 && rect.height > 0);
    });

    // Right-click a cell outside the current selection. Excel/Sheets move the active
    // cell to the clicked cell before showing the menu so commands apply to it.
    const d4Point = await page.evaluate(() => {
      const grid = document.getElementById("grid");
      if (!grid) throw new Error("Missing #grid container");
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const app: any = (window as any).__formulaApp;
      if (
        !app?.getCellRectA1 ||
        !app?.pickCellAtClientPoint ||
        typeof app.getCellRectA1 !== "function"
      ) {
        throw new Error("Missing required SpreadsheetApp test helpers");
      }

      const target = { row: 3, col: 3 };
      const rect = app.getCellRectA1("D4");
      if (!rect) throw new Error("Missing D4 rect");

      const gridRect = grid.getBoundingClientRect();
      // `getCellRectA1` is a test helper, but its coordinate space differs depending
      // on the underlying grid renderer. Use `pickCellAtClientPoint` to validate
      // which candidate coordinate maps back to D4.
      const candidates = [
        // Treat rect as already viewport-relative.
        { x: rect.x + rect.width / 2, y: rect.y + rect.height / 2 },
        // Treat rect as grid-root relative (need to add the grid's viewport offset).
        { x: gridRect.left + rect.x + rect.width / 2, y: gridRect.top + rect.y + rect.height / 2 },
      ];

      for (const point of candidates) {
        const picked = app.pickCellAtClientPoint(point.x, point.y);
        if (picked && picked.row === target.row && picked.col === target.col) return point;
      }

      const debug = {
        rect,
        gridRect: { left: gridRect.left, top: gridRect.top, width: gridRect.width, height: gridRect.height },
        picked: candidates.map((point) => ({ point, picked: app.pickCellAtClientPoint(point.x, point.y) })),
      };
      throw new Error(`Failed to locate D4 client coordinates for context menu test: ${JSON.stringify(debug)}`);
    });

    await page.evaluate((point) => {
      const grid = document.getElementById("grid");
      if (!grid) throw new Error("Missing #grid container");
      grid.dispatchEvent(
        new MouseEvent("contextmenu", {
          bubbles: true,
          cancelable: true,
          button: 2,
          clientX: point.x,
          clientY: point.y,
        }),
      );
    }, d4Point);

    const menu = page.getByTestId("context-menu");
    await expect(menu).toBeVisible();
    await expect(page.getByTestId("active-cell")).toHaveText("D4");
  });

  test("shared grid: right-click inside selection preserves it; outside selection moves active cell", async ({ page }) => {
    await gotoDesktop(page, "/?grid=shared");
    await grantSampleHelloPermissions(page);

    await page.evaluate(() => {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const app: any = (window as any).__formulaApp;
      if (!app) throw new Error("Missing window.__formulaApp (desktop e2e harness)");
      const sheetId = app.getCurrentSheetId();
      app.selectRange({
        sheetId,
        range: { startRow: 0, startCol: 0, endRow: 1, endCol: 1 },
      });
    });

    // Ensure the extensions host is running so the contributed context menu renders.
    await page.getByTestId("open-extensions-panel").click();
    await expect(page.getByTestId("panel-extensions")).toBeVisible();
    await expect(page.getByTestId("run-command-sampleHello.openPanel")).toBeVisible();

    // Right-click inside the selection on a non-active cell; selection should remain multi-cell.
    const b2 = await page.evaluate(() => {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const app: any = (window as any).__formulaApp;
      return app.getCellRectA1("B2");
    });
    if (!b2) throw new Error("Missing B2 rect");
    await page.locator("#grid").click({
      button: "right",
      position: { x: b2.x + b2.width / 2, y: b2.y + b2.height / 2 },
    });

    const menu = page.getByTestId("context-menu");
    await expect(menu).toBeVisible();
    const sumItem = menu.getByRole("button", { name: "Sample Hello: Sum Selection" });
    await expect(sumItem, "inside selection should keep hasSelection=true").toBeEnabled();

    // Close the menu so we can open it again on a different cell.
    await page.keyboard.press("Escape");
    await expect(menu).toBeHidden();

    // Right-click outside the selection should move active cell (and collapse selection).
    const d4 = await page.evaluate(() => {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const app: any = (window as any).__formulaApp;
      return app.getCellRectA1("D4");
    });
    if (!d4) throw new Error("Missing D4 rect");
    await page.locator("#grid").click({
      button: "right",
      position: { x: d4.x + d4.width / 2, y: d4.y + d4.height / 2 },
    });

    await expect(menu).toBeVisible();
    await expect(page.getByTestId("active-cell")).toHaveText("D4");
    const sumItemAfter = menu.getByRole("button", { name: "Sample Hello: Sum Selection" });
    await expect(sumItemAfter, "outside selection should collapse selection (hasSelection=false)").toBeDisabled();
  });
});
