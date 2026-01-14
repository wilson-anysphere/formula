import { expect, test, type Locator, type Page } from "@playwright/test";
import { writeFile } from "node:fs/promises";

import { gotoDesktop } from "./helpers";

const TINY_PNG_BASE64 =
  // 1×1 transparent PNG
  "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8/5+hHgAHggJ/PYpgVwAAAABJRU5ErkJggg==";

async function whenIdle(page: Page, timeoutMs: number = 15_000): Promise<void> {
  // Vite may trigger a one-time full reload after dependency optimization.
  // Retry once if the execution context is destroyed mid-wait.
  for (let attempt = 0; attempt < 2; attempt += 1) {
    try {
      await page.waitForFunction(() => typeof (window.__formulaApp as any)?.whenIdle === "function", undefined, {
        timeout: timeoutMs,
      });
      await page.evaluate(async (timeoutMs) => {
        const app = window.__formulaApp as any;
        if (!app || typeof app.whenIdle !== "function") return;
        await Promise.race([app.whenIdle(), new Promise<void>((resolve) => setTimeout(resolve, timeoutMs))]);
      }, timeoutMs);
      return;
    } catch (err) {
      const message = err instanceof Error ? err.message : String(err);
      if (attempt === 0 && message.includes("Execution context was destroyed")) {
        await page.waitForLoadState("domcontentloaded");
        continue;
      }
      throw err;
    }
  }
}

async function getImageDrawingCount(page: Page): Promise<number> {
  return await page.evaluate(() => {
    const app = window.__formulaApp as any;
    if (!app) throw new Error("Missing window.__formulaApp (desktop e2e harness)");

    const objects = (() => {
      if (typeof app.getDrawingsDebugState === "function") {
        const state = app.getDrawingsDebugState();
        if (Array.isArray(state)) return state;
        if (state && typeof state === "object") {
          const anyState = state as any;
          if (Array.isArray(anyState.objects)) return anyState.objects;
          if (Array.isArray(anyState.drawings)) return anyState.drawings;
          if (Array.isArray(anyState.sheetDrawings)) return anyState.sheetDrawings;
          if (Array.isArray(anyState.value)) return anyState.value;
        }
        // If the debug helper exists but returns an unexpected shape, fall back to other
        // debug APIs when available (keeps this test resilient while the harness evolves).
      }
      if (typeof app.getDrawingObjects === "function") {
        return app.getDrawingObjects();
      }
      throw new Error(
        "Missing drawings debug API. Expected window.__formulaApp.getDrawingsDebugState() or window.__formulaApp.getDrawingObjects().",
      );
    })();

    return objects.filter((obj: any) => {
      // `SpreadsheetApp.getDrawingsDebugState()` returns `{ drawings: [{ kind: string, ... }] }`
      // while `SpreadsheetApp.getDrawingObjects()` returns `{ kind: { type: string, ... }, ... }`.
      const kind = obj?.kind;
      if (typeof kind === "string") return kind === "image";
      if (kind && typeof kind === "object" && typeof (kind as any).type === "string") return (kind as any).type === "image";
      return false;
    }).length;
  });
}

async function resolveLocator(
  root: Locator,
  {
    testId,
    commandId,
    role,
    name,
  }: { testId: string; commandId?: string; role?: Parameters<Locator["getByRole"]>[0]; name?: string },
): Promise<Locator> {
  const isAttached = async (locator: Locator): Promise<boolean> => {
    try {
      await locator.first().waitFor({ state: "attached", timeout: 1_000 });
      return true;
    } catch {
      return false;
    }
  };

  const byTestId = root.getByTestId(testId);
  if (await isAttached(byTestId)) return byTestId;
  if (commandId) {
    const byCommand = root.locator(`[data-command-id="${commandId}"]`);
    if (await isAttached(byCommand)) return byCommand;
  }
  if (role && name) {
    const byRole = root.getByRole(role, { name });
    if (await isAttached(byRole)) return byRole;
  }
  throw new Error(`Failed to resolve ribbon locator for testId=${testId}${commandId ? ` commandId=${commandId}` : ""}`);
}

test.describe("Insert → Pictures", () => {
  test("Insert → Pictures → This Device opens file picker and inserts image drawings", async ({ page }, testInfo) => {
    await gotoDesktop(page);
    await whenIdle(page);

    const ribbon = page.getByTestId("ribbon-root");
    await ribbon.getByRole("tab", { name: "Insert" }).click();

    const picturesDropdown = await resolveLocator(ribbon, {
      testId: "ribbon-insert-pictures",
      commandId: "insert.illustrations.pictures",
      role: "button",
      name: "Pictures",
    });
    await picturesDropdown.click();

    const thisDevice = await resolveLocator(ribbon, {
      testId: "ribbon-insert-pictures-this-device",
      commandId: "insert.illustrations.pictures.thisDevice",
      role: "menuitem",
      name: "This Device…",
    });

    const image1Path = testInfo.outputPath("tiny-1.png");
    const image2Path = testInfo.outputPath("tiny-2.png");
    const pngBytes = Buffer.from(TINY_PNG_BASE64, "base64");
    await Promise.all([writeFile(image1Path, pngBytes), writeFile(image2Path, pngBytes)]);

    const beforeCount = await getImageDrawingCount(page);

    let fileChooser: import("@playwright/test").FileChooser;
    try {
      [fileChooser] = await Promise.all([page.waitForEvent("filechooser", { timeout: 10_000 }), thisDevice.click()]);
    } catch (err) {
      const message = err instanceof Error ? err.message : String(err);
      throw new Error(
        `Expected a file chooser to open after clicking Insert → Pictures → This Device… but none was observed.\n\nOriginal error: ${message}`,
      );
    }

    const selectedPaths = fileChooser.isMultiple() ? [image1Path, image2Path] : [image1Path];
    await fileChooser.setFiles(selectedPaths);

    await expect
      .poll(
        async () => {
          await whenIdle(page, 5_000);
          return await getImageDrawingCount(page);
        },
        {
          timeout: 20_000,
          message: `Expected inserting ${selectedPaths.length} image file(s) to create ${selectedPaths.length} image drawing(s).`,
        },
      )
      .toBe(beforeCount + selectedPaths.length);
  });
});
