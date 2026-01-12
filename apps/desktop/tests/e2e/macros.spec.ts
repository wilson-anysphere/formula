import { expect, test } from "@playwright/test";

import { gotoDesktop, waitForDesktopReady } from "./helpers";

test.describe("macros panel", () => {
  test("running a macro applies returned cell updates and is undoable as one step", async ({ page }) => {
    await page.addInitScript(() => {
      const listeners: Record<string, any> = {};
      (window as any).__tauriListeners = listeners;

      // Avoid interactive confirm() prompts from the macro security controller.
      window.confirm = () => true;

      (window as any).__TAURI__ = {
        core: {
          invoke: async (cmd: string, args: any) => {
            switch (cmd) {
              case "list_macros":
                return [{ id: "m1", name: "Module1.Macro1", language: "vba" }];

              case "get_macro_security_status":
                return {
                  has_macros: true,
                  origin_path: null,
                  workbook_fingerprint: null,
                  trust: "trusted_always",
                  signature: { status: "unsigned" },
                };

              case "set_macro_trust":
                return {
                  has_macros: true,
                  origin_path: null,
                  workbook_fingerprint: null,
                  trust: args?.decision ?? "trusted_always",
                  signature: { status: "unsigned" },
                };

              case "set_macro_ui_context":
                return null;

              case "run_macro":
                return {
                  ok: true,
                  output: [],
                  updates: [
                    {
                      sheet_id: "Sheet1",
                      row: 0,
                      col: 0,
                      value: "FromMacro",
                      formula: null,
                      display_value: "FromMacro",
                    },
                    {
                      sheet_id: "Sheet1",
                      row: 1,
                      col: 0,
                      value: "MacroA2",
                      formula: null,
                      display_value: "MacroA2",
                    },
                    {
                      sheet_id: "Sheet2",
                      row: 0,
                      col: 0,
                      value: "OtherSheetA1",
                      formula: null,
                      display_value: "OtherSheetA1",
                    },
                  ],
                  error: null,
                };

              // Host sync calls (no-op in this test harness).
              case "set_cell":
              case "set_range":
              case "save_workbook":
                return null;

              default:
                throw new Error(`Unexpected invoke: ${cmd} ${JSON.stringify(args)}`);
            }
          },
        },
        event: {
          listen: async (name: string, handler: any) => {
            listeners[name] = handler;
            return () => {
              delete listeners[name];
            };
          },
          emit: async () => {},
        },
        window: {
          getCurrentWebviewWindow: () => ({
            hide: async () => {
              (window as any).__tauriHidden = true;
            },
            close: async () => {
              (window as any).__tauriClosed = true;
            },
          }),
        },
      };
    });

    await gotoDesktop(page);
    await page.evaluate(() => localStorage.clear());
    await page.reload({ waitUntil: "domcontentloaded" });
    await waitForDesktopReady(page);

    await page.getByTestId("ribbon-root").getByTestId("open-macros-panel").click();
    const panel = page.getByTestId("dock-right").getByTestId("panel-macros");
    await expect(panel).toBeVisible();

    await expect(panel.locator("select")).toBeVisible();
    await panel.getByRole("button", { name: "Run" }).click();

    await expect(page.getByTestId("active-value")).toHaveText("FromMacro");
    await page.waitForFunction(async () => (await (window as any).__formulaApp.getCellValueA1("A2")) === "MacroA2");
    await expect(page.getByTestId("sheet-tabs").getByTestId("sheet-tab-Sheet2")).toBeVisible();
    const sheet2a1 = await page.evaluate(() => (window as any).__formulaApp.getDocument().getCell("Sheet2", "A1").value);
    expect(sheet2a1).toBe("OtherSheetA1");

    // Focus the grid to ensure keyboard shortcuts route to the SpreadsheetApp handler.
    // Click inside A1 (avoid the shared-grid corner header/select-all region).
    await page.click("#grid", { position: { x: 80, y: 40 } });

    const modifier = process.platform === "darwin" ? "Meta" : "Control";
    await page.keyboard.press(`${modifier}+Z`);

    await expect(page.getByTestId("active-value")).toHaveText("Seed");
    const a2 = await page.evaluate(() => (window as any).__formulaApp.getCellValueA1("A2"));
    expect(a2).toBe("A");
    const sheet2a1AfterUndo = await page.evaluate(
      () => (window as any).__formulaApp.getDocument().getCell("Sheet2", "A1").value,
    );
    expect(sheet2a1AfterUndo).toBeNull();
  });

  test("runs TypeScript + Python macros in the web demo", async ({ page }) => {
    test.setTimeout(120_000);

    await page.addInitScript(() => localStorage.clear());
    page.on("dialog", (dialog) => dialog.accept());

    await gotoDesktop(page);

    await page.getByTestId("ribbon-root").getByTestId("open-macros-panel").click();

    const panel = page.getByTestId("dock-right").getByTestId("panel-macros");
    await expect(panel).toBeVisible();

    const body = panel.locator(".dock-panel__body");
    const select = body.locator("select");
    await expect(select).toBeVisible();

    // Run TypeScript macro (writes E1).
    await select.selectOption({ label: "TypeScript: Write E1" });
    const runButton = body.getByRole("button", { name: "Run" });
    for (let attempt = 0; attempt < 2; attempt += 1) {
      await runButton.click();
      await expect(runButton).toBeDisabled();
      await expect(runButton).toBeEnabled({ timeout: 40_000 });

      try {
        await expect
          .poll(() => page.evaluate(() => (window as any).__formulaApp.getCellValueA1("E1")), { timeout: 10_000 })
          .toBe("hello from ts");
        break;
      } catch (err) {
        if (attempt === 0) continue;
        throw err;
      }
    }

    // Run Python macro (writes E2).
    await select.selectOption({ label: "Python: Write E2" });
    await runButton.click();
    await expect(runButton).toBeDisabled();
    await expect(runButton).toBeEnabled({ timeout: 90_000 });

    await expect
      .poll(() => page.evaluate(() => (window as any).__formulaApp.getCellValueA1("E2")), {
        timeout: 60_000,
      })
      .toBe("hello from python");
  });
});
