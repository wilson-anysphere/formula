import { expect, test } from "@playwright/test";

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

    await page.goto("/");
    await page.evaluate(() => localStorage.clear());
    await page.reload();

    await page.getByTestId("open-macros-panel").click();
    const panel = page.getByTestId("dock-right").getByTestId("panel-macros");
    await expect(panel).toBeVisible();

    await expect(panel.locator("select")).toBeVisible();
    await panel.getByRole("button", { name: "Run" }).click();

    await expect(page.getByTestId("active-value")).toHaveText("FromMacro");
    await page.waitForFunction(async () => (await (window as any).__formulaApp.getCellValueA1("A2")) === "MacroA2");

    // Focus the grid to ensure keyboard shortcuts route to the SpreadsheetApp handler.
    await page.click("#grid", { position: { x: 5, y: 5 } });

    const modifier = process.platform === "darwin" ? "Meta" : "Control";
    await page.keyboard.press(`${modifier}+Z`);

    await expect(page.getByTestId("active-value")).toHaveText("Seed");
    const a2 = await page.evaluate(() => (window as any).__formulaApp.getCellValueA1("A2"));
    expect(a2).toBe("A");
  });
});
