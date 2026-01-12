import { expect, test } from "@playwright/test";

import { gotoDesktop } from "./helpers";

test.describe("Page Layout print/export wiring (tauri)", () => {
  test("Page Setup, print area, and export PDF invoke the expected Tauri commands", async ({ page }) => {
    await page.addInitScript(() => {
      const listeners: Record<string, Array<(event: any) => void>> = {};
      const invokes: Array<{ cmd: string; args: any }> = [];
      const downloadClicks: Array<{ download: string | null; href: string | null }> = [];

      (window as any).__tauriListeners = listeners;
      (window as any).__tauriInvokes = invokes;
      (window as any).__downloadClicks = downloadClicks;

      // Avoid modal prompts blocking the test.
      window.confirm = () => true;

      // Avoid creating real browser downloads; capture intent instead.
      HTMLAnchorElement.prototype.click = function click() {
        try {
          downloadClicks.push({
            download: typeof (this as any).download === "string" ? (this as any).download : null,
            href: typeof (this as any).href === "string" ? (this as any).href : null,
          });
        } catch {
          // ignore
        }
      };

      const windowHandle = {
        hide: async () => {},
        close: async () => {},
        show: async () => {},
        setFocus: async () => {},
      };

      (window as any).__TAURI__ = {
        core: {
          invoke: async (cmd: string, args: any) => {
            invokes.push({ cmd, args });

            switch (cmd) {
              case "get_sheet_print_settings":
                return {
                  sheet_name: args?.sheet_id ?? "Sheet1",
                  print_area: null,
                  print_titles: null,
                  page_setup: {
                    orientation: "portrait",
                    paper_size: 1,
                    margins: { left: 0.7, right: 0.7, top: 0.75, bottom: 0.75, header: 0.3, footer: 0.3 },
                    scaling: { kind: "percent", percent: 100 },
                  },
                  manual_page_breaks: { row_breaks_after: [], col_breaks_after: [] },
                };

              case "set_sheet_page_setup":
              case "set_sheet_print_area":
                return null;

              case "export_sheet_range_pdf":
                // A tiny placeholder payload; need not be a valid PDF for frontend wiring tests.
                return btoa("pdf");

              // Minimal workbook backend stubs so the app can boot in \"tauri\" mode.
              case "list_defined_names":
              case "list_tables":
              case "get_workbook_theme_palette":
              case "get_macro_security_status":
              case "set_macro_ui_context":
              case "fire_workbook_open":
                return null;

              default:
                // Best-effort: ignore unrelated invocations so new backend calls don't break the test.
                return null;
            }
          },
        },
        event: {
          listen: async (name: string, handler: any) => {
            if (!Array.isArray(listeners[name])) listeners[name] = [];
            listeners[name].push(handler);
            return () => {
              const arr = listeners[name];
              if (!Array.isArray(arr)) return;
              const idx = arr.indexOf(handler);
              if (idx >= 0) arr.splice(idx, 1);
            };
          },
          emit: async () => {},
        },
        dialog: {
          open: async () => null,
          save: async () => null,
        },
        window: {
          // Provide all common handle accessors used by our Tauri abstractions so this
          // test stays resilient to future refactors.
          getCurrentWebviewWindow: () => windowHandle,
          getCurrentWindow: () => windowHandle,
          getCurrent: () => windowHandle,
          appWindow: windowHandle,
        },
        notification: {
          notify: async () => {},
        },
      };
    });

    await gotoDesktop(page);

    const ribbon = page.getByTestId("ribbon-root");
    await expect(ribbon).toBeVisible();

    // Ensure we have a stable selection for the set print area/export assertions.
    await expect(page.getByTestId("active-cell")).toHaveText("A1");

    await ribbon.getByRole("tab", { name: "Page Layout" }).click();

    // --- Page Setup dialog wiring ---------------------------------------------
    await ribbon.getByTestId("ribbon-page-setup").click();
    await expect(page.locator("dialog.page-setup-dialog")).toBeVisible();

    // Flip the orientation to trigger a backend update.
    await page.locator("dialog.page-setup-dialog select").first().selectOption("landscape");
    await page.waitForFunction(() => (window as any).__tauriInvokes?.some((e: any) => e?.cmd === "set_sheet_page_setup"));

    await page.locator("dialog.page-setup-dialog").getByRole("button", { name: "Close" }).click();
    await expect(page.locator("dialog.page-setup-dialog")).toHaveCount(0);

    // --- Print area wiring -----------------------------------------------------
    await ribbon.getByTestId("ribbon-set-print-area").click();
    await page.waitForFunction(
      () => (window as any).__tauriInvokes?.some((e: any) => e?.cmd === "set_sheet_print_area" && Array.isArray(e?.args?.print_area)),
    );

    await ribbon.getByTestId("ribbon-clear-print-area").click();
    await page.waitForFunction(
      () => (window as any).__tauriInvokes?.some((e: any) => e?.cmd === "set_sheet_print_area" && e?.args?.print_area === null),
    );

    // --- Export PDF wiring -----------------------------------------------------
    await ribbon.getByTestId("ribbon-export-pdf").click();
    await page.waitForFunction(() => (window as any).__tauriInvokes?.some((e: any) => e?.cmd === "export_sheet_range_pdf"));

    const exportCall = await page.evaluate(() => {
      const invokes = (window as any).__tauriInvokes as Array<{ cmd: string; args: any }> | undefined;
      if (!Array.isArray(invokes)) return null;
      return invokes.find((entry) => entry.cmd === "export_sheet_range_pdf") ?? null;
    });
    expect(exportCall).not.toBeNull();
    expect(exportCall!.args?.range?.start_row).toBe(1);
    expect(exportCall!.args?.range?.start_col).toBe(1);
    expect(exportCall!.args?.range?.end_row).toBe(1);
    expect(exportCall!.args?.range?.end_col).toBe(1);

    // Ensure the frontend attempted to start a download with a PDF filename.
    await page.waitForFunction(() => Array.isArray((window as any).__downloadClicks) && (window as any).__downloadClicks.length > 0);
    const download = await page.evaluate(() => (window as any).__downloadClicks?.[(window as any).__downloadClicks.length - 1] ?? null);
    expect(download?.download).toMatch(/\.pdf$/i);
  });
});
