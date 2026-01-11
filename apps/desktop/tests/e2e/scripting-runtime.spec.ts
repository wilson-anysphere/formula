import { expect, test } from "@playwright/test";

test("scripting: runs TypeScript in a WebWorker with RPC + network sandbox", async ({ page }) => {
  await page.goto("/scripting-test.html");
  await page.waitForFunction(() => Boolean((globalThis as any).__formulaScripting));

  const runInPage = async () => {
    const { ScriptRuntime, Workbook } = (globalThis as any).__formulaScripting;

    const workbook = new Workbook();
    workbook.addSheet("Sheet1");
    workbook.setActiveSheet("Sheet1");
    const sheet = workbook.getActiveSheet();

    sheet.setCellValue("A1", 10);
    sheet.setCellValue("B1", 32);

    const runtime = new ScriptRuntime(workbook);

    const mainRun = await runtime.run(`
await ctx.workbook.setSelection("Sheet1", "A1:B1");
const values = await ctx.activeSheet.getRange("A1:B1").getValues();
await ctx.activeSheet.getRange("C1").setValue(values[0][0] + values[0][1]);
await ctx.activeSheet.getRange("A1:B1").setFormat({ bold: true });
ctx.ui.log("sum", values[0][0] + values[0][1]);
`);

    const computed = sheet.getRange("C1").getValue();
    const format = sheet.getRange("A1").getFormat();
    const selection = workbook.getSelection();

    const blockedNetwork = await runtime.run(`
await fetch("https://example.com");
`);

    const allowlistedNetwork = await runtime.run(
      `
const res = await fetch("/scripting-test.html");
ctx.ui.log("status", res.status);
`,
      {
        permissions: { network: "allowlist", networkAllowlist: ["localhost"] },
      },
    );

    return { mainRun, computed, format, selection, blockedNetwork, allowlistedNetwork };
  };

  let result;
  // Vite may trigger a one-time full reload after dependency optimization. If
  // that happens mid-evaluate, retry once after the navigation completes.
  for (let attempt = 0; attempt < 2; attempt += 1) {
    try {
      result = await page.evaluate(runInPage);
      break;
    } catch (err) {
      const message = err instanceof Error ? err.message : String(err);
      if (attempt === 0 && message.includes("Execution context was destroyed")) {
        await page.waitForLoadState("load");
        await page.waitForFunction(() => Boolean((globalThis as any).__formulaScripting));
        continue;
      }
      throw err;
    }
  }

  expect(result).toBeTruthy();

  expect(result.mainRun.error).toBeUndefined();
  expect(result.computed).toBe(42);
  expect(result.format).toEqual({ bold: true });
  expect(result.selection).toEqual({ sheetName: "Sheet1", address: "A1:B1" });

  expect(result.blockedNetwork.error?.message).toContain("Network access");
  expect(result.allowlistedNetwork.error).toBeUndefined();
  expect(result.allowlistedNetwork.logs.some((entry) => entry.message.includes("status"))).toBe(true);
});
