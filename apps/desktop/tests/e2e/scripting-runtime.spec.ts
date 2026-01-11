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
export default async function main(ctx) {
  await ctx.workbook.setSelection("Sheet1", "A1:B1");
  const values = await ctx.activeSheet.getRange("A1:B1").getValues();
  await ctx.activeSheet.getRange("C1").setValue(values[0][0] + values[0][1]);
  await ctx.activeSheet.getRange("A1:B1").setFormat({ bold: true });
  ctx.ui.log("sum", values[0][0] + values[0][1]);
}
`);

    const computed = sheet.getRange("C1").getValue();
    const format = sheet.getRange("A1").getFormat();
    const selection = workbook.getSelection();

    const blockedNetwork = await runtime.run(`
export default async function main(ctx) {
  await ctx.fetch("https://example.com");
}
`);

    const allowlistedNetwork = await runtime.run(
      `
export default async function main(ctx) {
  const res = await ctx.fetch("/scripting-test.html");
  ctx.ui.log("status", res.status);
}
`,
      {
        permissions: { network: "allowlist", networkAllowlist: ["localhost"] },
      },
    );

    const allowlistDenied = await runtime.run(
      `
export default async function main(ctx) {
  await ctx.fetch("https://example.com");
}
`,
      {
        permissions: { network: "allowlist", networkAllowlist: ["localhost"] },
      },
    );

    const allowlistWebSocketDenied = await runtime.run(
      `
export default async function main(ctx) {
  new WebSocket("wss://example.com");
}
`,
      {
        permissions: { network: "allowlist", networkAllowlist: ["localhost"] },
      },
    );

    const dynamicImportDenied = await runtime.run(`
// Dynamic import is intentionally unsupported (it could otherwise bypass fetch/WebSocket sandboxing).
await import("https://example.com");
`);

    return {
      mainRun,
      computed,
      format,
      selection,
      blockedNetwork,
      allowlistedNetwork,
      allowlistDenied,
      allowlistWebSocketDenied,
      dynamicImportDenied,
    };
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

  expect(result.allowlistDenied.error?.message).toContain("example.com");
  expect(result.allowlistWebSocketDenied.error?.message).toContain("example.com");
  expect(result.dynamicImportDenied.error?.message).toContain("dynamic import");
  expect(result.dynamicImportDenied.error?.message).toContain("example.com");
});

test("scripting: times out hung scripts and ignores spoofed worker messages", async ({ page }) => {
  await page.goto("/scripting-test.html");
  await page.waitForFunction(() => Boolean((globalThis as any).__formulaScripting));

  const runInPage = async () => {
    const { ScriptRuntime, Workbook } = (globalThis as any).__formulaScripting;

    const workbook = new Workbook();
    workbook.addSheet("Sheet1");
    workbook.setActiveSheet("Sheet1");
    const sheet = workbook.getActiveSheet();
    sheet.setCellValue("A1", 10);

    const runtime = new ScriptRuntime(workbook);

    const simpleTimeout = await runtime.run(`await new Promise(() => {});`, { timeoutMs: 200 });

    const spoofed = await runtime.run(
      `
// Attempt to spoof the host protocol: should be ignored without the per-run token.
self.postMessage({
  type: "rpc",
  id: 123,
  method: "range.setValue",
  params: { sheetName: "Sheet1", address: "A1", value: 99 },
});
self.postMessage({ type: "result" });
await new Promise(() => {});
`,
      { timeoutMs: 200 },
    );

    return {
      simpleTimeout,
      spoofed,
      cellValue: sheet.getRange("A1").getValue(),
    };
  };

  const result = await page.evaluate(runInPage);
  expect(result.simpleTimeout.error?.message).toContain("timed out");
  expect(result.spoofed.error?.message).toContain("timed out");
  expect(result.cellValue).toBe(10);
});

test("scripting: forwards ctx.confirm/prompt/alert via RPC", async ({ page }) => {
  await page.goto("/scripting-test.html");
  await page.waitForFunction(() => Boolean((globalThis as any).__formulaScripting));

  const dialogs: Array<{ type: string; message: string }> = [];
  page.on("dialog", async (dialog) => {
    dialogs.push({ type: dialog.type(), message: dialog.message() });
    if (dialog.type() === "prompt") {
      await dialog.accept("Alice");
      return;
    }
    // confirm + alert
    await dialog.accept();
  });

  const result = await page.evaluate(async () => {
    const { ScriptRuntime, Workbook } = (globalThis as any).__formulaScripting;

    const workbook = new Workbook();
    workbook.addSheet("Sheet1");
    workbook.setActiveSheet("Sheet1");

    const runtime = new ScriptRuntime(workbook);
    return await runtime.run(`
export default async function main(ctx) {
  const ok = await ctx.confirm("Proceed?");
  const name = await ctx.prompt("Name?", "Unknown");
  await ctx.alert("done");
  ctx.ui.log("confirm", ok, "prompt", name);
}
`);
  });

  expect(result.error).toBeUndefined();
  expect(result.logs.some((entry) => entry.message.includes("confirm"))).toBe(true);
  expect(dialogs.map((d) => d.type)).toEqual(["confirm", "prompt", "alert"]);
  expect(dialogs[0]?.message).toBe("Proceed?");
  expect(dialogs[1]?.message).toBe("Name?");
});
