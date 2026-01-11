import { describe, expect, it } from "vitest";

import { ScriptRuntime, Workbook } from "../src/node.js";

function makeWorkbook() {
  const workbook = new Workbook();
  workbook.addSheet("Sheet1");
  workbook.addSheet("Sheet2");
  workbook.setActiveSheet("Sheet1");
  return workbook;
}

describe("ScriptRuntime", () => {
  it("runs a basic script and writes values", async () => {
    const workbook = makeWorkbook();
    const runtime = new ScriptRuntime(workbook);

    const result = await runtime.run(`
await ctx.activeSheet.getRange("A1:B1").setValues([[1, 2]]);
ctx.ui.log("done");
`);

    expect(result.error).toBeUndefined();
    expect(workbook.getActiveSheet().getRange("A1:B1").getValues()).toEqual([[1, 2]]);
    expect(result.logs.map((l) => l.message)).toContain("done");
  });

  it("supports module-style scripts that export a default function", async () => {
    const workbook = makeWorkbook();
    const runtime = new ScriptRuntime(workbook);

    const result = await runtime.run(`
export default async function main(ctx: ScriptContext) {
  await ctx.activeSheet.getRange("A1").setValue(99);
  ctx.ui.log("module ok");
}
`);

    expect(result.error).toBeUndefined();
    expect(workbook.getActiveSheet().getRange("A1").getValue()).toBe(99);
    expect(result.logs.map((l) => l.message)).toContain("module ok");
  });

  it("enforces network permissions (fetch denied by default)", async () => {
    const workbook = makeWorkbook();
    const runtime = new ScriptRuntime(workbook);

    const result = await runtime.run(`await fetch("https://example.com");`);

    expect(result.error).toBeDefined();
    expect(result.error?.name).toBe("PermissionDeniedError");
    expect(result.error?.code).toBe("PERMISSION_DENIED");
    expect(result.error?.message).toMatch(/Network access denied/i);
  });

  it("enforces filesystem permissions (fs.readFile denied by default)", async () => {
    const workbook = makeWorkbook();
    const runtime = new ScriptRuntime(workbook);

    const result = await runtime.run(`await fs.readFile("/tmp/forbidden.txt", "utf8");`);

    expect(result.error).toBeDefined();
    expect(result.error?.name).toBe("PermissionDeniedError");
    expect(result.error?.code).toBe("PERMISSION_DENIED");
    expect(result.error?.message).toMatch(/Filesystem read access denied/i);
  });

  it("surfaces TypeScript diagnostics", async () => {
    const workbook = makeWorkbook();
    const runtime = new ScriptRuntime(workbook);

    const result = await runtime.run(`
// Syntax error
const x: number = ;
`);

    expect(result.error).toBeDefined();
    expect(result.error?.name).toBe("TypeScriptCompileError");
    expect(result.error?.message).toMatch(/user-script\.ts/i);
  });

  it("delivers events in order", async () => {
    const workbook = makeWorkbook();
    const runtime = new ScriptRuntime(workbook);

    const result = await runtime.run(`
const seen: string[] = [];

ctx.events.onSelectionChange((evt) => {
  seen.push("sel:" + evt.address);
});
ctx.events.onEdit((evt) => {
  seen.push("edit:" + evt.address);
});

await ctx.workbook.setSelection("Sheet1", "A1");
await ctx.activeSheet.getRange("A1").setValue(123);
await ctx.workbook.setSelection("Sheet1", "B2");

await ctx.events.flush();
ctx.ui.log(JSON.stringify(seen));
`);

    expect(result.error).toBeUndefined();
    const last = result.logs.at(-1)?.message ?? "";
    expect(JSON.parse(last)).toEqual(["sel:A1", "edit:A1", "sel:B2"]);
  });

  it("supports getUsedRange/getSheets + formulas + formats", async () => {
    const workbook = makeWorkbook();
    const runtime = new ScriptRuntime(workbook);

    const result = await runtime.run(`
const sheets = await ctx.workbook.getSheets();
ctx.ui.log("sheets=" + sheets.map((s) => s.name).join(","));

await sheets[0].getRange("C3").setValue(42);
await sheets[0].getRange("B2").setFormulas([["=SUM(1,2)"]]);
await sheets[0].getRange("A1:B1").setFormats([[{ bold: true }, { italic: true }]]);

const used = await sheets[0].getUsedRange();
ctx.ui.log("used=" + used.address);
`);

    expect(result.error).toBeUndefined();
    expect(workbook.getSheet("Sheet1").getUsedRange().address).toBe("A1:C3");
    expect(workbook.getSheet("Sheet1").getRange("B2").getFormulas()).toEqual([["=SUM(1,2)"]]);
    expect(workbook.getSheet("Sheet1").getRange("A1:B1").getFormats()).toEqual([
      [{ bold: true }, { italic: true }],
    ]);
    expect(result.logs.some((l) => l.message.startsWith("used="))).toBe(true);
  });

  it(
    "supports cancellation via AbortSignal",
    { timeout: 10_000 },
    async () => {
      const workbook = makeWorkbook();
      const runtime = new ScriptRuntime(workbook);

      const controller = new AbortController();
      const promise = runtime.run(`await new Promise(() => {});`, { signal: controller.signal, timeoutMs: 5_000 });
      controller.abort();

      const result = await promise;
      expect(result.error).toBeDefined();
      expect(result.error?.name).toBe("AbortError");
    },
  );

  it(
    "configures sandbox memory limits (best-effort)",
    { timeout: 10_000 },
    async () => {
      const workbook = makeWorkbook();
      const runtime = new ScriptRuntime(workbook);

      const result = await runtime.run(`ctx.ui.log("ok");`, { memoryMb: 64 });

      expect(result.error).toBeUndefined();
      const spawn = result.audit.find((e) => e.eventType === "scripting.sandbox.spawn");
      expect(spawn).toBeDefined();
      expect(spawn.metadata.memoryMb).toBe(64);
      expect(spawn.metadata.resourceLimits?.maxOldGenerationSizeMb).toBeGreaterThanOrEqual(64);
    },
  );
});
