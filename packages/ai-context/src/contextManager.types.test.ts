import { describe, expect, it } from "vitest";

import { ContextManager, type BuildContextResult, type SheetSchema } from "./index.js";

// --- Compile-time assertions -------------------------------------------------
// These types intentionally produce a TypeScript error if the referenced type is `any`.
type IsAny<T> = 0 extends 1 & T ? true : false;
type Assert<T extends true> = T;

type _SchemaIsSheetSchema = Assert<BuildContextResult["schema"] extends SheetSchema ? true : false>;
type _SchemaNotAny = Assert<IsAny<BuildContextResult["schema"]> extends false ? true : false>;
type _RetrievedNotAny = Assert<IsAny<BuildContextResult["retrieved"][number]> extends false ? true : false>;
type _DlpOptionsNotAny = Assert<
  IsAny<NonNullable<Parameters<ContextManager["buildContext"]>[0]["dlp"]>> extends false ? true : false
>;

describe("ContextManager types", () => {
  it("buildContext().schema is a SheetSchema at runtime (and not `any` at compile time)", async () => {
    const cm = new ContextManager({
      tokenBudgetTokens: 10_000,
      // Avoid redaction so prompt strings are stable for snapshots/debugging.
      redactor: (text: string) => text,
    });

    const result = await cm.buildContext({
      sheet: { name: "Sheet1", values: [["A"], ["B"]] },
      query: "A",
    });

    // Runtime sanity check.
    expect(result.schema.name).toBe("Sheet1");

    // Compile-time check (reinforces the intent of this test file).
    const _schema: SheetSchema = result.schema;
    expect(_schema.tables).toBeDefined();
  });
});

