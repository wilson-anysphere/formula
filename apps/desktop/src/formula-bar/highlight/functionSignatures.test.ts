import { describe, expect, it } from "vitest";

import { getFunctionSignature, signatureParts } from "./functionSignatures.js";

describe("functionSignatures", () => {
  it("signatureParts marks the active parameter", () => {
    const sig = getFunctionSignature("IF");
    expect(sig).toBeTruthy();
    const parts = signatureParts(sig!, 1);
    expect(parts.some((p) => p.kind === "paramActive")).toBe(true);
  });

  it("resolves localized function names to canonical signatures (de-DE SUMME -> SUM)", () => {
    const sig = getFunctionSignature("SUMME", { localeId: "de-DE" });
    expect(sig).toBeTruthy();
    // Displayed name should match the localized formula text.
    expect(sig?.name).toBe("SUMME");
    // Params come from the canonical signature.
    expect(sig?.params[0]?.name).toBe("number1");
  });

  it("supports dotted localized function names (es-ES CONTAR.SI -> COUNTIF)", () => {
    const sig = getFunctionSignature("CONTAR.SI", { localeId: "es-ES" });
    expect(sig).toBeTruthy();
    expect(sig?.name).toBe("CONTAR.SI");
    expect(sig?.params.map((p) => p.name)).toEqual(["range", "criteria"]);
  });

  it("supports localized function names with non-ASCII letters (de-DE ZÄHLENWENN -> COUNTIF)", () => {
    const sig = getFunctionSignature("ZÄHLENWENN", { localeId: "de-DE" });
    expect(sig).toBeTruthy();
    expect(sig?.name).toBe("ZÄHLENWENN");
    expect(sig?.params.map((p) => p.name)).toEqual(["range", "criteria"]);
  });

  it("prefers curated signatures when available (XLOOKUP)", () => {
    const sig = getFunctionSignature("XLOOKUP");
    expect(sig).toBeTruthy();
    expect(sig?.name).toBe("XLOOKUP");
    expect(sig?.params[0]?.name).toBe("lookup_value");
    expect(sig?.summary).toContain("Looks up");
  });

  it("prefers curated signatures when available (SEQUENCE)", () => {
    const sig = getFunctionSignature("SEQUENCE");
    expect(sig).toBeTruthy();
    expect(sig?.name).toBe("SEQUENCE");
    expect(sig?.params[0]?.name).toBe("rows");
    expect(sig?.params[1]?.name).toBe("columns");
  });

  it("preserves _xlfn. prefix in displayed names (Excel compatibility)", () => {
    const sig = getFunctionSignature("_xlfn.sequence");
    expect(sig).toBeTruthy();
    expect(sig?.name).toBe("_XLFN.SEQUENCE");
  });
});
