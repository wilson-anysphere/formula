import { describe, expect, it } from "vitest";

import { getFunctionSignature, signatureParts } from "./functionSignatures.js";

describe("functionSignatures", () => {
  it("signatureParts marks the active parameter", () => {
    const sig = getFunctionSignature("IF");
    expect(sig).toBeTruthy();
    const parts = signatureParts(sig!, 1);
    expect(parts.some((p) => p.kind === "paramActive")).toBe(true);
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

