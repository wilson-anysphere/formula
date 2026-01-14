import { describe, expect, it } from "vitest";

import {
  isMissingGetCellPhoneticError,
  isMissingGetLocaleInfoError,
  isMissingGetRangeCompactError,
  isMissingSetCellPhoneticError,
  isMissingSupportedLocaleIdsError
} from "../compat.ts";

describe("isMissingGetRangeCompactError", () => {
  it("matches unknown RPC method errors", () => {
    expect(isMissingGetRangeCompactError(new Error("unknown method: getRangeCompact"))).toBe(true);
  });

  it("matches missing WASM export errors", () => {
    expect(
      isMissingGetRangeCompactError(new Error("getRangeCompact: WasmWorkbook.getRangeCompact is not available in this WASM build"))
    ).toBe(true);
  });

  it("does not match unrelated errors", () => {
    expect(isMissingGetRangeCompactError(new Error("boom"))).toBe(false);
    expect(isMissingGetRangeCompactError("boom")).toBe(false);
  });
});

describe("isMissingSupportedLocaleIdsError", () => {
  it("matches unknown RPC method errors", () => {
    expect(isMissingSupportedLocaleIdsError(new Error("unknown method: supportedLocaleIds"))).toBe(true);
  });

  it("matches missing WASM export errors", () => {
    expect(
      isMissingSupportedLocaleIdsError(new Error("supportedLocaleIds: wasm module does not export supportedLocaleIds()"))
    ).toBe(true);
  });

  it("does not match unrelated errors", () => {
    expect(isMissingSupportedLocaleIdsError(new Error("boom"))).toBe(false);
    expect(isMissingSupportedLocaleIdsError("boom")).toBe(false);
  });
});

describe("isMissingGetLocaleInfoError", () => {
  it("matches unknown RPC method errors", () => {
    expect(isMissingGetLocaleInfoError(new Error("unknown method: getLocaleInfo"))).toBe(true);
  });

  it("matches missing WASM export errors", () => {
    expect(isMissingGetLocaleInfoError(new Error("getLocaleInfo: wasm module does not export getLocaleInfo()"))).toBe(
      true
    );
  });

  it("does not match unrelated errors", () => {
    expect(isMissingGetLocaleInfoError(new Error("boom"))).toBe(false);
    expect(isMissingGetLocaleInfoError("boom")).toBe(false);
  });
});

describe("isMissingGetCellPhoneticError", () => {
  it("matches unknown RPC method errors", () => {
    expect(isMissingGetCellPhoneticError(new Error("unknown method: getCellPhonetic"))).toBe(true);
  });

  it("matches missing WASM export errors", () => {
    expect(
      isMissingGetCellPhoneticError(
        new Error("getCellPhonetic: WasmWorkbook.getCellPhonetic is not available in this WASM build")
      )
    ).toBe(true);
  });

  it("does not match unrelated errors", () => {
    expect(isMissingGetCellPhoneticError(new Error("boom"))).toBe(false);
    expect(isMissingGetCellPhoneticError("boom")).toBe(false);
  });
});

describe("isMissingSetCellPhoneticError", () => {
  it("matches unknown RPC method errors", () => {
    expect(isMissingSetCellPhoneticError(new Error("unknown method: setCellPhonetic"))).toBe(true);
  });

  it("matches missing WASM export errors", () => {
    expect(
      isMissingSetCellPhoneticError(
        new Error("setCellPhonetic: WasmWorkbook.setCellPhonetic is not available in this WASM build")
      )
    ).toBe(true);
  });

  it("does not match unrelated errors", () => {
    expect(isMissingSetCellPhoneticError(new Error("boom"))).toBe(false);
    expect(isMissingSetCellPhoneticError("boom")).toBe(false);
  });
});
