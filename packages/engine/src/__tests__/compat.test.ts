import { describe, expect, it } from "vitest";

import { isMissingGetRangeCompactError } from "../compat.ts";

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

