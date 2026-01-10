import { describe, expect, it } from "vitest";

import { formulaFiles } from "@formula/python-runtime";

describe("python runtime bundled formula API files", () => {
  it("includes core formula modules needed for Pyodide installation", () => {
    expect(formulaFiles).toBeTruthy();
    expect(formulaFiles).toHaveProperty("formula/__init__.py");
    expect(formulaFiles).toHaveProperty("formula/_a1.py");
    expect(formulaFiles).toHaveProperty("formula/_bridge.py");
    expect(formulaFiles).toHaveProperty("formula/_js_bridge.py");
    expect(formulaFiles).toHaveProperty("formula/runtime/sandbox.py");
  });
});
