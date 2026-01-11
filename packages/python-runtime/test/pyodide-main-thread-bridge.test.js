import test from "node:test";
import assert from "node:assert/strict";

import { registerFormulaBridge, setFormulaBridgeApi } from "../src/pyodide-main-thread.js";

test("pyodide main-thread formula_bridge reads API from the runtime instance and can be swapped", () => {
  /** @type {Record<string, any>} */
  const registered = {};

  const runtime = {
    registerJsModule(name, mod) {
      registered[name] = mod;
    },
  };

  setFormulaBridgeApi(runtime, {
    get_active_sheet_id(_params) {
      return "Sheet1";
    },
  });

  registerFormulaBridge(runtime);
  assert.ok(registered.formula_bridge);
  assert.equal(registered.formula_bridge.get_active_sheet_id(), "Sheet1");

  // Swap the API and ensure the previously-registered module uses the new API.
  setFormulaBridgeApi(runtime, {
    get_active_sheet_id(_params) {
      return "Sheet2";
    },
  });

  assert.equal(registered.formula_bridge.get_active_sheet_id(), "Sheet2");
});

