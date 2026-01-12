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

test("pyodide main-thread formula_bridge forwards create_sheet index when provided", () => {
  /** @type {Record<string, any>} */
  const registered = {};

  const runtime = {
    registerJsModule(name, mod) {
      registered[name] = mod;
    },
  };

  /** @type {any} */
  let lastParams = null;

  setFormulaBridgeApi(runtime, {
    create_sheet(params) {
      lastParams = params;
      return "sheet_new";
    },
  });

  registerFormulaBridge(runtime);
  assert.ok(registered.formula_bridge);

  // Without an index, the params should match the legacy shape.
  assert.equal(registered.formula_bridge.create_sheet("NoIndex"), "sheet_new");
  assert.deepEqual(lastParams, { name: "NoIndex" });

  // With an index, it should be forwarded to the host API.
  assert.equal(registered.formula_bridge.create_sheet("WithIndex", 2), "sheet_new");
  assert.deepEqual(lastParams, { name: "WithIndex", index: 2 });
});
