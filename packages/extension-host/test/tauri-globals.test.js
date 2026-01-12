const test = require("node:test");
const assert = require("node:assert/strict");

test("browser sandbox: tauri globals are locked down", async () => {
  const { TAURI_GLOBALS, lockDownTauriGlobals } = await import("../src/browser/tauri-globals.mjs");

  // Simulate a Tauri environment injecting these globals.
  for (const prop of TAURI_GLOBALS) {
    try {
      Object.defineProperty(globalThis, prop, {
        value: { injected: true },
        writable: true,
        configurable: true,
        enumerable: true
      });
    } catch {
      globalThis[prop] = { injected: true };
    }
  }

  /** @type {string[]} */
  const called = [];
  const lockDownGlobal = (prop, value) => {
    called.push(prop);
    Object.defineProperty(globalThis, prop, {
      value,
      writable: false,
      configurable: false,
      enumerable: true
    });
  };

  lockDownTauriGlobals(lockDownGlobal);

  assert.deepEqual(called.sort(), [...TAURI_GLOBALS].sort());

  for (const prop of TAURI_GLOBALS) {
    assert.equal(globalThis[prop], undefined);
    const desc = Object.getOwnPropertyDescriptor(globalThis, prop);
    assert.ok(desc, `expected ${prop} to be defined`);
    assert.equal(desc.value, undefined);
    assert.equal(desc.writable, false);
    assert.equal(desc.configurable, false);
  }
});

