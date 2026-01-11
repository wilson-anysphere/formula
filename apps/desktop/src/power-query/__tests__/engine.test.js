import assert from "node:assert/strict";
import test from "node:test";

import { createDesktopQueryEngine } from "../engine.ts";

test("createDesktopQueryEngine uses Tauri invoke file commands when FS plugin is unavailable", async () => {
  const originalTauri = globalThis.__TAURI__;

  /** @type {{ cmd: string, args: any }[]} */
  const calls = [];

  globalThis.__TAURI__ = {
    core: {
      invoke: async (cmd, args) => {
        calls.push({ cmd, args });
        if (cmd === "read_text_file") {
          return ["A,B", "1,2"].join("\n");
        }
        if (cmd === "read_binary_file") {
          // Base64 for bytes [1, 2, 3].
          return "AQID";
        }
        throw new Error(`Unexpected invoke: ${cmd}`);
      },
    },
  };

  try {
    const engine = createDesktopQueryEngine();

    const query = {
      id: "q_csv",
      name: "CSV",
      source: { type: "csv", path: "/tmp/test.csv", options: { hasHeaders: true } },
      steps: [],
    };

    const table = await engine.executeQuery(query, {}, {});
    assert.deepEqual(table.toGrid(), [["A", "B"], [1, 2]]);

    const bytes = await engine.fileAdapter.readBinary("/tmp/test.bin");
    assert.deepEqual(Array.from(bytes), [1, 2, 3]);

    assert.ok(calls.some((c) => c.cmd === "read_text_file"));
    assert.ok(calls.some((c) => c.cmd === "read_binary_file"));
  } finally {
    globalThis.__TAURI__ = originalTauri;
  }
});

