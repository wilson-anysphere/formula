import { createRequire } from "node:module";

/**
 * Load the CommonJS build of Yjs from an ESM test file.
 *
 * This is used by tests that simulate environments like y-websocket where
 * updates and/or root types can come from a different module loader instance
 * of Yjs than the one used by the code under test.
 *
 * @returns {typeof import("yjs")}
 */
export function requireYjsCjs() {
  const require = createRequire(import.meta.url);
  const prevError = console.error;
  console.error = (...args) => {
    if (typeof args[0] === "string" && args[0].startsWith("Yjs was already imported.")) return;
    prevError(...args);
  };
  try {
    // eslint-disable-next-line import/no-named-as-default-member
    return require("yjs");
  } finally {
    console.error = prevError;
  }
}
