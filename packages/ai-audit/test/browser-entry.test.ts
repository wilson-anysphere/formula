import { describe, expect, it } from "vitest";

import { spawnSync } from "node:child_process";

describe("@formula/ai-audit browser entrypoint", () => {
  it("imports without Node-only globals (process.versions.node, Buffer)", () => {
    const loaderUrl = new URL("./resolve-ts-loader.mjs", import.meta.url);

    const result = spawnSync(
      process.execPath,
      [
        "--no-warnings",
        `--experimental-loader=${loaderUrl.href}`,
        "--input-type=module",
        "--eval",
        `
          Object.defineProperty(process.versions, "node", { value: undefined, configurable: true });
          globalThis.Buffer = undefined;
          import { pathToFileURL } from "node:url";
          import { resolve } from "node:path";
          await import(pathToFileURL(resolve("packages/ai-audit/src/index.ts")).href);
        `
      ],
      { encoding: "utf8" }
    );

    expect(
      result.status,
      `child process failed:\nstdout:\n${result.stdout}\nstderr:\n${result.stderr}`
    ).toBe(0);
  });
});
