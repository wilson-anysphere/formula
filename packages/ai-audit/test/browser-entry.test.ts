import { describe, expect, it } from "vitest";

import { spawnSync } from "node:child_process";
import { createRequire } from "node:module";
import path from "node:path";
import { fileURLToPath } from "node:url";

const require = createRequire(import.meta.url);

function supportsRegister(): boolean {
  try {
    return typeof (require("node:module") as any)?.register === "function";
  } catch {
    return false;
  }
}

function resolveNodeLoaderArgs(loaderUrl: string): string[] {
  const allowedFlags =
    process.allowedNodeEnvironmentFlags && typeof process.allowedNodeEnvironmentFlags.has === "function"
      ? process.allowedNodeEnvironmentFlags
      : new Set<string>();

  if (supportsRegister() && allowedFlags.has("--import")) {
    const registerScript = `import { register } from "node:module"; register(${JSON.stringify(loaderUrl)});`;
    const dataUrl = `data:text/javascript;base64,${Buffer.from(registerScript, "utf8").toString("base64")}`;
    return ["--import", dataUrl];
  }

  if (allowedFlags.has("--loader")) return ["--loader", loaderUrl];
  if (allowedFlags.has("--experimental-loader")) return [`--experimental-loader=${loaderUrl}`];
  return [];
}

describe("@formula/ai-audit browser entrypoint", () => {
  it("imports without Node-only globals (process.versions.node, Buffer)", () => {
    // Use the repo's shared TS loader (used by the node:test runner) so this suite
    // stays in sync with how we execute workspace TS sources under Node.
    const loaderUrl = new URL("../../../scripts/resolve-ts-loader.mjs", import.meta.url);

    const result = spawnSync(
      process.execPath,
      [
        "--no-warnings",
        ...resolveNodeLoaderArgs(loaderUrl.href),
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

  it("imports the export entrypoint without Node-only globals (process.versions.node, Buffer)", () => {
    const loaderUrl = new URL("../../../scripts/resolve-ts-loader.mjs", import.meta.url);

    const result = spawnSync(
      process.execPath,
      [
        "--no-warnings",
        ...resolveNodeLoaderArgs(loaderUrl.href),
        "--input-type=module",
        "--eval",
        `
          Object.defineProperty(process.versions, "node", { value: undefined, configurable: true });
          globalThis.Buffer = undefined;
          import { pathToFileURL } from "node:url";
          import { resolve } from "node:path";
          await import(pathToFileURL(resolve("packages/ai-audit/src/export.ts")).href);
        `
      ],
      { encoding: "utf8" }
    );

    expect(
      result.status,
      `child process failed:\nstdout:\n${result.stdout}\nstderr:\n${result.stderr}`
    ).toBe(0);
  });

  it("imports @formula/ai-audit/export via package exports without Node-only globals (process.versions.node, Buffer)", () => {
    const loaderUrl = new URL("../../../scripts/resolve-ts-loader.mjs", import.meta.url);
    const pkgRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");

    const result = spawnSync(
      process.execPath,
      [
        "--no-warnings",
        ...resolveNodeLoaderArgs(loaderUrl.href),
        "--input-type=module",
        "--eval",
        `
          Object.defineProperty(process.versions, "node", { value: undefined, configurable: true });
          globalThis.Buffer = undefined;
          await import("@formula/ai-audit/export");
        `
      ],
      { encoding: "utf8", cwd: pkgRoot }
    );

    expect(
      result.status,
      `child process failed:\nstdout:\n${result.stdout}\nstderr:\n${result.stderr}`
    ).toBe(0);
  });
});
