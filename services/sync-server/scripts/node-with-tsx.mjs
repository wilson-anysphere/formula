import { spawn } from "node:child_process";
import { createRequire } from "node:module";

// tsx uses esbuild internally. In pnpm workspaces it's possible to have
// multiple versions of esbuild installed (e.g. vite/vitest vs tsx), and esbuild
// selects its binary via an optionalDependency lookup which can resolve to the
// wrong version. Force esbuild to use the binary that matches the esbuild
// version pinned by tsx.
const require = createRequire(import.meta.url);
const tsxPkgPath = require.resolve("tsx/package.json");
const tsxRequire = createRequire(tsxPkgPath);

const esbuildBinaryPath = tsxRequire.resolve("esbuild/bin/esbuild");

const env = {
  ...process.env,
  ESBUILD_BINARY_PATH: process.env.ESBUILD_BINARY_PATH ?? esbuildBinaryPath,
};

const child = spawn(process.execPath, ["--import", "tsx", ...process.argv.slice(2)], {
  stdio: "inherit",
  env,
});

for (const signal of ["SIGINT", "SIGTERM"]) {
  process.on(signal, () => child.kill(signal));
}

child.on("exit", (code, signal) => {
  if (signal) {
    // Re-raise the signal on this wrapper process so callers (and CI) observe the same
    // termination semantics as the underlying tsx child.
    //
    // Important: do this asynchronously. Some Node versions have been observed to hit
    // an internal assertion (e.g. `env_->execution_async_id() == 0`) if we `process.kill`
    // ourselves synchronously from inside the `exit` event callback.
    setImmediate(() => {
      try {
        process.kill(process.pid, signal);
      } catch {
        process.exit(1);
      }
    });
    return;
  }
  process.exit(code ?? 0);
});
