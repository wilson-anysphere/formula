import { spawnSync } from "node:child_process";
import { fileURLToPath } from "node:url";
import { dirname, resolve } from "node:path";

const __filename = fileURLToPath(import.meta.url);
const __dirname = dirname(__filename);

// `pnpm -C apps/desktop vitest run apps/desktop/src/...` is a common invocation pattern in
// tooling/docs, but when executed from within `apps/desktop` the extra `apps/desktop/` prefix
// prevents Vitest from finding the test file.
//
// Normalize any filter path arguments so both `src/...` and `apps/desktop/src/...` work.
const normalizedArgs = process.argv.slice(2).map((arg) => {
  if (typeof arg !== "string") return arg;
  if (arg.startsWith("apps/desktop/")) return arg.slice("apps/desktop/".length);
  if (arg.startsWith("apps\\desktop\\")) return arg.slice("apps\\desktop\\".length);
  return arg;
});

const vitestBin = resolve(__dirname, "../node_modules/.bin/vitest");
const result = spawnSync(vitestBin, normalizedArgs, { stdio: "inherit" });

if (result.error) {
  // eslint-disable-next-line no-console
  console.error(`Failed to run vitest: ${result.error.message}`);
  process.exit(1);
}

process.exit(result.status ?? 1);

