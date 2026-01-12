/**
 * Normalize CLI arguments forwarded from `pnpm test:vitest` to `vitest run`.
 *
 * Why this exists:
 * - pnpm will forward a literal `--` delimiter through to scripts when callers include it
 *   (npm/yarn muscle memory). Vitest treats a bare `--` as a test pattern, which can
 *   accidentally cause the full suite to run.
 * - Vitest's CLI treats `--silent <pattern>` as "silent has value <pattern>" and errors;
 *   use the explicit boolean form so `pnpm test:vitest --silent <file>` works.
 *
 * @param {string[]} rawArgs
 * @returns {string[]}
 */
export function normalizeVitestArgs(rawArgs) {
  /** @type {string[]} */
  const args = [];
  for (const arg of rawArgs) {
    if (arg === "--") continue;
    if (arg === "--silent") {
      args.push("--silent=true");
      continue;
    }
    args.push(arg);
  }
  return args;
}

