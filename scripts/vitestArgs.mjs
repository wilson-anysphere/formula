/**
 * Normalize CLI arguments forwarded from `pnpm test:vitest` to `vitest`.
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
  let args = rawArgs.slice();
  // Strip only the first occurrence so callers can still pass a literal `--` later
  // if they really need to.
  const delimiterIdx = args.indexOf("--");
  if (delimiterIdx >= 0) {
    args = [...args.slice(0, delimiterIdx), ...args.slice(delimiterIdx + 1)];
  }

  return args.map((arg) => {
    if (arg === "--silent") return "--silent=true";
    return arg;
  });
}
