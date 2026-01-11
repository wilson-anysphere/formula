import path from "node:path";
import { pathToFileURL } from "node:url";

const importEsm: (specifier: string) => Promise<any> = new Function(
  "specifier",
  "return import(specifier)"
) as unknown as (specifier: string) => Promise<any>;

function candidatesForRepoRelativePath(repoRelativePath: string): string[] {
  const candidates: string[] = [];

  // When running from compiled JS, `__dirname` typically points at
  //   services/api/dist/crypto
  // When running from tsx in dev, `__dirname` points at
  //   services/api/src/crypto
  // Both are 4 levels below repo root.
  if (typeof __dirname === "string") {
    candidates.push(pathToFileURL(path.resolve(__dirname, "../../../../", repoRelativePath)).href);
  }

  // When running services directly, process.cwd() is often:
  // - repoRoot/services/api
  // - repoRoot
  // Try both.
  candidates.push(
    pathToFileURL(path.resolve(process.cwd(), repoRelativePath)).href,
    pathToFileURL(path.resolve(process.cwd(), "..", "..", repoRelativePath)).href
  );

  return candidates;
}

/**
 * Import an ESM module from the repo's `packages/security/**` tree from CommonJS
 * code (services/api is compiled to CJS).
 *
 * This helper tries a small set of file URL candidates so it works in:
 * - local dev (repo checkout)
 * - Docker images where `packages/security/*` is copied to `/app/packages/security`
 */
export async function importSecurityModule(repoRelativePath: string): Promise<any> {
  let lastError: unknown;
  for (const specifier of candidatesForRepoRelativePath(repoRelativePath)) {
    try {
      return await importEsm(specifier);
    } catch (err) {
      lastError = err;
    }
  }

  throw lastError instanceof Error ? lastError : new Error(`Failed to load ${repoRelativePath}`);
}

