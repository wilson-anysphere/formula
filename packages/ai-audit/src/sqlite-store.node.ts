import { createRequire } from "node:module";

import { SqliteAIAuditStore } from "./sqlite-store.js";
import type { SqliteAIAuditStoreOptions } from "./sqlite-store.js";

export function locateSqlJsFileNode(file: string): string {
  const require = createRequire(import.meta.url);
  return require.resolve(`sql.js/dist/${file}`);
}

export async function createSqliteAIAuditStoreNode(
  options: Omit<SqliteAIAuditStoreOptions, "locateFile"> = {}
): Promise<SqliteAIAuditStore> {
  return SqliteAIAuditStore.create({
    ...options,
    locateFile: locateSqlJsFileNode
  });
}
