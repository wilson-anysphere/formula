import os from "node:os";
import path from "node:path";

import {
  AuditLogger,
  PermissionManager,
  SqliteAuditLogStore,
  runAiAction,
  runConnector,
  runExtension,
  runScript
} from "../../../../packages/security/src/index.js";

function defaultAuditDbPath() {
  // In a real Tauri application we'd use the platform-specific app data dir.
  // For this baseline, keep it deterministic and outside the repo.
  return path.join(os.homedir(), ".formula", "audit.sqlite");
}

export function createDesktopSecurity({ auditDbPath = defaultAuditDbPath(), onPrompt = null } = {}) {
  const store = new SqliteAuditLogStore({ path: auditDbPath });
  const auditLogger = new AuditLogger({ store });
  const permissionManager = new PermissionManager({ auditLogger, onPrompt });

  return {
    auditLogger,
    permissionManager,
    runAiAction: (options) => runAiAction({ ...options, permissionManager, auditLogger }),
    runConnector: (options) => runConnector({ ...options, permissionManager, auditLogger }),
    runExtension: (options) => runExtension({ ...options, permissionManager, auditLogger }),
    runScript: (options) => runScript({ ...options, permissionManager, auditLogger })
  };
}
