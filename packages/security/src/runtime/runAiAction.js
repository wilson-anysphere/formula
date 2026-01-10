import { createSecureFetch, createSecureFs } from "../secureApis/createSecureApis.js";

export async function runAiAction({
  aiSessionId,
  action,
  permissionManager,
  auditLogger = null,
  promptIfDenied = false
}) {
  const principal = { type: "ai", id: String(aiSessionId) };

  auditLogger?.log({
    eventType: "security.ai.action",
    actor: principal,
    success: true,
    metadata: { phase: "start", action: action?.type }
  });

  try {
    if (!action || typeof action.type !== "string") {
      throw new TypeError("AI action must have a type");
    }

    if (action.type === "filesystem.writeFile") {
      const secureFs = createSecureFs({ principal, permissionManager, auditLogger, promptIfDenied });
      await secureFs.writeFile(action.path, action.data);
      auditLogger?.log({
        eventType: "security.ai.action",
        actor: principal,
        success: true,
        metadata: { phase: "complete", action: action.type }
      });
      return null;
    }

    if (action.type === "filesystem.readFile") {
      const secureFs = createSecureFs({ principal, permissionManager, auditLogger, promptIfDenied });
      const data = await secureFs.readFile(action.path, action.options);
      auditLogger?.log({
        eventType: "security.ai.action",
        actor: principal,
        success: true,
        metadata: { phase: "complete", action: action.type }
      });
      return data;
    }

    if (action.type === "network.fetch") {
      const secureFetch = createSecureFetch({ principal, permissionManager, auditLogger, promptIfDenied });
      const res = await secureFetch(action.url, action.init);
      auditLogger?.log({
        eventType: "security.ai.action",
        actor: principal,
        success: true,
        metadata: { phase: "complete", action: action.type }
      });
      return res;
    }

    if (action.type === "file.export") {
      const secureFs = createSecureFs({ principal, permissionManager, auditLogger, promptIfDenied });
      await secureFs.writeFile(action.path, action.data);
      auditLogger?.log({
        eventType: "security.file.export",
        actor: principal,
        success: true,
        metadata: { path: action.path }
      });
      return null;
    }

    throw new Error(`Unknown AI action type: ${action.type}`);
  } catch (error) {
    auditLogger?.log({
      eventType: "security.ai.action",
      actor: principal,
      success: false,
      metadata: { phase: "error", action: action?.type, message: error.message }
    });
    throw error;
  }
}
