import { createSecureFetch } from "../secureApis/createSecureApis.js";

export async function runConnector({
  connectorId,
  request,
  permissionManager,
  auditLogger = null,
  promptIfDenied = false
}) {
  const principal = { type: "connector", id: String(connectorId) };
  const secureFetch = createSecureFetch({ principal, permissionManager, auditLogger, promptIfDenied });

  auditLogger?.log({
    eventType: "security.connector.request",
    actor: principal,
    success: true,
    metadata: { phase: "start", url: request?.url }
  });

  try {
    const res = await secureFetch(request.url, request.init);
    auditLogger?.log({
      eventType: "security.connector.request",
      actor: principal,
      success: true,
      metadata: { phase: "complete", url: request?.url }
    });
    return res;
  } catch (error) {
    auditLogger?.log({
      eventType: "security.connector.request",
      actor: principal,
      success: false,
      metadata: { phase: "error", url: request?.url, message: error.message }
    });
    throw error;
  }
}
