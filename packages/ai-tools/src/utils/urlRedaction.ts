/**
 * Shared URL sanitization helpers for tool results + audit logging.
 *
 * IMPORTANT: Any changes here affect both:
 * - fetch_external_data tool result provenance (`ToolExecutor`)
 * - audit log tool call parameter redaction (`runChatWithToolsAudited`)
 */
export function redactUrlSecrets(raw: string | URL): string {
  try {
    const url = typeof raw === "string" ? new URL(raw) : new URL(raw.toString());

    // Avoid leaking userinfo and fragments into tool results / audit logs.
    url.username = "";
    url.password = "";
    url.hash = "";

    if (url.search) {
      const params = new URLSearchParams(url.search);
      const keys = Array.from(new Set(Array.from(params.keys())));
      for (const key of keys) {
        if (!isSensitiveQueryParam(key)) continue;
        const count = params.getAll(key).length;
        params.delete(key);
        for (let i = 0; i < count; i++) params.append(key, "REDACTED");
      }
      const next = params.toString();
      url.search = next ? `?${next}` : "";
    }

    return url.toString();
  } catch {
    return typeof raw === "string" ? raw : raw.toString();
  }
}

export function isSensitiveQueryParam(key: string): boolean {
  const normalized = key.toLowerCase();
  return (
    normalized === "key" ||
    normalized === "api_key" ||
    normalized === "apikey" ||
    normalized === "token" ||
    normalized === "access_token" ||
    normalized === "auth" ||
    normalized === "authorization" ||
    normalized === "signature" ||
    normalized === "sig" ||
    normalized === "password" ||
    normalized === "secret" ||
    // Common OAuth parameter name. Keep in sync with other redaction surfaces.
    normalized === "client_secret"
  );
}

