/**
 * Minimal Cursor-backed tab completion client.
 *
 * Cursor controls authentication, prompt/harness, and model routing. This client
 * intentionally sends a structured request with no prompt engineering.
 */
export class CursorTabCompletionClient {
  /**
   * @param {{
   *   baseUrl?: string,
   *   fetchImpl?: typeof fetch,
   *   timeoutMs?: number
   * }} [options]
   */
  constructor(options = {}) {
    this.endpointUrl = resolveTabCompletionEndpoint(options.baseUrl);
    // Node 18+ provides global fetch. Allow injection for tests.
    // eslint-disable-next-line no-undef
    this.fetchImpl = options.fetchImpl ?? fetch;
    // Tab completion has a strict latency budget.
    this.timeoutMs = options.timeoutMs ?? 100;
  }

  /**
   * @param {{ input: string; cursorPosition: number; cellA1: string }} req
   * @returns {Promise<string>}
   */
  async completeTabCompletion(req) {
    const input = (req?.input ?? "").toString();
    const cursorPosition = Number.isInteger(req?.cursorPosition) ? req.cursorPosition : input.length;
    const cellA1 = (req?.cellA1 ?? "").toString();

    const controller = new AbortController();
    const timeoutMs = this.timeoutMs;
    const timeout =
      Number.isFinite(timeoutMs) && timeoutMs > 0 ? setTimeout(() => controller.abort(), timeoutMs) : null;

    try {
      const res = await this.fetchImpl(this.endpointUrl, {
        method: "POST",
        headers: { "content-type": "application/json" },
        // Cursor authentication is handled by the session (cookies). When the
        // backend is cross-origin, this requires CORS support + credentials.
        credentials: "include",
        body: JSON.stringify({ input, cursorPosition, cellA1 }),
        signal: controller.signal,
      });

      if (!res?.ok) return "";

      const data = await readJsonOrText(res);
      if (typeof data === "string") return data.trim();

      const completion =
        typeof data?.completion === "string" ? data.completion : typeof data?.text === "string" ? data.text : "";
      return completion.toString();
    } catch {
      // Cursor tab completion is optional; treat all failures as "no completion".
      return "";
    } finally {
      if (timeout) clearTimeout(timeout);
    }
  }
}

async function readJsonOrText(res) {
  try {
    return await res.json();
  } catch {
    try {
      return await res.text();
    } catch {
      return null;
    }
  }
}

function resolveTabCompletionEndpoint(baseUrl) {
  const raw = (baseUrl ?? "").toString().trim();
  if (!raw) return "/api/ai/tab-completion";

  const trimmed = raw.replace(/\/$/, "");
  // Allow callers to pass a fully qualified endpoint (common for tests/dev).
  if (/(^|\/)tab-completion$/.test(trimmed)) return trimmed;

  return joinUrl(trimmed, "/api/ai/tab-completion");
}

function joinUrl(baseUrl, path) {
  if (!baseUrl) return path;
  const base = baseUrl.replace(/\/$/, "");
  const suffix = path.startsWith("/") ? path : `/${path}`;
  return `${base}${suffix}`;
}
