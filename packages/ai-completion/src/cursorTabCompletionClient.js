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
   *   timeoutMs?: number,
   *   getAuthHeaders?: () => (Record<string,string> | Promise<Record<string,string>>)
   * }} [options]
   */
  constructor(options = {}) {
    this.endpointUrl = resolveTabCompletionEndpoint(options.baseUrl);
    // Node 18+ provides global fetch. Allow injection for tests.
    // eslint-disable-next-line no-undef
    this.fetchImpl = options.fetchImpl ?? fetch;
    // Tab completion has a strict latency budget.
    this.timeoutMs = options.timeoutMs ?? 100;
    this.getAuthHeaders = options.getAuthHeaders;
  }

  /**
   * @param {{ input: string; cursorPosition: number; cellA1: string; signal?: AbortSignal }} req
   * @returns {Promise<string>}
   */
  async completeTabCompletion(req) {
    const input = (req?.input ?? "").toString();
    const cursorPosition = Number.isInteger(req?.cursorPosition) ? req.cursorPosition : input.length;
    const cellA1 = (req?.cellA1 ?? "").toString();

    /** @type {AbortSignal | undefined} */
    const externalSignal = req?.signal;
    if (externalSignal?.aborted) return "";

    const controller = new AbortController();
    const timeoutMs = this.timeoutMs;
    const timeout =
      Number.isFinite(timeoutMs) && timeoutMs > 0 ? setTimeout(() => controller.abort(), timeoutMs) : null;

    const onExternalAbort = externalSignal
      ? () => {
          controller.abort(externalSignal.reason);
        }
      : null;

    if (externalSignal && onExternalAbort) {
      // Use a separate AbortController so we can combine caller cancellation with
      // the internal latency budget timeout.
      externalSignal.addEventListener("abort", onExternalAbort, { once: true });
      // In case the signal aborted between the early check above and registering
      // the event listener.
      if (externalSignal.aborted) onExternalAbort();
    }

    try {
      if (controller.signal.aborted) return "";

      const authHeaders =
        typeof this.getAuthHeaders === "function"
          ? await this.getAuthHeaders()
          : null;
      if (controller.signal.aborted) return "";

      // Use lowercase so tests (and any header-inspecting consumers) can treat this as a plain
      // record without worrying about case. Fetch treats header keys as case-insensitive.
      const headers = { ...(authHeaders ?? {}), "content-type": "application/json" };

      const res = await this.fetchImpl(this.endpointUrl, {
        method: "POST",
        headers,
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
      if (externalSignal && onExternalAbort) externalSignal.removeEventListener("abort", onExternalAbort);
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
