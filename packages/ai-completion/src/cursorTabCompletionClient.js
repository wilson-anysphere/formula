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
   *   // Cursor-managed auth headers (e.g. Authorization) for environments where
   *   // cookie-based session auth isn't available. Do not use user API keys.
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
          ? await raceWithAbort(this.getAuthHeaders(), controller.signal)
          : null;
      if (controller.signal.aborted) return "";

      // Merge Cursor-managed auth headers while forcing JSON Content-Type.
      //
      // IMPORTANT: ensure we only send a single Content-Type header. If we pass both
      // `Content-Type` and `content-type`, fetch() can combine them into a single comma-separated
      // value (e.g. "text/plain, application/json"), which can break server-side JSON parsing.
      //
      // Preserve header casing from `getAuthHeaders()` for compatibility with custom `fetchImpl`
      // implementations that inspect `init.headers` as a plain object.
      /** @type {Record<string, string>} */
      const headers = {};
      /** @type {Map<string, string>} */
      const keyByLower = new Map();

      /**
       * Set/overwrite a header name case-insensitively.
       *
       * Fetch treats header names case-insensitively and may combine duplicate
       * header names when it converts a plain object to `Headers`. We avoid this
       * by ensuring we never return the same header name twice with different
       * casing (e.g. `Authorization` and `authorization`).
       *
       * @param {string} key
       * @param {string} value
       */
      const setHeader = (key, value) => {
        const lower = String(key).toLowerCase();
        const canonicalKey =
          lower === "authorization" ? "Authorization" : lower === "content-type" ? "Content-Type" : key;

        const existingKey = keyByLower.get(lower);
        if (existingKey && existingKey !== canonicalKey) {
          delete headers[existingKey];
        }

        headers[canonicalKey] = value;
        keyByLower.set(lower, canonicalKey);
      };

      if (authHeaders && typeof authHeaders === "object") {
        for (const [key, value] of Object.entries(authHeaders)) {
          if (!key) continue;
          if (value === undefined || value === null) continue;
          if (key.toLowerCase() === "content-type") continue;
          setHeader(key, String(value));
        }
      }
      setHeader("Content-Type", "application/json");
      if (controller.signal.aborted) return "";

      const res = await raceWithAbort(
        this.fetchImpl(this.endpointUrl, {
        method: "POST",
        headers,
        // Cursor authentication is handled by the session (cookies). When the
        // backend is cross-origin, this requires CORS support + credentials.
        //
        // Some environments (e.g. desktop wrappers, servers) may not have access
        // to session cookies; callers can instead supply `getAuthHeaders()` to
        // provide Cursor-managed auth headers.
        credentials: "include",
        body: JSON.stringify({ input, cursorPosition, cellA1 }),
        signal: controller.signal,
        }),
        controller.signal,
      );

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

/**
 * Await a promise-like value, but reject early if the provided AbortSignal fires.
 *
 * This is used to ensure the tab completion latency budget applies to both
 * auth-header resolution and the fetch() call.
 *
 * @template T
 * @param {T | Promise<T>} promiseLike
 * @param {AbortSignal} signal
 * @returns {Promise<T>}
 */
async function raceWithAbort(promiseLike, signal) {
  if (!signal) return await promiseLike;
  if (signal.aborted) throw signal.reason;

  /** @type {(() => void) | null} */
  let cleanup = null;
  const abortPromise = new Promise((_, reject) => {
    const onAbort = () => reject(signal.reason);
    cleanup = () => signal.removeEventListener("abort", onAbort);
    signal.addEventListener("abort", onAbort, { once: true });
  });

  try {
    return await Promise.race([Promise.resolve(promiseLike), abortPromise]);
  } finally {
    cleanup?.();
  }
}
