import crypto from "node:crypto";
import type {
  CanonicalAuditEvent,
  SiemAuthConfig,
  SiemEndpointConfig,
  SiemRetryConfig,
  MaybeEncryptedSecret
} from "./types";
import { serializeBatch } from "./format";
import { fetchWithOrgTls, type OrgTlsPolicy } from "../http/tls";

type RetriableError = Error & { retriable?: boolean; status?: number; responseBody?: string };

function sleep(ms: number): Promise<void> {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

function withJitter(delayMs: number, jitter: boolean): number {
  if (!jitter) return delayMs;
  const min = delayMs * 0.5;
  const max = delayMs * 1.5;
  return Math.floor(min + Math.random() * (max - min));
}

export async function retryWithBackoff<T>(
  fn: (attempt: number) => Promise<T>,
  options: SiemRetryConfig & { retryOn?: (error: unknown) => boolean } = {}
): Promise<T> {
  const {
    maxAttempts = 5,
    baseDelayMs = 500,
    maxDelayMs = 30_000,
    jitter = true,
    retryOn = (error: unknown) => Boolean((error as RetriableError | undefined)?.retriable)
  } = options;

  let lastError: unknown;
  for (let attempt = 1; attempt <= maxAttempts; attempt++) {
    try {
      return await fn(attempt);
    } catch (error) {
      lastError = error;
      const shouldRetry = attempt < maxAttempts && retryOn(error);
      if (!shouldRetry) throw error;

      const exponential = baseDelayMs * 2 ** (attempt - 1);
      const delay = withJitter(Math.min(maxDelayMs, exponential), jitter);
      await sleep(delay);
    }
  }

  throw lastError;
}

function batchIdempotencyKey(events: Array<{ id: string }>): string {
  const ids = events.map((event) => event.id).join(",");
  return crypto.createHash("sha256").update(ids, "utf8").digest("hex");
}

async function resolveSecret(value: MaybeEncryptedSecret | undefined): Promise<string | undefined> {
  if (!value) return undefined;
  if (typeof value === "string") return value;

  if ("secretRef" in value && typeof value.secretRef === "string") {
    // Secret refs must be resolved (decrypted) by the config provider before the
    // sender is invoked.
    throw new Error("Unresolved secretRef in SIEM config");
  }

  if ("encrypted" in value && typeof value.encrypted === "string") {
    // Backwards-compatible placeholder for configs written before the API secret
    // store existed (or for self-encrypted/legacy configs).
    return value.encrypted;
  }

  if ("ciphertext" in value && typeof value.ciphertext === "string") {
    // Backwards-compatible placeholder for configs written before the API secret
    // store existed (or for self-encrypted/legacy configs).
    return value.ciphertext;
  }

  throw new Error("Unsupported secret value");
}

export async function buildAuthHeaders(auth: SiemAuthConfig | undefined): Promise<Record<string, string>> {
  if (!auth || auth.type === "none") return {};

  if (auth.type === "bearer") {
    const token = await resolveSecret(auth.token);
    if (!token) throw new Error("auth.token is required for bearer auth");
    return { Authorization: `Bearer ${token}` };
  }

  if (auth.type === "basic") {
    const username = await resolveSecret(auth.username);
    const password = await resolveSecret(auth.password);
    if (!username || !password) throw new Error("auth.username and auth.password are required for basic auth");
    const token = Buffer.from(`${username}:${password}`).toString("base64");
    return { Authorization: `Basic ${token}` };
  }

  if (auth.type === "header") {
    const value = await resolveSecret(auth.value);
    if (!auth.name || !value) throw new Error("auth.name and auth.value are required for header auth");
    return { [auth.name]: value };
  }

  // Exhaustiveness guard.
  throw new Error(`Unsupported auth.type: ${(auth as SiemAuthConfig).type}`);
}

async function postBatch(options: {
  endpointUrl: string;
  body: Buffer;
  contentType: string;
  headers: Record<string, string>;
  timeoutMs: number;
  tls?: OrgTlsPolicy;
}): Promise<void> {
  const res = await fetchWithOrgTls(
    options.endpointUrl,
    {
      method: "POST",
      headers: {
        "Content-Type": options.contentType,
        ...options.headers
      },
      body: options.body.toString("utf8"),
      signal: AbortSignal.timeout(options.timeoutMs)
    },
    { tls: options.tls }
  );

  if (res.ok) return;

  const responseBody = await res.text().catch(() => "");
  const error: RetriableError = new Error(`SIEM endpoint responded with status ${res.status}`);
  error.status = res.status;
  error.responseBody = responseBody;
  error.retriable = res.status >= 500 || res.status === 429 || res.status === 408;
  throw error;
}

export async function sendSiemBatch(
  config: SiemEndpointConfig,
  events: CanonicalAuditEvent[],
  options: { tls?: OrgTlsPolicy } = {}
): Promise<void> {
  if (!events || events.length === 0) return;

  const { body, contentType } = serializeBatch(events, {
    format: config.format ?? "json",
    redactionText: config.redactionOptions?.redactionText,
    sensitiveKeyPatterns: config.redactionOptions?.sensitiveKeyPatterns
  });

  const headers: Record<string, string> = {
    ...(await buildAuthHeaders(config.auth)),
    ...(config.headers ?? {})
  };

  if (config.idempotencyKeyHeader) {
    headers[config.idempotencyKeyHeader] = batchIdempotencyKey(events);
  }

  const timeoutMs = config.timeoutMs ?? 10_000;

  await retryWithBackoff(
    async () => {
      try {
        await postBatch({
          endpointUrl: config.endpointUrl,
          body,
          contentType,
          headers,
          timeoutMs,
          tls: options.tls
        });
      } catch (err) {
        const error: RetriableError =
          err instanceof Error
            ? (err as RetriableError)
            : Object.assign(new Error(`Failed to POST SIEM batch: ${String(err)}`), { retriable: true });

        // Treat network/timeouts as retriable.
        if (typeof error.retriable !== "boolean") error.retriable = true;
        throw error;
      }
    },
    config.retry ?? {}
  );
}
