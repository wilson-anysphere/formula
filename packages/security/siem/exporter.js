import * as http from "node:http";
import * as https from "node:https";
import { createHash } from "node:crypto";

import { serializeBatch } from "./format.js";

const UUID_REGEX = /^[0-9a-f]{8}-[0-9a-f]{4}-[1-5][0-9a-f]{3}-[89ab][0-9a-f]{3}-[0-9a-f]{12}$/i;

function assertUuid(id) {
  if (typeof id !== "string" || !UUID_REGEX.test(id)) {
    throw new Error("audit event id must be a UUID");
  }
}

function batchIdempotencyKey(events) {
  const ids = events.map((event) => event?.id ?? "");
  for (const id of ids) assertUuid(id);
  return createHash("sha256").update(ids.join(","), "utf8").digest("hex");
}

function sleep(ms) {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

function withJitter(delayMs, jitter) {
  if (!jitter) return delayMs;
  const min = delayMs * 0.5;
  const max = delayMs * 1.5;
  return Math.floor(min + Math.random() * (max - min));
}

export async function retryWithBackoff(fn, options) {
  const {
    maxAttempts = 5,
    baseDelayMs = 500,
    maxDelayMs = 30_000,
    jitter = true,
    retryOn = (error) => error && error.retriable
  } = options || {};

  let lastError;
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

export function buildAuthHeaders(auth) {
  if (!auth || auth.type === "none") return {};

  if (auth.type === "bearer") {
    if (!auth.token) throw new Error("auth.token is required for bearer auth");
    return { Authorization: `Bearer ${auth.token}` };
  }

  if (auth.type === "basic") {
    if (!auth.username || !auth.password) throw new Error("auth.username and auth.password are required for basic auth");
    const token = Buffer.from(`${auth.username}:${auth.password}`).toString("base64");
    return { Authorization: `Basic ${token}` };
  }

  if (auth.type === "header") {
    if (!auth.name || !auth.value) throw new Error("auth.name and auth.value are required for header auth");
    return { [auth.name]: auth.value };
  }

  throw new Error(`Unsupported auth.type: ${auth.type}`);
}

export function postBatch({ endpointUrl, body, contentType, headers = {}, timeoutMs = 10_000 }) {
  const url = new URL(endpointUrl);
  const moduleToUse = url.protocol === "https:" ? https : http;

  const requestOptions = {
    protocol: url.protocol,
    hostname: url.hostname,
    port: url.port ? Number(url.port) : url.protocol === "https:" ? 443 : 80,
    path: `${url.pathname}${url.search}`,
    method: "POST",
    headers: {
      "Content-Type": contentType,
      "Content-Length": body.length,
      ...headers
    },
    timeout: timeoutMs
  };

  return new Promise((resolve, reject) => {
    const req = moduleToUse.request(requestOptions, (res) => {
      const chunks = [];
      res.on("data", (chunk) => chunks.push(chunk));
      res.on("end", () => {
        const responseBody = Buffer.concat(chunks).toString("utf8");
        const status = res.statusCode || 0;
        if (status >= 200 && status < 300) return resolve({ status, body: responseBody });

        const error = new Error(`SIEM endpoint responded with status ${status}`);
        error.status = status;
        error.responseBody = responseBody;
        error.retriable = status >= 500 || status === 429 || status === 408;
        return reject(error);
      });
    });

    req.on("timeout", () => {
      const error = new Error("SIEM endpoint request timed out");
      error.retriable = true;
      req.destroy(error);
    });

    req.on("error", (error) => {
      if (error && typeof error.retriable === "boolean") return reject(error);
      const wrapped = new Error(`Failed to POST SIEM batch: ${error.message}`);
      wrapped.cause = error;
      wrapped.retriable = true;
      reject(wrapped);
    });

    req.write(body);
    req.end();
  });
}

export class SiemExporter {
  constructor(config) {
    if (!config || !config.endpointUrl) throw new Error("SIEM exporter requires endpointUrl");

    this.config = {
      format: "json",
      batchSize: 250,
      flushIntervalMs: 5_000,
      timeoutMs: 10_000,
      idempotencyKeyHeader: null,
      retry: {
        maxAttempts: 5,
        baseDelayMs: 500,
        maxDelayMs: 30_000,
        jitter: true
      },
      ...config
    };

    this.buffer = [];
    this.flushPromise = null;
    this.flushRequested = false;

    if (this.config.flushIntervalMs > 0) {
      this.interval = setInterval(() => {
        void this.flush().catch(() => {
          // Best-effort: avoid unhandled rejections when automatic background flush fails.
        });
      }, this.config.flushIntervalMs);
      this.interval.unref?.();
    }
  }

  enqueue(event) {
    if (!event || typeof event !== "object") throw new Error("SIEM audit event must be an object");
    assertUuid(event.id);
    this.buffer.push(event);
    if (this.buffer.length >= this.config.batchSize) {
      void this.flush().catch(() => {
        // Best-effort: avoid unhandled rejections when a fire-and-forget flush fails.
      });
    }
  }

  async flush() {
    this.flushRequested = false;
    if (this.flushPromise) return this.flushPromise;

    this.flushPromise = (async () => {
      while (true) {
        if (this.buffer.length === 0) break;

        const batch = this.buffer.splice(0, this.config.batchSize);
        try {
          await this.sendBatch(batch);
        } catch (error) {
          this.buffer = batch.concat(this.buffer);
          throw error;
        }
      }
    })();

    try {
      await this.flushPromise;
    } finally {
      this.flushPromise = null;
    }
  }

  async sendBatch(events) {
    if (!events || events.length === 0) return;
    for (const event of events) assertUuid(event?.id);

    const { body, contentType } = serializeBatch(events, {
      format: this.config.format,
      redactionOptions: this.config.redactionOptions
    });

    const headers = {
      ...buildAuthHeaders(this.config.auth),
      ...(this.config.headers || {})
    };

    if (this.config.idempotencyKeyHeader) {
      headers[this.config.idempotencyKeyHeader] = batchIdempotencyKey(events);
    }

    await retryWithBackoff(
      () =>
        postBatch({
          endpointUrl: this.config.endpointUrl,
          body,
          contentType,
          headers,
          timeoutMs: this.config.timeoutMs
        }),
      this.config.retry
    );
  }

  async stop({ flush = true } = {}) {
    if (this.interval) clearInterval(this.interval);
    this.interval = null;

    if (flush) await this.flush();
  }
}
