const http = require("node:http");
const crypto = require("node:crypto");

const { MarketplaceStore } = require("./store");

const CACHE_CONTROL_REVALIDATE = "public, max-age=0, must-revalidate";

class HttpError extends Error {
  constructor(statusCode, message) {
    super(message);
    this.statusCode = statusCode;
  }
}

function sendJson(res, statusCode, body, extraHeaders = {}) {
  const bytes = Buffer.from(JSON.stringify(body, null, 2));
  res.writeHead(statusCode, {
    "Content-Type": "application/json; charset=utf-8",
    "Content-Length": bytes.length,
    ...extraHeaders,
  });
  res.end(bytes);
}

function sendText(res, statusCode, text, extraHeaders = {}) {
  const bytes = Buffer.from(String(text));
  res.writeHead(statusCode, {
    "Content-Type": "text/plain; charset=utf-8",
    "Content-Length": bytes.length,
    ...extraHeaders,
  });
  res.end(bytes);
}

function notFound(res) {
  sendJson(res, 404, { error: "Not found" });
}

async function readJsonBody(req, { limitBytes = 5 * 1024 * 1024 } = {}) {
  const chunks = [];
  let size = 0;
  for await (const chunk of req) {
    size += chunk.length;
    if (size > limitBytes) {
      throw new Error("Request body too large");
    }
    chunks.push(chunk);
  }
  const raw = Buffer.concat(chunks).toString("utf8");
  if (!raw.trim()) return null;
  return JSON.parse(raw);
}

async function readBinaryBody(req, { limitBytes = 20 * 1024 * 1024 } = {}) {
  const chunks = [];
  let size = 0;
  const hash = crypto.createHash("sha256");
  for await (const chunk of req) {
    size += chunk.length;
    if (size > limitBytes) {
      throw new Error("Request body too large");
    }
    hash.update(chunk);
    chunks.push(chunk);
  }
  return { bytes: Buffer.concat(chunks), sha256: hash.digest("hex") };
}

function getBearerToken(req) {
  const auth = req.headers.authorization;
  if (!auth) return null;
  const match = /^Bearer\s+(.+)$/.exec(auth);
  return match ? match[1] : null;
}

function sha256Hex(input) {
  return crypto.createHash("sha256").update(input).digest("hex");
}

function normalizeEtag(value) {
  const raw = String(value || "").trim();
  if (!raw) return "";
  let tag = raw;
  if (tag.startsWith("W/")) tag = tag.slice(2).trim();
  if (tag.startsWith('"') && tag.endsWith('"')) tag = tag.slice(1, -1);
  return tag;
}

function etagMatches(ifNoneMatch, etag) {
  if (!ifNoneMatch || !etag) return false;
  const target = normalizeEtag(etag);
  if (!target) return false;
  return String(ifNoneMatch)
    .split(",")
    .map((part) => part.trim())
    .some((part) => {
      if (!part) return false;
      if (part === "*") return true;
      return normalizeEtag(part) === target;
    });
}

function parsePath(pathname) {
  return pathname.split("/").filter(Boolean);
}

function getClientIp(req) {
  const xff = req.headers["x-forwarded-for"];
  if (typeof xff === "string" && xff.trim()) return xff.split(",")[0].trim();
  return req.socket?.remoteAddress || "unknown";
}

function parseBooleanParam(value) {
  if (value === null || value === undefined) return undefined;
  const raw = String(value).toLowerCase().trim();
  if (raw === "1" || raw === "true" || raw === "yes") return true;
  if (raw === "0" || raw === "false" || raw === "no") return false;
  return undefined;
}

class TokenBucketRateLimiter {
  constructor({ capacity, refillMs, maxEntries = 10_000 }) {
    this.capacity = capacity;
    this.refillMs = refillMs;
    this.maxEntries = maxEntries;
    /** @type {Map<string, { tokens: number, updatedAt: number }>} */
    this.state = new Map();
    this._lastPrunedAt = 0;
  }

  _prune(now) {
    const pruneEveryMs = this.refillMs;
    const maxAgeMs = this.refillMs * 10;
    if (this.state.size <= this.maxEntries && now - this._lastPrunedAt < pruneEveryMs) return;
    this._lastPrunedAt = now;

    for (const [key, entry] of this.state.entries()) {
      if (now - entry.updatedAt > maxAgeMs) {
        this.state.delete(key);
      }
    }

    // Hard cap: if still too big (e.g. many unique abusive keys), drop everything.
    if (this.state.size > this.maxEntries * 2) {
      this.state.clear();
    }
  }

  take(key) {
    if (!Number.isFinite(this.capacity) || this.capacity <= 0) {
      return { ok: true, retryAfterMs: 0 };
    }

    const now = Date.now();
    this._prune(now);
    const existing = this.state.get(key) || { tokens: this.capacity, updatedAt: now };
    const elapsed = now - existing.updatedAt;
    const refill = (elapsed / this.refillMs) * this.capacity;
    const tokens = Math.min(this.capacity, existing.tokens + refill);

    if (tokens < 1) {
      this.state.set(key, { tokens, updatedAt: now });
      const refillRatePerMs = this.capacity / this.refillMs;
      const missing = Math.max(0, 1 - tokens);
      const retryAfterMs =
        Number.isFinite(refillRatePerMs) && refillRatePerMs > 0 ? Math.ceil(missing / refillRatePerMs) : this.refillMs;
      return { ok: false, retryAfterMs };
    }

    this.state.set(key, { tokens: tokens - 1, updatedAt: now });
    return { ok: true, retryAfterMs: 0 };
  }
}

function createLogger() {
  return {
    info(event) {
      // eslint-disable-next-line no-console
      console.log(JSON.stringify({ level: "info", time: new Date().toISOString(), ...event }));
    },
    error(event) {
      // eslint-disable-next-line no-console
      console.error(JSON.stringify({ level: "error", time: new Date().toISOString(), ...event }));
    },
  };
}

function formatPrometheusMetrics(metrics) {
  const lines = [];
  lines.push("# HELP marketplace_http_requests_total Total HTTP requests handled by the marketplace service");
  lines.push("# TYPE marketplace_http_requests_total counter");
  for (const [key, count] of metrics.requests.entries()) {
    const [method, route, status] = key.split(" ");
    lines.push(`marketplace_http_requests_total{method="${method}",route="${route}",status="${status}"} ${count}`);
  }
  lines.push("# HELP marketplace_rate_limited_total Total requests rejected by rate limiting");
  lines.push("# TYPE marketplace_rate_limited_total counter");
  lines.push(`marketplace_rate_limited_total ${metrics.rateLimited}`);
  return lines.join("\n") + "\n";
}

async function createMarketplaceServer({ dataDir, adminToken = null, rateLimits: rateLimitOverrides = {} } = {}) {
  if (!dataDir) throw new Error("dataDir is required");
  const store = new MarketplaceStore({ dataDir });
  await store.init();

  const logger = createLogger();
  const metrics = {
    requests: new Map(),
    rateLimited: 0,
  };

  const rateLimits = {
    publishPerPublisherPerMinute: 10,
    searchPerIpPerMinute: 30,
    getExtensionPerIpPerMinute: 60,
    downloadPerIpPerMinute: 120,
    ...rateLimitOverrides,
  };

  const publishLimiter = new TokenBucketRateLimiter({
    capacity: rateLimits.publishPerPublisherPerMinute,
    refillMs: 60_000,
  });
  const searchLimiter = new TokenBucketRateLimiter({
    capacity: rateLimits.searchPerIpPerMinute,
    refillMs: 60_000,
  });
  const getExtensionLimiter = new TokenBucketRateLimiter({
    capacity: rateLimits.getExtensionPerIpPerMinute,
    refillMs: 60_000,
  });
  const downloadLimiter = new TokenBucketRateLimiter({
    capacity: rateLimits.downloadPerIpPerMinute,
    refillMs: 60_000,
  });

  const server = http.createServer(async (req, res) => {
    const startedAt = process.hrtime.bigint();
    const ip = getClientIp(req);
    const method = req.method || "GET";
    let route = "unknown";
    let statusCode = 500;

    try {
      const url = new URL(req.url || "/", "http://localhost");
      const segments = parsePath(url.pathname);

      if (req.method === "GET" && url.pathname === "/api/health") {
        route = "/api/health";
        statusCode = 200;
        return sendJson(res, 200, { ok: true });
      }

      if (req.method === "GET" && (url.pathname === "/metrics" || url.pathname === "/api/internal/metrics")) {
        route = url.pathname;
        statusCode = 200;
        const body = formatPrometheusMetrics(metrics);
        res.writeHead(200, {
          "Content-Type": "text/plain; version=0.0.4; charset=utf-8",
          "Content-Length": Buffer.byteLength(body),
        });
        res.end(body);
        return;
      }

      if (req.method === "GET" && url.pathname === "/api/search") {
        route = "/api/search";
        const limited = searchLimiter.take(ip);
        if (!limited.ok) {
          metrics.rateLimited += 1;
          res.setHeader("Retry-After", String(Math.ceil(limited.retryAfterMs / 1000)));
          statusCode = 429;
          return sendJson(res, 429, { error: "Too Many Requests" });
        }

        const q = url.searchParams.get("q") || "";
        const category = url.searchParams.get("category") || "";
        const tag = url.searchParams.get("tag") || "";
        const verified = parseBooleanParam(url.searchParams.get("verified"));
        const featured = parseBooleanParam(url.searchParams.get("featured"));
        const sort = url.searchParams.get("sort") || "relevance";
        const limit = Number(url.searchParams.get("limit") || "20");
        const offset = Number(url.searchParams.get("offset") || "0");
        const cursor = url.searchParams.get("cursor");
        statusCode = 200;
        return sendJson(
          res,
          200,
          await store.search({ q, category, tag, verified, featured, sort, limit, offset, cursor })
        );
      }

      if (req.method === "GET" && segments[0] === "api" && segments[1] === "extensions" && segments.length === 3) {
        route = "/api/extensions/:id";
        const limited = getExtensionLimiter.take(ip);
        if (!limited.ok) {
          metrics.rateLimited += 1;
          res.setHeader("Retry-After", String(Math.ceil(limited.retryAfterMs / 1000)));
          statusCode = 429;
          return sendJson(res, 429, { error: "Too Many Requests" });
        }

        const id = segments[2];
        const ext = await store.getExtension(id);
        if (!ext) {
          statusCode = 404;
          return sendJson(res, 404, { error: "Not found" });
        }
        const etag = `"${sha256Hex(`${ext.id}|${ext.updatedAt || ""}`)}"`;
        if (etagMatches(req.headers["if-none-match"], etag)) {
          statusCode = 304;
          res.writeHead(304, { ETag: etag, "Cache-Control": CACHE_CONTROL_REVALIDATE });
          res.end();
          return;
        }
        statusCode = 200;
        return sendJson(res, 200, ext, { ETag: etag, "Cache-Control": CACHE_CONTROL_REVALIDATE });
      }

      if (
        req.method === "GET" &&
        segments[0] === "api" &&
        segments[1] === "extensions" &&
        segments.length === 5 &&
        segments[3] === "download"
      ) {
        route = "/api/extensions/:id/download/:version";
        const limited = downloadLimiter.take(ip);
        if (!limited.ok) {
          metrics.rateLimited += 1;
          res.setHeader("Retry-After", String(Math.ceil(limited.retryAfterMs / 1000)));
          statusCode = 429;
          return sendJson(res, 429, { error: "Too Many Requests" });
        }

        const id = segments[2];
        const version = segments[4];
        const pkgMeta = await store.getPackage(id, version, { includeBytes: false, incrementDownloadCount: false });
        if (!pkgMeta) {
          statusCode = 404;
          return sendJson(res, 404, { error: "Not found" });
        }

        const etag = `"${pkgMeta.sha256}"`;
        if (etagMatches(req.headers["if-none-match"], etag)) {
          statusCode = 304;
          res.writeHead(304, {
            ETag: etag,
            "Cache-Control": CACHE_CONTROL_REVALIDATE,
            "X-Package-Signature": pkgMeta.signatureBase64,
            "X-Package-Sha256": pkgMeta.sha256,
            "X-Package-Format-Version": String(pkgMeta.formatVersion ?? 1),
            "X-Publisher": pkgMeta.publisher,
          });
          res.end();
          return;
        }

        const pkg = await store.getPackage(id, version);
        if (!pkg) {
          statusCode = 404;
          return sendJson(res, 404, { error: "Not found" });
        }

        statusCode = 200;
        res.writeHead(200, {
          "Content-Type": "application/vnd.formula.extension-package",
          "Content-Length": pkg.bytes.length,
          ETag: etag,
          "Cache-Control": CACHE_CONTROL_REVALIDATE,
          "X-Package-Signature": pkg.signatureBase64,
          "X-Package-Sha256": pkg.sha256,
          "X-Package-Format-Version": String(pkg.formatVersion ?? 1),
          "X-Publisher": pkg.publisher,
        });
        res.end(pkg.bytes);
        return;
      }

      if (req.method === "POST" && url.pathname === "/api/publish-bin") {
        route = "/api/publish-bin";
        const token = getBearerToken(req);
        if (!token) {
          statusCode = 401;
          return sendJson(res, 401, { error: "Missing Authorization token" });
        }

        const tokenSha = sha256Hex(token);
        const publisherRecord = await store.getPublisherByTokenSha256(tokenSha);
        if (!publisherRecord) {
          statusCode = 403;
          return sendJson(res, 403, { error: "Invalid token" });
        }

        const publisherRate = publishLimiter.take(tokenSha);
        if (!publisherRate.ok) {
          metrics.rateLimited += 1;
          res.setHeader("Retry-After", String(Math.ceil(publisherRate.retryAfterMs / 1000)));
          statusCode = 429;
          return sendJson(res, 429, { error: "Too Many Requests" });
        }

        const contentType = String(req.headers["content-type"] || "").toLowerCase();
        if (!contentType.startsWith("application/vnd.formula.extension-package")) {
          statusCode = 415;
          return sendJson(res, 415, {
            error: "Unsupported Content-Type (expected application/vnd.formula.extension-package)",
          });
        }

        const MAX_UPLOAD_BYTES = 20 * 1024 * 1024;
        const declaredLength = Number(req.headers["content-length"] || "0");
        if (Number.isFinite(declaredLength) && declaredLength > MAX_UPLOAD_BYTES) {
          statusCode = 413;
          return sendJson(res, 413, { error: "Request body too large" });
        }

        const { bytes: packageBytes, sha256 } = await readBinaryBody(req, { limitBytes: MAX_UPLOAD_BYTES });
        const expectedShaHeader = req.headers["x-package-sha256"];
        const expectedSha256 =
          typeof expectedShaHeader === "string"
            ? expectedShaHeader
            : Array.isArray(expectedShaHeader)
              ? expectedShaHeader[0] || null
              : null;
        if (expectedSha256 && String(expectedSha256).toLowerCase() !== sha256) {
          statusCode = 400;
          return sendJson(res, 400, { error: "X-Package-Sha256 does not match uploaded bytes" });
        }
        const isV1 = packageBytes.length >= 2 && packageBytes[0] === 0x1f && packageBytes[1] === 0x8b;
        const signatureHeader = req.headers["x-package-signature"];
        const signatureBase64 =
          typeof signatureHeader === "string"
            ? signatureHeader
            : Array.isArray(signatureHeader)
              ? signatureHeader[0] || null
              : null;
        if (isV1 && !signatureBase64) {
          statusCode = 400;
          return sendJson(res, 400, { error: "X-Package-Signature is required for v1 packages" });
        }

        const published = await store.publishExtension({
          publisher: publisherRecord.publisher,
          packageBytes,
          signatureBase64,
        });

        statusCode = 200;
        return sendJson(res, 200, published);
      }

      if (req.method === "POST" && url.pathname === "/api/publish") {
        route = "/api/publish";
        const token = getBearerToken(req);
        if (!token) {
          statusCode = 401;
          return sendJson(res, 401, { error: "Missing Authorization token" });
        }

        const tokenSha = sha256Hex(token);
        const publisherRecord = await store.getPublisherByTokenSha256(tokenSha);
        if (!publisherRecord) {
          statusCode = 403;
          return sendJson(res, 403, { error: "Invalid token" });
        }

        const publisherRate = publishLimiter.take(tokenSha);
        if (!publisherRate.ok) {
          metrics.rateLimited += 1;
          res.setHeader("Retry-After", String(Math.ceil(publisherRate.retryAfterMs / 1000)));
          statusCode = 429;
          return sendJson(res, 429, { error: "Too Many Requests" });
        }

        const body = await readJsonBody(req, { limitBytes: 20 * 1024 * 1024 });
        if (!body?.packageBase64) {
          statusCode = 400;
          return sendJson(res, 400, { error: "packageBase64 is required" });
        }

        const packageBytes = Buffer.from(body.packageBase64, "base64");
        const isV1 = packageBytes.length >= 2 && packageBytes[0] === 0x1f && packageBytes[1] === 0x8b;
        if (isV1 && !body.signatureBase64) {
          statusCode = 400;
          return sendJson(res, 400, { error: "signatureBase64 is required for v1 packages" });
        }
        const signatureBase64 = body.signatureBase64 ? String(body.signatureBase64) : null;

        const published = await store.publishExtension({
          publisher: publisherRecord.publisher,
          packageBytes,
          signatureBase64,
        });

        statusCode = 200;
        return sendJson(res, 200, published);
      }

      if (req.method === "POST" && url.pathname === "/api/publishers/register") {
        route = "/api/publishers/register";
        if (!adminToken) {
          statusCode = 404;
          return sendJson(res, 404, { error: "Endpoint disabled" });
        }
        const token = getBearerToken(req);
        if (token !== adminToken) {
          statusCode = 403;
          return sendJson(res, 403, { error: "Forbidden" });
        }

        const body = await readJsonBody(req);
        if (!body?.publisher || !body?.token || !body?.publicKeyPem) {
          statusCode = 400;
          return sendJson(res, 400, { error: "publisher, token, publicKeyPem are required" });
        }

        await store.registerPublisher({
          publisher: body.publisher,
          tokenSha256: sha256Hex(body.token),
          publicKeyPem: body.publicKeyPem,
          verified: Boolean(body.verified),
        });

        statusCode = 200;
        return sendJson(res, 200, { ok: true });
      }

      if (
        req.method === "PATCH" &&
        segments[0] === "api" &&
        segments[1] === "extensions" &&
        segments[3] === "flags" &&
        segments.length === 4
      ) {
        route = "/api/extensions/:id/flags";
        if (!adminToken) {
          statusCode = 404;
          return sendJson(res, 404, { error: "Endpoint disabled" });
        }
        const token = getBearerToken(req);
        if (token !== adminToken) {
          statusCode = 403;
          return sendJson(res, 403, { error: "Forbidden" });
        }

        const id = segments[2];
        const body = await readJsonBody(req);
        try {
          const updated = await store.setExtensionFlags(
            id,
            {
              verified: body?.verified,
              featured: body?.featured,
              deprecated: body?.deprecated,
              blocked: body?.blocked,
              malicious: body?.malicious,
            },
            { actor: "admin", ip }
          );
          statusCode = 200;
          return sendJson(res, 200, updated);
        } catch (error) {
          if (String(error?.message || "").toLowerCase().includes("not found")) {
            statusCode = 404;
            return sendJson(res, 404, { error: "Extension not found" });
          }
          throw error;
        }
      }

      if (
        req.method === "PATCH" &&
        segments[0] === "api" &&
        segments[1] === "extensions" &&
        segments[3] === "versions" &&
        segments[5] === "flags" &&
        segments.length === 6
      ) {
        route = "/api/extensions/:id/versions/:version/flags";
        if (!adminToken) {
          statusCode = 404;
          return sendJson(res, 404, { error: "Endpoint disabled" });
        }
        const token = getBearerToken(req);
        if (token !== adminToken) {
          statusCode = 403;
          return sendJson(res, 403, { error: "Forbidden" });
        }

        const id = segments[2];
        const version = segments[4];
        const body = await readJsonBody(req);
        try {
          const updated = await store.setVersionFlags(id, version, { yanked: body?.yanked }, { actor: "admin", ip });
          statusCode = 200;
          return sendJson(res, 200, updated);
        } catch (error) {
          if (String(error?.message || "").toLowerCase().includes("not found")) {
            statusCode = 404;
            return sendJson(res, 404, { error: "Extension version not found" });
          }
          throw error;
        }
      }

      if (req.method === "GET" && url.pathname === "/api/admin/audit") {
        route = "/api/admin/audit";
        if (!adminToken) {
          statusCode = 404;
          return sendJson(res, 404, { error: "Endpoint disabled" });
        }
        const token = getBearerToken(req);
        if (token !== adminToken) {
          statusCode = 403;
          return sendJson(res, 403, { error: "Forbidden" });
        }

        const limit = Number(url.searchParams.get("limit") || "50");
        const offset = Number(url.searchParams.get("offset") || "0");
        statusCode = 200;
        return sendJson(res, 200, { entries: await store.listAuditLog({ limit, offset }) });
      }

      statusCode = 404;
      return sendJson(res, 404, { error: "Not found" });
    } catch (error) {
      if (error instanceof HttpError) {
        statusCode = error.statusCode;
        return sendJson(res, error.statusCode, { error: error.message || String(error) });
      }

      const message = String(error?.message || error);
      const lower = message.toLowerCase();
      if (lower.includes("request body too large")) {
        statusCode = 413;
        return sendJson(res, 413, { error: message });
      }
      if (
        error instanceof SyntaxError ||
        lower.includes("manifest") ||
        lower.includes("package") ||
        lower.includes("signature") ||
        lower.includes("invalid") ||
        lower.includes("disallowed") ||
        lower.includes("too many files") ||
        lower.includes("exceeds maximum")
      ) {
        statusCode = 400;
        return sendJson(res, 400, { error: message });
      }
      if (message.includes("already published")) {
        statusCode = 409;
        return sendJson(res, 409, { error: message });
      }
      statusCode = 500;
      return sendJson(res, 500, { error: message });
    } finally {
      const elapsedMs = Number(process.hrtime.bigint() - startedAt) / 1_000_000;
      const key = `${method} ${route} ${statusCode}`;
      metrics.requests.set(key, (metrics.requests.get(key) || 0) + 1);

      logger.info({
        msg: "request",
        method,
        path: req.url || "/",
        route,
        status: statusCode,
        ip,
        durationMs: Math.round(elapsedMs * 100) / 100,
      });
    }
  });

  server.on("close", () => {
    try {
      store.close();
    } catch {
      // ignore
    }
  });

  return { server, store };
}

module.exports = {
  createMarketplaceServer,
};
