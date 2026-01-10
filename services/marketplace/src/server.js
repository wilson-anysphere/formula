const http = require("node:http");
const crypto = require("node:crypto");

const { MarketplaceStore } = require("./store");

function sendJson(res, statusCode, body) {
  const bytes = Buffer.from(JSON.stringify(body, null, 2));
  res.writeHead(statusCode, {
    "Content-Type": "application/json; charset=utf-8",
    "Content-Length": bytes.length,
  });
  res.end(bytes);
}

function sendText(res, statusCode, text) {
  const bytes = Buffer.from(String(text));
  res.writeHead(statusCode, {
    "Content-Type": "text/plain; charset=utf-8",
    "Content-Length": bytes.length,
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

function getBearerToken(req) {
  const auth = req.headers.authorization;
  if (!auth) return null;
  const match = /^Bearer\s+(.+)$/.exec(auth);
  return match ? match[1] : null;
}

function sha256Hex(input) {
  return crypto.createHash("sha256").update(input).digest("hex");
}

function parsePath(pathname) {
  return pathname.split("/").filter(Boolean);
}

async function createMarketplaceServer({ dataDir, adminToken = null } = {}) {
  if (!dataDir) throw new Error("dataDir is required");
  const store = new MarketplaceStore({ dataDir });
  await store.init();

  const server = http.createServer(async (req, res) => {
    try {
      const url = new URL(req.url || "/", "http://localhost");
      const segments = parsePath(url.pathname);

      if (req.method === "GET" && url.pathname === "/api/health") {
        return sendJson(res, 200, { ok: true });
      }

      if (req.method === "GET" && url.pathname === "/api/search") {
        const q = url.searchParams.get("q") || "";
        const category = url.searchParams.get("category") || "";
        const limit = Number(url.searchParams.get("limit") || "20");
        const offset = Number(url.searchParams.get("offset") || "0");
        return sendJson(res, 200, store.search({ q, category, limit, offset }));
      }

      if (req.method === "GET" && segments[0] === "api" && segments[1] === "extensions" && segments.length === 3) {
        const id = segments[2];
        const ext = store.getExtension(id);
        if (!ext) return notFound(res);
        return sendJson(res, 200, ext);
      }

      if (
        req.method === "GET" &&
        segments[0] === "api" &&
        segments[1] === "extensions" &&
        segments.length === 5 &&
        segments[3] === "download"
      ) {
        const id = segments[2];
        const version = segments[4];
        const pkg = await store.getPackage(id, version);
        if (!pkg) return notFound(res);

        res.writeHead(200, {
          "Content-Type": "application/vnd.formula.extension-package",
          "Content-Length": pkg.bytes.length,
          "X-Package-Signature": pkg.signatureBase64,
          "X-Package-Sha256": pkg.sha256,
          "X-Publisher": pkg.publisher,
        });
        res.end(pkg.bytes);
        return;
      }

      if (req.method === "POST" && url.pathname === "/api/publish") {
        const token = getBearerToken(req);
        if (!token) return sendJson(res, 401, { error: "Missing Authorization token" });

        const publisherRecord = store.getPublisherByTokenSha256(sha256Hex(token));
        if (!publisherRecord) return sendJson(res, 403, { error: "Invalid token" });

        const body = await readJsonBody(req);
        if (!body?.packageBase64 || !body?.signatureBase64) {
          return sendJson(res, 400, { error: "packageBase64 and signatureBase64 are required" });
        }

        const packageBytes = Buffer.from(body.packageBase64, "base64");
        const signatureBase64 = String(body.signatureBase64);

        const published = await store.publishExtension({
          publisher: publisherRecord.publisher,
          packageBytes,
          signatureBase64,
        });

        return sendJson(res, 200, published);
      }

      if (req.method === "POST" && url.pathname === "/api/publishers/register") {
        if (!adminToken) return sendJson(res, 404, { error: "Endpoint disabled" });
        const token = getBearerToken(req);
        if (token !== adminToken) return sendJson(res, 403, { error: "Forbidden" });

        const body = await readJsonBody(req);
        if (!body?.publisher || !body?.token || !body?.publicKeyPem) {
          return sendJson(res, 400, { error: "publisher, token, publicKeyPem are required" });
        }

        await store.registerPublisher({
          publisher: body.publisher,
          tokenSha256: sha256Hex(body.token),
          publicKeyPem: body.publicKeyPem,
          verified: Boolean(body.verified),
        });

        return sendJson(res, 200, { ok: true });
      }

      if (
        req.method === "PATCH" &&
        segments[0] === "api" &&
        segments[1] === "extensions" &&
        segments[3] === "flags" &&
        segments.length === 4
      ) {
        if (!adminToken) return sendJson(res, 404, { error: "Endpoint disabled" });
        const token = getBearerToken(req);
        if (token !== adminToken) return sendJson(res, 403, { error: "Forbidden" });

        const id = segments[2];
        const body = await readJsonBody(req);
        const updated = await store.setExtensionFlags(id, {
          verified: body?.verified,
          featured: body?.featured,
        });
        return sendJson(res, 200, updated);
      }

      return notFound(res);
    } catch (error) {
      return sendJson(res, 500, { error: error.message || String(error) });
    }
  });

  return { server, store };
}

module.exports = {
  createMarketplaceServer,
};
