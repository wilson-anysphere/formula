import cookie from "@fastify/cookie";
import cors from "@fastify/cors";
import formbody from "@fastify/formbody";
import Fastify, { type FastifyBaseLogger, type FastifyInstance } from "fastify";
import type { Pool } from "pg";
import type { AppConfig } from "./config";
import { registerSecurityHeaders } from "./http/securityHeaders";
import { createLogger } from "./observability/logger";
import { createMetrics, instrumentDb, registerMetrics } from "./observability/metrics";
import { genRequestId, registerRequestId } from "./observability/request-id";
import { registerAuditRoutes } from "./routes/audit";
import { registerAuthRoutes } from "./routes/auth";
import { registerApiKeyRoutes } from "./routes/apiKeys";
import { registerDocRoutes } from "./routes/docs";
import { registerDlpRoutes } from "./routes/dlp";
import { registerInternalRoutes } from "./routes/internal";
import { registerOidcProviderRoutes } from "./routes/oidcProviders";
import { registerOrgRoutes } from "./routes/orgs";
import { AuditStreamHub } from "./audit/streamHub";
import { registerSamlProviderRoutes } from "./routes/samlProviders";
import { registerScimAdminRoutes } from "./routes/scimAdmin";
import { registerScimRoutes } from "./routes/scim";
import { registerSiemRoutes } from "./routes/siem";

export interface BuildAppOptions {
  db: Pool;
  config: AppConfig;
  logger?: FastifyBaseLogger;
}

export function buildApp(options: BuildAppOptions): FastifyInstance {
  const metrics = createMetrics();
  instrumentDb(options.db, metrics);

  const app = Fastify({
    loggerInstance: (options.logger ?? createLogger()) as FastifyBaseLogger,
    genReqId: genRequestId,
    requestIdLogLabel: "requestId",
    trustProxy: options.config.trustProxy ?? false
  });

  app.decorate("db", options.db);
  app.decorate("config", options.config);
  app.decorate("metrics", metrics);
  app.decorate("auditStreamHub", new AuditStreamHub(options.db, app.log));

  registerRequestId(app);
  registerMetrics(app, metrics);
  registerSecurityHeaders(app);

  app.register(cookie);
  // Needed for SAML IdP POST bindings (application/x-www-form-urlencoded).
  app.register(formbody);

  const allowedOrigins = new Set<string>(options.config.corsAllowedOrigins ?? []);
  const normalizeOrigin = (value: string): string | null => {
    try {
      const url = new URL(value);
      if (url.origin === "null") return null;
      return url.origin;
    } catch {
      return null;
    }
  };
  app.register(cors, {
    origin(origin, cb) {
      if (!origin) return cb(null, false);
      const normalized = normalizeOrigin(origin);
      if (!normalized) return cb(null, false);
      if (!allowedOrigins.has(normalized)) return cb(null, false);
      return cb(null, normalized);
    },
    credentials: true
  });

  app.get("/health", async () => ({ status: "ok" }));

  registerAuthRoutes(app);
  registerOrgRoutes(app);
  registerSamlProviderRoutes(app);
  registerApiKeyRoutes(app);
  registerScimAdminRoutes(app);
  registerScimRoutes(app);
  registerDocRoutes(app);
  registerDlpRoutes(app);
  registerAuditRoutes(app);
  registerSiemRoutes(app);
  registerInternalRoutes(app);
  registerOidcProviderRoutes(app);

  app.addHook("onReady", async () => {
    await app.auditStreamHub.start();
  });
  app.addHook("onClose", async () => {
    await app.auditStreamHub.stop();
  });

  return app;
}
