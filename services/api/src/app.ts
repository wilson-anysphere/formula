import cookie from "@fastify/cookie";
import cors from "@fastify/cors";
import Fastify, { type FastifyBaseLogger, type FastifyInstance } from "fastify";
import type { Pool } from "pg";
import type { AppConfig } from "./config";
import { createLogger } from "./observability/logger";
import { createMetrics, instrumentDb, registerMetrics } from "./observability/metrics";
import { genRequestId, registerRequestId } from "./observability/request-id";
import { registerAuditRoutes } from "./routes/audit";
import { registerAuthRoutes } from "./routes/auth";
import { registerApiKeyRoutes } from "./routes/apiKeys";
import { registerDocRoutes } from "./routes/docs";
import { registerInternalRoutes } from "./routes/internal";
import { registerOrgRoutes } from "./routes/orgs";

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
    requestIdLogLabel: "requestId"
  });

  app.decorate("db", options.db);
  app.decorate("config", options.config);
  app.decorate("metrics", metrics);

  registerRequestId(app);
  registerMetrics(app, metrics);

  app.register(cookie);
  app.register(cors, {
    origin: true,
    credentials: true
  });

  app.get("/health", async () => ({ status: "ok" }));

  registerAuthRoutes(app);
  registerOrgRoutes(app);
  registerApiKeyRoutes(app);
  registerDocRoutes(app);
  registerAuditRoutes(app);
  registerInternalRoutes(app);

  return app;
}
