import cookie from "@fastify/cookie";
import cors from "@fastify/cors";
import Fastify, { type FastifyInstance } from "fastify";
import type { Pool } from "pg";
import type { AppConfig } from "./config";
import { registerAuditRoutes } from "./routes/audit";
import { registerAuthRoutes } from "./routes/auth";
import { registerDocRoutes } from "./routes/docs";
import { registerInternalRoutes } from "./routes/internal";
import { registerOrgRoutes } from "./routes/orgs";

export interface BuildAppOptions {
  db: Pool;
  config: AppConfig;
}

export function buildApp(options: BuildAppOptions): FastifyInstance {
  const app = Fastify({
    logger: true
  });

  app.decorate("db", options.db);
  app.decorate("config", options.config);

  app.register(cookie);
  app.register(cors, {
    origin: true,
    credentials: true
  });

  app.get("/health", async () => ({ status: "ok" }));

  registerAuthRoutes(app);
  registerOrgRoutes(app);
  registerDocRoutes(app);
  registerAuditRoutes(app);
  registerInternalRoutes(app);

  return app;
}

