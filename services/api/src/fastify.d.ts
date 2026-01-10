import "fastify";
import type { Pool } from "pg";
import type { AppConfig } from "./config";
import type { AuthenticatedUser, SessionInfo } from "./auth/sessions";

declare module "fastify" {
  interface FastifyInstance {
    db: Pool;
    config: AppConfig;
  }

  interface FastifyRequest {
    user?: AuthenticatedUser;
    session?: SessionInfo;
  }
}

