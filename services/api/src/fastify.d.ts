import "fastify";
import type { Pool } from "pg";
import type { AppConfig } from "./config";
import type { AuthenticatedUser, SessionInfo } from "./auth/sessions";
import type { ApiKeyInfo } from "./auth/apiKeys";
import type { ScimTokenInfo } from "./auth/scim";
import type { ApiMetrics } from "./observability/metrics";

declare module "fastify" {
  interface FastifyInstance {
    db: Pool;
    config: AppConfig;
    metrics: ApiMetrics;
  }

  interface FastifyRequest {
    user?: AuthenticatedUser;
    session?: SessionInfo;
    apiKey?: ApiKeyInfo;
    scimToken?: ScimTokenInfo;
    authMethod?: "session" | "api_key" | "scim";
    authOrgId?: string;
    scim?: { orgId: string };
  }
}
