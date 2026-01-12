import "fastify";
import type { Pool } from "pg";
import type { AppConfig } from "./config";
import type { AuthenticatedUser, SessionInfo } from "./auth/sessions";
import type { ApiKeyInfo } from "./auth/apiKeys";
import type { ScimTokenInfo } from "./auth/scim";
import type { ApiMetrics } from "./observability/metrics";
import type { AuditStreamHub } from "./audit/streamHub";

declare module "fastify" {
  interface FastifyInstance {
    db: Pool;
    config: AppConfig;
    metrics: ApiMetrics;
    auditStreamHub: AuditStreamHub;
  }

  interface FastifyRequest {
    user?: AuthenticatedUser;
    session?: SessionInfo;
    apiKey?: ApiKeyInfo;
    scimToken?: ScimTokenInfo;
    authMethod?: "session" | "api_key" | "scim";
    authOrgId?: string;
    scim?: { orgId: string };
    rateLimitScope?: string;
  }
}
