import type { FastifyInstance, FastifyReply, FastifyRequest } from "fastify";
import { z } from "zod";
import { runRetentionSweep } from "../retention";
import { introspectSyncToken } from "../sync/introspection";

export function registerInternalRoutes(app: FastifyInstance): void {
  const requireInternalAdminToken = (request: FastifyRequest, reply: FastifyReply): boolean => {
    if (!app.config.internalAdminToken) {
      void reply.code(404).send({ error: "not_found" });
      return false;
    }
    const token = request.headers["x-internal-admin-token"];
    if (token !== app.config.internalAdminToken) {
      void reply.code(403).send({ error: "forbidden" });
      return false;
    }
    return true;
  };

  app.post("/internal/retention/sweep", async (request, reply) => {
    if (!requireInternalAdminToken(request, reply)) return;

    const syncServerInternalUrl = app.config.syncServerInternalUrl;
    const syncServerInternalAdminToken = app.config.syncServerInternalAdminToken;
    const onDocumentPurged =
      syncServerInternalUrl && syncServerInternalAdminToken
        ? async ({ orgId, docId }: { orgId: string; docId: string }) => {
            const purgeUrl = new URL(
              `/internal/docs/${encodeURIComponent(docId)}`,
              syncServerInternalUrl
            ).toString();

            let res: Response;
            try {
              res = await fetch(purgeUrl, {
                method: "DELETE",
                headers: {
                  "x-internal-admin-token": syncServerInternalAdminToken
                },
                signal: AbortSignal.timeout(5000)
              });
            } catch (err) {
              app.log.warn({ err, orgId, docId }, "sync_server_purge_request_failed");
              throw err;
            }

            if (!res.ok) {
              app.log.warn({ orgId, docId, status: res.status }, "sync_server_purge_failed");
              throw new Error(`sync-server purge failed (${res.status})`);
            }
          }
        : undefined;

    const result = await runRetentionSweep(app.db, { onDocumentPurged });
    if (result.syncPurgesFailed && result.syncPurgesFailed > 0) {
      app.log.warn(
        { syncPurgesFailed: result.syncPurgesFailed },
        "retention_sweep_sync_purge_failures"
      );
    }
    return reply.send({ ok: true, ...result });
  });

  const IntrospectBody = z.object({
    token: z.string().min(1),
    docId: z.string().min(1),
    clientIp: z.string().min(1).optional(),
    userAgent: z.string().min(1).optional()
  });

  app.post("/internal/sync/introspect", async (request, reply) => {
    if (!requireInternalAdminToken(request, reply)) return;

    const parsed = IntrospectBody.safeParse(request.body);
    if (!parsed.success) {
      app.metrics.syncTokenIntrospectFailuresTotal.inc();
      return reply.code(400).send({ error: "invalid_request" });
    }

    const result = await introspectSyncToken(app.db, {
      secret: app.config.syncTokenSecret,
      token: parsed.data.token,
      docId: parsed.data.docId,
      clientIp: parsed.data.clientIp ?? null,
      userAgent: parsed.data.userAgent ?? null
    });

    if (!result.active) {
      app.metrics.syncTokenIntrospectFailuresTotal.inc();
      return reply.send({
        ok: false,
        active: false,
        error: "forbidden",
        reason: result.reason,
        userId: result.userId ?? null,
        orgId: result.orgId ?? null,
        role: result.role ?? null,
        sessionId: result.sessionId ?? null
      });
    }

    return reply.send({
      ok: true,
      active: true,
      userId: result.userId,
      orgId: result.orgId,
      role: result.role,
      sessionId: result.sessionId ?? null
    });
  });
}
