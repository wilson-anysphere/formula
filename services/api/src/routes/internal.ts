import type { FastifyInstance } from "fastify";
import { z } from "zod";
import { runRetentionSweep } from "../retention";
import { verifySyncToken } from "../sync/token";
import type { DocumentRole } from "../rbac/roles";

export function registerInternalRoutes(app: FastifyInstance): void {
  app.post("/internal/retention/sweep", async (request, reply) => {
    if (!app.config.internalAdminToken) return reply.code(404).send({ error: "not_found" });
    const token = request.headers["x-internal-admin-token"];
    if (token !== app.config.internalAdminToken) return reply.code(403).send({ error: "forbidden" });

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

  const SyncIntrospectBody = z.object({
    token: z.string().min(1),
    docId: z.string().min(1)
  });

  app.post("/internal/sync/introspect", async (request, reply) => {
    if (!app.config.internalAdminToken) return reply.code(404).send({ error: "not_found" });
    const internalToken = request.headers["x-internal-admin-token"];
    if (internalToken !== app.config.internalAdminToken) {
      return reply.code(403).send({ ok: false, error: "forbidden" });
    }

    const parsed = SyncIntrospectBody.safeParse(request.body);
    if (!parsed.success) {
      app.metrics.syncTokenIntrospectFailuresTotal.inc();
      return reply.code(403).send({ ok: false, error: "forbidden" });
    }

    const { token: syncToken, docId } = parsed.data;

    let claims: { sub: string; docId: string; orgId: string; role: DocumentRole; sessionId?: string };
    try {
      claims = verifySyncToken({ token: syncToken, secret: app.config.syncTokenSecret });
    } catch {
      app.metrics.syncTokenIntrospectFailuresTotal.inc();
      return reply.code(403).send({ ok: false, error: "forbidden" });
    }

    if (claims.docId !== docId) {
      app.metrics.syncTokenIntrospectFailuresTotal.inc();
      return reply.code(403).send({ ok: false, error: "forbidden" });
    }

    if (claims.sessionId) {
      const sessionRes = await app.db.query(
        `
          SELECT 1
          FROM sessions
          WHERE id = $1
            AND user_id = $2
            AND revoked_at IS NULL
            AND expires_at > now()
          LIMIT 1
        `,
        [claims.sessionId, claims.sub]
      );

      if (sessionRes.rowCount !== 1) {
        app.metrics.syncTokenIntrospectFailuresTotal.inc();
        return reply.code(403).send({ ok: false, error: "forbidden" });
      }
    }

    const membershipRes = await app.db.query(
      `
        SELECT d.org_id, dm.role
        FROM documents d
        JOIN document_members dm
          ON dm.document_id = d.id AND dm.user_id = $2
        WHERE d.id = $1
        LIMIT 1
      `,
      [docId, claims.sub]
    );

    if (membershipRes.rowCount !== 1) {
      app.metrics.syncTokenIntrospectFailuresTotal.inc();
      return reply.code(403).send({ ok: false, error: "forbidden" });
    }

    const row = membershipRes.rows[0] as { org_id: string; role: DocumentRole };
    if (row.org_id !== claims.orgId) {
      app.metrics.syncTokenIntrospectFailuresTotal.inc();
      return reply.code(403).send({ ok: false, error: "forbidden" });
    }

    return reply.send({
      ok: true,
      userId: claims.sub,
      orgId: row.org_id,
      role: row.role,
      sessionId: claims.sessionId
    });
  });
}
