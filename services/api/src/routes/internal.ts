import type { FastifyInstance } from "fastify";
import { runRetentionSweep } from "../retention";

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
}
