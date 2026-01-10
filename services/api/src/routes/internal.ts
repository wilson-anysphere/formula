import type { FastifyInstance } from "fastify";
import { runRetentionSweep } from "../retention";

export function registerInternalRoutes(app: FastifyInstance): void {
  app.post("/internal/retention/sweep", async (request, reply) => {
    if (!app.config.internalAdminToken) return reply.code(404).send({ error: "not_found" });
    const token = request.headers["x-internal-admin-token"];
    if (token !== app.config.internalAdminToken) return reply.code(403).send({ error: "forbidden" });

    const result = await runRetentionSweep(app.db);
    return reply.send({ ok: true, ...result });
  });
}

