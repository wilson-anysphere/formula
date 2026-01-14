import path from "node:path";
import { loadConfig } from "./config";
import { createPool } from "./db/pool";
import { runMigrations } from "./db/migrations";
import { cleanupOidcAuthStates } from "./auth/oidc/oidc";
import { cleanupSamlAssertionReplays, cleanupSamlAuthStates, cleanupSamlRequestCache } from "./auth/saml/saml";
import { initOpenTelemetry } from "./observability/otel";
import { runRetentionSweep } from "./retention";
import { DbSiemConfigProvider } from "./siem/configProvider";
import { SiemExportWorker } from "./siem/worker";
import { closeCachedOrgTlsAgents } from "./http/tls";

const config = loadConfig();
const otel = initOpenTelemetry({ serviceName: "api" });
const pool = createPool(config.databaseUrl);

// Resolve relative to the current working directory so this works in:
// - local dev (cwd = services/api)
// - Docker image (cwd = /app)
const migrationsDir = path.resolve(process.cwd(), "migrations");

async function main(): Promise<void> {
  // docker-compose `depends_on` does not wait for Postgres readiness. Retry migrations
  // a few times so `docker-compose up` reliably brings the stack up.
  for (let attempt = 1; attempt <= 30; attempt++) {
    try {
      await runMigrations(pool, { migrationsDir });
      break;
    } catch (err) {
      if (attempt === 30) throw err;
      // eslint-disable-next-line no-console
      console.warn(`database not ready (attempt ${attempt}/30); retrying...`);
      await new Promise((resolve) => setTimeout(resolve, 1000));
    }
  }

  const { buildApp } = await import("./app");
  const app = buildApp({ db: pool, config });

  const siemWorker = new SiemExportWorker({
    db: pool,
    configProvider: new DbSiemConfigProvider(pool, config.secretStoreKeys, app.log),
    metrics: app.metrics,
    logger: app.log
  });
  siemWorker.start();
  app.addHook("onClose", async () => {
    siemWorker.stop();
    await closeCachedOrgTlsAgents();
  });

  if (config.retentionSweepIntervalMs) {
    const syncServerInternalUrl = config.syncServerInternalUrl;
    const syncServerInternalAdminToken = config.syncServerInternalAdminToken;
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

    // Fire-and-forget: we don't want retention issues to take down the API.
    const sweep = async () => {
      const result = await runRetentionSweep(pool, { onDocumentPurged });
      if (result.syncPurgesFailed && result.syncPurgesFailed > 0) {
        app.log.warn(
          { syncPurgesFailed: result.syncPurgesFailed },
          "retention_sweep_sync_purge_failures"
        );
      }
    };

    void sweep().catch((err) => {
      app.log.error({ err }, "retention sweep failed");
    });
    setInterval(() => {
      void sweep().catch((err) => {
        app.log.error({ err }, "retention sweep failed");
      });
    }, config.retentionSweepIntervalMs);
  }

  if (config.oidcAuthStateCleanupIntervalMs != null) {
    const sweep = async () => {
      const deleted = await cleanupOidcAuthStates(pool);
      if (deleted > 0) {
        app.log.debug({ deleted }, "oidc_auth_state_cleanup");
      }

      const samlStatesDeleted = await cleanupSamlAuthStates(pool);
      if (samlStatesDeleted > 0) {
        app.log.debug({ deleted: samlStatesDeleted }, "saml_auth_state_cleanup");
      }

      const samlRequestCacheDeleted = await cleanupSamlRequestCache(pool);
      if (samlRequestCacheDeleted > 0) {
        app.log.debug({ deleted: samlRequestCacheDeleted }, "saml_request_cache_cleanup");
      }

      const samlAssertionReplayDeleted = await cleanupSamlAssertionReplays(pool);
      if (samlAssertionReplayDeleted > 0) {
        app.log.debug({ deleted: samlAssertionReplayDeleted }, "saml_assertion_replay_cleanup");
      }
    };

    void sweep().catch((err) => {
      app.log.warn({ err }, "oidc_auth_state_cleanup_failed");
    });

    const timer = setInterval(() => {
      void sweep().catch((err) => {
        app.log.warn({ err }, "oidc_auth_state_cleanup_failed");
      });
    }, config.oidcAuthStateCleanupIntervalMs);

    app.addHook("onClose", async () => {
      clearInterval(timer);
    });
  }

  await app.listen({ port: config.port, host: "0.0.0.0" });
}

main().catch((err) => {
  // eslint-disable-next-line no-console
  console.error(err);
  process.exitCode = 1;
  void otel.shutdown().catch(() => {
    // Best-effort: avoid unhandled rejections during shutdown.
  });
});
