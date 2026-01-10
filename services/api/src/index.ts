import path from "node:path";
import { buildApp } from "./app";
import { loadConfig } from "./config";
import { createPool } from "./db/pool";
import { runMigrations } from "./db/migrations";
import { runRetentionSweep } from "./retention";

const config = loadConfig();
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

  const app = buildApp({ db: pool, config });

  if (config.retentionSweepIntervalMs) {
    // Fire-and-forget: we don't want retention issues to take down the API.
    void runRetentionSweep(pool).catch((err) => {
      app.log.error({ err }, "retention sweep failed");
    });
    setInterval(() => {
      void runRetentionSweep(pool).catch((err) => {
        app.log.error({ err }, "retention sweep failed");
      });
    }, config.retentionSweepIntervalMs);
  }

  await app.listen({ port: config.port, host: "0.0.0.0" });
}

main().catch((err) => {
  // eslint-disable-next-line no-console
  console.error(err);
  process.exitCode = 1;
});
