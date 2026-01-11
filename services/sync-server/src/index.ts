import "dotenv/config";

import { writeSync } from "node:fs";

import { createLogger } from "./logger.js";
import { loadConfigFromEnv } from "./config.js";
import { createSyncServer } from "./server.js";

const config = loadConfigFromEnv();
const logger = createLogger(config.logLevel);

if (config.auth.mode === "opaque" && config.auth.token === "dev-token") {
  logger.warn(
    {
      env: process.env.NODE_ENV ?? "development",
    },
    "Using default dev token (dev-token). Set SYNC_SERVER_AUTH_TOKEN or SYNC_SERVER_JWT_SECRET for production."
  );
}

const server = createSyncServer(config, logger);

try {
  await server.start();
} catch (err) {
  const message = err instanceof Error ? err.message : String(err);
  // Ensure the startup error is visible even if LOG_LEVEL=silent.
  writeSync(2, `sync-server failed to start: ${message}\n`);
  try {
    logger.error({ err }, "startup_failed");
  } catch {
    // ignore
  }
  process.exit(1);
}

let shuttingDown = false;

const shutdown = async (signal: string) => {
  if (shuttingDown) return;
  shuttingDown = true;

  logger.info({ signal }, "shutting_down");
  try {
    await server.stop();
  } catch (err) {
    logger.error({ err }, "shutdown_failed");
    process.exitCode = 1;
  }
};

process.on("SIGINT", () => void shutdown("SIGINT"));
process.on("SIGTERM", () => void shutdown("SIGTERM"));
