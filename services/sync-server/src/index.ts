import "dotenv/config";

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

await server.start();

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
