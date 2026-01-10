import pino from "pino";

export function createLogger(level: string) {
  return pino({
    level,
    base: {
      service: "sync-server",
    },
    redact: {
      paths: ["req.headers.authorization", "token", "authToken"],
      remove: true,
    },
  });
}

