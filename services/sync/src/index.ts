import { createSyncServer } from "./server";

function parseIntEnv(value: string | undefined, fallback: number): number {
  if (!value) return fallback;
  const parsed = Number.parseInt(value, 10);
  if (!Number.isFinite(parsed)) return fallback;
  return parsed;
}

const port = parseIntEnv(process.env.PORT, 1234);
const secret = process.env.SYNC_TOKEN_SECRET ?? "dev-sync-token-secret-change-me";

const server = createSyncServer({ port, syncTokenSecret: secret });

server.listen().catch((err) => {
  // eslint-disable-next-line no-console
  console.error(err);
  process.exitCode = 1;
});

