import type { DocumentRole } from "../rbac/roles";

type JwtModule = typeof import("jsonwebtoken");

function loadJwt(): JwtModule {
  // Use runtime require so Vitest/Vite SSR doesn't try to transform jsonwebtoken
  // (a CJS package that relies on relative requires like `./decode`).
  // eslint-disable-next-line @typescript-eslint/no-var-requires
  return require("jsonwebtoken") as JwtModule;
}

export interface SyncTokenClaims {
  sub: string;
  docId: string;
  orgId: string;
  role: DocumentRole;
  sessionId?: string;
}

export function signSyncToken(params: {
  secret: string;
  ttlSeconds: number;
  claims: SyncTokenClaims;
}): { token: string; expiresAt: Date } {
  const jwt = loadJwt();
  const expiresAt = new Date(Date.now() + params.ttlSeconds * 1000);
  const token = jwt.sign(params.claims, params.secret, {
    algorithm: "HS256",
    expiresIn: params.ttlSeconds,
    audience: "formula-sync"
  });
  return { token, expiresAt };
}
