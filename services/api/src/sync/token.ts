import jwt from "jsonwebtoken";
import type { DocumentRole } from "../rbac/roles";

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
  const expiresAt = new Date(Date.now() + params.ttlSeconds * 1000);
  const token = jwt.sign(params.claims, params.secret, {
    algorithm: "HS256",
    expiresIn: params.ttlSeconds,
    audience: "formula-sync"
  });
  return { token, expiresAt };
}

