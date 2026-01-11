import crypto from "node:crypto";
import { AsyncLocalStorage } from "node:async_hooks";
import type { IncomingMessage } from "node:http";
import type { FastifyInstance } from "fastify";

export const REQUEST_ID_HEADER = "x-request-id";

const requestIdStorage = new AsyncLocalStorage<string>();

export function getRequestId(): string | undefined {
  return requestIdStorage.getStore();
}

function normalizeIncomingRequestId(value: unknown): string | null {
  if (typeof value === "string") {
    const trimmed = value.trim();
    // Keep it small (avoid log spam) and ensure it can't break JSON logs.
    if (!trimmed) return null;
    if (trimmed.length > 200) return null;
    if (/[^\w.\-:@]/.test(trimmed)) return null;
    return trimmed;
  }

  if (Array.isArray(value) && value.length > 0) {
    return normalizeIncomingRequestId(value[0]);
  }

  return null;
}

export function genRequestId(req: IncomingMessage): string {
  const incoming = normalizeIncomingRequestId(req.headers[REQUEST_ID_HEADER]);
  return incoming ?? crypto.randomUUID();
}

export function registerRequestId(app: FastifyInstance): void {
  app.addHook("onRequest", (request, reply, done) => {
    const requestId = request.id || crypto.randomUUID();
    reply.header(REQUEST_ID_HEADER, requestId);
    requestIdStorage.enterWith(requestId);
    done();
  });
}

