import { context, trace } from "@opentelemetry/api";
import pino, { type DestinationStream, type Logger } from "pino";
import { getRequestId } from "./request-id";
import { sanitizeUrlPath } from "./redaction";

const LOG_REDACTIONS = [
  // Auth/session material.
  "req.headers.authorization",
  "req.headers.cookie",
  "req.headers.set-cookie",
  "req.headers.x-api-key",
  "req.headers.x-internal-admin-token",
  "req.headers.x-saml-assertion",
  "req.headers.x-oidc-token",
  "req.body.password",
  "req.body.token",
  "req.body.sessionToken",
  "req.body.apiKey",
  "req.body.samlAssertion",
  "req.body.SAMLResponse",
  "req.body.samlResponse",
  "req.body.assertion",
  "req.body.access_token",
  "req.body.refresh_token",
  "req.body.id_token",
  "password",
  "token",
  "apiKey",
  "sessionToken",
  "samlAssertion",
  "access_token",
  "refresh_token",
  "id_token"
];

export interface CreateLoggerOptions {
  level?: string;
  stream?: DestinationStream;
}

export function createLogger(options: CreateLoggerOptions = {}): Logger {
  return pino(
    {
      level: options.level ?? process.env.LOG_LEVEL ?? "info",
      base: {
        service: "api"
      },
      redact: {
        paths: LOG_REDACTIONS,
        remove: true
      },
      serializers: {
        // Fastify's default serializers intentionally avoid logging request headers.
        // Keep the same intent while ensuring we never log query params / tokens
        // embedded in the URL.
        req(req) {
          const url = typeof req.url === "string" ? sanitizeUrlPath(req.url) : undefined;
          return {
            method: req.method,
            url,
            hostname: req.hostname,
            remoteAddress: req.remoteAddress,
            remotePort: req.remotePort
          };
        },
        res(res) {
          return { statusCode: res.statusCode };
        }
      },
      mixin(_mergeObject, _level, logger) {
        const fields: Record<string, string> = {};

        const bindings = typeof (logger as any).bindings === "function" ? (logger as any).bindings() : {};
        const hasReqId = bindings && (("requestId" in bindings) || ("reqId" in bindings));
        if (!hasReqId) {
          const requestId = getRequestId();
          if (requestId) fields.requestId = requestId;
        }

        const span = trace.getSpan(context.active());
        const spanContext = span?.spanContext();
        if (spanContext && spanContext.traceId !== "00000000000000000000000000000000") {
          fields.traceId = spanContext.traceId;
          fields.spanId = spanContext.spanId;
        }

        return fields;
      }
    },
    options.stream
  );
}
