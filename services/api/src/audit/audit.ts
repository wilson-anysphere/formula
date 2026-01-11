import type { Pool } from "pg";
import { context, trace } from "@opentelemetry/api";
import { getRequestId } from "../observability/request-id";

import {
  assertAuditEvent,
  buildPostgresAuditLogInsert,
  createAuditEvent,
  type AuditEvent
} from "@formula/audit-core";

export { createAuditEvent };
export type { AuditEvent };

function currentTraceId(): string | undefined {
  const span = trace.getSpan(context.active());
  const spanContext = span?.spanContext();
  if (!spanContext) return undefined;
  if (spanContext.traceId === "00000000000000000000000000000000") return undefined;
  return spanContext.traceId;
}

function enrichCorrelation(event: AuditEvent): AuditEvent {
  const requestId = getRequestId();
  const traceId = currentTraceId();

  if (!requestId && !traceId) return event;

  const correlation = {
    requestId: event.correlation?.requestId ?? requestId ?? null,
    traceId: event.correlation?.traceId ?? traceId ?? null
  };

  const enriched = {
    ...event,
    correlation
  };
  assertAuditEvent(enriched);
  return enriched;
}

export async function writeAuditEvent(pool: Pool, event: AuditEvent): Promise<void> {
  const enriched = enrichCorrelation(event);
  const { text, values } = buildPostgresAuditLogInsert(enriched);
  await pool.query(text, values);
}
