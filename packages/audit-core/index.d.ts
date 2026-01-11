export type AuditSchemaVersion = 1;

export interface AuditActor {
  type: string;
  id: string;
}

export interface AuditContext {
  orgId?: string | null;
  userId?: string | null;
  userEmail?: string | null;
  ipAddress?: string | null;
  userAgent?: string | null;
  sessionId?: string | null;
}

export interface AuditResource {
  type: string;
  id?: string | null;
  name?: string | null;
}

export interface AuditError {
  code?: string | null;
  message?: string | null;
}

export interface AuditCorrelation {
  requestId?: string | null;
  traceId?: string | null;
}

export interface AuditEvent {
  schemaVersion: AuditSchemaVersion;
  id: string;
  timestamp: string;
  eventType: string;
  actor: AuditActor;
  context?: AuditContext;
  resource?: AuditResource;
  success: boolean;
  error?: AuditError;
  details?: Record<string, unknown>;
  correlation?: AuditCorrelation;
}

export interface CreateAuditEventInput extends Omit<AuditEvent, "schemaVersion" | "id" | "timestamp"> {
  id?: string;
  timestamp?: string;
  schemaVersion?: AuditSchemaVersion;
}

export interface ValidateAuditEventResult {
  valid: boolean;
  errors: string[];
}

export const AUDIT_EVENT_SCHEMA_VERSION: AuditSchemaVersion;
export const auditEventJsonSchema: Record<string, unknown>;

export function createAuditEvent(input: CreateAuditEventInput): AuditEvent;
export function validateAuditEvent(event: unknown): ValidateAuditEventResult;
export function assertAuditEvent(event: unknown): asserts event is AuditEvent;

export const DEFAULT_REDACTION_TEXT: string;
export const DEFAULT_SENSITIVE_KEY_PATTERNS: RegExp[];

export function redactValue<T = unknown>(
  value: T,
  options?: {
    redactionText?: string;
    sensitiveKeyPatterns?: RegExp[];
  }
): T;

export function redactAuditEvent<T = AuditEvent>(
  event: T,
  options?: {
    redactionText?: string;
    sensitiveKeyPatterns?: RegExp[];
  }
): T;

export function escapeCefHeaderField(value: unknown): string;
export function escapeCefExtensionValue(value: unknown): string;

export function toCef(
  event: AuditEvent,
  options?: {
    redact?: boolean;
    redactionOptions?: { redactionText?: string; sensitiveKeyPatterns?: RegExp[] };
    vendor?: string;
    product?: string;
    deviceVersion?: string;
    severity?: number;
  }
): string;

export function toLeef(
  event: AuditEvent,
  options?: {
    redact?: boolean;
    redactionOptions?: { redactionText?: string; sensitiveKeyPatterns?: RegExp[] };
    vendor?: string;
    product?: string;
    productVersion?: string;
    delimiter?: string;
  }
): string;

export function serializeBatch(
  events: AuditEvent[],
  options?: {
    format?: "json" | "cef" | "leef";
    redact?: boolean;
    redactionOptions?: { redactionText?: string; sensitiveKeyPatterns?: RegExp[] };
  }
): { contentType: string; body: Buffer };

export function auditEventToSqliteRow(event: AuditEvent): {
  id: string;
  ts: number;
  timestamp: string;
  eventType: string;
  actorType: string;
  actorId: string;
  orgId: string | null;
  userId: string | null;
  userEmail: string | null;
  ipAddress: string | null;
  userAgent: string | null;
  sessionId: string | null;
  resourceType: string | null;
  resourceId: string | null;
  resourceName: string | null;
  success: 0 | 1;
  errorCode: string | null;
  errorMessage: string | null;
  details: string;
  requestId: string | null;
  traceId: string | null;
};

export function buildPostgresAuditLogInsert(event: AuditEvent): {
  text: string;
  values: unknown[];
};

export function retentionCutoffMs(now: number, retentionDays: number): number;

