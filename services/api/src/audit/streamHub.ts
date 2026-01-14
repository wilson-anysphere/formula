import type { Pool, PoolClient, Notification } from "pg";
import type { FastifyBaseLogger } from "fastify";
import { auditLogRowToAuditEvent, type AuditEvent } from "@formula/audit-core";

export type AuditPgNotification = { orgId: string; id: string };

const AUDIT_PG_CHANNEL = "formula_audit_events";
const RECENT_EVENT_ID_LIMIT = 5000;

function isUuid(value: string): boolean {
  return /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i.test(value);
}

function parseNotification(payload: string | null): AuditPgNotification | null {
  if (!payload) return null;
  try {
    const parsed = JSON.parse(payload) as unknown;
    if (!parsed || typeof parsed !== "object") return null;
    const record = parsed as Record<string, unknown>;
    const orgId = record.orgId;
    const id = record.id;
    if (typeof orgId !== "string" || !isUuid(orgId)) return null;
    if (typeof id !== "string" || !isUuid(id)) return null;
    return { orgId, id };
  } catch {
    return null;
  }
}

export class AuditStreamHub {
  private readonly subscribersByOrgId = new Map<string, Set<(event: AuditEvent) => void>>();
  private readonly recentEventIds = new Map<string, true>();
  private listenerClient: PoolClient | null = null;

  constructor(
    private readonly db: Pool,
    private readonly logger: FastifyBaseLogger
  ) {}

  subscribe(orgId: string, handler: (event: AuditEvent) => void): () => void {
    const set = this.subscribersByOrgId.get(orgId) ?? new Set();
    set.add(handler);
    this.subscribersByOrgId.set(orgId, set);

    return () => {
      const existing = this.subscribersByOrgId.get(orgId);
      if (!existing) return;
      existing.delete(handler);
      if (existing.size === 0) this.subscribersByOrgId.delete(orgId);
    };
  }

  publish(event: AuditEvent): void {
    const orgId = event.context?.orgId;
    if (!orgId || typeof orgId !== "string") return;
    if (!this.trackEventId(event.id)) return;

    const subscribers = this.subscribersByOrgId.get(orgId);
    if (!subscribers || subscribers.size === 0) return;

    for (const handler of subscribers) {
      try {
        handler(event);
      } catch (err) {
        // Don't allow a single broken subscriber to take down streaming for everyone.
        this.logger.warn({ err }, "audit stream subscriber failed");
      }
    }
  }

  async start(): Promise<void> {
    if (this.listenerClient) return;

    try {
      const client = await this.db.connect();
      this.listenerClient = client;

      client.on("notification", (msg: Notification) => {
        if (msg.channel !== AUDIT_PG_CHANNEL) return;
        void this.handlePgNotification(msg.payload ?? null).catch((err) => {
          // Best-effort: avoid unhandled rejections from EventEmitter callbacks.
          this.logger.warn({ err }, "audit stream pg notification handler failed");
        });
      });

      client.on("error", (err) => {
        this.logger.warn({ err }, "audit stream pg listener error");
      });

      await client.query(`LISTEN ${AUDIT_PG_CHANNEL}`);
    } catch (err) {
      // pg-mem (used in tests) does not fully implement LISTEN/NOTIFY; don't fail
      // the whole API if the database can't listen. Unit tests can inject events
      // via `injectNotification`.
      const message = typeof (err as any)?.message === "string" ? ((err as any).message as string) : "";
      const log = message.includes("pg-mem") ? this.logger.debug.bind(this.logger) : this.logger.warn.bind(this.logger);
      log({ err }, "audit stream LISTEN failed; continuing without pg notifications");
      await this.stop();
    }
  }

  async stop(): Promise<void> {
    const client = this.listenerClient;
    this.listenerClient = null;
    if (!client) return;

    try {
      client.removeAllListeners("notification");
      client.removeAllListeners("error");
      await client.query("UNLISTEN *");
    } catch {
      // Ignore cleanup errors; the connection may already be closed.
    } finally {
      client.release();
    }
  }

  async injectNotification(notification: AuditPgNotification): Promise<void> {
    await this.handleNotification(notification);
  }

  private async handlePgNotification(payload: string | null): Promise<void> {
    const parsed = parseNotification(payload);
    if (!parsed) return;
    await this.handleNotification(parsed);
  }

  private async handleNotification(notification: AuditPgNotification): Promise<void> {
    // Prevent dupes when the same event is published locally and also arrives via
    // LISTEN/NOTIFY on the same instance.
    if (this.recentEventIds.has(notification.id)) return;

    const columns =
      "id, org_id, user_id, user_email, event_type, resource_type, resource_id, ip_address, user_agent, session_id, success, error_code, error_message, details, created_at";
    const result = await this.db.query(
      `
        SELECT ${columns}
        FROM audit_log
        WHERE org_id = $1 AND id = $2
        LIMIT 1
      `,
      [notification.orgId, notification.id]
    );

    if (result.rowCount !== 1) return;
    const event = auditLogRowToAuditEvent(result.rows[0] as any);
    this.publish(event);
  }

  private trackEventId(id: string): boolean {
    if (!id || typeof id !== "string") return false;
    if (this.recentEventIds.has(id)) return false;
    this.recentEventIds.set(id, true);
    if (this.recentEventIds.size > RECENT_EVENT_ID_LIMIT) {
      const oldest = this.recentEventIds.keys().next().value as string | undefined;
      if (oldest) this.recentEventIds.delete(oldest);
    }
    return true;
  }
}
