import type { Pool } from "pg";

export interface OrgSiemExportState {
  orgId: string;
  lastCreatedAt: Date | null;
  lastEventId: string | null;
  updatedAt: Date;
  lastError: string | null;
  consecutiveFailures: number;
  disabledUntil: Date | null;
}

export type OrgSiemExportCursor = {
  lastCreatedAt: Date;
  lastEventId: string;
};

function toDateOrNull(value: unknown): Date | null {
  if (!value) return null;
  if (value instanceof Date) return value;
  const date = new Date(String(value));
  if (Number.isNaN(date.getTime())) return null;
  return date;
}

function toDate(value: unknown): Date {
  const date = toDateOrNull(value);
  if (!date) throw new Error(`Invalid date: ${String(value)}`);
  return date;
}

function sanitizeError(err: unknown): string {
  const message =
    err instanceof Error
      ? err.message
      : typeof err === "string"
        ? err
        : err && typeof err === "object" && "message" in err && typeof (err as any).message === "string"
          ? (err as any).message
          : String(err);

  return message.replace(/\s+/g, " ").trim().slice(0, 500);
}

function failureBackoffUntil(now: Date, consecutiveFailures: number): Date {
  const baseDelayMs = 30_000;
  const maxDelayMs = 30 * 60_000;
  const delay = Math.min(maxDelayMs, baseDelayMs * 2 ** Math.max(0, consecutiveFailures - 1));
  return new Date(now.getTime() + delay);
}

function mapRow(row: any): OrgSiemExportState {
  return {
    orgId: String(row.org_id),
    lastCreatedAt: toDateOrNull(row.last_created_at),
    lastEventId: row.last_event_id ? String(row.last_event_id) : null,
    updatedAt: toDate(row.updated_at),
    lastError: row.last_error ? String(row.last_error) : null,
    consecutiveFailures: Number(row.consecutive_failures ?? 0),
    disabledUntil: toDateOrNull(row.disabled_until)
  };
}

export class OrgSiemExportStateStore {
  constructor(private readonly db: Pool) {}

  async getOrCreate(orgId: string): Promise<OrgSiemExportState> {
    await this.db.query(
      `
        INSERT INTO org_siem_export_state (org_id)
        VALUES ($1)
        ON CONFLICT (org_id) DO NOTHING
      `,
      [orgId]
    );

    const res = await this.db.query("SELECT * FROM org_siem_export_state WHERE org_id = $1", [orgId]);
    if (res.rowCount !== 1) throw new Error(`Missing org_siem_export_state row for org ${orgId}`);
    return mapRow(res.rows[0]);
  }

  async markSuccess(orgId: string, cursor: OrgSiemExportCursor): Promise<void> {
    await this.db.query(
      `
        INSERT INTO org_siem_export_state (
          org_id,
          last_created_at,
          last_event_id,
          updated_at,
          last_error,
          consecutive_failures,
          disabled_until
        )
        VALUES ($1, $2, $3, now(), NULL, 0, NULL)
        ON CONFLICT (org_id) DO UPDATE
        SET
          last_created_at = EXCLUDED.last_created_at,
          last_event_id = EXCLUDED.last_event_id,
          updated_at = now(),
          last_error = NULL,
          consecutive_failures = 0,
          disabled_until = NULL
      `,
      [orgId, cursor.lastCreatedAt, cursor.lastEventId]
    );
  }

  async markFailure(orgId: string, error: unknown): Promise<OrgSiemExportState> {
    const current = await this.getOrCreate(orgId);
    const consecutiveFailures = current.consecutiveFailures + 1;
    const now = new Date();
    const disabledUntil = failureBackoffUntil(now, consecutiveFailures);
    const lastError = sanitizeError(error);

    await this.db.query(
      `
        UPDATE org_siem_export_state
        SET
          updated_at = now(),
          last_error = $2,
          consecutive_failures = $3,
          disabled_until = $4
        WHERE org_id = $1
      `,
      [orgId, lastError, consecutiveFailures, disabledUntil]
    );

    return {
      ...current,
      updatedAt: now,
      lastError,
      consecutiveFailures,
      disabledUntil
    };
  }
}

