import type { Pool } from "pg";
import { SpanStatusCode, trace } from "@opentelemetry/api";
import { createAuditEvent, writeAuditEvent } from "../audit/audit";
import type { ApiMetrics } from "../observability/metrics";
import {
  assertOutboundRegionAllowed,
  resolvePrimaryStorageRegion,
  DataResidencyViolationError
} from "../policies/dataResidency";
import { fetchNextAuditEvents, type AuditCursor } from "./auditSource";
import { sendSiemBatch } from "./sender";
import { OrgSiemExportStateStore } from "./stateStore";
import type { EnabledSiemOrg, SiemConfigProvider } from "./configProvider";
import type { OrgTlsPolicy } from "../http/tls";

type OrgDataResidencyPolicy = {
  region: string;
  allowedRegions: unknown;
  allowCrossRegionProcessing: boolean;
};

async function loadOrgSiemPolicies(
  db: Pool,
  orgId: string
): Promise<{ tls: OrgTlsPolicy; residency: OrgDataResidencyPolicy }> {
  const res = await db.query<{
    certificate_pinning_enabled: boolean;
    certificate_pins: unknown;
    data_residency_region: string;
    data_residency_allowed_regions: unknown;
    allow_cross_region_processing: boolean;
  }>(
    `
      SELECT
        certificate_pinning_enabled,
        certificate_pins,
        data_residency_region,
        data_residency_allowed_regions,
        allow_cross_region_processing
      FROM org_settings
      WHERE org_id = $1
    `,
    [orgId]
  );
  if (res.rowCount !== 1) {
    throw new Error(`Missing org_settings row for org ${orgId}`);
  }

  const row = res.rows[0] as any;

  return {
    tls: {
      certificatePinningEnabled: Boolean(row.certificate_pinning_enabled),
      certificatePins: row.certificate_pins
    },
    residency: {
      region: String(row.data_residency_region ?? "us"),
      allowedRegions: row.data_residency_allowed_regions,
      allowCrossRegionProcessing: Boolean(row.allow_cross_region_processing)
    }
  };
}

export interface SiemExportWorkerOptions {
  db: Pool;
  configProvider: SiemConfigProvider;
  metrics: ApiMetrics;
  logger: {
    info: (...args: any[]) => void;
    warn: (...args: any[]) => void;
    error: (...args: any[]) => void;
    debug?: (...args: any[]) => void;
  };
  pollIntervalMs?: number;
  maxConcurrentOrgs?: number;
  maxBatchesPerOrgRun?: number;
  defaultBatchSize?: number;
}

async function runWithConcurrencyLimit<T>(
  items: T[],
  limit: number,
  fn: (item: T) => Promise<void>
): Promise<void> {
  if (items.length === 0) return;
  const queue = items.slice();

  const workers = Array.from({ length: Math.min(limit, queue.length) }, async () => {
    while (queue.length > 0) {
      const next = queue.shift();
      if (!next) break;
      await fn(next);
    }
  });

  await Promise.all(workers);
}

export class SiemExportWorker {
  private readonly stateStore: OrgSiemExportStateStore;
  private readonly inFlightOrgs = new Set<string>();
  private timer: NodeJS.Timeout | null = null;
  private cycleInFlight = false;
  private readonly tracer = trace.getTracer("api.siem");

  private readonly pollIntervalMs: number;
  private readonly maxConcurrentOrgs: number;
  private readonly maxBatchesPerOrgRun: number;
  private readonly defaultBatchSize: number;

  constructor(private readonly options: SiemExportWorkerOptions) {
    this.stateStore = new OrgSiemExportStateStore(options.db);
    this.pollIntervalMs = options.pollIntervalMs ?? 10_000;
    this.maxConcurrentOrgs = options.maxConcurrentOrgs ?? 3;
    this.maxBatchesPerOrgRun = options.maxBatchesPerOrgRun ?? 10;
    this.defaultBatchSize = options.defaultBatchSize ?? 250;
  }

  start(): void {
    if (this.timer) return;
    this.timer = setInterval(() => {
      void this.tick().catch((err) => {
        this.options.logger.error({ err }, "siem_export_tick_failed");
      });
    }, this.pollIntervalMs);
    this.timer.unref?.();
  }

  stop(): void {
    if (!this.timer) return;
    clearInterval(this.timer);
    this.timer = null;
  }

  async tick(): Promise<void> {
    if (this.cycleInFlight) return;
    this.cycleInFlight = true;
    try {
      const enabledOrgs = await this.options.configProvider.listEnabledOrgs();
      const orgsToProcess = enabledOrgs.filter((org) => !this.inFlightOrgs.has(org.orgId));

      let lagMaxSeconds = 0;

      await runWithConcurrencyLimit(orgsToProcess, this.maxConcurrentOrgs, async (org) => {
        this.inFlightOrgs.add(org.orgId);
        try {
          const lagSeconds = await this.exportOrg(org);
          if (typeof lagSeconds === "number") lagMaxSeconds = Math.max(lagMaxSeconds, lagSeconds);
        } finally {
          this.inFlightOrgs.delete(org.orgId);
        }
      });

      this.options.metrics.siemExportLagSeconds.set(lagMaxSeconds);
    } finally {
      this.cycleInFlight = false;
    }
  }

  private async exportOrg(org: EnabledSiemOrg): Promise<number | null> {
    return this.tracer.startActiveSpan(
      "siem.export.org",
      { attributes: { orgId: org.orgId } },
      async (span) => {
        const now = new Date();
        let residencyContext:
          | {
              dataResidencyRegion: string;
              allowCrossRegionProcessing: boolean;
              siemDataRegion: string;
              siemConfiguredRegion: string | null;
            }
          | null = null;
        try {
          const state = await this.stateStore.getOrCreate(org.orgId);
          if (state.disabledUntil && state.disabledUntil.getTime() > now.getTime()) {
            this.options.metrics.siemBatchesTotal.inc({ status: "disabled" });
            span.setStatus({ code: SpanStatusCode.OK });
            return null;
          }

          const { tls: tlsPolicy, residency } = await loadOrgSiemPolicies(this.options.db, org.orgId);

          const primaryRegion = resolvePrimaryStorageRegion({
            region: residency.region,
            allowedRegions: residency.allowedRegions
          });
          const dataRegion = org.config.dataRegion ?? primaryRegion;

          residencyContext = {
            dataResidencyRegion: residency.region,
            allowCrossRegionProcessing: residency.allowCrossRegionProcessing,
            siemDataRegion: dataRegion,
            siemConfiguredRegion: org.config.dataRegion ?? null
          };

          assertOutboundRegionAllowed({
            orgId: org.orgId,
            requestedRegion: dataRegion,
            operation: "siem.export.send",
            region: residency.region,
            allowedRegions: residency.allowedRegions,
            allowCrossRegionProcessing: residency.allowCrossRegionProcessing
          });

          let cursor: AuditCursor = {
            lastCreatedAt: state.lastCreatedAt,
            lastEventId: state.lastEventId
          };

          let batches = 0;
          let lastLagSeconds: number | null = null;

          while (batches < this.maxBatchesPerOrgRun) {
            const batchSize = Math.max(1, Math.min(1000, org.config.batchSize ?? this.defaultBatchSize));

            const events = await this.tracer.startActiveSpan(
              "siem.export.fetch_batch",
              { attributes: { orgId: org.orgId, batchSize } },
              async (fetchSpan) => {
                try {
                  const result = await fetchNextAuditEvents(this.options.db, org.orgId, cursor, batchSize);
                  fetchSpan.setStatus({ code: SpanStatusCode.OK });
                  return result;
                } catch (err) {
                  fetchSpan.recordException(err as Error);
                  fetchSpan.setStatus({ code: SpanStatusCode.ERROR });
                  throw err;
                } finally {
                  fetchSpan.end();
                }
              }
            );

            if (events.length === 0) {
              if (batches === 0) this.options.metrics.siemBatchesTotal.inc({ status: "noop" });
              break;
            }

            const batchStart = process.hrtime.bigint();

            await this.tracer.startActiveSpan(
              "siem.export.send_batch",
              { attributes: { orgId: org.orgId, events: events.length } },
              async (sendSpan) => {
                try {
                  const payload = events.map(({ event }) => event);
                  await sendSiemBatch(org.config, payload, { tls: tlsPolicy });
                  sendSpan.setStatus({ code: SpanStatusCode.OK });
                } catch (err) {
                  sendSpan.recordException(err as Error);
                  sendSpan.setStatus({ code: SpanStatusCode.ERROR });
                  throw err;
                } finally {
                  sendSpan.end();
                }
              }
            );

            const durationSeconds = Number(process.hrtime.bigint() - batchStart) / 1e9;
            this.options.metrics.siemBatchDurationSeconds.observe(durationSeconds);
            this.options.metrics.siemBatchesTotal.inc({ status: "ok" });
            this.options.metrics.siemEventsTotal.inc({ status: "ok" }, events.length);

            const last = events[events.length - 1]!;
            await this.stateStore.markSuccess(org.orgId, {
              lastCreatedAt: last.createdAt,
              lastEventId: last.event.id
            });

            cursor = {
              lastCreatedAt: last.createdAt,
              lastEventId: last.event.id
            };

            lastLagSeconds = Math.max(0, (Date.now() - last.createdAt.getTime()) / 1000);
            batches += 1;
          }

          span.setStatus({ code: SpanStatusCode.OK });
          return lastLagSeconds;
        } catch (err) {
          if (err instanceof DataResidencyViolationError) {
            this.options.metrics.dataResidencyBlockedTotal.inc({ operation: err.operation });
            try {
              await writeAuditEvent(
                this.options.db,
                createAuditEvent({
                  eventType: "org.data_residency.blocked",
                  actor: { type: "system", id: "siem_export_worker" },
                  context: { orgId: org.orgId },
                  resource: { type: "integration", id: "siem", name: "siem" },
                  success: false,
                  error: { code: "data_residency_violation", message: err.message },
                  details: {
                    operation: err.operation,
                    requestedRegion: err.requestedRegion,
                    allowedRegions: err.allowedRegions,
                    ...(residencyContext ?? {})
                  }
                })
              );
            } catch (auditErr) {
              this.options.logger.warn(
                { err: auditErr, orgId: org.orgId },
                "data_residency_blocked_audit_failed"
              );
            }
          }

          this.options.metrics.siemBatchesTotal.inc({ status: "error" });
          await this.stateStore.markFailure(org.orgId, err);
          this.options.logger.warn({ err, orgId: org.orgId }, "siem_export_org_failed");
          span.recordException(err as Error);
          span.setStatus({ code: SpanStatusCode.ERROR });
          return null;
        } finally {
          span.end();
        }
      }
    );
  }
}
