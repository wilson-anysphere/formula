import * as http from "node:http";
import { createAuditEvent } from "../../packages/audit-core/index.js";

import { AuditStore } from "./auditStore.js";
import { SiemConfigStore } from "./siemConfigStore.js";
import { SiemExporter } from "../../packages/security/siem/exporter.js";

function jsonResponse(res, status, payload) {
  const body = payload === undefined ? "" : JSON.stringify(payload);
  res.writeHead(status, {
    "Content-Type": "application/json",
    "Content-Length": Buffer.byteLength(body)
  });
  res.end(body);
}

function textResponse(res, status, text, headers = {}) {
  const body = text ?? "";
  res.writeHead(status, {
    "Content-Type": "text/plain",
    "Content-Length": Buffer.byteLength(body),
    ...headers
  });
  res.end(body);
}

async function readJsonBody(req) {
  const chunks = [];
  for await (const chunk of req) chunks.push(chunk);
  const raw = Buffer.concat(chunks).toString("utf8");
  if (!raw.trim()) return null;
  return JSON.parse(raw);
}

function matchRoute(method, url, pattern) {
  if (method !== pattern.method) return null;
  const match = url.pathname.match(pattern.path);
  if (!match) return null;
  return match.groups || {};
}

function isNonEmptyString(value) {
  return typeof value === "string" && value.length > 0;
}

function isActor(value) {
  return (
    value &&
    typeof value === "object" &&
    isNonEmptyString(value.type) &&
    isNonEmptyString(value.id)
  );
}

export function createApiServer(options = {}) {
  const configStore = options.configStore ?? new SiemConfigStore();
  const auditStore = options.auditStore ?? new AuditStore();
  const exportersByOrgId = options.exportersByOrgId ?? new Map();

  async function upsertExporter(orgId, config) {
    const existing = exportersByOrgId.get(orgId);
    if (existing) await existing.stop({ flush: true });

    if (!config) {
      exportersByOrgId.delete(orgId);
      return null;
    }

    const exporter = new SiemExporter(config);
    exportersByOrgId.set(orgId, exporter);
    return exporter;
  }

  const server = http.createServer(async (req, res) => {
    const url = new URL(req.url, "http://localhost");

    try {
      const paramsGetSiem = matchRoute(req.method, url, { method: "GET", path: /^\/orgs\/(?<orgId>[^/]+)\/siem$/ });
      if (paramsGetSiem) {
        const config = configStore.getSanitized(paramsGetSiem.orgId);
        if (!config) return jsonResponse(res, 404, { error: "SIEM config not found" });
        return jsonResponse(res, 200, config);
      }

      const paramsPutSiem = matchRoute(req.method, url, { method: "PUT", path: /^\/orgs\/(?<orgId>[^/]+)\/siem$/ });
      if (paramsPutSiem) {
        const body = await readJsonBody(req);
        if (!body || typeof body.endpointUrl !== "string") return jsonResponse(res, 400, { error: "endpointUrl is required" });
        configStore.set(paramsPutSiem.orgId, body);
        await upsertExporter(paramsPutSiem.orgId, body);
        return jsonResponse(res, 200, configStore.getSanitized(paramsPutSiem.orgId));
      }

      const paramsDeleteSiem = matchRoute(req.method, url, {
        method: "DELETE",
        path: /^\/orgs\/(?<orgId>[^/]+)\/siem$/
      });
      if (paramsDeleteSiem) {
        configStore.delete(paramsDeleteSiem.orgId);
        await upsertExporter(paramsDeleteSiem.orgId, null);
        res.writeHead(204);
        res.end();
        return;
      }

      const paramsPostAudit = matchRoute(req.method, url, {
        method: "POST",
        path: /^\/orgs\/(?<orgId>[^/]+)\/audit$/
      });
      if (paramsPostAudit) {
        const body = await readJsonBody(req);
        if (!body || typeof body.eventType !== "string") return jsonResponse(res, 400, { error: "eventType is required" });

        const orgId = paramsPostAudit.orgId;

        const actor = isActor(body.actor)
          ? body.actor
          : isNonEmptyString(body.userId)
            ? { type: "user", id: String(body.userId) }
            : isNonEmptyString(body.userEmail)
              ? { type: "anonymous", id: String(body.userEmail) }
              : { type: "system", id: "api" };

        const context = {
          ...(body.context && typeof body.context === "object" ? body.context : {}),
          orgId,
          userId: body.context?.userId ?? body.userId ?? null,
          userEmail: body.context?.userEmail ?? body.userEmail ?? null,
          ipAddress: body.context?.ipAddress ?? body.ipAddress ?? null,
          userAgent: body.context?.userAgent ?? body.userAgent ?? null,
          sessionId: body.context?.sessionId ?? body.sessionId ?? null
        };

        const resource =
          body.resource && typeof body.resource === "object"
            ? body.resource
            : typeof body.resourceType === "string" && body.resourceType.length > 0
              ? { type: body.resourceType, id: body.resourceId ?? null, name: body.resourceName ?? null }
              : undefined;

        const error =
          body.error && typeof body.error === "object"
            ? body.error
            : body.errorCode || body.errorMessage
              ? { code: body.errorCode ?? null, message: body.errorMessage ?? null }
              : undefined;

        const event = createAuditEvent({
          eventType: body.eventType,
          actor,
          context,
          resource,
          success: typeof body.success === "boolean" ? body.success : true,
          error,
          details: body.details ?? body.metadata ?? {},
          correlation: body.correlation
        });

        auditStore.append(orgId, event);

        const exporter = exportersByOrgId.get(orgId);
        if (exporter) exporter.enqueue(event);

        return jsonResponse(res, 202, { id: event.id });
      }

      const paramsGetAudit = matchRoute(req.method, url, { method: "GET", path: /^\/orgs\/(?<orgId>[^/]+)\/audit$/ });
      if (paramsGetAudit) {
        const limit = url.searchParams.get("limit") ? Number(url.searchParams.get("limit")) : 100;
        return jsonResponse(res, 200, { events: auditStore.list(paramsGetAudit.orgId, Number.isFinite(limit) ? limit : 100) });
      }

      const paramsStream = matchRoute(req.method, url, {
        method: "GET",
        path: /^\/orgs\/(?<orgId>[^/]+)\/audit\/stream$/
      });
      if (paramsStream) {
        res.writeHead(200, {
          "Content-Type": "text/event-stream",
          "Cache-Control": "no-cache",
          Connection: "keep-alive"
        });
        res.write(":ok\n\n");

        const emitter = auditStore.getEmitter(paramsStream.orgId);
        const handler = (event) => {
          res.write(`data: ${JSON.stringify(event)}\n\n`);
        };
        emitter.on("event", handler);
        req.on("close", () => {
          emitter.off("event", handler);
        });
        return;
      }

      return textResponse(res, 404, "Not Found");
    } catch (error) {
      return jsonResponse(res, 500, { error: error.message });
    }
  });

  return {
    server,
    configStore,
    auditStore,
    exportersByOrgId
  };
}
