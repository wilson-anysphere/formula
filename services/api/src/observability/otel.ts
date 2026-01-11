import { diag, DiagConsoleLogger, DiagLogLevel } from "@opentelemetry/api";
import { AsyncLocalStorageContextManager } from "@opentelemetry/context-async-hooks";
import { registerInstrumentations } from "@opentelemetry/instrumentation";
import { FastifyInstrumentation } from "@opentelemetry/instrumentation-fastify";
import { HttpInstrumentation } from "@opentelemetry/instrumentation-http";
import { UndiciInstrumentation } from "@opentelemetry/instrumentation-undici";
import { OTLPTraceExporter } from "@opentelemetry/exporter-trace-otlp-http";
import { resourceFromAttributes } from "@opentelemetry/resources";
import { BatchSpanProcessor, SimpleSpanProcessor, type SpanExporter } from "@opentelemetry/sdk-trace-base";
import { NodeTracerProvider } from "@opentelemetry/sdk-trace-node";
import { SemanticResourceAttributes } from "@opentelemetry/semantic-conventions";
import { sanitizeUrlPath } from "./redaction";

export type OtelInitOptions = {
  serviceName?: string;
  spanExporter?: SpanExporter;
};

let tracerProvider: NodeTracerProvider | null = null;

function isOtelDisabled(env: NodeJS.ProcessEnv): boolean {
  const disabled = (env.OTEL_SDK_DISABLED ?? "").toLowerCase();
  if (disabled === "true") return true;

  const tracesExporter = (env.OTEL_TRACES_EXPORTER ?? "").toLowerCase();
  if (tracesExporter === "none") return true;

  return false;
}

function buildOtlpTraceExporter(env: NodeJS.ProcessEnv): OTLPTraceExporter | null {
  const explicit = env.OTEL_EXPORTER_OTLP_TRACES_ENDPOINT;
  const base = env.OTEL_EXPORTER_OTLP_ENDPOINT;
  const endpoint = explicit ?? base;
  if (!endpoint) return null;

  const trimmed = endpoint.replace(/\/+$/, "");
  const url = trimmed.endsWith("/v1/traces") ? trimmed : `${trimmed}/v1/traces`;

  return new OTLPTraceExporter({ url });
}

export function initOpenTelemetry(options: OtelInitOptions = {}): { shutdown: () => Promise<void> } {
  if (tracerProvider) {
    return { shutdown: async () => tracerProvider?.shutdown() };
  }

  const env = process.env;
  if (isOtelDisabled(env)) {
    return { shutdown: async () => {} };
  }

  const logLevel = (env.OTEL_LOG_LEVEL ?? "").toLowerCase();
  if (logLevel === "debug") {
    diag.setLogger(new DiagConsoleLogger(), DiagLogLevel.DEBUG);
  }

  const exporter = options.spanExporter ?? buildOtlpTraceExporter(env);
  const spanProcessors = exporter
    ? [options.spanExporter ? new SimpleSpanProcessor(exporter) : new BatchSpanProcessor(exporter)]
    : [];

  tracerProvider = new NodeTracerProvider({
    resource: resourceFromAttributes({
      [SemanticResourceAttributes.SERVICE_NAME]: options.serviceName ?? env.OTEL_SERVICE_NAME ?? "api"
    }),
    spanProcessors
  });

  tracerProvider.register({
    contextManager: new AsyncLocalStorageContextManager().enable()
  });

  registerInstrumentations({
    tracerProvider: tracerProvider,
    instrumentations: [
      new HttpInstrumentation({
        ignoreIncomingRequestHook: (req) => {
          const path = req.url ? sanitizeUrlPath(req.url) : "";
          return path === "/health" || path === "/metrics";
        },
        requestHook: (span, req) => {
          // `requestHook` runs for both incoming (IncomingMessage) and outgoing (ClientRequest)
          // spans. Only sanitize incoming request URLs here.
          if (!("url" in req) || typeof (req as any).url !== "string") return;

          const path = sanitizeUrlPath((req as any).url as string);
          // Old semantic conventions.
          span.setAttribute("http.target", path);

          // Newer semantic conventions (OTel >= 1.23): best-effort.
          span.setAttribute("url.path", path);
          span.setAttribute("url.query", "");

          const host = (req as any).headers?.host as unknown;
          if (typeof host === "string" && host.length > 0) {
            span.setAttribute("http.url", `http://${host}${path}`);
            span.setAttribute("url.full", `http://${host}${path}`);
          }
        }
      }),
      new FastifyInstrumentation(),
      new UndiciInstrumentation()
    ]
  });

  return {
    shutdown: async () => tracerProvider?.shutdown()
  };
}
