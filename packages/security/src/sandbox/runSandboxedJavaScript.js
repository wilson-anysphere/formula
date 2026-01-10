import { Worker } from "node:worker_threads";

import { PermissionDeniedError, SandboxTimeoutError } from "../errors.js";

function deserializeWorkerError(serialized) {
  if (!serialized || typeof serialized !== "object") return new Error("Unknown sandbox error");

  if (serialized.code === "PERMISSION_DENIED") {
    const err = new PermissionDeniedError({
      principal: serialized.principal,
      request: serialized.request,
      reason: serialized.reason ?? serialized.message
    });
    if (serialized.stack) err.stack = serialized.stack;
    return err;
  }

  const err = new Error(serialized.message ?? "Sandbox error");
  err.name = serialized.name ?? "Error";
  if (serialized.code) err.code = serialized.code;
  if (serialized.stack) err.stack = serialized.stack;
  return err;
}

export async function runSandboxedJavaScript({
  principal,
  code,
  permissionSnapshot,
  auditLogger = null,
  timeoutMs = 5_000,
  memoryMb = 64,
  label = "script"
}) {
  if (!principal || typeof principal.type !== "string" || typeof principal.id !== "string") {
    throw new TypeError("runSandboxedJavaScript requires a principal");
  }
  if (typeof code !== "string") throw new TypeError("runSandboxedJavaScript requires code string");

  const workerUrl = new URL("./sandboxWorker.js", import.meta.url);

  const worker = new Worker(workerUrl, {
    type: "module",
    resourceLimits: {
      // V8 memory limits are approximate; this is a best-effort baseline.
      maxOldGenerationSizeMb: Math.max(16, memoryMb),
      maxYoungGenerationSizeMb: Math.max(16, Math.min(64, Math.floor(memoryMb / 4)))
    }
  });

  auditLogger?.log({
    eventType: `security.${label}.run`,
    actor: principal,
    success: true,
    metadata: { phase: "start" }
  });

  return new Promise((resolve, reject) => {
    let settled = false;

    const timeout = setTimeout(() => {
      worker.terminate().catch(() => {});
      if (settled) return;
      settled = true;
      auditLogger?.log({
        eventType: `security.${label}.run`,
        actor: principal,
        success: false,
        metadata: { phase: "timeout", timeoutMs }
      });
      reject(new SandboxTimeoutError({ timeoutMs }));
    }, timeoutMs);

    const finalize = async (callback) => {
      clearTimeout(timeout);
      if (settled) return;
      settled = true;
      try {
        await worker.terminate();
      } catch {
        // Ignore termination errors; the worker may already be gone.
      }
      callback();
    };

    worker.on("message", (message) => {
      if (!message || typeof message !== "object") return;

      if (message.type === "audit") {
        try {
          auditLogger?.log(message.event);
        } catch {
          // Audit logging should not crash sandbox execution.
        }
        return;
      }

      if (message.type === "result") {
        finalize(() => {
          auditLogger?.log({
            eventType: `security.${label}.run`,
            actor: principal,
            success: true,
            metadata: { phase: "complete" }
          });
          resolve(message.result);
        });
        return;
      }

      if (message.type === "error") {
        finalize(() => {
          const err = deserializeWorkerError(message.error);
          auditLogger?.log({
            eventType: `security.${label}.run`,
            actor: principal,
            success: false,
            metadata: { phase: "error", name: err.name, message: err.message }
          });
          reject(err);
        });
      }
    });

    worker.on("error", (err) => {
      finalize(() => reject(err));
    });

    worker.postMessage({
      type: "run",
      principal,
      permissions: permissionSnapshot,
      timeoutMs,
      code
    });
  });
}
