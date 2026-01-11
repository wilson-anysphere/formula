import { Worker } from "node:worker_threads";

import {
  PermissionDeniedError,
  SandboxMemoryLimitError,
  SandboxOutputLimitError,
  SandboxTimeoutError
} from "../errors.js";

function deserializeWorkerError(serialized, { timeoutMs } = {}) {
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

  if (serialized.code === "ERR_SCRIPT_EXECUTION_TIMEOUT" || serialized.code === "SANDBOX_TIMEOUT") {
    return new SandboxTimeoutError({ timeoutMs: serialized.timeoutMs ?? timeoutMs ?? 0 });
  }

  if (serialized.code === "SANDBOX_OUTPUT_LIMIT") {
    return new SandboxOutputLimitError({ maxBytes: serialized.maxBytes ?? 0 });
  }

  if (serialized.code === "SANDBOX_MEMORY_LIMIT") {
    return new SandboxMemoryLimitError({ memoryMb: serialized.memoryMb ?? 0, usedMb: serialized.usedMb });
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
  maxOutputBytes = 128 * 1024,
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
    metadata: { phase: "start", language: "javascript" }
  });

  return new Promise((resolve, reject) => {
    let settled = false;
    let outputBytes = 0;
    let stdout = "";
    let stderr = "";

    const timeout = setTimeout(() => {
      worker.terminate().catch(() => {});
      if (settled) return;
      settled = true;
      auditLogger?.log({
        eventType: `security.${label}.run`,
        actor: principal,
        success: false,
        metadata: { phase: "timeout", timeoutMs, language: "javascript" }
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

      if (message.type === "output") {
        const text = typeof message.text === "string" ? message.text : "";
        const chunkBytes = Buffer.byteLength(text);
        outputBytes += chunkBytes;

        if (message.stream === "stderr") {
          if (stderr.length < maxOutputBytes) stderr += text;
        } else {
          if (stdout.length < maxOutputBytes) stdout += text;
        }

        if (outputBytes > maxOutputBytes) {
          finalize(() => {
            const err = new SandboxOutputLimitError({ maxBytes: maxOutputBytes });
            err.stdout = stdout;
            err.stderr = stderr;
            auditLogger?.log({
              eventType: `security.${label}.run`,
              actor: principal,
              success: false,
              metadata: { phase: "output_limit", maxOutputBytes, language: "javascript" }
            });
            reject(err);
          });
        }
        return;
      }

      if (message.type === "limit") {
        if (message.limit === "memory") {
          finalize(() => {
            const err = new SandboxMemoryLimitError({
              memoryMb: message.memoryMb ?? memoryMb,
              usedMb: message.usedMb ?? null
            });
            err.stdout = stdout;
            err.stderr = stderr;
            auditLogger?.log({
              eventType: `security.${label}.run`,
              actor: principal,
              success: false,
              metadata: {
                phase: "memory_limit",
                memoryMb: message.memoryMb ?? memoryMb,
                usedMb: message.usedMb,
                language: "javascript"
              }
            });
            reject(err);
          });
        }
        return;
      }

      if (message.type === "result") {
        finalize(() => {
          auditLogger?.log({
            eventType: `security.${label}.run`,
            actor: principal,
            success: true,
            metadata: { phase: "complete", language: "javascript" }
          });
          resolve(message.result);
        });
        return;
      }

      if (message.type === "error") {
        finalize(() => {
          const err = deserializeWorkerError(message.error, { timeoutMs });
          err.stdout = stdout;
          err.stderr = stderr;
          const phase = err.code === "SANDBOX_TIMEOUT" ? "timeout" : "error";
          auditLogger?.log({
            eventType: `security.${label}.run`,
            actor: principal,
            success: false,
            metadata: { phase, name: err.name, message: err.message, language: "javascript", code: err.code }
          });
          reject(err);
        });
      }
    });

    worker.on("error", (err) => {
      finalize(() => {
        if (err?.code === "ERR_WORKER_OUT_OF_MEMORY") {
          const oom = new SandboxMemoryLimitError({ memoryMb, usedMb: null });
          oom.stdout = stdout;
          oom.stderr = stderr;
          auditLogger?.log({
            eventType: `security.${label}.run`,
            actor: principal,
            success: false,
            metadata: { phase: "memory_limit", memoryMb, language: "javascript", code: err.code }
          });
          reject(oom);
          return;
        }

        auditLogger?.log({
          eventType: `security.${label}.run`,
          actor: principal,
          success: false,
          metadata: { phase: "error", language: "javascript", message: err?.message, code: err?.code }
        });
        reject(err);
      });
    });

    worker.on("exit", (code) => {
      if (settled) return;
      finalize(() => {
        const err = new Error(`Sandbox worker exited unexpectedly with code ${code}`);
        err.code = "SANDBOX_WORKER_EXIT";
        reject(err);
      });
    });

    worker.postMessage({
      type: "run",
      principal,
      permissions: permissionSnapshot,
      timeoutMs,
      memoryMb,
      code
    });
  });
}
