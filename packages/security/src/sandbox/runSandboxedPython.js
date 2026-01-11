import { spawn } from "node:child_process";
import { fileURLToPath } from "node:url";

import {
  PermissionDeniedError,
  SandboxMemoryLimitError,
  SandboxOutputLimitError,
  SandboxTimeoutError
} from "../errors.js";

function deserializePythonError(payload, { timeoutMs, maxOutputBytes, memoryMb } = {}) {
  if (payload?.code === "PERMISSION_DENIED") {
    const err = new PermissionDeniedError({
      principal: payload.principal,
      request: payload.request,
      reason: payload.reason ?? payload.message
    });
    if (payload.stack) err.stack = payload.stack;
    return err;
  }

  if (payload?.code === "SANDBOX_OUTPUT_LIMIT") {
    const err = new SandboxOutputLimitError({ maxBytes: payload.maxBytes ?? maxOutputBytes ?? 0 });
    if (payload.stack) err.stack = payload.stack;
    return err;
  }

  if (payload?.code === "SANDBOX_MEMORY_LIMIT") {
    const err = new SandboxMemoryLimitError({
      memoryMb: payload.memoryMb ?? memoryMb ?? 0,
      usedMb: payload.usedMb ?? null
    });
    if (payload.stack) err.stack = payload.stack;
    return err;
  }

  if (payload?.code === "SANDBOX_TIMEOUT") {
    const err = new SandboxTimeoutError({ timeoutMs: payload.timeoutMs ?? timeoutMs ?? 0 });
    if (payload.stack) err.stack = payload.stack;
    return err;
  }

  const err = new Error(payload?.message ?? "Python sandbox error");
  err.name = payload?.name ?? "Error";
  if (payload?.stack) err.stack = payload.stack;
  if (payload?.code) err.code = payload.code;
  return err;
}

export async function runSandboxedPython({
  principal,
  code,
  permissionSnapshot,
  auditLogger = null,
  timeoutMs = 5_000,
  memoryMb = 128,
  maxOutputBytes = 128 * 1024,
  label = "script"
}) {
  if (!principal || typeof principal.type !== "string" || typeof principal.id !== "string") {
    throw new TypeError("runSandboxedPython requires a principal");
  }
  if (typeof code !== "string") throw new TypeError("runSandboxedPython requires code string");

  const pythonEntrypoint = fileURLToPath(new URL("./pythonSandbox.py", import.meta.url));

  auditLogger?.log({
    eventType: `security.${label}.run`,
    actor: principal,
    success: true,
    metadata: { phase: "start", language: "python" }
  });

  return new Promise((resolve, reject) => {
    const child = spawn("python3", [pythonEntrypoint], {
      stdio: ["pipe", "pipe", "pipe"],
      env: {
        ...process.env,
        PYTHONUNBUFFERED: "1"
      }
    });

    let stdout = "";
    let stderr = "";
    let stdoutBytes = 0;
    let stderrBytes = 0;
    const transportLimitBytes = Math.max(1024, maxOutputBytes * 2);
    let settled = false;
    let timeout = null;

    const finalize = (callback) => {
      if (settled) return;
      settled = true;
      if (timeout) clearTimeout(timeout);
      try {
        if (!child.killed) child.kill("SIGKILL");
      } catch {
        // ignore
      }
      callback();
    };

    child.stdout.on("data", (chunk) => {
      if (settled) return;
      const text = chunk.toString("utf8");
      stdout += text;
      stdoutBytes += Buffer.byteLength(text);

      if (stdoutBytes > transportLimitBytes) {
        finalize(() => {
          const err = new SandboxOutputLimitError({ maxBytes: maxOutputBytes });
          err.stdout = stdout;
          err.stderr = stderr;
          auditLogger?.log({
            eventType: `security.${label}.run`,
            actor: principal,
            success: false,
            metadata: { phase: "output_limit", language: "python", maxOutputBytes }
          });
          reject(err);
        });
      }
    });
    child.stderr.on("data", (chunk) => {
      if (settled) return;
      const text = chunk.toString("utf8");
      stderr += text;
      stderrBytes += Buffer.byteLength(text);

      if (stderrBytes > transportLimitBytes) {
        finalize(() => {
          const err = new SandboxOutputLimitError({ maxBytes: maxOutputBytes });
          err.stdout = stdout;
          err.stderr = stderr;
          auditLogger?.log({
            eventType: `security.${label}.run`,
            actor: principal,
            success: false,
            metadata: { phase: "output_limit", language: "python", maxOutputBytes }
          });
          reject(err);
        });
      }
    });

    timeout = setTimeout(() => {
      finalize(() => {
        auditLogger?.log({
          eventType: `security.${label}.run`,
          actor: principal,
          success: false,
          metadata: { phase: "timeout", timeoutMs, language: "python" }
        });
        reject(new SandboxTimeoutError({ timeoutMs }));
      });
    }, timeoutMs);

    child.on("error", (err) => {
      finalize(() => reject(err));
    });

    child.on("close", (code) => {
      if (settled) return;
      settled = true;
      if (timeout) clearTimeout(timeout);
      timeout = null;

      try {
        const parsed = JSON.parse(stdout.trim() || "{}");

        if (Array.isArray(parsed.audit)) {
          for (const event of parsed.audit) {
            try {
              auditLogger?.log(event);
            } catch {
              // ignore audit store failures
            }
          }
        }

        if (parsed.ok) {
          auditLogger?.log({
            eventType: `security.${label}.run`,
            actor: principal,
            success: true,
            metadata: { phase: "complete", language: "python" }
          });
          resolve(parsed.result ?? null);
          return;
        }

        const err = deserializePythonError(parsed.error, { timeoutMs, maxOutputBytes, memoryMb });
        err.stdout = parsed.stdout ?? "";
        err.stderr = parsed.stderr ?? "";
        err.transportStderr = stderr;
        err.exitCode = code;

        const phase =
          err.code === "SANDBOX_OUTPUT_LIMIT"
            ? "output_limit"
            : err.code === "SANDBOX_MEMORY_LIMIT"
              ? "memory_limit"
              : "error";

        auditLogger?.log({
          eventType: `security.${label}.run`,
          actor: principal,
          success: false,
          metadata: { phase, language: "python", message: err.message, code: err.code }
        });
        reject(err);
      } catch (error) {
        const err = new Error(`Failed to parse python sandbox output: ${error.message}`);
        err.stdout = stdout;
        err.stderr = stderr;
        err.exitCode = code;
        auditLogger?.log({
          eventType: `security.${label}.run`,
          actor: principal,
          success: false,
          metadata: { phase: "error", language: "python", message: err.message, code: "PYTHON_SANDBOX_PARSE" }
        });
        reject(err);
      }
    });

    const payload = {
      principal,
      permissions: permissionSnapshot,
      timeoutMs,
      memoryMb,
      maxOutputBytes,
      code
    };

    child.stdin.end(JSON.stringify(payload));
  });
}
