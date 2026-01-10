import { spawn } from "node:child_process";
import { fileURLToPath } from "node:url";

import { SandboxTimeoutError } from "../errors.js";

function deserializePythonError(payload) {
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
    child.stdout.on("data", (chunk) => {
      stdout += chunk.toString("utf8");
    });
    child.stderr.on("data", (chunk) => {
      stderr += chunk.toString("utf8");
    });

    const timeout = setTimeout(() => {
      child.kill("SIGKILL");
      auditLogger?.log({
        eventType: `security.${label}.run`,
        actor: principal,
        success: false,
        metadata: { phase: "timeout", timeoutMs, language: "python" }
      });
      reject(new SandboxTimeoutError({ timeoutMs }));
    }, timeoutMs);

    child.on("error", (err) => {
      clearTimeout(timeout);
      reject(err);
    });

    child.on("close", (code) => {
      clearTimeout(timeout);
      try {
        const parsed = JSON.parse(stdout.trim() || "{}");
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

        const err = deserializePythonError(parsed.error);
        err.stderr = stderr;
        err.exitCode = code;
        auditLogger?.log({
          eventType: `security.${label}.run`,
          actor: principal,
          success: false,
          metadata: { phase: "error", language: "python", message: err.message }
        });
        reject(err);
      } catch (error) {
        const err = new Error(`Failed to parse python sandbox output: ${error.message}`);
        err.stdout = stdout;
        err.stderr = stderr;
        err.exitCode = code;
        reject(err);
      }
    });

    const payload = {
      principal,
      permissions: permissionSnapshot,
      timeoutMs,
      memoryMb,
      code
    };

    child.stdin.end(JSON.stringify(payload));
  });
}
