import { parentPort, Worker } from "node:worker_threads";

import { buildModuleRunnerJavaScript, buildSandboxedScript, serializeError, transpileTypeScript } from "./sandbox.js";

if (!parentPort) {
  throw new Error("sandbox-worker.node must be run in a worker thread");
}

const SECURITY_WORKER_URL = new URL("../../security/src/sandbox/sandboxWorker.js", import.meta.url);

/** @type {import("node:worker_threads").Worker | null} */
let sandboxWorker = null;
let settled = false;

async function cleanup() {
  if (sandboxWorker) {
    try {
      sandboxWorker.removeAllListeners();
      await sandboxWorker.terminate();
    } catch {
      // ignore
    }
    sandboxWorker = null;
  }
  try {
    parentPort.close();
  } catch {
    // ignore
  }
}

function forwardToSandbox(message) {
  if (!sandboxWorker) return;
  sandboxWorker.postMessage(message);
}

parentPort.on("message", async (message) => {
  if (!message || typeof message !== "object") return;

  if (message.type === "cancel") {
    if (settled) return;
    settled = true;
    parentPort.postMessage({
      type: "error",
      error: { name: "AbortError", message: "Script execution cancelled" },
    });
    await cleanup();
    return;
  }

  if (message.type === "rpcResult" || message.type === "rpcError" || message.type === "event") {
    forwardToSandbox(message);
    return;
  }

  if (message.type !== "run") return;
  if (settled) return;

  const { code, activeSheetName, selection, principal, permissions, timeoutMs = 5_000, memoryMb = 64 } = message;

  try {
    const { bootstrap, ts, moduleKind, kind } = buildSandboxedScript({
      code: String(code ?? ""),
      activeSheetName: String(activeSheetName),
      selection,
    });

    const { js } = transpileTypeScript(ts, { moduleKind });
    const userScript = kind === "module" ? buildModuleRunnerJavaScript({ moduleJs: js }) : js;
    const fullScript = `${bootstrap}\n${userScript}`;

    sandboxWorker = new Worker(SECURITY_WORKER_URL, {
      type: "module",
      resourceLimits: {
        maxOldGenerationSizeMb: Math.max(16, Math.floor(memoryMb)),
        maxYoungGenerationSizeMb: Math.max(16, Math.min(64, Math.floor(memoryMb / 4))),
      },
    });
    parentPort.postMessage({
      type: "audit",
      event: {
        eventType: "scripting.sandbox.spawn",
        actor: principal,
        success: true,
        metadata: { memoryMb, resourceLimits: sandboxWorker.resourceLimits ?? null },
      },
    });

    sandboxWorker.on("message", (innerMessage) => {
      if (!innerMessage || typeof innerMessage !== "object") return;

      if (innerMessage.type === "result" || innerMessage.type === "error") {
        if (settled) return;
        settled = true;
        parentPort.postMessage(innerMessage);
        cleanup();
        return;
      }

      // console / audit / rpc messages are forwarded verbatim.
      parentPort.postMessage(innerMessage);
    });

    sandboxWorker.on("error", async (err) => {
      if (settled) return;
      settled = true;
      parentPort.postMessage({ type: "error", error: serializeError(err) });
      await cleanup();
    });

    sandboxWorker.on("exit", async (code) => {
      if (settled) return;
      settled = true;
      parentPort.postMessage({
        type: "error",
        error: { name: "SandboxWorkerExitError", message: `Sandbox worker exited with code ${code}` },
      });
      await cleanup();
    });

    sandboxWorker.postMessage({
      type: "run",
      principal,
      permissions,
      timeoutMs,
      code: fullScript,
      enableHostRpc: true,
      captureConsole: true,
      wrapAsyncIife: false,
    });
  } catch (err) {
    if (settled) return;
    settled = true;
    parentPort.postMessage({ type: "error", error: serializeError(err) });
    await cleanup();
  }
});
