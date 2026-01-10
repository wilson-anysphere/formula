import { spawn } from "node:child_process";
import { createInterface } from "node:readline";
import { fileURLToPath } from "node:url";
import path from "node:path";
import { dispatchRpc } from "./rpc.js";

function resolveRepoRoot() {
  const here = path.dirname(fileURLToPath(import.meta.url));
  // packages/python-runtime/src -> repo root
  return path.resolve(here, "../../..");
}

function resolveFormulaApiPath() {
  return path.join(resolveRepoRoot(), "python", "formula_api");
}

function withPythonPath(env, formulaApiPath) {
  const existing = env.PYTHONPATH;
  const next = existing ? `${formulaApiPath}${path.delimiter}${existing}` : formulaApiPath;
  return { ...env, PYTHONPATH: next };
}

/**
 * Minimal native Python runner for desktop (or Node tests).
 *
 * Uses a short-lived Python subprocess that executes user code and performs
 * spreadsheet operations via JSON-RPC over stdio.
 */
export class NativePythonRuntime {
  constructor(options = {}) {
    this.pythonExecutable = options.pythonExecutable ?? "python3";
    this.timeoutMs = options.timeoutMs ?? 5_000;
    this.maxMemoryBytes = options.maxMemoryBytes ?? 256 * 1024 * 1024;
    this.permissions = options.permissions ?? { filesystem: "none", network: "none" };
    this.formulaApiPath = options.formulaApiPath ?? resolveFormulaApiPath();
  }

  /**
   * Execute a Python script.
   *
   * @param {string} code
   * @param {{ api: any, timeoutMs?: number, maxMemoryBytes?: number, permissions?: any }} opts
   */
  async execute(code, opts) {
    const api = opts?.api;
    if (!api) {
      throw new Error("NativePythonRuntime.execute requires opts.api (spreadsheet bridge)");
    }

    const timeoutMs = opts?.timeoutMs ?? this.timeoutMs;
    const maxMemoryBytes = opts?.maxMemoryBytes ?? this.maxMemoryBytes;
    const permissions = opts?.permissions ?? this.permissions;

    const child = spawn(this.pythonExecutable, ["-u", "-m", "formula.runtime.stdio_runner"], {
      cwd: resolveRepoRoot(),
      env: withPythonPath(process.env, this.formulaApiPath),
      stdio: ["pipe", "pipe", "pipe"],
    });

    let timeoutId;
    const killWith = (signal) => {
      try {
        child.kill(signal);
      } catch {
        // ignore
      }
    };

    const resultPromise = new Promise((resolve, reject) => {
      let done = false;

      const stdoutLines = createInterface({ input: child.stdout });
      stdoutLines.on("line", async (line) => {
        if (done) return;
        if (!line.trim()) return;

        let msg;
        try {
          msg = JSON.parse(line);
        } catch (err) {
          done = true;
          reject(new Error(`Python runtime protocol error (non-JSON line): ${line}`));
          killWith("SIGKILL");
          return;
        }

        if (msg.type === "rpc") {
          const { id, method, params } = msg;
          try {
            const rpcResult = await dispatchRpc(api, method, params);
            child.stdin.write(JSON.stringify({ type: "rpc_response", id, result: rpcResult, error: null }) + "\n");
          } catch (rpcErr) {
            child.stdin.write(
              JSON.stringify({
                type: "rpc_response",
                id,
                result: null,
                error: rpcErr instanceof Error ? rpcErr.message : String(rpcErr),
              }) + "\n",
            );
          }
          return;
        }

        if (msg.type === "result") {
          done = true;
          resolve(msg);
          killWith("SIGTERM");
          return;
        }

        done = true;
        reject(new Error(`Python runtime protocol error (unknown message type "${msg.type}")`));
        killWith("SIGKILL");
      });

      child.on("error", (err) => {
        if (done) return;
        done = true;
        reject(err);
      });

      child.on("exit", (code, signal) => {
        if (done) return;
        done = true;
        reject(new Error(`Python process exited unexpectedly (code=${code}, signal=${signal})`));
      });

      child.stderr.on("data", () => {
        // Intentionally ignored by default; caller can attach their own listeners
        // by overriding this class or spawning with stdio pipes.
      });
    });

    if (Number.isFinite(timeoutMs) && timeoutMs > 0) {
      timeoutId = setTimeout(() => {
        killWith("SIGKILL");
      }, timeoutMs);
    }

    // Kick off execution once listeners are attached.
    child.stdin.write(
      JSON.stringify({
        type: "execute",
        code,
        permissions,
        max_memory_bytes: maxMemoryBytes,
      }) + "\n",
    );

    try {
      const result = await resultPromise;
      if (!result.success) {
        const err = new Error(result.error || "Python script failed");
        err.stack = result.traceback || err.stack;
        throw err;
      }
      return result;
    } finally {
      if (timeoutId) clearTimeout(timeoutId);
      killWith("SIGTERM");
    }
  }
}
