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

    // Stdout is reserved for protocol messages; user prints are redirected to
    // stderr by the Python runner (see `formula.runtime.stdio_runner`).
    let stderrText = "";
    child.stderr.setEncoding("utf8");
    child.stderr.on("data", (chunk) => {
      stderrText += chunk;
    });

    let timeoutId;
    let done = false;
    /** @type {((err: any) => void) | null} */
    let rejectPromise = null;
    const killWith = (signal) => {
      try {
        child.kill(signal);
      } catch {
        // ignore
      }
    };

    const resultPromise = new Promise((resolve, reject) => {
      let rejectOnce = (err) => {
        if (done) return;
        done = true;
        reject(err);
      };
      rejectPromise = rejectOnce;

      const stdoutLines = createInterface({ input: child.stdout });
      stdoutLines.on("line", (line) => {
        void (async () => {
          if (done) return;
          if (!line.trim()) return;

          let msg;
          try {
            msg = JSON.parse(line);
          } catch (err) {
            rejectOnce(new Error(`Python runtime protocol error (non-JSON line): ${line}`));
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

          rejectOnce(new Error(`Python runtime protocol error (unknown message type "${msg.type}")`));
          killWith("SIGKILL");
        })().catch((err) => {
          rejectOnce(err);
          killWith("SIGKILL");
        });
      });

      child.on("error", (err) => {
        rejectOnce(err);
      });

      // Note: `exit` can fire before the stdout stream has flushed its final
      // chunks. Rejecting immediately can race with the "result" line being read,
      // producing flaky "exited unexpectedly" errors even though a valid result
      // was printed. `close` fires after stdio streams are closed; additionally,
      // defer the rejection one tick to allow readline to emit any pending lines.
      child.on("close", (code, signal) => {
        if (done) return;
        setImmediate(() => {
          if (done) return;
          const err = new Error(`Python process exited unexpectedly (code=${code}, signal=${signal})`);
          err.stderr = stderrText;
          rejectOnce(err);
        });
      });
    });

    if (Number.isFinite(timeoutMs) && timeoutMs > 0) {
      timeoutId = setTimeout(() => {
        if (done) return;
        const err = new Error(`Python script timed out after ${timeoutMs}ms`);
        err.stderr = stderrText;
        rejectPromise?.(err);
        killWith("SIGKILL");
      }, timeoutMs);
    }

    // Kick off execution once listeners are attached.
    child.stdin.write(
      JSON.stringify({
        type: "execute",
        code,
        permissions,
        timeout_ms: timeoutMs,
        max_memory_bytes: maxMemoryBytes,
      }) + "\n",
    );

    try {
      const result = await resultPromise;
      if (!result.success) {
        const err = new Error(result.error || "Python script failed");
        const combinedStderr = result.stderr ?? stderrText;
        if (typeof combinedStderr === "string" && combinedStderr.length > 0) {
          err.stderr = combinedStderr;
        }
        if (result.traceback) {
          err.stack =
            typeof combinedStderr === "string" && combinedStderr.length > 0
              ? `${result.traceback}\n\n--- Captured stderr ---\n${combinedStderr}`
              : result.traceback;
        }
        throw err;
      }
      return { ...result, stdout: "", stderr: stderrText };
    } finally {
      if (timeoutId) clearTimeout(timeoutId);
      killWith("SIGTERM");
    }
  }
}
