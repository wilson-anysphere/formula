import { execFile, spawn } from "node:child_process";
import { access, mkdir } from "node:fs/promises";
import os from "node:os";
import path from "node:path";
import { promisify } from "node:util";
import { fileURLToPath } from "node:url";

import { Workbook } from "../workbook.js";
import { executeVbaModuleSub } from "./execute.js";

const execFileAsync = promisify(execFile);

function spawnWithInput(command, args, { cwd, input, maxBuffer = 10 * 1024 * 1024 } = {}) {
  return new Promise((resolve, reject) => {
    const child = spawn(command, args, { cwd, stdio: ["pipe", "pipe", "pipe"] });
    const stdoutChunks = [];
    const stderrChunks = [];
    let stdoutLen = 0;
    let stderrLen = 0;

    child.stdout.on("data", (chunk) => {
      stdoutChunks.push(chunk);
      stdoutLen += chunk.length;
      if (stdoutLen > maxBuffer) {
        child.kill();
        reject(new Error(`Process stdout exceeded maxBuffer (${maxBuffer})`));
      }
    });

    child.stderr.on("data", (chunk) => {
      stderrChunks.push(chunk);
      stderrLen += chunk.length;
      if (stderrLen > maxBuffer) {
        child.kill();
        reject(new Error(`Process stderr exceeded maxBuffer (${maxBuffer})`));
      }
    });

    child.on("error", reject);
    child.on("close", (code) => {
      resolve({
        code: code ?? 0,
        stdout: Buffer.concat(stdoutChunks),
        stderr: Buffer.concat(stderrChunks),
      });
    });

    if (input !== undefined) {
      child.stdin.end(input);
    } else {
      child.stdin.end();
    }
  });
}

function repoRootFromHere() {
  // packages/vba-migrate/src/vba/oracle.js -> repo root
  const here = path.dirname(fileURLToPath(import.meta.url));
  return path.resolve(here, "../../../../");
}

/**
 * @typedef {Object} VbaOracleRunRequest
 * @property {Buffer|Uint8Array|string} workbookBytes
 * @property {string} macroName
 * @property {any[]} [inputs]
 */

/**
 * @typedef {Object} VbaOracleRunResult
 * @property {boolean} ok
 * @property {Buffer} workbookAfter
 * @property {string[]} logs
 * @property {string[]} errors
 * @property {any} [report] - Raw JSON report emitted by the oracle.
 */

/**
 * Mock oracle for unit tests. It runs a *very* small JS VBA subset (same as the
 * old validator did) so tests can be deterministic without requiring Rust.
 */
export class MockOracle {
  /**
   * @param {object} [options]
   * @param {(req: VbaOracleRunRequest) => Promise<VbaOracleRunResult>|VbaOracleRunResult} [options.run]
   */
  constructor(options = {}) {
    this.run = options.run ?? null;
  }

  /**
   * @param {VbaOracleRunRequest} request
   * @returns {Promise<VbaOracleRunResult>}
   */
  async runMacro({ workbookBytes, macroName }) {
    if (this.run) {
      return await this.run({ workbookBytes, macroName });
    }

    const workbook = Workbook.fromBytes(workbookBytes);
    const payload = JSON.parse(Buffer.from(workbookBytes).toString("utf8"));
    const module = payload?.vbaModules?.[0];
    if (!module?.code) {
      throw new Error("MockOracle requires workbookBytes to include vbaModules[0].code");
    }

    const before = workbook.clone();
    executeVbaModuleSub({ workbook, module, entryPoint: macroName });

    const workbookAfter = workbook.toBytes({ vbaModules: payload?.vbaModules ?? [] });
    return {
      ok: true,
      workbookAfter,
      logs: [],
      errors: [],
      report: {
        ok: true,
        macroName,
        // Provide the same workbook schema as Rust oracle so validator can run.
        workbookAfter: JSON.parse(workbookAfter.toString("utf8")),
      },
    };
  }
}

let buildPromise = null;

async function ensureBuilt({ repoRoot, binPath }) {
  if (buildPromise) return buildPromise;
  buildPromise = (async () => {
    try {
      await access(binPath);
      return;
    } catch {
      // Build the oracle CLI once. Subsequent invocations will use the binary directly.
      const defaultGlobalCargoHome = path.resolve(os.homedir(), ".cargo");
      const envCargoHome = process.env.CARGO_HOME;
      const normalizedEnvCargoHome = envCargoHome ? path.resolve(envCargoHome) : null;
      const cargoHome =
        !envCargoHome ||
        (!process.env.CI &&
          !process.env.FORMULA_ALLOW_GLOBAL_CARGO_HOME &&
          normalizedEnvCargoHome === defaultGlobalCargoHome)
          ? path.join(repoRoot, "target", "cargo-home")
          : envCargoHome;
      await mkdir(cargoHome, { recursive: true });
      // Default to conservative parallelism. On very high core-count hosts, the Rust linker
      // (lld) can spawn many worker threads per link step; combining that with Cargo-level
      // parallelism can exceed sandbox process/thread limits and cause flaky
      // "Resource temporarily unavailable" failures.
      const cpuCount = os.cpus().length;
      const defaultJobs = cpuCount >= 64 ? "2" : "4";
      const jobs = process.env.FORMULA_CARGO_JOBS ?? process.env.CARGO_BUILD_JOBS ?? defaultJobs;
      const rayonThreads = process.env.RAYON_NUM_THREADS ?? process.env.FORMULA_RAYON_NUM_THREADS ?? jobs;
      const rustcWrapper = process.env.RUSTC_WRAPPER ?? process.env.CARGO_BUILD_RUSTC_WRAPPER ?? "";
      const rustcWorkspaceWrapper =
        process.env.RUSTC_WORKSPACE_WRAPPER ??
        process.env.CARGO_BUILD_RUSTC_WORKSPACE_WRAPPER ??
        "";
      const baseEnv = {
        ...process.env,
        CARGO_HOME: cargoHome,
        // Keep builds safe in high-core-count environments (e.g. agent sandboxes) even
        // if the caller didn't initialize via `scripts/agent-init.sh`.
        CARGO_BUILD_JOBS: jobs,
        MAKEFLAGS: process.env.MAKEFLAGS ?? `-j${jobs}`,
        CARGO_PROFILE_DEV_CODEGEN_UNITS: process.env.CARGO_PROFILE_DEV_CODEGEN_UNITS ?? jobs,
        // Rayon defaults to spawning one worker per core; cap it for multi-agent hosts unless
        // callers explicitly override it.
        RAYON_NUM_THREADS: rayonThreads,
        // Some environments configure Cargo globally with `build.rustc-wrapper`. When the
        // wrapper is unavailable/misconfigured, builds can fail even for `cargo metadata`.
        // Default to disabling any configured wrapper unless the user explicitly overrides it.
        RUSTC_WRAPPER: rustcWrapper,
        RUSTC_WORKSPACE_WRAPPER: rustcWorkspaceWrapper,
        // Cargo config can also be controlled via `CARGO_BUILD_RUSTC_WRAPPER`; set these so a
        // global config doesn't unexpectedly re-enable a flaky wrapper.
        CARGO_BUILD_RUSTC_WRAPPER: rustcWrapper,
        CARGO_BUILD_RUSTC_WORKSPACE_WRAPPER: rustcWorkspaceWrapper,
      };

      // `RUSTUP_TOOLCHAIN` overrides the repo's `rust-toolchain.toml` pin. Some environments set it
      // globally (often to `stable`), which would bypass the pinned toolchain when we fall back to
      // invoking `cargo` directly (notably on Windows).
      if (baseEnv.RUSTUP_TOOLCHAIN) {
        try {
          await access(path.join(repoRoot, "rust-toolchain.toml"));
          delete baseEnv.RUSTUP_TOOLCHAIN;
        } catch {
          // ignore
        }
      }

      let useCargoAgent = false;
      const cargoAgentPath = path.join(repoRoot, "scripts", "cargo_agent.sh");
      if (process.platform !== "win32") {
        try {
          await access(cargoAgentPath);
          useCargoAgent = true;
        } catch {
          useCargoAgent = false;
        }
      }

      const buildCommand = useCargoAgent ? "bash" : "cargo";
      const buildArgs = useCargoAgent
        ? [cargoAgentPath, "build", "-q", "-p", "formula-vba-oracle-cli"]
        : ["build", "-q", "-p", "formula-vba-oracle-cli"];
      try {
        await execFileAsync(buildCommand, buildArgs, { cwd: repoRoot, env: baseEnv });
      } catch (err) {
        // Some sandbox environments have flaky sccache daemons; retry once with sccache
        // disabled to keep tests deterministic.
        const stderr = String(err?.stderr ?? "");
        const message = String(err?.message ?? "");
        const looksLikeSccacheFailure =
          (stderr.includes("sccache") || message.includes("sccache")) &&
          (stderr.includes("Connection reset") ||
            stderr.includes("Failed to send data") ||
            message.includes("Connection reset") ||
            message.includes("Failed to send data"));
        if (!looksLikeSccacheFailure) throw err;

        const noSccacheEnv = {
          ...baseEnv,
          SCCACHE_DISABLE: "1",
          // Disable any rustc wrapper so Cargo can't end up invoking a flaky sccache daemon.
          RUSTC_WRAPPER: "",
          RUSTC_WORKSPACE_WRAPPER: "",
          CARGO_BUILD_RUSTC_WRAPPER: "",
          CARGO_BUILD_RUSTC_WORKSPACE_WRAPPER: "",
        };
        await execFileAsync(buildCommand, buildArgs, { cwd: repoRoot, env: noSccacheEnv });
      }
    }
  })();
  return buildPromise;
}

/**
 * Oracle implementation that shells out to the Rust `formula-vba-oracle-cli`
 * binary (built from the monorepo workspace).
 */
export class RustCliOracle {
  /**
   * @param {object} [options]
   * @param {string} [options.binPath] - Path to an existing oracle binary. If omitted, the oracle is built via Cargo.
   * @param {string} [options.repoRoot] - Cargo workspace root. Defaults to monorepo root.
   */
  constructor(options = {}) {
    this.repoRoot = options.repoRoot ?? repoRootFromHere();
    const defaultBin = path.join(
      this.repoRoot,
      "target",
      "debug",
      process.platform === "win32" ? "formula-vba-oracle-cli.exe" : "formula-vba-oracle-cli",
    );
    this.binPath = options.binPath ?? process.env.VBA_ORACLE_BIN ?? defaultBin;
  }

  /**
   * @param {VbaOracleRunRequest} request
   * @returns {Promise<VbaOracleRunResult>}
   */
  async runMacro({ workbookBytes, macroName, inputs }) {
    const bytes = Buffer.isBuffer(workbookBytes) ? workbookBytes : Buffer.from(workbookBytes);
    await ensureBuilt({ repoRoot: this.repoRoot, binPath: this.binPath });

    const args = ["run", "--macro", macroName];
    if (inputs && inputs.length) {
      args.push("--args", JSON.stringify(inputs));
    }

    const result = await spawnWithInput(this.binPath, args, { cwd: this.repoRoot, input: bytes });
    const stdout = result.stdout;
    const stderr = result.stderr;

    const text = stdout.toString("utf8").trim();
    const report = JSON.parse(text || "{}");
    const workbookAfter = Buffer.from(JSON.stringify(report.workbookAfter ?? report.workbook_after ?? {}), "utf8");

    const ok = Boolean(report.ok);
    const logs = Array.isArray(report.logs) ? report.logs : [];
    const errors = report.error ? [String(report.error)] : [];
    if (result.code !== 0 && errors.length === 0) {
      errors.push(stderr.toString("utf8").trim() || `Rust CLI exited with code ${result.code}`);
    }

    return { ok, workbookAfter, logs, errors, report };
  }

  /**
   * Extract macros + workbook snapshot from an `.xlsm` (or oracle JSON payload).
   * @param {{workbookBytes: Buffer|Uint8Array|string}} request
   * @returns {Promise<{ok: boolean, workbook: any, workbookBytes: Buffer, procedures: any[], error?: string, report: any}>}
   */
  async extract({ workbookBytes }) {
    const bytes = Buffer.isBuffer(workbookBytes) ? workbookBytes : Buffer.from(workbookBytes);
    await ensureBuilt({ repoRoot: this.repoRoot, binPath: this.binPath });

    const result = await spawnWithInput(this.binPath, ["extract"], { cwd: this.repoRoot, input: bytes });
    const stdout = result.stdout;

    const text = stdout.toString("utf8").trim();
    const report = JSON.parse(text || "{}");
    const workbook = report.workbook ?? {};
    return {
      ok: Boolean(report.ok),
      workbook,
      workbookBytes: Buffer.from(JSON.stringify(workbook), "utf8"),
      procedures: Array.isArray(report.procedures) ? report.procedures : [],
      error: report.error ? String(report.error) : undefined,
      report,
    };
  }
}
