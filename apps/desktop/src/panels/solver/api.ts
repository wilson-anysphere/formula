import type { SolverConfig, SolverOutcome, SolverProgress } from "./types";

/**
 * Desktop solver IPC stub.
 *
 * In the real application this should call into the Rust formula engine
 * (Tauri command / MessageChannel). For now we keep the API surface here so
 * the UI can be wired up without blocking on the IPC layer.
 */
export async function runSolver(
  _config: SolverConfig,
  _onProgress: (p: SolverProgress) => void,
  _signal?: AbortSignal,
): Promise<SolverOutcome> {
  throw new Error("runSolver is not wired to the backend yet");
}

