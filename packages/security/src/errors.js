export class PermissionDeniedError extends Error {
  /**
   * @param {object} options
   * @param {{type: string, id: string}} options.principal
   * @param {object} options.request
   * @param {string} options.reason
   */
  constructor({ principal, request, reason }) {
    super(reason);
    this.name = "PermissionDeniedError";
    this.code = "PERMISSION_DENIED";
    this.principal = principal;
    this.request = request;
    this.reason = reason;
  }
}

export class SandboxTimeoutError extends Error {
  /**
   * @param {object} options
   * @param {number} options.timeoutMs
   */
  constructor({ timeoutMs }) {
    super(`Sandbox execution timed out after ${timeoutMs}ms`);
    this.name = "SandboxTimeoutError";
    this.code = "SANDBOX_TIMEOUT";
    this.timeoutMs = timeoutMs;
  }
}

export class SandboxOutputLimitError extends Error {
  /**
   * @param {object} options
   * @param {number} options.maxBytes
   */
  constructor({ maxBytes }) {
    super(`Sandbox output exceeded limit of ${maxBytes} bytes`);
    this.name = "SandboxOutputLimitError";
    this.code = "SANDBOX_OUTPUT_LIMIT";
    this.maxBytes = maxBytes;
  }
}

export class SandboxMemoryLimitError extends Error {
  /**
   * @param {object} options
   * @param {number} options.memoryMb
   * @param {number} [options.usedMb]
   */
  constructor({ memoryMb, usedMb = null }) {
    const suffix = usedMb ? ` (used ~${usedMb}MB)` : "";
    super(`Sandbox exceeded memory limit of ${memoryMb}MB${suffix}`);
    this.name = "SandboxMemoryLimitError";
    this.code = "SANDBOX_MEMORY_LIMIT";
    this.memoryMb = memoryMb;
    this.usedMb = usedMb;
  }
}
