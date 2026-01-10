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
