export class SiemConfigStore {
  constructor() {
    this.configByOrgId = new Map();
  }

  get(orgId) {
    return this.configByOrgId.get(orgId) || null;
  }

  set(orgId, config) {
    this.configByOrgId.set(orgId, config);
  }

  delete(orgId) {
    return this.configByOrgId.delete(orgId);
  }

  getSanitized(orgId) {
    const config = this.get(orgId);
    if (!config) return null;

    const sanitized = { ...config };
    if (sanitized.auth) {
      sanitized.auth = { ...sanitized.auth };
      if (sanitized.auth.type === "bearer" && sanitized.auth.token) sanitized.auth.token = "***";
      if (sanitized.auth.type === "basic") {
        if (sanitized.auth.username) sanitized.auth.username = "***";
        if (sanitized.auth.password) sanitized.auth.password = "***";
      }
      if (sanitized.auth.type === "header" && sanitized.auth.value) sanitized.auth.value = "***";
    }

    return sanitized;
  }
}
