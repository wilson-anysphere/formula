const STORAGE_KEY = "formula.gpuAcceleration.enabled";

export function loadGpuAccelerationSettings() {
  if (typeof localStorage === "undefined") {
    return { enabled: true };
  }
  const value = localStorage.getItem(STORAGE_KEY);
  if (value == null) return { enabled: true };
  return { enabled: value === "true" };
}

export function saveGpuAccelerationSettings(settings) {
  if (typeof localStorage === "undefined") return;
  localStorage.setItem(STORAGE_KEY, settings.enabled ? "true" : "false");
}

