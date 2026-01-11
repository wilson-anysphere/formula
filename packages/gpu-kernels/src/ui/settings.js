const STORAGE_KEY = "formula.gpuAcceleration.enabled";
const STORAGE_KEY_PRECISION = "formula.gpuAcceleration.precision";

export function loadGpuAccelerationSettings() {
  if (typeof localStorage === "undefined") {
    return { enabled: true, precision: "excel" };
  }
  const value = localStorage.getItem(STORAGE_KEY);
  const precision = localStorage.getItem(STORAGE_KEY_PRECISION);
  const enabled = value == null ? true : value === "true";
  return { enabled, precision: precision === "fast" ? "fast" : "excel" };
}

export function saveGpuAccelerationSettings(settings) {
  if (typeof localStorage === "undefined") return;
  if ("enabled" in settings) {
    localStorage.setItem(STORAGE_KEY, settings.enabled ? "true" : "false");
  }
  if ("precision" in settings) {
    localStorage.setItem(STORAGE_KEY_PRECISION, settings.precision === "fast" ? "fast" : "excel");
  }
}
