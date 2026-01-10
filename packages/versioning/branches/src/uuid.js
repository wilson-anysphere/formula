/**
 * Small `randomUUID()` helper that works in both modern browsers and Node.
 *
 * We avoid importing `node:crypto` so this module can be used in the desktop
 * frontend bundle without relying on Node built-ins.
 *
 * @returns {string}
 */
export function randomUUID() {
  if (globalThis.crypto && typeof globalThis.crypto.randomUUID === "function") {
    return globalThis.crypto.randomUUID();
  }

  // Best-effort fallback for environments without Web Crypto.
  const rand32 = () => Math.floor(Math.random() * 0xffffffff).toString(16).padStart(8, "0");
  return `uuid-${Date.now().toString(16)}-${rand32()}-${rand32()}-${rand32()}-${rand32()}`;
}

