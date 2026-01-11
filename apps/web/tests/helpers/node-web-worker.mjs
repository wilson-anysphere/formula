import { parentPort, workerData } from "node:worker_threads";

if (!parentPort) {
  throw new Error("node-web-worker wrapper must run inside a worker_threads Worker");
}

// Minimal Web Worker compatibility layer for `packages/extension-host/src/browser/extension-worker.mjs`.
// The browser worker script expects `self.addEventListener("message", ...)` and `postMessage(...)`.
const listeners = new Map();

function ensureSet(type) {
  if (!listeners.has(type)) listeners.set(type, new Set());
  return listeners.get(type);
}

globalThis.self = globalThis;

globalThis.addEventListener = (type, listener) => {
  if (typeof listener !== "function") return;
  ensureSet(String(type)).add(listener);
};

globalThis.removeEventListener = (type, listener) => {
  const set = listeners.get(String(type));
  if (!set) return;
  set.delete(listener);
};

globalThis.postMessage = (message) => {
  parentPort.postMessage(message);
};

parentPort.on("message", (message) => {
  const set = listeners.get("message");
  if (!set) return;
  for (const listener of [...set]) {
    try {
      listener({ data: message });
    } catch {
      // ignore
    }
  }
});

await import(workerData.url);

