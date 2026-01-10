// This file is meant to run in a browser Worker context.
// It is not exercised in Node tests.

let pyodide = null;

async function loadPyodideOnce() {
  if (pyodide) return pyodide;

  // Load Pyodide from the official CDN by default. Integrators can host this
  // locally and override `indexURL` when bundling.
  // eslint-disable-next-line no-undef
  importScripts("https://cdn.jsdelivr.net/pyodide/v0.25.1/full/pyodide.js");

  // eslint-disable-next-line no-undef
  pyodide = await self.loadPyodide({
    indexURL: "https://cdn.jsdelivr.net/pyodide/v0.25.1/full/",
  });

  return pyodide;
}

self.onmessage = async (event) => {
  const msg = event.data;
  if (!msg || typeof msg.type !== "string") return;

  if (msg.type === "init") {
    try {
      const runtime = await loadPyodideOnce();
      // Preload common packages if desired; kept minimal here.
      // await runtime.loadPackage(["numpy", "pandas"]);
      self.postMessage({ type: "ready" });
    } catch (err) {
      self.postMessage({ type: "ready", error: err?.message ?? String(err) });
    }
    return;
  }

  if (msg.type === "execute") {
    const requestId = msg.requestId;
    try {
      const runtime = await loadPyodideOnce();
      const result = await runtime.runPythonAsync(msg.code);
      self.postMessage({ type: "result", requestId, success: true, result });
    } catch (err) {
      self.postMessage({ type: "result", requestId, success: false, error: err?.message ?? String(err) });
    }
  }
};

